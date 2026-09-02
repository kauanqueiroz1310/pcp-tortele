import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx-js-style";

/* ============================================================
   PCP TORTELÊ — WEB v2
   Paridade completa com a planilha V3:
   - PCP Semanal: 8 semanas, parcial, Ajuste Combo, Média, DP, CV,
     Mín, Máx, Tendência, ES, Margem %, Sugerida, Estoque, Líquida,
     ABC, Alertas, Líquida por dia + total
   - Programação de produção (salgados em dias alternados) +
     tabela auxiliar quinzenal (Sugerida por dia)
   - PCP por Loja: mix, Sugerida/Estoque/Líquida POR loja (estoque
     físico real, com CD), Envio semanal por dia com seletor de loja
   - Auditoria CMV: qde, venda, custo, custo unit., preço médio, CMV%
   - Uploads: categorias (AUX CATEGORIA), combos (AUX COMBOS),
     estoque (formato colunar OU blocos horizontais, com CD)
   - Persistência: window.storage (dados voltam ao reabrir)
   ============================================================ */

const Z_TABLE = { 0.8: 0.84, 0.9: 1.28, 0.95: 1.65, 0.99: 2.33 };
const LOJAS = ["ALDEOTA", "MEIRELES", "RIO MAR", "SUL"];
const LOCAIS_ESTOQUE = [...LOJAS, "CD"];
const DIAS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"];

/* ---------------- helpers ---------------- */
function parseDateBR(s) {
  if (s instanceof Date) return isNaN(s) ? null : s;
  if (typeof s === "number") {
    const d = XLSX.SSF.parse_date_code(s);
    return d ? new Date(d.y, d.m - 1, d.d) : null;
  }
  if (typeof s !== "string") return null;
  const m = s.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  const dt = new Date(+m[3], +m[2] - 1, +m[1]);
  return isNaN(dt) ? null : dt;
}
const weekStart = (d) => { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); x.setHours(0,0,0,0); return x; };
const fmtDM = (d) => `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}`;
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const normLoja = (s) => {
  const n = String(s || "").toUpperCase().replace(/\s+/g, "");
  if (n.includes("ALDEOTA")) return "ALDEOTA";
  if (n.includes("MEIRELES")) return "MEIRELES";
  if (n.includes("RIOMAR") || n === "RM") return "RIO MAR";
  if (n === "SUL" || n.includes("SUL")) return "SUL";
  if (n === "CD" || n.includes("DISTRIB")) return "CD";
  return null;
};
const isSalgado = (cat) => {
  const c = String(cat || "").toLowerCase();
  return c.includes("salgado");
};

// Distribui um total inteiro entre N grupos com pesos, garantindo sum(result) === total
function distribute(total, weights) {
  const wsum = weights.reduce((a,b)=>a+b,0);
  if (!wsum || !total) return weights.map(()=>0);
  const exact = weights.map(w=>(total*w)/wsum);
  const result = exact.map(Math.floor);
  let rem = total - result.reduce((a,b)=>a+b,0);
  const order = exact.map((e,i)=>[e-Math.floor(e),i]).sort((a,b)=>b[0]-a[0]);
  for (let k=0; k<rem; k++) result[order[k][1]]++;
  return result;
}

/* ---------------- parsers ---------------- */
function parseVendas(wb) {
  const rows = [], warnings = [];
  for (const name of wb.SheetNames) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    if (!data.length) continue;
    let hi = -1;
    for (let i = 0; i < Math.min(data.length, 8); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      if (r.some((c) => c.includes("data_emissao") || c === "data")) { hi = i; break; }
    }
    if (hi < 0) {
      // Tenta formato nativo do sistema (izzyway): col0=vazio, col1=Cod., col2=Produto, col3=Qde.
      // Múltiplas datas intercaladas: "Data Emissão: DD/MM/YYYY" em col0
      const h0 = (data[0] || []).map((c) => String(c || "").toLowerCase().trim());
      if (h0[1] === "cod." && h0[2] === "produto" && (h0[3] || "").startsWith("qde")) {
        let currentDate = null;
        for (let i = 1; i < data.length; i++) {
          const r = data[i] || [];
          const c0 = String(r[0] ?? "").trim();
          const dm = c0.match(/(\d{2}\/\d{2}\/\d{4})/);
          if (dm) { currentDate = parseDateBR(dm[1]); continue; }
          if (!currentDate || !c0 || !/^\d+$/.test(c0)) continue;
          const cod = parseInt(String(r[1] || ""), 10);
          const qde = parseFloat(r[3]);
          if (!cod || isNaN(cod) || !qde || isNaN(qde) || qde <= 0) continue;
          rows.push({ dt: currentDate, cod, produto: String(r[2] || "").trim(),
            qde, vendido: parseFloat(r[5]) || 0, custo: parseFloat(r[6]) || 0 });
        }
        if (rows.length) {
          const nd = new Set(rows.map((r) => r.dt.getTime())).size;
          warnings.push(`Formato exportação sistema: ${rows.length} registros em ${nd} datas.`);
          break;
        }
      }
      continue;
    }
    const header = data[hi].map((c) => String(c || "").toLowerCase().trim());
    const iData = header.findIndex((c) => c.includes("data"));
    const iCod = header.findIndex((c) => c === "cod" || c === "cod." || c === "código" || c === "codigo");
    let iProd = header.findIndex((c) => c.includes("produto") || c.includes("descri") || c === "nome" || c === "item");
    // Fallback: primeira coluna de texto após cod (ex: col "Produto" com outro nome)
    if (iProd < 0 && iCod >= 0) {
      for (let k = iCod + 1; k < header.length; k++) {
        const h = header[k];
        if (!h.match(/^(qde|qtd|val|cust|prec|dat|med)/)) { iProd = k; break; }
      }
    }
    const iQde = header.findIndex((c) => c.startsWith("qde") || c.startsWith("qtd"));
    const iVend = header.findIndex((c) => c === "vendido");
    const iCusto = header.findIndex((c) => c === "custo");
    const composite = iCod < 0;

    for (let i = hi + 1; i < data.length; i++) {
      const r = data[i] || [];
      const dt = parseDateBR(r[iData]);
      if (!dt) continue;
      let cod, produto, qde, vendido, custo;
      if (composite) {
        const m = String(r[iData + 1] || "").match(/^(\d+)\s*-\s*(.+)$/);
        if (!m) continue;
        cod = +m[1]; produto = m[2].trim();
        qde = parseFloat(r[iData + 2]);
        vendido = parseFloat(r[iData + 4]) || 0;
        custo = parseFloat(r[iData + 5]) || 0;
      } else {
        cod = parseInt(r[iCod], 10);
        produto = String(r[iProd] || "").trim();
        qde = parseFloat(r[iQde]);
        vendido = iVend >= 0 ? parseFloat(r[iVend]) || 0 : 0;
        custo = iCusto >= 0 ? parseFloat(r[iCusto]) || 0 : 0;
      }
      if (!cod || isNaN(cod) || !qde || isNaN(qde) || qde <= 0) continue;
      rows.push({ dt, cod, produto, qde, vendido, custo });
    }
    if (rows.length) {
      if (composite) warnings.push(`Aba "${name}": formato composto detectado e tratado.`);
      break;
    }
  }
  return { rows, warnings };
}

function parseEstoquePorLoja(wb) {
  // retorna { cod: {ALDEOTA: n, MEIRELES: n, "RIO MAR": n, SUL: n, CD: n} }
  const map = {};
  const add = (cod, loja, v) => {
    if (!cod || isNaN(cod) || !loja || isNaN(v)) return;
    if (!map[cod]) map[cod] = {};
    map[cod][loja] = (map[cod][loja] || 0) + v;
  };
  let formato = null;
  for (const name of wb.SheetNames) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    // localizar linha de cabeçalho com "cod"
    let hi = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      if (r.filter((c) => c === "cod" || c === "código" || c === "codigo").length >= 1) { hi = i; break; }
    }
    if (hi < 0) continue;
    const header = data[hi].map((c) => String(c || "").trim());
    const lower = header.map((c) => c.toLowerCase());

    // FORMATO A (colunar): Cod | Produto | ALDEOTA | MEIRELES | RIOMAR | SUL | CD | ...
    const lojaCols = [];
    header.forEach((h, idx) => { const l = normLoja(h); if (l) lojaCols.push([idx, l]); });
    const codCols = [];
    lower.forEach((h, idx) => { if (h === "cod" || h === "código" || h === "codigo") codCols.push(idx); });

    if (codCols.length === 1 && lojaCols.length >= 2) {
      formato = "colunar";
      for (let i = hi + 1; i < data.length; i++) {
        const r = data[i] || [];
        const cod = parseInt(r[codCols[0]], 10);
        for (const [idx, loja] of lojaCols) add(cod, loja, parseFloat(r[idx]));
      }
      break;
    }
    // FORMATO B (blocos): Cod|Produto|LOJA repetidos horizontalmente
    if (codCols.length >= 2) {
      formato = "blocos";
      // para cada coluna de cod, a loja é o header 2 posições à direita
      const blocos = codCols.map((ci) => {
        for (let k = ci + 1; k <= ci + 3 && k < header.length; k++) {
          const l = normLoja(header[k]);
          if (l) return [ci, k, l];
        }
        return null;
      }).filter(Boolean);
      for (let i = hi + 1; i < data.length; i++) {
        const r = data[i] || [];
        for (const [ci, vi, loja] of blocos) add(parseInt(r[ci], 10), loja, parseFloat(r[vi]));
      }
      break;
    }
    // FORMATO C (simples): Cod + primeiro número — vira total no CD
    formato = "simples";
    for (let i = hi + 1; i < data.length; i++) {
      const r = data[i] || [];
      const cod = parseInt(r[0], 10);
      for (let k = 1; k < Math.min(4, r.length); k++) {
        const v = parseFloat(r[k]);
        if (typeof r[k] === "number" && !isNaN(v)) { add(cod, "CD", v); break; }
      }
    }
    break;
  }
  return { map, formato };
}

// Formato simples de estoque: um arquivo por loja (Cod | Produto | Qty)
// Aceita também o formato izzyway (col0 vazio, col1=Cod., col3=Qde.)
function parseEstoqueSimples(wb) {
  const map = {};
  for (const name of wb.SheetNames) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    let hi = -1, iCod = -1, iQty = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      const ic = r.findIndex((c) => c === "cod" || c === "cod." || c === "código" || c === "codigo");
      if (ic < 0) continue;
      for (let k = 0; k < r.length; k++) {
        if (k === ic) continue;
        if (r[k].startsWith("saldo") || r[k].startsWith("qde") || r[k].startsWith("qtd") || r[k] === "estoque" || r[k] === "quantidade") {
          hi = i; iCod = ic; iQty = k; break;
        }
      }
      if (hi >= 0) break;
    }
    // Fallback: formato izzyway vendas (col0=vazio, col1=Cod., col3=Qde.)
    if (hi < 0) {
      const h0 = (data[0] || []).map((c) => String(c || "").toLowerCase().trim());
      if (h0[1] === "cod." && h0[2] === "produto" && (h0[3] || "").startsWith("qde")) {
        hi = 0; iCod = 1; iQty = 3;
      }
    }
    // Fallback: formato estoque bruto Izzyway (Produto|Medida|_|Quantidade|_ — cod em col0, qty em col4)
    if (hi < 0) {
      const h0 = (data[0] || []).map((c) => String(c || "").toLowerCase().trim());
      if (h0[0] === "produto" && h0[3] === "quantidade") {
        hi = 0; iCod = 0; iQty = 4;
      }
    }
    if (hi < 0 || iQty < 0) continue;
    for (let i = hi + 1; i < data.length; i++) {
      const r = data[i] || [];
      const cod = parseInt(String(r[iCod] || ""), 10);
      const qty = parseFloat(r[iQty]);
      if (!cod || isNaN(cod) || isNaN(qty)) continue;
      map[cod] = (map[cod] || 0) + qty;
    }
    if (Object.keys(map).length) break;
  }
  return map;
}

function parseCategorias(wb) {
  const map = {};
  for (const name of wb.SheetNames) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    let hi = -1, iCod = -1, iCat = -1, iSetor = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      const ic = r.findIndex((c) => c === "código" || c === "codigo" || c === "cod");
      const ik = r.findIndex((c) => c === "categoria");
      if (ic >= 0 && ik >= 0) {
        hi = i; iCod = ic; iCat = ik;
        iSetor = r.findIndex((c) => c === "setor");
        break;
      }
    }
    if (hi < 0) continue;
    for (let i = hi + 1; i < data.length; i++) {
      const r = data[i] || [];
      const cod = parseInt(r[iCod], 10);
      const cat = String(r[iCat] || "").trim();
      const setor = iSetor >= 0 ? String(r[iSetor] || "").trim() : "";
      if (cod && !isNaN(cod) && cat) map[cod] = { cat, setor };
    }
    if (Object.keys(map).length) break;
  }
  return map;
}

function parseCombos(wb) {
  const list = [];
  for (const name of wb.SheetNames) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    let hi = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      if (r.some((c) => c.includes("combo")) && r.some((c) => c.includes("componente"))) { hi = i; break; }
    }
    const start = hi >= 0 ? hi + 1 : 0;
    for (let i = start; i < data.length; i++) {
      const r = data[i] || [];
      const codCombo = parseInt(r[0], 10);
      const codComp = parseInt(r[2], 10);
      const qde = parseFloat(r[4]);
      if (codCombo && codComp && qde && !isNaN(codCombo) && !isNaN(codComp) && !isNaN(qde)) {
        list.push({ codCombo, nomeCombo: String(r[1] || ""), codComp, nomeComp: String(r[3] || ""), qde });
      }
    }
    if (list.length) break;
  }
  return list;
}

function parseFichasTecnicas(wb) {
  const ft = {};
  const sh = wb.Sheets[wb.SheetNames[0]];
  if (!sh) return { ft: {}, sf: {}, sfCodes: new Set() };
  const data = XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: null });

  let currentCod = null, currentRend = 1, inIng = false;
  for (const r of data) {
    const r0 = String(r[0] || "").trim();
    const r1 = String(r[1] || "").trim();
    // Product header: col 0 = numeric code, col 2 = product name, col 12 = rendimento
    if (/^\d{3,}$/.test(r0) && r[2]) {
      currentCod = +r0;
      const rend = r[12];
      currentRend = (rend != null && +rend > 0) ? +rend : 1;
      inIng = false;
    }
    // Ingredient section header: col 1 = "Produto", col 7 = "Sigla"
    if (r1 === "Produto" && String(r[7] || "").trim() === "Sigla") { inIng = true; continue; }
    // End of ingredient section
    if (r1.startsWith("Produzir") || r1.startsWith("Observa")) inIng = false;
    // Ingredient row: col 1 = "CODE  NAME", col 7 = unit, col 9 = quantity
    if (inIng && currentCod && r[9] != null) {
      const m = r1.match(/^(\d+)\s+(.+)$/);
      if (m) {
        const codInsumo = +m[1], nome = m[2].trim(), quant = parseFloat(r[9]);
        if (codInsumo && !isNaN(quant) && quant > 0)
          (ft[currentCod] ||= []).push({ codInsumo, nome, quant: quant / currentRend, unit: String(r[7] || "").trim() });
      }
    }
  }

  // Sub-products: ingredient codes that also have their own FT entry
  const allIngCodes = new Set();
  for (const ings of Object.values(ft)) for (const ing of ings) allIngCodes.add(ing.codInsumo);
  const sfCodes = new Set([...allIngCodes].filter(c => ft[c]));
  const sf = {};
  for (const cod of sfCodes) sf[cod] = ft[cod];

  return { ft, sf, sfCodes };
}

/* ---------------- cálculo ---------------- */
function computeAll(sales, params, estoqueLoja, categorias, combos) {
  const { nivelServico, janela, dataRef, usarCombos } = params;
  const Z = Z_TABLE[nivelServico] ?? 1.28;
  const ref = new Date(dataRef);
  const dowRef = (ref.getDay() + 6) % 7;
  const lastSunday = dowRef === 6 ? ref : addDays(ref, -(dowRef + 1));
  const lastWeekStart = addDays(lastSunday, -6);
  lastWeekStart.setHours(0, 0, 0, 0); // normaliza para meia-noite — timestamps devem bater com weekStart()
  const weekStarts = []; for (let i = 7; i >= 0; i--) weekStarts.push(addDays(lastWeekStart, -7*i));
  const partialWeekStart = addDays(lastWeekStart, 7);
  const mix4Start = addDays(lastWeekStart, -21), mix4End = addDays(lastWeekStart, 6);

  const byPW = {}, byPL = {}, byPD = {}, names = {}, cmv = {};
  for (const s of sales) {
    const ws_ = weekStart(s.dt).getTime();
    (byPW[s.cod] ||= {})[ws_] = (byPW[s.cod][ws_] || 0) + s.qde;
    names[s.cod] ||= s.produto;
    const c = (cmv[s.cod] ||= { qde: 0, venda: 0, custo: 0 });
    c.qde += s.qde; c.venda += s.vendido || 0; c.custo += s.custo || 0;
    if (s.dt >= mix4Start && s.dt <= mix4End) {
      (byPL[s.cod] ||= {})[s.loja] = (byPL[s.cod][s.loja] || 0) + s.qde;
      const dow = (s.dt.getDay() + 6) % 7;
      (byPD[s.cod] ||= {})[dow] = (byPD[s.cod][dow] || 0) + s.qde;
    }
  }

  // vendas médias de cada combo na janela
  const janelaStarts = weekStarts.slice(8 - janela);
  const comboMedias = {};
  if (usarCombos && combos.length) {
    const combosSet = new Set(combos.map((c) => c.codCombo));
    for (const cc of combosSet) {
      const vals = janelaStarts.map((w) => byPW[cc]?.[w.getTime()] || 0);
      comboMedias[cc] = vals.reduce((a, b) => a + b, 0) / vals.length;
    }
  }
  const ajusteCombo = {};
  if (usarCombos) for (const c of combos) {
    ajusteCombo[c.codComp] = (ajusteCombo[c.codComp] || 0) + c.qde * (comboMedias[c.codCombo] || 0);
  }

  const rows = Object.keys(byPW).map((k) => {
    const cod = +k;
    const weeks = weekStarts.map((w) => byPW[cod][w.getTime()] || 0);
    const partial = byPW[cod][partialWeekStart.getTime()] || 0;
    const jv = janelaStarts.map((w) => byPW[cod][w.getTime()] || 0);
    const mediaBase = jv.reduce((a,b)=>a+b,0) / jv.length;
    const combo = ajusteCombo[cod] || 0;
    const media = mediaBase + combo;
    const dp = jv.length > 1 ? Math.sqrt(jv.reduce((a,b)=>a+(b-mediaBase)**2,0)/(jv.length-1)) : 0;
    const cv = media > 0 ? dp / media : 0;
    const min = Math.min(...jv), max = Math.max(...jv);
    const es = Z * dp;
    const margem = media > 0 ? es / media : 0;
    const sugerida = Math.ceil(media + es);
    const estL = estoqueLoja[cod] || {};
    const estoquePorLocal = LOCAIS_ESTOQUE.map((l) => Math.max(0, estL[l] || 0));
    const estoque = estoquePorLocal.reduce((a,b)=>a+b,0);
    const liquida = Math.max(0, sugerida - estoque);
    const ultima = weeks[7];
    const tend = media > 0 ? (ultima - media) / media : 0;
    const semVenda = weeks.filter((v)=>v>0).length;
    const catObj = categorias[cod];
    const categoria = typeof catObj === "string" ? catObj : (catObj?.cat || "S/CAT");
    const setor = typeof catObj === "string" ? "" : (catObj?.setor || "");

    const lojaTot = LOJAS.reduce((a,l)=>a+(byPL[cod]?.[l]||0),0);
    const mixLoja = LOJAS.map((l)=> lojaTot>0 ? (byPL[cod]?.[l]||0)/lojaTot : 0);
    const sugLoja = mixLoja.map((m)=>Math.ceil(sugerida*m));
    const estPorLoja = LOJAS.map((l)=>Math.max(0, estL[l]||0));
    const liqLoja = sugLoja.map((s,i)=>Math.max(0, s - estPorLoja[i]));

    const dowTot = Array.from({length:7},(_,d)=>byPD[cod]?.[d]||0);
    const dowSum = dowTot.reduce((a,b)=>a+b,0);
    const mixDia = dowTot.map((v)=>dowSum>0?v/dowSum:0);
    const liqDia = mixDia.map((m)=>Math.ceil(liquida*m));
    const sugDia = mixDia.map((m)=>Math.ceil(sugerida*m)); // quinzenal aux

    // PROGRAMAÇÃO: salgados produzem em dias alternados
    const salg = isSalgado(categoria);
    const prog = salg
      ? [liqDia[1]+liqDia[2], null, liqDia[3]+liqDia[4], null, liqDia[5]+liqDia[6], liqDia[0], null]
      : [liqDia[0], liqDia[1], liqDia[2], liqDia[3], liqDia[4], liqDia[5]+liqDia[6], null];

    const alertas = [];
    if (categoria === "S/CAT") alertas.push("S/CAT");
    if (semVenda < 4) alertas.push("POUCO HIST");
    if (weeks[7]===0 && weeks[6]===0) alertas.push("PARADO");
    if (cv > 0.5) alertas.push("ALTA VAR");
    if (tend < -0.4) alertas.push("QUEDA");
    if (tend > 0.5) alertas.push("ALTA");

    const cm = cmv[cod];
    const cmvPct = cm.venda > 0 ? cm.custo / cm.venda : (cm.custo > 0 ? null : 0);

    return { cod, produto: names[cod], categoria, setor, weeks, partial, combo, media, dp, cv, min, max,
      tend, es, margem, sugerida, estoque, estoquePorLocal, liquida, mixLoja, sugLoja, estPorLoja,
      liqLoja, mixDia, liqDia, sugDia, prog, alertas,
      cmvQde: cm.qde, cmvVenda: cm.venda, cmvCusto: cm.custo,
      custoUnit: cm.qde>0?cm.custo/cm.qde:0, precoMedio: cm.qde>0?cm.venda/cm.qde:0, cmvPct };
  });

  const withVol = rows.map((r)=>({...r, vol4: r.weeks.slice(4).reduce((a,b)=>a+b,0)}));
  withVol.sort((a,b)=>b.vol4-a.vol4);
  const totVol = withVol.reduce((a,r)=>a+r.vol4,0);
  let ac = 0;
  for (const r of withVol) { ac += r.vol4; r.abc = totVol===0?"C":ac/totVol<=0.8?"A":ac/totVol<=0.95?"B":"C"; }
  return { rows: withVol, weekStarts, partialWeekStart, lastWeekStart, Z, nCombos: combos.length };
}

/* ---------------- persistência (localStorage) ---------------- */
// Shim: se rodar fora do runtime de artifact (Vercel, navegador comum), usa localStorage
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    set: (k, v) => { try { localStorage.setItem(k, v); } catch(_){} return Promise.resolve(); },
    get: (k) => { try { const v = localStorage.getItem(k); return Promise.resolve(v !== null ? { value: v } : null); } catch(_){ return Promise.resolve(null); } },
    delete: (k) => { try { localStorage.removeItem(k); } catch(_){} return Promise.resolve(); },
  };
}
const CHUNK = 20000;
async function saveState({ files, estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos, params, progEdits, progWeekStart, fichasTec, ftInfo }) {
  try {
    const flat = [];
    const fmeta = files.map((f) => ({ name: f.name, loja: f.loja, n: f.rows.length }));
    for (const f of files) if (f.loja) {
      const li = LOJAS.indexOf(f.loja);
      for (const r of f.rows) flat.push([Math.round(r.dt.getTime()/86400000), r.cod, r.qde, li, r.vendido||0, r.custo||0]);
    }
    const nChunks = Math.ceil(flat.length / CHUNK);
    for (let i = 0; i < nChunks; i++) {
      await window.storage.set(`sales:${i}`, JSON.stringify(flat.slice(i*CHUNK,(i+1)*CHUNK)));
    }
    await window.storage.set("meta", JSON.stringify({ nChunks, fmeta, savedAt: Date.now() }));
    const prodNames = {};
    for (const f of files) for (const r of f.rows) if (r.cod && r.produto) prodNames[r.cod] = r.produto;
    const ftSer = fichasTec ? { ft: fichasTec.ft, sf: fichasTec.sf, sfCodes: [...fichasTec.sfCodes] } : null;
    await window.storage.set("aux", JSON.stringify({ estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos, params, prodNames, progEdits: progEdits || {}, progWeekStart: progWeekStart || "", fichasTec: ftSer, ftInfo: ftInfo || null }));
    return true;
  } catch (e) { console.error("save", e); return false; }
}
async function loadState() {
  try {
    const metaR = await window.storage.get("meta");
    if (!metaR) return null;
    const meta = JSON.parse(metaR.value);
    const flat = [];
    for (let i = 0; i < meta.nChunks; i++) {
      const r = await window.storage.get(`sales:${i}`);
      if (r) flat.push(...JSON.parse(r.value));
    }
    const auxR = await window.storage.get("aux");
    const aux = auxR ? JSON.parse(auxR.value) : {};
    const prodNames = aux.prodNames || {};
    const byLoja = {};
    for (const [d, cod, qde, li, vendido, custo] of flat) {
      const loja = LOJAS[li];
      (byLoja[loja] ||= []).push({ dt: new Date(d*86400000), cod, produto: prodNames[cod] || "", qde, vendido, custo, loja });
    }
    const files = meta.fmeta.filter((f)=>f.loja).map((f)=>({ name: f.name, loja: f.loja, rows: [], warnings: [], fromCache: true }));
    // reconstruir rows por loja na ordem dos files (aproximação: agrupar por loja)
    const grouped = {};
    for (const l of LOJAS) grouped[l] = byLoja[l] || [];
    const outFiles = [];
    for (const l of LOJAS) if (grouped[l].length) outFiles.push({ name: `(sessão anterior) ${l}`, loja: l, rows: grouped[l], warnings: [], fromCache: true });
    const ftRaw = aux.fichasTec;
    const fichasTec = ftRaw ? { ft: ftRaw.ft, sf: ftRaw.sf, sfCodes: new Set(ftRaw.sfCodes || []) } : null;
    return { files: outFiles, ...aux, fichasTec, savedAt: meta.savedAt };
  } catch (e) { console.error("load", e); return null; }
}
async function clearState() {
  try {
    const metaR = await window.storage.get("meta");
    if (metaR) {
      const meta = JSON.parse(metaR.value);
      for (let i = 0; i < meta.nChunks; i++) await window.storage.delete(`sales:${i}`).catch(()=>{});
    }
    await window.storage.delete("meta").catch(()=>{});
    await window.storage.delete("aux").catch(()=>{});
  } catch (e) {}
}

/* ---------------- export ---------------- */
const XS = {
  hdr: { font:{bold:true,color:{rgb:"FFFFFF"},sz:11}, fill:{patternType:"solid",fgColor:{rgb:"25211C"}}, alignment:{horizontal:"center",vertical:"center",wrapText:true}, border:{bottom:{style:"medium",color:{rgb:"B96A1B"}}} },
  hdrOrange: { font:{bold:true,color:{rgb:"FFFFFF"},sz:11}, fill:{patternType:"solid",fgColor:{rgb:"B96A1B"}}, alignment:{horizontal:"center",vertical:"center",wrapText:true} },
  rowA: { fill:{patternType:"solid",fgColor:{rgb:"F6F3EE"}}, alignment:{vertical:"center"} },
  rowB: { fill:{patternType:"solid",fgColor:{rgb:"FFFFFF"}}, alignment:{vertical:"center"} },
  num: (alt) => ({ ...(alt?XS.rowA:XS.rowB), alignment:{horizontal:"right",vertical:"center"}, numFmt:"#,##0" }),
  dec: (alt) => ({ ...(alt?XS.rowA:XS.rowB), alignment:{horizontal:"right",vertical:"center"}, numFmt:"#,##0.0" }),
  pct: (alt) => ({ ...(alt?XS.rowA:XS.rowB), alignment:{horizontal:"right",vertical:"center"}, numFmt:'0.0"%"' }),
  bold: (alt) => ({ ...(alt?XS.rowA:XS.rowB), font:{bold:true}, alignment:{vertical:"center"} }),
  red: { fill:{patternType:"solid",fgColor:{rgb:"FFEEEA"}}, font:{bold:true,color:{rgb:"C4501E"}}, alignment:{horizontal:"right",vertical:"center"}, numFmt:"#,##0" },
  grn: { fill:{patternType:"solid",fgColor:{rgb:"EAF4EE"}}, font:{bold:true,color:{rgb:"2D6A4F"}}, alignment:{horizontal:"right",vertical:"center"}, numFmt:"#,##0" },
};

const LOGO_HDR_STYLE = { font:{bold:true,sz:14,color:{rgb:"F0DBBF"}}, fill:{patternType:"solid",fgColor:{rgb:"3C2008"}}, alignment:{horizontal:"left",vertical:"center"} };
const LOGO_SUB_STYLE = { font:{sz:10,italic:true,color:{rgb:"D4B896"}}, fill:{patternType:"solid",fgColor:{rgb:"3C2008"}}, alignment:{horizontal:"right",vertical:"center"} };

function styledSheet(data, cols, title) {
  const ws = {};
  const nr = data.length, nc = data[0]?.length || 0;
  const dt = new Date().toLocaleDateString("pt-BR");
  // Linha 0: cabeçalho da empresa
  ws[XLSX.utils.encode_cell({r:0, c:0})] = { v: "tortelê", t:"s", s: LOGO_HDR_STYLE };
  for (let c = 1; c < nc-1; c++) ws[XLSX.utils.encode_cell({r:0,c})] = {v:"",t:"s",s:LOGO_HDR_STYLE};
  ws[XLSX.utils.encode_cell({r:0, c:nc-1})] = { v: `${title||""} · ${dt}`, t:"s", s: LOGO_SUB_STYLE };
  ws["!merges"] = [{s:{r:0,c:0},e:{r:0,c:Math.max(0,nc-2)}}];
  // Linhas 1+: header (linha 1) e dados (linhas 2+)
  for (let r = 0; r < nr; r++) {
    for (let c = 0; c < nc; c++) {
      const addr = XLSX.utils.encode_cell({r:r+1, c});
      const v = data[r][c];
      let cell;
      if (r === 0) {
        cell = { v: v ?? "", t: "s", s: XS.hdr };
      } else {
        const alt = r % 2 === 0;
        if (v == null || v === "") { cell = { v: "", t: "s", s: alt?XS.rowA:XS.rowB }; }
        else if (typeof v === "number") { cell = { v, t: "n", s: XS.num(alt) }; }
        else { cell = { v: String(v), t: "s", s: alt?XS.rowA:XS.rowB }; }
      }
      ws[addr] = cell;
    }
  }
  ws["!ref"] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:nr,c:nc-1}});
  ws["!cols"] = (cols || []).map((w) => ({wch: w}));
  ws["!rows"] = [{hpt:28}, {hpt:22}];
  ws["!freeze"] = {xSplit:3, ySplit:2};
  return ws;
}

function exportAll(result) {
  const { rows, weekStarts } = result;
  const wb = XLSX.utils.book_new();
  const dt = new Date().toLocaleDateString("pt-BR");

  // ── PCP Semanal ────────────────────────────────────────────────────
  const h1 = ["Cód","Produto","Categoria",...weekStarts.map(fmtDM),"Parcial","Combo","Média/sem","DP","CV%","Mín","Máx","Tend%","ES","Marg%","Sugerida","Estoque","LÍQUIDA","ABC","Alertas",...DIAS.map(d=>`Líq ${d}`),"Total/dia"];
  const d1 = rows.map((r) => [
    r.cod, r.produto, r.categoria,
    ...r.weeks, r.partial,
    +r.combo.toFixed(1), +r.media.toFixed(1), +r.dp.toFixed(1), +(r.cv*100).toFixed(1),
    r.min, r.max, +(r.tend*100).toFixed(1), +r.es.toFixed(1), +(r.margem*100).toFixed(1),
    r.sugerida, r.estoque, r.liquida, r.abc, r.alertas.join(" | "),
    ...r.liqDia, r.liqDia.reduce((a,b)=>a+b,0),
  ]);
  const ws1 = styledSheet([h1,...d1], [7,38,20,...weekStarts.map(()=>9),9,9,10,8,7,7,7,7,8,7,10,10,10,5,30,...DIAS.map(()=>9),9], "PCP Semanal");
  // Destaque: SUGERIDA verde, LÍQUIDA vermelha (dados começam na linha 2 = logo(0)+header(1)+dados(2+))
  for (let r = 2; r < d1.length + 2; r++) {
    const alt = r % 2 === 0;
    const sugIdx = 3 + weekStarts.length + 12;
    const liqIdx = sugIdx + 2;
    const sugCell = XLSX.utils.encode_cell({r, c: sugIdx});
    const liqCell = XLSX.utils.encode_cell({r, c: liqIdx});
    if (ws1[sugCell]) ws1[sugCell].s = { ...XS.num(alt), font:{bold:true,color:{rgb:"2D6A4F"}} };
    if (ws1[liqCell] && ws1[liqCell].v === 0) ws1[liqCell].s = { ...XS.num(alt), fill:{patternType:"solid",fgColor:{rgb:"F6F3EE"}} };
    else if (ws1[liqCell]) ws1[liqCell].s = { ...XS.red };
  }
  XLSX.utils.book_append_sheet(wb, ws1, "PCP Semanal");

  // ── PCP por Loja ───────────────────────────────────────────────────
  const h3 = ["Cód","Produto","Categoria","Sugerida","Estoque","LÍQUIDA",...LOJAS.map(l=>`${l} %`),...LOJAS.map(l=>`Sug ${l}`),...LOJAS.map(l=>`Est ${l}`),"Est CD",...LOJAS.map(l=>`Líq ${l}`)];
  const d3 = rows.map((r) => [r.cod,r.produto,r.categoria,r.sugerida,r.estoque,r.liquida,...r.mixLoja.map(m=>+(m*100).toFixed(1)),...r.sugLoja,...r.estPorLoja,r.estoquePorLocal[4],...r.liqLoja]);
  const ws3 = styledSheet([h3,...d3], [7,38,20,10,10,10,...LOJAS.map(()=>8),...LOJAS.map(()=>10),...LOJAS.map(()=>10),10,...LOJAS.map(()=>10)], "PCP por Loja");
  XLSX.utils.book_append_sheet(wb, ws3, "PCP por Loja");

  // ── Envio Diário por Loja ──────────────────────────────────────────
  // Base = SUGERIDA (= média + ES), mesma coluna exibida no PCP Semanal.
  // distribute() garante: sum(dias de uma loja) = total loja, sum(totais lojas) = sugerida.
  const diasAbrev = ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"];
  const h5 = ["Cód","Produto","Categoria","Sugerida","Média","ES",
    ...LOJAS.flatMap(l => diasAbrev.map(d => `${l.slice(0,3)} ${d}`)),
    ...LOJAS.map(l => `Total ${l.slice(0,3)}`)];
  const d5 = rows.map(r => {
    const weekTot = r.sugerida;                                // mesma base do PCP Semanal
    const lojaTots = distribute(weekTot, r.mixLoja);           // soma = weekTot ✓
    const dayCells = lojaTots.map(lt => distribute(lt, r.mixDia)); // soma por loja = lojaTot ✓
    return [
      r.cod, r.produto, r.categoria, r.sugerida, Math.round(r.media), Math.round(r.es),
      ...LOJAS.flatMap((_, li) => dayCells[li]),
      ...lojaTots,
    ];
  });
  const ws5 = styledSheet([h5,...d5],
    [7, 38, 20, 10, 9, 9, ...LOJAS.flatMap(() => diasAbrev.map(() => 8)), ...LOJAS.map(() => 11)],
    "Envio Diário por Loja");
  XLSX.utils.book_append_sheet(wb, ws5, "Envio Diário");

  // ── Programação ────────────────────────────────────────────────────
  const h2 = ["Cód","Produto","Categoria","Sugerida","Estoque","LÍQUIDA","Seg","Ter","Qua","Qui","Sex","Sáb","Dom","Total programado"];
  const d2 = rows.map((r) => [
    r.cod, r.produto, r.categoria, r.sugerida, r.estoque, r.liquida,
    ...r.prog.map((v) => v ?? ""),
    r.prog.reduce((a,b)=>a+(b||0),0),
  ]);
  const ws2 = styledSheet([h2,...d2], [7,38,20,10,10,10,8,8,8,8,8,8,8,14], "Programação");
  for (let r = 2; r < d2.length + 2; r++) {
    for (let c = 6; c <= 12; c++) {
      const addr = XLSX.utils.encode_cell({r, c});
      if (ws2[addr] && ws2[addr].v !== "" && ws2[addr].v > 0)
        ws2[addr].s = { fill:{patternType:"solid",fgColor:{rgb:"EDF1F8"}}, font:{bold:true,color:{rgb:"264478"}}, alignment:{horizontal:"center",vertical:"center"}, numFmt:"#,##0" };
    }
  }
  XLSX.utils.book_append_sheet(wb, ws2, "Programação");

  // ── Auditoria CMV ──────────────────────────────────────────────────
  const h4 = ["Cód","Produto","Categoria","Qde Vendida","Venda Total","Custo Total","Custo Unit Médio","Preço Médio","CMV %"];
  const d4 = rows.map((r) => [r.cod,r.produto,r.categoria,+r.cmvQde.toFixed(0),+r.cmvVenda.toFixed(2),+r.cmvCusto.toFixed(2),+r.custoUnit.toFixed(2),+r.precoMedio.toFixed(2),r.cmvPct==null?"SEM VENDA":+(r.cmvPct*100).toFixed(1)]);
  const ws4 = styledSheet([h4,...d4], [7,38,20,12,14,14,16,12,8], "Auditoria CMV");
  for (let r = 2; r < d4.length + 2; r++) {
    const cmvCell = XLSX.utils.encode_cell({r, c:8});
    if (ws4[cmvCell] && typeof ws4[cmvCell].v === "number" && ws4[cmvCell].v > 45)
      ws4[cmvCell].s = { fill:{patternType:"solid",fgColor:{rgb:"FFEEEA"}}, font:{bold:true,color:{rgb:"C4501E"}}, alignment:{horizontal:"right",vertical:"center"}, numFmt:'0.0"%"' };
    else if (ws4[cmvCell] && typeof ws4[cmvCell].v === "number")
      ws4[cmvCell].s = { ...XS.pct(r%2===0) };
  }
  XLSX.utils.book_append_sheet(wb, ws4, "Auditoria CMV");

  XLSX.writeFile(wb, `PCP_Tortele_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ================= UI ================= */
const BRAND = { dark:"#3C2008", cream:"#F0DBBF", amber:"#B96A1B", bg:"#F7F2EB", border:"#E4DDD2", muted:"#7A6450" };
const S = {
  page: { minHeight: "100vh", background: BRAND.bg, color: BRAND.dark, fontFamily: "'Inter', system-ui, sans-serif", fontSize: 14 },
  mono: { fontFamily: "'JetBrains Mono','SF Mono',Consolas,monospace" },
  panel: { background: "#fff", border: `1px solid ${BRAND.border}`, borderRadius: 10, padding: 18, marginBottom: 14 },
  btn: { background: BRAND.amber, color: "#fff", border: "none", borderRadius: 8, padding: "10px 18px", fontWeight: 600, cursor: "pointer", fontSize: 14 },
  btnGhost: { background: "transparent", color: BRAND.amber, border: `1.5px solid ${BRAND.amber}`, borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 },
  tag: (bg, fg) => ({ display: "inline-block", background: bg, color: fg, borderRadius: 5, fontSize: 11, fontWeight: 700, padding: "2px 7px", marginRight: 4 }),
  tabBtn: (active) => ({
    padding: "10px 18px", border: "none", cursor: "pointer", fontWeight: 600, fontSize: 14,
    background: active ? BRAND.dark : "transparent", color: active ? BRAND.cream : BRAND.muted,
    borderRadius: "8px 8px 0 0",
  }),
  thBase: { padding: "7px 8px", fontSize: 11, fontWeight: 700, color: BRAND.muted, textTransform: "uppercase", letterSpacing: "0.03em", borderBottom: "2px solid #E4DDD2", whiteSpace: "nowrap", position: "sticky", top: 0, background: "#fff", zIndex: 2 },
};

function CheckboxDrop({ label, options, selected, setSelected }) {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef(null);
  React.useEffect(() => {
    if (!open) return;
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);
  const toggle = (v) => setSelected(prev => {
    const next = new Set(prev);
    next.has(v) ? next.delete(v) : next.add(v);
    return next;
  });
  const count = selected.size;
  return (
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={() => setOpen(p => !p)} style={{
        padding:"7px 12px", borderRadius:8, border:"1px solid #D8D0C2",
        background: count ? BRAND.amber : "#fff", color: count ? "#fff" : BRAND.dark,
        fontSize:13, cursor:"pointer", display:"flex", alignItems:"center", gap:6, fontWeight: count ? 600 : 400,
      }}>
        {label}{count ? ` (${count})` : ""} <span style={{fontSize:10}}>▾</span>
      </button>
      {open && (
        <div style={{
          position:"absolute", top:"calc(100% + 4px)", left:0, zIndex:200,
          background:"#fff", border:"1px solid #D8D0C2", borderRadius:10,
          boxShadow:"0 4px 16px rgba(0,0,0,0.12)", padding:"6px 0", minWidth:180, maxHeight:280, overflowY:"auto",
        }}>
          {count > 0 && (
            <button onClick={() => setSelected(new Set())} style={{
              width:"100%", textAlign:"left", padding:"5px 14px", background:"none", border:"none",
              fontSize:12, color:BRAND.amber, cursor:"pointer", borderBottom:"1px solid #EEE9E2", marginBottom:4,
            }}>Limpar filtro</button>
          )}
          {options.map(opt => (
            <label key={opt} style={{display:"flex", alignItems:"center", gap:8, padding:"5px 14px", cursor:"pointer", fontSize:13, color:BRAND.dark}}>
              <input type="checkbox" checked={selected.has(opt)} onChange={() => toggle(opt)}
                style={{width:14, height:14, accentColor:BRAND.amber, cursor:"pointer"}} />
              {opt}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

export default function PCPTorteleWeb() {
  const [tab, setTab] = useState("pcp");
  const [files, setFiles] = useState([]);
  const [estoqueLoja, setEstoqueLoja] = useState({});
  const [estoqueInfo, setEstoqueInfo] = useState(null);
  const [estoqueArqs, setEstoqueArqs] = useState([]); // [{name, loja, data:{cod:qty}}]
  const [categorias, setCategorias] = useState({});
  const [catInfo, setCatInfo] = useState(null);
  const [combos, setCombos] = useState([]);
  const [comboInfo, setComboInfo] = useState(null);
  const [usarCombos, setUsarCombos] = useState(true);
  const [fichasTec, setFichasTec] = useState(null);
  const [ftInfo, setFtInfo] = useState(null);
  const [subDetalhado, setSubDetalhado] = useState(false);
  const [comprasDetalhado, setComprasDetalhado] = useState(false);
  const [subSearch, setSubSearch] = useState("");
  const [comprasSearch, setComprasSearch] = useState("");
  const [nivelServico, setNivelServico] = useState(0.9);
  const [janela, setJanela] = useState(4);
  const [dataRef, setDataRef] = useState(() => new Date().toISOString().slice(0, 10));
  const [progEdits, setProgEdits] = useState({});
  const [progResetKey, setProgResetKey] = useState(0);
  const [progWeekStart, setProgWeekStart] = useState(() => {
    const today = new Date(); const dow = (today.getDay()+6)%7;
    return addDays(today, dow===0?0:7-dow).toISOString().slice(0,10);
  });
  const refEstArq = useRef(null);

  const [busca, setBusca] = useState("");
  const [filtABCs, setFiltABCs] = useState(new Set());
  const [filtCats, setFiltCats] = useState(new Set());
  const [filtSetores, setFiltSetores] = useState(new Set());
  const [rowResetKeys, setRowResetKeys] = useState({});
  const [soAlertas, setSoAlertas] = useState(false);
  const [sortKey, setSortKey] = useState("vol4");
  const [sortDesc, setSortDesc] = useState(true);
  const [lojaEnvio, setLojaEnvio] = useState("ALDEOTA");
  const [erro, setErro] = useState(null);
  const [saveStatus, setSaveStatus] = useState("");
  const [loadedFromCache, setLoadedFromCache] = useState(false);
  const refVendas = useRef(null), refEst = useRef(null), refCat = useRef(null), refCombo = useRef(null), refFT = useRef(null);

  /* carregar da sessão anterior */
  useEffect(() => {
    (async () => {
      const st = await loadState();
      if (st && st.files?.length) {
        setFiles(st.files);
        if (st.estoqueLoja) setEstoqueLoja(st.estoqueLoja);
        if (st.estoqueInfo) setEstoqueInfo(st.estoqueInfo);
        if (st.estoqueArqs) setEstoqueArqs(st.estoqueArqs);
        if (st.categorias) setCategorias(st.categorias);
        if (st.combos) setCombos(st.combos);
        if (st.fichasTec) { setFichasTec(st.fichasTec); setFtInfo(st.ftInfo || null); }
        if (st.params) {
          setNivelServico(st.params.nivelServico ?? 0.9);
          setJanela(st.params.janela ?? 4);
          if (st.params.dataRef) setDataRef(st.params.dataRef);
          setUsarCombos(st.params.usarCombos ?? true);
        }
        if (Object.keys(st.categorias || {}).length) setCatInfo(`${Object.keys(st.categorias).length} produtos (sessão anterior)`);
        if ((st.combos || []).length) setComboInfo(`${st.combos.length} linhas (sessão anterior)`);
        if (st.progEdits && Object.keys(st.progEdits).length) setProgEdits(st.progEdits);
        if (st.progWeekStart) setProgWeekStart(st.progWeekStart);
        setLoadedFromCache(true);
      }
    })();
  }, []);

  const detectLoja = (name) => {
    const n = name.toLowerCase();
    if (n.includes("aldeota")) return "ALDEOTA";
    if (n.includes("meireles")) return "MEIRELES";
    if (n.includes("rio")) return "RIO MAR";
    if (n.includes("sul")) return "SUL";
    return "";
  };

  const handleVendas = useCallback(async (fl) => {
    setErro(null);
    for (const f of Array.from(fl)) {
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
        const { rows, warnings } = parseVendas(wb);
        if (!rows.length) { setErro(`"${f.name}": nenhuma venda reconhecida.`); continue; }
        setFiles((p) => [...p, { name: f.name, loja: detectLoja(f.name), rows, warnings }]);
      } catch (e) { setErro(`Falha em "${f.name}": ${e.message}`); }
    }
  }, []);

  const handleEstoque = useCallback(async (fl) => {
    const f = fl[0]; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const { map, formato } = parseEstoquePorLoja(wb);
      const n = Object.keys(map).length;
      if (!n) { setErro("Estoque: nenhum produto reconhecido."); return; }
      setEstoqueLoja(map);
      setEstoqueInfo(`${f.name} — ${n} produtos, formato ${formato}`);
    } catch (e) { setErro(`Estoque: ${e.message}`); }
  }, []);

  const handleEstoqueArq = useCallback(async (fl) => {
    setErro(null);
    for (const f of Array.from(fl)) {
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
        const data = parseEstoqueSimples(wb);
        const n = Object.keys(data).length;
        if (!n) { setErro(`"${f.name}": nenhum produto de estoque reconhecido.`); continue; }
        const loja = normLoja(f.name) || "";
        setEstoqueArqs((prev) => [...prev, { name: f.name, loja, data }]);
      } catch (e) { setErro(`Estoque "${f.name}": ${e.message}`); }
    }
  }, []);

  const handleCat = useCallback(async (fl) => {
    const f = fl[0]; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const map = parseCategorias(wb);
      const n = Object.keys(map).length;
      if (!n) { setErro("Categorias: colunas Código+Categoria não encontradas."); return; }
      setCategorias(map); setCatInfo(`${f.name} — ${n} produtos`);
    } catch (e) { setErro(`Categorias: ${e.message}`); }
  }, []);

  const handleCombo = useCallback(async (fl) => {
    const f = fl[0]; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const list = parseCombos(wb);
      if (!list.length) { setErro("Combos: nenhuma linha válida (Cod Combo | ... | Cod Componente | ... | Qde)."); return; }
      setCombos(list); setComboInfo(`${f.name} — ${list.length} linhas, ${new Set(list.map(c=>c.codCombo)).size} combos`);
    } catch (e) { setErro(`Combos: ${e.message}`); }
  }, []);

  const handleFT = useCallback(async (fl) => {
    const f = fl[0]; if (!f) return;
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const data = parseFichasTecnicas(wb);
      const nFT = Object.keys(data.ft).length;
      const nSF = data.sfCodes.size;
      if (!nFT) { setErro("Fichas Técnicas: nenhuma ficha reconhecida. Verifique se o arquivo é o export de fichas técnicas do Izzyway (XLS/XLSX com blocos de produto e seção 'Itens da composição')."); return; }
      setFichasTec(data); setFtInfo(`${f.name} — ${nFT} fichas, ${nSF} subprodutos detectados`);
    } catch (e) { setErro(`Fichas Técnicas: ${e.message}`); }
  }, []);

  const allSales = useMemo(() => {
    const out = [];
    for (const f of files) if (f.loja) for (const r of f.rows) out.push(r.loja ? r : { ...r, loja: f.loja });
    return out;
  }, [files]);

  const semLoja = files.some((f) => !f.loja);

  // Combina estoque do arquivo único (multi-loja) com arquivos por loja
  const estoqueEfetivo = useMemo(() => {
    if (!estoqueArqs.length) return estoqueLoja;
    const merged = {};
    for (const [cod, lojaMap] of Object.entries(estoqueLoja)) {
      merged[+cod] = { ...lojaMap };
    }
    for (const arq of estoqueArqs) {
      if (!arq.loja) continue;
      for (const [cod, qty] of Object.entries(arq.data)) {
        const c = +cod;
        if (!merged[c]) merged[c] = {};
        merged[c][arq.loja] = (merged[c][arq.loja] || 0) + qty;
      }
    }
    return merged;
  }, [estoqueLoja, estoqueArqs]);

  const result = useMemo(() => {
    if (!allSales.length) return null;
    return computeAll(allSales, { nivelServico, janela, dataRef: new Date(dataRef + "T12:00:00"), usarCombos }, estoqueEfetivo, categorias, combos);
  }, [allSales, nivelServico, janela, dataRef, usarCombos, estoqueEfetivo, categorias, combos]);

  // Sincroniza semana de produção com a semana parcial do resultado
  useEffect(() => {
    if (result?.partialWeekStart) setProgWeekStart(result.partialWeekStart.toISOString().slice(0,10));
  }, [result]);

  const progWeekDates = useMemo(() => {
    const start = new Date(progWeekStart + "T12:00:00");
    return Array.from({length:14}, (_,i) => addDays(start, i));
  }, [progWeekStart]);

  const cats = useMemo(() => result ? [...new Set(result.rows.map((r)=>r.categoria))].sort() : [], [result]);
  const setores = useMemo(() => result ? [...new Set(result.rows.map((r)=>r.setor).filter(Boolean))].sort() : [], [result]);

  const visRows = useMemo(() => {
    if (!result) return [];
    let r = result.rows;
    if (busca) { const b = busca.toLowerCase(); r = r.filter((x)=>String(x.cod).includes(b)||x.produto.toLowerCase().includes(b)); }
    if (filtABCs.size) r = r.filter((x)=>filtABCs.has(x.abc));
    if (filtCats.size) r = r.filter((x)=>filtCats.has(x.categoria));
    if (filtSetores.size) r = r.filter((x)=>filtSetores.has(x.setor));
    if (soAlertas) r = r.filter((x)=>x.alertas.length>0);
    r = [...r].sort((a,b)=>{ const va=a[sortKey]??0, vb=b[sortKey]??0;
      if (typeof va==="string") return sortDesc?vb.localeCompare(va):va.localeCompare(vb);
      return sortDesc?vb-va:va-vb; });
    return r;
  }, [result, busca, filtABCs, filtCats, filtSetores, soAlertas, sortKey, sortDesc]);

  const totais = useMemo(() => result ? {
    produtos: result.rows.length,
    sugerida: result.rows.reduce((a,r)=>a+r.sugerida,0),
    liquida: result.rows.reduce((a,r)=>a+r.liquida,0),
    estoque: result.rows.reduce((a,r)=>a+r.estoque,0),
    alertas: result.rows.filter((r)=>r.alertas.length).length,
  } : null, [result]);

  const subprodData = useMemo(() => {
    if (!result || !fichasTec) return null;
    const map = {};
    for (const r of result.rows) {
      if (!r.liquida) continue;
      for (const e of (fichasTec.ft[r.cod] || [])) {
        if (!fichasTec.sfCodes.has(e.codInsumo)) continue;
        if (!map[e.codInsumo]) map[e.codInsumo] = { nome: e.nome, total: 0, rows: [] };
        const subtotal = e.quant * r.liquida;
        map[e.codInsumo].total += subtotal;
        map[e.codInsumo].rows.push({ produto: r.produto || `cod ${r.cod}`, codProd: r.cod, ftQuant: e.quant, liquida: r.liquida, subtotal });
      }
    }
    return Object.entries(map)
      .map(([cod, d]) => ({ cod: +cod, nome: d.nome, total: d.total, rows: d.rows.sort((a,b)=>b.subtotal-a.subtotal) }))
      .sort((a, b) => b.total - a.total);
  }, [result, fichasTec]);

  const comprasData = useMemo(() => {
    if (!result || !fichasTec) return null;
    const { ft, sfCodes } = fichasTec;
    const map = {};
    function expand(codInsumo, nome, ftFactor, liquida, via, depth) {
      if (depth > 6) return;
      if (sfCodes.has(codInsumo) && ft[codInsumo]?.length) {
        for (const sub of ft[codInsumo])
          expand(sub.codInsumo, sub.nome, ftFactor * sub.quant, liquida, `${via} → ${nome}`, depth + 1);
      } else {
        const subtotal = ftFactor * liquida;
        if (!map[codInsumo]) map[codInsumo] = { nome, total: 0, rows: [] };
        map[codInsumo].total += subtotal;
        map[codInsumo].rows.push({ via, ftFactor, liquida, subtotal });
      }
    }
    for (const r of result.rows) {
      if (!r.liquida) continue;
      const prod = r.produto || `cod ${r.cod}`;
      for (const e of (ft[r.cod] || []))
        expand(e.codInsumo, e.nome, e.quant, r.liquida, prod, 0);
    }
    return Object.entries(map)
      .map(([cod, d]) => ({ cod: +cod, nome: d.nome, total: d.total, rows: d.rows.sort((a,b)=>b.subtotal-a.subtotal) }))
      .sort((a, b) => b.total - a.total);
  }, [result, fichasTec]);

  const doSave = async () => {
    setSaveStatus("salvando…");
    const ok = await saveState({ files, estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos,
      params: { nivelServico, janela, dataRef, usarCombos }, progEdits, progWeekStart, fichasTec, ftInfo });
    setSaveStatus(ok ? "✓ salvo" : "falha ao salvar");
    setTimeout(()=>setSaveStatus(""), 3000);
  };
  const doClear = async () => {
    await clearState();
    setEstoqueArqs([]);
    setSaveStatus("dados da sessão apagados");
    setTimeout(()=>setSaveStatus(""),3000);
  };

  const th = (label, key, extra={}) => (
    <th onClick={()=>{ if(key){ setSortKey(key); setSortDesc(sortKey===key?!sortDesc:true);} }}
      style={{...S.thBase, textAlign: extra.left?"left":"right", cursor:key?"pointer":"default",
        background: sortKey===key?"#FDF6EC":"#fff", ...extra.style}}>
      {label}{sortKey===key?(sortDesc?" ↓":" ↑"):""}
    </th>
  );
  const td = (v, extra={}) => (
    <td title={extra.title} style={{ padding:"5px 8px", textAlign: extra.left?"left":"right", whiteSpace:"nowrap", ...extra.style }}>{v}</td>
  );
  const num = (v, dec=0) => v==null?"·":(+v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec});

  return (
    <div style={S.page}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
        * { box-sizing: border-box; } input, select, button { font-family: inherit; }
        tbody tr:hover { background: #FBF8F2; }
        ::-webkit-scrollbar { height:10px; width:10px; } ::-webkit-scrollbar-thumb { background:#D8D0C2; border-radius:6px; }
      `}</style>

      {/* HEADER */}
      <div style={{ background:"#3C2008", color:"#F0DBBF", padding:"0 26px", display:"flex", alignItems:"center", gap:20, flexWrap:"wrap", minHeight:68, borderBottom:"3px solid #B96A1B" }}>
        <img src="/tortele-logo.png" alt="Tortelê" style={{ height:44, objectFit:"contain", display:"block" }} />
        <div style={{ width:1, height:36, background:"rgba(240,219,191,0.25)" }} />
        <div style={{ fontSize:13, fontWeight:600, color:"#F0DBBF", letterSpacing:"0.04em", opacity:0.9 }}>
          PCP — Planejamento e Controle de Produção
        </div>
        {result && (
          <div style={{...S.mono, fontSize:12, color:"#D4B896", marginLeft:4}}>
            semana {fmtDM(result.lastWeekStart)}–{fmtDM(addDays(result.lastWeekStart,6))}
            {" · "}Z={result.Z}
            {result.nCombos>0 && usarCombos && " · combos"}
          </div>
        )}
        <div style={{ marginLeft:"auto", display:"flex", gap:10, alignItems:"center" }}>
          {saveStatus && <span style={{fontSize:12, color:"#E8C98C"}}>{saveStatus}</span>}
          <button style={{...S.btnGhost, borderColor:"#E8C98C", color:"#E8C98C", fontSize:12, padding:"7px 14px"}} onClick={doSave} disabled={!files.length}>Salvar sessão</button>
          <button style={{...S.btnGhost, borderColor:"rgba(240,219,191,0.3)", color:"#A08060", fontSize:11, padding:"6px 10px"}} onClick={doClear}>limpar</button>
        </div>
      </div>

      <div style={{ maxWidth: 1560, margin: "0 auto", padding: "20px 26px" }}>

        {loadedFromCache && (
          <div style={{ ...S.panel, background:"#FDF6EC", borderColor:"#E8C99A", fontSize:13, padding:"10px 16px" }}>
            ⚡ Dados restaurados da sessão anterior. Suba novas bases para atualizar ou continue de onde parou.
          </div>
        )}

        {/* CADASTROS / UPLOADS */}
        <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr", gap:12 }}>
          <div style={S.panel}>
            <div style={{ marginBottom:10 }}><span style={S.tag(BRAND.dark,BRAND.cream)}>1</span><b> Bases de vendas</b></div>
            <div onDragOver={(e)=>e.preventDefault()} onDrop={(e)=>{e.preventDefault();handleVendas(e.dataTransfer.files);}}
              onClick={()=>refVendas.current?.click()}
              style={{ border:"2px dashed #C9BFA9", borderRadius:8, padding:"16px", textAlign:"center", cursor:"pointer", background:"#FCFAF6", color:"#8A8073", fontSize:13 }}>
              Arraste as exportações (.xlsx, uma por loja) ou clique
              <input ref={refVendas} type="file" multiple accept=".xlsx,.xls" style={{display:"none"}}
                onChange={(e)=>{handleVendas(e.target.files);e.target.value="";}} />
            </div>
            {files.length>0 && (
              <div style={{ marginTop:10, fontSize:12 }}>
                {files.map((f,i)=>(
                  <div key={i}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 0", borderBottom: f.warnings?.length?"none":"1px solid #F0EBE2" }}>
                      <span style={{flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{f.name}</span>
                      <span style={{...S.mono, color:"#8A8073"}}>{f.rows.length.toLocaleString("pt-BR")}</span>
                      <select value={f.loja} onChange={(e)=>setFiles((p)=>p.map((x,k)=>k===i?{...x,loja:e.target.value}:x))}
                        style={{ padding:"3px 6px", borderRadius:5, border:f.loja?"1px solid #D8D0C2":"2px solid #C4501E", fontSize:12 }}>
                        <option value="">loja?</option>
                        {LOJAS.map((l)=><option key={l} value={l}>{l}</option>)}
                      </select>
                      <button onClick={()=>setFiles((p)=>p.filter((_,k)=>k!==i))}
                        style={{ border:"none", background:"none", color:"#C4501E", cursor:"pointer", fontSize:14 }}>×</button>
                    </div>
                    {f.warnings?.length > 0 && (
                      <div style={{paddingLeft:4, paddingBottom:4, fontSize:11, color:"#B96A1B", borderBottom:"1px solid #F0EBE2"}}>
                        {f.warnings.map((w,j)=><div key={j}>ℹ {w}</div>)}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
            {semLoja && <div style={{marginTop:8, color:"#C4501E", fontWeight:600, fontSize:12}}>⚠ Selecione a loja dos arquivos marcados.</div>}
          </div>

          {/* Painel Estoque */}
          <div style={S.panel}>
            <div style={{ marginBottom:8 }}><span style={S.tag(BRAND.dark,BRAND.cream)}>2</span><b> Estoque atual</b></div>
            {/* Opção A: arquivo único com todas as lojas */}
            <div onDragOver={(e)=>e.preventDefault()} onDrop={(e)=>{e.preventDefault();handleEstoque(e.dataTransfer.files);}}
              onClick={()=>refEst.current?.click()}
              style={{border:"2px dashed #C9BFA9",borderRadius:8,padding:"10px 12px",textAlign:"center",cursor:"pointer",background:"#FCFAF6",color:"#8A8073",fontSize:12}}>
              Arraste ou clique — todas as lojas
              <input ref={refEst} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                onChange={(e)=>{handleEstoque(e.target.files);e.target.value="";}} />
            </div>
            {estoqueInfo
              ? <div style={{marginTop:6, fontSize:12, color:"#2D6A4F", fontWeight:600}}>✓ {estoqueInfo}
                  <button onClick={()=>{setEstoqueLoja({});setEstoqueInfo(null);}} style={{border:"none",background:"none",color:"#C4501E",cursor:"pointer",marginLeft:6,fontSize:11}}>limpar</button>
                </div>
              : <div style={{marginTop:6, fontSize:11, color:"#8A8073"}}>Formato blocos (Cod|Produto|Loja…) ou colunar (Cod|…|ALDEOTA|…|CD).</div>}
            {/* Opção B: um arquivo por loja (exportação direta do sistema) */}
            <div style={{borderTop:"1px solid #F0EBE2", marginTop:10, paddingTop:8}}>
              <div style={{fontSize:11, color:"#6B6153", fontWeight:600, marginBottom:6}}>
                Ou suba por loja (exportação direta do sistema):
              </div>
              <div onDragOver={(e)=>e.preventDefault()} onDrop={(e)=>{e.preventDefault();handleEstoqueArq(e.dataTransfer.files);}}
                onClick={()=>refEstArq.current?.click()}
                style={{border:"2px dashed #C9BFA9",borderRadius:8,padding:"8px 12px",textAlign:"center",cursor:"pointer",background:"#FCFAF6",color:"#8A8073",fontSize:11}}>
                Arraste ou clique — por loja (vários)
                <input ref={refEstArq} type="file" multiple accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={(e)=>{handleEstoqueArq(e.target.files);e.target.value="";}} />
              </div>
              {estoqueArqs.length > 0 && (
                <div style={{marginTop:6, fontSize:12}}>
                  {estoqueArqs.map((arq, i) => (
                    <div key={i} style={{display:"flex", alignItems:"center", gap:6, padding:"3px 0", borderBottom:"1px solid #F0EBE2"}}>
                      <span style={{flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontSize:11}}>{arq.name}</span>
                      <span style={{...S.mono, color:"#8A8073", fontSize:11}}>{Object.keys(arq.data).length}p</span>
                      <select value={arq.loja} onChange={(e)=>setEstoqueArqs((prev)=>prev.map((x,k)=>k===i?{...x,loja:e.target.value}:x))}
                        style={{padding:"2px 5px", borderRadius:4, border:arq.loja?"1px solid #D8D0C2":"2px solid #C4501E", fontSize:11}}>
                        <option value="">loja?</option>
                        {[...LOJAS,"CD"].map((l)=><option key={l} value={l}>{l}</option>)}
                      </select>
                      <button onClick={()=>setEstoqueArqs((prev)=>prev.filter((_,k)=>k!==i))}
                        style={{border:"none",background:"none",color:"#C4501E",cursor:"pointer",fontSize:13}}>×</button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          {[
            { n:"3", t:"Categorias", info: catInfo, ref: refCat, handler: handleCat,
              desc:"Suba a AUX CATEGORIA (colunas Código + Categoria). Pode ser a própria planilha V3.",
              clear: ()=>{setCategorias({});setCatInfo(null);} },
            { n:"4", t:"Combos", info: comboInfo, ref: refCombo, handler: handleCombo,
              desc:"Suba a AUX COMBOS (Cod Combo | Nome | Cod Componente | Nome | Qde).",
              clear: ()=>{setCombos([]);setComboInfo(null);} },
            { n:"5", t:"Fichas Técnicas", info: ftInfo, ref: refFT, handler: handleFT,
              desc:"Suba o export de Fichas Técnicas do Izzyway (XLS). Sub-produtos são detectados automaticamente. Habilita as abas Subprodutos e Sugestão de Compras.",
              clear: ()=>{setFichasTec(null);setFtInfo(null);} },
          ].map((b)=>(
            <div key={b.n} style={S.panel}>
              <div style={{ marginBottom:8 }}><span style={S.tag(BRAND.dark,BRAND.cream)}>{b.n}</span><b> {b.t}</b></div>
              <div onDragOver={(e)=>e.preventDefault()} onDrop={(e)=>{e.preventDefault();b.handler(e.dataTransfer.files);}}
                onClick={()=>b.ref.current?.click()}
                style={{border:"2px dashed #C9BFA9",borderRadius:8,padding:"12px",textAlign:"center",cursor:"pointer",background:"#FCFAF6",color:"#8A8073",fontSize:12}}>
                Arraste ou clique para subir
                <input ref={b.ref} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={(e)=>{b.handler(e.target.files);e.target.value="";}} />
              </div>
              {b.info
                ? <div style={{marginTop:8, fontSize:12, color:"#2D6A4F", fontWeight:600}}>✓ {b.info}
                    <button onClick={b.clear} style={{border:"none",background:"none",color:"#C4501E",cursor:"pointer",marginLeft:6,fontSize:11}}>limpar</button>
                  </div>
                : <div style={{marginTop:8, fontSize:11, color:"#8A8073"}}>{b.desc}</div>}
              {b.n==="4" && combos.length>0 && (
                <label style={{fontSize:12, display:"flex", gap:6, marginTop:6, cursor:"pointer"}}>
                  <input type="checkbox" checked={usarCombos} onChange={(e)=>setUsarCombos(e.target.checked)} />
                  Aplicar ajuste de combos
                </label>
              )}
            </div>
          ))}
        </div>

        {erro && (
          <div style={{...S.panel, background:"#FEF2F2", border:"1px solid #FCA5A5", color:"#991B1B", display:"flex", gap:12, alignItems:"center", padding:"12px 16px"}}>
            <span style={{fontSize:18}}>⚠</span>
            <span style={{flex:1, fontSize:13}}>{erro}</span>
            <button onClick={()=>setErro(null)} style={{border:"none", background:"none", color:"#991B1B", cursor:"pointer", fontSize:18, fontWeight:700, lineHeight:1}}>×</button>
          </div>
        )}

        {/* PARÂMETROS */}
        <div style={{ ...S.panel, display:"flex", gap:18, alignItems:"flex-end", flexWrap:"wrap" }}>
          <label style={{fontSize:12, color:"#6B6153", fontWeight:600}}>Nível de serviço<br/>
            <select value={nivelServico} onChange={(e)=>setNivelServico(+e.target.value)}
              style={{marginTop:4, padding:"7px 10px", borderRadius:6, border:"1px solid #D8D0C2", fontSize:14}}>
              <option value={0.8}>80% (Z 0,84)</option><option value={0.9}>90% (Z 1,28)</option>
              <option value={0.95}>95% (Z 1,65)</option><option value={0.99}>99% (Z 2,33)</option>
            </select></label>
          <label style={{fontSize:12, color:"#6B6153", fontWeight:600}}>Janela (semanas)<br/>
            <select value={janela} onChange={(e)=>setJanela(+e.target.value)}
              style={{marginTop:4, padding:"7px 10px", borderRadius:6, border:"1px solid #D8D0C2", fontSize:14}}>
              {[2,3,4,5,6,7,8].map((n)=><option key={n} value={n}>{n}</option>)}
            </select></label>
          <label style={{fontSize:12, color:"#6B6153", fontWeight:600}}>Data de referência<br/>
            <input type="date" value={dataRef} onChange={(e)=>setDataRef(e.target.value)}
              style={{marginTop:4, padding:"6px 10px", borderRadius:6, border:"1px solid #D8D0C2", fontSize:14}}/></label>
          {result && <button style={{...S.btn, marginLeft:"auto"}} onClick={()=>exportAll(result)}>⬇ Exportar tudo (Excel)</button>}
        </div>

        {result && totais && (
          <>
            {/* CARDS */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:10, marginBottom:14 }}>
              {[["Produtos",num(totais.produtos),BRAND.dark],["Sugerida total",num(totais.sugerida),BRAND.amber],
                ["Estoque abatido",num(totais.estoque),"#6B6153"],["Produção LÍQUIDA",num(totais.liquida),"#2D6A4F"],
                ["Com alerta",num(totais.alertas),"#C4501E"]].map(([l,v,c])=>(
                <div key={l} style={{...S.panel, marginBottom:0, padding:"12px 16px"}}>
                  <div style={{fontSize:10, color:"#8A8073", textTransform:"uppercase", fontWeight:700}}>{l}</div>
                  <div style={{...S.mono, fontSize:24, fontWeight:600, color:c}}>{v}</div>
                </div>
              ))}
            </div>

            {/* TABS */}
            <div style={{ display:"flex", gap:2, borderBottom:`2px solid ${BRAND.dark}` }}>
              {[["pcp","PCP Semanal"],["prog","Programação"],["loja","PCP por Loja"],["cmv","Auditoria CMV"],
                ...(fichasTec?[["sub","Subprodutos"],["compras","Sugestão de Compras"]]:[])]
                .map(([k,l])=>(
                <button key={k} style={S.tabBtn(tab===k)} onClick={()=>setTab(k)}>{l}</button>
              ))}
            </div>

            {/* FILTROS */}
            <div style={{ ...S.panel, borderRadius:"0 0 10px 10px", borderTop:"none", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
              <input placeholder="Buscar código ou produto…" value={busca} onChange={(e)=>setBusca(e.target.value)}
                style={{padding:"8px 12px", borderRadius:8, border:"1px solid #D8D0C2", width:220, fontSize:13}}/>
              {setores.length > 0 && (
                <CheckboxDrop label="Setor" options={setores} selected={filtSetores} setSelected={setFiltSetores} />
              )}
              <CheckboxDrop label="Categoria" options={cats} selected={filtCats} setSelected={setFiltCats} />
              <CheckboxDrop label="ABC" options={["A","B","C"]} selected={filtABCs} setSelected={setFiltABCs} />
              <label style={{fontSize:13, display:"flex", gap:6, alignItems:"center", cursor:"pointer"}}>
                <input type="checkbox" checked={soAlertas} onChange={(e)=>setSoAlertas(e.target.checked)}/>Só com alerta
              </label>
              <span style={{color:"#8A8073", fontSize:12}}>{visRows.length} produtos</span>
              {tab==="loja" && (
                <label style={{marginLeft:"auto", fontSize:12, color:"#6B6153", fontWeight:600}}>Loja p/ envio diário:{" "}
                  <select value={lojaEnvio} onChange={(e)=>setLojaEnvio(e.target.value)}
                    style={{padding:"7px 10px", borderRadius:8, border:"1px solid #D8D0C2", fontSize:13}}>
                    {LOJAS.map((l)=><option key={l} value={l}>{l}</option>)}
                  </select></label>
              )}
            </div>

            {/* ===== TAB: PCP SEMANAL ===== */}
            {tab==="pcp" && (
              <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"64vh"}}>
                <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5, minWidth:1750}}>
                  <thead><tr>
                    {th("Cod","cod")}{th("Produto","produto",{left:true})}{th("Categ.","categoria",{left:true})}{th("Setor","setor",{left:true})}
                    {result.weekStarts.map((w)=>th(fmtDM(w),null))}
                    {th("Parc.",null)}{th("Combo","combo")}
                    {th("Méd","media")}{th("DP","dp")}{th("CV","cv")}{th("Mín","min")}{th("Máx","max")}{th("Tend","tend")}
                    {th("ES","es")}{th("Marg","margem")}{th("Sug.","sugerida")}{th("Estq","estoque")}
                    {th("LÍQ.","liquida",{style:{background:"#2D6A4F",color:"#fff"}})}
                    {th("ABC","abc")}
                    {DIAS.map((d)=>th(d,null))}{th("Tot",null)}
                    {th("Alertas",null,{left:true})}
                  </tr></thead>
                  <tbody style={S.mono}>
                    {visRows.slice(0,400).map((r)=>(
                      <tr key={r.cod} style={{borderBottom:"1px solid #F0EBE2"}}>
                        {td(r.cod,{style:{color:"#8A8073"}})}
                        <td title={r.produto||`cod ${r.cod}`} style={{padding:"5px 8px",textAlign:"left",whiteSpace:"nowrap"}}>
                          <div style={{width:210,overflow:"hidden",textOverflow:"ellipsis",fontFamily:"'Inter',sans-serif",color:BRAND.dark}}>
                            {r.produto||<span style={{color:"#B0A794",fontStyle:"italic"}}>cod {r.cod}</span>}
                          </div>
                        </td>
                        {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:"#6B6153",maxWidth:110,overflow:"hidden",textOverflow:"ellipsis"}})}
                        <td style={{padding:"4px 8px",whiteSpace:"nowrap"}}>
                          {r.setor
                            ? <span style={{fontSize:11,fontWeight:600,padding:"2px 7px",borderRadius:3,
                                background:isSalgado(r.categoria)?"#FFF3E0":"#FCE4EC",
                                color:isSalgado(r.categoria)?"#B96A1B":"#AD1457"}}>
                                {r.setor}
                              </span>
                            : <span style={{color:"#C0B9B0",fontSize:11}}>—</span>}
                        </td>
                        {r.weeks.map((v,i)=>td(v?num(v):"·",{style:{color:i===7?BRAND.dark:"#8A8073",fontWeight:i===7?600:400}}))}
                        {td(r.partial?num(r.partial):"·",{style:{color:"#B0A794"}})}
                        {td(r.combo?num(r.combo,1):"·",{style:{background:r.combo?"#FFF3D6":"transparent"}})}
                        {td(num(r.media))}{td(num(r.dp),{style:{color:"#8A8073"}})}
                        {td(`${num(r.cv*100)}%`,{style:{color:r.cv>0.5?"#C4501E":"#8A8073"}})}
                        {td(num(r.min),{style:{color:"#8A8073"}})}{td(num(r.max),{style:{color:"#8A8073"}})}
                        {td(`${r.tend>=0?"+":""}${num(r.tend*100)}%`,{style:{color:r.tend<-0.4?"#C4501E":r.tend>0.5?"#2D6A4F":"#8A8073"}})}
                        {td(num(r.es),{style:{color:"#8A8073"}})}
                        {td(`${num(r.margem*100)}%`,{style:{color:"#8A8073"}})}
                        {td(num(r.sugerida),{style:{fontWeight:600,color:"#B96A1B"}})}
                        {td(r.estoque?num(r.estoque):"·",{style:{color:"#8A8073"}})}
                        {td(num(r.liquida),{style:{fontWeight:700,color:"#fff",background:"#2D6A4F"}})}
                        {td(r.abc,{style:{textAlign:"center",fontWeight:700}})}
                        {r.liqDia.map((v,i)=>td(v||"·",{style:{color:"#4A6B5A",background:"#F2F7F3"}}))}
                        {td(num(r.liqDia.reduce((a,b)=>a+b,0)),{style:{fontWeight:600,color:"#2D6A4F"}})}
                        {td(r.alertas.map((a)=><span key={a} style={S.tag("#FCE4D6","#C4501E")}>{a}</span>),{left:true,style:{fontFamily:"'Inter',sans-serif"}})}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {visRows.length>400 && <div style={{padding:10,textAlign:"center",color:"#8A8073",fontSize:12}}>Exibindo 400 de {visRows.length} — refine ou exporte.</div>}
              </div>
            )}

            {/* ===== TAB: PROGRAMAÇÃO ===== */}
            {tab==="prog" && (
              <>
                {/* Controles da semana */}
                <div style={{...S.panel, display:"flex", gap:18, alignItems:"center", flexWrap:"wrap", padding:"12px 18px"}}>
                  <label style={{fontSize:12, fontWeight:600, color:BRAND.muted, display:"flex", alignItems:"center", gap:8}}>
                    Início do período (2 semanas):
                    <input type="date" value={progWeekStart} onChange={(e)=>setProgWeekStart(e.target.value)}
                      style={{padding:"5px 10px", borderRadius:6, border:"1px solid #D8D0C2", fontSize:13}} />
                  </label>
                  <div style={{fontSize:12, color:BRAND.muted}}>
                    {fmtDM(progWeekDates[0])} (Seg) — {fmtDM(progWeekDates[13])} (Dom)
                  </div>
                  <div style={{marginLeft:"auto", display:"flex", gap:10, alignItems:"center"}}>
                    <span style={{fontSize:11, color:"#8A8073"}}>
                      vnd = venda prevista · campo editável = produção · número colorido = estoque projetado (verde ≥ ES, vermelho &lt; ES)
                    </span>
                    <button style={{...S.btn, fontSize:12, padding:"7px 14px"}}
                      onClick={()=>{
                        const wk0 = addDays(new Date(progWeekStart+"T12:00:00"), 0).toISOString().slice(0,10);
                        const wk1 = addDays(new Date(progWeekStart+"T12:00:00"), 7).toISOString().slice(0,10);
                        const hdr = ["Cód","Produto","Categoria","Setor","Est. Atual","ES","Líq/sem",
                          ...progWeekDates.map(d=>fmtDM(d)+" "+DIAS[(d.getDay()+6)%7]),"Total 2sem","Saldo"];
                        const body = visRows.slice(0,400).map(r=>{
                          const vals = progWeekDates.map((_,di)=>{
                            const wk = Math.floor(di/7); const dow = di%7;
                            return progEdits[`${r.cod}_${wk===0?wk0:wk1}_${dow}`]??(r.prog[dow]??0);
                          });
                          const tot = vals.reduce((s,v)=>s+v,0);
                          return [r.cod,r.produto,r.categoria,r.setor??"-",r.estoque,Math.ceil(r.es),r.liquida,...vals,tot,tot-r.liquida*2];
                        });
                        const wb2 = XLSX.utils.book_new();
                        const ws = styledSheet([hdr,...body],
                          [7,36,16,8,9,7,9,...progWeekDates.map(()=>10),10,9],
                          `Programação ${progWeekStart}`);
                        ws["!freeze"] = {xSplit:4, ySplit:2};
                        // colorir colunas dos dias
                        const NF = 7;
                        for (let row = 2; row < body.length+2; row++) {
                          const alt = row%2===0;
                          const liqC = ws[XLSX.utils.encode_cell({r:row,c:6})];
                          if (liqC) liqC.s = {...XS.num(alt), font:{bold:true,color:{rgb:"2D6A4F"}}};
                          for (let col=NF; col<NF+14; col++) {
                            const cell = ws[XLSX.utils.encode_cell({r:row,c:col})];
                            if (!cell) continue;
                            const wk = Math.floor((col-NF)/7);
                            if (!cell.v) { cell.s = {...(alt?XS.rowA:XS.rowB), alignment:{horizontal:"center",vertical:"center"}}; continue; }
                            cell.s = wk===0
                              ? {fill:{patternType:"solid",fgColor:{rgb:"EDF1F8"}},font:{bold:true,color:{rgb:"264478"}},alignment:{horizontal:"center",vertical:"center"},numFmt:"#,##0"}
                              : {fill:{patternType:"solid",fgColor:{rgb:"E8F0E4"}},font:{bold:true,color:{rgb:"2D6A4F"}},alignment:{horizontal:"center",vertical:"center"},numFmt:"#,##0"};
                          }
                          const sC = ws[XLSX.utils.encode_cell({r:row,c:NF+15})];
                          if (sC && typeof sC.v==="number") sC.s = {...XS.num(alt), font:{bold:true,color:{rgb:sC.v>=0?"2D6A4F":"C4501E"}}};
                        }
                        XLSX.utils.book_append_sheet(wb2, ws, "Programação");
                        XLSX.writeFile(wb2, `Prog_Tortele_${progWeekStart}.xlsx`);
                      }}>⬇ Excel</button>
                    <button style={{...S.btnGhost, fontSize:11, padding:"6px 10px"}}
                      onClick={()=>{
                        const wk0 = addDays(new Date(progWeekStart+"T12:00:00"), 0).toISOString().slice(0,10);
                        const wk1 = addDays(new Date(progWeekStart+"T12:00:00"), 7).toISOString().slice(0,10);
                        const rows = visRows.slice(0,400);
                        let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Programação ${progWeekStart} — Tortelê</title>
<style>body{font-family:Arial,sans-serif;font-size:8.5px;color:#222;margin:6mm}
h2{font-size:11px;color:#3C2008;margin:0 0 5px}
table{border-collapse:collapse;width:100%}
th{background:#3C2008;color:#F0DBBF;padding:3px 3px;font-size:8px;white-space:nowrap;text-align:center}
td{border:1px solid #E4DDD2;padding:2px 3px;text-align:right;white-space:nowrap}
.nome{text-align:left;max-width:110px;overflow:hidden}
.s1{background:#EDF1F8;color:#264478;font-weight:700}
.s2{background:#E8F0E4;color:#2D6A4F;font-weight:700}
.liq{background:#2D6A4F;color:#fff;font-weight:700;text-align:right}
.ok{color:#2D6A4F;font-weight:700}.nok{color:#C4501E;font-weight:700}
.zero{color:#CCC}
tr:nth-child(even) td{background:#FAFAFA}
.s1,.s2,.liq{background-color:inherit!important}
@page{size:landscape;margin:6mm}</style></head><body>
<h2>Programação de Produção — Tortelê · ${progWeekStart} – ${addDays(new Date(progWeekStart+"T12:00:00"),13).toISOString().slice(0,10)}</h2>
<table><thead><tr>
<th>Cód</th><th>Produto</th><th>Cat</th><th>Líq/sem</th>
${progWeekDates.map((d,i)=>`<th class="${i<7?'s1':'s2'}">${fmtDM(d)}<br/>${DIAS[i%7]}</th>`).join("")}
<th>Total</th><th>Saldo</th></tr></thead><tbody>`;
                        for (const r of rows) {
                          const vals = progWeekDates.map((_,di)=>{
                            const wk=Math.floor(di/7),dow=di%7;
                            return progEdits[`${r.cod}_${wk===0?wk0:wk1}_${dow}`]??(r.prog[dow]??0);
                          });
                          const tot=vals.reduce((s,v)=>s+v,0);
                          const saldo=tot-r.liquida*2;
                          html+=`<tr><td style="color:#8A8073">${r.cod}</td><td class="nome">${r.produto||'-'}</td>
<td style="text-align:left;font-size:7.5px;color:#7A6450">${r.categoria||'-'}</td>
<td class="liq">${r.liquida.toLocaleString('pt-BR')}</td>
${vals.map((v,i)=>`<td class="${i<7?'s1':'s2'}${v===0?' zero':''}">${v||''}</td>`).join("")}
<td style="font-weight:700">${tot.toLocaleString('pt-BR')}</td>
<td class="${saldo>=0?'ok':'nok'}">${saldo>=0?'+':''}${saldo.toLocaleString('pt-BR')}</td></tr>`;
                        }
                        html+=`</tbody></table></body></html>`;
                        const w=window.open("","_blank");
                        w.document.write(html); w.document.close(); w.print();
                      }}>🖨 PDF</button>
                    <button style={{...S.btnGhost, fontSize:11, padding:"6px 10px"}}
                      onClick={()=>{ setProgEdits({}); setRowResetKeys({}); setProgResetKey(k=>k+1); }}>Resetar ajustes</button>
                    <button style={{...S.btnGhost, fontSize:11, padding:"6px 10px", color:"#C4501E", borderColor:"#E8C8C0"}}
                      onClick={()=>{
                        const wk0 = addDays(new Date(progWeekStart+"T12:00:00"), 0).toISOString().slice(0,10);
                        const wk1 = addDays(new Date(progWeekStart+"T12:00:00"), 7).toISOString().slice(0,10);
                        const next = {...progEdits};
                        for (const r of visRows) {
                          for (const wKey of [wk0, wk1]) {
                            for (let dow=0; dow<7; dow++) next[`${r.cod}_${wKey}_${dow}`] = 0;
                          }
                        }
                        setProgEdits(next);
                        setProgResetKey(k=>k+1);
                      }}>Zerar todos</button>
                  </div>
                </div>

                {/* Tabela interativa */}
                <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"62vh"}}>
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:11.5, minWidth:1900}}>
                    <thead>
                      <tr>
                        <th colSpan={5} style={{...S.thBase,background:"#fff",borderBottom:"none"}} />
                        <th colSpan={7} style={{background:"#EDF1F8",color:"#264478",textAlign:"center",fontSize:11,fontWeight:700,padding:"4px 3px",borderBottom:"2px solid #264478",position:"sticky",top:0,zIndex:2}}>
                          Semana 1 · {fmtDM(progWeekDates[0])} – {fmtDM(progWeekDates[6])}
                        </th>
                        <th colSpan={7} style={{background:"#E8F0E4",color:"#2D6A4F",textAlign:"center",fontSize:11,fontWeight:700,padding:"4px 3px",borderLeft:"3px solid #B96A1B",borderBottom:"2px solid #2D6A4F",position:"sticky",top:0,zIndex:2}}>
                          Semana 2 · {fmtDM(progWeekDates[7])} – {fmtDM(progWeekDates[13])}
                        </th>
                        <th colSpan={3} style={{...S.thBase,background:"#fff",borderBottom:"none"}} />
                      </tr>
                      <tr>
                        {th("Cód","cod")}
                        {th("Produto","produto",{left:true})}
                        {th("Cat.","categoria",{left:true})}
                        {th("Líq/sem","liquida",{style:{background:"#2D6A4F",color:"#fff"}})}
                        {th("Estq","estoque")}
                        {progWeekDates.map((d,i)=>(
                          <th key={i} style={{...S.thBase,
                            background: Math.floor(i/7)===0 ? "#EDF1F8" : "#E8F0E4",
                            color: Math.floor(i/7)===0 ? "#264478" : "#2D6A4F",
                            borderLeft: i===7 ? "3px solid #B96A1B" : undefined,
                            padding:"4px 3px", minWidth:72, textAlign:"center"}}>
                            {DIAS[i%7]}<br/>
                            <span style={{fontWeight:400,fontSize:10}}>{fmtDM(d)}</span>
                          </th>
                        ))}
                        {th("Total",null,{style:{fontWeight:700}})}
                        {th("Saldo 2sem",null)}
                        {th("",null,{style:{width:38}})}
                      </tr>
                    </thead>
                    <tbody style={S.mono}>
                      {(() => {
                        // Constantes computadas UMA VEZ fora do loop de 400 linhas
                        const _partialTs = result.partialWeekStart.getTime();
                        const _selTs = new Date(progWeekStart + "T12:00:00").getTime();
                        const _weeksDiff = Math.max(0, Math.round((_selTs - _partialTs) / (7 * 86400000)));
                        // Apenas 2 wKeys distintos (semana 1 e semana 2)
                        const progWKeys = [0, 7].map(off =>
                          addDays(new Date(progWeekStart + "T12:00:00"), off).toISOString().slice(0, 10)
                        );
                        return visRows.slice(0,400).map((r)=>{
                        let startEst = r.estoque || 0;
                        for (let w = 0; w < _weeksDiff; w++) {
                          const wKey = addDays(result.partialWeekStart, w * 7).toISOString().slice(0, 10);
                          let weekProd = 0;
                          for (let di = 0; di < 7; di++) weekProd += progEdits[`${r.cod}_${wKey}_${di}`] ?? (r.prog[di] ?? 0);
                          startEst = Math.round(startEst + weekProd - r.media);
                        }
                        let runEst = startEst;
                        const dayData = progWeekDates.map((d, di)=>{
                          const wk = Math.floor(di / 7);
                          const dow = di % 7;
                          const wKey = progWKeys[wk];
                          const prod = progEdits[`${r.cod}_${wKey}_${dow}`] ?? (r.prog[dow] ?? 0);
                          const venda = Math.ceil(r.mixDia[dow] * r.media);
                          runEst = Math.round(runEst + prod - venda);
                          return { prod, venda, estProj: runEst, wKey, dow };
                        });
                        const totalProg = dayData.reduce((s,d)=>s+d.prod,0);
                        const saldo = totalProg - r.liquida * 2;
                        return (
                          <tr key={r.cod} style={{borderBottom:"2px solid #F0EBE2"}}>
                            {td(r.cod,{style:{color:"#8A8073",fontSize:11}})}
                            <td title={r.produto||`cod ${r.cod}`} style={{padding:"5px 8px",textAlign:"left",whiteSpace:"nowrap"}}>
                          <div style={{width:190,overflow:"hidden",textOverflow:"ellipsis",fontFamily:"'Inter',sans-serif",color:BRAND.dark}}>
                            {r.produto||<span style={{color:"#B0A794",fontStyle:"italic"}}>cod {r.cod}</span>}
                          </div>
                        </td>
                            {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:10,color:isSalgado(r.categoria)?BRAND.amber:BRAND.muted}})}
                            {td(num(r.liquida),{style:{fontWeight:700,color:"#fff",background:"#2D6A4F"}})}
                            {td(r.estoque?num(r.estoque):"·",{style:{color:"#8A8073"}})}
                            {dayData.map((dd,di)=>{
                              const estOk = dd.estProj >= r.es;
                              const hasVal = dd.prod > 0;
                              const isWeekend = (di%7)===5||(di%7)===6;
                              return (
                                <td key={di} style={{
                                  padding:"3px 3px",
                                  background: isWeekend ? "#F5F3EE" : (Math.floor(di/7)===1 ? "#FAFDF8" : "transparent"),
                                  borderLeft: di===7 ? "3px solid #B96A1B" : undefined,
                                  textAlign:"center", minWidth:72}}>
                                  <div style={{fontSize:10,color:"#9A8E7F",marginBottom:2,fontFamily:"'Inter',sans-serif",fontWeight:500}}>
                                    vnd {dd.venda||0}
                                  </div>
                                  <input
                                    key={`${r.cod}_${dd.wKey}_${dd.dow}_r${rowResetKeys[r.cod]||0}_g${progResetKey}`}
                                    type="number" min="0"
                                    defaultValue={progEdits[`${r.cod}_${dd.wKey}_${dd.dow}`] !== undefined
                                      ? progEdits[`${r.cod}_${dd.wKey}_${dd.dow}`]
                                      : (r.prog[dd.dow] ?? "")}
                                    onBlur={(e)=>{
                                      const raw = e.target.value.trim();
                                      const v = raw === "" ? 0 : Math.max(0, parseInt(raw) || 0);
                                      const key = `${r.cod}_${dd.wKey}_${dd.dow}`;
                                      const cur = progEdits[key] ?? (r.prog[dd.dow] ?? 0);
                                      if (v !== cur) setProgEdits(p=>({...p,[key]:v}));
                                    }}
                                    style={{width:54,padding:"4px 2px",borderRadius:4,border:"1px solid",
                                      borderColor:hasVal?"#264478":"#D8D0C2",
                                      background:hasVal?"#EDF1F8":"#FAFAFA",
                                      color:hasVal?"#264478":"#8A8073",
                                      fontWeight:600,fontSize:13,fontFamily:"'JetBrains Mono',monospace",textAlign:"center"}}
                                  />
                                  <div style={{fontSize:11,color:estOk?"#2D6A4F":"#C4501E",fontWeight:700,marginTop:2,fontFamily:"'JetBrains Mono',monospace"}}>
                                    {num(dd.estProj)}
                                  </div>
                                </td>
                              );
                            })}
                            <td style={{padding:"5px 8px",textAlign:"right",fontWeight:700,color:saldo>=0?"#2D6A4F":"#C4501E"}}>
                              {num(totalProg)}
                            </td>
                            <td style={{padding:"5px 8px",textAlign:"right",fontWeight:600,fontSize:11,color:saldo>=0?"#2D6A4F":"#C4501E"}}>
                              {saldo>=0?`+${num(saldo)}`:num(saldo)}
                            </td>
                            <td style={{padding:"3px 6px",textAlign:"center"}}>
                              <button
                                title="Zerar produção desta linha nas 2 semanas"
                                onClick={()=>{
                                  const next = {...progEdits};
                                  progWKeys.forEach((wKey)=>{
                                    for (let dow=0; dow<7; dow++) next[`${r.cod}_${wKey}_${dow}`] = 0;
                                  });
                                  setProgEdits(next);
                                  // apenas os 14 inputs desta linha remontam
                                  setRowResetKeys(p=>({...p,[r.cod]:(p[r.cod]||0)+1}));
                                }}
                                style={{border:"1px solid #E4DDD2",background:"#FFF8F5",color:"#C4501E",
                                  borderRadius:4,padding:"2px 6px",fontSize:11,cursor:"pointer",
                                  fontWeight:700,fontFamily:"'JetBrains Mono',monospace",lineHeight:1.4}}>
                                ×0
                              </button>
                            </td>
                          </tr>
                        );
                      }); })()}
                    </tbody>
                  </table>
                </div>
              </>
            )}

            {/* ===== TAB: PCP POR LOJA ===== */}
            {tab==="loja" && (
              <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"64vh"}}>
                <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5, minWidth:1700}}>
                  <thead><tr>
                    {th("Cod","cod")}{th("Produto","produto",{left:true})}{th("Categ.","categoria",{left:true})}
                    {th("Sug.","sugerida")}{th("Estq","estoque")}{th("LÍQ.","liquida",{style:{background:"#2D6A4F",color:"#fff"}})}
                    {LOJAS.map((l)=>th(`${l.slice(0,3)} %`,null,{style:{color:"#8A8073"}}))}
                    {LOJAS.map((l)=>th(`Sug ${l.slice(0,3)}`,null,{style:{color:"#B96A1B"}}))}
                    {LOJAS.map((l)=>th(`Est ${l.slice(0,3)}`,null))}
                    {th("Est CD",null,{style:{color:"#6B6153"}})}
                    {LOJAS.map((l)=>th(`Líq ${l.slice(0,3)}`,null,{style:{background:"#2D6A4F",color:"#fff"}}))}
                    {DIAS.map((d)=>th(`${d} → ${lojaEnvio.slice(0,3)}`,null,{style:{background:"#264478",color:"#fff"}}))}
                  </tr></thead>
                  <tbody style={S.mono}>
                    {visRows.slice(0,400).map((r)=>{
                      const li = LOJAS.indexOf(lojaEnvio);
                      const envio = r.liqDia.map((v)=>Math.ceil(v*r.mixLoja[li]));
                      return (
                        <tr key={r.cod} style={{borderBottom:"1px solid #F0EBE2"}}>
                          {td(r.cod,{style:{color:"#8A8073"}})}
                          <td title={r.produto||`cod ${r.cod}`} style={{padding:"5px 8px",textAlign:"left",whiteSpace:"nowrap"}}>
                          <div style={{width:190,overflow:"hidden",textOverflow:"ellipsis",fontFamily:"'Inter',sans-serif",color:BRAND.dark}}>
                            {r.produto||<span style={{color:"#B0A794",fontStyle:"italic"}}>cod {r.cod}</span>}
                          </div>
                        </td>
                          {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:"#6B6153",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis"}})}
                          {td(num(r.sugerida),{style:{fontWeight:600,color:"#B96A1B"}})}
                          {td(r.estoque?num(r.estoque):"·",{style:{color:"#8A8073"}})}
                          {td(num(r.liquida),{style:{fontWeight:700,color:"#fff",background:"#2D6A4F"}})}
                          {r.mixLoja.map((m,i)=>td(`${num(m*100)}%`,{style:{color:"#B0A794",fontSize:11}}))}
                          {r.sugLoja.map((v,i)=>td(num(v),{style:{color:"#B96A1B"}}))}
                          {r.estPorLoja.map((v,i)=>td(v?num(v):"·",{style:{color:"#8A8073"}}))}
                          {td(r.estoquePorLocal[4]?num(r.estoquePorLocal[4]):"·",{style:{color:"#6B6153",background:"#F6F3EE"}})}
                          {r.liqLoja.map((v,i)=>td(num(v),{style:{fontWeight:600,color:"#2D6A4F",background:"#F2F7F3"}}))}
                          {envio.map((v,i)=>td(v||"·",{style:{color:"#264478",background:"#EDF1F8"}}))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* ===== TAB: SUBPRODUTOS ===== */}
            {tab==="sub" && fichasTec && (()=>{
              const sq = subSearch.trim().toLowerCase();
              const subVis = subprodData ? subprodData.filter(sp=>
                !sq || sp.nome.toLowerCase().includes(sq) || sp.rows.some(r=>r.produto.toLowerCase().includes(sq))
              ) : [];
              return (
              <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"64vh"}}>
                <div style={{padding:"10px 16px", borderBottom:"1px solid #F0EBE2", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap"}}>
                  <input value={subSearch} onChange={e=>setSubSearch(e.target.value)} placeholder="Buscar subproduto ou produto…"
                    style={{flex:"0 0 220px", padding:"5px 9px", border:"1px solid #D8D0C2", borderRadius:6, fontSize:12, background:"#FDFAF6"}} />
                  <div style={{fontSize:11, color:"#6B6153", flex:1}}>
                    {subDetalhado
                      ? <>Cálculo: <b>FT (unid/prod acabado) × Líquida = Total subproduto</b>. Unidade = coluna Quant Insumo da ficha técnica.</>
                      : <>Quantidade de cada subproduto para a produção líquida. <b>Detalhado</b> mostra o cálculo linha a linha.</> }
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>setSubDetalhado(false)}
                      style={{...S.btnGhost, fontSize:11, padding:"5px 10px", background:!subDetalhado?BRAND.amber:"transparent", color:!subDetalhado?"#fff":BRAND.amber}}>
                      Resumo
                    </button>
                    <button onClick={()=>setSubDetalhado(true)}
                      style={{...S.btnGhost, fontSize:11, padding:"5px 10px", background:subDetalhado?BRAND.amber:"transparent", color:subDetalhado?"#fff":BRAND.amber}}>
                      Detalhado
                    </button>
                  </div>
                  <button style={{...S.btn, fontSize:11, padding:"6px 12px"}} onClick={()=>{
                    if (!subprodData) return;
                    const hdr = ["Código Sub","Subproduto","Produto","FT (unid/prod)","Líquida (un)","Subtotal","Total Subproduto"];
                    const body = [];
                    for (const sp of subprodData) {
                      for (const row of sp.rows)
                        body.push([sp.cod, sp.nome, row.produto, +row.ftQuant.toFixed(2), row.liquida, +row.subtotal.toFixed(2), ""]);
                      body.push(["","","","","TOTAL",+sp.total.toFixed(2),+sp.total.toFixed(2)]);
                    }
                    const ws = styledSheet([hdr,...body],[8,36,40,14,10,12,14],"Subprodutos Detalhado");
                    const wb2=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2,ws,"Subprodutos");
                    XLSX.writeFile(wb2,`Subprodutos_${new Date().toISOString().slice(0,10)}.xlsx`);
                  }}>⬇ Excel</button>
                </div>

                {/* ---- MODO RESUMO ---- */}
                {!subDetalhado && (
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5}}>
                    <thead><tr>
                      <th style={{...S.thBase,textAlign:"right",width:80}}>Código</th>
                      <th style={{...S.thBase,textAlign:"left"}}>Subproduto</th>
                      <th style={{...S.thBase,textAlign:"right",width:150}} title="Unidade = coluna Quant Insumo da ficha técnica">Total (unid. FT)</th>
                      <th style={{...S.thBase,textAlign:"right",width:90}}>Nº produtos</th>
                      <th style={{...S.thBase,textAlign:"left"}}>Produtos que usam</th>
                    </tr></thead>
                    <tbody style={S.mono}>
                      {subVis.map((sp,i)=>(
                        <tr key={sp.cod} style={{borderBottom:"1px solid #F0EBE2",background:i%2===0?"#F6F3EE":"#fff"}}>
                          <td style={{padding:"5px 8px",textAlign:"right",color:"#8A8073"}}>{sp.cod}</td>
                          <td style={{padding:"5px 8px",textAlign:"left",fontFamily:"'Inter',sans-serif",fontWeight:600}}>{sp.nome}</td>
                          <td style={{padding:"5px 8px",textAlign:"right",fontWeight:700,color:BRAND.amber}}>{num(sp.total,2)}</td>
                          <td style={{padding:"5px 8px",textAlign:"right",color:"#8A8073"}}>{sp.rows.length}</td>
                          <td style={{padding:"5px 8px",textAlign:"left",fontSize:11,color:"#6B6153",fontFamily:"'Inter',sans-serif",maxWidth:400,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                            title={sp.rows.map(r=>r.produto).join(", ")}>
                            {sp.rows.slice(0,4).map(r=>r.produto).join(", ")}{sp.rows.length>4?` +${sp.rows.length-4} mais`:""}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* ---- MODO DETALHADO ---- */}
                {subDetalhado && (
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:12}}>
                    <thead><tr>
                      <th style={{...S.thBase,textAlign:"left",width:200}}>Subproduto</th>
                      <th style={{...S.thBase,textAlign:"left"}}>Produto que usa</th>
                      <th style={{...S.thBase,textAlign:"right",width:130}} title="Quantidade da ficha técnica por unidade do produto acabado">FT (unid/prod acabado)</th>
                      <th style={{...S.thBase,textAlign:"right",width:90}}>Líquida (un)</th>
                      <th style={{...S.thBase,textAlign:"right",width:160}}>= Total (unid. FT)</th>
                    </tr></thead>
                    <tbody style={S.mono}>
                      {subVis.map((sp)=>(
                        <React.Fragment key={sp.cod}>
                          <tr style={{background:"#F0DBBF"}}>
                            <td colSpan={4} style={{padding:"5px 10px",fontWeight:700,color:BRAND.dark,fontFamily:"'Inter',sans-serif",fontSize:12}}>
                              {sp.nome} <span style={{fontWeight:400,color:"#7A6450",fontSize:11}}>(cod {sp.cod})</span>
                            </td>
                            <td style={{padding:"5px 10px",textAlign:"right",fontWeight:700,color:BRAND.amber,fontSize:13}}>{num(sp.total,2)}</td>
                          </tr>
                          {sp.rows.map((row,j)=>(
                            <tr key={j} style={{borderBottom:"1px solid #F0EBE2",background:j%2===0?"#FAFAF8":"#fff"}}>
                              <td style={{padding:"4px 10px 4px 20px",color:"#8A8073",fontSize:11}}>{sp.nome}</td>
                              <td style={{padding:"4px 8px",textAlign:"left",fontFamily:"'Inter',sans-serif",fontSize:11}}>{row.produto}</td>
                              <td style={{padding:"4px 8px",textAlign:"right",color:"#264478",fontWeight:600}}>{num(row.ftQuant,2)}</td>
                              <td style={{padding:"4px 8px",textAlign:"right",color:"#2D6A4F"}}>{num(row.liquida)}</td>
                              <td style={{padding:"4px 8px",textAlign:"right",fontWeight:600,color:BRAND.amber}}>
                                {num(row.ftQuant,2)} × {num(row.liquida)} = {num(row.subtotal,2)}
                              </td>
                            </tr>
                          ))}
                        </React.Fragment>
                      ))}
                    </tbody>
                  </table>
                )}
                <div style={{padding:"8px 16px",fontSize:10,color:"#9A8E7F",borderTop:"1px solid #F0EBE2",fontStyle:"italic"}}>
                  Unidade: coluna <b>Quant Insumo</b> da ficha técnica. Confira no arquivo original qual unidade cada subproduto usa (kg, g, litros, unidades).
                  {sq && <> · <b>{subVis.length}</b> resultado{subVis.length!==1?"s":""} para <i>"{subSearch}"</i></>}
                </div>
                {(!subprodData||!subprodData.length)&&<div style={{padding:"24px",textAlign:"center",color:"#8A8073",fontSize:13}}>Nenhum produto do PCP tem ficha técnica com subprodutos mapeados.</div>}
              </div>
            );})()}

            {/* ===== TAB: SUGESTÃO DE COMPRAS ===== */}
            {tab==="compras" && fichasTec && (()=>{
              const cq = comprasSearch.trim().toLowerCase();
              const comprasVis = comprasData ? comprasData.filter(item=>
                !cq || item.nome.toLowerCase().includes(cq) || item.rows.some(r=>r.via.toLowerCase().includes(cq))
              ) : [];
              return (
              <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"64vh"}}>
                <div style={{padding:"10px 16px", borderBottom:"1px solid #F0EBE2", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap"}}>
                  <input value={comprasSearch} onChange={e=>setComprasSearch(e.target.value)} placeholder="Buscar insumo ou produto…"
                    style={{flex:"0 0 220px", padding:"5px 9px", border:"1px solid #D8D0C2", borderRadius:6, fontSize:12, background:"#FDFAF6"}} />
                  <div style={{fontSize:11, color:"#6B6153", flex:1}}>
                    {comprasDetalhado
                      ? <>Expansão recursiva de sub-produtos até insumos brutos. Fator FT = multiplicador acumulado da cadeia. Subtotal = Fator × Líquida.</>
                      : <>Insumos brutos totais — sub-produtos expandidos completamente. <b>Detalhado</b> mostra a cadeia de cada contribuição.</> }
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>setComprasDetalhado(false)}
                      style={{...S.btnGhost, fontSize:11, padding:"5px 10px", background:!comprasDetalhado?BRAND.amber:"transparent", color:!comprasDetalhado?"#fff":BRAND.amber}}>
                      Resumo
                    </button>
                    <button onClick={()=>setComprasDetalhado(true)}
                      style={{...S.btnGhost, fontSize:11, padding:"5px 10px", background:comprasDetalhado?BRAND.amber:"transparent", color:comprasDetalhado?"#fff":BRAND.amber}}>
                      Detalhado
                    </button>
                  </div>
                  <button style={{...S.btn, fontSize:11, padding:"6px 12px"}} onClick={()=>{
                    if (!comprasData) return;
                    const hdr = ["Código","Insumo","Via (cadeia produto → sub-produto)","Fator FT acumulado","Líquida (un)","Subtotal","Total Insumo"];
                    const body = [];
                    for (const item of comprasData) {
                      for (const row of item.rows)
                        body.push([item.cod, item.nome, row.via, +row.ftFactor.toFixed(4), row.liquida, +row.subtotal.toFixed(2), ""]);
                      body.push(["","","","","TOTAL",+item.total.toFixed(2),+item.total.toFixed(2)]);
                    }
                    const ws = styledSheet([hdr,...body],[8,40,60,14,10,12,14],"Compras Detalhado");
                    const wb2=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2,ws,"Compras");
                    XLSX.writeFile(wb2,`Compras_${new Date().toISOString().slice(0,10)}.xlsx`);
                  }}>⬇ Excel</button>
                </div>

                {/* ---- MODO RESUMO ---- */}
                {!comprasDetalhado && (
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5}}>
                    <thead><tr>
                      <th style={{...S.thBase,textAlign:"right",width:80}}>Código</th>
                      <th style={{...S.thBase,textAlign:"left"}}>Insumo</th>
                      <th style={{...S.thBase,textAlign:"right",width:150}} title="Unidade = coluna Quant Insumo da ficha técnica">Total (unid. FT)</th>
                      <th style={{...S.thBase,textAlign:"right",width:90}}>Nº contrib.</th>
                    </tr></thead>
                    <tbody style={S.mono}>
                      {comprasVis.map((item,i)=>(
                        <tr key={item.cod} style={{borderBottom:"1px solid #F0EBE2",background:i%2===0?"#F6F3EE":"#fff"}}>
                          <td style={{padding:"5px 8px",textAlign:"right",color:"#8A8073"}}>{item.cod}</td>
                          <td style={{padding:"5px 8px",textAlign:"left",fontFamily:"'Inter',sans-serif",fontWeight:600}}>{item.nome}</td>
                          <td style={{padding:"5px 8px",textAlign:"right",fontWeight:700,color:"#2D6A4F"}}>{num(item.total,2)}</td>
                          <td style={{padding:"5px 8px",textAlign:"right",color:"#8A8073",fontSize:11}}>{item.rows.length}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {/* ---- MODO DETALHADO ---- */}
                {comprasDetalhado && (
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:11.5}}>
                    <thead><tr>
                      <th style={{...S.thBase,textAlign:"left",width:200}}>Insumo</th>
                      <th style={{...S.thBase,textAlign:"left"}}>Cadeia (produto → sub-produtos → insumo)</th>
                      <th style={{...S.thBase,textAlign:"right",width:120}} title="Multiplicador acumulado de toda a cadeia FT, por unidade do produto acabado">Fator FT</th>
                      <th style={{...S.thBase,textAlign:"right",width:90}}>Líquida (un)</th>
                      <th style={{...S.thBase,textAlign:"right",width:140}}>= Subtotal</th>
                    </tr></thead>
                    <tbody style={S.mono}>
                      {comprasVis.map((item)=>(
                        <React.Fragment key={item.cod}>
                          <tr style={{background:"#E8F0E4"}}>
                            <td colSpan={4} style={{padding:"5px 10px",fontWeight:700,color:"#1A4A2F",fontFamily:"'Inter',sans-serif",fontSize:12}}>
                              {item.nome} <span style={{fontWeight:400,color:"#6B6153",fontSize:11}}>(cod {item.cod})</span>
                            </td>
                            <td style={{padding:"5px 10px",textAlign:"right",fontWeight:700,color:"#2D6A4F",fontSize:13}}>{num(item.total,2)}</td>
                          </tr>
                          {item.rows.map((row,j)=>(
                            <tr key={j} style={{borderBottom:"1px solid #F0EBE2",background:j%2===0?"#FAFDF8":"#fff"}}>
                              <td style={{padding:"3px 10px 3px 18px",color:"#8A8073",fontSize:10,fontFamily:"'Inter',sans-serif"}}>{item.nome}</td>
                              <td title={row.via} style={{padding:"3px 8px",textAlign:"left",fontFamily:"'Inter',sans-serif",fontSize:11,maxWidth:380,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.via}</td>
                              <td style={{padding:"3px 8px",textAlign:"right",color:"#264478",fontWeight:600}}>{num(row.ftFactor,4)}</td>
                              <td style={{padding:"3px 8px",textAlign:"right",color:"#2D6A4F"}}>{num(row.liquida)}</td>
                              <td style={{padding:"3px 8px",textAlign:"right",fontWeight:600,color:"#2D6A4F",fontSize:11}}>
                                {`${num(row.ftFactor,4)} × ${num(row.liquida)} = ${num(row.subtotal,2)}`}
                              </td>
                            </tr>
                          ))}
                        </React.Fragment>
                      ))}
                    </tbody>
                  </table>
                )}
                <div style={{padding:"8px 16px",fontSize:10,color:"#9A8E7F",borderTop:"1px solid #F0EBE2",fontStyle:"italic"}}>
                  Unidade: coluna <b>Quant Insumo</b> da ficha técnica. Sub-produtos são expandidos recursivamente — somente insumos brutos aparecem aqui.
                  {cq && <> · <b>{comprasVis.length}</b> resultado{comprasVis.length!==1?"s":""} para <i>"{comprasSearch}"</i></>}
                </div>
                {(!comprasData||!comprasData.length)&&<div style={{padding:"24px",textAlign:"center",color:"#8A8073",fontSize:13}}>Nenhum produto do PCP tem ficha técnica carregada ou a produção líquida é zero para todos.</div>}
              </div>
            );})()}

            {/* ===== TAB: CMV ===== */}
            {tab==="cmv" && (
              <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"64vh"}}>
                <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5, minWidth:1100}}>
                  <thead><tr>
                    {th("Cod","cod")}{th("Produto","produto",{left:true})}{th("Categ.","categoria",{left:true})}
                    {th("Qde Vendida","cmvQde")}{th("Venda Total R$","cmvVenda")}{th("Custo Total R$","cmvCusto")}
                    {th("Custo Unit.","custoUnit")}{th("Preço Médio","precoMedio")}{th("CMV %","cmvPct")}
                  </tr></thead>
                  <tbody style={S.mono}>
                    {visRows.slice(0,400).map((r)=>(
                      <tr key={r.cod} style={{borderBottom:"1px solid #F0EBE2"}}>
                        {td(r.cod,{style:{color:"#8A8073"}})}
                        <td title={r.produto||`cod ${r.cod}`} style={{padding:"5px 8px",textAlign:"left",whiteSpace:"nowrap"}}>
                          <div style={{width:220,overflow:"hidden",textOverflow:"ellipsis",fontFamily:"'Inter',sans-serif",color:BRAND.dark}}>
                            {r.produto||<span style={{color:"#B0A794",fontStyle:"italic"}}>cod {r.cod}</span>}
                          </div>
                        </td>
                        {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:"#6B6153"}})}
                        {td(num(r.cmvQde))}
                        {td(num(r.cmvVenda,2))}
                        {td(num(r.cmvCusto,2))}
                        {td(num(r.custoUnit,2))}
                        {td(num(r.precoMedio,2))}
                        {td(r.cmvPct==null
                            ? <span style={S.tag("#FCE4D6","#C4501E")}>SEM VENDA</span>
                            : `${num(r.cmvPct*100,1)}%`,
                          {style:{fontWeight:600, color:r.cmvPct!=null&&r.cmvPct>0.45?"#C4501E":BRAND.dark}})}
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div style={{padding:"10px 16px", fontSize:11, color:"#8A8073", borderTop:"1px solid #F0EBE2"}}>
                  CMV % em vermelho quando acima de 45%. "SEM VENDA" = custo registrado sem receita — auditar lançamento.
                </div>
              </div>
            )}
          </>
        )}

        {!result && (
          <div style={{ ...S.panel, textAlign:"center", padding:44 }}>
            {files.length > 0 ? (
              <>
                <div style={{fontSize:38, marginBottom:6}}>⚠️</div>
                <div style={{fontSize:16, fontWeight:600, color:"#C4501E"}}>PCP não calculado</div>
                <div style={{fontSize:13, marginTop:6, color:"#6B6153"}}>
                  {semLoja
                    ? "Selecione a loja para cada arquivo de vendas marcado em vermelho."
                    : allSales.length === 0
                      ? "Nenhuma venda foi reconhecida nos arquivos. Verifique se o formato está correto (exportação do sistema ou Modelo de Vendas)."
                      : "Verifique os parâmetros e os arquivos carregados."}
                </div>
              </>
            ) : (
              <>
                <div style={{fontSize:38, marginBottom:6}}>📦</div>
                <div style={{fontSize:16, fontWeight:600, color:BRAND.dark}}>Suba as bases de vendas para gerar o PCP</div>
                <div style={{fontSize:13, marginTop:6, color:"#8A8073"}}>Processamento 100% no navegador. Use "Salvar sessão" para não perder os dados ao fechar.</div>
              </>
            )}
          </div>
        )}

        <div style={{textAlign:"center", padding:"16px 0 26px", color:"#A08060", fontSize:12}}>
          <img src="/tortele-logo.png" alt="tortelê" style={{height:22, opacity:0.5, verticalAlign:"middle", marginRight:8}} />
          Sistema de PCP · Liberdata Consultoria
        </div>
      </div>
    </div>
  );
}
