import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

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
    const iProd = header.findIndex((c) => c.includes("produto"));
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
    // Fallback: formato izzyway (col0=vazio, col1=Cod., col3=Qde.)
    if (hi < 0) {
      const h0 = (data[0] || []).map((c) => String(c || "").toLowerCase().trim());
      if (h0[1] === "cod." && h0[2] === "produto" && (h0[3] || "").startsWith("qde")) {
        hi = 0; iCod = 1; iQty = 3;
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
    let hi = -1, iCod = -1, iCat = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const r = (data[i] || []).map((c) => String(c || "").toLowerCase().trim());
      const ic = r.findIndex((c) => c === "código" || c === "codigo" || c === "cod");
      const ik = r.findIndex((c) => c === "categoria");
      if (ic >= 0 && ik >= 0) { hi = i; iCod = ic; iCat = ik; break; }
    }
    if (hi < 0) continue;
    for (let i = hi + 1; i < data.length; i++) {
      const r = data[i] || [];
      const cod = parseInt(r[iCod], 10);
      const cat = String(r[iCat] || "").trim();
      if (cod && !isNaN(cod) && cat) map[cod] = cat;
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
    const categoria = categorias[cod] || "S/CAT";

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

    return { cod, produto: names[cod], categoria, weeks, partial, combo, media, dp, cv, min, max,
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

/* ---------------- persistência (window.storage) ---------------- */
const CHUNK = 20000;
async function saveState({ files, estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos, params }) {
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
    await window.storage.set("aux", JSON.stringify({ estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos, params }));
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
    const byLoja = {};
    for (const [d, cod, qde, li, vendido, custo] of flat) {
      const loja = LOJAS[li];
      (byLoja[loja] ||= []).push({ dt: new Date(d*86400000), cod, produto: "", qde, vendido, custo, loja });
    }
    const files = meta.fmeta.filter((f)=>f.loja).map((f)=>({ name: f.name, loja: f.loja, rows: [], warnings: [], fromCache: true }));
    // reconstruir rows por loja na ordem dos files (aproximação: agrupar por loja)
    const grouped = {};
    for (const l of LOJAS) grouped[l] = byLoja[l] || [];
    const outFiles = [];
    for (const l of LOJAS) if (grouped[l].length) outFiles.push({ name: `(sessão anterior) ${l}`, loja: l, rows: grouped[l], warnings: [], fromCache: true });
    return { files: outFiles, ...aux, savedAt: meta.savedAt };
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
function exportAll(result) {
  const { rows, weekStarts } = result;
  const wb = XLSX.utils.book_new();
  // PCP Semanal
  const h1 = ["Cod","Produto","Categoria",...weekStarts.map(fmtDM),"Parcial","Combo","Média","DP","CV%","Mín","Máx","Tend%","ES","Marg%","Sugerida","Estoque","LÍQUIDA","ABC","Alertas",...DIAS.map(d=>`Líq ${d}`),"Total dia"];
  const d1 = rows.map((r)=>[r.cod,r.produto,r.categoria,...r.weeks,r.partial,+r.combo.toFixed(1),+r.media.toFixed(1),+r.dp.toFixed(1),+(r.cv*100).toFixed(1),r.min,r.max,+(r.tend*100).toFixed(1),+r.es.toFixed(1),+(r.margem*100).toFixed(1),r.sugerida,r.estoque,r.liquida,r.abc,r.alertas.join("|"),...r.liqDia,r.liqDia.reduce((a,b)=>a+b,0)]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h1,...d1]), "PCP Semanal");
  // Programação
  const h2 = ["Cod","Produto","Categoria","Prog Seg","Prog Ter","Prog Qua","Prog Qui","Prog Sex","Prog Sáb","Total",...DIAS.map(d=>`Sug ${d} (quinz)`)];
  const d2 = rows.map((r)=>[r.cod,r.produto,r.categoria,...r.prog.slice(0,6).map(v=>v==null?"-":v),r.liqDia.reduce((a,b)=>a+b,0),...r.sugDia]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h2,...d2]), "Programação");
  // PCP por Loja
  const h3 = ["Cod","Produto","Categoria","Sugerida","Estoque","LÍQUIDA",...LOJAS.map(l=>`${l} %`),...LOJAS.map(l=>`Sug ${l}`),...LOJAS.map(l=>`Est ${l}`),"Est CD",...LOJAS.map(l=>`Líq ${l}`)];
  const d3 = rows.map((r)=>[r.cod,r.produto,r.categoria,r.sugerida,r.estoque,r.liquida,...r.mixLoja.map(m=>+(m*100).toFixed(1)),...r.sugLoja,...r.estPorLoja,r.estoquePorLocal[4],...r.liqLoja]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h3,...d3]), "PCP por Loja");
  // CMV
  const h4 = ["Cod","Produto","Categoria","Qde Vendida","Venda Total","Custo Total","Custo Unit Médio","Preço Médio","CMV %"];
  const d4 = rows.map((r)=>[r.cod,r.produto,r.categoria,+r.cmvQde.toFixed(0),+r.cmvVenda.toFixed(2),+r.cmvCusto.toFixed(2),+r.custoUnit.toFixed(2),+r.precoMedio.toFixed(2),r.cmvPct==null?"SEM VENDA":+(r.cmvPct*100).toFixed(1)]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h4,...d4]), "Auditoria CMV");
  XLSX.writeFile(wb, `PCP_Tortele_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ================= UI ================= */
const S = {
  page: { minHeight: "100vh", background: "#F6F3EE", color: "#25211C", fontFamily: "'Inter', system-ui, sans-serif", fontSize: 14 },
  mono: { fontFamily: "'JetBrains Mono','SF Mono',Consolas,monospace" },
  panel: { background: "#fff", border: "1px solid #E4DED4", borderRadius: 10, padding: 18, marginBottom: 14 },
  btn: { background: "#B96A1B", color: "#fff", border: "none", borderRadius: 8, padding: "10px 18px", fontWeight: 600, cursor: "pointer", fontSize: 14 },
  btnGhost: { background: "transparent", color: "#B96A1B", border: "1.5px solid #B96A1B", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 },
  tag: (bg, fg) => ({ display: "inline-block", background: bg, color: fg, borderRadius: 5, fontSize: 11, fontWeight: 700, padding: "2px 7px", marginRight: 4 }),
  tabBtn: (active) => ({
    padding: "10px 18px", border: "none", cursor: "pointer", fontWeight: 600, fontSize: 14,
    background: active ? "#25211C" : "transparent", color: active ? "#F6F3EE" : "#6B6153",
    borderRadius: "8px 8px 0 0",
  }),
  thBase: { padding: "7px 8px", fontSize: 11, fontWeight: 700, color: "#6B6153", textTransform: "uppercase", letterSpacing: "0.03em", borderBottom: "2px solid #E4DED4", whiteSpace: "nowrap", position: "sticky", top: 0, background: "#fff", zIndex: 2 },
};

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
  const [nivelServico, setNivelServico] = useState(0.9);
  const [janela, setJanela] = useState(4);
  const [dataRef, setDataRef] = useState(() => new Date().toISOString().slice(0, 10));
  const refEstArq = useRef(null);

  const [busca, setBusca] = useState("");
  const [filtroABC, setFiltroABC] = useState("");
  const [filtroCat, setFiltroCat] = useState("");
  const [soAlertas, setSoAlertas] = useState(false);
  const [sortKey, setSortKey] = useState("vol4");
  const [sortDesc, setSortDesc] = useState(true);
  const [lojaEnvio, setLojaEnvio] = useState("ALDEOTA");
  const [erro, setErro] = useState(null);
  const [saveStatus, setSaveStatus] = useState("");
  const [loadedFromCache, setLoadedFromCache] = useState(false);
  const refVendas = useRef(null), refEst = useRef(null), refCat = useRef(null), refCombo = useRef(null);

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
        if (st.params) {
          setNivelServico(st.params.nivelServico ?? 0.9);
          setJanela(st.params.janela ?? 4);
          if (st.params.dataRef) setDataRef(st.params.dataRef);
          setUsarCombos(st.params.usarCombos ?? true);
        }
        if (Object.keys(st.categorias || {}).length) setCatInfo(`${Object.keys(st.categorias).length} produtos (sessão anterior)`);
        if ((st.combos || []).length) setComboInfo(`${st.combos.length} linhas (sessão anterior)`);
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

  const cats = useMemo(() => result ? [...new Set(result.rows.map((r)=>r.categoria))].sort() : [], [result]);

  const visRows = useMemo(() => {
    if (!result) return [];
    let r = result.rows;
    if (busca) { const b = busca.toLowerCase(); r = r.filter((x)=>String(x.cod).includes(b)||x.produto.toLowerCase().includes(b)); }
    if (filtroABC) r = r.filter((x)=>x.abc===filtroABC);
    if (filtroCat) r = r.filter((x)=>x.categoria===filtroCat);
    if (soAlertas) r = r.filter((x)=>x.alertas.length>0);
    r = [...r].sort((a,b)=>{ const va=a[sortKey]??0, vb=b[sortKey]??0;
      if (typeof va==="string") return sortDesc?vb.localeCompare(va):va.localeCompare(vb);
      return sortDesc?vb-va:va-vb; });
    return r;
  }, [result, busca, filtroABC, filtroCat, soAlertas, sortKey, sortDesc]);

  const totais = useMemo(() => result ? {
    produtos: result.rows.length,
    sugerida: result.rows.reduce((a,r)=>a+r.sugerida,0),
    liquida: result.rows.reduce((a,r)=>a+r.liquida,0),
    estoque: result.rows.reduce((a,r)=>a+r.estoque,0),
    alertas: result.rows.filter((r)=>r.alertas.length).length,
  } : null, [result]);

  const doSave = async () => {
    setSaveStatus("salvando…");
    const ok = await saveState({ files, estoqueLoja, estoqueInfo, estoqueArqs, categorias, combos,
      params: { nivelServico, janela, dataRef, usarCombos } });
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
    <td style={{ padding:"5px 8px", textAlign: extra.left?"left":"right", whiteSpace:"nowrap", ...extra.style }}>{v}</td>
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
      <div style={{ background:"#25211C", color:"#F6F3EE", padding:"16px 26px", display:"flex", alignItems:"center", gap:16, flexWrap:"wrap" }}>
        <div style={{ fontSize:21, fontWeight:700 }}>PCP <span style={{color:"#E8A34E"}}>Tortelê</span></div>
        {result && (
          <div style={{...S.mono, fontSize:13, opacity:0.9}}>
            Semana ref: <b style={{color:"#E8A34E"}}>{fmtDM(result.lastWeekStart)}–{fmtDM(addDays(result.lastWeekStart,6))}</b>
            {" · "}Z=<b style={{color:"#E8A34E"}}>{result.Z}</b>
            {result.nCombos>0 && usarCombos && <> · combos ativos</>}
          </div>
        )}
        <div style={{ marginLeft:"auto", display:"flex", gap:10, alignItems:"center" }}>
          {saveStatus && <span style={{fontSize:12, color:"#E8A34E"}}>{saveStatus}</span>}
          <button style={{...S.btnGhost, borderColor:"#E8A34E", color:"#E8A34E"}} onClick={doSave} disabled={!files.length}>💾 Salvar sessão</button>
          <button style={{...S.btnGhost, borderColor:"#8A8073", color:"#B0A794", fontSize:12, padding:"6px 10px"}} onClick={doClear}>limpar salvos</button>
        </div>
      </div>

      <div style={{ maxWidth: 1560, margin: "0 auto", padding: "20px 26px" }}>

        {loadedFromCache && (
          <div style={{ ...S.panel, background:"#FDF6EC", borderColor:"#E8C99A", fontSize:13, padding:"10px 16px" }}>
            ⚡ Dados restaurados da sessão anterior. Suba novas bases para atualizar ou continue de onde parou.
          </div>
        )}

        {/* CADASTROS / UPLOADS */}
        <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr 1fr", gap:12 }}>
          <div style={S.panel}>
            <div style={{ marginBottom:10 }}><span style={S.tag("#25211C","#F6F3EE")}>1</span><b> Bases de vendas</b></div>
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
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 0", borderBottom:"1px solid #F0EBE2" }}>
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
                ))}
              </div>
            )}
            {semLoja && <div style={{marginTop:8, color:"#C4501E", fontWeight:600, fontSize:12}}>⚠ Selecione a loja dos arquivos marcados.</div>}
          </div>

          {/* Painel Estoque */}
          <div style={S.panel}>
            <div style={{ marginBottom:8 }}><span style={S.tag("#25211C","#F6F3EE")}>2</span><b> Estoque atual</b></div>
            {/* Opção A: arquivo único com todas as lojas */}
            <button style={{...S.btnGhost, width:"100%", fontSize:12}} onClick={()=>refEst.current?.click()}>
              Subir arquivo (todas as lojas)
            </button>
            <input ref={refEst} type="file" accept=".xlsx,.xls" style={{display:"none"}}
              onChange={(e)=>{handleEstoque(e.target.files);e.target.value="";}} />
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
              <button style={{...S.btnGhost, width:"100%", fontSize:12}} onClick={()=>refEstArq.current?.click()}>
                Subir por loja (vários arquivos)
              </button>
              <input ref={refEstArq} type="file" multiple accept=".xlsx,.xls" style={{display:"none"}}
                onChange={(e)=>{handleEstoqueArq(e.target.files);e.target.value="";}} />
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
          ].map((b)=>(
            <div key={b.n} style={S.panel}>
              <div style={{ marginBottom:8 }}><span style={S.tag("#25211C","#F6F3EE")}>{b.n}</span><b> {b.t}</b></div>
              <button style={{...S.btnGhost, width:"100%", fontSize:12}} onClick={()=>b.ref.current?.click()}>Subir arquivo</button>
              <input ref={b.ref} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                onChange={(e)=>{b.handler(e.target.files);e.target.value="";}} />
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
              {[["Produtos",num(totais.produtos),"#25211C"],["Sugerida total",num(totais.sugerida),"#B96A1B"],
                ["Estoque abatido",num(totais.estoque),"#6B6153"],["Produção LÍQUIDA",num(totais.liquida),"#2D6A4F"],
                ["Com alerta",num(totais.alertas),"#C4501E"]].map(([l,v,c])=>(
                <div key={l} style={{...S.panel, marginBottom:0, padding:"12px 16px"}}>
                  <div style={{fontSize:10, color:"#8A8073", textTransform:"uppercase", fontWeight:700}}>{l}</div>
                  <div style={{...S.mono, fontSize:24, fontWeight:600, color:c}}>{v}</div>
                </div>
              ))}
            </div>

            {/* TABS */}
            <div style={{ display:"flex", gap:2, borderBottom:"2px solid #25211C" }}>
              {[["pcp","PCP Semanal"],["prog","Programação"],["loja","PCP por Loja"],["cmv","Auditoria CMV"]].map(([k,l])=>(
                <button key={k} style={S.tabBtn(tab===k)} onClick={()=>setTab(k)}>{l}</button>
              ))}
            </div>

            {/* FILTROS */}
            <div style={{ ...S.panel, borderRadius:"0 0 10px 10px", borderTop:"none", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
              <input placeholder="Buscar código ou produto…" value={busca} onChange={(e)=>setBusca(e.target.value)}
                style={{padding:"8px 12px", borderRadius:8, border:"1px solid #D8D0C2", width:240, fontSize:13}}/>
              <select value={filtroCat} onChange={(e)=>setFiltroCat(e.target.value)}
                style={{padding:"8px 10px", borderRadius:8, border:"1px solid #D8D0C2", fontSize:13, maxWidth:200}}>
                <option value="">Categoria: todas</option>
                {cats.map((c)=><option key={c} value={c}>{c}</option>)}
              </select>
              <select value={filtroABC} onChange={(e)=>setFiltroABC(e.target.value)}
                style={{padding:"8px 10px", borderRadius:8, border:"1px solid #D8D0C2", fontSize:13}}>
                <option value="">ABC: todos</option><option value="A">A</option><option value="B">B</option><option value="C">C</option>
              </select>
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
                    {th("Cod","cod")}{th("Produto","produto",{left:true})}{th("Categ.","categoria",{left:true})}
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
                        {td(r.produto,{left:true,style:{fontFamily:"'Inter',sans-serif",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}})}
                        {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:"#6B6153",maxWidth:110,overflow:"hidden",textOverflow:"ellipsis"}})}
                        {r.weeks.map((v,i)=>td(v?num(v):"·",{style:{color:i===7?"#25211C":"#8A8073",fontWeight:i===7?600:400}}))}
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
                <div style={{...S.panel, fontSize:12, color:"#6B6153", background:"#FDF6EC"}}>
                  <b>Regra de produção:</b> categorias com "Salgado" produzem em dias alternados —
                  Seg cobre Ter+Qua · Qua cobre Qui+Sex · Sex cobre Sáb+Dom · Sáb cobre Seg seguinte.
                  Demais categorias produzem no dia (Sáb cobre Sáb+Dom).
                  A tabela auxiliar quinzenal usa a <b>Sugerida</b> (bruta, sem descontar estoque).
                </div>
                <div style={{...S.panel, padding:0, overflow:"auto", maxHeight:"60vh"}}>
                  <table style={{borderCollapse:"collapse", width:"100%", fontSize:12.5, minWidth:1300}}>
                    <thead><tr>
                      {th("Cod","cod")}{th("Produto","produto",{left:true})}{th("Categ.","categoria",{left:true})}
                      {["Seg","Ter","Qua","Qui","Sex","Sáb"].map((d)=>th(`Prog ${d}`,null,{style:{background:"#264478",color:"#fff"}}))}
                      {th("Total",null)}
                      {DIAS.map((d)=>th(`Sug ${d}`,null,{style:{color:"#B96A1B"}}))}
                      {th("Tot Sug",null)}
                    </tr></thead>
                    <tbody style={S.mono}>
                      {visRows.slice(0,400).map((r)=>(
                        <tr key={r.cod} style={{borderBottom:"1px solid #F0EBE2"}}>
                          {td(r.cod,{style:{color:"#8A8073"}})}
                          {td(r.produto,{left:true,style:{fontFamily:"'Inter',sans-serif",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}})}
                          {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:isSalgado(r.categoria)?"#B96A1B":"#6B6153"}})}
                          {r.prog.slice(0,6).map((v,i)=>td(v==null?"—":num(v),{style:{background:"#EDF1F8",color:"#264478",fontWeight:v?600:400}}))}
                          {td(num(r.liqDia.reduce((a,b)=>a+b,0)),{style:{fontWeight:600}})}
                          {r.sugDia.map((v,i)=>td(v||"·",{style:{color:"#B96A1B",background:"#FDF6EC"}}))}
                          {td(num(r.sugDia.reduce((a,b)=>a+b,0)),{style:{fontWeight:600,color:"#B96A1B"}})}
                        </tr>
                      ))}
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
                          {td(r.produto,{left:true,style:{fontFamily:"'Inter',sans-serif",maxWidth:190,overflow:"hidden",textOverflow:"ellipsis"}})}
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
                        {td(r.produto,{left:true,style:{fontFamily:"'Inter',sans-serif",maxWidth:230,overflow:"hidden",textOverflow:"ellipsis"}})}
                        {td(r.categoria,{left:true,style:{fontFamily:"'Inter',sans-serif",fontSize:11,color:"#6B6153"}})}
                        {td(num(r.cmvQde))}
                        {td(num(r.cmvVenda,2))}
                        {td(num(r.cmvCusto,2))}
                        {td(num(r.custoUnit,2))}
                        {td(num(r.precoMedio,2))}
                        {td(r.cmvPct==null
                            ? <span style={S.tag("#FCE4D6","#C4501E")}>SEM VENDA</span>
                            : `${num(r.cmvPct*100,1)}%`,
                          {style:{fontWeight:600, color:r.cmvPct!=null&&r.cmvPct>0.45?"#C4501E":"#25211C"}})}
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
          <div style={{ ...S.panel, textAlign:"center", padding:44, color:"#8A8073" }}>
            <div style={{fontSize:38, marginBottom:6}}>📦</div>
            <div style={{fontSize:16, fontWeight:600, color:"#25211C"}}>Suba as bases de vendas para gerar o PCP</div>
            <div style={{fontSize:13, marginTop:6}}>Processamento 100% no navegador. Use "Salvar sessão" para não perder os dados ao fechar.</div>
          </div>
        )}

        <div style={{textAlign:"center", padding:"16px 0 26px", color:"#B0A794", fontSize:12}}>
          PCP Tortelê · Liberdata Consultoria · client-side
        </div>
      </div>
    </div>
  );
}
