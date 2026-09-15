# Histórico de versões — PCP Tortelê Web

## 15/09/2026 — v1.24 — Sugestão de Compras: estoque do insumo e unidade de compra real

### Novidades
- **Coluna "Estoque Atual" na Sugestão de Compras**: reaproveita o mesmo arquivo de Estoque Atual já usado no PCP — se o insumo tiver estoque cadastrado lá (comum quando o relatório do Izzyway já cobre todos os produtos, inclusive matérias-primas), ele aparece aqui automaticamente. Sem necessidade de nova planilha.
- **Coluna "A Comprar (líquido)"**: Compra Mínima menos o Estoque Atual do insumo — a quantidade líquida que realmente falta comprar.
- **Novo painel opcional "6 — Unidade de Compra"**: planilha com Código do Insumo + Unidade de Compra + Fator de Conversão (ex: farinha usada em KG na receita, comprada em saco de 25kg → fator 25). Quando cadastrada, mostra a coluna "Comprar (unid. real)" já convertida e arredondada para cima (nº de sacos/caixas/fardos a comprar).
- Exportação em Excel atualizada com todas as colunas novas.

### Pendente (aguardando decisão)
- Categoria dos insumos: adiada a pedido do Kauan — sem planilha de categoria dedicada por enquanto.
- Custo estimado da compra: depende do relatório de compras do sistema, ainda não integrado.

---

## 15/09/2026 — v1.23 — Sugestão de Compras: renomeação, tendência semanal

### Ajustes
- **Coluna "Total (unid. FT)" renomeada para "Compra Mínima"** na aba Sugestão de Compras, deixando mais claro que esse número é o mínimo necessário para cobrir a produção da semana.
- **Coluna "Nº contribuições" removida** da tela e da exportação — não agregava valor para quem está montando a lista de compras.
- **Nova coluna "Tendência (sem.)"**: compara a demanda real de cada insumo na última semana fechada com a semana anterior, mostrando seta de alta/queda/estável e o percentual de variação. Ajuda a perceber se o consumo de um insumo está subindo ou caindo antes de fechar a compra.
- Exportação em Excel agora traz Compra Mínima, Semana Atual e Semana Anterior lado a lado.

### Pendente (aguardando decisão)
- Estoque atual, categoria e unidade de compra real dos insumos: a planilha de fichas técnicas do Izzyway não traz essas informações prontas — item em aberto até definirmos a fonte de dado.

---

## 09/09/2026 — v1.22 — Correção de datas do PCP Semanal e vendas diretas de subprodutos

### Correções
- **PCP Semanal — deslocamento de 1 dia corrigido**: ao recarregar uma sessão salva, as datas de venda estavam sendo reconstruídas 3 horas adiantadas (fuso), fazendo vendas de segunda-feira caírem na semana errada. Corrigido a reconstrução da data para usar o dia local correto.
- **Subprodutos — vendas diretas contabilizadas**: quando um subproduto também é vendido individualmente (não só como componente de outro produto), a venda direta agora entra no total da aba Subprodutos.
- **Sugestão de Compras — exportação só em Resumo**: removida a opção de exportar em formato Detalhado; a exportação Excel agora traz sempre uma linha por insumo com o total.

---

## 02/09/2026 — v1.21 — Busca nas abas de ficha técnica, explosão recursiva e sessão persistente

### Novidades
- **Busca por nome** nas abas Subprodutos e Sugestão de Compras.
- **Explosão de BOM recursiva**: subprodutos que usam outros subprodutos (qualquer profundidade) são totalmente destrinchados até os insumos brutos na Sugestão de Compras.
- **Sessão salva agora inclui a ficha técnica**: antes, ao salvar a sessão no navegador e reabrir depois, era preciso subir a ficha técnica de novo.

---

## 02/09/2026 — v1.20 — Novo leitor de Fichas Técnicas (formato nativo Izzyway)

### Novidades
- **Leitor de planilha reescrito** para ler o arquivo exportado direto do Izzyway (sem precisar reformatar manualmente). Detecta produtos, insumos e rendimento automaticamente a partir do formato de blocos do relatório.

---

## 01/09/2026 — v1.19 — Casas decimais e unidades nas abas de ficha técnica

### Ajustes
- Números das abas Subprodutos e Sugestão de Compras agora mostram no máximo 2 casas decimais.
- Cabeçalhos de coluna e rodapé explicando que a unidade vem da coluna "Quant Insumo" da ficha técnica.

---

## 02/09/2026 — v1.17 — Fichas Técnicas: Subprodutos, Sugestão de Compras, Excel/PDF na Programação

### Novidades
- **Upload de Fichas Técnicas (painel 5)**: suba a planilha `Fichas_Tecnicas_Organizado.xlsx` (ou qualquer arquivo com as abas "Ficha Técnica" e "Subfichas") diretamente na tela. O sistema detecta automaticamente quais insumos são subprodutos e faz a explosão em dois níveis.
- **Aba Subprodutos**: mostra quanto de cada subproduto (ex.: Massa Coxinha, Recheio Frango, Pão de Ló) precisa ser produzido para cobrir a produção líquida da semana. Exibe os produtos que consomem cada subproduto e exporta para Excel.
- **Aba Sugestão de Compras**: explosão BOM completa — insumos diretos das fichas técnicas + insumos dos subprodutos expandidos pelas subfichas. Resultado: total de cada matéria-prima bruta necessária para a semana. Exporta para Excel.
- **Exportação Excel da Programação melhorada**: substituída a saída bruta por planilha estilizada com cabeçalho da marca, semana 1 em azul, semana 2 em verde, coluna Líquida destacada, saldo colorido (verde/vermelho).
- **Exportação PDF da Programação**: novo botão "🖨 PDF" — abre janela otimizada para impressão em paisagem, lista completa de produtos com os dias de produção e saldo.

---

## 31/08/2026 — v1.14 — Setor no modelo de categorias + Filtros em checkbox

### Novidades
- **Coluna Setor no modelo de categorias**: a planilha `Modelo_Categorias.xlsx` agora tem uma quarta coluna "Setor". O sistema lê o setor de cada produto e permite filtrar por ele. Setores já vêm pré-preenchidos (Salgados, Doces, Bebidas, Combos, Insumos, Materiais, Administrativo) — basta ajustar e re-subir o arquivo.
- **Filtros em checkbox**: os filtros de Setor, Categoria e ABC nas quatro abas deixaram de ser listas drop-down de seleção única e passaram a ser menus de checkbox — selecione múltiplos valores ao mesmo tempo. Botão "Limpar filtro" aparece quando há seleção ativa.
- **Setor como filtro independente**: o filtro de Setor só aparece quando a base de categorias carregada tiver a coluna Setor preenchida.

---

## 27/08/2026 — v1.13 — Correção: Envio Diário agora usa Sugerida como base

### Correção
- **Base do Envio Diário corrigida**: o relatório Excel "Envio Diário" agora usa o valor **Sugerida** (= Média + Estoque de Segurança) como total semanal — o mesmo número exibido na coluna Sugerida do PCP Semanal. Antes usava apenas a Média, gerando uma diferença do tamanho do ES (ex.: item 900 mostrava 3076 vs 3448 no PCP). Agora ambos mostram o mesmo valor.
- **Novas colunas no Excel**: adicionadas colunas "Média" e "ES" ao lado de "Sugerida" para facilitar conferência.

---

## 27/08/2026 — v1.6 — Drag-and-drop + Aba de Programação interativa

### Novidades
- **Drag-and-drop em todos os painéis**: agora é possível arrastar diretamente os arquivos de estoque, categorias e combos para as áreas de upload — igual ao painel de vendas.
- **Aba Programação reformulada**: agora mostra as datas reais da semana de produção como colunas (Seg 25/08, Ter 26/08…). Para cada produto e cada dia:
  - Linha cinza: venda prevista para aquele dia (baseada no histórico de mix)
  - Campo editável: quantidade a produzir naquele dia (pré-preenchida pelo PCP, editável manualmente)
  - Número colorido: estoque projetado ao final do dia — verde se acima do estoque de segurança, vermelho se abaixo
- **Saldo de programação**: coluna final mostra se a programação da semana cobre a produção líquida recomendada (+) ou deixa faltando (-).
- **Exportação da programação em Excel**: botão para baixar a tabela com os ajustes feitos.
- **Resetar ajustes**: botão para voltar aos valores calculados automaticamente pelo PCP.

---

## 26/08/2026 — v1.5 — Identidade visual Tortelê no site e na exportação

### Novidades
- **Logo da Tortelê no cabeçalho**: o topo do sistema agora exibe a logo oficial com o fundo marrom chocolate da marca.
- **Cores da marca aplicadas ao site**: marrom `#3C2008` e creme `#F0DBBF` alinhados com a identidade visual da Tortelê em todo o sistema.
- **Logo na exportação Excel**: cada aba do arquivo exportado abre com uma linha de cabeçalho escura com o nome "tortelê" e o título da aba, no padrão visual da marca.
- **Rodapé discreto com a logo** na parte inferior da página.

---

## 26/08/2026 — v1.4 — Estoque bruto Izzyway + Excel estilizado

### Novidades
- **Estoque bruto do Izzyway**: agora é possível subir o arquivo de estoque exportado diretamente do sistema (formato "ALD ESTOQUE 25.08.xlsx", "MEI ESTOQUE 25.08.xlsx", etc.) sem nenhum tratamento — o sistema reconhece automaticamente o formato.
- **Excel exportado reformulado**: cabeçalhos com fundo escuro e texto branco, linhas alternadas para facilitar leitura, largura de colunas ajustada, CMV acima de 45% destacado em vermelho, programação de produção destacada em azul. O arquivo está pronto para imprimir.

---

## 26/08/2026 — v1.3 — Correção: salvar sessão funcionando no navegador

### Correção
- **"Erro ao salvar" resolvido**: o sistema usava internamente um mecanismo de armazenamento que só existe no ambiente de desenvolvimento — no link do Vercel (navegador comum) ele simplesmente não existia, fazendo o salvar falhar sempre. Corrigido para usar o armazenamento padrão do navegador (`localStorage`). Agora "Salvar sessão" e "Restaurar sessão" funcionam normalmente para todos.

---

## 26/08/2026 — v1.2 — Indicadores de erro e diagnóstico

### Novidades
- **Banner de erro visível**: quando um arquivo não é reconhecido ou falha ao ser lido, aparece agora um aviso vermelho no topo da tela com a mensagem do problema. O operacional pode fechar o aviso depois de ver.
- **Avisos por arquivo**: o sistema agora exibe, embaixo de cada arquivo carregado, informações sobre o formato detectado e quantos registros foram lidos (ex.: "Formato exportação sistema: 240 registros em 3 datas").
- **Diagnóstico quando o PCP não calcula**: se os arquivos foram carregados mas o PCP continua em branco, o sistema agora explica o motivo — loja não selecionada, formato não reconhecido, ou outro problema — em vez de exibir apenas a tela de "suba as bases".

## 25/08/2026 — v1.1 — Correção de dados zerados + suporte ao formato do sistema

### Correções
- **Bug crítico resolvido**: todas as colunas semanais mostravam ponto (·) mesmo com os arquivos de venda carregados corretamente. O sistema estava comparando os dados em horários diferentes internamente. Corrigido — agora os números aparecem como esperado.

### Novidades
- **Upload de vendas no formato nativo do sistema (izzyway)**: não é mais necessário usar o Modelo de Vendas como intermediário. Basta exportar do izzyway e subir o arquivo diretamente — o sistema reconhece automaticamente o formato com as datas por seção.
- **Estoque por loja**: novo botão "Subir por loja (vários arquivos)" no painel de estoque. Agora é possível exportar o estoque de cada loja separadamente e subir cada arquivo indicando qual loja é — sem precisar fazer PROCX no Excel para combinar as lojas.

---

## Versão inicial
- PCP Semanal com 8 semanas, estoque de segurança, produção sugerida e líquida
- Programação por dia (salgados em dias alternados)
- PCP por loja com mix e envio diário
- Auditoria CMV
- Persistência entre sessões (salvar/restaurar)
