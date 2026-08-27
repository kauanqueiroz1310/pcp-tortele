# Histórico de versões — PCP Tortelê Web

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
