# Histórico de versões — PCP Tortelê Web

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
