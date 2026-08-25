# Histórico de versões — PCP Tortelê Web

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
