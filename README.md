# PCP Tortelê — Web

Planejamento e controle de produção da Tortelê, direto no navegador. Sem
instalação, sem servidor — você sobe as planilhas de vendas e o site calcula
a produção sugerida da semana.

## Como usar (equipe, dia a dia)

1. Abra o link do site (ver Vercel após o deploy).
2. Arraste as exportações de vendas do sistema (uma por loja, .xlsx).
   Confirme a loja de cada arquivo no seletor.
3. Opcional: suba estoque atual, categorias e combos nos quadros ao lado.
4. Ajuste nível de serviço, janela de semanas e data de referência.
5. Navegue pelas abas: PCP Semanal, Programação, PCP por Loja, Auditoria CMV.
6. Use Exportar tudo (Excel) para baixar os resultados.
7. Use Salvar neste navegador para não perder os dados ao fechar a aba
   — isso fica salvo só neste computador/navegador, não é compartilhado
   com o resto da equipe.

## Metodologia

Mesma lógica da planilha original:

- Média das últimas N semanas completas (segunda a domingo).
- Estoque de Segurança = Z x Desvio Padrão, onde Z depende do nível de
  serviço escolhido (80/90/95/99%).
- Produção Sugerida = arredondar para cima (Média + Estoque de Segurança).
- Produção Líquida = máximo entre 0 e (Sugerida - Estoque Atual).
- Rateio por loja e por dia da semana usa o mix histórico das últimas
  4 semanas completas.
- Categorias "Salgado" entram na regra de produção em dias alternados
  (Segunda cobre Terça+Quarta, Quarta cobre Quinta+Sexta, Sexta cobre
  Sábado+Domingo, Sábado cobre a Segunda seguinte).

## Rodando localmente

npm install e depois npm run dev

## Gerando a versão de produção

npm run build

Os arquivos finais ficam na pasta dist. Qualquer serviço de hospedagem
estática (Vercel, Netlify, GitHub Pages, Cloudflare Pages) publica esse
projeto sem configuração adicional — é um app Vite mais React padrão.

## Importante sobre os dados

Todo o processamento roda no navegador de quem está usando. Nada é
enviado para nenhum servidor. Isso também significa que cada pessoa
vê seus próprios dados salvos — se a equipe toda precisa enxergar a
mesma base compartilhada, isso exige evoluir o projeto com um banco de
dados e login.
