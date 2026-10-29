# Painel horário do laboratório: fontes

Levantamento das fontes do painel `Dados_BI_horário.pbix` (dados de 27/10/2026),
ponto de partida para o painel do laboratório. O `.pbix` original não foi alterado
e não está guardado aqui.

- `Catalogo_Fontes_Painel_Horario.xlsx`: o catálogo, com quatro abas.
  - **Leia-me:** conexão, colunas da view e como ler os nomes.
  - **Elementos:** um elemento por linha.
  - **Limites por campanha.**
  - **Consultas e medidas:** M e DAX.
- `catalogo_elementos.json`: os mesmos dados, para gerar consultas e telas.

**Fonte:** `dbo.vw_ingredientinput_gpu` no lakehouse `lh_mes_horario` (SQL endpoint do Fabric).
Ela tem 57 elementos: 47 de qualidade (LQU e LFU) e 10 de produção (U03 e U04).
A view só guarda 2 dias de lançamentos, ou seja, ontem e hoje.
