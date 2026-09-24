# 14 — Metas, limites e de onde vem cada número

Guia para responder "de onde saiu isso?". Para cada número: onde ele nasce,
onde você altera e o que acontece no painel.

## 1. Metas do painel (as 4 frentes da Visão Geral)

| Meta | Valor hoje | Onde está | Como alterar |
|---|---|---|---|
| Disponibilidade | 95% | Base `Base_PowerBI_Gestao_Equipamentos.xlsx`, aba **Configuração**, linha `Meta Disponibilidade (%)` | Troque o valor na aba e atualize o painel |
| Disponibilidade por grupo | 95% | Base, aba **Grupos**, coluna `Meta Disponibilidade (%)` | Um valor por grupo |
| Inspeções no mês | 100 | Base, aba **Configuração**, linha `Meta Inspeções/Mês` | Troque o valor na aba |
| Calibração em dia | 90% | **valor-padrão do painel** (não está em planilha) | Acrescente na aba Configuração a linha `Meta Calibração em Dia (%)` com o valor (ex.: 90) |
| Aferições conformes | 95% | **valor-padrão do painel** | Acrescente na aba Configuração a linha `Meta Aferições Conformes (%)` |
| Aderência à rotina | 90% | **valor-padrão do painel** | Acrescente na aba Configuração a linha `Meta Aderência Rotina (%)` |

Como funciona:

- O painel procura cada linha na aba Configuração. Achou: usa o valor.
  Não achou: usa o padrão da tabela.
- Os três padrões foram as metas que eu propus no mockup.
- Para que elas passem a ser **oficiais e rastreáveis**, acrescente as 3
  linhas na aba Configuração.
- O nome na coluna Parâmetro precisa ser **exatamente** o da tabela, e o
  valor vai em %, sem o sinal (90, 95…).

**Situação da frente** (● na meta · ▲ atenção · ◆ crítico): regra fixa do
painel.

| Frente | ● Na meta | ▲ Atenção | ◆ Crítico |
|---|---|---|---|
| Disponibilidade e calibração | valor ≥ meta | até 10% abaixo da meta | mais que isso |
| Aferições | valor ≥ meta | até 5 pontos abaixo | mais que isso |
| Inspeções | realizado ≥ esperado até hoje | até 10% abaixo do esperado | mais que isso |

Nas inspeções, o esperado até hoje é a meta do mês × dia de hoje ÷ dias do
mês.

## 2. Limites técnicos das aferições: aba `BD_Limites`

A aba oculta **BD_Limites**, dentro de cada planilha mensal CALIBRAÇÃO,
guarda nominal, limite inferior, limite superior, tolerância e voltas de cada
ensaio. As fórmulas da aba oculta `BD_Afericoes` leem dali para dizer se cada
registro está **Conforme** ou **Não conforme**.

- **Mudou um limite na BD_Limites e salvou a planilha?** Na próxima
  atualização o painel já mostra o resultado novo. O Power BI lê o resultado
  que a planilha gravou.
- **Vale só para aquela planilha (aquele mês).** O mês seguinte, criado com
  "Salvar como", herda os limites.
- **São coisas diferentes:** a BD_Limites decide se *cada aferição* passou.
  A meta de "Aferições conformes" (95%) decide se *o mês* está bom.

**Como saber se a planilha da pasta é a versão atual:** abra a BD_Limites.

- A versão atual tem a linha **COM16 com nominal 360** e as linhas
  **V_TA05 a V_TA07** (voltas dos tambores).
- Se aparecer 355 e não houver as linhas V_TA, é a primeira versão: troque
  pela planilha mais recente.

## 3. Mapa de aferições (ensaio × dia): de onde vem cada marca

| Marca | Regra | De onde vem |
|---|---|---|
| **✔ azul (feito)** | existe **pelo menos um registro** daquele ensaio naquela data | coluna **Data** que o operador preenche em cada aba (Tambor A6:A34, Alpine A7:A21, Blaine A6:A71, Umidade A5:A18, Compressão B6:B23, Peneiradores AC de cada bloco, Granulometria e Tamb 5×15 no campo Data de cada bloco), levada para a `BD_Afericoes` |
| **✖ laranja (não feito)** | o dia **estava na rotina**, **já passou** (até ontem) e **não há nenhum registro** do ensaio naquela data | rotina da tabela **DimEnsaio** do painel (abaixo) |
| **em branco** | o dia não estava na rotina e não houve registro | — |

A rotina foi tirada do cabeçalho de cada aba da planilha:

| Ensaio | Rotina no painel | Dias (1 = seg … 7 = dom) |
|---|---|---|
| Tambor de abrasão | toda segunda | 1 |
| Peneiradores | ter, qui e sáb | 2, 4, 6 |
| Compressão | toda segunda | 1 |
| Tamb 5 × 15 kg | toda segunda | 1 |
| Alpine | seg, qua e sáb | 1, 3, 6 |
| Blaine | início de cada turno (todos os dias) | 1 a 7 |
| Granulometria | 1 por semana | não marca ✖ por dia |
| Umidade e Fisher | sem rotina definida | não marca ✖ |

**Limites da regra, para reportar com segurança:**

- **O ✔ significa que o ensaio foi registrado no dia, não que passou.** Se
  passou ou falhou está na tabela "Não conformidades por equipamento".
- **Blaine: basta um turno com registro para o dia ficar ✔.** A contagem
  exata por turno está na coluna "Rotina" da tabela Situação por ensaio.
- **Ensaio feito mas não lançado aparece como ✖.** O painel só enxerga o
  que está na planilha.
- **Data digitada errada:** o registro vai para o dia errado no mapa.

**Como corrigir:**

- **Um ✖ indevido:** lance a data na aba da planilha e salve. Na próxima
  atualização vira ✔.
- **Uma rotina:** Power BI Desktop > Transformar dados > **DimEnsaio** >
  Editor Avançado > coluna "Dias Semana". Exemplo: `"1,3,6"` = seg, qua e
  sáb.

## 4. Os demais números

| Número | Fonte |
|---|---|
| Calibração em dia, vencidas, a vencer | Base, aba Calibração (coluna Próximo Vencimento, contada a partir do dia da atualização) |
| Disponibilidade e horas paradas | Base, aba Paradas (horas somadas no mês da data da parada) × nº de equipamentos × horas do mês |
| Inspeções | Base, aba de inspeções + planilha do Microsoft Forms |
| Aferições | pasta do SharePoint "Calibração Integrada mensal", aba oculta BD_Afericoes de cada planilha mensal |
