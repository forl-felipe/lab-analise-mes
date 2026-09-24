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

## 3. Mapa de aferições: a regra do dia certo

(v19: **cada ensaio tem um dia certo**, o dia da rotina. A marca fica
**sempre na coluna do dia certo**. Substitui a regra de prazo da v18.)

| Marca | Significado | Quando aparece |
|---|---|---|
| **✔ azul** | feito no dia certo | há registro no próprio dia da rotina (Blaine: os 2 turnos) |
| **amarelo, sem texto** | feito com atraso | não houve registro no dia certo, mas houve **depois** dele e **antes da próxima data da rotina** |
| **✔ azul-claro** | feito fora do dia programado | dia que não é da rotina e tem registro: é o dia em que a atrasada foi feita, ou um registro extra. Umidade e Fisher, que não têm rotina, só mostram esta marca |
| **✖ laranja** | não feito | passou o dia certo e não há registro depois dele (até a próxima data da rotina) |
| **○** | programado | dia certo de hoje (ainda sem registro) ou que ainda vai chegar |
| **½** | incompleto | Blaine com só 1 dos 2 turnos |
| **cinza, vazio** | nada | não é dia da rotina e não houve registro |

A primeira linha do mapa mostra o **dia da semana** de cada coluna, e
**hoje** aparece em amarelo.

**Exemplo: tambor de abrasão (toda segunda) em setembro/2026, com hoje = 24/09.**

| Dia certo | O que aconteceu | O mapa mostra |
|---|---|---|
| 07 (seg) | feito na quarta 09 | 07 **amarelo** · 09 **✔ azul-claro** |
| 14 (seg) | feito no dia | 14 **✔ azul** |
| 21 (seg) | sem registro até hoje | 21 **✖**. Se for feito no dia 25, o 21 vira **amarelo** e o 25 **✔ azul-claro** |
| 28 (seg) | ainda não chegou | 28 **○** |

**Até quando o atraso conta.** Até a véspera da **próxima** data da
rotina. No tambor, o dia 07 aceita registro até 13/09. Feito só no dia
14, o dia 07 fica ✖ e o registro conta para o próprio dia 14.

**Onde ver a regra de cada ensaio.** Na tabela **"Regras de cada ensaio ·
o dia certo"**, embaixo do mapa. Ela mostra:

- quando o ensaio deve ser feito;
- os dias certos do mês (ex.: 07, 14, 21, 28);
- o resultado no mês (ex.: "1 de 3 · 1 com atraso · 1 não feitas").

**Clicar no nome de um ensaio no mapa** deixa a tabela só com a regra
dele; clicar de novo volta a mostrar todos. O clique no mapa não filtra
o resto da página.

**A rotina de cada ensaio** foi tirada do cabeçalho de cada aba da planilha:

| Ensaio | Dia certo | Dias (1 = seg … 7 = dom) |
|---|---|---|
| Tambor de abrasão | toda segunda, turno da noite | 1 |
| Peneiradores | terça, quinta e sábado | 2, 4, 6 |
| Compressão | toda segunda | 1 |
| Tamb 5 × 15 kg | toda segunda | 1 |
| Alpine | segunda, quarta e sábado | 1, 3, 6 |
| Blaine | todos os dias, 2 turnos | 1 a 7 |
| Granulometria | 1 por semana: segunda | 1 |
| Umidade e Fisher | sem rotina definida | nunca recebem ✖ |

**Os mesmos números no resto da página.** O cartão **"Rotina · feitas
no dia certo"** e as colunas "No dia certo", "Com atraso" e "Não feitos"
da tabela Situação por ensaio seguem esta regra. Também seguem a linha
de rotina do Resumo automático e o "rotina x%" da Visão Geral. A meta é
a "Meta Aderência Rotina (%)" (90% se não houver a linha em Configuração).

**De onde vêm as datas:** coluna **Data** que o operador preenche em
cada aba, levada para a `BD_Afericoes`:

- Tambor: A6:A34
- Alpine: A7:A21
- Blaine: A6:A71
- Umidade: A5:A18
- Compressão: B6:B23
- Peneiradores: AC de cada bloco
- Granulometria e Tamb 5×15: campo Data de cada bloco

**Limites da regra, para reportar com segurança:**

- **O ✔ significa que o ensaio foi registrado no dia, não que passou.**
  Se passou ou falhou está na página "Detalhe por ensaio".
- **Uma linha por ensaio:** basta um equipamento do ensaio ter registro
  no dia para o dia contar como feito.
- **Ensaio feito mas não lançado** aparece como ✖. O painel só enxerga o
  que está na planilha.
- **Data digitada errada** move o registro de dia. Em setembro há
  registros de peneiradores em domingo (06 e 13): vale conferir.
- **Turno da noite que passa da meia-noite:** vale a data lançada na
  planilha. Lance a data da segunda para contar no dia certo.

**Como ajustar:**

- **Um ✖ indevido:** lance a data na aba da planilha e salve. Na próxima
  atualização some.
- **Uma rotina ou o texto da regra:** Power BI Desktop > Transformar
  dados > **DimEnsaio** > Editor Avançado.
  - Coluna "Dias Semana": `"1,3,6"` = seg, qua e sáb. É ela que define o
    dia certo.
  - Coluna "Regra": o texto que aparece na tabela de regras.

## 4. Os demais números

| Número | Fonte |
|---|---|
| Calibração em dia, vencidas, a vencer | Base, aba Calibração (coluna Próximo Vencimento, contada a partir do dia da atualização). O **LDP não é mais monitorado** (v18): fica fora mesmo que alguma linha LDP volte à planilha |
| Disponibilidade e horas paradas | Base, aba Paradas (horas somadas no mês da data da parada) × nº de equipamentos × horas do mês |
| Inspeções | Base, aba de inspeções + planilha do Microsoft Forms |
| Aferições | pasta do SharePoint "Calibração Integrada mensal", aba oculta BD_Afericoes de cada planilha mensal |
