# Medidas DAX

Crie uma tabela vazia para guardar as medidas (**Página Inicial > Inserir dados** > nome `_Medidas` > Carregar) e cole cada medida em **Nova medida**.
Os nomes de tabela são os das consultas: `tbl_Afericoes`, `tbl_Limites`, `Equipamentos`, `Lancamentos`, `Previstos`, `Calendario`.

## Tabela de datas (Modelagem > Nova tabela)

```DAX
Calendario =
ADDCOLUMNS(
    CALENDAR(DATE(2026, 10, 1), TODAY() + 60),
    "Ano", YEAR([Date]),
    "Mês", FORMAT([Date], "mmm/yyyy"),
    "MêsNum", YEAR([Date]) * 100 + MONTH([Date]),
    "Semana", WEEKNUM([Date], 2),
    "Dia da semana", FORMAT([Date], "ddd")
)
```
Marque como tabela de datas (coluna `Date`).

## Relacionamentos (Modelo)

| De | Para | Cardinalidade |
|---|---|---|
| `Calendario[Date]` | `tbl_Afericoes[Data]` | 1 : * |
| `Calendario[Date]` | `Previstos[Data]` | 1 : * |
| `Calendario[Date]` | `Lancamentos[Data]` | 1 : * |
| `Equipamentos[TAG]` | `tbl_Afericoes[TAG]` | 1 : * |

## Resultados das aferições

```DAX
Resultados = COUNTROWS(tbl_Afericoes)

Conformes = CALCULATE([Resultados], tbl_Afericoes[Resultado] = "Conforme")

Não conformes = CALCULATE([Resultados], tbl_Afericoes[Resultado] = "Não conforme")

Informativos = CALCULATE([Resultados], tbl_Afericoes[Resultado] = "Informativo")

% Conformidade = DIVIDE([Conformes], [Conformes] + [Não conformes])

Lançamentos = DISTINCTCOUNT(tbl_Afericoes[Lançamento])

% Lançado fora do turno =
DIVIDE(
    CALCULATE(DISTINCTCOUNT(tbl_Afericoes[Lançamento]), tbl_Afericoes[Fora do turno] = "Sim"),
    [Lançamentos]
)

Último registro = MAX(tbl_Afericoes[Data])
```

## Agenda: aderência (mesma regra da aba Painel da planilha)

```DAX
Previstos = CALCULATE(COUNTROWS(Previstos), Previstos[Conta] = TRUE())

Realizados = CALCULATE(COUNTROWS(Previstos), Previstos[Realizado] = TRUE())

Não feitos = CALCULATE(COUNTROWS(Previstos), Previstos[Situação] = "Não feito")

Feitos com atraso = CALCULATE(COUNTROWS(Previstos), Previstos[Situação] = "Com atraso")

Aderência = DIVIDE([Realizados], [Previstos])

Situação do ensaio =
VAR naoFeitos = [Não feitos]
RETURN IF(naoFeitos = 0, "Em dia", "Pendente (" & naoFeitos & ")")
```

Para a **Gestão por Letras**, use `Previstos[Letra]` no eixo ou na tabela (letra que estava escalada no turno).

## Equipamentos

```DAX
Equipamentos ativos = COUNTROWS(Equipamentos)

Equipamentos bloqueados = CALCULATE(COUNTROWS(Equipamentos), Equipamentos[Situação] = "Bloqueado")

% Equipamentos liberados = DIVIDE(CALCULATE(COUNTROWS(Equipamentos), Equipamentos[Situação] = "Liberado"), [Equipamentos ativos])

Calibrações vencidas = CALCULATE(COUNTROWS(Equipamentos), Equipamentos[Status] = "Bloqueado · Calibração vencida")

Calibrações a vencer 30 dias =
CALCULATE(
    COUNTROWS(Equipamentos),
    Equipamentos[ControlaCalibracao] = TRUE(),
    Equipamentos[ProximaCalibracao] >= TODAY(),
    Equipamentos[ProximaCalibracao] <= TODAY() + 30
)

Sem data de calibração = CALCULATE(COUNTROWS(Equipamentos), Equipamentos[Status] = "Liberado · Sem data de calibração")
```

## Causas de não conformidade

```DAX
NC com causa = CALCULATE([Não conformes], NOT ISBLANK(tbl_Afericoes[Causa]), tbl_Afericoes[Causa] <> "")
```
Gráfico de barras: eixo `tbl_Afericoes[Causa]`, valor `[Não conformes]`. Para peneiradores, eixo `tbl_Afericoes[ItensNaoOK]`.

## Sugestão de páginas

1. **Visão geral**: cartões (% Conformidade, Aderência, Equipamentos bloqueados, Calibrações vencidas), aderência por ensaio (barras), resultados por dia (colunas empilhadas por Resultado).
2. **Agenda e letras**: matriz Ensaio × Letra com Previstos, Realizados e Aderência. Tabela de "Não feito" com data e turno.
3. **Tendência por equipamento**: segmentação por TAG; linha de `Valor` com linhas constantes de `Lim. Inferior`/`Lim. Superior`.
4. **Calibração**: tabela de Equipamentos com Status, Próxima calibração e Motivo; cartões de vencidas e a vencer.
