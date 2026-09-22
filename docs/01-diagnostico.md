# Diagnóstico — estado do projeto

Levantado lendo o `.pbix` (modelo descomprimido + JSON dos 60 visuais) e a base
`Base_PowerBI_Gestao_Equipamentos.xlsx`. Todo número aqui é medido, não estimado.

## Volume real dos dados

| Tabela | Carregadas | Com conteúdo | Observação |
|---|---:|---:|---|
| `tbl_Calibracao` | 214 | 208 | 6 linhas em branco do ListObject |
| `tbl_Equipamentos` | 34 | 34 | |
| `tbl_Inspecoes` | 200 | **5** | 195 linhas em branco no modelo |
| `tbl_Medicoes_Operacionais` | 9 | 3 | |
| `tbl_Meta_Inspecoes` | 9 | 5 | |
| `tbl_Intervencoes` | 3 | 3 | |
| `Tbl_Paradas` | 1 | 1 | |
| `tbl_Notas` | 0 | **0** | base de 2 cards do painel |
| `tbl_Sobressalentes` | — | 27 | existe na planilha, nunca carregada |

## Chaves

| Relacionamento | Casam | Total |
|---|---:|---:|
| `tbl_Calibracao[Tag]` → `tbl_Equipamentos[Tag]` | **0** | 214 |
| `tbl_Inspecoes[Equip_Tag]` → `tbl_Equipamentos[Tag]` | **0** | 5 |
| `DimCalendario` | sem relacionamento | |
| `DimTecnicos` | sem relacionamento | |
| `tbl_Meta_Inspecoes` | sem relacionamento | |

Padrões de TAG por domínio: `U03-02CR001` (equipamentos) · `66 AN 13`, `449212`
(instrumentos) · `AM-CT-01` (inspeções). Três convenções independentes.

## Indicadores que mostram número errado hoje

| Indicador | Mostra | Deveria | Causa |
|---|---|---|---|
| `% Conformidade` | 63,5% | 57,0% | denominador exclui os 22 "Sem Data" |
| `Equipamentos Críticos` | 34 | — | `Crítico` = "Sim" nas 34 linhas |
| `OMs Abertas` | 0 | vazio | `COUNTROWS(tbl_Notas) + 0` sobre tabela vazia |
| `Lead Time Médio` | 0 | vazio | `COALESCE(…, 0)` |
| Inspeções por responsável | todas as barras iguais | — | `DimTecnicos` sem relacionamento |
| `Cumprimento Meta` | soma de todas as metas (50) | por responsável | `tbl_Meta_Inspecoes` sem relacionamento |
| Velocidade por equipamento | contagem de texto | média de velocidade | coluna não tipada |
| `Nao Conformidades` | soma de texto | soma numérica | `Qtd_NC` não tipada |
| `CriticidadeCalibracao` | 22 em branco | "Sem Data" | `null < 0` em M devolve null, não false |

## Qualidade das datas em `tbl_Calibracao`

**`Última Calibração`** — seis formatos concorrentes na mesma coluna:

| Formato | Linhas |
|---|---:|
| `dd/mm/aaaa` | 100 |
| `mm/aa` ou `mm/aaaa` — **sem dia** | 80 |
| texto livre (`NOVA`, `nova`, `SEM `) | 10 |
| `dd/mm/aa` | 6 |
| serial do Excel | 6 |
| vazio | 6 |

**`Próximo Vencimento`** — 92% em bom estado (191 seriais do Excel). Os 17
restantes: `NOVA`/`nova` (8), vazio (6) e dois casos em que o texto
*"SEM CALIBRAÇÃO"* transbordou da coluna anterior, deixando `CALIBRAÇÃO` e
`CALIBR` dentro de uma coluna de data.

Há também registros com `Última Calibração` **no futuro** — até 15/07/2027.

O parser de `powerquery/05-fx_ParseData.pq` recupera **192 de 208** em ambas as
colunas e reproduz os status atuais (Em Dia 122 · Vencido 65 · A Vencer 5),
eliminando as 6 linhas fantasma.

## Modelo

- **9 tabelas de data automáticas** (`DateTableTemplate_*` + 8 `LocalDateTable_*`)
  enquanto `DimCalendario` existe com **uma única coluna** e nenhuma relação
- **duas tabelas de medidas** (`_Medidas`, `_Medidas_Inspecoes`), nenhuma medida
  com pasta de exibição, nenhuma com formato definido
- **3 medidas resíduo** com prefixo `DAX` colado junto da fórmula
- **2 duplicatas funcionais**: `Qtd Inspecoes por Criticidade` = `Inspecoes
  Realizadas`; `Calibracoes Criticas` = `Calibracoes A Vencer` (mesma definição,
  cards diferentes na mesma página)
- **1 relacionamento bidirecional** (`tbl_Notas`), único assim no modelo
- **12 consultas** abrindo separadamente o mesmo arquivo do SharePoint

## Layout

Sobreposições de visuais, medidas no JSON — **nas quatro páginas**:

| Página | Sobreposição | Ocupação do canvas |
|---|---|---:|
| VISÃO GERAL | rosca × barras, 1px | 46% |
| CALIBRAÇÕES | barras × rosca, **17px** | 64% |
| MANUTENÇÃO | card × card, **15px** | 35% |
| INSPEÇÕES | colunas × colunas **31px**, rosca × tabela 9px | 47% |

Outros desalinhamentos: cards da VISÃO GERAL em `y = 154…154,155`; cards da
CALIBRAÇÕES em `x = 0,1,0,0,0,2`; espaçamento dos cards da MANUTENÇÃO em
180/196/165px; cabeçalhos de 150/196/166/175px; navegação em `x=1320` em três
páginas e `x=1044` na quarta.

Formatação aplicada visual a visual: o card `% Conformidade` tem raio de borda,
sombra customizada e espaçamentos ajustados; o card ao lado tem apenas fundo e
sombra padrão.
