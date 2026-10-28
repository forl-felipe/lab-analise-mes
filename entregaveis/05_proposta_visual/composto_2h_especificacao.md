# Composto de 2 h da prensa automática: o que precisa existir no PIMS

## Por que

O laboratório analisa uma amostra **composta de 2 h**, formada por 4 cortes (um a cada 30 min). A prensa automática ensaia **cada corte**. Comparar o último ensaio da prensa com o resultado do laboratório mistura uma amostra pontual com uma composta e gera divergências que não existem.

A comparação correta é **composto com composto**: a média dos 4 ensaios da prensa nos mesmos cortes que formaram o composto do laboratório.

## O que a tela mostra (versão 6)

| Painel | Conteúdo | Origem |
|---|---|---|
| Média prensa automática | Último ensaio, desvio padrão e horário | Tags existentes (`307-GI-001`, `307-GI-002`) |
| Composto de 2 h: prensa x laboratório | Composto da prensa, composto do laboratório, fim da janela e diferença | Laboratório: tag existente do MES. Prensa e diferença: **tags calculadas** |
| Tendência | Cada ensaio (linha fina), composto da prensa (linha grossa azul) e composto do laboratório (linha grossa branca) | Ensaios e laboratório: tags existentes. Composto da prensa: **tag calculada** |
| Cortes do composto atual | Hora de cada um dos 4 cortes e o resultado da prensa | **Tags calculadas** |

## Tags calculadas a pedir ao administrador do PIMS (Aspen Calc)

Nomes provisórios; podem seguir o padrão da área.

| Tag | Cálculo | Quando calcula | Horário gravado |
|---|---|---|---|
| `307-GI-001-C2H` | Média de `307-GI-001` **dos cortes com resultado** na janela de 2 h (1 a 4 cortes) | A cada novo ensaio da janela (o valor vai se completando) e no fim da janela | Fim da janela (horário do último corte + 30 min) |
| `307-GI-001-N2H` | Quantidade de cortes com resultado na janela (0 a 4) | Junto com `307-GI-001-C2H` | Fim da janela |
| `307-GI-001-DIF2H` | `307-GI-001-C2H` menos o resultado do MES (`M4710050020-0031-HHLFU`) da mesma janela | Quando entra o resultado do MES | Fim da janela |
| `307-GI-001-CORTE1` a `CORTE4` | Resultado de `307-GI-001` de cada corte da janela atual; sem valor (ou qualidade Bad) quando o corte não teve resultado | A cada novo ensaio | Horário do corte |

Pontos que o cálculo precisa respeitar:

1. **Defasagem do PIMS:** o ensaio do corte das 15:00 entra no PIMS às 15:20 (20 min depois). A janela deve ser montada pelo **horário do corte**, não pelo horário de entrada no PIMS. Se a tag guardar o horário de entrada, o cálculo subtrai os 20 min.
2. **Janela incompleta:** se a prensa falhar e a janela tiver só 1, 2 ou 3 ensaios, o composto é a **média dos ensaios que existem**, e `307-GI-001-N2H` informa quantos foram. Com 0 ensaios, o composto fica sem valor.
3. **Amostragem em manual:** se 07CR001 ou 07CR002 estiverem em Manual na janela, a comparação não é válida (mesma regra da planilha de comparação).
4. **Valores zero:** a tag `307-GI-001` volta a 0 entre os ensaios (aparece na tendência). Valores 0 e ensaios com qualidade diferente de Good não entram na média.

## Antes de pedir as tags: teste rápido no editor

Algumas versões da tendência do Aspen permitem mostrar a pena como **média por período** (agregado). Na tendência principal, abra as propriedades da pena `307-GI-001` e procure algo como "Agregado", "Média" ou "Período". Se existir, uma pena com **média de 2 h** desenha o composto da prensa na tendência sem precisar de tag calculada. Os números do painel e a tabela de cortes continuam dependendo das tags calculadas.

## Perguntas para fechar a regra

1. Quais cortes formam o composto do laboratório? Exemplo: o composto "das 16:00" usa os cortes de 14:00, 14:30, 15:00 e 15:30, ou de 14:30 a 16:00?
2. O horário do resultado no MES é o fim da janela, o horário da análise ou o do lançamento?
3. A defasagem de 20 min entre o corte e a entrada no PIMS é fixa?
