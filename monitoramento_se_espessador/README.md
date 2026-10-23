# Monitoramento S.E. — Espessador × Batch

Planilha `Monitoramento_SE_Espessador_Batch.xlsm`: versão automática da `Monitoramento SE - Após fator ESP 6.2.xlsx`
(aba *Média ponderada*). O botão **Buscar no MES** traz os resultados do dia, **Salvar no histórico** grava o dia e o
**Report** monta os 3 gráficos gerenciais dos últimos N dias, prontos para copiar como imagem. Mesmo padrão do
Comparativo Germano × Ubu (`comparativo_interlab/`): visual Samarco, consulta `GetCalculationValues`, gerador Python + VBA.

## Uso no dia a dia

1. Aba **Dia** → data do report → **Buscar no MES**. Confira os valores (células amarelas) e o Nº de resultados.
2. Lance o que não tem tag: **TK 01 e TK 02 - EB6** (e o Fator, se mudar; dia novo repete o do dia anterior).
3. **Salvar no histórico** → o Report passa a terminar nesse dia.
4. Aba **Report** → **Copiar imagem** → colar no e-mail/Teams.

Dias atrasados: na aba Dia, **Buscar um período** (De/Até) busca e salva cada dia, mantendo TK, Fator e Observação
já lançados no histórico.

## Abas

| Aba | Para quê |
|---|---|
| **Report** | Cartões do último dia (EB6, ESP C2+C3, as duas diferenças, % lama) com a média do período e os gráficos: *Espessador C2 + C3 × Batch EB6*, *Batch EB4 + Espessador C3 × Batch EB6* (barras = diferença, linhas = S.E.) e *% de lama no batch* (barras EB6/TK01/TK02, linha % lama). Data final e Nº de dias (2 a 60) na barra de cima. |
| **Dia** | Busca, conferência e lançamento do dia; resultado do dia (ponderadas, diferenças, % lama); busca de período. |
| **Histórico** | Um dia por linha. A base de 28/01 a 15/10/2026 (254 dias) veio da planilha antiga. Colunas cinza são calculadas. |
| **Configurações** (oculta, senha `Samarco@Lab`) | Tag, servidor, mapa, janela, agregação e faixa válida de cada item; janelas de horário; fator padrão; fonte (MES ou SIMULAÇÃO). |
| Dados_MES, _Graf (ocultas) | Consultas ao Aspen e séries dos gráficos. |

## Regras

- **Janela E (espessador)**: D−1 23:00 a D 21:00 — produção C2 e S.E. dos espessadores (2 h antes do batch).
- **Janela D (dia)**: D 00:00 a 23:59 — batch EB4/EB6, tanques e DVR (valor diário gravado às 23:59).
- Consultas de **1 h** começando 30 min antes do horário (00:30–01:30 pega o resultado das 01:00; 23:30–00:30 pega o
  fechamento das 23:59). Agregação por item: **Soma** (produção C2), **Média** (S.E.), **Último** (DVR).
- ESP C2 + C3 = (ESP C3 × Prod C3 + ESP C2 × Prod C2) / (Prod C3 + Prod C2) ·
  EB4 + ESP C3 = (EB4 × Prod C2 + ESP C3 × Prod C3) / (Prod C3 + Prod C2) · diferenças = EB6 − ponderada ·
  % lama = (lama C3 + lama C2) / (conc C3 + conc C2). Conferido com 18/08/2026 (2266,07 · 2155,21 · −82,07 · +28,79).
- Tag recusada pelo MES: tenta o mesmo código no outro formato (`-0596-` ↔ `-596-`); mapa recusado: tenta
  `IP_AnalogMap`, `IP_MESVALOR`, `IP_DiscreteMap`. O que o MES aceitar fica gravado em Configurações.

## Tags (aba Configurações)

| Item | Tag | Janela · agregação |
|---|---|---|
| Produção C3 | `M4239010000-1445-DDDVRCC` (Concen C3 transf Mineroduto 3) | D · Último |
| Produção C2 | `M0239010000-0596-HHCC` (Ritmo Produção Concentrador II) | E · Soma |
| S.E. espessador C3 / C2 | `M4238010000-2888-HHLQ` / `M0238010000-0017-HHLQ` | E · Média |
| S.E. batch EB4 / EB6 | `M250010000-17-ABLQ` / `M4250010000-0017-ABLQ` (do projeto do batch) | D · Média |
| TK 01 / TK 02 - EB6 | sem tag (lançar à mão) | — |
| Lama incorporada C3 / C2 | `M4239000000-0781-DDDVR` / `M0239000000-0781-DDDVR` | D · Último |
| Concentrado C3 | `M4224010000-0317-DDDVRCC` (no MES: Benef ROM Aliment C3) | D · Último |
| Concentrado C2 | `M0238010000-1445-DDDVR` (Esp. Conc DVR (tms)) | D · Último |

Levantamento completo (planilha antiga × caderno × Word × lista de tags): [`dicionario_campos.csv`](dicionario_campos.csv).

## A conferir na primeira busca real

- **Produção C3**: a tag informada é diária (DVR), por isso está na janela D com *Último*. Comparar com o histórico
  (18/08: 26 983 t) — se o valor vier diferente, ajustar janela/agregação em Configurações.
- **Nº de resultados**: EB6 deve dar 13 no dia (como na tela do MES) e o espessador 12. Se der o dobro, o MES está
  repetindo o valor nas janelas de 1 h.
- **Concentrado C2**: conferir se `M0238…-1445` é o valor do Boletim DVR (19 601 t em 18/08).
- **Fator**: só é guardado (não entra em nenhuma conta nem gráfico). Confirmar o que representa.

## Gerar de novo

```
cd build
python3 build_monitor.py            # gera ../Monitoramento_SE_Espessador_Batch.xlsm
python3 build_monitor.py --teste    # versão com o teste automático (LibreOffice)
python3 extrair_historico.py "Monitoramento SE - Após fator ESP 6.2.xlsx"   # refaz build/historico.json
```
