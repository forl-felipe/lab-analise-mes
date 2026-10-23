# Monitoramento S.E. — Espessador × Batch (levantamento)

Fase 0 do projeto: leitura da planilha atual (`Monitoramento SE - Após fator ESP 6.2.xlsx`), das anotações
do caderno, do Word "Dados para montar report" (prints do MES) e da lista de fases-item/tags, para montar uma
planilha automática no mesmo modelo do **Comparativo Interlab Germano × Ubu** (`comparativo_interlab/`, branch
`claude/peaceful-cray-dmfxkz`): botão **Buscar no MES**, histórico, gráficos prontos e Configurações com as tags.

O dicionário completo (coluna da planilha → tela/receita/item do MES → janela → tag) está em
[`dicionario_campos.csv`](dicionario_campos.csv).

## O que a planilha atual faz

A aba que vai para a gerência é **Média ponderada**: uma linha por dia, com os gráficos a partir de 09/10/2026
(intervalo fixo `256:400`, ajustado à mão).

| Gráfico | Barras | Linhas |
|---|---|---|
| SE - Espessador C2 + C3 x Batch EB 6 | K = EB6 − ESP(C2+C3) | F = ESP C2+C3 ponderado · I = Batch EB6 |
| SE - Batch EB4 + ESP C3 x Batch EB 6 | J = EB6 − (EB4+ESP C3) | H = EB4+ESP C3 ponderado · I = Batch EB6 |
| % LAMA NO BATCH | I = Batch EB6 · L = TK01 · M = TK02 | R = % lama |

Fórmulas (conferidas com o dia 18/08/2026, que é o exemplo dos prints do Word):

- F = (ESP C3 · Prod C3 + ESP C2 · Prod C2) / (Prod C3 + Prod C2) → 2266,07
- H = (Batch EB4 · Prod C2 + ESP C3 · Prod C3) / (Prod C3 + Prod C2) → 2155,21
- J = EB6 − H · K = EB6 − F · R = Lama / (Conc C3 + Conc C2)

As demais abas são estudos antigos ou parados: *SE - ESP x BATCH* (horário, parou em 25/02), *Média aritmética*
(parou em 24/02), *SE Tanques* (2 leituras/dia, parou em 25/02), *Média ponderada - separado* (março, com
"Aumento SE" = EB6 − (0,98·ESP C3 + %lama·10625)), *Mineralogia* (HP/HE, fev) e *SE LDP*.

## Horários (caderno × Word)

| Dado | Janela | Observação |
|---|---|---|
| Produção C2/C3 e Alimentação do espessador (S.E.) | **D−1 23:00 a D 21:00** | defasagem de 2 h em relação ao batch (a aba antiga já comparava ESP h−2 × batch h) |
| Batch EB4 / EB6 (S.E.) | **D 00:00 a D 23:59** | o MES traz 13 resultados (01:00 … 23:00 e o fechamento 23:59); a média do report é a de todos |
| Tanques TK01/TK02 EB6 | D 00:00 a D 23:59 | — |
| DVR (concentrados e lama) | D 00:00 a D 23:59 | valor diário gravado às 23:59 |

## Confronto: o que bate e o que falta

**Confirmado** (caderno = Word = lista de tags = valor na planilha de 18/08):
Produção C2 (Ritmo Concentrador II), Alimentação Espessador C2 e C3, Batch EB6 (2184,08 no MES → 2184),
Concentrado C3 (23979) e C2 (19601, "Esp. Conc DVR"), lama do Boletim DVR.

**Pendências para fechar antes de construir**

1. **Ritmo Produção Concentrador III** (Produção C3) não está na lista de tags (só o II, `M0239010000-0596-HHCC`).
   Hipótese pelo padrão C2 sem "4" / C3 com "4": `M4239010000-0596-HHCC`.
2. **Produção é soma ou média?** Os valores da planilha (~25 000 t) indicam a **soma** do ritmo horário na janela.
3. **Batch EB4 e EB6** não estão na lista. Do projeto do batch: EB6 = `M4250010000-0017-ABLQ` (Mineroduto 3, já
   validada); EB4 = `M250010000-17-ABLQ` (Mineroduto 2, ainda não usada lá).
4. **TK01 e TK02 EB6**: sem tela no Word e sem tag.
5. **Lama C2**: o Word mostra *00-G2-DD BALANCO DVR* e o caderno manda olhar *25-G2-HH LAMA (USINA)*. Na planilha a
   lama é lançada como soma (ex.: `=412+50`) — C3 + C2.
6. **Concentrado C3** ("Beneficiamento - Concentrado C3 (tms)"): sem tag clara na lista. Candidata
   `M4239010000-1445-DDDVRCC` (Concen C3 transf Mineroduto 3). `M4224010000-0317-DDDVRCC` é alimentação de ROM.
7. **Esp. Conc DVR**: há duas tags (`M0238…-1445-DDDVR` e `M4238…-1445-DDDVR`); conferir qual é a do boletim.
8. **S.E. do espessador C3** usa o item 2888 (C2 usa 17): conferir se é a mesma análise.
9. **Fator** (180 → 200 → 160) é manual: o que significa e quando muda.
10. **Tags da lista sem uso na planilha**: Lama Incorporada Filtragem, Rejeito Arenoso Lama Incorp, Lama para Cava
    Sul, Benef ROM Aliment C3, Concen C3 transf MD3 e o PF Elutriação (S.E., FeO, SiO₂, P — as mesmas do PFUF do
    batch). Entram no novo report?

## Proposta da planilha (mesmo padrão do batch)

- **Dia**: data do report → *Buscar no MES* faz as consultas `GetCalculationValues` (horárias D−1 23:00–D 21:00,
  batch D 00:00–23:59 em janelas de 1 h como o "Germano em sequência", DVR diário) e mostra a conferência.
  Também busca um **intervalo de datas** para preencher dias atrasados.
- **Report**: os 3 gráficos acima com os últimos N dias (sem ajustar intervalo à mão) e *Copiar imagem*.
- **Histórico**: `tbl_Dias`, uma linha por dia (inclui a base atual importada de 28/01 a 15/10/2026).
- **Configurações** (com senha): tags, servidor, mapas, janelas e fator.
- Reaproveita do batch: consulta com troca de mapa/variante de tag, contador de tempo, aviso de `#NOME?`, padrão
  visual Samarco, gerador Python + `vba/*.bas`.
