# Monitoramento S.E. — Espessador × Batch

Guia rápido de uso: [`guia/Guia_rapido_Monitoramento_SE.png`](guia/Guia_rapido_Monitoramento_SE.png).

Planilha `Monitoramento_SE_Espessador_Batch.xlsm`: versão automática da `Monitoramento SE - Após fator ESP 6.2.xlsx`
(aba *Média ponderada*). O botão **Buscar no MES** traz os resultados do dia, **Salvar no histórico** grava o dia e o
**Painel** monta o report gerencial dos últimos N dias, pronto para copiar como imagem. O Painel segue o mesmo padrão
visual da aba Painel do Comparativo Germano × Ubu (`comparativo_interlab/`): cabeçalho Samarco, quadro de status
OK/NOK, comentários automáticos e gráficos com barras por dia + barra escura da média, linha laranja da diferença,
◆ diferença média e faixa verde tracejada da tolerância.

## Uso no dia a dia

1. Aba **Dia** → dia dos dados (**ontem**: o DVR só entra no MES no fechamento do dia) → **Buscar no MES**.
   O dia já é **gravado no histórico** e o **Painel passa a terminar nele**. Confira valores, Nº de resultados e Situação.
2. Se corrigir algum valor ou o Fator à mão → **Salvar no histórico** de novo.
3. Comentários do painel: bloco **④** da aba Dia (texto automático editável, Sim/Não para entrar no painel, 3 tópicos
   livres, **Restaurar automáticos**; ao trocar a data, voltam ao automático). No Painel, **✎ Editar comentários**
   leva até lá. Tolerâncias: bloco **⑤**, uma por linha do quadro do Painel (vazio = "—"; diferenças: |valor| ≤ tol.;
   demais: |último dia − média do período| ≤ tol.). Padrão: ±150 nas duas diferenças.
4. Aba **Painel** → **Copiar imagem** (e **Copiar resumo** para o corpo do e-mail). Ao abrir a planilha, o Painel
   já está no último dia salvo.

Lançamento à mão: dá para digitar um dia direto na aba **Histórico** (data + valores nas colunas brancas); as
colunas cinza calculam sozinhas, o Painel lê o dia e a aba Dia abre o dia ao escolher a data.

Item sem resultado (Nº de resultados 0, em vermelho): confira no MES. Se o MES tiver o valor, aba **Diagnóstico** →
dia já fechado → **Rodar diagnóstico**: ele testa as formas da tag (`-0596-` ↔ `-596-`, `M0…` ↔ `M…`) com
`IP_MESVALOR` e `IP_AnalogMap` e grava em Configurações a que trouxer valor. Mostra tag, mapa, resultado, valor do
dia e os horários com valor.

## Abas

| Aba | Para quê |
|---|---|
| **Painel** | Report gerencial (padrão do Comparativo): quadro *Último dia / Média do período / Tolerância / Status* (EB6, ESP C2+C3, EB4+ESP C3, as duas diferenças, % lama), **comentários automáticos** (último dia, período, dias fora da tolerância, % lama máxima, observação do dia) e caixa *Produção e DVR*. Gráficos: *Espessador C2 + C3 × Batch EB6*, *Batch EB4 + Espessador C3 × Batch EB6* e *% de lama no batch × S.E. do EB6 e dos tanques*. Barra de cima (fora da imagem): data final, nº de dias (2 a 60), Último dia, Ir para o Dia, Atualizar painel, Copiar imagem, Copiar resumo. |
| **Dia** | ① dia dos dados, ② resultados do MES (item, valor, nº de resultados, janela) com a legenda de D / D−1 em datas, ③ resultado do dia, ④ comentários do painel, ⑤ tolerâncias e a última busca. |
| **Histórico** | Um dia por linha, de 28/01/2026 em diante (planilha anterior + buscas no MES). Colunas cinza são calculadas; dá para lançar dias à mão. |
| **Diagnóstico** | Só quando algum item vier sem valor (ver acima). |
| **Configurações** (oculta, senha `Samarco@Lab`) | Tag, servidor, mapa, janela, agregação e faixa válida de cada item; janelas de horário; fator padrão; fonte (MES ou SIMULAÇÃO). |
| Dados_MES, _Graf (ocultas) | Consultas ao Aspen (fórmula e resposta de cada consulta) e séries dos gráficos. |

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
- Status OK/NOK do Painel: por linha, com a tolerância do bloco ⑤ da aba Dia (vazio = sem status).
- Busca de **hoje**: a planilha avisa que o dia não terminou (DVR e fechamento do batch ainda não estão no MES) e
  pergunta antes de salvar.

## Tags (aba Configurações)

Conferidas pelo laboratório no MES em 26/10/2026 (exceto a S.E. do espessador C3). Exemplo de horários: o report de
26/10 usa os dados de 25/10 — produção e espessadores de 24/10 23:00 a 25/10 21:00; batch e DVR de 25/10 00:00 a 23:59.

| Item | Tag | Mapa | Janela · agregação |
|---|---|---|---|
| Produção C3 | `M4239010000-596-HHCC` (Ritmo Produção Concentrador III) | IP_MESVALOR | E · Média |
| Produção C2 | `M0239010000-0596-HHCC` (Ritmo Produção Concentrador II) | IP_AnalogMap | E · Média |
| S.E. espessador C3 / C2 | `M4238010000-2888-HHLQ` / `M0238010000-0017-HHLQ` | IP_AnalogMap | E · Média |
| S.E. batch EB4 / EB6 | `M250010000-17-ABLQ` / `M4250010000-17-ABLQ` | IP_MESVALOR | D · Média |
| TK 01 / TK 02 | `M239040000-17-TBLQ` / `M239050000-17-TBLQ` (Tk1 / Tk2 U2 - SE) | IP_MESVALOR | D · Média |
| Lama incorporada C3 / C2 | `M4239000000-0781-DDDVR` / `M0239000000-0781-DDDVR` | IP_AnalogMap | D · Último |
| Concentrado C3 | `M4238010000-1445-DDDVR` (Beneficiamento - Concentrado C3) | IP_AnalogMap | D · Último |
| Concentrado C2 | `M0238010000-1445-DDDVR` (Esp. Conc DVR (tms)) | IP_AnalogMap | D · Último |

Levantamento completo (planilha antiga × caderno × Word × lista de tags): [`dicionario_campos.csv`](dicionario_campos.csv).

## Histórico da validação no MES

- **26/10/2026 (1ª busca real, dia 26/10 = hoje)**: 5 de 10 itens com resultado (produção C2, espessadores C2/C3,
  batch EB4/EB6). Os 5 do DVR vieram vazios porque o DVR do dia só é gravado no fechamento (23:59). O MES recusou
  `IP_MESVALOR` nas tags HH/DVR e aceitou `IP_AnalogMap`; para o EB6 aceitou `M4250010000-17-ABLQ`. Esses ajustes
  viraram o padrão da planilha.
- **A conferir com um dia fechado** (aba Diagnóstico): DVR (produção C3, lama, concentrados) com valor às 23:59;
  EB6 com 13 resultados no dia; produção C3 parecida com o histórico (18/08: 26 983 t); concentrado C2 = Boletim DVR
  (19 601 t em 18/08).
- **26/10/2026 (conferência manual das tags)**: produção C3 passou para o ritmo do Concentrador III
  (`M4239010000-596-HHCC`, janela do espessador) e concentrado C3 para `M4238010000-1445-DDDVR`; produção C2 e C3 pela
  **média** (a planilha antiga usa a média: 18/08 = 26 983 / 18 006 t).
- **30/10/2026 (diagnóstico do dia 15/10)**: com `IP_MESVALOR` vieram produção C3 (15 877, igual ao histórico) e batch
  EB4/EB6 (EB6 2 210, igual ao histórico). Com `IP_AnalogMap` o MES respondia "Sucesso" sem nenhum dado — a troca
  automática do dia 26/10 tinha levado as tags para esse mapa. Padrão voltou para `IP_MESVALOR` e o Diagnóstico passou a
  descobrir a forma da tag e o mapa sozinho.
- **30/10/2026 (diagnóstico do dia 29/10)**: 12 de 12 tags com valor; o Diagnóstico gravou em Configurações a forma
  e o mapa de cada uma. Tanques com resultado às 04:00 e 16:00 (TK 01 sem resultado em 29/10). DVR aparece no
  Diagnóstico como "D+1 00:00" = centro da janela 23:30–00:30, ou seja, o valor gravado às 23:59 do dia D.
- **Fator**: só é guardado (não entra em nenhuma conta nem gráfico). Confirmar o que representa.

## Gerar de novo

```
cd build
python3 build_monitor.py            # gera ../Monitoramento_SE_Espessador_Batch.xlsm
python3 build_monitor.py --teste    # versão com o teste automático (LibreOffice, 54 verificações)
python3 extrair_historico.py "Monitoramento SE - Após fator ESP 6.2.xlsx"   # refaz build/historico.json
python3 mesclar_historico.py "planilha anterior.xlsx" "Monitoramento_SE_Espessador_Batch.xlsm"   # junta as duas
```
