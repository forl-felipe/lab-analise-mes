# Monitoramento S.E. — Espessador × Batch

Planilha `Monitoramento_SE_Espessador_Batch.xlsm`: versão automática da `Monitoramento SE - Após fator ESP 6.2.xlsx`
(aba *Média ponderada*). O botão **Buscar no MES** traz os resultados do dia, **Salvar no histórico** grava o dia e o
**Painel** monta o report gerencial dos últimos N dias, pronto para copiar como imagem. O Painel segue o mesmo padrão
visual da aba Painel do Comparativo Germano × Ubu (`comparativo_interlab/`): cabeçalho Samarco, quadro de status
OK/NOK, comentários automáticos e gráficos com barras por dia + barra escura da média, linha laranja da diferença,
◆ diferença média e faixa verde tracejada da tolerância.

## Uso no dia a dia

1. Aba **Dia** → data do report (**ontem**: o DVR só entra no MES no fechamento do dia) → **Buscar no MES**.
   Confira os valores (células amarelas), o Nº de resultados e a coluna Situação.
2. Lance o Fator, se mudar (dia novo repete o do dia anterior). TK 01 e TK 02 já vêm do MES.
3. **Salvar no histórico** → o Painel passa a terminar nesse dia.
4. Aba **Painel** → **Copiar imagem** (e **Copiar resumo** para o corpo do e-mail).

Dias atrasados: na aba Dia, **Buscar um período** (De/Até, no máximo até ontem) busca e salva cada dia, mantendo
Fator, Observação e valores lançados à mão em itens que o MES não trouxe.

Item sem resultado: aba **Diagnóstico** → dia já fechado → **Rodar diagnóstico**. Cada tag é consultada sozinha de
D−1 12:00 a D+1 12:00 e a aba mostra se o MES aceitou a tag e o mapa, quantos horários têm valor, o 1º e o último
horário e o valor que a janela do item daria.

## Abas

| Aba | Para quê |
|---|---|
| **Painel** | Report gerencial (padrão do Comparativo): quadro *Último dia / Média do período / Tolerância / Status* (EB6, ESP C2+C3, EB4+ESP C3, as duas diferenças, % lama), **comentários automáticos** (último dia, período, dias fora da tolerância, % lama máxima, observação do dia) e caixa *Produção e DVR*. Gráficos: *Espessador C2 + C3 × Batch EB6*, *Batch EB4 + Espessador C3 × Batch EB6* e *% de lama no batch × S.E. do EB6 e dos tanques*. Barra de cima (fora da imagem): data final, nº de dias (2 a 60), Último dia, Ir para o Dia, Atualizar painel, Copiar imagem, Copiar resumo. |
| **Dia** | ① data do report, ② resultados do MES (valor, nº de resultados, janela, tag, situação), ③ resultado do dia, ④ busca de período e a última busca. |
| **Histórico** | Um dia por linha. A base de 28/01 a 15/10/2026 (254 dias) veio da planilha antiga. Colunas cinza são calculadas. |
| **Diagnóstico** | Teste tag por tag no MES (ver acima). |
| **Configurações** (oculta, senha `Samarco@Lab`) | Tag, servidor, mapa, janela, agregação e faixa válida de cada item; janelas de horário; **tolerância da diferença (±150)**; fator padrão; fonte (MES ou SIMULAÇÃO). |
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
- Status OK/NOK do Painel: |diferença do último dia| ≤ tolerância (Configurações, padrão ±150 — o mesmo da S.E. no
  Comparativo; confirmar com o laboratório).
- Busca de **hoje**: a planilha avisa que o dia não terminou (DVR e fechamento do batch ainda não estão no MES) e
  pergunta antes de salvar.

## Tags (aba Configurações)

Conferidas pelo laboratório no MES em 26/10/2026 (exceto a S.E. do espessador C3). Exemplo de horários: o report de
26/10 usa os dados de 25/10 — produção e espessadores de 24/10 23:00 a 25/10 21:00; batch e DVR de 25/10 00:00 a 23:59.

| Item | Tag | Mapa | Janela · agregação |
|---|---|---|---|
| Produção C3 | `M4239010000-596-HHCC` (Ritmo Produção Concentrador III) | IP_AnalogMap | E · Média |
| Produção C2 | `M0239010000-0596-HHCC` (Ritmo Produção Concentrador II) | IP_AnalogMap | E · Média |
| S.E. espessador C2 | `M0238010000-0017-HHLQ` | IP_AnalogMap | E · Média |
| S.E. espessador C3 | `M4238010000-2888-HHLQ` (**não está na conferência — confirmar**) | IP_AnalogMap | E · Média |
| S.E. batch EB4 / EB6 | `M250010000-17-ABLQ` / `M4250010000-17-ABLQ` | IP_MESVALOR | D · Média |
| TK 01 / TK 02 | `M0239040000-0017-TBLQ` / `M0239050000-0017-TBLQ` (Tk1 / Tk2 U2 - SE) | IP_MESVALOR | D · Média |
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
- **Fator**: só é guardado (não entra em nenhuma conta nem gráfico). Confirmar o que representa.

## Gerar de novo

```
cd build
python3 build_monitor.py            # gera ../Monitoramento_SE_Espessador_Batch.xlsm
python3 build_monitor.py --teste    # versão com o teste automático (LibreOffice, 54 verificações)
python3 extrair_historico.py "Monitoramento SE - Após fator ESP 6.2.xlsx"   # refaz build/historico.json
```
