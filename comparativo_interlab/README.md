# Comparativo de Medições de Qualidade — Germano × Ubu

Planilha `Comparativo_Interlab_Germano_Ubu.xlsm`: versão em Excel do painel HTML
"Comparativo Interlaboratorial Germano × Ubu" (Report Gerencial), com os resultados puxados do MES.

## Abas

| Aba | Para quê |
|---|---|
| **Painel** | O report: cabeçalho, **farol** (médias, diferença, tolerância, OK/NOK), **comentários**, **pellet feed ultra fino** (composto diário do dia anterior: Fe, SiO₂, P, S.E.) e os gráficos de **P, SiO₂, PPC, <325# e S.E.** (barras Germano × Ubu por horário + média, linha da diferença com rótulos, ◆ diferença das médias e linhas tracejadas da tolerância). Na barra de cima: *Batch em exibição*, Editar batch, Atualizar, Copiar imagem e Copiar resumo (texto curto para o corpo do e-mail). |
| **Batch** | ① **batch Nº** e ③ **nível dos tanques** à esquerda; ② área grande para **colar as Fases do Evento** do MES (sem células mescladas em volta); linha de conferência (Germano EB06, Ubu UBU3, início no processo, bi-horários, defasagem, data, mineroduto); ④ **pellet feed ultra fino** (data do composto e os 4 resultados, buscados no MES junto com as leituras e editáveis); **comentários do report** (parâmetros fora e acompanhamento automáticos e editáveis + 6 tópicos); a grade de **leituras por horário**; no fim, ajustes opcionais. As diferenças por horário ficam só no Painel. Linhas e colunas sem uso ficam ocultas (também no Painel). Botões: Buscar no MES, Salvar no histórico, Ver painel, Limpar leituras. **Salvar no histórico** já deixa a aba em branco para o próximo batch (tolerâncias, títulos e comentários voltam aos padrões de Configurações). Abre com um batch em branco. |
| **Acomp. Químico** / **Acomp. Físico** | Batch **em processamento**: um bloco por batch (o mais recente em cima), como as planilhas de acompanhamento do laboratório — Data, Início aproximado no processo (horário de Ubu), e para cada análise: Germano (MES), Alimentação 04 (Ubu, lançado a mão a cada 2 h nas células amarelas) e Diferença (Ubu − Germano), com a **Média**; média fora da tolerância em vermelho. Químico: SiO₂, P, PPC. Físico: S.E., <325#. Pode ficar no SharePoint para o físico e o químico lançarem juntos (as macros não rodam no Excel online, as fórmulas sim). |
| **Histórico** | Tabela `tbl_Batches`: um batch por linha, com médias, diferenças e status de cada análise. Botões: abrir e excluir o batch selecionado. |
| **BD_Leituras** (oculta) | Tabela `tbl_Leituras`: uma linha por batch × horário × análise (Germano, Ubu, diferença, tolerância, status). Base para o Power BI. |
| **Configurações** | Tags do MES (batch e PFUF), consulta, atraso por nível dos tanques, padrões de batch novo e títulos dos gráficos. |

Ocultas: `BD_Leituras`, `Dados_MES` (consultas ao Aspen), `_Graf` (séries dos gráficos, farol e resumo) e `_Acomp`
(registro dos batches em acompanhamento: chegada/término em Ubu, nível, Germano).

## Regras (iguais às do painel HTML)

- Média de cada laboratório = média de **todas** as leituras dele no batch.
- Diferença = avaliado − referência (padrão: **Ubu − Germano**; a referência pode ser trocada no batch).
- **OK** quando |diferença arredondada| ≤ tolerância. Tolerâncias padrão: S.E. 150 · <325# 1,0 · SiO₂ 0,19 · P 0,003 · PPC 0,08.
- Os gráficos mostram só os horários com leitura nos dois laboratórios (ex.: química de Ubu de 4 em 4 h ×
  Germano de 2 em 2 h), mais a coluna **Média**, que é a média de todas as leituras de cada laboratório.
- **Escala de cada análise** (Configurações: frequência e 1º horário): o MES devolve a análise de 4 h em duas
  janelas de 2 h (valor repetido); só o horário da escala fica. Padrão: Ubu SiO₂, P e PPC de 4 em 4 h a partir de
  03:30 (03:30, 07:30, 11:30, 15:30, 19:30, 23:30); demais de 2 em 2 h. Ubu de 2 em 2 h: frequência 2.

## Consulta ao MES

- Uma fórmula `GetCalculationValues` por laboratório, igual aos relatórios do laboratório (tags, servidores e mapas
  em texto literal, intervalo `2h`). Tags recusadas pelo MES saem da consulta, que é refeita com as demais.
- **Nível dos tanques** em Ubu: atraso até o batch prevalecer no processo — **Alto 8 h, Médio 6 h, Baixo 4 h**
  (tabela em Configurações). Início considerado no processo = chegada em Ubu (UBU3) + atraso; o ajuste manual,
  quando preenchido, vale no lugar.
- Janelas de 2 h: **Ubu** na grade 07:00, 09:00, 11:00… (horário 07:30); **Germano** na grade 00:00, 02:00…, para
  os resultados de Germano (01:00, 03:00…) caírem no meio da janela (horário mostrado 01:00). A 1ª janela é a
  primeira completa depois do início. Grades em Configurações.
- **Nº de bi-horários = o de Germano** (EB06, do início ao término do bombeamento); Ubu tem os mesmos.
  O bi-horário 1 de Germano é comparado com o 1 de Ubu, e assim por diante. **Horário oficial = Ubu**: a hora de
  Germano aparece só como referência de origem.
- A fórmula do Aspen é escrita testando as formas do nome da função (com e sem `_xll.`) e a que o Excel reconhece
  fica guardada. Sem o suplemento (#NOME?), a planilha avisa uma vez e não fica esperando. Depois de cada consulta a
  fórmula vira texto (uma consulta ativa refaz a busca no MES a cada recálculo e deixa a planilha lenta).
- Tag recusada pelo MES: a planilha tenta o mesmo código no outro formato (`-0008-` ↔ `-8-`); se o MES aceitar,
  a tag é corrigida sozinha na aba Configurações. Resposta "Sucesso" sem resultados: a planilha monta a matriz
  `ShowCalculationValues` e espera mais uma vez.
- Análise **sem tag** fica fora da consulta (lançamento manual na grade). Para treinar sem MES: Fonte = SIMULAÇÃO.
- Tags (aba Configurações): Germano **Mineroduto 3** `M4250010000-0017-ABLQ` (S.E.), `-0016-ABLF` (<325#),
  `-0004-ABLQ` (SiO₂), `-0008-ABLQ` (P), `-0013-ABLQ` (PPC); Ubu `M4650060010-0017-HHLFU` (S.E.),
  `M4650060010-0016-HHLFU` (<325#), `M650030010-0004/-0008/-0013-HHLQU` (SiO₂, P, PPC).
  Tags do Mineroduto 2 (para a futura ponderação por volume): `M250010000-17-ABLQ`, `-16-ABLF`, `-4-ABLQ`,
  `-8-ABLQ`, `-13-ABLQ`. Tabulações e quebras de linha coladas junto com a tag são removidas.

## Batch em processamento (acompanhamento)

- Na aba Batch, **Tipo da busca no MES** (embaixo dos botões): *Comparativo (batch processado)* ou
  *Acompanhamento (em processamento)*.
- Acompanhamento: informar o batch, colar as Fases do Evento (Germano EB06) e, quando houver, a chegada em Ubu e o
  nível dos tanques. **Buscar no MES** traz só Germano e monta/atualiza o bloco do batch nas abas Acomp. Químico e
  Acomp. Físico. Buscar de novo atualiza Germano e os horários e **mantém o que já foi lançado de Ubu**.
- Sem a chegada em Ubu ainda, o bloco sai sem os horários; ao informar a chegada e buscar de novo, eles entram.
- Salvar no histórico fica para o comparativo (depois de processado). Ao digitar no Batch um número que já está em
  acompanhamento, chegada, término, nível e Germano já vêm preenchidos.
- **Últimos batches em Ubu** (aba Batch, abaixo da conferência): os 3 batches anteriores ao da aba (histórico e
  acompanhamento) com nível dos tanques, chegada e término em Ubu, início no processo e tempo em Ubu — referência
  para preencher o próximo.

## Pellet feed ultra fino (elutriação) — composto diário

- **Composto do dia anterior** à data do report (ex.: report de 22/10 → composto de 21/10). A data fica no Batch
  (`data do report − 1`, configurável em *Dias antes da data do report*) e pode ser trocada à mão.
- Um botão só: o **Buscar no MES** traz as leituras do batch e o PFUF (sem a chegada em Ubu ainda, traz só o PFUF).
- O composto do dia é gravado no MES às **23:30 do próprio dia** (Configurações: *Horário de lançamento no MES*).
  A consulta usa janelas de **2 h**, como as leituras do batch, de 12 h antes a 12 h depois do dia do composto; vale a
  janela que contém as 23:30. As demais janelas com resultado aparecem na mensagem, para conferir o horário.
  Análise sem valor às 23:30 usa o último valor do **mesmo dia** do composto (a mensagem avisa o horário); valor
  fora da faixa válida (Configurações) é descartado e a mensagem mostra qual foi. Faixa da S.E. do PFUF: 100 a 20000.
  Os valores que o MES não devolver ficam como estão.
- Tags (servidor UBU, aba Configurações): `M0650024000-0003-DDLQU` (Fe — no MES "PF Elutriação - FeO"),
  `M0650024000-0004-DDLQU` (SiO₂), `M0650024000-0008-DDLQU` (P), `M0650024000-0017-DDLFU` (S.E.).
  Mapa `IP_AnalogMap` (o MES recusou `IP_MESVALOR` para essas tags). O rótulo mostrado no painel também é editável lá.
  Mesmo tratamento de tag recusada/variante de código do batch.
- Os 4 valores e a data do composto vão para o histórico (colunas *PFUF ...*).

## Tempo da consulta

- Contador de segundos na barra de status do Excel durante a consulta; o tempo total aparece na mensagem final e
  em Dados_MES (*Última consulta*).
- Resposta de erro do Aspen (`Erro: ...`) encerra a espera na hora, sem aguardar o tempo máximo.
- **Mapa recusado** (`Map Name ... is invalid`): a planilha tenta `IP_AnalogMap`, `IP_MESVALOR` e `IP_DiscreteMap`, e o
  que o MES aceitar fica gravado na aba Configurações.

## Histórico

A versão gerada traz o histórico salvo pelo laboratório (`build/historico.json`: batch 284 e as leituras dele).

## Gerar de novo

```
cd build
python3 build_interlab.py            # gera ../Comparativo_Interlab_Germano_Ubu.xlsm
python3 build_interlab.py --teste    # versão com o módulo de teste automático (LibreOffice)
```

`historico_inicial.json` tem os batches importados; `vba/*.bas` é o código das macros.
