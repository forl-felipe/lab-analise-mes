# Comparativo de Medições de Qualidade — Germano × Ubu

Planilha `Comparativo_Interlab_Germano_Ubu.xlsm`: versão em Excel do painel HTML
"Comparativo Interlaboratorial Germano × Ubu" (Report Gerencial), com os resultados puxados do MES.

## Abas

| Aba | Para quê |
|---|---|
| **Painel** | O report: cabeçalho, **farol** (médias, diferença, tolerância, OK/NOK), **comentários**, **Pellet Feed Ultra Fino**, e os gráficos de **P, SiO₂, PPC, <325# e S.E.** (barras Germano × Ubu por horário + média, linha da diferença com rótulos, ◆ diferença das médias e linhas tracejadas da tolerância). Na barra de cima: *Batch em exibição* (lista do histórico), Copiar imagem, Salvar PNG, PDF e E-mail. |
| **Batch** | Cabeçalho do batch, horários para o MES, tolerâncias, comentários (com a frase automática dos parâmetros NOK), PFUF e a grade de **leituras por horário** (até 24 horários de 2 h). Botões: Buscar no MES, Salvar no histórico, Novo batch, Limpar leituras e Importar da planilha antiga. |
| **Enviar** | Resumo pronto para o corpo do e-mail e o assunto; Copiar resumo, Criar e-mail no Outlook (painel no corpo + PNG anexo), imagem, PNG e PDF. |
| **Histórico** | Tabela `tbl_Batches`: um batch por linha, com médias, diferenças e status de cada análise. Botões: abrir e excluir o batch selecionado. |
| **BD_Leituras** | Tabela `tbl_Leituras`: uma linha por batch × horário × análise (Germano, Ubu, diferença, tolerância, status). |
| **Configurações** | **Tags do MES** (uma por análise e laboratório, com servidor, mapa, faixa válida e valor típico da simulação), consulta (fonte MES/SIMULAÇÃO, tempo de espera, formato da data, minutos antes do horário), padrões de batch novo (laboratórios, referência, mineroduto, tolerâncias, títulos) e títulos dos gráficos. |

Ocultas: `Dados_MES` (consultas ao Aspen), `_Graf` (séries dos gráficos e cálculos do farol) e `Importar`.

## Regras (iguais às do painel HTML)

- Média de cada laboratório = média de **todas** as leituras dele no batch.
- Diferença = avaliado − referência (padrão: **Ubu − Germano**; a referência pode ser trocada no batch).
- **OK** quando |diferença arredondada| ≤ tolerância. Tolerâncias padrão: S.E. 150 · <325# 1,0 · SiO₂ 0,19 · P 0,003 · PPC 0,08.
- Os gráficos mostram só os horários com leitura nos dois laboratórios, mais a coluna **Média**.

## Consulta ao MES

- Uma fórmula `GetCalculationValues` por laboratório, igual aos relatórios do laboratório (tags, servidores e mapas
  em texto literal, intervalo `2h`). Tags recusadas pelo MES saem da consulta, que é refeita com as demais.
- Horários do batch copiados da **tela de eventos do MES** (Fases do Evento): **Germano** = início e término do
  bombeamento (ex.: EB06); **Ubu** = Terminal Ponta Ubu (UBU3).
- Janelas de 2 h na grade 07:00, 09:00, 11:00… (aba Configurações). A 1ª janela é a **primeira completa depois do
  início** (UBU3 13/10 09:18 → janela 11:00–13:00, mostrada como 11:30). O Nº de horários sai do término de Ubu.
- Germano: mesma regra a partir do início do bombeamento, até o término (EB06 11/10 00:00 → 01:30, 03:30…).
  O horário 1 de Germano é comparado com o horário 1 de Ubu, e assim por diante (a **defasagem** aparece no batch).
  Sem início de Germano = mesmos horários de Ubu.
- Análise **sem tag** fica fora da consulta (lançamento manual na grade). Para treinar sem MES: Fonte = SIMULAÇÃO.
- Tags cadastradas pelo laboratório em 14/10/2026 (aba Configurações).

## Histórico inicial

Os batches 286, 288 e 290 foram importados da planilha "Comparativo Batch GER × UBU 286 – 290" (287 e 289 estavam
vazios). A planilha abre no batch 288, que é o mais completo.

## Gerar de novo

```
cd build
python3 build_interlab.py            # gera ../Comparativo_Interlab_Germano_Ubu.xlsm
python3 build_interlab.py --teste    # versão com o módulo de teste automático (LibreOffice)
```

`historico_inicial.json` tem os batches importados; `vba/*.bas` é o código das macros.
