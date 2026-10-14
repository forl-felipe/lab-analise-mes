# Comparativo de Medições de Qualidade — Germano × Ubu

Planilha `Comparativo_Interlab_Germano_Ubu.xlsm`: versão em Excel do painel HTML
"Comparativo Interlaboratorial Germano × Ubu" (Report Gerencial), com os resultados puxados do MES.

## Abas

| Aba | Para quê |
|---|---|
| **Painel** | O report: cabeçalho, **farol** (médias, diferença, tolerância, OK/NOK), **comentários** e os gráficos de **P, SiO₂, PPC, <325# e S.E.** (barras Germano × Ubu por horário + média, linha da diferença com rótulos, ◆ diferença das médias e linhas tracejadas da tolerância). Na barra de cima: *Batch em exibição* (lista do histórico), Copiar imagem, Salvar PNG, PDF e E-mail. |
| **Batch** | Só 3 passos: ① **número do batch**, ② **colar as Fases do Evento** da tela do MES (preenche Germano EB06, Ubu UBU3, mineroduto e data), ③ **nível dos tanques**. Abaixo, a conferência (dá para corrigir), comentários, a grade de **leituras por horário** e, no fim, os ajustes opcionais (laboratórios, referência, títulos, tolerâncias e ajuste manual do início no processo). Botões: Buscar no MES, Salvar no histórico, Novo batch, Limpar leituras e Importar da planilha antiga. A planilha abre com um batch em branco. |
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
- **Nível dos tanques** em Ubu: atraso até o batch prevalecer no processo — **Alto 8 h, Médio 6 h, Baixo 4 h**
  (tabela em Configurações). Início considerado no processo = chegada em Ubu (UBU3) + atraso; o ajuste manual,
  quando preenchido, vale no lugar.
- Janelas de 2 h na grade 07:00, 09:00, 11:00… (Configurações); a 1ª é a primeira completa depois do início.
- **Nº de bi-horários = o de Germano** (EB06, do início ao término do bombeamento); Ubu tem os mesmos.
  O bi-horário 1 de Germano é comparado com o 1 de Ubu, e assim por diante. **Horário oficial = Ubu**: a hora de
  Germano aparece só como referência de origem.
- A fórmula do Aspen é escrita testando as formas do nome da função (com e sem `_xll.`) e a que o Excel reconhece
  fica guardada. Sem o suplemento (#NOME?), a planilha avisa uma vez e não fica esperando. Depois de cada consulta a
  fórmula vira texto (uma consulta ativa refaz a busca no MES a cada recálculo e deixa a planilha lenta).
- Análise **sem tag** fica fora da consulta (lançamento manual na grade). Para treinar sem MES: Fonte = SIMULAÇÃO.
- Tags cadastradas pelo laboratório em 14/10/2026 (aba Configurações), com as correções: S.E. de Germano estava
  igual à de PPC (→ `M250010000-17-ABLF`), S.E. de Ubu na US3 (→ `M4650060010-0017-HHLFU`, Alimentação 04) e tags
  de Ubu com o código de 4 dígitos (`-0016-`, `-0008-`, `-0013-`).

## Histórico inicial

Os batches 286, 288 e 290 foram importados da planilha "Comparativo Batch GER × UBU 286 – 290" (287 e 289 estavam
vazios) e ficam no Histórico; a planilha abre com um batch em branco.

## Gerar de novo

```
cd build
python3 build_interlab.py            # gera ../Comparativo_Interlab_Germano_Ubu.xlsm
python3 build_interlab.py --teste    # versão com o módulo de teste automático (LibreOffice)
```

`historico_inicial.json` tem os batches importados; `vba/*.bas` é o código das macros.
