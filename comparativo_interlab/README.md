# Comparativo de Medições de Qualidade — Germano × Ubu

Padrão visual Samarco (kit visual para painéis): Azul Samarco `#00335A`, azul médio `#004070`, azul de título
`#003B6C`, amarelo `#FFC000` (linhas de destaque e botão Buscar no MES), azul claro `#59C6F2`, laranja `#F37021`
(diferenças e horários em evidência), cinza da marca `#B5BEC4`, fonte Segoe UI e a logo do kit.

Duas planilhas, as duas entregues **limpas** (modelo padrão, sem batch nenhum):

| Planilha | Quem usa | Abas visíveis |
|---|---|---|
| `Acompanhamento_Batch_Germano_Ubu.xlsm` | técnicos de turno — batch **em processamento** | **Preenchimento** e **Batch** |
| `Comparativo_Interlab_Germano_Ubu.xlsm` | report diário — batch **processado** | **Batch**, **Painel** e **Histórico** |

As duas têm o mesmo código VBA (constante `VARIANTE` em `ModLayout`: `ACOMP` ou `COMP`); as abas que uma variante
não usa ficam muito ocultas, sem ser apagadas, para o código compilar igual nas duas.

## Planilha 1 — Acompanhamento do batch em processamento

### Aba Preenchimento

Só o necessário, em etapas:

1. **Identificação do batch**: batch Nº; **nível dos tanques** (lista Alto / Médio / Baixo — define o atraso entre a
   chegada em Ubu e o início no processo) **ou** **início no processo (manual)** — preencher um apaga o outro e o
   que não vale fica cinza; **estender o batch (horas)** — ver abaixo.
2. **Fases do evento**: área amarela do tamanho da tabela do MES (4 colunas × 3 linhas; cabeçalho ou coluna a mais
   também são aceitos) para colar a tabela de fases (Germano EB06 e Ubu UBU3 entram sozinhos
   na conferência).
3. **Conferência**: Germano início/término, Ubu chegada/término, início no processo, bi-horários, **estendido**
   (+N h, quando o batch foi estendido), data e mineroduto — dá para corrigir. Essa linha é a que se copia para o
   comparativo (abaixo).
4. **Conferência dos batches**: uma linha por batch (o atual e os anteriores, o mais recente em cima), sem fórmulas,
   no mesmo desenho da linha da conferência: BATCH, Germano início/término, Ubu chegada/término, início no processo,
   bi-horários, estendido, nível e tempo em Ubu. É daqui que se copia para o comparativo (da coluna BATCH a TEMPO EM
   UBU, Ctrl+C) — cada batch tem a sua linha, mesmo depois do Próximo batch.

Só dois botões: **Buscar no MES** (Germano e Ubu do batch; vão para o bloco dele na aba Batch) e **Próximo batch**
(limpa a aba Preenchimento; o que foi lançado na aba Batch fica). Não há PFUF, ajustes nem comentários do report
nesta planilha.

**Estender o batch (X horas)**: acrescenta `X / 2` bi-horários depois do último resultado de Germano (arredonda para
o bi-horário mais próximo). Nesses bi-horários Germano = **média das leituras do MES** do batch, repetida, com o
mesmo visual dos demais horários; a observação do bloco diz quantos foram estendidos. Ubu também vem do MES nesses
horários. Em branco =
só os resultados do MES.

### Aba Batch (5 batches por planilha)

Cinco blocos, um por batch, no padrão do Painel. Cada **Buscar no MES** atualiza o bloco do batch (Germano, Ubu e
horários); batch novo entra no primeiro lugar livre (lugar livre = só uma faixa cinza "LIVRE").
Com os 5 blocos ocupados a busca de um 6º batch não muda nada e avisa: comece uma cópia nova da planilha padrão
limpa.

Cada bloco: faixa do batch; **Resumo do batch**
(média de Germano e de Alimentação 04, diferença, tolerância, status OK/NOK); **Observações** — um
campo só, que já vem escrito com chegada em Ubu, nível dos tanques, início no processo, Germano (EB06) e, se
houver, "Batch estendido +6 h" (atualiza sozinho enquanto ninguém edita; o que for escrito à mão fica); e **Leituras por bi-horário**
com os horários de Ubu nas colunas e, para cada análise (SiO₂, P, PPC, S.E., <325#), as linhas Germano /
✎ Alimentação 04 / Diferença. Ubu vem do MES nos bi-horários já iniciados; onde o MES não tiver resultado, os
técnicos lançam à mão nas células amarelas (só elas são digitáveis). Fora da tolerância só aparece no resumo
(pela média: diferença em vermelho e status NOK); a grade fica sem marcação. Sem botões na aba: o bloco é
pensado para um print direto (inclusive no Excel online) para o e-mail — faixa só com "BATCH 300", resumo compacto
(uma coluna por campo) e espaço livre nas Observações e à direita para os horários estendidos. Colunas à direita do bloco
mais largo e linhas abaixo do último bloco ficam ocultas.

Regras do Germano no acompanhamento:
- Janelas de 1 h começando 30 min depois do horário da grade; bi-horário sem resultado não deixa coluna vazia —
  o resultado seguinte entra na sequência.
- Ubu: a cada busca, o que estiver na aba Preenchimento é limpo e o MES é consultado só para este batch (nunca vem
  Ubu de outro batch). Valor do MES vale; nos horários sem resultado no MES fica o que foi lançado à mão no bloco.
  A observação escrita à mão também fica.
- Janela de Germano já passada sem resultado no MES aparece na mensagem, para conferir no MES.
- Sem a chegada em Ubu ainda, o bloco sai sem os horários; ao informar a chegada e buscar de novo, eles entram.

### No SharePoint

Salve o .xlsm numa biblioteca do SharePoint. Quem busca no MES abre pelo aplicativo (Abrir no aplicativo da área de
trabalho — macros e Aspen só rodam no aplicativo), roda o Buscar no MES e fecha. Os técnicos abrem o mesmo arquivo
pelo navegador e lançam juntos na aba Batch: no Excel online as fórmulas, cores e a proteção funcionam; os botões
não (nada quebra, só não rodam). Cuidados: não rodar o Buscar no MES enquanto alguém lança no navegador; a política
de TI precisa permitir abrir .xlsm no navegador.

**Abas protegidas com senha** (`fislab`, a mesma de Configurações): Preenchimento e Batch. Os blocos ficam
em linhas e colunas fixas; apagar ou inserir uma linha/coluna desloca tudo e a macro deixa de achar o batch (o bloco
"trava"). Largura de coluna e altura de linha dá para mudar; mudança de desenho vai no gerador.

## Planilha 2 — Comparativo (report diário)

| Aba | Para quê |
|---|---|
| **Batch** | Preenchimento completo em etapas: **1 · Identificação** (batch Nº, nível dos tanques e atraso), **2 · Conferência** (colada da planilha dos técnicos), **Referência · últimos batches em Ubu**, **3 · Pellet feed ultra fino**, **4 · Comentários do report** e **5 · Leituras por horário** (Germano e Ubu); ajustes opcionais no fim. Ações: Buscar no MES (batch + PFUF), Salvar no histórico, Ver painel, Limpar leituras; na faixa do título: Configurações (com senha). **Salvar no histórico** já deixa a aba em branco para o próximo batch. |
| **Painel** | O report: cabeçalho, **farol** (médias, diferença, tolerância, OK/NOK), **comentários**, **pellet feed ultra fino** (composto diário do dia anterior: Fe, SiO₂, P, S.E.) e os gráficos de **P, SiO₂, PPC, <325# e S.E.** (barras Germano × Ubu por horário + média, linha da diferença com rótulos, ◆ diferença das médias e linhas tracejadas da tolerância). Na barra de cima: *Batch em exibição*, Editar batch, Atualizar, Copiar imagem e Copiar resumo. |
| **Conferência dos técnicos** | Com a planilha dos técnicos aberta no mesmo Excel (pode ser a do SharePoint, aberta no aplicativo), basta digitar o número do batch: a conferência vem do registro dela (Germano, Ubu, nível ou início manual, bi-horários e a extensão do batch). O Buscar no MES também tenta isso quando a conferência está vazia. |
| **Colar conferência** | Na planilha dos técnicos, aba Preenchimento, tabela **CONFERÊNCIA DOS BATCHES**: selecione a linha do batch (da coluna BATCH a TEMPO EM UBU) e Ctrl+C; aqui, na aba Batch, etapa 2 · CONFERÊNCIA, clique na célula BATCH e Ctrl+V — as duas linhas têm o mesmo desenho. O número do batch vem junto. A tabela não tem fórmulas, então o Excel não pergunta sobre nomes repetidos (copiar a linha 3 · Conferência dos técnicos, que tem fórmulas, faz o Excel perguntar "o nome já existe"). A macro lê o que foi copiado e refaz as fórmulas daqui (início no processo, defasagem, data). Se o Ctrl+V não for aceito, o botão **Colar conferência** faz o mesmo. Entram Germano início/término, Ubu chegada/término, bi-horários, mineroduto e o início no processo: se ele for a chegada + o atraso de um nível dos tanques (Configurações), fica o nível; senão, vira início manual. Se o batch foi estendido, o Buscar no MES completa os bi-horários a mais com a média de Germano, como na planilha dos técnicos. O comparativo não tem mais a etapa das fases do MES. |
| **Histórico** | Tabela `tbl_Batches`: um batch por linha, com médias, diferenças e status de cada análise. Botões: abrir e excluir o batch selecionado. Vem vazia. |
| **BD_Leituras** (oculta) | Tabela `tbl_Leituras`: uma linha por batch × horário × análise (Germano, Ubu, diferença, tolerância, status). Base para o Power BI. |

## Abas ocultas (nas duas)

`Configurações` — na planilha de acompanhamento é uma aba oculta comum (botão direito numa aba > Reexibir; as demais
abas técnicas ficam muito ocultas, então só ela aparece); no comparativo é muito oculta e abre pelo botão
Configurações com a senha da constante `SENHA_CFG` do gerador. Nas duas, as células de valor são editáveis e as
demais protegidas com essa senha. Conteúdo: tags do MES (batch e PFUF), consulta, atraso por nível dos tanques,
padrões de batch novo e títulos dos gráficos. `Dados_MES` (consultas ao Aspen), `_Graf` (séries dos gráficos, farol
e resumo) e `_Acomp` (registro dos batches: chegada/término em Ubu, nível, Germano — base da referência dos últimos
batches).

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
- **Germano em sequência** (Configurações, padrão Sim): como na tela do MES e nos relatórios do laboratório, cada
  resultado de Germano vira o próximo bi-horário — horário sem resultado no MES **não deixa coluna vazia**. A consulta
  de Germano usa janelas de **1 h** começando 30 min depois da grade (00:30–01:30 pega o resultado de 01:00), então
  cada resultado fica separado, inclusive o de fechamento do batch (ex.: 23:00 e 23:59 em colunas próprias; antes a
  janela de 2 h tirava a média dos dois). O horário real de cada resultado vai para "Origem em Germano". Se houver
  mais resultados que bi-horários, o Nº de horários do batch passa a ser o Nº de resultados.
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

## Pellet feed ultra fino (elutriação) — composto diário

Lançado à mão (a consulta automática ao MES saiu): na aba Batch, etapa 3, a data do composto (padrão: dia anterior
ao report) e os 4 resultados copiados do MES na vertical (Fe, SiO₂, P, S.E.) — clique em Fe e Ctrl+V. Em cada linha
vale o 1º número (texto com vírgula vira número; colunas a mais são limpas). Ao Salvar no histórico, data e
resultados ficam gravados com o batch e aparecem no Painel e no resumo do e-mail.

## Tempo da consulta

- Contador de segundos na barra de status do Excel durante a consulta; o tempo total aparece na mensagem final e
  em Dados_MES (*Última consulta*).
- Resposta de erro do Aspen (`Erro: ...`) encerra a espera na hora, sem aguardar o tempo máximo.
- **Mapa recusado** (`Map Name ... is invalid`): a planilha tenta `IP_AnalogMap`, `IP_MESVALOR` e `IP_DiscreteMap`, e o
  que o MES aceitar fica gravado na aba Configurações.

## Histórico

As planilhas saem com o histórico vazio. (A variante de teste `COMPLETA` importa `build/historico.json` para os
testes.)

Para levar o histórico de uma planilha do comparativo já em uso para a versão nova:

```
cd build
python3 extrair_historico.py "Comparativo_Interlab_Germano_Ubu_2.xlsm" hist.json
python3 build_interlab.py --historico hist.json     # -> ../Comparativo_Interlab_Germano_Ubu_COM_HISTORICO.xlsm
```

## Verificações do gerador

Além dos testes automáticos (LibreOffice), o gerador recusa o VBA com: declaração de módulo depois de rotina;
variável/parâmetro com o nome de uma rotina ou constante (o Excel não compila esse módulo); e variável/parâmetro com
nome de palavra reservada do VBA (`Next`, `For`, `Name`, `Date`…). Erro de compilação no VBA faz **todas** as
macros pararem, por isso essas verificações barram a geração.

## Gerar de novo

```
cd build
python3 build_interlab.py            # gera as duas planilhas limpas em ../
python3 build_interlab.py --teste    # versões _TESTE (COMPLETA, ACOMP e COMP) com o teste automático (LibreOffice)
```

`vba/*.bas` é o código das macros (o mesmo nas duas planilhas).
