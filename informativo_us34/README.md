# Informativo de Qualidade & Passagem de Turno: Usinas 3 e 4

Planilha Excel com macros (`Informativo_Qualidade_US3_US4.xlsm`) que substitui a antiga
"TRABALHO INFORMATIVO QUALIDADE (planilha padrão)". Ela foi refeita para a operação atual:

- **somente as Usinas 3 e 4**;
- **turnos de 12 horas**: Dia 07h–19h e Noite 19h–07h;
- **turmas A, B, C e D**;
- **resultados puxados do MES** (Aspen IP.21), com a mesma função do suplemento Aspen usada na planilha antiga;
- **identidade visual Samarco**: logo, paleta e padrão do kit visual (`build/assets/`);
- **modelo diário**: um arquivo por dia (copiado do modelo), com os turnos Dia e Noite. Não há histórico acumulado dentro da planilha: o histórico é a própria pasta do mês, com um arquivo por dia.

## Abas

| Aba | Para que serve |
|---|---|
| **Painel** | Menu principal: seleção do turno (data, Dia/Noite, turma, responsável), 4 botões de ação, situação do turno e indicadores US3 × US4 |
| **Resumo Dia** / **Resumo Noite** | Um resumo por turno, pronto para o e-mail, em **2 imagens**: (1) resultados de 2 em 2 h de US3 e US4, com resultado do turno, mín. e máx. (sem LIE/LSE/Status; fora da especificação em laranja) e (2) passagem de turno em texto. Botões **Copiar RESULTADOS** e **Copiar PASSAGEM**, ou print com Windows+Shift+S. Os valores são gravados como uma "foto" do turno: o turno seguinte pode atualizar o MES e preencher a passagem sem alterar o resumo já finalizado |
| **Passagem de Turno** | Equipe (técnico físico, técnico químico e 4 laboratoristas), turma que recebe, testes e pendências, minerodutos 02 e 03, status de 14 equipamentos (Operando / Não operando), comentários por usina (produto, aglomerante, combustível sólido, qualidade das pelotas), embarque em andamento e observações |
| **Informativo** (completo) | Resultados do MES de 2 em 2 h de **qualquer período de até 24 h**: o técnico preenche **Início** e **Fim** no alto da aba e clica em **Atualizar dados do MES** (acima de 12 h, a macro faz duas consultas de 12 h, iguais às do turno, e junta os resultados). Mostra só as janelas do período, o resultado, o mín. e o máx.; LIE/LSE/Status ficam ocultos. Botão **Copiar como imagem**. O botão 1 do Painel também carrega aqui o turno selecionado |
| **Configurações** (oculta, sem botão) | Fonte dos dados, servidor, horário, formato da data, **tags do MES**, **limites LIE/LSE** e faixa válida |
| Dados_MES (oculta) | Fórmulas do suplemento Aspen (`GetCalculationValues` / `ShowCalculationValues`) |

## Os 4 botões do Painel

Ao abrir um arquivo novo, o Painel já vem com a **data e o turno pelo relógio** (na primeira hora de um turno, sugere o turno que acabou de terminar). Um arquivo de outro dia, já usado, abre com a **sua própria data e turno**, para completar ou corrigir depois.

Os botões **Editar turno Dia** e **Editar turno Noite** trocam o turno sem mudar a data. Um turno que ainda não começou pode ser selecionado e preenchido: o MES não é consultado e os resultados ficam em branco. Ao **Finalizar turno**, os dados do MES são buscados de novo sempre que o Informativo foi carregado antes do fim do turno.

1. **Atualizar dados do MES**: busca os resultados do turno selecionado.
   - A macro espera todas as consultas responderem e os valores pararem de mudar, para não trazer dados incompletos.
   - Num turno em andamento, os horários futuros ficam em branco. Basta atualizar de novo mais tarde.
2. **Passagem de Turno**: preencher os campos amarelos.
3. **Resumo do turno**: gera (ou refaz) o resumo do turno do Painel na aba Resumo Dia ou Resumo Noite e abre a aba. O resumo também é refeito sozinho ao atualizar o MES e antes de copiar as imagens, enquanto o turno não for finalizado.
4. **Finalizar turno**:
   - busca os dados do MES sozinho, se ainda não foram buscados;
   - refaz e **fixa** o resumo do turno (Resumo Dia ou Resumo Noite);
   - salva o arquivo e oferece limpar a passagem para o próximo turno.
   - **Não muda o turno do Painel.**

As imagens são copiadas com o Excel em 200% de zoom (e o zoom volta ao normal em seguida). Assim a imagem tem o dobro de pixels e as letras continuam nítidas ao ampliar no e-mail.

Para enviar por e-mail: na aba Resumo do turno, **Copiar RESULTADOS** (imagem 1) e **Copiar PASSAGEM** (imagem 2), colando cada uma com Ctrl+V. A cópia usa o formato bitmap, que cola no Outlook (inclusive web), no Teams e no WhatsApp Web.

### Rotina diária

1. Copie o arquivo modelo (`Informativo_Qualidade_US3_US4.xlsm`) para a pasta do mês e renomeie com a data, por exemplo `Informativo_2026-09-28.xlsm`.
2. O turno Dia usa o arquivo e finaliza; o turno Noite continua no **mesmo arquivo** e finaliza o seu.
3. No fim do mês, a pasta tem um arquivo por dia, cada um com o Resumo Dia e o Resumo Noite.

Guia de uma página para os técnicos: [`guia/Guia_rapido_Informativo_US3_US4.png`](guia/Guia_rapido_Informativo_US3_US4.png).

## Primeiro uso

1. **Desbloqueie o arquivo** baixado: botão direito → Propriedades → marcar **Desbloquear**. Depois abra e clique em **Habilitar conteúdo**.
2. Em **Configurações**, **no arquivo modelo** (as cópias diárias herdam os ajustes). A aba fica oculta e sem botão: clique com o botão direito numa guia → **Reexibir** → Configurações.
   - **confira as tags** das Usinas 3 e 4 (foram herdadas da planilha de 2014);
   - preencha **LIE/LSE** dos parâmetros que devem ser avaliados.
3. Para treinar ou testar sem o MES, mude **Fonte dos dados** para `SIMULAÇÃO`. Para uso real, deixe `MES`.
4. O **suplemento Aspen Process Data** precisa estar ativo, como na planilha antiga. Sem ele, as fórmulas mostram `#NOME?`, e o botão 1 avisa e explica o que verificar.

O logo vem embutido. Para regenerá-lo a partir do kit visual: `python3 build/gerar_logos.py <samarco-kit-visual.html>` (requer `pip install cairosvg`).

## Consulta ao MES (formato da planilha de referência)

**Matrizes de resultado:** o suplemento Aspen cria e redimensiona sozinho a matriz de resultados de cada consulta (no arquivo em uso, a saída vira `ROWS(...)&"#"&COLUMNS(...)`). Se sobrar a matriz de uma consulta anterior com outro número de linhas, o Excel recusa com *"Não é possível alterar parte de uma matriz"*. Por isso, antes de cada consulta a macro apaga as matrizes de saída inteiras (`LimparSaidasMES`). Se nenhuma consulta responder depois disso, ela recria as matrizes no formato original (`RestaurarSaidasMES`) e tenta mais uma vez.

A aba oculta `Dados_MES` segue o mesmo formato da `TRABALHO_INFORMATIVO_QUALIDADE.xlsm` (out/2014), que funciona no Excel. São três consultas `GetCalculationValues`:

| Consulta | Célula | Tags | Mapa | Cálculo |
|---|---|---|---|---|
| Qualidade | A7 → saída A9 | Grelha + 20 parâmetros de laboratório (US3/US4) | IP_ANALOGMAP / IP_MESVALOR | "1" |
| Produção | A17 → saída A19 | Produção e Ritmo (MES) | IP_MESVALOR | "1" |
| Ritmo de processo | A27 → saída A29 | 306GERAL-FIT003-R (US3), 406-RITMO (US4) | IP_ANALOGMAP | "0" |

- Tags, servidores (`UBU`) e mapas ficam em **texto literal** dentro da fórmula, quebrados em pedaços de até 250 caracteres com `&`, como faz o próprio suplemento.
- A cada **Atualizar dados do MES**, o VBA (`ModMES.RegerarFormulasMES`) reescreve as fórmulas a partir da aba Configurações. Assim, uma tag alterada ali entra na consulta.
- O texto completo das fórmulas, lado a lado com a referência, está em [`FORMULAS_MES.md`](FORMULAS_MES.md).
- **Tag de compressão:** a referência usa `…-0030-HHLFU`; esta planilha usa `…-0031-HHLFU`. Confirme no MES e troque na aba Configurações, se necessário.

## Pontos a confirmar com a área

- **Tags:** as tags de US3/US4 vieram da planilha padrão antiga. Algumas podem ter mudado no MES.
- **Horários da consulta:** usa intervalos de 2 h a partir do início do turno (07h/19h), como a planilha antiga fazia a partir de 01h30. Confirmar se os resultados do MES ficam corretos nessas janelas. O horário de início do turno Dia pode ser ajustado em Configurações.
- **Produção:** somada nos 6 intervalos. A planilha antiga somava com peso dobrado no último intervalo do dia; confirmar a regra.
- **Lista de equipamentos e equipe:** os nomes são editáveis direto na aba Passagem de Turno.

## Como foi validado

- O projeto VBA é gerado por código (`build/`).
- Foi validado com **olevba** e com o **LibreOffice**. O LibreOffice executou o fluxo completo em modo simulação: atualizar MES → preencher passagem → finalizar Dia (Resumo Dia fixado) → turno Noite no mesmo arquivo → finalizar Noite, com o Resumo Dia intacto.
- Não foi possível testar no **Excel para Windows** nem com o **MES real** neste ambiente. Recomenda-se o primeiro uso em modo `SIMULAÇÃO` e depois um turno real acompanhado.

## Regerar a planilha

```bash
pip install xlsxwriter
cd build
python3 build_workbook.py          # gera ../Informativo_Qualidade_US3_US4.xlsm
```

- Código VBA: pasta `vba/`. O `ModLayout.bas` é gerado automaticamente pelo `build_workbook.py`.
- Layout, parâmetros, tags e cores: `build/build_workbook.py`.
- Gerador do `vbaProject.bin` (MS-OVBA/MS-CFB): `build/vba_project.py`.
