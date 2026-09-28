# Informativo de Qualidade & Passagem de Turno: Usinas 3 e 4

Planilha Excel com macros (`Informativo_Qualidade_US3_US4.xlsm`) que substitui a antiga
"TRABALHO INFORMATIVO QUALIDADE (planilha padrão)". Ela foi refeita para a operação atual:

- **somente as Usinas 3 e 4**;
- **turnos de 12 horas**: Dia 07h–19h e Noite 19h–07h;
- **turmas A, B, C e D**;
- **resultados puxados do MES** (Aspen IP.21), com a mesma função do suplemento Aspen usada na planilha antiga;
- **identidade visual Samarco** (paleta institucional);
- **tudo numa única planilha**, com histórico acumulado de todos os turnos.

## Abas

| Aba | Para que serve |
|---|---|
| **Painel** | Menu principal: seleção do turno (data, Dia/Noite, turma, responsável), 4 botões de ação, situação do turno e indicadores US3 × US4 |
| **Informativo** | Resultados do MES do turno: 22 parâmetros × 2 usinas, 6 janelas de 2 h, resultado do turno, mín./máx., LIE/LSE e status OK/Fora |
| **Passagem de Turno** | Formulário: equipe, segurança (SSMA), testes e pendências, minerodutos/batch, status de 14 equipamentos (US3/US4), comentários por usina, embarques e observações |
| **Histórico** | Uma linha por turno fechado, com os resultados de US3 e US4, a contagem de equipamentos NÃO OK, SSMA, pendências e o PDF |
| **Registro Passagem** | Cada informação da passagem vira uma linha (formato pronto para filtros e Power BI) |
| **Tendências** | Gráfico dos últimos 30 turnos (US3 × US4 × LIE/LSE) do parâmetro escolhido, com estatísticas |
| **Configurações** | Fonte dos dados, servidor, horário, pasta dos PDFs, e-mail, **tags do MES**, **limites LIE/LSE**, faixa válida e botão **Inserir logo** |
| Dados_MES (oculta) | Fórmulas do suplemento Aspen (`GetCalculationValues` / `ShowCalculationValues`) |

## Os 4 botões do Painel

1. **Atualizar dados do MES**: consulta o MES para a janela do turno selecionado (6 intervalos de 2 h).
   - Descarta valores fora da faixa válida e preenche o Informativo.
   - Cada parâmetro recebe um resultado do turno: média; Produção = soma; Ritmo = último valor. A regra pode ser mudada em Configurações.
2. **Passagem de Turno**: abre o formulário.
   - Há um botão para **carregar as pendências do turno anterior** direto do Histórico.
3. **Informativo do Turno**: mostra o relatório, com botões para **copiar como imagem** (colar no e-mail, Teams ou WhatsApp) e **gerar PDF**.
4. **Fechar Turno**:
   - valida os dados;
   - grava no Histórico e no Registro;
   - gera o PDF (Informativo + Passagem);
   - opcionalmente abre um e-mail no Outlook com resumo e anexo;
   - salva o arquivo;
   - oferece **preparar o próximo turno**: avança Dia → Noite → Dia, define a turma que recebe e leva as pendências.

## Primeiro uso

1. **Desbloqueie o arquivo** baixado: botão direito → Propriedades → marcar **Desbloquear**. Depois abra e clique em **Habilitar conteúdo**.
2. Em **Configurações**, clique em **Inserir logo nas abas** e escolha o arquivo do logo Samarco (PNG).
3. Ainda em **Configurações**:
   - **confira as tags** das Usinas 3 e 4 (foram herdadas da planilha de 2014);
   - preencha **LIE/LSE** dos parâmetros que devem ser avaliados.
4. Para treinar ou testar sem o MES, mude **Fonte dos dados** para `SIMULAÇÃO`. Para uso real, deixe `MES`.
5. O **suplemento Aspen Process Explorer (Excel Add-in)** precisa estar ativo, como na planilha antiga. Sem ele, as fórmulas mostram `#NOME?`, e o botão 1 avisa e explica o que verificar.

## Pontos a confirmar com a área

- **Tags:** as tags de US3/US4 vieram da planilha padrão antiga. Algumas podem ter mudado no MES.
- **Horários da consulta:** usa intervalos de 2 h a partir do início do turno (07h/19h), como a planilha antiga fazia a partir de 01h30. Confirmar se os resultados do MES ficam corretos nessas janelas. O horário de início do turno Dia pode ser ajustado em Configurações.
- **Produção:** somada nos 6 intervalos. A planilha antiga somava com peso dobrado no último intervalo do dia; confirmar a regra.
- **Lista de equipamentos e equipe:** os nomes são editáveis direto na aba Passagem de Turno.

## Como foi validado

- O projeto VBA é gerado por código (`build/`).
- Foi validado com **olevba** e com o **LibreOffice**. O LibreOffice executou o fluxo completo em modo simulação: atualizar → preencher passagem → fechar turno → Histórico/Registro gravados → próximo turno preparado → pendências recarregadas.
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
