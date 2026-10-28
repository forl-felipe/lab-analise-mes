# Etapa 4 · Power BI (a gestão)

O app só lança e consulta. **A gestão fica no Power BI**: conformidade, aderência à agenda por letra, tendência por equipamento e situação de calibração.

A ideia é **aproveitar o relatório que você já tem**: só a fonte de dados muda, de Excel para as listas do SharePoint. As consultas abaixo devolvem as tabelas com **o mesmo nome e as mesmas colunas** da planilha (`tbl_Afericoes` e `tbl_Limites`), então os gráficos e medidas atuais continuam funcionando.

## 4.1 Antes de começar

- Faça uma **cópia** do arquivo .pbix atual (ex.: `Metrologia - SharePoint.pbix`). O original continua como está até a virada.
- As listas já precisam ter a carga inicial (os 400 registros de outubro), para você comparar os números.
- Não precisa de gateway: o Power BI lê o SharePoint Online direto.

## 4.2 Parâmetro e funções

**Página Inicial > Transformar dados** (abre o Power Query). Em cada passo, **Nova fonte > Consulta nula**, abra o **Editor avançado**, apague tudo, cole o arquivo e renomeie a consulta com o nome indicado.

| Arquivo | Nome da consulta | O que é |
|---|---|---|
| `consultas-M/00-SiteUrl.pq` | `SiteUrl` | Parâmetro com o endereço do site. Troque pelo endereço real |
| `consultas-M/01-fnLista.pq` | `fnLista` | Lê uma lista do SharePoint |
| `consultas-M/02-fnData.pq` | `fnData` | Converte datas do SharePoint sem deslocar o dia |

## 4.3 Trocar a fonte das tabelas atuais

1. Clique na consulta **tbl_Afericoes** existente > **Editor avançado** > apague tudo > cole `10-tbl_Afericoes.pq` > Concluído.
2. Faça o mesmo com **tbl_Limites** e `11-tbl_Limites.pq`.

Se o seu relatório usa outros nomes de consulta, use os seus nomes. O importante é manter o nome que os gráficos usam.

> **Colunas novas** que passam a existir em tbl_Afericoes: `TAG` (liga com Equipamentos), `Causa`, `ItensNaoOK`, `Lançamento`, `Fonte` (App ou Planilha), `ResponsavelEmail` e `Modo de leitura` (QR, Digitado ou Link: mostra o que foi lançado sem a leitura da etiqueta).

## 4.4 Consultas novas

| Arquivo | Nome da consulta | Para que serve |
|---|---|---|
| `20-Equipamentos.pq` | `Equipamentos` | Situação de cada equipamento (Liberado/Bloqueado e motivo), calibração |
| `21-Lancamentos.pq` | `Lancamentos` | Um registro por envio |
| `22-Previstos.pq` | `Previstos` | Agenda expandida com a situação de cada ensaio previsto: base de **Aderência** e **Gestão por Letras** |

A consulta **Previstos** usa a **mesma regra da macro da planilha** (feito no turno, no dia, com atraso, não feito). Ela foi conferida contra o painel de outubro: **91 previstos e 83 realizados**, ensaio por ensaio.

**Fechar e aplicar.** Na primeira vez, o Power BI pede credenciais: escolha **Conta organizacional** > Entrar.

## 4.5 Modelo e medidas

1. **Modelagem > Nova tabela**: cole a tabela `Calendario` de `medidas-DAX.md`.
2. Crie os relacionamentos da tabela em `medidas-DAX.md`.
3. Crie as medidas. Se o relatório atual já tem medidas com os mesmos nomes, mantenha as suas.
4. **Exibir > Temas > Procurar temas**: escolha `tema-metrologia.json` (cores iguais às do app).

## 4.6 Conferir antes de publicar

Compare com o painel da planilha, de 01 a 22/10:

| Indicador | Planilha | Power BI |
|---|---|---|
| Resultados | 400 | ☐ |
| Não conformes | 12 | ☐ |
| Previstos (agenda) | 91 | ☐ |
| Realizados | 83 | ☐ |
| Blaine: previstos / realizados | 44 / 42 | ☐ |

## 4.7 Publicar e atualizar sozinho

1. **Publicar** no workspace do laboratório.
2. No Power BI online: conjunto de dados > **Configurações > Credenciais da fonte de dados** > OAuth2 com sua conta.
3. **Atualização agendada**: ligue e coloque horários (ex.: 07:15, 12:00, 19:15, 23:00). Com licença Pro são até 8 por dia.
4. **Assinatura diária** (relatório por e-mail com a imagem do painel): veja o final de `04-power-automate/GUIA-FLUXOS.md`.
5. **Compartilhar**: o relatório precisa de licença Pro para quem for ver (ou workspace Premium). Confirme com a TI.
