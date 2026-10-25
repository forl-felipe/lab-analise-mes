# Sistema Metrologia · calibração e aferição do laboratório físico

Um QR Code em cada equipamento, um app com **três acessos** e um painel de gestão no Power BI.
Substitui a planilha de aferição e integra o controle de calibração da versão 1.

| Perfil | Quem | Faz |
|---|---|---|
| **Colaborador** | Qualquer pessoa da empresa | Lê o QR Code e vê se o equipamento pode ser usado, a validade e o histórico |
| **Laboratorista** | Cadastrado na lista Acessos | Lança as aferições do turno (agenda, cálculo automático, justificativa de não conformidade) |
| **Instrumentação** | Cadastrado na lista Acessos | Registra calibrações e acompanha os equipamentos bloqueados ou vencendo |

**Quem tem qual acesso:** lista **Acessos** do SharePoint (seção 1.4 do guia do SharePoint).
**Onde ficam as tolerâncias:** lista **Limites** (seção 1.5 do guia do SharePoint).

## Conteúdo

| Pasta | O que tem |
|---|---|
| [`01-arquitetura`](01-arquitetura/ARQUITETURA.md) | Arquitetura, regras de negócio, perfis, decisões pendentes e [dicionário de dados](01-arquitetura/DICIONARIO-DE-DADOS.md) |
| [`02-sharepoint`](02-sharepoint/GUIA-SHAREPOINT.md) | Guia das 9 listas, script automático `criar-estrutura.ps1`, formatação de cores |
| [`03-powerapps`](03-powerapps/GUIA-POWERAPPS.md) | 5 telas prontas para colar (`telas/*.yaml`) e o guia |
| [`04-power-automate`](04-power-automate/GUIA-FLUXOS.md) | Fluxos F0 (acessos), F1 (alerta de não conformidade), F2 (rotina e relatório diário) e os modelos de e-mail |
| [`05-powerbi`](05-powerbi/GUIA-POWERBI.md) | Consultas M (mantêm as colunas do painel atual), tabela de previstos, medidas DAX, tema |
| [`06-etiquetas`](06-etiquetas/) | Gerador de etiquetas QR (funciona offline) |
| [`07-testes`](07-testes/LEIAME.md) | Validação automática das regras contra a planilha |
| [`ferramentas`](ferramentas/) | Scripts que geram a carga inicial, o script PnP, a documentação e as telas |

**Pacote de dados (fora do GitHub):** a carga inicial (`carga-inicial.xlsx`, CSVs), o roteiro de testes com valores esperados, o relatório de validação e a lista para as etiquetas são gerados a partir da planilha do laboratório e entregues à parte. Eles **não** ficam neste repositório, que é público.

## Ordem de implantação

| Etapa | O que | Tempo |
|---|---|---|
| 0 | Confirmar com a TI o **ambiente** do app e as **licenças** do Power BI | — |
| 1 | **SharePoint:** criar as listas (script ou manual), permissões, carga inicial, cadastrar Acessos | 2 h (manual) |
| 2 | **Power Apps:** criar o app, conectar as listas, colar as 5 telas, testar com o roteiro | 1h30 |
| 3 | **Power Automate:** fluxos F0, F1 e F2 | 1h30 |
| 4 | **Power BI:** trocar a fonte do painel atual para as listas, conferir os números de outubro | 1 h |
| 5 | **Etiquetas:** conferir TAGs, gerar e colar | 1 h |
| 6 | **Piloto:** 2 semanas com app e planilha em paralelo, comparando resultados | 2 semanas |
| 7 | **Virada:** desligar a planilha, apagar a lista Calibradores da versão 1 | — |

## Checklist de implantação

- [ ] Ambiente de produção e licenças confirmados com a TI
- [ ] 9 listas criadas, com índices e permissões
- [ ] Carga inicial importada (Configuracoes, Limites, Ensaios, PontosAfericao, Equipamentos, Lancamentos, Afericoes)
- [ ] TAGs repetidas e datas de calibração conferidas
- [ ] Acessos cadastrados (laboratoristas e instrumentação)
- [ ] App publicado, compartilhado e com coproprietário
- [ ] Roteiro de testes executado (26 cálculos + 11 testes de fluxo)
- [ ] Fluxos F0, F1 e F2 ligados e testados
- [ ] Power BI conferido com o painel da planilha (400 resultados, 91 previstos, 83 realizados)
- [ ] Etiquetas impressas e coladas
- [ ] Piloto concluído e planilha desligada

## Como regenerar os arquivos

Os arquivos de SharePoint, telas e documentação são gerados pelos scripts da pasta `ferramentas`, a partir de `ferramentas/modelo.py`. Para mudar uma coluna ou uma tela, altere o script e rode:

```
python ferramentas/gerar_pnp.py      # 02-sharepoint/criar-estrutura.ps1
python ferramentas/gerar_docs.py     # guia do SharePoint e dicionário de dados
python ferramentas/gerar_telas.py    # 03-powerapps/telas
python ferramentas/gerar_carga.py "planilha.xlsm" dados     # pacote de dados (não versionar)
python 07-testes/validar_regras.py "planilha.xlsm" dados
```
Requisitos: Python 3.10+, `openpyxl`, `PyYAML`.
