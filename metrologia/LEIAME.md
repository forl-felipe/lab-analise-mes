# Sistema Metrologia · calibração e aferição do laboratório físico

Um QR Code em cada equipamento, um app com **três acessos** e um painel de gestão no Power BI.
**Um QR, um equipamento, duas leituras:** o QR de cada etiqueta é o endereço da imagem do equipamento no SharePoint. A câmera do celular abre a imagem (consulta de qualquer pessoa); o scanner do app abre a ficha daquele equipamento para quem registra. O app não tem lista de equipamentos.
Substitui a planilha de aferição e integra o controle de calibração da versão 1.

| Perfil | Quem | Faz |
|---|---|---|
| **Colaborador** | Qualquer pessoa da empresa | Lê o QR com a **câmera do celular** e vê a **imagem** do equipamento: CALIBRADO (válida até…) ou NÃO UTILIZAR (calibração vencida). Não precisa do app |
| **Laboratorista** | Cadastrado na lista Acessos | No app: **Ler QR Code** (ou digita a TAG no computador) e lança a aferição daquele equipamento, com cálculo automático e justificativa de não conformidade. Vê a agenda do turno |
| **Instrumentação** | Cadastrado na lista Acessos | No app: **Ler QR Code** (ou digita a TAG) e vê a última calibração e as 3 últimas observações dos técnicos; registra a calibração. Recebe as pendências no e-mail diário |

**Quem tem qual acesso:** lista **Acessos** do SharePoint (seção 1.4 do guia do SharePoint).
**Onde ficam as tolerâncias:** lista **Limites** (seção 1.5 do guia do SharePoint).

## Conteúdo

| Pasta | O que tem |
|---|---|
| [`01-arquitetura`](01-arquitetura/ARQUITETURA.md) | Arquitetura, regras de negócio, perfis, decisões pendentes e [dicionário de dados](01-arquitetura/DICIONARIO-DE-DADOS.md) |
| [`02-sharepoint`](02-sharepoint/GUIA-SHAREPOINT.md) | Guia das 9 listas, script automático `criar-estrutura.ps1`, formatação de cores |
| [`03-powerapps`](03-powerapps/GUIA-POWERAPPS.md) | 5 telas prontas para colar (`telas/*.yaml`) e o guia |
| [`04-power-automate`](04-power-automate/GUIA-FLUXOS.md) | Fluxos F0 (acessos), F1 (alerta de não conformidade), F2 (rotina, relatório diário e aviso de calibração vencida), F3 (calibração registrada) e os modelos de e-mail |
| [`05-powerbi`](05-powerbi/GUIA-POWERBI.md) | Consultas M (mantêm as colunas do painel atual), tabela de previstos, medidas DAX, tema |
| [`06-etiquetas`](06-etiquetas/LEIAME.md) | Gerador de imagens de consulta (o que a câmera abre) e gerador de etiquetas QR. Os dois funcionam offline |
| [`07-testes`](07-testes/LEIAME.md) | Validação automática das regras contra a planilha |
| [`ferramentas`](ferramentas/) | Scripts que geram a carga inicial, o script PnP, a documentação e as telas |

**Pacote de dados (fora do GitHub):** a carga inicial (`carga-inicial.xlsx`, CSVs), o roteiro de testes com valores esperados, o relatório de validação e as listas para as etiquetas são gerados a partir da planilha do laboratório e entregues à parte. Eles **não** ficam neste repositório, que é público.

## Ordem de implantação

| Etapa | O que | Tempo |
|---|---|---|
| 0 | Confirmar com a TI o **ambiente** do app e as **licenças** do Power BI | — |
| 1 | **SharePoint:** criar as listas (script ou manual), permissões, carga inicial, cadastrar Acessos | 2 h (manual) |
| 2 | **Power Apps:** criar o app, conectar as listas, colar as 5 telas, inserir o leitor de QR, testar com o roteiro | 2 h |
| 3 | **Power Automate:** fluxos F0, F1, F2 e F3 | 2 h |
| 4 | **Power BI:** trocar a fonte do painel atual para as listas, conferir os números de outubro | 1 h |
| 5 | **Imagens e etiquetas:** gerar as imagens de consulta, carregar na ConsultaQR, gerar e colar as etiquetas | 1h30 |
| 6 | **Piloto:** 2 semanas com app e planilha em paralelo, comparando resultados | 2 semanas |
| 7 | **Virada:** desligar a planilha, apagar a lista Calibradores da versão 1 | — |

## Checklist de implantação

- [ ] Ambiente de produção e licenças confirmados com a TI
- [ ] 9 listas e a biblioteca ConsultaQR criadas, com índices e permissões
- [ ] Carga inicial importada (Configuracoes, Limites, Ensaios, PontosAfericao, Equipamentos, Lancamentos, Afericoes)
- [ ] TAGs repetidas e datas de calibração conferidas
- [ ] Acessos cadastrados (laboratoristas e instrumentação)
- [ ] App publicado, compartilhado e com coproprietário
- [ ] Leitor de QR inserido na tela Início e testado no celular
- [ ] Roteiro de testes executado (26 cálculos + testes de perfil, fluxo, peneiradores e QR)
- [ ] Fluxos F0, F1, F2 e F3 ligados e testados
- [ ] Power BI conferido com o painel da planilha (400 resultados, 91 previstos, 83 realizados)
- [ ] Imagens de consulta carregadas na ConsultaQR
- [ ] Etiquetas impressas e coladas (testadas com a câmera e com o scanner do app)
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
