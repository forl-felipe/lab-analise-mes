# Registro de hipóteses e premissas

Projeto: tela de monitoramento do ensaio de compressão de pelotas (prensa automática, PLC 307-PLC-001), Laboratório de Controle de Embarque, Unidade UBU.

Situação: **Confirmada** (informada pelo LCE ou verificada em arquivo), **Pendente** (valor padrão adotado até confirmação) ou **Aguardando aprovação** (proposta a ser validada).

| # | Premissa | Valor adotado | Situação | Fonte ou como confirmar | Impacto se estiver errada |
|---|---|---|---|---|---|
| H01 | Plataforma do PIMS | AspenTech: Aspen InfoPlus.21 e Aspen Process Graphics Editor V14.2 (aspenONE). O planejamento inicial previa PI System | Confirmada | Arquivo `PIMS laboratório atualizado.atgraphic` e captura da tela do editor (22/10/2026) | Substitui PI Vision, PI AF e PI Web API nos entregáveis 02, 03 e 05 |
| H02 | Unidade da carga média e do desvio padrão | kgf/pel | Confirmada | Informada pelo LCE e coerente com a tela atual do Laboratório Químico | Escalas e linhas de referência |
| H03 | Meta de especificação e alarmes | Sem meta. Linhas em 100, 150 e 200 kgf/pel apenas como referência visual | Confirmada | Informada pelo LCE | Nenhum alarme de processo será configurado nesta versão |
| H04 | Nome dos registros no IP.21 e fonte de dados | Igual à tag (ex.: 307-GI-001), na fonte UBU | Confirmada | Informada pelo LCE em 22/10/2026: ao digitar a tag no editor, o IP.21 preenche automaticamente as informações do registro (IP_AnalogMap para analógicos) | Ligação dos campos da tela |
| H05 | Significado dos discretos 307-QI-001 a 003 | Estado 0/1, cumulativos, referentes à carga média do ensaio | Pendente | Automação (lógica do 307-PLC-001), ou observação após um ensaio com carga baixa | Lógica do painel de faixas e severidade |
| H06 | Periodicidade de atualização | A cada ensaio, por evento | Pendente | Tendência de 24 horas da 307-GI-001 no Process Graphics | Limites de defasagem (H07) |
| H07 | Limites de defasagem | "Sem ensaio recente" após 2 h; "Possível perda de comunicação" com qualidade diferente de Good, estado de erro do Cim-IO ou mais de 8 h sem atualização | Pendente | Validar com o LCE conforme a rotina de ensaios | Indicações de comunicação na tela |
| H08 | Valor 0 na carga média | Tratado como "sem ensaio registrado", sem acionar faixa | Pendente | Na última leitura (30/09/2026 11:28) todas as tags estavam em 0 com qualidade Good. Confirmar com a Automação se 0 é o valor de repouso do PLC | Sem esse tratamento, a tela indicaria "menor que 100" sem haver ensaio |
| H09 | Tipo do parâmetro PV_D nas tags analógicas | Valor analógico com casas decimais | Pendente | No DeltaV, PV_D costuma designar parâmetro discreto. Conferir com a Automação se a carga média chega com decimais ao IP.21 | Perda de resolução ou leitura incorreta da carga |
| H10 | Acesso programático aos dados (entregável 05) | Script entregue como opcional e não testado, usando Aspen Process Data REST ou SQLplus | Pendente | TI ou administração do PIMS | O entregável 05 pode ser dispensado |
| H11 | Forma da tela | Tela nova criada a partir de uma cópia de `PIMS laboratório atualizado`, apagando todos os elementos exceto Comp. Fx Maior. O original é preservado com "Salvar como" | Confirmada | Decisão do LCE em 22/10/2026 | Layout do protótipo e roteiro |
| H12 | Permissão para publicar telas | Edição no Process Graphics Editor disponível para o LCE; publicação a confirmar | Pendente | Administração do PIMS | Etapa final do roteiro |
| H13 | Identidade visual | Logotipo e paleta do kit visual fornecido. O kit informa que o logotipo foi reconstruído a partir de slide e que as cores foram medidas em imagens | Confirmada para uso interno | Confirmar com a Comunicação se existe arquivo oficial do manual de marca antes de uso amplo | Fidelidade da marca |
| H14 | Padrão visual de IHM | ISA-101: fundo cinza neutro, cor apenas para situação anormal, estado sempre também em texto. A tela atual usa fundo cinza, o que facilita a integração | Confirmada | Requisito do projeto | Nenhum |
| H15 | Severidade das faixas | Menor que 200: Atenção (prioridade 3, triângulo amarelo). Menor que 150: Alerta (prioridade 2, losango laranja). Menor que 100: Crítico (prioridade 1, quadrado vermelho). Sempre com texto "Ativo" ou "Inativo". São indicações visuais, sem alarme no IP.21 | Aguardando aprovação | Validação do LCE | Cores e textos do painel de faixas |
| H16 | Escala da carga média (307-GI-001) | 0 a 400 kgf/pel | Confirmada | Informada pelo LCE em 22/10/2026 | Escala das tendências e do minigráfico |
| H17 | Escala do desvio padrão (307-GI-002) | 0 a 100 kgf/pel | Pendente | Conferir após os primeiros ensaios registrados | Escala do minigráfico do desvio padrão |
| H18 | Faixas com carga sem qualidade Good | Se a 307-GI-001 ou qualquer discreto estiver sem qualidade Good, as três faixas mostram "Sem dado confiável" e nenhuma é indicada | Aguardando aprovação | Protótipo, cenário "Qualidade Bad na carga" | Comportamento mais conservador: evita indicar faixa com base em dado inválido |
| H19 | Campo de comparação com o laboratório | Comp. Fx Maior mantido da tela atual. Usina a definir: Usina 03 (`M710050020-0031-HHLFU`) ou Usina 04 (`M4710050020-0031-HHLFU`) | Pendente | Decisão do LCE | Tag do campo de referência e da pena verde da tendência |
| H20 | Cores das séries da tendência | Carga média em azul #1F5F99; Comp. Fx Maior em verde #2E8B57, com marcador quadrado. Amarelo, laranja e vermelho reservados às faixas | Confirmada | Paleta verificada com validador de contraste e de daltonismo | Nenhum |

## Observações sobre a tela atual do Laboratório Químico

Registradas durante a análise do arquivo, sem relação direta com a tela nova:

- O texto "SP Calcário" aparece mais de 3.000 vezes no arquivo, o que sugere objetos duplicados ou ocultos. Recomenda-se verificar, pois isso pode deixar a tela mais lenta.
- A legenda "-325# Saída da Prensa" aparece repetida junto aos campos de compressão e pode ter sido copiada de outro objeto.
- Os valores de compressão (Comp. Fx Maior e Comp. Fx Menor) são exibidos em vermelho mesmo quando não há situação anormal identificada. Pela ISA-101, o vermelho deveria ficar reservado para condição anormal.
