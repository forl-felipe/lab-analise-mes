# Hipóteses e premissas da tela de compressão de pelotas

Situação de cada item: **Confirmado** (informado pelo LCE), **Adotado** (valor padrão em uso até confirmação) ou **Pendente** (precisa de informação para avançar).

| Nº | Assunto | Hipótese em uso | Situação |
|---|---|---|---|
| H1 | Plataforma do PIMS | O PIMS da UBU usado nesta tela é AspenTech (aspenONE, Aspen Process Graphics Editor V14.2, histórico InfoPlus.21). A especificação original citava PI Vision e PI AF; os entregáveis foram adaptados ao Aspen. | Confirmado pela captura de tela |
| H2 | Nome das tags no PIMS | O nome usado no PIMS é o da coluna "Tag" do chamado RITM0273793 (307-GI-001, 307-GI-002, 307-QI-001, 307-QI-002, 307-QI-003). | Confirmado pelo LCE |
| H3 | Campos do InfoPlus.21 | Valor em IP_INPUT_VALUE, data e hora em IP_INPUT_TIME e qualidade em IP_INPUT_QUALITY, que são os campos padrão dos registros IP_AnalogDef e IP_DiscreteDef. | Adotado, conferir no Tag Browser |
| H4 | Unidade de engenharia | daN/pelota para carga média e desvio padrão. | Adotado |
| H5 | Escala de exibição | Carga média de 0 a 500 daN/pelota; desvio padrão de 0 a 150 daN/pelota. | Adotado |
| H6 | Significado dos discretos | Estado binário (0 = Inativo, 1 = Ativo) da carga média do ensaio abaixo de 100, 150 e 200 daN/pelota. Faixas cumulativas: com carga abaixo de 100, as três ficam ativas. | Adotado, confirmar com a automação |
| H7 | Severidade das faixas | Menor que 200: Atenção (amarelo, símbolo ▲). Menor que 150: Alerta (laranja, símbolo ◆). Menor que 100: Crítico (vermelho, símbolo ■). | Proposta, aguarda aprovação do LCE |
| H8 | Especificação e alarmes | Sem meta oficial informada. Nenhum alarme sonoro ou notificação, apenas sinalização visual na tela. | Pendente |
| H9 | Periodicidade | Atualização por ensaio (por evento), não contínua. | Adotado |
| H10 | Defasagem | Comunicação considerada perdida quando o último ensaio tem mais de 4 h. | Adotado, ajustar ao intervalo real entre ensaios |
| H11 | Resolução da tela | Monitor do laboratório em 1920 x 1080. Para outra resolução, as imagens podem ser regeradas. | Adotado |
| H12 | Logotipo | Usar somente o arquivo oficial fornecido pelo LCE ou pela comunicação. O protótipo reserva o espaço e não reproduz a marca. | Pendente (arquivo não recebido) |
| H13 | Valores atuais iguais a zero | A leitura de 30/09/2026 com todas as tags em 0 e qualidade Good indica integração criada e ainda sem ensaios reais. Valor 0 na carga média não deve ser tratado como pelota fraca até haver ensaio real. | Adotado |
| H14 | Tags de cálculo (opcionais) | Quantidade de tags Good, tempo desde o último ensaio, banda de mais ou menos um desvio padrão e linhas de referência dependem de tags calculadas no Aspen Calc, a criar pelo administrador do PIMS. Os nomes sugeridos no roteiro são provisórios. | Pendente |
| H15 | Foco da tela | O módulo de compressão é o destaque. A comparação com o MES e o estado da amostragem são complementares, num painel discreto e numa planilha à parte. | Confirmado pelo LCE |
| H16 | Tag do laboratório | Compressão do MES em `M4710050020-0031-HHLFU` (mapa IP_MESVALOR), associada ao rótulo "Compressão" na tela PQ US04 pela posição. | Adotado, conferir no editor |
| H17 | Janela do MES | Um resultado a cada 2 h; o resultado no horário T representa a amostra de T menos 2 h até T. | Confirmado (2 h); referência do horário pendente |
| H18 | Modo da amostragem | Modo dos cortadores em `U04-07CR001` e `U04-07CR002` pelo mapa IP_PIDMAP03 (Automático ou Manual). Comparação válida só com os dois em automático e sem falha durante toda a janela. | Adotado |
| H19 | Usina operando | Alimentação `406-FX-001` acima de 300 t/h. | Adotado, valor a confirmar |
| H20 | Pelotas por ensaio | 60 no laboratório e 20 no módulo, usadas só para estimar a diferença esperada pelo acaso. | Pendente |
| H21 | Vínculo de campos no Aspen | Campos de dados usam tag, mapa e atributo VAL, como na tela PQ US04. | Adotado |

## O que não pôde ser validado

O ambiente de desenvolvimento não tem acesso à rede da Samarco. Não foram testados: leitura das tags, nomes de campos do InfoPlus.21, comportamento dos objetos do Aspen Process Graphics Editor e aparência final no monitor do laboratório. O protótipo usa dados simulados.
