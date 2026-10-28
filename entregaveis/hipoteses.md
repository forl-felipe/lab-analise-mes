# Hipóteses e premissas da tela de compressão de pelotas

Situação de cada item: **Confirmado** (informado pelo LCE), **Adotado** (valor padrão em uso até confirmação) ou **Pendente** (precisa de informação para avançar).

| Nº | Assunto | Hipótese em uso | Situação |
|---|---|---|---|
| H1 | Plataforma do PIMS | O PIMS da UBU usado nesta tela é AspenTech (aspenONE, Aspen Process Graphics Editor V14.2, histórico InfoPlus.21). A especificação original citava PI Vision e PI AF; os entregáveis foram adaptados ao Aspen. | Confirmado pela captura de tela |
| H2 | Nome das tags no PIMS | O nome usado no PIMS é o da coluna "Tag" do chamado RITM0273793 (307-GI-001, 307-GI-002, 307-QI-001, 307-QI-002, 307-QI-003). | Confirmado pelo LCE |
| H3 | Campos do InfoPlus.21 | Valor em IP_INPUT_VALUE, data e hora em IP_INPUT_TIME e qualidade em IP_INPUT_QUALITY, que são os campos padrão dos registros IP_AnalogDef e IP_DiscreteDef. | Adotado, conferir no Tag Browser |
| H4 | Unidade de engenharia | kgf/pelota para a média e o desvio padrão da prensa e para o laboratório. | Confirmado pelo LCE |
| H5 | Escala de exibição | Carga média de 0 a 500 daN/pelota; desvio padrão de 0 a 150 daN/pelota. | Adotado |
| H6 | Significado dos "discretos" 307-QI-001/002/003 | Quantidade de pelotas do ensaio abaixo de 100, 150 e 200 daN (valores lidos no PIMS: 0, 4 e 5, cumulativos). | Adotado a partir do PIMS; confirmar com a automação |
| H7 | Severidade das faixas | Menor que 200: Atenção (amarelo, símbolo ▲). Menor que 150: Alerta (laranja, símbolo ◆). Menor que 100: Crítico (vermelho, símbolo ■). | Proposta, aguarda aprovação do LCE |
| H8 | Especificação e alarmes | Sem meta oficial informada. Nenhum alarme sonoro ou notificação, apenas sinalização visual na tela. | Pendente |
| H9 | Periodicidade | Atualização por ensaio (por evento), não contínua. | Adotado |
| H10 | Defasagem | Comunicação considerada perdida quando o último ensaio tem mais de 4 h. | Adotado, ajustar ao intervalo real entre ensaios |
| H11 | Tamanho e construção da tela | Tela de 50 x 35 unidades (padrão do editor, igual à PQ US04), montada só com objetos nativos (polígonos, textos, campos de dados e tendências), sem imagem de fundo, para não perder nitidez ao maximizar. Arquivo `.atgraphic` gerado por engenharia reversa do formato da tela PQ US04. | Adotado; abertura no Aspen pendente de teste pelo LCE |
| H12 | Logotipo | Logotipo oficial do kit visual Samarco fornecido pelo LCE, convertido em polígonos nativos do Aspen. | Confirmado |
| H13 | Valores atuais iguais a zero | A leitura de 30/09/2026 com todas as tags em 0 e qualidade Good indica integração criada e ainda sem ensaios reais. Valor 0 na carga média não deve ser tratado como pelota fraca até haver ensaio real. | Adotado |
| H14 | Tags de cálculo (opcionais) | Quantidade de tags Good, tempo desde o último ensaio, banda de mais ou menos um desvio padrão e linhas de referência dependem de tags calculadas no Aspen Calc, a criar pelo administrador do PIMS. Os nomes sugeridos no roteiro são provisórios. | Pendente |
| H15 | Foco e comparação | Módulo em destaque. Comparação com o laboratório feita composto com composto (média dos 4 cortes de 2 h da prensa contra o composto do MES), versão 6 aprovada pelo LCE. | Confirmado pelo LCE |
| H16 | Tag do laboratório | Compressão do MES em `M4710050020-0031-HHLFU` (mapa IP_MESVALOR), associada ao rótulo "Compressão" na tela PQ US04 pela posição. | Adotado, conferir no editor |
| H17 | Janela do MES | Um resultado a cada 2 h; o resultado no horário T representa a amostra de T menos 2 h até T. | Confirmado (2 h); referência do horário pendente |
| H18 | Modo da amostragem | Modo dos cortadores em `U04-07CR001` e `U04-07CR002` pelo mapa IP_PIDMAP03 (Automático ou Manual). Comparação válida só com os dois em automático e sem falha durante toda a janela. | Adotado |
| H19 | Usina operando | Alimentação `406-FX-001` acima de 300 t/h. | Adotado, valor a confirmar |
| H20 | Pelotas por ensaio | 60 no laboratório e 20 no módulo, usadas só para estimar a diferença esperada pelo acaso. | Pendente |
| H22 | Composto da prensa | Média dos cortes com resultado na janela de 2 h (1 a 4), ignorando zeros e qualidade ruim. Tags calculadas 307-GI-001-C2H, -N2H, -DIF2H e -CORTE1 a 4, a criar no Aspen Calc; a tela já vem ligada a esses nomes. | Pendente (administrador do PIMS) |
| H23 | Modo dos cortadores | Valor 32 = Automático e 16 = Manual no mapa IP_PIDMAP03, deduzido dos indicadores da tela PQ US04. | Adotado, conferir em campo |
| H21 | Vínculo de campos no Aspen | Campos de dados usam tag, mapa e atributo VAL, como na tela PQ US04. Os campos de hora e de qualidade vêm com VAL e precisam ter o atributo trocado no editor. Mapa de discretos: IP_DISCRETEMAP, a confirmar. | Adotado |

## O que não pôde ser validado

O ambiente de desenvolvimento não tem acesso à rede da Samarco. Não foram testados: leitura das tags, nomes de campos do InfoPlus.21, comportamento dos objetos do Aspen Process Graphics Editor e aparência final no monitor do laboratório. O protótipo usa dados simulados.
