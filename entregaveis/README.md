# Tela PIMS: monitoramento do ensaio de compressão de pelotas

Laboratório de Controle de Embarque (LCE), Unidade UBU. Prensa automática de compressão, PLC 307-PLC-001, fonte DeltaV01. Plataforma: Aspen InfoPlus.21 e Aspen Process Graphics V14.2.

| Fase | Arquivo | Situação |
|---|---|---|
| 1 | `01_dicionario_tags.csv` | Concluída (pendências da TI registradas) |
| 1 | `02_modelo_IP21.md` | Concluída (pendências da TI registradas) |
| 1 | `hipoteses.md` | Atualizado continuamente |
| 2 | `04_prototipo/index.html` | Em validação |
| 2 e 3 | `03_roteiro_process_graphics.md` | Parte 1 (preparação) pronta; Partes 2 e 3 na Fase 3 |
| 3 | `05_validar_tags.py` | Opcional (hipótese H10) |
| 4 | `06_checklist_testes.md` e `07_manual_operacao.md` | A fazer |

## Caminho até a tela pronta

| Etapa | Quem executa | O que acontece |
|---|---|---|
| 1. Enviar perguntas no RITM0273793 | LCE | Confirmar nome dos registros, lógica dos discretos e parâmetro PV_D (H04, H05, H09) |
| 2. Validar o protótipo | LCE e equipe | Abrir `04_prototipo/index.html` no navegador e aprovar layout e severidade |
| 3. Conferir as tags no PIMS | LCE | No Process Graphics Editor, localizar as cinco tags e anotar valor, horário e qualidade |
| 4. Montar a tela | LCE | Seguir `03_roteiro_process_graphics.md` no Process Graphics Editor, salvando como arquivo novo |
| 5. Criar o botão de acesso | LCE | Inserir na tela do Laboratório Químico um link para a tela nova |
| 6. Testar | LCE | Executar `06_checklist_testes.md` |
| 7. Publicar | Administração do PIMS | Disponibilizar a tela no servidor para todos os usuários |
| 8. Treinar | LCE | Apresentar a tela à equipe com `07_manual_operacao.md` |

O CSV usa ponto e vírgula como separador e codificação UTF-8, para abrir corretamente no Excel com configuração regional do Brasil.
