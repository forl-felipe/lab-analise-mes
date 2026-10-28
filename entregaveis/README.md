# Tela de monitoramento do ensaio de compressão de pelotas (LCE, UBU)

| Pasta ou arquivo | Conteúdo |
|---|---|
| `hipoteses.md` | Premissas adotadas e o que ainda precisa de confirmação |
| `01_dicionario_tags.csv` | Dicionário das cinco tags (separador ponto e vírgula, abre direto no Excel) |
| `02_prototipo/index.html` | Protótipo navegável com dados simulados. Abra no navegador e troque o cenário no canto inferior direito |
| `03_aspen/` | **Pacote de execução no Aspen:** `roteiro_aspen_graphics.pdf` (passo a passo com uma ficha por objeto; também em `.html` e `.md`), imagem de fundo e mapa em PNG e BMP, `posicoes_objetos.xlsx` e telas de exemplo |
| `04_comparacao_modulo_mes/` | Análise da tela PQ US04, método de comparação módulo x MES e planilha `comparacao_modulo_mes.xlsx` (dados simulados) |
| `ferramentas/ler_atgraphic.py` | Lê um arquivo `.atgraphic` e lista objetos, tags e posições (`pip install olefile`) |
| `ferramentas/gerar_planilha_comparacao.py` | Regera a planilha de comparação |
| `ferramentas/gerar_imagens.js` e `montar_pacote_aspen.py` | Regeram o pacote do Aspen a partir do protótipo: `node gerar_imagens.js` e depois `python montar_pacote_aspen.py` |

Comece por `03_aspen/roteiro_aspen_graphics.pdf`.
