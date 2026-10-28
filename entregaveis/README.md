# Tela de monitoramento do ensaio de compressão de pelotas (LCE, UBU)

| Pasta ou arquivo | Conteúdo |
|---|---|
| `hipoteses.md` | Premissas adotadas e o que ainda precisa de confirmação |
| `01_dicionario_tags.csv` | Dicionário das cinco tags (separador ponto e vírgula, abre direto no Excel) |
| `02_prototipo/index.html` | Protótipo navegável com dados simulados. Abra no navegador e troque o cenário no canto inferior direito |
| `03_aspen/` | **Tela pronta para o Aspen:** `LCE_Compressao_Pelotas_Prensa01.atgraphic` (objetos nativos, sem imagem de fundo), `teste_minimo.atgraphic` para diagnóstico, passo a passo `roteiro_aspen_graphics.pdf` (também `.html` e `.md`), mapa e telas de exemplo |
| `04_comparacao_modulo_mes/` | Análise da tela PQ US04, método de comparação módulo x MES e planilha `comparacao_modulo_mes.xlsx` (dados simulados) |
| `ferramentas/ler_atgraphic.py` | Lê um arquivo `.atgraphic` e lista objetos, tags e posições (`pip install olefile`) |
| `ferramentas/gerar_planilha_comparacao.py` | Regera a planilha de comparação |
| `ferramentas/gerar_imagens.js`, `gerar_atgraphic.py` e `montar_pacote_aspen.py` | Regeram a tela e o pacote a partir do protótipo: `node gerar_imagens.js`, `python gerar_atgraphic.py` e `python montar_pacote_aspen.py` |
| `ferramentas/conferir_atgraphic.py` | Relê um `.atgraphic` gerado e redesenha a tela em SVG para conferência |

Comece por `03_aspen/roteiro_aspen_graphics.pdf`.
