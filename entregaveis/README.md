# Tela Compressão Online Usina 4 (LCE, UBU)

| Pasta ou arquivo | Conteúdo |
|---|---|
| `hipoteses.md` | Premissas adotadas e o que ainda precisa de confirmação |
| `01_dicionario_tags.csv` | Dicionário das cinco tags (separador ponto e vírgula, abre direto no Excel) |
| `02_prototipo/index.html` | Protótipo navegável com dados simulados. Abra no navegador e troque o cenário no canto inferior direito |
| `03_aspen/` | **Tela pronta para o Aspen (versão 8):** `LCE_Compressao_Online_Usina4.atgraphic` (objetos nativos, sem composto de 2 h e sem tags calculadas), `teste_minimo.atgraphic` para diagnóstico, passo a passo `roteiro_aspen_graphics.pdf` (também `.html` e `.md`), mapa e telas de exemplo |
| `04_comparacao_modulo_mes/` | Análise da tela PQ US04, método de comparação módulo x MES e planilha `comparacao_modulo_mes.xlsx` (dados simulados) |
| `06_planilha_correlacao/` | **Planilha com macros `correlacao_prensa_laboratorio.xlsm`**: busca no PIMS pelo suplemento Aspen, como o Relatório Físico, a prensa automática (307-GI-001) e a Resistência à Compressão do laboratório (M4710050020-0031-HHLFU) e compara composto com composto (2 h). Sai sem dados. Também o código da macro `ModPIMS.bas`. A versão antiga com dados simulados ficou em `06_planilha_correlacao/versao_anterior/` |
| `ferramentas/gerar_planilha_aspen.py` | Regera a planilha com macros (`python gerar_planilha_aspen.py saida.xlsm`; usa `vba/ModPIMS.bas` e `vba_projeto.py`) |
| `ferramentas/gerar_planilha_correlacao.py` | Regera a versão anterior da planilha de correlação (dados simulados) |
| `ferramentas/ler_atgraphic.py` | Lê um arquivo `.atgraphic` e lista objetos, tags e posições (`pip install olefile`) |
| `ferramentas/gerar_planilha_comparacao.py` | Regera a planilha de comparação |
| `ferramentas/gerar_imagens.js`, `gerar_atgraphic.py` e `montar_pacote_aspen.py` | Regeram a tela e o pacote a partir do protótipo: `node gerar_imagens.js`, `python gerar_atgraphic.py` e `python montar_pacote_aspen.py` |
| `ferramentas/conferir_atgraphic.py` | Relê um `.atgraphic` gerado e redesenha a tela em SVG para conferência |

Comece por `03_aspen/roteiro_aspen_graphics.pdf`.
