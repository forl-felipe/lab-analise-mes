# Gerador das abas BD_Afericoes / BD_Limites

Só para manutenção. Os operadores não usam.

- `spec.py`: o mapa de cada ensaio, ou seja, qual aba e qual célula viram
  qual linha da tabela, e os limites.
- `build.py`: insere as duas abas ocultas direto no XML do .xlsm. Não abre
  nem salva pelo openpyxl, então VBA, gráficos, desenhos e modos de exibição
  ficam intactos.
- `avalia.py`: calcula todas as fórmulas com a biblioteca `formulas` e gera
  os valores que vão gravados no arquivo. O Power BI lê o valor gravado, não
  a fórmula.

Refazer a partir da planilha original:

    python3 avalia.py ORIGINAL.xlsm cache.json
    python3 build.py  ORIGINAL.xlsm SAIDA.xlsm cache.json

Se o layout de uma aba mudar (linhas ou colunas inseridas), ajuste o bloco
correspondente em `spec.py` e rode de novo.
