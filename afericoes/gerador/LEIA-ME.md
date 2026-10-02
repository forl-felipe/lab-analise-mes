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

## Planilhas mensais (desde outubro/2026)

A planilha de **setembro** é o layout de referência. O Alpine dela vai até a
linha 24, porque ganhou 3 linhas. `spec.py` já lê as linhas 7 a 24.

- `limpeza.py`: as células de preenchimento de cada aba. É o que se apaga
  para fazer a padrão.
- `mensal.py`: gera as planilhas a partir da de setembro. Edita o texto do
  XML só nas células que mudam; o resto do arquivo fica byte a byte igual.
  Refaz as abas ocultas e grava os valores calculados.

      python3 mensal.py padrao   SETEMBRO.xlsm "PADRAO - CALIBRAÇÃO.xlsm"
      python3 mensal.py outubro  SETEMBRO.xlsm "CALIBRAÇÃO 10 - OUTUBRO.xlsm" OUTUBRO_DOS_OPERADORES.xlsm
      python3 mensal.py setembro SETEMBRO.xlsm "CALIBRAÇÃO 09 - SETEMBRO.xlsm"

- `verifica.py ARQUIVO.xlsm`: confere o XML, as tabelas e quantos registros
  o Power BI vai ler, por ensaio e por mês.

**Mês novo:** copie a PADRÃO, renomeie para `CALIBRAÇÃO MM - MÊS.xlsm` e
salve na pasta `Calibração mensal AAAA`. A PADRÃO deve ficar fora das pastas
de ano.
