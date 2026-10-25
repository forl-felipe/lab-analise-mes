# Testes

## Validação automática das regras

`validar_regras.py` confere se as regras do sistema (as mesmas do app, do fluxo e do Power BI, escritas em `ferramentas/regras.py`) dão os mesmos resultados que a planilha de aferição.

```
python ferramentas/gerar_carga.py "planilha.xlsm" dados
python 07-testes/validar_regras.py "planilha.xlsm" dados
```

| Teste | O que confere |
|---|---|
| A | Cada resultado da BD_Afericoes recalculado com o critério gravado na própria linha |
| B | Cada resultado recalculado com os critérios atuais da lista Limites |
| C | Cálculos da aba Lançamento (rpm, estufa, granulometria, tamboramento) refeitos a partir das entradas |
| D | Aderência à agenda (previstos x realizados por ensaio) comparada com a aba Painel |

O relatório sai em `dados/relatorio-validacao.md`.

**Resultado com a planilha de outubro/2026:** A 184/184 · B 184/184 · C 16/16 · D 8/8.
(As 216 linhas de verificação dos peneiradores gravam só a contagem de itens; a regra delas foi conferida pela fórmula da planilha.)

## Roteiro de aceitação no app

O arquivo `roteiro-de-testes.md` (no pacote de dados, porque usa valores do laboratório) traz 26 casos de cálculo com o valor esperado e 11 testes de perfil e fluxo. Faça antes de liberar o app e repita depois de qualquer mudança nas listas de configuração.
