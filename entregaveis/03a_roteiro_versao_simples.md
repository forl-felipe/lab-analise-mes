# Versão simples da tela: reaproveitar a tela atual

A ideia é não construir nada do zero: aproveitar os objetos que já existem na tela do Laboratório Químico e apenas trocar textos e tags. Tempo estimado: 30 a 40 minutos.

## Passo 1. Fazer a cópia (2 min)

**Arquivo > Salvar como** e salve como `PIMS compressao prensa 307`. Confira o nome novo na barra de título antes de continuar.

## Passo 2. Apagar o que não será usado (10 min)

Arraste um retângulo sobre cada bloco e pressione **Delete**. Use Ctrl+Z se apagar algo por engano.

Apague: níveis dos tanques, Carregamento, Campanha, **Parâmetros USINA 04** inteiro, as tendências **LM 02** das duas usinas, a tendência **LM 01 da Usina 04**, Set Point de Insumos, Tamboramento e Comp. Fx Menor.

Mantenha: o bloco **Parâmetros USINA 03**, a tendência **LM 01 da Usina 03** e o **Comp. Fx Maior**.

## Passo 3. Transformar a tabela de parâmetros (15 min)

No bloco Parâmetros USINA 03, use as cinco primeiras linhas e apague as outras quatro (Mn, B2, Kg/T e %C).

Em cada linha, clique duas vezes no rótulo amarelo para trocar o texto, e clique duas vezes no campo branco para trocar a tag. Mantenha o mapa e a fonte UBU que o editor preencher.

| Linha atual | Novo rótulo | Nova tag |
|---|---|---|
| Fe | Carga média (kgf/pel) | 307-GI-001 |
| SiO2 | Desvio padrão (kgf/pel) | 307-GI-002 |
| CaO | Menor que 200 | 307-QI-003 |
| MgO | Menor que 150 | 307-QI-002 |
| P | Menor que 100 | 307-QI-001 |

As três últimas tags são discretas: ao digitar a tag, o editor deve trocar o mapa para IP_DiscreteMap. Se continuar IP_MESVALOR ou IP_AnalogMap, troque para IP_DiscreteMap.

## Passo 4. Ajustar a tendência (5 min)

Clique duas vezes na tendência LM 01. Substitua as penas atuais por:

1. `307-GI-001` (carga média)
2. A tag do Comp. Fx Maior mantido (`M710050020-0031-HHLFU`)

Defina a escala de 0 a 400. Apague o texto "LM 01 OPERANDO" acima da tendência.

## Passo 5. Títulos e salvar (3 min)

1. Clique duas vezes em "Laboratório Químico" e troque por "Ensaio de compressão de pelotas".
2. Troque "Parâmetros USINA 03" por "Prensa automática | 307-PLC-001".
3. **Arquivo > Salvar**.

## Pronto

O resultado é uma tela com os cinco valores da prensa, a tendência da carga média e o Comp. Fx Maior para comparação. Melhorias como cores por faixa, banda de desvio padrão e aviso de dado defasado podem ser feitas depois, uma por vez, se houver interesse.
