# Passo a passo: tela Tanques e Fluxos no Aspen (versão 1)

**Tela:** Tanques e Fluxos, padrão visual Samarco. Refaz a tela Tanques com as mesmas tags, fórmulas e escalas: estoque dos 9 tanques e totais, qualidade do MES e 12 fluxos com tendência.

**Resultado esperado:** igual a `tela_exemplo_normal.png` (dados simulados)

**Tempo estimado:** 15 minutos

![Resultado esperado](tela_exemplo_normal.png)

## O que vem pronto no arquivo

O arquivo `Tanques_e_Fluxos.atgraphic` foi gerado pelo mesmo processo das telas Compressão Online Usina 4 e Qualidade de Embarque. Tudo é objeto nativo do Aspen e não perde nitidez ao maximizar. A tela tem 70 x 35 unidades, como a original.

| Conteúdo | Quantidade |
|---|---|
| Barras de nível dos tanques (com a fórmula da tela original) | {{N_BARRAS}} |
| Campos de dados ligados às tags ou às fórmulas | {{N_CAMPOS}} |
| Tendências com penas e escala definidas | {{N_TEND}} |
| Textos fixos | {{N_TEXTOS}} |

## Parte A. Abrir

**A1.** Copie `Tanques_e_Fluxos.atgraphic` para a pasta de rede das telas. A tela antiga pode ficar até esta ser conferida.

**A2.** Abra no **Aspen Process Graphics Editor**, compare com `tela_exemplo_normal.png` e maximize.

**Se der erro ao abrir:** tente `teste_minimo.atgraphic` (uma barra, um campo e uma tendência) e me envie a captura da mensagem.

## Parte B. Conferência rápida

**B1. Estoque dos tanques.** Cada tanque usa a fórmula da tela original: estoque em tms = (nível - 1,5 m) x 400, zero abaixo de 1,5 m. Compare os valores com a tela antiga aberta ao lado: devem ser iguais.

| Grupo | Tanques e tags de nível | Escala das barras |
|---|---|---|
| Preparação 1 | TQ 01 `lt-803`, TQ 02 `lt-805`, TQ 03 `lt-807`, TQ 04 `li-3802` | 0 a 5000 por tanque; total 0 a 20000 |
| Preparação 2 | TQ 01 `301HG001-LIT001-R`, TQ 02 `301HG002-LIT002-R`, TQ 08 `401-LI-001`, TQ 09 `401-LI-002`, TQ 10 `401-LI-004` | 0 a 5000 por tanque; total 0 a 25000 |

**B2. Nomes dos grupos.** Os nomes Preparação 1 e Preparação 2 seguem o texto "PREP 1 + PREP 2" da tela original. Se a fileira estiver trocada, me avise que eu inverto.

**B3. Qualidade e fluxos.** As tags, os mapas e as escalas são os da tela original. Nada a fazer, a menos que algum campo mostre `***`.

## Parte C. Tendências

**C1.** Na barra de tempo de cada tendência, use o período que preferir (a tela original não define). Sugestão: **`24:00:00`**.

**C2.** Não altere a escala das penas: as escalas escritas na tela (por exemplo, "Vazão, 0 a 2000 t/h") são texto fixo.

## Parte D. Salvar

**D1.** **Arquivo > Salvar**. Abra no Process Explorer e maximize para conferir.

## Parte E. Problemas comuns

| Problema | Solução |
|---|---|
| Erro ao abrir | Tente `teste_minimo.atgraphic` e envie a captura (Parte A) |
| Campo com `***` | Tag sem valor no PIMS: confira a mesma tag na tela antiga |
| Estoque diferente da tela antiga | Me envie as duas telas lado a lado |

## Anexo. Fichas dos objetos dinâmicos

Posições em unidades da tela de 70 x 35.

![Mapa de posicionamento](mapa_posicionamento.png)

{{FICHAS}}
