**Assunto:** Solicitação de tags calculadas no PIMS (Aspen Calc) para a tela do ensaio de compressão de pelotas

Prezados,

O Laboratório de Controle de Embarque (LCE) está colocando em uso uma tela no PIMS (Aspen Process Explorer) para acompanhar a prensa automática de compressão de pelotas. A tela já está pronta e funcionando com as tags existentes. Para completá-la, precisamos de algumas **tags calculadas**, se possível no Aspen Calc.

**Contexto**

- A prensa automática ensaia uma amostra a cada 30 minutos e grava o resultado na tag `307-GI-001` (média de compressão do ensaio, em kgf/pelota). O resultado entra no PIMS cerca de 20 minutos depois do corte. Exemplo: o corte das 15:00 aparece no PIMS às 15:20.
- O laboratório analisa uma amostra composta de 2 horas, formada por 4 cortes (um a cada 30 min). O resultado vem do MES pela tag `M4710050020-0031-HHLFU` (mapa `IP_MESVALOR`).
- Para comparar a prensa com o laboratório de forma correta, precisamos da média dos cortes da prensa na mesma janela de 2 horas do composto do laboratório.

**Tags solicitadas**

| Tag (nome sugerido) | O que calcula | Horário gravado |
|---|---|---|
| `307-GI-001-C2H` | Média de `307-GI-001` dos cortes com resultado na janela de 2 h | Fim da janela |
| `307-GI-001-N2H` | Quantidade de cortes com resultado na janela (0 a 4) | Fim da janela |
| `307-GI-001-DIF2H` | `307-GI-001-C2H` menos `M4710050020-0031-HHLFU` da mesma janela | Fim da janela |
| `307-GI-001-CORTE1` | Resultado de `307-GI-001` do 1º corte da janela atual | Horário do corte |
| `307-GI-001-CORTE2` | Resultado do 2º corte da janela atual | Horário do corte |
| `307-GI-001-CORTE3` | Resultado do 3º corte da janela atual | Horário do corte |
| `307-GI-001-CORTE4` | Resultado do 4º corte da janela atual | Horário do corte |

Todas em kgf/pelota com 1 casa decimal, exceto `307-GI-001-N2H`, que é um número inteiro.

**Regras do cálculo**

1. **Janela de 2 horas pelo horário do corte.** As janelas começam nas horas pares. Exemplo: a janela das 14:00 às 16:00 reúne os cortes de 14:00, 14:30, 15:00 e 15:30. Como o resultado entra no PIMS 20 minutos depois do corte, o horário de entrada deve ser ajustado para o horário do corte antes de montar a janela.
2. **Janela incompleta.** Se a prensa falhar e a janela tiver só 1, 2 ou 3 cortes com resultado, a média é feita só com os cortes que existem, e `307-GI-001-N2H` informa quantos foram. Sem nenhum corte, a média fica sem valor.
3. **Valores que não entram na média.** A tag `307-GI-001` volta a 0 entre os ensaios. Valores iguais a 0 e resultados com qualidade diferente de Good não devem entrar na média.
4. **Atualização.** `307-GI-001-C2H` e `307-GI-001-N2H` podem ser recalculadas a cada novo resultado da prensa, para o valor ir se completando ao longo da janela. `307-GI-001-DIF2H` é calculada quando o resultado do MES daquela janela chegar.
5. **Cortes da janela atual.** As tags `CORTE1` a `CORTE4` mostram os resultados dos cortes da janela em andamento (ou da última janela fechada). Corte sem resultado fica sem valor.

**Exemplo**

Janela das 14:00 às 16:00, com a prensa sem resultado no corte das 15:00:

| Corte | `307-GI-001` |
|---|---|
| 14:00 | 312,4 |
| 14:30 | 305,8 |
| 15:00 | sem resultado |
| 15:30 | 298,1 |

Resultado esperado: `307-GI-001-C2H` = 305,4 (média de 3 cortes), `307-GI-001-N2H` = 3, gravados às 16:00. Se o laboratório informar 299 para essa janela, `307-GI-001-DIF2H` = 6,4.

**Observações**

- A tela já está configurada com esses nomes. Assim que as tags forem criadas, os valores aparecem automaticamente, sem nenhuma alteração na tela. Se for necessário usar outro padrão de nomes, basta nos informar.
- Fico à disposição para validar os primeiros resultados junto com vocês.

Atenciosamente,

Felipe
Laboratório de Controle de Embarque (LCE), UBU
