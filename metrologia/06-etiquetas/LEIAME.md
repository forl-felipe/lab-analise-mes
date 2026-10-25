# Etapa 5 · Etiquetas QR Code

São **três modelos**, todos com o mesmo link do app. O app decide o que mostrar pelo tipo do equipamento.

| Modelo | Tamanho | Onde cola | Quem usa | Lista no pacote de dados |
|---|---|---|---|---|
| **Equipamento** | 80 × 40 mm | Balanças, tambores, prensas, analisadores, Alpine, Blaine... | Todos | `etiquetas-equipamentos.txt` |
| **Peneirador** | 100 × 50 mm | Na máquina de cada peneirador e no suporte das peneiras circulares | Laboratorista verifica as peneiras do conjunto | `etiquetas-peneiradores.txt` |
| **Peneira** | mini 50 × 25 mm | No aro de cada peneira | Instrumentação (calibração e troca de posição) | `etiquetas-peneiras.txt` |

## Como gerar

1. Antes, confira na lista **Equipamentos** (coluna `ObservacaoCadastro`): TAGs repetidas em dois peneiradores, peneira sem TAG e a TAG provisória do conjunto de peneiras circulares.
2. Abra `gerador-etiquetas.html` com dois cliques (funciona sem internet).
3. Cole o **Link da Web** do app Metrologia.
4. Escolha o **Modelo**. Tamanho, cabeçalho e rodapé são ajustados sozinhos.
5. Cole a lista do modelo (arquivo `.txt` correspondente) e clique em **Gerar etiquetas**. Toque numa etiqueta para testar o link e imprima.
6. Repita para os outros modelos.

Use material resistente a limpeza e calor. Para as peneiras, prefira etiqueta de poliéster ou plaqueta presa ao aro: a peneira é lavada e manuseada com frequência.
