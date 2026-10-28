# Etapa 5 · Imagens de consulta e etiquetas QR Code

Duas ferramentas, **nesta ordem**. As duas funcionam offline: abra com dois cliques no Edge ou no Chrome.

```
1. gerador-imagens-consulta.html  ─► TAG.png de cada equipamento ─► biblioteca ConsultaQR no SharePoint
2. gerador-etiquetas.html          ─► etiquetas com QR = endereço da imagem ─► imprimir e colar
```

O QR Code de cada etiqueta é o endereço da imagem do equipamento: `PastaImagensQR/TAG.png`.
- **Câmera do celular:** abre a imagem (qualquer pessoa da empresa, sem app).
- **Scanner do app Metrologia:** lê o mesmo QR e abre a ficha (laboratorista e instrumentação).

Se a imagem ainda não estiver na biblioteca, a câmera abre uma página de "arquivo não encontrado". Por isso as imagens vêm primeiro.

## 1. Imagens de consulta (primeira carga, todas de uma vez)

1. No SharePoint, abra a lista **Equipamentos** > exibição **Imagens de consulta** > **Editar no modo de grade**.
2. Selecione todas as linhas (clique na primeira célula, Shift + clique na última) > **Ctrl+C**.
3. Abra `gerador-imagens-consulta.html` > cole no campo **Equipamentos** > **Gerar imagens**.
4. Confira os avisos em amarelo (equipamento sem data aparece como **A CONFIRMAR**). Ajuste a situação no seletor de cada imagem, se precisar. Equipamento sem calibração periódica: escolha **SEM CALIBRAÇÃO PERIÓDICA**.
5. **Baixar todas (.zip)** > extraia o .zip numa pasta.
6. Biblioteca **ConsultaQR** > **Carregar > Arquivos** > selecione todas as imagens.

> Sem acesso ao SharePoint ainda? O arquivo `imagens-consulta.txt` do pacote de dados tem as mesmas linhas, tiradas da planilha (quase todas sem data de calibração: confira com a instrumentação antes).

## 2. Imagens de consulta (no dia a dia)

As imagens só mudam em dois momentos, e um e-mail avisa você nos dois:

| Quando | E-mail | O que fazer |
|---|---|---|
| Calibração **aprovada** registrada no app | F3 · "Calibração registrada", com uma linha pronta | Cole a linha no gerador > baixe `TAG.png` > ConsultaQR > Carregar > **Substituir** |
| Calibração **venceu** ontem | F2, parte E · "Atualize a imagem", com as linhas NÃO UTILIZAR | Igual |

Aferição não conforme **não** muda a imagem (regra decidida): ela bloqueia o equipamento só no app, para a equipe.

> **Regra de ouro:** carregue sempre com o **mesmo nome** e escolha **Substituir**. Nunca apague nem renomeie a imagem: a etiqueta aponta para esse nome. A imagem anterior fica no histórico de versões da biblioteca.

## 3. Etiquetas

São três modelos, todos com QR = endereço da imagem.

| Modelo | Tamanho | Onde cola | Lista no pacote de dados |
|---|---|---|---|
| **Equipamento** | 80 × 40 mm | Balanças, tambores, prensas, analisadores, Alpine, Blaine... | `etiquetas-equipamentos.txt` |
| **Peneirador** | 100 × 50 mm | Na máquina de cada peneirador e no suporte das peneiras circulares | `etiquetas-peneiradores.txt` |
| **Peneira** | mini 50 × 25 mm | No aro de cada peneira | `etiquetas-peneiras.txt` |

1. Antes, confira na lista **Equipamentos** (coluna `ObservacaoCadastro`): TAGs repetidas em dois peneiradores, peneira sem TAG e a TAG provisória do conjunto de peneiras circulares.
2. Abra `gerador-etiquetas.html`. Cole o **endereço da biblioteca ConsultaQR** (o mesmo de `PastaImagensQR` na lista Configuracoes).
3. Escolha o **Modelo**. Tamanho, cabeçalho e rodapé são ajustados sozinhos.
4. Cole a lista do modelo (arquivo `.txt` correspondente) > **Gerar etiquetas**.
5. **Teste antes de imprimir:** toque numa etiqueta na tela (abre a imagem) e leia uma com a câmera do celular e com o scanner do app.
6. **Imprimir / salvar em PDF**. Repita para os outros modelos.

Use material resistente a limpeza e calor. Para as peneiras, prefira etiqueta de poliéster ou plaqueta presa ao aro: a peneira é lavada e manuseada com frequência. Na etiqueta mini, imprima em boa resolução (impressora a laser ou térmica de 300 dpi) para o QR ler bem.
