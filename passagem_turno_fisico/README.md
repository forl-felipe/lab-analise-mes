# Relatório de Passagem de Turno do Laboratório Físico: Usinas 3 e 4

Planilha `Relatorio_Fisico_US3_US4.xlsm`, no mesmo padrão do Relatório Químico (`../informativo_quimico_us34`):
modelo diário (um arquivo por dia, turnos Dia 07h–19h e Noite 19h–07h), um técnico por letra.

## Abas

| Aba | Conteúdo |
|---|---|
| **Preenchimento** | Única aba de digitação. Turno Dia em cima e Turno Noite embaixo: letra, técnico, produto US3/US4 e, na ordem do informativo enviado hoje: 1 Produção – Usinas 03/04 (comentário de qualidade por usina), 2 Dragas, 3 Mineroduto 03 (batch e teores médios), 4 Filtragem, 5 Solicitações, 6 Insumos, 7 Mercado interno, 8 Carregamento, 9 Equipamentos em operação, 10 Observações gerais |
| **Resumo Dia** / **Resumo Noite** | Uma página: identificação, ocorrências (só o que foi preenchido) e resultados do MES por horário de amostra (07:30, 09:30…). Botões **Atualizar dados do MES** e **Copiar imagem** |
| **Resultados gerais** | Resultados de qualquer período de até 24 h, nas mesmas janelas do turno |
| Limites (oculta) | Farol por produto (SMIN-POP-GEA-001 rev. 12), editável |
| Configurações (oculta) | Fonte dos dados, servidor e tags do MES |

## Preenchimento → Resumo

- Assunto sem nada preenchido não aparece no Resumo (título incluído). Linha vazia também não.
- Mineroduto: o Resumo monta a frase "Processando batch 267 com teores: SiO2: 1,83; P: 0,060; PPC: 3,64; SE: 1983; -325#: 85,4" com os campos preenchidos.
- Texto longo quebra a linha e a altura se ajusta sozinha.

## Resultados do MES

Ordem no relatório: alimentação da grelha, depois os químicos (Pellet Feed · Químico, Pelota Queimada · Químico), depois os físicos.
Análises e tags:

- Alimentação da grelha;
- Prensa de rolos (SE, #325, H2O);
- Pellet Feed (PPC, SiO2, CaO), em **uma linha US3/4**;
- **Filtrado para o pátio**: umidade (`M0650030020-0001-HHLFU`), em uma linha US3/4; vermelha acima de 11,30 % (máximo editável na aba Limites, Limites gerais);
- **Pelota Crua** (US3 `M690025000`, US4 `M4690025000`): umidade (0001), resiliência (0120), compressão seca (0095) e compressão úmida (0094). **Regra:** a umidade da pelota crua fica vermelha quando é maior ou igual à H2O de saída da prensa de rolos (mesma usina e mesmo horário) e verde quando é menor. A margem fica na aba Limites (Limites gerais, padrão 0);
- Mistura (SiO2, CaO, B2, carvão, carbono fixo);
- Pelota queimada (+16 −8 mm, relação granulométrica, tamboramento, compressão, < 200 kgf, finos −6,3 mm);
- Produção e ritmo.

Resultado do turno: média; Produção = soma; Ritmo = último valor.
Três consultas ao Aspen (Qualidade, Produção, Ritmo), com o mesmo formato de fórmula do informativo atual.

Farol (verde dentro do limite / vermelho fora), pelo produto escolhido:

| Produto | SiO2 máx. | CaO mín. | B2 | Tamboramento mín. | Finos máx. | Rel. granulométrica |
|---|---|---|---|---|---|---|
| PDR/MX | 1,54 | 0,70 | – | 93,0 | 1,80 | 0,5 a 1,2 |
| PDR/STD | 2,05 | 0,65 | – | 93,0 | 1,80 | 0,5 a 1,2 |
| PBF/MB45 | 3,45 | – | máx. 0,55 | 92,8 | 2,10 | 0,5 a 1,2 |
| PBF/STD | 3,20 | – | mín. 0,75 | 92,8 | 2,10 | 0,5 a 1,2 |
| PBF/HB | 2,80 | – | mín. 0,95 | 92,8 | 2,10 | 0,5 a 1,2 |
| PBF/SF | 3,00 | – | mín. 1,10 | 92,8 | 2,10 | 0,5 a 1,2 |
| PBF/SA | 5,30 | – | mín. 0,40 | 92,8 | 2,10 | 0,5 a 1,2 |

Pellet Feed, por concentrado (CLS, CNS, CHS ou CSP):
- SiO2 máx.: 1,36 / 1,99 / 2,50 / 5,30;
- PPC máx.: 4,30.

A compressão média não tem farol, porque o padrão define limites por faixa granulométrica (CCS 16+12,5 e 12,5+10).

**Desempenho:** igual ao Químico.
- Cada consulta é calculada uma vez por clique.
- Depois da leitura, a fórmula do Aspen vira texto.
- Uma tag recusada pelo MES sai da consulta, com aviso.
- O período em andamento é consultado só até a janela atual.

## Gerar

```
cd build
python3 build_fisico.py           # ../Relatorio_Fisico_US3_US4.xlsm
python3 build_fisico.py --teste   # versão com teste automático (LibreOffice)
```
