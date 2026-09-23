# 11 — Integração com o Microsoft Forms

## O que passou a acontecer

A planilha de respostas do formulário **Inspeção de Amostradores** agora
alimenta o painel junto com a planilha antiga. Uma coluna nova, `Origem`,
diz de onde cada linha veio: `Planilha` ou `Forms`.

Uma submissão do formulário carrega **dois fatos diferentes**, e por isso
alimenta duas tabelas:

| O que | Vai para | Alimenta |
|---|---|---|
| checklist, turno, responsável, tag, observações | `tbl_Inspecoes` | página INSPEÇÕES |
| tempos, distância, velocidade | `tbl_Medicoes_Operacionais` | gráfico dos cortadores |

Nenhuma medida, nenhum gráfico e nenhum filtro precisou mudar.

## O formulário é ramificado — e a consulta lida com isso

O formulário mostra um bloco de perguntas diferente conforme o tipo de
equipamento escolhido. São **três blocos**, e só o escolhido vem
preenchido; os outros ficam vazios na mesma linha.

Como o Forms nomeia cada coluna com o **texto da pergunta**, e acrescenta
sufixo numérico quando duas perguntas têm o mesmo enunciado, a planilha de
respostas tem colunas assim:

```
Tempo 1                                    Tempo 11
Checklist do cortador.Limpeza do sistema   Checklist do cortador.Limpeza do sistema1
          Observações   (opcional)                   Observações   (opcional)        1
```

Repare nos espaços duplicados e nos espaços não separáveis. Escrever a
consulta contra nomes exatos seria garantir que ela quebre.

**Nenhuma linha desta consulta referencia nome exato de coluna.** As
colunas são achadas por padrão sobre o nome normalizado (sem espaço, sem
pontuação, minúsculo), e os valores são recolhidos de todos os blocos —
o que estiver preenchido é o que vale:

| Campo | Como é achado |
|---|---|
| tempos | toda coluna cujo nome começa com `tempo` |
| checklist | contém `checklistdocortador` ou `divisorrotativo` — 29 colunas |
| observações | contém `observações` — 3 colunas, uma por bloco |
| distância | contém `distânciadepercurso` — 2 colunas |

Acrescentar uma pergunta nova de tempo, ou um quarto bloco, **não quebra
nada**: a consulta acha sozinha.

## As traduções que a consulta faz

**Turno.** O formulário usa `Letra A/B/C/D` e `ADM`; o modelo usa
`Turno A/B/C` e `Adm`. A consulta converte. E isso responde a sua dúvida
de ontem: se o formulário tiver `Letra D`, o **Turno D aparece sozinho**
no painel, sem mexer em nada.

**Responsável.** A planilha tem **duas** colunas `Nome`: a identidade de
quem enviou (o Forms coloca sozinho) e a pergunta com o nome do inspetor.
São coisas diferentes — na resposta 5, quem enviou foi você e quem
inspecionou foi o Luiz Guilherme. A consulta usa a **pergunta**, que é o
inspetor de verdade.

**TAG.** Vem como `U04-07AL001 · Correia Alimentadora PQ 04`. A consulta
separa no `·`: a tag vira a chave de ligação com o cadastro, o resto vira
o nome. Conferi as duas tags respondidas contra as 34 do cadastro —
**as duas batem**. A lista suspensa do formulário está correta.

**ID_Inspecao.** Gerado como `FORM-AAAAMMDD-TAG-NNNN`. Não repete o bug do
`yyyy` literal que existe na planilha antiga.

## Uma decisão que eu tomei e você precisa confirmar

**`Qtd_NC` e `Resultado_Inspecao` vêm da CONTAGEM dos itens do checklist,
não da pergunta "Encontrou alguma Não Conformidade?".**

Motivo: as duas respostas de teste já chegaram se contradizendo. Nas duas,
a pergunta-resumo foi respondida **"Não"**, mas o checklist tem **4 itens
marcados como "Não Conforme"** em cada uma.

Confiar na pergunta-resumo faria o painel dizer que está tudo conforme
enquanto quatro itens falharam. Então a consulta conta os itens. A
resposta da pergunta-resumo é preservada na coluna `Requer_Acao`, para
você comparar.

Se a intenção da pergunta for outra — tipo "precisa de ação imediata?" —
vale mudar o enunciado no formulário, porque do jeito que está ela induz
ao erro.

## Uma suposição que eu documentei

**A distância de percurso está sendo lida como CENTÍMETROS.**

O enunciado da pergunta traz `ref.: CR001 = 2,60 · CR002 = 1,68`, que são
metros. As duas respostas vieram com **260** — o número sem a vírgula.
Testei as três leituras possíveis:

| Se 260 for | Distância | Velocidade resultante |
|---|---|---|
| milímetros | 0,26 m | 0,04 a 0,06 m/s |
| **centímetros** | **2,60 m** | **0,44 e 0,61 m/s** |
| metros | 260 m | 44 a 61 m/s |

Só centímetros produz velocidade plausível para cortador de amostragem.
Adotei essa leitura, e a consulta converte para milímetros, que é a
unidade da coluna na planilha antiga.

**Mas isso é inferência, não especificação.** Se o campo for de texto
livre, uma hora alguém vai digitar `2,6` e outra `2600`, e o painel vai
mostrar velocidade errada sem avisar. A correção certa é no formulário:
diga a unidade no enunciado ("distância em centímetros") e, se der,
restrinja o campo a número.

## Outros pontos do formulário que valem revisão

- **Só 4 tempos.** O modelo antigo prevê 5. Não é problema — a média usa
  os que existirem —, mas se a norma pedir 5 passadas, falta uma.
- **`Criticidade` veio vazia** nas duas respostas. Se for para classificar
  a não conformidade, vale tornar obrigatória quando houver NC.
- **A pergunta-resumo de NC aparece 3 vezes** (uma por bloco). A consulta
  junta as três, mas duas sempre virão vazias.
- **As fotos** são guardadas como URL do SharePoint e a consulta traz a
  do equipamento na coluna `Foto`. Quem abrir o painel precisa de acesso
  à pasta das fotos para vê-las.

## Regra que continua valendo

**Mudar o enunciado de uma pergunta muda o nome da coluna.** A consulta
foi escrita para aguentar acréscimos e pequenas variações de espaço, mas
se você reescrever uma pergunta inteira, o padrão pode deixar de casar.
Pode mexer nas opções de resposta à vontade; me avise antes de reescrever
enunciado.

---

## Adendo v12 — linha vazia truncando o M

A v11 não abriu. O Power BI acusou:

```
Erro de mecanismo M: 'Esperava-se o token ','.'
```

Causa: ao injetar o bloco do Forms nas partições, eu gerei **linhas
totalmente vazias** entre os passos do M. No TMDL, o valor de uma
propriedade multi-linha é delimitado por **indentação**, e uma linha
totalmente vazia **encerra o bloco**. O M foi cortado logo depois de

```
DaPlanilha = Table.AddColumn( IdValido, "Origem", each "Planilha", type text ),
```

— uma vírgula sem nada depois. Daí a mensagem.

As linhas em branco do arquivo original não eram vazias: traziam os
quatro tabuleiros de indentação. A correção foi essa, e só ela: dar o
recuo às linhas em branco de dentro do M, **mantendo vazia** a linha que
fecha o bloco antes do `annotation`.

### A verificação que passou a existir

Um erro desses não devia chegar até você. O projeto ganhou um validador
que, para cada partição `= m`:

- extrai o bloco do M do jeito que o TMDL o delimita;
- recusa linha totalmente vazia dentro dele;
- exige que o bloco contenha `in` e não termine em branco;
- confere o balanceamento de `( )`, `[ ]` e `{ }`, ignorando o que está
  dentro de string e de comentário.

Rodando nas 14 partições M do modelo: **todas fechadas e balanceadas**.
Se a v11 tivesse passado por ele, teria sido barrada aqui.
