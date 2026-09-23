# 10 — Os quatro ajustes da v8

## 1. Ícones dos cartões: todos azuis

Você tem razão e o erro foi meu. Os ícones eram coloridos por **tema do
indicador** (verde para "em dia", vermelho para "vencidas", âmbar para "a
vencer"), mas ficam parados quando o número muda. Isso é exatamente o que
um farol não pode fazer: se a cor não reage ao valor, ela mente.

Os 22 ícones passaram todos para o azul institucional `#1B6FB0`. O que
distingue um cartão do outro agora é só o **desenho** — cubo, alerta,
selo, relógio, chave — que é um rótulo, não um julgamento.

Semáforo continua existindo onde ele de fato reage ao dado: na coluna
`Situação` da tabela de calibrações e nas faixas de criticidade.

## 2. "Última Atualização" era mentira — agora não é

A medida antiga era:

```dax
FORMAT ( UTCNOW () - TIME ( 3, 0, 0 ), "dd/MM/yyyy HH:mm" )
```

`UTCNOW()` é recalculado **toda vez que a tela é desenhada**. O cartão
mostrava a hora de agora, não a da última carga — e com o título "Última
Atualização" em cima. Você identificou certo: era um relógio disfarçado.

A correção é uma tabela nova de uma linha só, `Atualizacao`, cujo valor é
gravado **no momento da atualização**:

```m
let
    Agora  = DateTime.From( DateTimeZone.SwitchZone( DateTimeZone.UtcNow(), -3 ) ),
    Origem = #table( type table [ Carimbo = datetime ], { { Agora } } )
in
    Origem
```

Power Query roda isso uma vez por atualização e **congela** o resultado
dentro do modelo. A medida passou a ler esse carimbo:

```dax
Última Atualização =
VAR c = MAX ( Atualizacao[Carimbo] )
RETURN IF ( ISBLANK ( c ), "sem atualização", FORMAT ( c, "dd/MM/yyyy HH:mm" ) )
```

Agora o cartão só muda quando os dados mudam. Se ele estiver velho, é
porque a atualização não rodou — que é justamente a informação que o
cartão deveria dar.

### Sobre atualizar de hora em hora

Precisa ser dito com clareza: **no Power BI Pro o limite é 8 atualizações
por dia.** De hora em hora (24 por dia) só em capacidade **Premium por
usuário (PPU)** ou **Premium/Fabric**, onde o limite é 48.

Com 8 por dia dá para cobrir bem o horário de trabalho, por exemplo
06:00, 08:00, 10:00, 12:00, 14:00, 16:00, 18:00 e 20:00. Para estes
dados — calibração, inspeção, paradas — isso é mais do que suficiente.
Se a área quiser de hora em hora mesmo, é conversa de licença com a TI,
não de configuração.

## 3. EQUIPAMENTOS: os dois gráficos mostravam a mesma coisa

Confirmado no dado, e você está certo:

| Grupo Operacional | | Área | |
|---|---|---|---|
| PF · Usina 03 | 9 | PF-US03 | 9 |
| PF · Usina 04 | 9 | PF-US04 | 9 |
| Embarque · Torre 3 | 8 | EMB | 8 |
| PQ · Usina 03 | 4 | PQ-US0304 | 8 |
| PQ · Usina 04 | 4 | | |

`Área` é o **mesmo corte** que `Grupo Operacional`, só com as duas usinas
do PQ somadas. Dois gráficos para a mesma informação.

O segundo gráfico passou a ser **Equipamentos por Tipo**, que é outro
eixo de verdade — diz *o que* é o equipamento, não *onde* ele está:

| Tipo | Equipamentos |
|---|---|
| Cortador Primário | 10 |
| Correia Alimentadora | 10 |
| Cortador Secundário | 10 |
| Correia Primária | 2 |
| Correia Secundária | 2 |

O filtro de Área continua na barra lateral — você não perde o corte, só
deixa de gastar meia página com ele.

## 4. Calibrações vencidas: conferi, e o número está certo

Sua desconfiança era legítima, então não acreditei na minha própria
medida: contei por fora, direto da planilha, sem passar pelo Power BI.

Lendo `Próximo Vencimento` das 208 linhas e comparando com a data de
hoje:

| | |
|---|---|
| **Vencido** | **65** |
| Em dia | 126 |
| Sem data interpretável | 17 |

O painel mostra **65**. Bate.

E fui além, porque a sua hipótese era de confusão entre as duas datas.
Testei se `Última Calibração + Frequência (meses)` fecha com
`Próximo Vencimento` nas linhas em que os três campos existem:

| | |
|---|---|
| Bate | **190** |
| Diverge | 2 |
| Sem dado para testar | 16 |

As duas divergências são erros de digitação na planilha, não do painel:

- **Balança 71** — calibrada em 01/03/2025, frequência 12 meses, mas o
  vencimento está gravado como 31/03/2027 (dois anos depois).
- **Peneira Circular 200#** — `Última Calibração` está em **15/07/2027**,
  uma data no futuro.

Ou seja: a métrica está correta, e a planilha tem duas linhas para
revisar.

### O que mudou na tabela

Você estava certo em pedir: eu tinha trazido `Próximo Vencimento` mas
deixado de fora `Última Calibração`, que é o que permite conferir. A
tabela da página CALIBRAÇÕES agora tem cinco colunas:

| Equipamento | Laboratório | **Última Calibração** | **Vencimento** | **Situação** |
|---|---|---|---|---|

A coluna `Situação` é o farol de verdade — 🔴 Vencido, 🟠 Atenção,
🟡 Crítico, 🟢 Em Dia — e ele reage ao dado, diferente dos ícones dos
cartões.

Vale lembrar de uma distinção que confunde: a planilha tem uma coluna
`Status Calibração` com OK / N/OK / E/C. Isso é **status metrológico**
(aprovado, reprovado, em calibração) — não diz se a calibração está
vencida. São duas perguntas diferentes, e por isso são duas colunas
diferentes no modelo: `Status Metrologico` e `StatusAtual`.

---

## Adendo v9 — por que as páginas abriam em branco

Não era o download, nem a cópia, nem o Power BI. Era o **limite de 260
caracteres de caminho do Windows**, e a conta fecha no caractere:

| | |
|---|---|
| `C:\Users\<usuário>\Downloads\` | 28 |
| pasta criada pelo "Extrair tudo" | 42 |
| pasta que vem dentro do zip | 42 |
| `Gestão equioamentos power pbip\` | 31 |
| `Gestão equioamentos power pbip.Report\` | 38 |
| `definition\pages\` | 17 |
| id da página | 21 |
| `visuals\` | 8 |
| id do visual | 21 |
| `visual.json` | 11 |
| **total** | **259** |

O limite útil do Windows é **259 caracteres**. O `page.json` fica em 228 e
passa; o `visual.json` fica em 259 e não passa. Por isso as 9 páginas
apareciam e **todas** ficavam vazias, de forma uniforme — o corte não era
por página, era por profundidade.

O Explorer não avisa quando isso acontece. Ele extrai o que cabe e
termina sem erro. O sintoma que denunciou foi a contagem: 57 arquivos na
pasta do relatório contra os 487 esperados. Os 430 que faltavam eram
exatamente os do nível mais fundo.

**A correção foi encurtar os nomes**, que são a única parte do caminho
sob meu controle:

| Antes | Depois | Economia |
|---|---|---|
| `Gestão equioamentos power pbip` | `Painel Samarco` | 16 |
| `Gestão equioamentos power pbip.Report` | `Painel.Report` | 24 |
| `Gestão equioamentos power pbip.SemanticModel` | `Painel.SemanticModel` | 24 |

São 49 caracteres a menos por caminho. O pior caso passa de 259 para
**219**, com 40 de folga.

Os dois arquivos que amarram esses nomes foram atualizados junto:
`Painel.pbip` aponta para `Painel.Report`, e
`Painel.Report/definition.pbir` aponta para `../Painel.SemanticModel`.
O nome exibido dentro do Power BI, que fica em `.platform`, não mudou.

Lição para a próxima entrega: nome de pasta bonito custa caracteres, e
caractere de caminho é um recurso escasso no Windows quando a estrutura
tem seis níveis e um arquivo por visual.

---

## Adendo v10 — o `$schema` do `.pbip` estava errado

Depois de resolvido o caminho longo, o Power BI passou a abrir o arquivo
e recusou com uma mensagem precisa:

```
Expected '$schema' property in '...\Painel.pbip' to follow patterns:
^https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.[0-9]+.[0-9]+/schema.json$
```

O `.pbip` foi o único arquivo deste projeto que **eu escrevi do zero**, na
v5, quando passei a entregar a pasta do projeto inteira. Escrevi:

```
.../json-schemas/fabric/item/pbipProperties/1.0.0/schema.json
                         ^^^^
```

quando o correto é:

```
.../json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json
                         ^^^^
```

Uma palavra. Todos os outros arquivos de controle vieram do `.pbip`
original exportado pelo Power BI na máquina do Felipe, e por isso estavam
certos desde o começo — o erro estava exatamente no único que eu inventei
de memória.

O validador do projeto agora confere o `$schema` dos quatro arquivos de
controle (`.pbip`, `definition.pbir`, `report.json`, `pages.json`) contra
os padrões que o Power BI exige, para que esse tipo de erro não passe de
novo.
