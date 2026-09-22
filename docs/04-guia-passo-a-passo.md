# Guia passo a passo — como aplicar tudo isso no Power BI

Escrito para quem está começando. Cada passo diz **onde clicar**.

---

## Antes de tudo: faça uma cópia

Abra a pasta onde está o `.pbix` e **duplique o arquivo**:

```
Gestão equipamentos power.pbix            ← guarde, não mexa
Gestão equipamentos power - TRABALHO.pbix ← é neste que você trabalha
```

Se algo der errado, você volta ao original. Nenhum dos passos abaixo é
irreversível **se** você tiver essa cópia.

---

## As três formas de fazer isso

### Caminho A — manual, copiando e colando
Você abre o Power BI, vai em cada consulta e cola o código. **Não precisa de
nenhuma ferramenta nova.** Leva cerca de 3 a 4 horas, divididas em sessões.

### Caminho B — semi-automático com o Tabular Editor
O [Tabular Editor 2](https://github.com/TabularEditor/TabularEditor/releases) é
gratuito. Ele conecta no Power BI Desktop aberto e cria as **50 medidas de uma
vez**, com pasta e formato, em vez de você criar uma a uma. Economiza a parte
mais repetitiva. As consultas do Power Query continuam sendo coladas à mão.

### Caminho C — automático, eu editando os arquivos
Existe e funciona, mas exige uma conversão: no Power BI Desktop,
*Arquivo ▸ Salvar como ▸ **Projeto do Power BI (.pbip)***. O modelo deixa de ser
um binário e vira arquivos de texto (TMDL). Aí você me envia a pasta, eu reescrevo
os arquivos direto e te devolvo prontos — você só abre.

---

## O que eu **não** consigo fazer

Vale ser claro para você não perder tempo:

- **Não consigo abrir o Power BI Desktop nem clicar por você.** Ele roda no
  Windows, na sua máquina. Eu rodo num servidor Linux, sem acesso ao seu
  computador. Não existe configuração que mude isso.
- **O Git/GitHub não "aplica" nada no Power BI.** Git é um sistema de controle
  de versão: ele guarda histórico de arquivos de texto. Ele não conversa com o
  Power BI nem injeta código nele. O que o Git resolve é *outro* problema — ter
  histórico do que mudou e poder voltar atrás. Útil, mas depois.

---

## Recomendação para o seu caso

**Comece pelo Caminho A.** Não é o mais rápido, é o que te ensina o modelo.
Colando consulta por consulta você vai entender por que cada etapa existe — e
quando o painel der problema daqui a três meses, você vai saber onde olhar.
Quem pula essa parte fica dependente de quem escreveu.

Depois que a Fase 1 estiver funcionando, o Caminho C passa a valer muito a pena
para a Fase 2 (o layout), porque aí o trabalho é repetitivo e visual.

---

## SESSÃO 1 — a base (40 min)

### 1.1 Abrir o editor
No Power BI Desktop: **Página Inicial ▸ Transformar dados**.
Abre o *Editor do Power Query*, com a lista de consultas à esquerda.

### 1.2 Criar o parâmetro
**Página Inicial ▸ Gerenciar Parâmetros ▸ Novo Parâmetro**

| Campo | Valor |
|---|---|
| Nome | `Parametro_CaminhoBase` |
| Tipo | Texto |
| Valor Atual | a URL do `.xlsx` no SharePoint |

Para pegar a URL: clique em `tbl_Equipamentos` na lista, veja a primeira etapa
em *Etapas Aplicadas* à direita, e copie o endereço que aparece.

### 1.3 Criar a consulta Fonte_Base
**Página Inicial ▸ Nova Fonte ▸ Consulta Nula**

Uma consulta chamada `Consulta1` aparece. Com ela selecionada:
**Página Inicial ▸ Editor Avançado** → apague tudo → cole o conteúdo de
`powerquery/01-Fonte_Base.pq` → **Concluído**.

Renomeie para `Fonte_Base` (clique com o botão direito na consulta ▸ Renomear).

> ⚠ **Clique com o botão direito em `Fonte_Base` e DESMARQUE "Habilitar carga".**
> Se esquecer, ela vira uma tabela inútil no seu modelo. O nome dela fica em
> *itálico* quando estiver certo.

### 1.4 Criar as 5 funções
Repita o mesmo procedimento (Consulta Nula ▸ Editor Avançado ▸ colar ▸ renomear
▸ desmarcar "Habilitar carga") para cada uma, **nesta ordem**:

1. `fx_Tabela` — arquivo `02-`
2. `fx_Config` — arquivo `03-`
3. `fx_NormalizarTag` — arquivo `04-`
4. `fx_ParseData` — arquivo `05-`
5. `fx_QualidadeData` — arquivo `06-`

A ordem importa: `fx_Config` usa `fx_Tabela`. Se inverter, aparece um erro que
some sozinho quando a outra for criada — mas é melhor não se assustar à toa.

> 💡 Para organizar: botão direito na área das consultas ▸ **Novo Grupo** ▸
> chame de `_Funções` e arraste as 6 para lá.

### 1.5 Ainda NÃO clique em "Fechar e Aplicar"
Continue na Sessão 2. Aplicar agora não quebra nada, só demora à toa.

---

## SESSÃO 2 — as consultas (60 min)

Para **cada** consulta da tabela abaixo: clique nela na lista à esquerda →
**Página Inicial ▸ Editor Avançado** → apague tudo → cole o arquivo → **Concluído**.

| Ordem | Consulta no Power BI | Arquivo |
|---|---|---|
| 1 | `tbl_Config` | `20-` |
| 2 | `tbl_Equipamentos` | `11-` |
| 3 | `tbl_Calibracao` | `10-` |
| 4 | `tbl_Grupos` | `17-` |
| 5 | `Tbl_Paradas` | `15-` |
| 6 | `tbl_Intervencoes` | `16-` |
| 7 | `tbl_Notas` | `14-` |
| 8 | `tbl_Inspecoes` | `12-` |
| 9 | `tbl_Medicoes_Operacionais` | `13-` |
| 10 | `tbl_Meta_Inspecoes` | `18-` |

`tbl_Config` vem primeiro porque `tbl_Calibracao` depende dela para ler os
limiares de 30 e 15 dias.

### Consultas novas
Três não existem ainda. Crie do mesmo jeito da Sessão 1 (Consulta Nula ▸ colar
▸ renomear), mas **deixe "Habilitar carga" MARCADO** — estas são tabelas de verdade:

- `tbl_Sobressalentes` — arquivo `19-`
- `DimLaboratorio` — arquivo `21-`
- `DePara_Inspecao_Equipamento` — arquivo `22-`

### Agora sim: Fechar e Aplicar
**Página Inicial ▸ Fechar e Aplicar**. Vai levar um ou dois minutos.

### O que conferir
Se aparecer erro, ele diz o nome da consulta. Os mais prováveis:

| Mensagem | Causa | Solução |
|---|---|---|
| `fx_Tabela não foi reconhecido` | função não criada ou nome com erro de digitação | confira o nome exato |
| `Não foi possível localizar o item` | nome da tabela no Excel mudou | confira a aba na planilha |
| Credenciais do SharePoint | o parâmetro mudou a URL | *Fontes de Dados Recentes* ▸ editar permissões |

---

## SESSÃO 3 — o modelo (45 min)

### 3.1 Desligar a data/hora automática
**Arquivo ▸ Opções e configurações ▸ Opções ▸ Arquivo Atual ▸ Carregar Dados**
→ desmarque **"Data/hora automática"** → OK.

Isso apaga 9 tabelas ocultas que estavam inflando o arquivo à toa.

### 3.2 Criar o calendário
Vá para a exibição de **Tabela** (ícone da grade, à esquerda) →
**Página Inicial ▸ Nova Tabela** → cole `dax/01-DimCalendario.dax`.

Depois:
1. Apague a `DimCalendario` **antiga** (a de uma coluna só).
2. Renomeie a coluna `Date` para `Data` (duplo clique no cabeçalho).
3. Selecione a tabela → **Ferramentas de Tabela ▸ Marcar como tabela de datas**
   → escolha a coluna `Data`.

### 3.3 Recriar DimTecnicos
Apague a `DimTecnicos` atual → **Nova Tabela** → cole `dax/02-DimTecnicos.dax`.

### 3.4 Os relacionamentos
Exibição de **Modelo** (terceiro ícone à esquerda). A lista completa está em
`docs/03-plano-implementacao.md`, seção 1.6. Dois cuidados:

- **Apague** o relacionamento `tbl_Calibracao → tbl_Equipamentos`. Ele parece
  certo e não casa nenhuma linha das 214.
- No relacionamento `tbl_Notas → tbl_Equipamentos`, dê duplo clique e mude
  *Direção do filtro cruzado* de **Ambos** para **Única**.

Para criar: arraste a coluna de uma tabela sobre a coluna da outra.

---

## SESSÃO 4 — as medidas (60 min, ou 15 com o Tabular Editor)

### Manual
Clique na tabela `_Medidas` → **Página Inicial ▸ Nova Medida** → cole a fórmula
→ Enter. Depois, no painel direito, preencha **Pasta de Exibição** e **Formato**
(os dois estão anotados no comentário acima de cada medida).

### Com o Tabular Editor 2 (recomendado aqui)
1. Baixe, instale e abra o **Power BI Desktop com o arquivo aberto**.
2. No Power BI: **Ferramentas Externas ▸ Tabular Editor**.
3. Lá dentro: botão direito em `_Medidas` ▸ *Create* ▸ *Measure*, cole, repita —
   ou use o *Advanced Scripting* para criar todas de uma vez.
4. **Ctrl+S** no Tabular Editor grava de volta no Power BI.

### Limpeza (não pule)
1. Mova as 9 medidas de `_Medidas_Inspecoes` para `_Medidas`: selecione cada
   uma e mude a propriedade **Tabela Inicial**. Isso **não quebra visual nenhum**.
2. Apague a tabela `_Medidas_Inspecoes`, agora vazia.
3. Apague as 3 medidas resíduo: `DAXQtd Equipamentos Distintos`,
   `DAXInstrumentos Sem Data`, `DAXQtd Inspecoes`. Conferi no arquivo do
   relatório — nenhuma está em uso, pode apagar sem medo.
4. Oculte a coluna `Value` de `_Medidas`.

---

## SESSÃO 5 — o tema (5 min)

**Exibir ▸ Temas ▸ Procurar temas** → escolha `tema/samarco-tema.json`.

O painel inteiro muda de cara de uma vez. Depois disso, vale remover a
formatação que foi aplicada manualmente em alguns visuais (um card tem sombra
customizada e o vizinho não) — com o tema, o padrão vem sozinho.

---

## Armadilhas que você vai encontrar

**1. A tabela da página CALIBRAÇÕES ficar vazia.**
Ela tem um filtro fixo em `CriticidadeCalibracao IN ('🔴 Vencido', '🟠 Atenção',
'🟡 Crítico')`. Eu mantive os emojis exatamente iguais por causa disso — mas se
você editar esses rótulos, a tabela esvazia **sem mensagem de erro**. Se isso
acontecer, é aqui.

**2. Os números vão mudar, e está certo.**
`% Conformidade` cai de 63,5% para 57,0%. Não é bug: o cálculo antigo excluía 22
registros do denominador em silêncio. Se alguém perguntar, a explicação está em
`docs/02-decisoes-arquitetura.md`, decisão D7.

**3. Cards vão ficar em branco.**
`OMs Abertas` e `Lead Time Médio` param de mostrar zero e passam a ficar vazios,
porque `tbl_Notas` não tem nenhuma linha. **Isso é o comportamento correto** —
zero afirma que não há OM aberta; vazio afirma que não se sabe.

**4. "Habilitar carga" esquecido.**
Se aparecerem tabelas `Fonte_Base` ou `fx_...` no painel de campos, você
esqueceu de desmarcar. Volte em *Transformar dados*, botão direito, desmarque.

**5. Não faça tudo de uma vez.**
Feche e aplique ao fim de cada sessão e confira se o painel ainda abre. É muito
mais fácil achar o que quebrou entre 10 mudanças do que entre 60.

---

## Pendência na planilha: cadastrar os 3 equipamentos

Confirmado que Cross Belt, Bomba de Polpa e Divisor Rotativo entram no mestre.
Na aba **Equipamentos**, acrescente uma linha para cada, preenchendo:

| Coluna | O que colocar |
|---|---|
| `Tag` | a TAG real de campo |
| `Equipamento` | Amostrador Cross Belt 01 · Bomba de Polpa Amostragem 01 · Divisor Rotativo 01 |
| `Categoria` | Amostragem |
| `Área` / `Sistema` / `Tipo` | conforme os demais |
| `Grupo Operacional` | um dos 5 grupos existentes |
| `Laboratório` | LCP ou LCE |
| `Crítico` | **Sim ou Não de verdade** — hoje as 34 linhas estão "Sim", e é por isso que o card "Equipamentos Críticos" repete o total |
| `Status Operacional` | idem — hoje as 34 estão "Operante" |

Depois, preencha `TagEquipamento` para essas 3 linhas em
`powerquery/22-DePara_Inspecao_Equipamento.pq`.

> **Atalho:** se no campo esses equipamentos são conhecidos justamente por
> `AM-CB-01`, `AM-BP-01` e `AM-DR-01`, use essas TAGs no cadastro. O de-para
> vira identidade e some um ponto de manutenção.

**Correção de fundo:** a lista suspensa de equipamento no formulário de inspeção
precisa ser alimentada pelo cadastro mestre. Sem isso, o de-para dessincroniza
de novo a cada equipamento novo.
