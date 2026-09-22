# Instalação do modelo corrigido — passo a passo

Para quem baixou o repositório em
`C:\Users\00125723\Downloads\lab-analise-mes-claude-funny-dirac-ic9mx0`.

⏱ Uns 10 minutos. Tudo é reversível enquanto você mantiver o backup do passo 3.

---

## 1. Feche o Power BI Desktop

Feche **todas** as janelas. Se o arquivo estiver aberto, o Windows não deixa
renomear a pasta e o Power BI pode sobrescrever o que você copiar.

---

## 2. Abra a pasta do seu projeto

É a pasta onde está o arquivo `.pbip`. Você deve ver três itens:

```
📄 Gestão equioamentos power pbip.pbip
📁 Gestão equioamentos power pbip.Report
📁 Gestão equioamentos power pbip.SemanticModel   ← esta vai ser substituída
```

> **Anote o nome exato da pasta `...SemanticModel`.** Pode ser ligeiramente
> diferente do que está escrito acima. É esse nome que vale, não o meu.
> Dica: clique nela uma vez, aperte **F2**, copie o nome com **Ctrl+C**,
> e aperte **Esc** para não renomear ainda.

---

## 3. Faça o backup

Clique na pasta `...SemanticModel` → **F2** → acrescente `.BACKUP` no fim do nome:

```
Gestão equioamentos power pbip.SemanticModel.BACKUP
```

Não apague nada. Este é o seu ponto de retorno.

---

## 4. Copie a pasta nova

1. Abra `C:\Users\00125723\Downloads\lab-analise-mes-claude-funny-dirac-ic9mx0`
2. Entre em `pbip`
3. Clique uma vez na pasta `SemanticModel` e **Ctrl+C**
4. Volte à pasta do seu projeto e **Ctrl+V**

Agora existe uma pasta chamada apenas `SemanticModel` ali.

---

## 5. Renomeie para o nome exato

Clique na pasta `SemanticModel` → **F2** → cole (**Ctrl+V**) o nome que você
copiou no passo 2, **sem** o `.BACKUP`.

O resultado tem de ficar assim:

```
📄 Gestão equioamentos power pbip.pbip
📁 Gestão equioamentos power pbip.Report
📁 Gestão equioamentos power pbip.SemanticModel          ← a nova
📁 Gestão equioamentos power pbip.SemanticModel.BACKUP   ← a antiga
```

> Se o nome não bater **exatamente** — incluindo o acento em "Gestão" e o
> "equioamentos" escrito assim — o Power BI abre sem encontrar o modelo.

---

## 6. Abra o projeto

Duplo clique no arquivo `.pbip`.

**Vai pedir as credenciais do SharePoint.** É esperado: as consultas foram
reescritas, então o Power BI trata como origem nova. Escolha
**Conta organizacional** e entre com o seu login da Samarco.

A primeira atualização leva um pouco mais que o normal.

---

## 7. Confira se deu certo

| Onde olhar | O que tem de aparecer |
|---|---|
| Painel de campos → `_Medidas` | 8 pastas: `00 Cabeçalho` … `99 Auxiliares` |
| Painel de campos | `_Medidas_Inspecoes` **não existe mais** |
| Exibição de Modelo | 16 tabelas, **nenhuma** `LocalDateTable_…` |
| `DimCalendario` | 19 colunas (antes tinha 1) |
| Card **% Conformidade** | **57,0%** — antes era 63,5% |
| Cards **OMs Abertas** / **Lead Time Médio** | **em branco** — antes mostravam `0` |
| Página CALIBRAÇÕES | a tabela da direita continua preenchida |

As duas últimas linhas são as que mais confundem: **são correções, não defeitos.**
`% Conformidade` caiu porque o cálculo antigo descartava 22 registros do
denominador sem avisar. Os cards ficaram em branco porque `tbl_Notas` está
vazia — e branco significa "não sei", enquanto zero significa "não há nenhuma".

---

## 8. Aplique o tema

Ainda no Power BI: **Exibir ▸ Temas ▸ Procurar temas**, e escolha

```
C:\Users\00125723\Downloads\lab-analise-mes-claude-funny-dirac-ic9mx0\tema\samarco-tema.json
```

O painel inteiro muda de aparência de uma vez.

---

## 9. Salve

**Ctrl+S**. O Power BI grava de volta nos arquivos do projeto.

---

## Se der errado

1. Feche o Power BI Desktop
2. Apague a pasta `...SemanticModel` nova
3. Renomeie `...SemanticModel.BACKUP` de volta, tirando o `.BACKUP`
4. Abra o `.pbip` — está como antes

E me mande **a mensagem de erro exata** (print serve). São arquivos de texto;
dá para corrigir e gerar de novo.

---

## Erros comuns

| Sintoma | Causa | Solução |
|---|---|---|
| "Não foi possível carregar o modelo" | nome da pasta diferente | confira letra por letra com o `.BACKUP` |
| Não consegue renomear a pasta | Power BI ainda aberto | feche tudo e tente de novo |
| Tabela da página CALIBRAÇÕES vazia | não deveria acontecer — os emojis foram preservados | me avise |
| Pede credencial a cada atualização | nível de privacidade do SharePoint | Arquivo ▸ Opções ▸ Privacidade ▸ *Ignorar os níveis de privacidade* |
| Acentos estranhos nos nomes | ZIP extraído com ferramenta antiga | extraia com o próprio Windows |

---

## Depois que funcionar

A pasta `.Report` **não foi tocada** — as páginas continuam exatamente como
estão, com as sobreposições de visual. Essa é a Fase 2: me mande a pasta
`Gestão equioamentos power pbip.Report` que eu corrijo os desalinhamentos e
aplico a grade 1920×1080 do mockup.
