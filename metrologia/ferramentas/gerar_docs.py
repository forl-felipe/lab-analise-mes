"""
Gera as partes da documentação que dependem do modelo de dados:
  - 02-sharepoint/GUIA-SHAREPOINT.md   (passo a passo manual com as tabelas de colunas)
  - 01-arquitetura/DICIONARIO-DE-DADOS.md

Uso:  python gerar_docs.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from modelo import GRUPOS, LISTAS  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]

TIPO_SP = {
    "Text": "Linha única de texto",
    "Note": "Várias linhas de texto (texto sem formatação)",
    "Number": "Número",
    "Date": "Data e hora · **Incluir hora: Não**",
    "DateTime": "Data e hora · **Incluir hora: Sim**",
    "Boolean": "Sim/Não",
    "Choice": "Escolha",
    "User": "Pessoa ou Grupo · seleção múltipla: Não",
}


def config_coluna(c) -> str:
    partes = []
    if c.obrigatoria:
        partes.append("Obrigatória")
    if c.tipo == "Choice":
        partes.append("Opções: " + ", ".join(f"`{x}`" for x in c.choices))
    if c.padrao is not None and c.tipo == "Boolean":
        partes.append("Padrão: " + ("Sim" if c.padrao == "1" else "Não"))
    elif c.padrao is not None:
        partes.append(f"Padrão: `{c.padrao}`")
    if c.indexada:
        partes.append("**Indexar**")
    return " · ".join(partes)


def tabela_lista(L) -> str:
    linhas = [
        f"**Título (Title):** {L.titulo_significa}" + (" · **Impor valores exclusivos: Sim** (e indexar)" if L.titulo_unico else " · **indexar**" if L.titulo_indexado else ""),
        "",
        "| Nome exato da coluna | Tipo no SharePoint | Configuração | Para que serve |",
        "|---|---|---|---|",
    ]
    for c in L.colunas:
        linhas.append(f"| `{c.nome}` | {TIPO_SP[c.tipo]} | {config_coluna(c)} | {c.descricao} |")
    return "\n".join(linhas)


def guia_sharepoint() -> str:
    novas = [L for L in LISTAS if not L.existente]
    exist = [L for L in LISTAS if L.existente]
    s = []
    s.append("""# Etapa 1 · SharePoint (a base de dados)

> **Resumo:** são **9 listas** no mesmo site. Duas já existem desde a versão 1 (**Equipamentos** e **Calibracoes**):
> nelas você só **acrescenta colunas**. As outras sete são novas.
> A lista **Calibradores** da versão 1 é substituída pela **Acessos**. Mantenha a antiga até o app novo entrar no ar e depois apague.

## 1.0 Duas formas de fazer

| Caminho | Quem faz | Tempo |
|---|---|---|
| **A. Script automático** `criar-estrutura.ps1` | TI ou quem tiver PnP.PowerShell | 5 minutos (cria listas, colunas, índices, grupos, permissões e importa os dados) |
| **B. Manual**, seguindo este guia | Você | 1h30 a 2h |

O resultado é o mesmo. Se for pelo **caminho A**, pule para a seção **1.4 Acessos**, que sempre é manual.

```
./criar-estrutura.ps1 -SiteUrl https://EMPRESA.sharepoint.com/sites/SITE -ClientId <id do app do PnP> -ImportarDados -PastaCsv <pasta>\\csv
```

## 1.1 Regras de ouro (caminho manual)

1. **Nome da coluna exatamente como na tabela:** sem acento, sem espaço, com as mesmas maiúsculas. O app procura esses nomes.
2. **Nunca renomeie a coluna Título.** Em cada lista ela guarda uma coisa diferente (está indicado em cada tabela).
3. Crie a coluna com **+ Adicionar coluna > tipo > nome**. Para Sim/Não, Escolha e Data, ajuste as opções antes de salvar.
4. **Índices:** depois de criar as colunas, vá em **Configurações da lista > Colunas indexadas > Criar um novo índice** e indexe as colunas marcadas com **Indexar**. Sem isso, o app fica lento quando a lista passar de 5.000 itens.
5. Todas as listas no **mesmo site** (o mesmo do app atual).

## 1.2 Listas que já existem: acrescentar colunas
""")
    for L in exist:
        s.append(f"### `{L.nome}` (já existe)\n\n{L.descricao}\n\nCrie **somente as colunas que ainda não existem**. As da versão 1 continuam iguais.\n\n{tabela_lista(L)}\n")
    s.append("## 1.3 Listas novas\n\nPara cada uma: **+ Novo > Lista > Lista em branco**, digite o nome exato e crie as colunas.\n")
    for L in novas:
        if L.nome == "Acessos":
            continue
        s.append(f"### `{L.nome}`\n\n{L.descricao}\n\n**Quem edita:** {L.quem_edita}\n\n{tabela_lista(L)}\n")
    a = next(L for L in LISTAS if L.nome == "Acessos")
    s.append(f"""## 1.4 Acessos: é aqui que se define quem é quem

> **Esta lista é o único lugar onde você define o tipo de acesso de cada pessoa.**

| Perfil no app | Como a pessoa ganha | O que ela pode fazer |
|---|---|---|
| **Colaborador** | Automático: qualquer pessoa da empresa que **não** está nesta lista, ou que está com Ativo = Não | Consultar equipamentos, validade e histórico |
| **Laboratorista** | Linha nesta lista com **Perfil = Laboratorista** e Ativo = Sim | Consultar e **lançar aferições** |
| **Instrumentação** | Linha nesta lista com **Perfil = Instrumentação** e Ativo = Sim | Consultar e **registrar calibrações** |

Crie a lista **Acessos** com estas colunas:

{tabela_lista(a)}

**Para dar acesso a alguém:** abra a lista Acessos > **+ Novo** > em **Pessoa**, digite o nome e escolha no seletor > escolha o **Perfil** e a **Letra** > Salvar.
A aba **Acessos** da carga inicial tem os nomes e as letras da planilha, para você conferir.

**Para tirar o acesso:** desmarque **Ativo**. O histórico da pessoa continua.

**Permissão no SharePoint (automática):** o fluxo **F0 · Sincronizar acessos** (etapa 3) coloca e tira a pessoa do grupo certo do SharePoint sempre que você muda esta lista. Assim você mantém **um lugar só**. Sem o fluxo, adicione a pessoa também no grupo correspondente (seção 1.6).
""")
    lim = next(L for L in LISTAS if L.nome == "Limites")
    s.append("""## 1.5 Limites: é aqui que se define cada tolerância

> **Toda tolerância, limite e valor nominal de ensaio fica na lista `Limites`.** É a mesma tabela da aba BD_Limites da planilha.

| Situação | Preencha | O app considera Conforme quando |
|---|---|---|
| Faixa (ex.: rotação entre um mínimo e um máximo) | `LimInferior` e/ou `LimSuperior` (e `Nominal` para mostrar a diferença) | LimInferior ≤ valor ≤ LimSuperior |
| Diferença máxima (ex.: analisador x estufa) | `Tolerancia` (e `Nominal` quando a referência é fixa) | \\|valor − referência\\| ≤ Tolerância |
| Sem critério | nada | grava como **Informativo** |

**Para mudar uma tolerância:** abra a lista Limites > clique na linha do critério > altere `Tolerancia` > Salvar. **Vale a partir do próximo lançamento.** Cada aferição grava o critério que estava valendo naquele momento, então o histórico não muda.

**Qual ponto usa qual critério:** a coluna `ChaveLimite` da lista **PontosAfericao** liga cada campo do formulário a uma linha da Limites.

**Voltas dos tambores e diâmetros da granulometria** ficam na coluna `Fator` da lista **PontosAfericao**. **Massas de ensaio do tamboramento** ficam na lista **Configuracoes**.
""")
    perm = ["## 1.6 Permissões (segurança real)", "",
            "O app esconde os botões conforme o perfil, mas quem garante que ninguém grave por fora é a permissão do SharePoint.", "",
            "1. **Configurações do site > Permissões do site > Criar grupo**: crie os dois grupos abaixo, com permissão **Leitura** no site.",
            "2. No grupo **Visitantes** do site, inclua **Todos, exceto usuários externos**. Assim todos leem.",
            "3. Em cada lista da tabela: **Configurações da lista > Permissões para esta lista > Parar de herdar permissões > Conceder permissões**: grupo + **Colaborar**.", "",
            "| Lista | Colaborador (Visitantes) | Metrologia - Laboratoristas | Metrologia - Instrumentação | Você (Proprietários) |", "|---|---|---|---|---|"]
    for L in LISTAS:
        lab = "Colaborar" if L.nome in GRUPOS["Metrologia - Laboratoristas"] else "Leitura"
        ins = "Colaborar" if L.nome in GRUPOS["Metrologia - Instrumentação"] else "Leitura"
        perm.append(f"| {L.nome} | Leitura | {lab} | {ins} | Controle total |")
    perm += ["", "> **Os fluxos do Power Automate** devem ser criados por uma conta **Proprietária** do site (a sua). Eles atualizam Equipamentos e Status mesmo quando quem lançou é laboratorista.", ""]
    s.append("\n".join(perm))
    s.append("""## 1.7 Carregar os dados iniciais (caminho manual)

O arquivo **carga-inicial.xlsx** tem uma aba por lista, com as colunas na mesma ordem.

1. Abra a lista no SharePoint > **Editar no modo de grade**.
2. Confira se a ordem das colunas na grade é igual à da aba. Se não for, arraste as colunas da grade.
3. No Excel, selecione da linha 2 até a última (sem cabeçalho) e copie.
4. Clique na primeira célula vazia da grade e cole (Ctrl+V). Aguarde salvar e clique em **Sair do modo de grade**.

**Ordem:** Configuracoes > Limites > Ensaios > PontosAfericao > Equipamentos > Lancamentos > Afericoes (400 linhas: cole em 2 ou 3 partes) > Acessos (pelo seletor, um a um).

## 1.8 Cores nas listas

- **Equipamentos > coluna Status > Formatar esta coluna > Modo avançado**: cole `formatacao/equipamentos-status.json`.
- **Afericoes > coluna Resultado**: cole `formatacao/afericoes-resultado.json`.

## 1.9 Conferência final

- [ ] As 9 listas estão no mesmo site, com os nomes exatos
- [ ] Índices criados (Equipamentos.Title, Afericoes.Title/DataEnsaio/TAG/Ensaio, Lancamentos.Title/DataEnsaio/Ensaio, PontosAfericao.Ensaio, Calibracoes.EquipamentoID)
- [ ] Grupos criados e permissões aplicadas
- [ ] Configuracoes: `EmailsRelatorioDiario` e `LinkApp` preenchidos
- [ ] Limites: 13 critérios conferidos com o gestor
- [ ] Equipamentos: TAGs repetidas e datas de calibração conferidas com a instrumentação (veja `ObservacaoCadastro`)
- [ ] Equipamentos da versão 1 (os que já estavam na lista): `ControlaCalibracao` = Sim e `Ativo` = Sim
- [ ] Acessos: laboratoristas e instrumentação cadastrados pelo seletor
""")
    return "\n".join(s)


def dicionario() -> str:
    s = ["# Dicionário de dados", "", "Gerado de `ferramentas/modelo.py`. Uma tabela por lista do SharePoint.", ""]
    for L in LISTAS:
        s += [f"## {L.nome}" + (" (existente)" if L.existente else ""), "", L.descricao, "", f"**Quem edita:** {L.quem_edita}", "", tabela_lista(L), ""]
    return "\n".join(s)


if __name__ == "__main__":
    (RAIZ / "02-sharepoint" / "GUIA-SHAREPOINT.md").write_text(guia_sharepoint(), encoding="utf-8")
    (RAIZ / "01-arquitetura" / "DICIONARIO-DE-DADOS.md").write_text(dicionario(), encoding="utf-8")
    print("documentação gerada")
