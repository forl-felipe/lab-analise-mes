// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_Tabela      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Busca uma tabela nomeada dentro de Fonte_Base, normaliza texto só-espaço para
// null (o SharePoint manda U+00A0 em célula "vazia") e remove as linhas em branco
// que o Excel arrasta junto (o ListObject de tbl_Inspeções vai até a linha 201,
// mas só 5 têm conteúdo — as outras 195 entravam no modelo).

let
    fx = ( nomeTabela as text ) as table =>
        let
            Dados = Fonte_Base{[ Item = nomeTabela, Kind = "Table" ]}[Data],
            Colunas = Table.ColumnNames( Dados ),

            // U+00A0 conta como conteúdo se não for tratado: a linha entra no modelo
            // e o valor vira uma categoria em branco no gráfico
            Branco = { " ", "#(00A0)", "#(tab)", "#(cr)", "#(lf)" },
            Limpo = Table.TransformColumns(
                Dados, {},
                ( v ) => if v is text
                        then ( let t = Text.Trim( v, Branco ) in if t = "" then null else t )
                        else v
            ),

            // mantém apenas linhas em que ao menos uma coluna tem conteúdo real
            SemVazias = Table.SelectRows(
                Limpo,
                each List.AnyTrue(
                    List.Transform( Colunas, (c) => Record.Field( _, c ) <> null )
                )
            )
        in
            SemVazias
in
    fx
