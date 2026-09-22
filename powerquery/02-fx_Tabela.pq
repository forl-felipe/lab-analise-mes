// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_Tabela      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Busca uma tabela nomeada dentro de Fonte_Base e já remove as linhas em branco
// que o Excel arrasta junto (o ListObject de tbl_Inspeções vai até a linha 201,
// mas só 5 têm conteúdo — as outras 195 entravam no modelo).

let
    fx = ( nomeTabela as text ) as table =>
        let
            Dados = Fonte_Base{[ Item = nomeTabela, Kind = "Table" ]}[Data],
            Colunas = Table.ColumnNames( Dados ),

            // mantém apenas linhas em que ao menos uma coluna tem conteúdo real
            SemVazias = Table.SelectRows(
                Dados,
                each List.AnyTrue(
                    List.Transform(
                        Colunas,
                        (c) =>
                            let v = Record.Field( _, c )
                            in  v <> null and Text.Trim( Text.From( v ) ) <> ""
                    )
                )
            )
        in
            SemVazias
in
    fx
