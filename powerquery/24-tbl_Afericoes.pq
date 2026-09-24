// ─────────────────────────────────────────────────────────────────────────────
// tbl_Afericoes — aferições, calibrações internas e comparativos (v15)
// ─────────────────────────────────────────────────────────────────────────────
// Cópia de LEITURA da partição de tbl_Afericoes
// (Painel.SemanticModel/definition/tables/tbl_Afericoes.tmdl).
//
// Fonte: a PASTA do SharePoint com as planilhas mensais "CALIBRAÇÃO MM - MÊS.xlsm".
// Cada uma tem a aba oculta BD_Afericoes com a tabela tbl_Afericoes. Todo mês
// novo que entra na pasta aparece sozinho, com o histórico acumulado.
//
// Enquanto PastaLink não for preenchido, a consulta devolve uma tabela vazia
// com as colunas certas. Assim o resto do painel atualiza normalmente.

let
    // ════════════════════════════════════════════════════════════════════════
    // ►► COLE AQUI o link da PASTA do SharePoint onde ficam os arquivos mensais
    //    "CALIBRAÇÃO MM - MÊS.xlsm" (o endereço do navegador com a pasta aberta,
    //    ou o link "Copiar link" da pasta). Enquanto estiver com o texto abaixo,
    //    as páginas de Aferições ficam vazias e o resto do painel funciona.
    // ════════════════════════════════════════════════════════════════════════
    PastaLink = "COLE_AQUI_O_LINK_DA_PASTA",

    Colunas = { "Ensaio", "Tipo", "Data", "Equipamento", "Parâmetro", "Unidade", "Valor",
                "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância",
                "Resultado", "Responsável", "Letra", "Observação", "Origem" },
    Numericas = { "Data", "Valor", "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância" },

    Configurado = Text.StartsWith( Text.Lower( Text.Trim( PastaLink ) ), "https://" ),

    // link do navegador ou de compartilhamento: tira "?..." e o prefixo "/:f:/r"
    SemQuery = let t = Text.Trim( PastaLink ) in if Text.Contains( t, "?" ) then Text.BeforeDelimiter( t, "?" ) else t,
    Limpo    = Text.Replace( Text.Replace( SemQuery, "/:f:/r/", "/" ), "/:f:/s/", "/" ),
    Partes   = Text.Split( Limpo, "/" ),
    // https: | "" | servidor | sites | <site>  ->  as 5 primeiras partes formam o site
    Site     = Text.Combine( List.FirstN( Partes, 5 ), "/" ),
    // compara sempre no formato codificado (%20, %C3%A7...), em minúsculas
    Codificar = ( t as text ) as text =>
        Text.Lower( Text.Combine( List.Transform( Text.Split( t, "/" ),
            each if Text.Contains( _, ":" ) or Text.Contains( _, "%" ) then _ else Uri.EscapeDataString( _ ) ), "/" ) ),
    Alvo     = Codificar( Limpo ),

    Lido =
        if not Configurado then null
        else
            let
                Arquivos  = SharePoint.Files( Site, [ ApiVersion = 15 ] ),
                DaPasta   = Table.SelectRows( Arquivos, each
                                Text.StartsWith( Codificar( [Folder Path] ), Alvo )
                                and Text.EndsWith( Text.Lower( [Name] ), ".xlsm" )
                                and not Text.StartsWith( [Name], "~$" ) ),
                // arquivo sem a tabela (mês antigo, cópia de trabalho) é ignorado, não quebra
                ComDados  = Table.AddColumn( DaPasta, "Dados", each
                                let
                                    wb = try Excel.Workbook( [Content], null, true ) otherwise null,
                                    t  = if wb = null then null
                                         else Table.SelectRows( wb, each [Kind] = "Table" and [Name] = "tbl_Afericoes" )
                                in
                                    if t = null or Table.IsEmpty( t ) then null else t{0}[Data] ),
                Validos   = Table.SelectRows( ComDados, each [Dados] <> null ),
                Enxuto    = Table.SelectColumns( Validos, { "Name", "Dados" } ),
                Renomeado = Table.RenameColumns( Enxuto, { { "Name", "Arquivo" } } ),
                Expandido = Table.ExpandTableColumn( Renomeado, "Dados", Colunas )
            in
                Expandido,

    Vazio = #table( List.Combine( { { "Arquivo" }, Colunas } ), {} ),
    Base  = if Lido = null then Vazio else Lido,

    // fórmula que devolve "" chega como texto vazio: vira null antes de tipar
    SemVazio = Table.ReplaceValue( Base, "", null, Replacer.ReplaceValue, Numericas ),
    Tipado   = Table.TransformColumnTypes( SemVazio, {
                   { "Data", type date }, { "Valor", type number }, { "Referência", type number },
                   { "Diferença", type number }, { "Lim. Inferior", type number },
                   { "Lim. Superior", type number }, { "Tolerância", type number },
                   { "Ensaio", type text }, { "Tipo", type text }, { "Equipamento", type text },
                   { "Parâmetro", type text }, { "Unidade", type text }, { "Resultado", type text },
                   { "Responsável", type text }, { "Letra", type text }, { "Observação", type text },
                   { "Origem", type text }, { "Arquivo", type text } } ),
    SemErro  = Table.ReplaceErrorValues( Tipado, List.Transform( Numericas, each { _, null } ) ),

    // só o que foi de fato registrado: tem data e tem resultado
    Registros = Table.SelectRows( SemErro, each [Data] <> null and [Resultado] <> null and [Resultado] <> "" ),

    // o mesmo lançamento não entra duas vezes (arquivo copiado para o mês seguinte sem
    // limpar): Origem é a célula da aba de preenchimento, única dentro do arquivo
    Chave  = Table.AddColumn( Registros, "Chave", each
                 [Origem] & "|" & [Parâmetro] & "|" & Date.ToText( [Data], "yyyy-MM-dd" ), type text ),
    Unicos = Table.Distinct( Chave, { "Chave" } ),

    Conforme = Table.AddColumn( Unicos, "Conforme",
                   each if [Resultado] = "Não conforme" then 0 else if [Resultado] = "Conforme" then 1 else null,
                   Int64.Type ),

    // TAG normalizada: "66AN11 x Estufa" -> 66AN11 · "66TA05 (5 kg) ..." -> 66TA05 · "PN  497" -> PN497
    TagKey = Table.AddColumn( Conforme, "TagKey", each
                 let
                     e  = if [Equipamento] = null then "" else [Equipamento],
                     b1 = if Text.Contains( e, " x " ) then Text.BeforeDelimiter( e, " x " ) else e,
                     b2 = if Text.Contains( b1, " (" ) then Text.BeforeDelimiter( b1, " (" ) else b1
                 in
                     Text.Upper( Text.Remove( b2, { " ", "-", Character.FromNumber( 160 ) } ) ), type text ),

    // família de instrumento: liga a aferição à "saúde por grupo" da Visão Geral
    Grupo = Table.AddColumn( TagKey, "Grupo Operacional", each
                if      Text.StartsWith( [TagKey], "66TA" ) and [Ensaio] = "Tambor de Abrasão" then "Tambores de Abrasão"
                else if Text.StartsWith( [TagKey], "66AG" ) then "Peneiradores Alpine"
                else if Text.StartsWith( [TagKey], "66PS" ) then "Prensas de Compressão"
                else if Text.StartsWith( [TagKey], "66AN" ) then "Analisadores de Umidade"
                else null, type nullable text ),

    PctTol = Table.AddColumn( Grupo, "% da Tolerância", each
                 if [Tolerância] <> null and [Tolerância] > 0 and [Diferença] <> null
                 then Number.Abs( [Diferença] ) / [Tolerância] else null, type nullable number ),

    Comparativo = Table.AddColumn( PctTol, "Comparativo", each
                      [Equipamento] & " · " & [Parâmetro], type text ),

    Semana = Table.AddColumn( Comparativo, "Semana Início", each Date.StartOfWeek( [Data], Day.Monday ), type date )
in
    Semana

