let
    // ════════════════════════════════════════════════════════════════════════
    // Planilhas mensais de calibração/aferição no SharePoint:
    //   automacao / laboratorios / Laboratório de Controle da Produção - Ubu /
    //   Resultados rotina / Laboratorio Fisico Ubu / Calibração Integrada  -LCP Ubu /
    //   Calibração Integrada mensal / Calibração mensal AAAA / CALIBRAÇÃO MM - MÊS.xlsm
    //
    // Endereço FIXO (o serviço do Power BI só agenda atualização com endereço fixo).
    // São lidas todas as pastas "Calibração mensal AAAA" a partir de AnoInicial:
    // o ano novo entra sozinho. Atenção ao espaço DUPLO em "Integrada  -LCP".
    //
    // Se a consulta tbl_Afericoes_Status disser "Pasta não encontrada",
    // "laboratorios" é um subsite e não uma biblioteca. Nesse caso:
    //   Site    = "https://smineracao.sharepoint.com/sites/automacao/laboratorios"
    //   Caminho = a mesma lista SEM o primeiro item "laboratorios"
    // (trocar nas DUAS consultas: tbl_Afericoes e tbl_Afericoes_Status)
    // ════════════════════════════════════════════════════════════════════════
    Site = "https://smineracao.sharepoint.com/sites/automacao",
    Caminho = { "laboratorios", "Laboratório de Controle da Produção - Ubu", "Resultados rotina",
                "Laboratorio Fisico Ubu", "Calibração Integrada  -LCP Ubu", "Calibração Integrada mensal" },
    AnoInicial = 2026,

    // navega pasta a pasta (lista só o caminho, não o site inteiro)
    Mensal = try List.Accumulate( Caminho, SharePoint.Contents( Site, [ ApiVersion = 15 ] ),
                 ( t, n ) => t{[Name = n]}[Content] ) otherwise null,
    PastasAno =
        if Mensal = null then null
        else Table.SelectRows( Mensal, each
                 Value.Is( [Content], type table )
                 and Text.StartsWith( Text.Lower( Text.Trim( [Name] ) ), "calibração mensal" )
                 and ( try Number.From( Text.End( Text.Trim( [Name] ), 4 ) ) >= AnoInicial otherwise false ) ),
    ListaArquivos =
        if PastasAno = null then {}
        else List.Combine( List.Transform( PastasAno[Content], each
                 Table.ToRecords( Table.SelectColumns(
                     Table.SelectRows( _, each
                         not Value.Is( [Content], type table )
                         and Text.EndsWith( Text.Lower( [Name] ), ".xlsm" )
                         and not Text.StartsWith( [Name], "~$" ) ),
                     { "Name", "Content" } ) ) ) ),
    Arquivos = Table.FromRecords( ListaArquivos, { "Name", "Content" } ),

    Colunas = { "Ensaio", "Tipo", "Data", "Equipamento", "Parâmetro", "Unidade", "Valor",
                "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância",
                "Resultado", "Responsável", "Letra", "Observação", "Origem" },
    Numericas = { "Data", "Valor", "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância" },

    // arquivo sem a aba oculta (o dos operadores antes da troca, cópia antiga) é ignorado
    ComDados  = Table.AddColumn( Arquivos, "Dados", each
                    let
                        wb = try Excel.Workbook( [Content], null, true ) otherwise null,
                        t  = if wb = null then null
                             else Table.SelectRows( wb, each [Kind] = "Table" and [Name] = "tbl_Afericoes" )
                    in
                        if t = null or Table.IsEmpty( t ) then null else t{0}[Data] ),
    Validos   = Table.SelectRows( ComDados, each [Dados] <> null ),
    Renomeado = Table.RenameColumns( Table.SelectColumns( Validos, { "Name", "Dados" } ), { { "Name", "Arquivo" } } ),
    Lido      = if Table.IsEmpty( Renomeado ) then null else Table.ExpandTableColumn( Renomeado, "Dados", Colunas ),

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
