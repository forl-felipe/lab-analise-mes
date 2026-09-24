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

    NArquivos = List.Count( ListaArquivos ),
    Status =
        if Mensal = null then
            "Pasta não encontrada no SharePoint. Provável causa: ""laboratorios"" é um subsite. Troque Site e Caminho nas consultas tbl_Afericoes e tbl_Afericoes_Status (instrução no início delas)."
        else if NArquivos = 0 then
            "Pasta encontrada, mas sem planilhas .xlsm nas pastas ""Calibração mensal"" a partir de " & Text.From( AnoInicial ) & "."
        else "ok",
    Resultado = #table(
        type table [ Status = text, #"Planilhas na Pasta" = Int64.Type, #"Pastas de Ano" = Int64.Type ],
        { { Status, NArquivos, if PastasAno = null then 0 else Table.RowCount( PastasAno ) } } )
in
    Resultado
