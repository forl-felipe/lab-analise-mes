// ─────────────────────────────────────────────────────────────────────────────
// BLOCO: leitura da planilha de respostas do Microsoft Forms
// ─────────────────────────────────────────────────────────────────────────────
// Vive DENTRO das partições de tbl_Inspecoes e tbl_Medicoes_Operacionais (cada
// partição do modelo é autocontida). Reproduzido aqui só para leitura.
//
// Uma submissão carrega DOIS fatos:
//   a inspeção (checklist) -> tbl_Inspecoes
//   a medição  (tempos)    -> tbl_Medicoes_Operacionais
// O mesmo bloco alimenta as duas, com projeções diferentes no fim.

// ═══ Microsoft Forms ══════════════════════════════════════════════════════
    // A planilha de respostas nomeia cada coluna com o TEXTO da pergunta e
    // acrescenta sufixo numerico quando duas perguntas tem o mesmo enunciado
    // (Tempo 1 / Tempo 11). Os nomes ainda chegam com espaco duplicado e espaco
    // nao separavel. Por isso NADA aqui referencia nome exato de coluna: elas
    // sao achadas por padrao sobre o nome normalizado. Assim, mexer no
    // formulario nao quebra a consulta em silencio.
    // URL inteira do arquivo, na mesma forma da planilha antiga (que ja autentica).
    // Com pasta + RelativePath o Power BI testava a credencial contra a PASTA e recusava.
    FonteForms = Excel.Workbook(
        Web.Contents( "https://smineracao.sharepoint.com/sites/automacao/laboratorios/ZTESTE/Inspe%C3%A7%C3%A3o%20de%20Amostradores.xlsx" ),
        null, true ),
    // o arquivo de respostas tem uma tabela so: pega a primeira, qualquer que seja o nome
    // (tbl_forms, tbl_inspecoes... renomear no SharePoint nao quebra mais a consulta)
    BrutoForms = Table.SelectRows( FonteForms, each [Kind] = "Table" ){0}[Data],

    BrancoF = { " ", "#(00A0)", "#(tab)", "#(cr)", "#(lf)" },
    TextoF = ( v as any ) as nullable text =>
        let t = if v = null then null else Text.Trim( Text.From( v ), BrancoF )
        in  if t = null or t = "" then null else t,
    NumF = ( v as any ) as nullable number =>
        let t = TextoF( v )
        in  if t = null then null
            else try Number.FromText( Text.Replace( t, ",", "." ), "en-US" ) otherwise null,

    AchataF = ( n as text ) as text =>
        Text.Lower( Text.Remove( n, { " ", "#(00A0)", ".", "?", "(", ")", ",", ":", "-", "–", "/", "=", "·" } ) ),
    ColsF = Table.ColumnNames( BrutoForms ),
    ContemF = ( trecho as text ) as list => List.Select( ColsF, each Text.Contains( AchataF( _ ), trecho ) ),
    PrimeiraF = ( trecho as text ) as nullable text =>
        let l = ContemF( trecho ) in if List.Count( l ) = 0 then null else List.First( l ),
    ExataF = ( nome as text ) as nullable text =>
        let l = List.Select( ColsF, each AchataF( _ ) = nome )
        in  if List.Count( l ) = 0 then null else List.Last( l ),

    cId     = ExataF( "id" ),
    cFim    = ExataF( "horadeconclusão" ),
    cTurno  = ExataF( "turno" ),
    // ha DUAS colunas "Nome": a identidade de quem enviou (do Forms) e a
    // pergunta com o nome do inspetor. List.Last pega a pergunta, que vem depois.
    cNome   = ExataF( "nome" ),
    cCritic = ExataF( "criticidade" ),
    cTag    = PrimeiraF( "tagdoequipamento" ),
    colsDist = ContemF( "distânciadepercurso" ),   // uma por ramificação do formulário; só a escolhida vem preenchida
    cFoto   = PrimeiraF( "fotodoequipamento" ),
    cDescNC = PrimeiraF( "descrevaanãoconformidade" ),
    colsTempo = List.Select( ColsF, each Text.StartsWith( AchataF( _ ), "tempo" ) ),
    colsChk   = List.Select( ColsF, each Text.Contains( AchataF( _ ), "checklistdocortador" )
                                   or Text.Contains( AchataF( _ ), "divisorrotativo" ) ),
    colsObs   = ContemF( "observações" ),
    colsReq   = ContemF( "encontroualguma" ),

    CampoF = ( r as record, c as nullable text ) as any =>
        if c = null then null else Record.FieldOrDefault( r, c, null ),
    JuntaF = ( r as record, cols as list ) as nullable text =>
        let v = List.Select( List.Transform( cols, ( c ) => TextoF( CampoF( r, c ) ) ), each _ <> null )
        in  if List.Count( v ) = 0 then null else Text.Combine( v, " | " ),

    LinhasF = Table.SelectRows( BrutoForms, each CampoF( _, cFim ) <> null ),
    RegistrosF = List.Transform( Table.ToRecords( LinhasF ), ( r ) =>
        let
            fim    = CampoF( r, cFim ),
            data   = if fim = null then null else Date.From( fim ),
            bruto  = TextoF( CampoF( r, cTag ) ),
            tag    = if bruto = null then null else TextoF( Text.BeforeDelimiter( bruto & "·", "·" ) ),
            equip  = if bruto = null then null else TextoF( Text.AfterDelimiter( bruto, "·" ) ),
            // so o bloco do tipo de equipamento escolhido vem preenchido;
            // os outros ficam nulos, entao basta recolher os que tem valor
            tempos = List.Select( List.Transform( colsTempo, ( c ) => NumF( CampoF( r, c ) ) ), each _ <> null ),
            medio  = if List.Count( tempos ) = 0 then null else List.Average( tempos ),
            checks = List.Select( List.Transform( colsChk, ( c ) => TextoF( CampoF( r, c ) ) ), each _ <> null ),
            qtdNC  = List.Count( List.Select( checks, each Text.Contains( Text.Lower( _ ), "não conforme" ) ) ),
            // a distancia e digitada em CENTIMETROS (ver LEIA-ME: o enunciado da
            // pergunta usa metros com virgula e as respostas chegam sem a virgula)
            distCm = List.First( List.Select( List.Transform( colsDist, ( c ) => NumF( CampoF( r, c ) ) ), each _ <> null ), null ),
            veloc  = if medio = null or medio = 0 or distCm = null then null
                     else Number.Round( distCm / 100 / medio, 3 ),
            turnoF = let s = TextoF( CampoF( r, cTurno ) ) in
                     if s = null then null
                     else if Text.Upper( s ) = "ADM" then "Adm"
                     else if Text.StartsWith( Text.Lower( s ), "letra " )
                          then "Turno " & Text.Upper( Text.End( s, 1 ) )
                     else if Text.StartsWith( Text.Lower( s ), "turno " )
                          then "Turno " & Text.Upper( Text.End( s, 1 ) )
                     else s,
            idIns  = "FORM-" & ( if data = null then "" else Date.ToText( data, "yyyyMMdd" ) )
                     & "-" & ( if tag = null then "SEMTAG" else tag )
                     & "-" & Text.PadStart( Text.From( CampoF( r, cId ) ), 4, "0" )
        in
            [ ID_Inspecao = idIns, Carimbo = fim, Data = data, Tag = tag, Equipamento = equip,
              Turno = turnoF, Responsavel = TextoF( CampoF( r, cNome ) ),
              Tempos = tempos, TempoMedio = medio, DistanciaCm = distCm, Velocidade = veloc,
              QtdNC = qtdNC, Criticidade = TextoF( CampoF( r, cCritic ) ),
              Observacoes = JuntaF( r, colsObs & { cDescNC } ),
              RequerAcao = JuntaF( r, colsReq ),
              Foto = TextoF( CampoF( r, cFoto ) ),
              Checks = checks ] ),
    // #table com lista de nomes e lista de linhas: a forma mais simples de
    // montar tabela em M, sem declaracao de tipo embutida
    DoForms =
        Table.TransformColumnTypes(
            #table(
                { "ID_Inspecao", "Carimbo_DataHora", "Data", "Semana", "Mes", "Ano", "Turno", "Responsavel", "Equip_Tag", "Equipamento", "Qtd_NC", "Resultado_Inspecao", "Requer_Acao", "Criticidade", "Observacoes", "Foto", "TagKey Inspecao", "ID Válido", "Origem" },
                List.Transform( RegistrosF, ( x ) =>
                    { x[ID_Inspecao],
                      if x[Carimbo] = null then null else DateTime.ToText( x[Carimbo], "dd/MM/yyyy HH:mm" ),
                      x[Data],
                      if x[Data] = null then null else Text.From( Date.Year( x[Data] ) ) & "-S" & Text.PadStart( Text.From( Date.WeekOfYear( x[Data], Day.Monday ) ), 2, "0" ),
                      if x[Data] = null then null else Text.From( Date.Month( x[Data] ) ),
                      if x[Data] = null then null else Text.From( Date.Year( x[Data] ) ),
                      x[Turno],
                      x[Responsavel],
                      x[Tag],
                      x[Equipamento],
                      x[QtdNC],
                      if List.Count( x[Checks] ) = 0 then null else if x[QtdNC] = 0 then "Conforme" else "Nao Conforme",
                      x[RequerAcao],
                      x[Criticidade],
                      x[Observacoes],
                      x[Foto],
                      fxTag( x[Tag] ),
                      true,
                      "Forms"
                    } )
            ),
            { { "Data", type date }, { "Qtd_NC", Int64.Type }, { "ID Válido", type logical } }
        ),
