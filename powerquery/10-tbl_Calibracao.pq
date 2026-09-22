// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Calibracao          (grão: 1 linha por instrumento · 208 linhas)
// ─────────────────────────────────────────────────────────────────────────────
// Reescrita completa. Mudanças em relação à versão atual:
//   1. Usa Fonte_Base (1 download em vez de 12)
//   2. Parser explícito de 6 formatos de data  -> 192/208 interpretadas
//   3. Colunas de QUALIDADE para dirigir a limpeza na origem
//   4. "Hoje" em horário de Brasília, não UTC
//   5. Limiares lidos de tbl_Config (30 / 15 dias), não fixos no código
//   6. Guarda de nulo em CriticidadeCalibracao  (o if antigo quebrava com null)
//   7. STATUS METROLÓGICO promovido a coluna própria — 48 instrumentos
//      reprovados estavam invisíveis no painel

let
    Origem = fx_Tabela( "tbl_Calibracao" ),

    // ── limiares vindos da planilha, não do código ───────────────────────────
    DiasAlerta  = fx_Config( "Dias p/ Alerta Vencimento Calibração", 30 ),
    DiasCritico = fx_Config( "Dias p/ Alerta Vencimento Crítico",    15 ),
    Hoje        = Date.From( DateTimeZone.SwitchZone( DateTimeZone.UtcNow(), -3 ) ),

    Renomear = Table.RenameColumns(
        Origem,
        {
            { "Laboratório",        "Laboratorio"            },
            { "Última Calibração",  "UltimaCalibracaoOrigem" },
            { "Próximo Vencimento", "ProxVencimentoOrigem"   }
        },
        MissingField.Ignore
    ),

    Tipos = Table.TransformColumnTypes(
        Renomear,
        {
            { "Equipamento",        type text },
            { "Laboratorio",        type text },
            { "Local",              type text },
            { "Frequência (meses)", Int64.Type },
            { "Status Calibração",  type text },
            { "Certificado",        type text },
            { "Observações",        type text }
        },
        "pt-BR"
    ),

    // ── chave de relacionamento normalizada ─────────────────────────────────
    ComTagKey = Table.AddColumn( Tipos, "TagKey", each fx_NormalizarTag( [Tag] ), type nullable text ),
    TagTexto  = Table.TransformColumns( ComTagKey, {{ "Tag", each if _ = null then null else Text.Trim( Text.From( _ ) ), type nullable text }} ),

    // ── datas interpretadas + qualidade ─────────────────────────────────────
    UltimaCalib = Table.AddColumn( TagTexto,     "UltimaCalibData",     each fx_ParseData( [UltimaCalibracaoOrigem], false ), type nullable date ),
    Vencimento  = Table.AddColumn( UltimaCalib,  "DataVencimento",      each fx_ParseData( [ProxVencimentoOrigem],  true  ), type nullable date ),
    QualUltima  = Table.AddColumn( Vencimento,   "Qualidade Última Calibração", each fx_QualidadeData( [UltimaCalibracaoOrigem] ), type text ),
    QualVenc    = Table.AddColumn( QualUltima,   "Qualidade Vencimento",        each fx_QualidadeData( [ProxVencimentoOrigem]   ), type text ),

    // sinaliza o absurdo lógico: calibração "realizada" com data futura
    DataFutura = Table.AddColumn(
        QualVenc, "Última Calibração no Futuro",
        each [UltimaCalibData] <> null and [UltimaCalibData] > Hoje,
        type logical
    ),

    // ── prazo ───────────────────────────────────────────────────────────────
    Dias = Table.AddColumn(
        DataFutura, "DiasParaVencer",
        each if [DataVencimento] = null then null else Duration.Days( [DataVencimento] - Hoje ),
        Int64.Type
    ),

    StatusAtual = Table.AddColumn(
        Dias, "StatusAtual",
        each
            if [DiasParaVencer] = null    then "Sem Data"
            else if [DiasParaVencer] < 0  then "Vencido"
            else if [DiasParaVencer] <= DiasAlerta then "A Vencer"
            else "Em Dia",
        type text
    ),

    // ordem explícita para os visuais (evita ordenação alfabética)
    OrdemStatus = Table.AddColumn(
        StatusAtual, "Ordem Status",
        each if [StatusAtual] = "Vencido" then 1
             else if [StatusAtual] = "A Vencer" then 2
             else if [StatusAtual] = "Em Dia" then 3
             else 4,
        Int64.Type
    ),

    // faixa de criticidade — agora com guarda de nulo (o if antigo produzia
    // 22 valores em branco porque  null < 0  não é false em M, é null)
    Criticidade = Table.AddColumn(
        OrdemStatus, "CriticidadeCalibracao",
        each
            if [DiasParaVencer] = null                then "Sem Data"
            else if [DiasParaVencer] < -180           then "Vencido há +180d"
            else if [DiasParaVencer] < 0              then "Vencido"
            else if [DiasParaVencer] <= DiasCritico   then "Crítico"
            else if [DiasParaVencer] <= DiasAlerta    then "Atenção"
            else "OK",
        type text
    ),

    // ── status METROLÓGICO (o que o painel atual ignora por completo) ───────
    // origem: OK = 137 · N/OK = 47 · E/C = 23 · NOK = 1 (erro de digitação)
    StatusMetro = Table.AddColumn(
        Criticidade, "Status Metrologico",
        each
            let Bruto = if [Status Calibração] = null then "" else Text.Upper( Text.Remove( Text.Trim( [Status Calibração] ), { "/", " ", "-" } ) )
            in  if Bruto = "OK"    then "Aprovado"
                else if Bruto = "NOK" then "Reprovado"
                else if Bruto = "EC"  then "Em calibração"
                else "Não informado",
        type text
    ),

    OrdemMetro = Table.AddColumn(
        StatusMetro, "Ordem Metrologico",
        each if [Status Metrologico] = "Reprovado" then 1
             else if [Status Metrologico] = "Em calibração" then 2
             else if [Status Metrologico] = "Aprovado" then 3
             else 4,
        Int64.Type
    ),

    // instrumento em risco = reprovado OU fora do prazo.
    // Um instrumento "Em Dia" de prazo e REPROVADO está em uso medindo errado.
    EmRisco = Table.AddColumn(
        OrdemMetro, "Em Risco",
        each [Status Metrologico] = "Reprovado" or [StatusAtual] = "Vencido",
        type logical
    ),

    // ── tipo de instrumento, derivado da descrição ──────────────────────────
    // Peneira 120 · Balança 38 · Estufa 11 · Prensa 7 · Analisador 5 · Tambor 5
    TipoInstrumento = Table.AddColumn(
        EmRisco, "Tipo Instrumento",
        each
            let D = if [Equipamento] = null then "" else Text.Upper( Text.From( [Equipamento] ) )
            in  if      Text.Contains( D, "PENEIRA"        ) then "Peneira"
                else if Text.Contains( D, "BALAN"          ) then "Balança"
                else if Text.Contains( D, "ESTUFA"         ) then "Estufa"
                else if Text.Contains( D, "PRENSA"         ) then "Prensa"
                else if Text.Contains( D, "ANALISADOR"     ) then "Analisador"
                else if Text.Contains( D, "TAMBOR"         ) then "Tambor"
                else if Text.Contains( D, "PERMEAB"        ) then "Permeabilímetro"
                else if Text.Contains( D, "TITULA"         ) then "Titulador"
                else if Text.Contains( D, "FORNO"          ) then "Forno"
                else if Text.Contains( D, "DINAM"          ) then "Dinamômetro"
                else if Text.Contains( D, "PESAGEM"        ) then "Pesagem"
                else "Outros",
        type text
    ),

    // laboratório nulo (6 linhas) vira "N/A" para casar com DimLaboratorio
    LabTratado = Table.TransformColumns(
        TipoInstrumento,
        {{ "Laboratorio", each if _ = null or Text.Trim( Text.From( _ ) ) = "" then "N/A" else Text.Upper( Text.Trim( Text.From( _ ) ) ), type text }}
    )
in
    LabTratado
