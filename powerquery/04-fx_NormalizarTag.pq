// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_NormalizarTag      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// A coluna Tag de tbl_Calibracao mistura número (449212), texto ('66EF15'),
// texto com espaços irregulares ('66 AN 13', '66AN 06', '66BL 98') e 1 nulo.
// Normalizado: 205 valores distintos em 207 não-nulos — 2 duplicados reais
// ('66BL98' e '66PS01') que precisam ser resolvidos na origem.
//
// A Tag original é preservada para exibição; esta gera a CHAVE de relacionamento.

let
    fx = ( tag as any ) as nullable text =>
        let
            Texto      = if tag = null then null else Text.From( tag ),
            Limpo      = if Texto = null then null else Text.Upper( Text.Trim( Texto ) ),
            SemEspacos = if Limpo = null then null else Text.Remove( Limpo, { " ", Character.FromNumber(160) } ),
            Resultado  = if SemEspacos = null or SemEspacos = "" then null else SemEspacos
        in
            Resultado
in
    fx
