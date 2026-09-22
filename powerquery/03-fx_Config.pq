// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_Config      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Lê um parâmetro de tbl_Config. As regras de negócio já estão na planilha
// (Meta Disponibilidade 95, Horas do Mês 744, Dias p/ Alerta 30 e 15) e hoje
// não são usadas em lugar nenhum — os "30 dias" estão fixos no código.
//
//   fx_Config("Dias p/ Alerta Vencimento Calibração")  ->  30

let
    fx = ( parametro as text, optional padrao as any ) as any =>
        let
            Tabela  = fx_Tabela( "tbl_Config" ),
            Filtro  = Table.SelectRows( Tabela, each Text.Trim( Text.From( [Parâmetro] ) ) = parametro ),
            Bruto   = if Table.RowCount( Filtro ) = 0 then null else Filtro{0}[Valor],
            Numero  = try Number.From( Bruto ) otherwise null,
            Escolha = if Numero <> null then Numero
                      else if Bruto <> null then Bruto
                      else padrao
        in
            Escolha
in
    fx
