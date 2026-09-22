// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_QualidadeData      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Classifica a CONFIABILIDADE do valor original, para que o painel possa
// mostrar quantos registros dependem de uma interpretação e dirigir a limpeza
// na origem — em vez de esconder o problema atrás de um try/otherwise.

let
    fx = ( valor as any ) as text =>
        let
            Texto = if valor = null then null else Text.Trim( Text.From( valor ) ),
            Vazio = Texto = null or Texto = "",

            EhData   = valor is date or valor is datetime or valor is datetimezone,
            EhNumero = valor is number,

            Partes  = if Vazio then {} else Text.Split( Texto, "/" ),
            NPartes = List.Count( Partes ),
            Numeros = List.Transform( Partes, each try Number.FromText( Text.Trim( _ ) ) otherwise null ),
            TodosNumericos = NPartes >= 2 and not List.Contains( Numeros, null ),

            Resultado =
                if Vazio                              then "Vazio"
                else if EhData or EhNumero            then "OK"
                else if TodosNumericos and NPartes = 3 then "OK"
                else if TodosNumericos and NPartes = 2 then "Mês sem dia"
                else "Não interpretável"
        in
            Resultado
in
    fx
