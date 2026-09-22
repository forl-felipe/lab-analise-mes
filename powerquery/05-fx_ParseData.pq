// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÃO: fx_ParseData      ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Interpreta os SEIS formatos que convivem na coluna "Última Calibração":
//
//   serial do Excel   46174          ->  data direta
//   data nativa       (date/datetime)->  data direta
//   dd/mm/aaaa        01/06/2025     ->  2025-06-01
//   dd/mm/aa          01/07/26       ->  2026-07-01
//   mm/aa   (sem dia) 05/26          ->  maio/2026   (dia 1 ou último — ver flag)
//   mm/aaaa (sem dia) 05/2026        ->  maio/2026
//   texto livre       NOVA, SEM      ->  null
//
// Parâmetro fimDoMes:
//   false (padrão) -> mm/aa vira o dia 1.  Use para "Última Calibração":
//                     assumir o início do mês é conservador (vence antes).
//   true           -> mm/aa vira o último dia. Use para vencimentos.
//
// O código antigo fazia  try Date.From([UltimaCalibracao]) otherwise ...
// que silenciosamente devolvia null — ou pior, uma data errada — para os 61
// registros no formato mm/aa, sem deixar rastro.

let
    fx = ( valor as any, optional fimDoMes as nullable logical ) as nullable date =>
        let
            UsarFim = fimDoMes = true,

            EhData   = valor is date or valor is datetime or valor is datetimezone,
            EhNumero = valor is number,

            Texto = if valor = null then null else Text.Trim( Text.From( valor ) ),
            Vazio = Texto = null or Texto = "",

            // 1 · data nativa
            PorData = if EhData then try Date.From( valor ) otherwise null else null,

            // 2 · serial do Excel (base 1899-12-30)
            PorSerial =
                if EhNumero and valor > 1 and valor < 100000
                then try Date.From( valor ) otherwise null
                else null,

            // 3 · texto com barras
            Partes  = if Vazio then {} else Text.Split( Texto, "/" ),
            NPartes = List.Count( Partes ),
            Numeros = List.Transform( Partes, each try Number.FromText( Text.Trim( _ ) ) otherwise null ),
            TodosNumericos = NPartes >= 2 and not List.Contains( Numeros, null ),

            AnoCheio = ( a as number ) as number => if a < 100 then 2000 + a else a,

            // 3a · dd/mm/aaaa  ou  dd/mm/aa
            PorCompleta =
                if TodosNumericos and NPartes = 3 then
                    try #date( AnoCheio( Numeros{2} ), Numeros{1}, Numeros{0} ) otherwise null
                else null,

            // 3b · mm/aa  ou  mm/aaaa   (sem dia)
            PorMesAno =
                if TodosNumericos and NPartes = 2 then
                    let
                        Mes       = Numeros{0},
                        Ano       = AnoCheio( Numeros{1} ),
                        Primeiro  = try #date( Ano, Mes, 1 ) otherwise null
                    in
                        if Primeiro = null then null
                        else if UsarFim then Date.EndOfMonth( Primeiro )
                        else Primeiro
                else null,

            Resultado =
                if PorData     <> null then PorData
                else if PorSerial   <> null then PorSerial
                else if PorCompleta <> null then PorCompleta
                else if PorMesAno   <> null then PorMesAno
                else null
        in
            Resultado
in
    fx
