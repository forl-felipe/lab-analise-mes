// Nova consulta: Previstos
// Um registro por ensaio previsto na agenda (lista Ensaios) em cada dia e turno, com a situação
// calculada pela MESMA regra da macro ModPainel da planilha (abas Painel e Gestão por Letras):
//   No turno      registro com a data e o turno previstos (mesmo que lançado depois)
//   No dia        registro na mesma data, no outro turno (quando o outro turno não é da rotina)
//   Com atraso    registro depois do turno previsto e antes da próxima data da rotina
//   Não feito     o turno terminou e não há registro até a próxima data da rotina
//   Em andamento / Futuro   não contam
// Conferido em Python contra o painel de outubro da planilha (91 previstos, 83 realizados).
let
    Config = fnLista("Configuracoes"),
    Cfg = (k as text) as text => Table.SelectRows(Config, each [Title] = k){0}[Valor],
    Inicio = Date.FromText(Cfg("InicioControle"), [Format = "yyyy-MM-dd"]),
    Referencia = Date.FromText(Cfg("EscalaDataReferencia"), [Format = "yyyy-MM-dd"]),
    Letras07 = Text.Split(Cfg("EscalaLetras07x19"), ","),
    Letras19 = Text.Split(Cfg("EscalaLetras19x07"), ","),
    Agora = DateTimeZone.RemoveZone(DateTimeZone.SwitchZone(DateTimeZone.UtcNow(), -3)),
    Hoje = Date.From(Agora),

    // ---- agenda: para cada ensaio, 14 marcações na ordem SegD, SegN, TerD ... DomN
    Ensaios = Table.SelectRows(fnLista("Ensaios"), each [Ativo] = true),
    Colunas = {"SegD", "SegN", "TerD", "TerN", "QuaD", "QuaN", "QuiD", "QuiN", "SexD", "SexN", "SabD", "SabN", "DomD", "DomN"},
    Agenda = Record.FromList(
        List.Transform(Table.ToRecords(Ensaios), (e) => List.Transform(Colunas, (c) => Record.Field(e, c) = true)),
        Ensaios[NomeBD]
    ),
    // índice de turno: dia (número da data) * 2 + 0 (07x19) ou 1 (19x07)
    PrevistoSlot = (ens as text, s as number) as logical =>
        let d = Date.From(Number.IntegerDivide(s, 2)), idx = Date.DayOfWeek(d, Day.Monday) * 2 + Number.Mod(s, 2)
        in Record.Field(Agenda, ens){idx},
    ProximoSlot = (ens as text, s as number) as number =>
        List.First(List.Select(List.Numbers(s + 1, 121), each PrevistoSlot(ens, _)), s + 121),

    // ---- registros realizados (a partir das aferições: um por ensaio, data e turno)
    Regs = Table.Distinct(Table.SelectColumns(tbl_Afericoes, {"Ensaio", "Data", "Turno"})),
    RegsSlot = Table.AddColumn(Regs, "Slot", each Number.From([Data]) * 2 + (if [Turno] = "19x07" then 1 else 0)),
    PorEnsaio = Table.Group(RegsSlot, {"Ensaio"}, {
        {"Slots", each List.Buffer([Slot])},
        {"Dias", each List.Buffer(List.Transform([Data], Number.From))}
    }),
    SlotsDe = (ens as text) as list => try PorEnsaio{[Ensaio = ens]}[Slots] otherwise {},
    DiasDe = (ens as text) as list => try PorEnsaio{[Ensaio = ens]}[Dias] otherwise {},

    // ---- grade de dias x turnos x ensaios previstos
    Dias = List.Dates(Inicio, Duration.Days(Hoje - Inicio) + 1, #duration(1, 0, 0, 0)),
    TabDias = Table.FromList(Dias, Splitter.SplitByNothing(), {"Data"}, null, ExtraValues.Error),
    ComTurnos = Table.ExpandListColumn(Table.AddColumn(TabDias, "Turno", each {"07x19", "19x07"}), "Turno"),
    ComEnsaio = Table.ExpandListColumn(Table.AddColumn(ComTurnos, "Ensaio", each Ensaios[NomeBD]), "Ensaio"),
    ComSlot = Table.AddColumn(ComEnsaio, "Slot", each Number.From([Data]) * 2 + (if [Turno] = "19x07" then 1 else 0)),
    SoPrevistos = Table.SelectRows(ComSlot, each PrevistoSlot([Ensaio], [Slot])),

    ComSituacao = Table.AddColumn(SoPrevistos, "Situação", each
        let
            ens = [Ensaio], s = [Slot], g = SlotsDe(ens), dia = Number.From([Data]),
            outro = dia * 2 + (1 - Number.Mod(s, 2))
        in
            if List.Contains(g, s) then "No turno"
            else if not PrevistoSlot(ens, outro) and List.Contains(DiasDe(ens), dia) then "No dia"
            else
                let prox = ProximoSlot(ens, s) in
                if List.AnyTrue(List.Transform(g, (r) => r > s and r < prox)) then "Com atraso"
                else
                    let
                        ini = DateTime.From([Data]) + (if Number.Mod(s, 2) = 0 then #duration(0, 7, 0, 0) else #duration(0, 19, 0, 0)),
                        fim = ini + #duration(0, 12, 0, 0)
                    in
                        if Agora < ini then "Futuro" else if Agora < fim then "Em andamento" else "Não feito",
        type text),
    ComLetra = Table.AddColumn(ComSituacao, "Letra", each
        let pos = Number.Mod(Duration.Days([Data] - Referencia), List.Count(Letras07)) in
        if [Turno] = "07x19" then Letras07{pos} else Letras19{pos}, type text),
    ComConta = Table.AddColumn(ComLetra, "Conta", each not List.Contains({"Futuro", "Em andamento"}, [Situação]), type logical),
    ComRealizado = Table.AddColumn(ComConta, "Realizado", each List.Contains({"No turno", "No dia", "Com atraso"}, [Situação]), type logical),
    NomeApp = Table.AddColumn(ComRealizado, "Ensaio (app)", each Table.SelectRows(Ensaios, (e) => e[NomeBD] = [Ensaio]){0}[Nome], type text),
    Final = Table.TransformColumnTypes(Table.RemoveColumns(NomeApp, {"Slot"}), {{"Data", type date}})
in
    Final
