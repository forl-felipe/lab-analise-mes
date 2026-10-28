"""
Gera as telas do app Metrologia (Power Apps) em YAML, prontas para colar no
Power Apps Studio.

Uso:  python gerar_telas.py
Saída: 03-powerapps/telas/*.yaml  e  03-powerapps/alternativa-somente-controles/*.yaml

As fórmulas estão no formato invariável (vírgula como separador), que é o
formato do YAML do Power Apps: o Studio converte sozinho para português.
Controles usados: os mesmos da versão 1 (clássicos), que já foram colados com
sucesso no ambiente.
"""

import textwrap
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "03-powerapps" / "telas"
DESTINO_CTRL = RAIZ / "03-powerapps" / "alternativa-somente-controles"

# ------------------------------------------------------------------ paleta
NAVY = "RGBA(11, 37, 64, 1)"
NAVY2 = "RGBA(22, 59, 98, 1)"
ACCENT = "RGBA(14, 107, 168, 1)"
BG = "RGBA(242, 245, 248, 1)"
WHITE = "RGBA(255, 255, 255, 1)"
TXT = "RGBA(23, 35, 52, 1)"
SUB = "RGBA(93, 107, 125, 1)"
LINE = "RGBA(225, 231, 238, 1)"
HEAD_SUB = "RGBA(169, 194, 220, 1)"
OK, OK_S = "RGBA(30, 127, 79, 1)", "RGBA(226, 243, 234, 1)"
BAD, BAD_S = "RGBA(180, 35, 24, 1)", "RGBA(253, 236, 235, 1)"
WARN, WARN_S = "RGBA(181, 71, 8, 1)", "RGBA(254, 240, 199, 1)"
MUT, MUT_S = "RGBA(85, 98, 116, 1)", "RGBA(236, 239, 243, 1)"
LAB, LAB_S = "RGBA(23, 92, 211, 1)", "RGBA(230, 239, 255, 1)"
INS, INS_S = "RGBA(107, 47, 191, 1)", "RGBA(241, 232, 255, 1)"
OVERLAY = "RGBA(6, 14, 24, 0.62)"


def F(s: str) -> str:
    """Normaliza uma fórmula multilinha e acrescenta o '='."""
    return "=" + textwrap.dedent(s).strip("\n").rstrip()


def status_fill(expr):
    return F(f'''
        If(
            StartsWith({expr}, "Bloqueado"), {BAD_S},
            {expr} = "Liberado · Calibração vence em breve", {WARN_S},
            {expr} = "Liberado · Sem data de calibração", {MUT_S},
            {OK_S}
        )''')


def status_color(expr):
    return F(f'''
        If(
            StartsWith({expr}, "Bloqueado"), {BAD},
            {expr} = "Liberado · Calibração vence em breve", {WARN},
            {expr} = "Liberado · Sem data de calibração", {MUT},
            {OK}
        )''')


def res_fill(expr):
    return F(f'If({expr} = "Conforme", {OK_S}, {expr} = "Não conforme", {BAD_S}, {MUT_S})')


def res_color(expr):
    return F(f'If({expr} = "Conforme", {OK}, {expr} = "Não conforme", {BAD}, {MUT})')


# ------------------------------------------------------------------ fábrica de controles
def _props(**kw):
    out = {}
    for k, v in kw.items():
        if v is None:
            continue
        if isinstance(v, (int, float)):
            v = f"={v}"
        elif isinstance(v, bool):
            v = "=true" if v else "=false"
        elif not v.startswith("="):
            v = "=" + v
        out[k] = v
    return out


def ctl(nome, tipo, variant=None, children=None, **props):
    d = {"Control": tipo}
    if variant:
        d["Variant"] = variant
    d["Properties"] = _props(**props)
    if children:
        d["Children"] = children
    return {nome: d}


def label(nome, text, x, y, w, h, size=13, bold=False, color=TXT, **kw):
    color = kw.pop("Color", color)
    peso = kw.pop("FontWeight", "FontWeight.Bold" if bold else None)
    return ctl(nome, "Label@2.5.1", Text=text, X=x, Y=y, Width=w, Height=h, Size=size,
               FontWeight=peso, Color=color, **kw)


def rect(nome, x, y, w, h, fill=WHITE, **kw):
    return ctl(nome, "Rectangle@2.3.0", Fill=fill, X=x, Y=y, Width=w, Height=h, **kw)


def radius(r):
    return dict(RadiusTopLeft=r, RadiusTopRight=r, RadiusBottomLeft=r, RadiusBottomRight=r)


def button(nome, text, x, y, w, h, onselect, fill=NAVY, color=WHITE, size=15, **kw):
    base = dict(Text=text, X=x, Y=y, Width=w, Height=h, OnSelect=onselect, Fill=fill, Color=color, Size=size,
                FontWeight="FontWeight.Bold", HoverFill=kw.pop("HoverFill", fill), PressedFill=kw.pop("PressedFill", fill),
                HoverColor=kw.pop("HoverColor", color), PressedColor=kw.pop("PressedColor", color), **radius(12))
    base.update(kw)
    return ctl(nome, "Classic/Button@2.2.0", **base)


def textinput(nome, x, y, w, h, hint, default='""', size=15, **kw):
    return ctl(nome, "Classic/TextInput@2.3.2", Default=default, HintText=hint, X=x, Y=y, Width=w, Height=h, Size=size,
               BorderColor="RGBA(205, 214, 225, 1)", **radius(10), **kw)


def dropdown(nome, items, x, y, w, h, default=None, size=14, **kw):
    return ctl(nome, "Classic/DropDown@2.3.1", Items=items, Default=default, X=x, Y=y, Width=w, Height=h, Size=size,
               BorderColor="RGBA(205, 214, 225, 1)", **kw)


def icon(nome, ic, x, y, w, h, onselect=None, color=WHITE, **kw):
    return ctl(nome, "Classic/Icon@2.5.0", Icon=ic, X=x, Y=y, Width=w, Height=h, Color=color, OnSelect=onselect,
               PaddingTop=8, PaddingBottom=8, PaddingLeft=8, PaddingRight=8, **kw)


def gallery(nome, items, x, y, w, h, template, children, variant_override=None, **kw):
    return ctl(nome, "Gallery@2.15.0", variant=variant_override or "Vertical", children=children, Items=items, X=x, Y=y, Width=w, Height=h,
               TemplateSize=template, TemplatePadding=0, TemplateFill="RGBA(0, 0, 0, 0)", **kw)


def cabecalho(prefixo, titulo, subtitulo, voltar=None, direita=None):
    """Faixa azul do topo. voltar = fórmula do botão voltar (None = logo)."""
    c = [rect(f"rec{prefixo}Topo", 0, 0, "Parent.Width", 112, fill=NAVY)]
    if voltar:
        c.append(icon(f"ico{prefixo}Voltar", "Icon.ChevronLeft", 12, 34, 48, 48, onselect=voltar, Tooltip='"Voltar"'))
    else:
        c.append(rect(f"rec{prefixo}Logo", 20, 38, 42, 42, fill=ACCENT, **{}))
        c.append(icon(f"ico{prefixo}Logo", "Icon.CheckBadge", 20, 38, 42, 42))
    c.append(label(f"lbl{prefixo}Titulo", titulo, 74, 26, "Parent.Width - 94 - " + ("170" if direita else "0"), 40, size=19, bold=True, color=WHITE))
    c.append(label(f"lbl{prefixo}Subtitulo", subtitulo, 74, 64, "Parent.Width - 94 - " + ("170" if direita else "0"), 30, size=12, color=HEAD_SUB))
    if direita:
        c += direita
    return c


# ------------------------------------------------------------------ fórmulas compartilhadas
INIT = '''
If(
    !gblIniciado,
    Concurrent(
        ClearCollect(colConfig, Configuracoes),
        ClearCollect(colAcessos, Filter(Acessos, Ativo = true)),
        ClearCollect(colEnsaios, Sort(Filter(Ensaios, Ativo = true), Ordem)),
        ClearCollect(
            colPontos,
            AddColumns(Filter(PontosAfericao, Ativo = true), EntradaTxt, Entrada.Value, TipoTxt, TipoRegistro.Value)
        ),
        ClearCollect(colLimites, Limites)
    );
    // Perfil: quem não está na lista Acessos (ou está inativo) é Colaborador
    Set(gblEu, LookUp(colAcessos, Lower(Pessoa.Email) = Lower(User().Email)));
    Set(gblPerfil, If(IsBlank(gblEu), "Colaborador", gblEu.Perfil.Value));
    Set(gblNome, User().FullName);
    Set(
        gblIniciais,
        Upper(Left(First(Split(User().FullName, " ")).Value, 1) & Left(Last(Split(User().FullName, " ")).Value, 1))
    );
    // Turno e letra da escala (lista Configuracoes)
    Set(gblHoraIni, Coalesce(Value(LookUp(colConfig, Title = "TurnoDiaInicio").Valor), 7));
    Set(gblHoraFim, Coalesce(Value(LookUp(colConfig, Title = "TurnoDiaFim").Valor), 19));
    Set(gblDiasAviso, Coalesce(Value(LookUp(colConfig, Title = "DiasAvisoCalibracao").Valor), 30));
    Set(gblAgora, Now());
    Set(gblTurnoAtual, If(Hour(gblAgora) >= gblHoraIni && Hour(gblAgora) < gblHoraFim, "07x19", "19x07"));
    Set(gblDataTurnoAtual, If(Hour(gblAgora) < gblHoraIni, DateAdd(Today(), -1, TimeUnit.Days), Today()));
    Set(
        gblLetraTurno,
        With(
            {
                wRef: DateValue(LookUp(colConfig, Title = "EscalaDataReferencia").Valor),
                w07: Split(LookUp(colConfig, Title = "EscalaLetras07x19").Valor, ","),
                w19: Split(LookUp(colConfig, Title = "EscalaLetras19x07").Valor, ",")
            },
            Index(
                If(gblTurnoAtual = "07x19", w07, w19),
                Mod(DateDiff(wRef, gblDataTurnoAtual, TimeUnit.Days), CountRows(w07)) + 1
            ).Value
        )
    );
    Set(gblMinhaLetra, Coalesce(gblEu.Letra.Value, gblLetraTurno));
    Set(gblIniciado, true)
)'''

CARREGAR_EQUIP = '''
ClearCollect(
    colEquip,
    AddColumns(
        Filter(Equipamentos, Ativo = true),
        StatusCalc,
        If(
            ControlaCalibracao && UltimoResultado = "Recusada", "Bloqueado · Calibração recusada",
            ControlaCalibracao && ManutencaoSolicitada, "Bloqueado · Em manutenção",
            ControlaCalibracao && !IsBlank(ProximaCalibracao) && ProximaCalibracao < Today(), "Bloqueado · Calibração vencida",
            UltimaAfericaoResultado = "Não conforme" &&
                !(UltimoResultado = "Aprovada" && !IsBlank(UltimaCalibracao) && UltimaCalibracao >= UltimaAfericao),
                "Bloqueado · Aferição não conforme",
            ControlaCalibracao && IsBlank(ProximaCalibracao), "Liberado · Sem data de calibração",
            ControlaCalibracao && ProximaCalibracao <= DateAdd(Today(), gblDiasAviso, TimeUnit.Days), "Liberado · Calibração vence em breve",
            "Liberado"
        )
    )
)'''

CAUSAS = '["Amostra fora do padrão", "Equipamento sujo ou mal posicionado", "Variação de temperatura ou umidade", "Padrão ou referência com problema", "Suspeita de descalibração", "Falha do equipamento", "Outro"]'


def salvar(prefixo):
    """Fórmula do botão 'Confirmar envio' (igual nas duas telas de lançamento)."""
    return F(f'''
        If(
            gblNC > 0 && IsBlank(drp{prefixo}Causa.Selected.Value),
                Notify("Selecione a causa provável da não conformidade.", NotificationType.Error),
            gblNC > 0 && Len(Trim(txt{prefixo}Just.Text)) < 10,
                Notify("Descreva o que foi verificado (mínimo de 10 caracteres).", NotificationType.Error),
            IsBlank(dtp{prefixo}Data.SelectedDate) || dtp{prefixo}Data.SelectedDate > Today(),
                Notify("A data do ensaio não pode ser futura.", NotificationType.Error),

            UpdateContext({{locSalvando: true}});
            Set(gblLote, "LCP-" & Text(Now(), "yymmdd-hhmmss") & "-" & gblIniciais);
            With(
                {{
                    wData: dtp{prefixo}Data.SelectedDate,
                    wTurno: drp{prefixo}Turno.Selected.Value,
                    wFora: If(
                        dtp{prefixo}Data.SelectedDate = gblDataTurnoAtual && drp{prefixo}Turno.Selected.Value = gblTurnoAtual,
                        "Não",
                        "Sim"
                    ),
                    wCausa: If(gblNC > 0, drp{prefixo}Causa.Selected.Value, ""),
                    wJust: If(gblNC > 0, drp{prefixo}Causa.Selected.Value & ": " & Trim(txt{prefixo}Just.Text), ""),
                    wObs: Trim(txt{prefixo}Obs.Text)
                }},
                IfError(
                    Collect(
                        Afericoes,
                        ForAll(
                            colResultados As r,
                            {{
                                Title: gblLote,
                                Ensaio: gblEnsaio.NomeBD,
                                TipoRegistro: r.Tipo,
                                DataEnsaio: wData,
                                Equipamento: r.Equipamento,
                                TAG: r.TAG,
                                Parametro: r.Parametro,
                                Unidade: r.Unidade,
                                Valor: r.Valor,
                                Referencia: r.Referencia,
                                Diferenca: r.Diferenca,
                                LimInferior: r.LI,
                                LimSuperior: r.LS,
                                Tolerancia: r.Tol,
                                Resultado: r.Resultado,
                                Responsavel: gblNome,
                                ResponsavelEmail: Lower(User().Email),
                                Letra: gblMinhaLetra,
                                Turno: wTurno,
                                Observacao: Concat(
                                    Filter(
                                        Table({{t: r.ObsItem}}, {{t: If(r.Resultado = "Não conforme", wJust, "")}}, {{t: wObs}}),
                                        !IsBlank(t)
                                    ),
                                    t,
                                    "; "
                                ),
                                Causa: If(r.Resultado = "Não conforme", wCausa, ""),
                                ItensNaoOK: r.ItensNaoOK,
                                RegistradoEm: Now(),
                                TurnoRegistro: gblTurnoAtual,
                                ForaDoTurno: wFora,
                                Origem: "App",
                                ModoLeitura: gblModoLeitura
                            }}
                        )
                    ),
                    Notify("Não foi possível gravar os resultados: " & FirstError.Message, NotificationType.Error),

                    Patch(
                        Lancamentos,
                        Defaults(Lancamentos),
                        {{
                            Title: gblLote,
                            Ensaio: gblEnsaio.NomeBD,
                            DataEnsaio: wData,
                            Turno: wTurno,
                            Letra: gblMinhaLetra,
                            Responsavel: gblNome,
                            ResponsavelEmail: Lower(User().Email),
                            QtdResultados: CountRows(colResultados),
                            QtdNaoConformes: gblNC,
                            TAGsNaoConformes: Concat(
                                Filter(colResultados, Resultado = "Não conforme"),
                                Coalesce(TAG, Equipamento) & " (" & Parametro & ")",
                                ", "
                            ),
                            Causa: wCausa,
                            Observacao: Concat(Filter(Table({{t: wJust}}, {{t: wObs}}), !IsBlank(t)), t, "; "),
                            RegistradoEm: Now(),
                            ForaDoTurno: wFora,
                            Situacao: "Ativo",
                            Origem: "App",
                            ModoLeitura: gblModoLeitura
                        }}
                    ),
                    Notify("Resultados gravados, mas o resumo do lançamento falhou: " & FirstError.Message, NotificationType.Warning),

                    Set(gblUltimoLote, gblLote);
                    Set(gblUltimoLoteTexto, gblEnsaio.Nome & " · " & CountRows(colResultados) & " resultados");
                    Notify(
                        CountRows(colResultados) & " resultados enviados" &
                        If(gblNC > 0, " · " & gblNC & " não conforme(s): a instrumentação será avisada", ""),
                        If(gblNC > 0, NotificationType.Warning, NotificationType.Success)
                    );
                    UpdateContext({{locRevisar: false}});
                    Navigate(scrInicio, ScreenTransition.UnCover)
                )
            );
            UpdateContext({{locSalvando: false}})
        )''')


def topo_lancamento(prefixo):
    """Instruções, responsável, data e turno do ensaio (telas de lançamento)."""
    return [
        label(f"lbl{prefixo}Instr", "gblEnsaio.Instrucoes", 20, 120, "Parent.Width - 40", 58, size=12, color=SUB,
              VerticalAlign="VerticalAlign.Top"),
        label(f"lbl{prefixo}Quem", F('gblNome & " · Letra " & gblMinhaLetra & If(gblPerfil <> "Laboratorista", "  (perfil sem permissão de lançar)", "")'),
              20, 178, "Parent.Width - 40", 26, size=12, bold=True, color=NAVY),
        ctl(f"dtp{prefixo}Data", "Classic/DatePicker@2.6.0", DefaultDate="gblDataTurnoAtual", Format="DateTimeFormat.ShortDate",
            X=20, Y=208, Width="(Parent.Width - 50) / 2", Height=50, Size=14),
        dropdown(f"drp{prefixo}Turno", '["07x19", "19x07"]', "30 + (Parent.Width - 50) / 2", 208, "(Parent.Width - 50) / 2", 50,
                 default="gblTurnoAtual"),
    ]


def barra_e_revisao(prefixo, revisar_formula, resumo_formula):
    """Barra inferior com resumo + botão 'Revisar e enviar' e o painel de revisão."""
    vis = "locRevisar"
    nc = "gblNC > 0"
    return [
        rect(f"rec{prefixo}Barra", 0, "Parent.Height - 104", "Parent.Width", 104, fill=WHITE, BorderColor=LINE, BorderThickness=1),
        label(f"lbl{prefixo}Resumo", resumo_formula, 20, "Parent.Height - 94", "Parent.Width - 270", 84, size=12, color=SUB),
        button(f"btn{prefixo}Revisar", '"Revisar e enviar"', "Parent.Width - 240", "Parent.Height - 88", 220, 70, revisar_formula,
               DisplayMode=F('If(gblPerfil = "Laboratorista", DisplayMode.Edit, DisplayMode.Disabled)')),
        # ---- painel de revisão (fica por cima de tudo)
        rect(f"rec{prefixo}RevFundo", 0, 0, "Parent.Width", "Parent.Height", fill=OVERLAY, Visible=vis),
        rect(f"rec{prefixo}RevCartao", 20, 170, "Parent.Width - 40", 760, fill=WHITE, Visible=vis),
        label(f"lbl{prefixo}RevTitulo", '"Revisar e enviar"', 44, 190, "Parent.Width - 88", 40, size=19, bold=True, color=NAVY, Visible=vis),
        label(f"lbl{prefixo}RevResumo", F(f'''
            gblEnsaio.Nome & Char(10) &
            Text(dtp{prefixo}Data.SelectedDate, "dd/mm/yyyy") & " · turno " & drp{prefixo}Turno.Selected.Value &
            If(dtp{prefixo}Data.SelectedDate = gblDataTurnoAtual && drp{prefixo}Turno.Selected.Value = gblTurnoAtual, "", "  (fora do turno)") & Char(10) &
            CountRows(colResultados) & " resultados · " &
            CountRows(Filter(colResultados, Resultado = "Conforme")) & " conformes · " &
            gblNC & " não conformes · " &
            CountRows(Filter(colResultados, Resultado = "Informativo")) & " informativos"'''),
              44, 232, "Parent.Width - 88", 96, size=13, color=TXT, VerticalAlign="VerticalAlign.Top", Visible=vis),
        label(f"lbl{prefixo}RevNC", F('"Não conformes: " & Concat(Filter(colResultados, Resultado = "Não conforme"), Coalesce(TAG, Equipamento) & " (" & Parametro & ")", ", ")'),
              44, 332, "Parent.Width - 88", 84, size=12, bold=True, color=BAD, Fill=BAD_S, PaddingLeft=12, PaddingRight=12,
              VerticalAlign="VerticalAlign.Top", Visible=f"{vis} && {nc}"),
        label(f"lbl{prefixo}RevCausaRot", '"CAUSA PROVÁVEL *"', 44, 424, "Parent.Width - 88", 26, size=10, bold=True, color=SUB, Visible=f"{vis} && {nc}"),
        dropdown(f"drp{prefixo}Causa", CAUSAS, 44, 450, "Parent.Width - 88", 50, Visible=f"{vis} && {nc}", AllowEmptySelection=True, default='""'),
        label(f"lbl{prefixo}RevJustRot", '"O QUE FOI VERIFICADO *"', 44, 508, "Parent.Width - 88", 26, size=10, bold=True, color=SUB, Visible=f"{vis} && {nc}"),
        textinput(f"txt{prefixo}Just", 44, 534, "Parent.Width - 88", 110, '"Ex.: repeti o ensaio, conferi o padrão e avisei o supervisor"',
                  Mode="TextMode.MultiLine", size=13, Visible=f"{vis} && {nc}"),
        label(f"lbl{prefixo}RevObsRot", '"OBSERVAÇÃO (OPCIONAL)"', 44, 652, "Parent.Width - 88", 26, size=10, bold=True, color=SUB, Visible=vis),
        textinput(f"txt{prefixo}Obs", 44, 678, "Parent.Width - 88", 54, '"Opcional"', size=13, Visible=vis),
        button(f"btn{prefixo}RevVoltar", '"Voltar e corrigir"', 44, 840, "(Parent.Width - 100) / 2", 66,
               "UpdateContext({locRevisar: false})", fill=WHITE, color=NAVY, BorderColor=NAVY, BorderThickness=2, Visible=vis),
        button(f"btn{prefixo}RevEnviar", F('If(locSalvando, "Enviando...", "Confirmar envio")'), "56 + (Parent.Width - 100) / 2", 840,
               "(Parent.Width - 100) / 2", 66, salvar(prefixo), Visible=vis,
               DisplayMode=F('If(locSalvando, DisplayMode.Disabled, DisplayMode.Edit)')),
    ]


# ------------------------------------------------------------------ TELA 1 · Início
def previsto(r=""):
    """Ensaio previsto no turno atual (grade SegD … DomN da lista Ensaios). r = prefixo do registro ('en.')."""
    return f'''
        With(
            {{wd: Weekday(gblDataTurnoAtual, StartOfWeek.Monday), wDia: gblTurnoAtual = "07x19"}},
            Switch(
                wd,
                1, If(wDia, {r}SegD, {r}SegN), 2, If(wDia, {r}TerD, {r}TerN), 3, If(wDia, {r}QuaD, {r}QuaN),
                4, If(wDia, {r}QuiD, {r}QuiN), 5, If(wDia, {r}SexD, {r}SexN), 6, If(wDia, {r}SabD, {r}SabN),
                If(wDia, {r}DomD, {r}DomN)
            )
        )'''


# Agenda do turno: uma linha por equipamento previsto (peneiradores: uma por conjunto).
# Feito = existe aferição daquele equipamento no turno atual. Só leitura: para lançar, lê-se o QR do equipamento.
AGENDA = f'''
ClearCollect(colLancTurno, Filter(Lancamentos, DataEnsaio = gblDataTurnoAtual && Turno = gblTurnoAtual && Situacao = "Ativo"));
ClearCollect(colAfTurno, Filter(Afericoes, DataEnsaio = gblDataTurnoAtual && Turno = gblTurnoAtual));
Clear(colAgendaUn);
ForAll(
    Filter(colEnsaios As en, {textwrap.indent(textwrap.dedent(previsto("en.")).strip(), "    ").strip()}) As en,
    Collect(
        colAgendaUn,
        If(
            en.Title = "PEN",
            ForAll(
                Distinct(Filter(colPontos, Ensaio = "PEN" && !IsBlank(Conjunto) && !IsBlank(TAG)), Conjunto) As u,
                {{
                    Ensaio: en.Title,
                    Un: u.Value,
                    Feito: CountRows(
                        Filter(colPontos, Ensaio = "PEN" && Conjunto = u.Value && TAG in Filter(colAfTurno, Ensaio = en.NomeBD).TAG)
                    ) > 0
                }}
            ),
            ForAll(
                Distinct(Filter(colPontos, Ensaio = en.Title && !IsBlank(TAG)), TAG) As u,
                {{Ensaio: en.Title, Un: u.Value, Feito: !IsBlank(LookUp(colAfTurno, Ensaio = en.NomeBD && TAG = u.Value))}}
            )
        )
    )
);
ClearCollect(
    colAgenda,
    ForAll(
        Filter(colEnsaios As en, {textwrap.indent(textwrap.dedent(previsto("en.")).strip(), "    ").strip()}) As en,
        With(
            {{w: Filter(colAgendaUn, Ensaio = en.Title), wEnv: !IsBlank(LookUp(colLancTurno, Ensaio = en.NomeBD))}},
            {{
                Title: en.Title,
                Nome: en.Nome,
                Ordem: en.Ordem,
                Total: CountRows(w),
                Feitos: CountIf(w, Feito),
                Lista: Concat(w, Un & If(Feito, " ✓", ""), "   ·   "),
                Completo: If(CountRows(w) = 0, wEnv, CountIf(w, Feito) = CountRows(w))
            }}
        )
    )
)'''


# Abrir um equipamento pela TAG (scanner do app ou TAG digitada). Mesma regra nos dois casos:
# só abre equipamento cadastrado e ativo; nada de lista nem busca.
def abrir_tag(expr_tag, modo, limpar=""):
    return f'''
        With(
            {{wTag: Upper(Trim({expr_tag}))}},
            If(
                IsBlank(wTag), Notify("Digite a TAG do equipamento.", NotificationType.Error),
                IsBlank(LookUp(Equipamentos, Title = wTag && Ativo = true)),
                    Notify("Equipamento " & wTag & " não encontrado. Confira a TAG ou o QR Code.", NotificationType.Error),
                Set(gblTag, wTag);
                Set(gblModoLeitura, "{modo}");{limpar}
                Navigate(scrEquipamento, ScreenTransition.Cover)
            )
        )'''


# OnScan do leitor de QR. O conteúdo lido pode ser o endereço da imagem
# (…/ConsultaQR/TAG.png), um link do app (…&equip=TAG) ou a própria TAG.
ONSCAN = '''
    With(
        {wLido: Trim(First(bcrIniScanner.Barcodes).Value)},
        With(
            {wArq: Last(Split(First(Split(wLido, "?")).Value, "/")).Value},
            With(
                {
                    wTag: Upper(
                        If(
                            "equip=" in Lower(wLido), First(Split(Last(Split(wLido, "equip=")).Value, "&")).Value,
                            EndsWith(Lower(wArq), ".png"), Substitute(Left(wArq, Len(wArq) - 4), "%20", " "),
                            wLido
                        )
                    )
                },
                If(
                    IsBlank(LookUp(Equipamentos, Title = wTag && Ativo = true)),
                        Notify("Este QR Code não é de um equipamento cadastrado (" & wTag & ").", NotificationType.Error),
                    Set(gblTag, wTag);
                    Set(gblModoLeitura, "QR");
                    Navigate(scrEquipamento, ScreenTransition.Cover)
                )
            )
        )
    )'''


def tela_inicio():
    pill_perfil = label("lblIniPerfil", "Upper(gblPerfil)", "Parent.Width - 200", 128, 180, 34, size=10, bold=True,
                        Align="Align.Center",
                        Fill=F(f'Switch(gblPerfil, "Laboratorista", {LAB_S}, "Instrumentação", {INS_S}, {MUT_S})'),
                        Color=F(f'Switch(gblPerfil, "Laboratorista", {LAB}, "Instrumentação", {INS}, {MUT})'))
    lab = 'gblPerfil = "Laboratorista"'
    ins = 'gblPerfil = "Instrumentação"'
    equipe = 'gblPerfil <> "Colaborador"'
    y_scan = F(f'If({lab}, 206, 176)')
    y_tag = F(f'If({lab}, 322, 292)')
    y_prog = 390
    y_ag = F(f'If(IsBlank(gblUltimoLote), {y_prog + 94}, {y_prog + 154})')
    completos = 'CountIf(colAgenda, Completo)'
    vencidas = 'CountIf(colEquip, ControlaCalibracao && !IsBlank(ProximaCalibracao) && ProximaCalibracao < Today())'
    vencendo = ('CountIf(colEquip, ControlaCalibracao && !IsBlank(ProximaCalibracao) && ProximaCalibracao >= Today() && '
                'ProximaCalibracao <= DateAdd(Today(), gblDiasAviso, TimeUnit.Days))')

    filhos = cabecalho("Ini", '"Metrologia"', '"Laboratório físico · calibração e aferição"') + [
        label("lblIniOla", F('"Olá, " & First(Split(gblNome, " ")).Value'), 20, 124, "Parent.Width - 240", 42, size=18, bold=True),
        pill_perfil,
        label("lblIniTurno", F('Text(gblDataTurnoAtual, "dd/mm/yyyy") & "   ·   Turno " & gblTurnoAtual & "   ·   Letra " & gblMinhaLetra'),
              20, 168, "Parent.Width - 40", 30, size=12, bold=True, color=NAVY, Visible=lab),
        # ----- leitor de QR Code (laboratorista e instrumentação)
        # O controle "Leitor de código de barras" (bcrIniScanner) é inserido à mão por cima deste cartão (guia, seção 2.5).
        rect("recIniScan", 20, y_scan, "Parent.Width - 40", 104, fill=NAVY, Visible=equipe, **radius(16)),
        rect("recIniScanIco", 40, F(f'{y_scan[1:]} + 20'), 64, 64, fill=ACCENT, Visible=equipe, **radius(14)),
        icon("icoIniScan", "Icon.Camera", 52, F(f'{y_scan[1:]} + 32'), 40, 40, Visible=equipe),
        label("lblIniScanTit", '"Ler QR Code"', 124, F(f'{y_scan[1:]} + 18'), "Parent.Width - 164", 38, size=18, bold=True, color=WHITE, Visible=equipe),
        label("lblIniScanSub", '"Toque e aponte a câmera para a etiqueta do equipamento"', 124, F(f'{y_scan[1:]} + 54'), "Parent.Width - 164", 34,
              size=11, color=HEAD_SUB, Visible=equipe),
        textinput("txtIniTag", 20, y_tag, "Parent.Width - 196", 56, '"Sem câmera? Digite a TAG (ex.: BAL-01)"', Visible=equipe),
        button("btnIniAbrir", '"Abrir"', "Parent.Width - 168", y_tag, 148, 56,
               F(abrir_tag("txtIniTag.Text", "Digitado", "\n                Reset(txtIniTag);")), fill=WHITE, color=NAVY, size=14,
               BorderColor=NAVY, BorderThickness=2, Visible=equipe),
        # ----- laboratorista: agenda do turno (só leitura)
        rect("recIniProg", 20, y_prog, "Parent.Width - 40", 84, fill=WHITE, Visible=lab),
        label("lblIniProgRot", '"AGENDA DO TURNO"', 36, y_prog + 6, "Parent.Width - 72", 26, size=10, bold=True, color=SUB, Visible=lab),
        label("lblIniProg", F(f'{completos} & " de " & CountRows(colAgenda) & " ensaios completos"'), 36, y_prog + 30, "Parent.Width - 72", 30,
              size=16, bold=True, color=F(f'If({completos} = CountRows(colAgenda), {OK}, {NAVY})'), Visible=lab),
        label("lblIniDica", '"Para lançar, leia o QR Code do equipamento."', 36, y_prog + 56, "Parent.Width - 72", 24,
              size=11, color=SUB, Visible=lab),
        label("lblIniDesfazer", F('"Último envio: " & gblUltimoLoteTexto'), 20, y_prog + 94, "Parent.Width - 200", 50, size=12, color=NAVY,
              Fill=LAB_S, PaddingLeft=12, Visible=f'{lab} && !IsBlank(gblUltimoLote)'),
        button("btnIniDesfazer", '"Desfazer"', "Parent.Width - 172", y_prog + 94, 152, 50, F(f'''
            IfError(
                Remove(Afericoes, Filter(Afericoes, Title = gblUltimoLote)),
                Notify("Não foi possível desfazer: " & FirstError.Message, NotificationType.Error),
                Patch(Lancamentos, LookUp(Lancamentos, Title = gblUltimoLote), {{Situacao: "Desfeito"}}),
                Notify("Resultados removidos, mas o resumo não foi atualizado.", NotificationType.Warning),
                Set(gblUltimoLote, Blank());
                {textwrap.indent(textwrap.dedent(AGENDA).strip(), "                ").strip()};
                Notify("Envio desfeito.", NotificationType.Success)
            )'''), fill=WHITE, color=BAD, size=13, BorderColor=BAD, BorderThickness=1,
               Visible=f'{lab} && !IsBlank(gblUltimoLote)'),
        gallery("galIniAgenda", F('Sort(colAgenda, Ordem)'), 20, y_ag, "Parent.Width - 40", F(f'Parent.Height - ({y_ag[1:]}) - 10'), 112, [
            rect("recAgCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 10", fill=WHITE),
            rect("recAgFaixa", 0, 0, 6, "Parent.TemplateHeight - 10", fill=F(f'If(ThisItem.Completo, {OK}, {WARN})')),
            label("lblAgSigla", "ThisItem.Title", 18, 14, 58, 40, size=11, bold=True, Align="Align.Center",
                  Fill=LAB_S, Color=LAB),
            label("lblAgNome", "ThisItem.Nome", 88, 10, "Parent.TemplateWidth - 270", 40, size=14, bold=True),
            label("lblAgSit", F('If(ThisItem.Completo, "FEITO", ThisItem.Feitos > 0, ThisItem.Feitos & " DE " & ThisItem.Total, "FALTA")'),
                  "Parent.TemplateWidth - 150", 16, 136, 32, size=10, bold=True, Align="Align.Center",
                  Fill=F(f'If(ThisItem.Completo, {OK_S}, {WARN_S})'), Color=F(f'If(ThisItem.Completo, {OK}, {WARN})')),
            label("lblAgEquip", "ThisItem.Lista", 18, 56, "Parent.TemplateWidth - 32", 44, size=11, color=SUB,
                  VerticalAlign="VerticalAlign.Top"),
        ], Visible=lab),
        label("lblIniAgVazia", '"Nenhum ensaio previsto para este turno."', 20, F(f'{y_ag[1:]} + 20'), "Parent.Width - 40", 60,
              size=13, color=SUB, Align="Align.Center", Visible=f'{lab} && CountRows(colAgenda) = 0'),
        # ----- instrumentação: só contagens (a lista das pendências chega no e-mail diário)
        rect("recIniVenc", 20, 368, "(Parent.Width - 52) / 2", 104, fill=BAD_S, Visible=ins, **radius(14)),
        label("lblIniVencN", F(vencidas), 36, 378, "(Parent.Width - 52) / 2 - 32", 50, size=26, bold=True, color=BAD, Visible=ins),
        label("lblIniVencT", '"calibrações vencidas"', 36, 428, "(Parent.Width - 52) / 2 - 32", 34, size=12, color=BAD, Visible=ins),
        rect("recIniBreve", "32 + (Parent.Width - 52) / 2", 368, "(Parent.Width - 52) / 2", 104, fill=WARN_S, Visible=ins, **radius(14)),
        label("lblIniBreveN", F(vencendo), "48 + (Parent.Width - 52) / 2", 378, "(Parent.Width - 52) / 2 - 32", 50, size=26, bold=True, color=WARN, Visible=ins),
        label("lblIniBreveT", F('"vencem em até " & gblDiasAviso & " dias"'), "48 + (Parent.Width - 52) / 2", 428, "(Parent.Width - 52) / 2 - 32", 34,
              size=12, color=WARN, Visible=ins),
        label("lblIniInsNota", '"A lista com as TAGs chega no e-mail diário das 7h. Para registrar uma calibração, leia o QR Code do equipamento ou digite a TAG."',
              20, 484, "Parent.Width - 40", 70, size=12, color=SUB, Visible=ins),
        # ----- colaborador: a consulta é pela câmera (imagem no SharePoint)
        rect("recIniQR", 20, 186, "Parent.Width - 40", 400, fill=WHITE, Visible=f'!({equipe})'),
        rect("recIniQRIco", "(Parent.Width - 120) / 2", 220, 120, 120, fill=NAVY, Visible=f'!({equipe})', **radius(24)),
        icon("icoIniQR", "Icon.Camera", "(Parent.Width - 72) / 2", 244, 72, 72, Visible=f'!({equipe})'),
        label("lblIniQRTit", '"Consulte pela câmera do celular"', 40, 362, "Parent.Width - 80", 44, size=18, bold=True, color=NAVY,
              Align="Align.Center", Visible=f'!({equipe})'),
        label("lblIniQRTxt", '"Aponte a câmera do celular para a etiqueta do equipamento: abre a situação da calibração." & Char(10) & Char(10) & "Este app é usado pela equipe do laboratório e da instrumentação para registrar aferições e calibrações."',
              48, 410, "Parent.Width - 96", 160, size=13, color=SUB, Align="Align.Center", VerticalAlign="VerticalAlign.Top",
              Visible=f'!({equipe})'),
        icon("icoIniAtualizar", "Icon.Reset", "Parent.Width - 64", 38, 46, 46, onselect=F(f'''
            Set(gblIniciado, false);
            {textwrap.indent(textwrap.dedent(INIT).strip(), "            ").strip()};
            If(
                gblPerfil = "Laboratorista",
                {textwrap.indent(textwrap.dedent(AGENDA).strip(), "                ").strip()}
            );
            If(gblPerfil = "Instrumentação", {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "                ").strip()});
            Notify("Dados atualizados.", NotificationType.Success)'''), Tooltip='"Atualizar"'),
    ]
    onvisible = F(f'''
        {textwrap.dedent(INIT).strip()};
        If(
            gblPerfil = "Laboratorista",
            {textwrap.indent(textwrap.dedent(AGENDA).strip(), "            ").strip()}
        );
        If(gblPerfil = "Instrumentação", {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "            ").strip()})''')
    return {"scrInicio": {"Properties": {"Fill": f"={BG}", "OnVisible": onvisible}, "Children": filhos}}


# ------------------------------------------------------------------ TELA 2 · Equipamento
STATUS_EQ = '''
    With(
        {e: gblEquip},
        If(
            IsBlank(e), "",
            e.ControlaCalibracao && e.UltimoResultado = "Recusada", "Bloqueado · Calibração recusada",
            e.ControlaCalibracao && e.ManutencaoSolicitada, "Bloqueado · Em manutenção",
            e.ControlaCalibracao && !IsBlank(e.ProximaCalibracao) && e.ProximaCalibracao < Today(), "Bloqueado · Calibração vencida",
            gblUltAfNC && !(e.UltimoResultado = "Aprovada" && !IsBlank(e.UltimaCalibracao) && e.UltimaCalibracao >= gblUltAfData),
                "Bloqueado · Aferição não conforme",
            e.ControlaCalibracao && IsBlank(e.ProximaCalibracao), "Liberado · Sem data de calibração",
            e.ControlaCalibracao && e.ProximaCalibracao <= DateAdd(Today(), gblDiasAviso, TimeUnit.Days), "Liberado · Calibração vence em breve",
            "Liberado"
        )
    )'''

CARREGAR_FICHA = f'''
Set(gblEquip, LookUp(Equipamentos, Title = gblTag));
ClearCollect(colHist, FirstN(Sort(Filter(Afericoes, TAG = gblTag), DataEnsaio, SortOrder.Descending), 40));
// Instrumentação: as 3 últimas calibrações, com as observações do técnico
ClearCollect(colCalHist, FirstN(Sort(Filter(Calibracoes, Title = gblTag), DataCalibracao, SortOrder.Descending), 3));
Set(gblUltAfData, First(colHist).DataEnsaio);
Set(gblUltAfNC, CountRows(Filter(colHist, DataEnsaio = gblUltAfData && Resultado = "Não conforme")) > 0);
// Peneirador = conjunto com etiqueta própria; Peneira = item calibrado pela instrumentação
Set(gblEhPeneirador, gblEquip.TipoEquipamento = "Peneirador" && !IsBlank(LookUp(colPontos, Ensaio = "PEN" && Conjunto = gblEquip.Title)));
Set(gblEhPeneira, gblEquip.TipoEquipamento = "Peneira");
Set(gblConjuntoRef, If(gblEhPeneirador, gblEquip.Title, gblEhPeneira, gblEquip.Conjunto, ""));
// Ensaios lançados a partir deste QR: os da coluna EnsaioAfericao; peneira instalada = verificação do peneirador dela
Set(
    gblCodigos,
    Filter(
        Split(If(gblEhPeneira, If(IsBlank(gblEquip.Conjunto), "", "PEN"), Coalesce(gblEquip.EnsaioAfericao, "")), ";"),
        !IsBlank(Trim(Value))
    )
);
ClearCollect(
    colPeneirasConj,
    AddColumns(
        Sort(Filter(colPontos, Ensaio = "PEN" && Conjunto = gblEquip.Title), Ordem),
        StatusPen, Coalesce(LookUp(colEquip, Title = TAG).StatusCalc, "Posição vazia")
    )
);
Set(gblStatus, {textwrap.indent(textwrap.dedent(STATUS_EQ).strip(), "")})'''


def texto_lancar(cod, curto):
    """Texto do botão de lançamento para o ensaio de código `cod`."""
    pen = '"Verificar peneiras"' if curto else '"Verificar as peneiras do " & Coalesce(LookUp(colEquip, Title = gblConjuntoRef).Descricao, "peneirador")'
    return f'If({cod} = "PEN", {pen}, "Lançar " & LookUp(colEnsaios, Title = {cod}).Nome)'


def acao_lancar(cod):
    """Abre o formulário do ensaio `cod` só com os pontos deste equipamento (peneiras: só o peneirador dele)."""
    return F(f'''
        With(
            {{wCod: {cod}}},
            Set(gblEnsaio, LookUp(colEnsaios, Title = wCod));
            Set(gblConjuntoFiltro, If(wCod = "PEN", gblConjuntoRef, ""));
            Set(gblDestacar, If(wCod = "PEN", If(gblEhPeneira, gblEquip.Title, ""), gblEquip.Title))
        );
        Navigate(If(gblEnsaio.Formulario.Value = "Checklist", scrVerificacao, scrEnsaio), ScreenTransition.Cover)''')


def tela_equipamento():
    onvisible = F(f'''
        {textwrap.dedent(INIT).strip()};
        // Aberto por um link (botão de e-mail) com ?equip=TAG
        If(IsBlank(gblTag), Set(gblTag, Upper(Param("equip"))); Set(gblModoLeitura, "Link"));
        If(CountRows(colEquip) = 0, {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "    ").strip()});
        {textwrap.indent(textwrap.dedent(CARREGAR_FICHA).strip(), "        ").strip()};
        If(
            gblPerfil = "Laboratorista",
            {textwrap.indent(textwrap.dedent(AGENDA).strip(), "            ").strip()}
        );
        UpdateContext({{locInstalar: false}})''')
    vivo = "!IsBlank(gblEquip)"
    bloq = 'StartsWith(gblStatus, "Bloqueado")'
    dias = "DateDiff(Today(), gblEquip.ProximaCalibracao, TimeUnit.Days)"
    lancar_visivel = f'{vivo} && gblPerfil = "Laboratorista" && CountRows(gblCodigos) >= 1'
    lancar2_visivel = f'{vivo} && gblPerfil = "Laboratorista" && CountRows(gblCodigos) >= 2'
    agenda_eq = 'Filter(colAgendaUn, Un = If(Ensaio = "PEN", gblConjuntoRef, gblEquip.Title))'
    pos_item = 'Parametro & "  ·  atual: " & Coalesce(TAG, "vazia")'
    instalar = F(f'''
        With(
            {{pos: LookUp(colPontos, Ensaio = "PEN" && {pos_item} = drpEqPosicao.Selected.Value)}},
            If(
                IsBlank(pos), Notify("Escolha a posição.", NotificationType.Error),
                pos.TAG = gblEquip.Title, Notify("Esta peneira já está nessa posição.", NotificationType.Information),
                IfError(
                    // 1) a posição passa a ter esta peneira
                    Patch(PontosAfericao, LookUp(PontosAfericao, Title = pos.Title), {{TAG: gblEquip.Title, EquipamentoBD: gblEquip.Title}}),
                    Notify("Não foi possível alterar a posição: " & FirstError.Message, NotificationType.Error),
                    // 2) a peneira que estava lá fica fora de uso
                    If(
                        !IsBlank(pos.TAG) && !IsBlank(LookUp(Equipamentos, Title = pos.TAG)),
                        Patch(Equipamentos, LookUp(Equipamentos, Title = pos.TAG), {{Conjunto: "", Grupo: ""}})
                    ),
                    Notify("Posição alterada, mas a peneira anterior não foi atualizada: " & FirstError.Message, NotificationType.Warning),
                    // 3) se esta peneira estava em outra posição, aquela posição fica vazia
                    ForAll(
                        Filter(colPontos, Ensaio = "PEN" && TAG = gblEquip.Title && Title <> pos.Title) As velho,
                        Patch(PontosAfericao, LookUp(PontosAfericao, Title = velho.Title), {{TAG: "", EquipamentoBD: ""}})
                    );
                    Patch(Equipamentos, gblEquip, {{Conjunto: pos.Conjunto, Grupo: pos.Secao}}),
                    Notify("Erro ao atualizar a peneira: " & FirstError.Message, NotificationType.Error),

                    ClearCollect(colPontos, AddColumns(Filter(PontosAfericao, Ativo = true), EntradaTxt, Entrada.Value, TipoTxt, TipoRegistro.Value));
                    {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "                    ").strip()};
                    {textwrap.indent(textwrap.dedent(CARREGAR_FICHA).strip(), "                    ").strip()};
                    UpdateContext({{locInstalar: false}});
                    Notify(gblEquip.Title & " instalada em " & pos.Secao & ".", NotificationType.Success)
                )
            )
        )''')
    filhos = cabecalho("Eq", F('Coalesce(gblEquip.Title, gblTag)'), F('Coalesce(gblEquip.Descricao, "Equipamento")'),
                       voltar="Navigate(scrInicio, ScreenTransition.UnCover)") + [
        label("lblEqNaoEncontrado", F('"Equipamento não encontrado: " & Coalesce(gblTag, "(sem TAG)") & Char(10) & Char(10) & "Confira se a TAG está cadastrada na lista Equipamentos e marcada como Ativo."'),
              20, 140, "Parent.Width - 40", 200, size=14, color=BAD, Fill=BAD_S, Align="Align.Center", Visible="IsBlank(gblEquip)"),
        rect("recEqVeredito", 0, 112, "Parent.Width", 118, fill=F(f'If({bloq}, RGBA(176, 36, 26, 1), RGBA(30, 127, 79, 1))'), Visible=vivo),
        icon("icoEqVeredito", F(f'If({bloq}, Icon.Blocked, Icon.CheckBadge)'), 22, 138, 66, 66, Visible=vivo),
        label("lblEqVeredito", F(f'If({bloq}, "NÃO UTILIZAR", "LIBERADO PARA USO")'), 100, 124, "Parent.Width - 120", 44, size=20, bold=True, color=WHITE, Visible=vivo),
        label("lblEqMotivo", F(f'''
            If(
                gblStatus = "Liberado", "Calibração e aferição em dia",
                Trim(Last(Split(gblStatus, "·")).Value)
            ) &
            If(
                gblEquip.ControlaCalibracao && !IsBlank(gblEquip.ProximaCalibracao),
                " · calibração " & If({dias} < 0, "venceu em ", "até ") & Text(gblEquip.ProximaCalibracao, "dd/mm/yyyy"),
                ""
            )'''), 100, 168, "Parent.Width - 120", 54, size=12, color=WHITE, VerticalAlign="VerticalAlign.Top", Visible=vivo),
        # identificação
        rect("recEqId", 0, 230, "Parent.Width", 104, fill=WHITE, Visible=vivo),
        label("lblEqTag", "gblEquip.Title", 20, 238, "Parent.Width - 250", 44, size=24, bold=True, color=NAVY, Visible=vivo),
        label("lblEqAgenda", F(f'If(CountIf({agenda_eq}, !Feito) = 0, "FEITO NESTE TURNO ✓", "PREVISTO · FALTA")'),
              "Parent.Width - 230", 244, 210, 34, size=10, bold=True, Align="Align.Center",
              Fill=F(f'If(CountIf({agenda_eq}, !Feito) = 0, {OK_S}, {WARN_S})'),
              Color=F(f'If(CountIf({agenda_eq}, !Feito) = 0, {OK}, {WARN})'),
              Visible=f'{vivo} && gblPerfil = "Laboratorista" && CountRows({agenda_eq}) > 0'),
        label("lblEqInfo", F('''
            gblEquip.Descricao &
            If(
                gblEhPeneira,
                "  ·  " & If(
                    IsBlank(gblEquip.Conjunto), "fora de uso (sem posição)",
                    "instalada no " & Coalesce(LookUp(colEquip, Title = gblEquip.Conjunto).Descricao, gblEquip.Conjunto)
                ),
                ""
            ) &
            If(
                gblEhPeneirador,
                "  ·  " & CountRows(colPeneirasConj) & " posições  ·  " &
                CountIf(colPeneirasConj, StartsWith(StatusPen, "Bloqueado")) & " peneira(s) bloqueada(s)",
                ""
            ) &
            If(!gblEhPeneira && !gblEhPeneirador && IsBlank(gblEquip.EnsaioAfericao), "  ·  sem aferição semanal", "")'''),
              20, 282, "Parent.Width - 40", 44, size=12, color=SUB, Visible=vivo),
        # cartões
        rect("recEqCal", 20, 348, "(Parent.Width - 52) / 2", 132, fill=WHITE, Visible=vivo),
        label("lblEqCalRot", '"CALIBRAÇÃO"', 34, 356, "(Parent.Width - 52) / 2 - 28", 26, size=10, bold=True, color=SUB, Visible=vivo),
        label("lblEqCalVal", F(f'''
            If(
                !gblEquip.ControlaCalibracao, "Não controlada",
                IsBlank(gblEquip.ProximaCalibracao), "Sem data",
                {dias} < 0, "Vencida",
                {dias} & " dias"
            )'''), 34, 382, "(Parent.Width - 52) / 2 - 28", 40, size=19, bold=True,
              color=F(f'If(gblEquip.ControlaCalibracao && !IsBlank(gblEquip.ProximaCalibracao) && {dias} < 0, {BAD}, gblEquip.ControlaCalibracao && !IsBlank(gblEquip.ProximaCalibracao) && {dias} <= gblDiasAviso, {WARN}, {NAVY})'),
              Visible=vivo),
        label("lblEqCalSub", F(f'''
            If(
                !gblEquip.ControlaCalibracao || IsBlank(gblEquip.ProximaCalibracao), "Cadastre a data na lista Equipamentos",
                If({dias} < 0, "desde ", "até ") & Text(gblEquip.ProximaCalibracao, "dd/mm/yyyy")
            ) &
            If(IsBlank(gblEquip.UltimaCalibracao), "", Char(10) & "Última: " & Text(gblEquip.UltimaCalibracao, "dd/mm/yyyy"))'''),
              34, 424, "(Parent.Width - 52) / 2 - 28", 50, size=11, color=SUB, VerticalAlign="VerticalAlign.Top", Visible=vivo),
        rect("recEqAf", "32 + (Parent.Width - 52) / 2", 348, "(Parent.Width - 52) / 2", 132, fill=WHITE, Visible=vivo),
        label("lblEqAfRot", F('If(gblEhPeneirador, "PENEIRAS INSTALADAS", "ÚLTIMA AFERIÇÃO" & If(IsBlank(gblUltAfData), "", " · " & Text(gblUltAfData, "dd/mm")))'),
              "46 + (Parent.Width - 52) / 2", 356, "(Parent.Width - 52) / 2 - 28", 26, size=10, bold=True, color=SUB, Visible=vivo),
        label("lblEqAfVal", F('''
            If(
                gblEhPeneirador, CountIf(colPeneirasConj, StartsWith(StatusPen, "Liberado")) & " de " & CountRows(colPeneirasConj) & " liberadas",
                IsBlank(First(colHist)), "Sem registro",
                First(colHist).Unidade = "itens OK", First(colHist).Valor & " de " & First(colHist).Referencia & " itens OK",
                Text(First(colHist).Valor, "#,##0.###") & " " & First(colHist).Unidade
            )'''), "46 + (Parent.Width - 52) / 2", 382, "(Parent.Width - 52) / 2 - 28", 40, size=17, bold=True,
              color=F(f'If(gblEhPeneirador, If(CountIf(colPeneirasConj, StartsWith(StatusPen, "Bloqueado")) > 0, {BAD}, {OK}), gblUltAfNC, {BAD}, IsBlank(First(colHist)), {MUT}, {OK})'), Visible=vivo),
        label("lblEqAfSub", F('''
            If(
                gblEhPeneirador, "A verificação é registrada em cada peneira",
                IsBlank(First(colHist)),
                    If(gblEhPeneira, "Nenhuma verificação ainda", IsBlank(gblEquip.EnsaioAfericao), "Equipamento sem aferição", "Nenhum lançamento ainda"),
                If(gblUltAfNC, "Não conforme", "Conforme") & " · " & First(colHist).Parametro
            )'''), "46 + (Parent.Width - 52) / 2", 424, "(Parent.Width - 52) / 2 - 28", 50, size=11, color=SUB,
              VerticalAlign="VerticalAlign.Top", Visible=vivo),
        # justificativa
        label("lblEqJust", F('''
            With(
                {nc: LookUp(colHist, DataEnsaio = gblUltAfData && Resultado = "Não conforme")},
                "Justificativa registrada: " & Coalesce(nc.Observacao, "(sem observação)") & Char(10) &
                nc.Responsavel & " · Letra " & nc.Letra & " · " & Text(nc.DataEnsaio, "dd/mm/yyyy")
            )'''), 20, 492, "Parent.Width - 40", 96, size=12, color=BAD, Fill=BAD_S, PaddingLeft=14, PaddingRight=14,
              VerticalAlign="VerticalAlign.Top", BorderColor=BAD, BorderThickness=1, Visible=f"{vivo} && gblUltAfNC && !gblEhPeneirador"),
        label("lblEqHistRot", F('If(gblEhPeneirador, "POSIÇÕES DESTE PENEIRADOR", gblPerfil = "Instrumentação", "ÚLTIMAS CALIBRAÇÕES E OBSERVAÇÕES", "HISTÓRICO DE AFERIÇÕES")'),
              20, F('If(gblUltAfNC && !gblEhPeneirador, 600, 494)'), "Parent.Width - 40", 30, size=11, bold=True, color=SUB, Visible=vivo),
        gallery("galEqHist", "colHist", 20, F('If(gblUltAfNC, 632, 526)'), "Parent.Width - 40",
                F('Parent.Height - If(gblUltAfNC, 632, 526) - 112'), 80, [
            rect("recHiCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 6", fill=WHITE),
            label("lblHiLinha1", F('Text(ThisItem.DataEnsaio, "dd/mm") & " · " & ThisItem.Parametro'), 14, 6, "Parent.TemplateWidth - 200", 32, size=13, bold=True),
            label("lblHiLinha2", F('ThisItem.Responsavel & " · Letra " & ThisItem.Letra & " · " & ThisItem.Turno'), 14, 38, "Parent.TemplateWidth - 200", 30, size=11, color=SUB),
            label("lblHiValor", F('If(ThisItem.Unidade = "itens OK", ThisItem.Valor & "/" & ThisItem.Referencia & " OK", Text(ThisItem.Valor, "#,##0.###") & " " & ThisItem.Unidade)'),
                  "Parent.TemplateWidth - 186", 6, 172, 32, size=12, bold=True, Align="Align.Right",
                  Color=res_color("ThisItem.Resultado")),
            label("lblHiRes", F('Upper(ThisItem.Resultado)'), "Parent.TemplateWidth - 186", 40, 172, 26, size=9, bold=True,
                  Align="Align.Center", Fill=res_fill("ThisItem.Resultado"), Color=res_color("ThisItem.Resultado")),
        ], Visible=f'{vivo} && !gblEhPeneirador && gblPerfil <> "Instrumentação"'),
        # instrumentação: o que os técnicos anotaram nas últimas calibrações
        gallery("galEqCal", "colCalHist", 20, F('If(gblUltAfNC, 632, 526)'), "Parent.Width - 40",
                F('Parent.Height - If(gblUltAfNC, 632, 526) - 112'), 128, [
            rect("recCaCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 8", fill=WHITE),
            rect("recCaFaixa", 0, 0, 6, "Parent.TemplateHeight - 8", fill=F(f'If(ThisItem.Resultado = "Recusada", {BAD}, {INS})')),
            label("lblCaLinha1", F('Text(ThisItem.DataCalibracao, "dd/mm/yyyy") & " · " & ThisItem.Responsavel'), 20, 6,
                  "Parent.TemplateWidth - 200", 30, size=13, bold=True, color=NAVY),
            label("lblCaRes", F('Upper(ThisItem.Resultado) & If(ThisItem.ManutencaoSolicitada, " · MANUTENÇÃO", "")'),
                  "Parent.TemplateWidth - 186", 8, 172, 28, size=9, bold=True, Align="Align.Center",
                  Fill=F(f'If(ThisItem.Resultado = "Recusada", {BAD_S}, {OK_S})'), Color=F(f'If(ThisItem.Resultado = "Recusada", {BAD}, {OK})')),
            label("lblCaCert", F('If(IsBlank(ThisItem.Certificado), "Sem nº de certificado", "Certificado " & ThisItem.Certificado) & If(IsBlank(ThisItem.MotivoRecusa), "", " · Recusa: " & ThisItem.MotivoRecusa)'),
                  20, 36, "Parent.TemplateWidth - 34", 26, size=11, color=SUB),
            label("lblCaObs", F('Coalesce(ThisItem.Observacoes, "Sem observação.")'), 20, 62, "Parent.TemplateWidth - 34", 54, size=12,
                  color=F(f'If(IsBlank(ThisItem.Observacoes), {SUB}, {TXT})'), VerticalAlign="VerticalAlign.Top",
                  Italic=F('IsBlank(ThisItem.Observacoes)')),
        ], Visible=f'{vivo} && !gblEhPeneirador && gblPerfil = "Instrumentação"'),
        label("lblEqCalVazio", '"Nenhuma calibração registrada no app ainda. A primeira que você registrar aparece aqui para o próximo técnico."',
              20, F('If(gblUltAfNC, 646, 540)'), "Parent.Width - 40", 70, size=12, color=SUB, Align="Align.Center",
              Visible=f'{vivo} && CountRows(colCalHist) = 0 && !gblEhPeneirador && gblPerfil = "Instrumentação"'),
        gallery("galEqPeneiras", "colPeneirasConj", 20, 526, "Parent.Width - 40", "Parent.Height - 526 - 112", 76, [
            rect("recPcCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 6", fill=WHITE),
            rect("recPcFaixa", 0, 0, 6, "Parent.TemplateHeight - 6", fill=status_color("ThisItem.StatusPen")),
            label("lblPcTag", "ThisItem.TAG", 20, 6, "Parent.TemplateWidth - 230", 32, size=14, bold=True, color=NAVY),
            label("lblPcMalha", F('Trim(First(Split(ThisItem.Parametro, "(")).Value)'), 20, 36, "Parent.TemplateWidth - 230", 28, size=11, color=SUB),
            label("lblPcStat", F('Upper(If(StartsWith(ThisItem.StatusPen, "Bloqueado"), "Bloqueada", ThisItem.StatusPen = "Posição vazia", "Vazia", "Liberada"))'),
                  "Parent.TemplateWidth - 196", 8, 180, 30, size=10, bold=True, Align="Align.Center",
                  Fill=status_fill("ThisItem.StatusPen"), Color=status_color("ThisItem.StatusPen")),
            label("lblPcMot", F('Trim(Last(Split(ThisItem.StatusPen, "·")).Value)'), "Parent.TemplateWidth - 236", 40, 220, 26,
                  size=10, color=SUB, Align="Align.Right"),
        ], Visible=f"{vivo} && gblEhPeneirador"),
        label("lblEqHistVazio", '"Nenhuma aferição registrada para este equipamento."', 20, 540, "Parent.Width - 40", 60, size=12, color=SUB,
              Align="Align.Center", Visible=f'{vivo} && CountRows(colHist) = 0 && !gblEhPeneirador && gblPerfil <> "Instrumentação"'),
        # rodapé por perfil
        rect("recEqRodape", 0, "Parent.Height - 104", "Parent.Width", 104, fill=WHITE, BorderColor=LINE, BorderThickness=1, Visible=vivo),
        button("btnEqLancar", F(f'If(CountRows(gblCodigos) >= 2, {texto_lancar("First(gblCodigos).Value", True)}, {texto_lancar("First(gblCodigos).Value", False)})'),
               20, "Parent.Height - 88", F('If(CountRows(gblCodigos) >= 2, (Parent.Width - 50) / 2, Parent.Width - 40)'), 70,
               acao_lancar("First(gblCodigos).Value"), size=F('If(CountRows(gblCodigos) >= 2, 13, 15)'), Visible=lancar_visivel),
        button("btnEqLancar2", F(texto_lancar("Last(gblCodigos).Value", True)),
               "30 + (Parent.Width - 50) / 2", "Parent.Height - 88", "(Parent.Width - 50) / 2", 70,
               acao_lancar("Last(gblCodigos).Value"), size=13, Visible=lancar2_visivel),
        button("btnEqCalibrar", '"Registrar calibração"', 20, "Parent.Height - 88", F('If(gblEhPeneira, (Parent.Width - 50) / 2, Parent.Width - 40)'), 70,
               "Navigate(scrCalibrar, ScreenTransition.Cover)", Visible=f'{vivo} && gblPerfil = "Instrumentação"'),
        button("btnEqInstalar", '"Instalar em posição"', "30 + (Parent.Width - 50) / 2", "Parent.Height - 88", "(Parent.Width - 50) / 2", 70,
               "UpdateContext({locInstalar: true}); Reset(drpEqPosicao)", fill=WHITE, color=NAVY, BorderColor=NAVY, BorderThickness=2,
               Visible=f'{vivo} && gblPerfil = "Instrumentação" && gblEhPeneira'),
        label("lblEqConsulta", F('If(gblPerfil = "Laboratorista", "Este equipamento não tem aferição semanal.", "Modo consulta: só a equipe autorizada registra.")'),
              20, "Parent.Height - 88", "Parent.Width - 40", 70, size=12, color=SUB, Align="Align.Center",
              Visible=f'{vivo} && (gblPerfil = "Colaborador" || (gblPerfil = "Laboratorista" && !({lancar_visivel})))'),
        # painel: instalar peneira em uma posição (instrumentação)
        rect("recEqInsFundo", 0, 0, "Parent.Width", "Parent.Height", fill=OVERLAY, Visible="locInstalar"),
        rect("recEqInsCartao", 20, 300, "Parent.Width - 40", 470, fill=WHITE, Visible="locInstalar"),
        label("lblEqInsTitulo", F('"Instalar " & gblEquip.Title & " (" & gblEquip.Descricao & ")"'), 44, 320, "Parent.Width - 88", 40, size=17, bold=True, color=NAVY, Visible="locInstalar"),
        label("lblEqInsTexto", F('''
            "Escolha a posição onde a peneira foi colocada. A peneira que estava nessa posição fica como fora de uso." &
            If(IsBlank(gblEquip.Conjunto), "", Char(10) & "Hoje ela está no " & Coalesce(LookUp(colEquip, Title = gblEquip.Conjunto).Descricao, gblEquip.Conjunto) & ".")'''),
              44, 364, "Parent.Width - 88", 96, size=12, color=SUB, VerticalAlign="VerticalAlign.Top", Visible="locInstalar"),
        label("lblEqInsRot", '"POSIÇÃO (PENEIRADOR · MALHA)"', 44, 468, "Parent.Width - 88", 26, size=10, bold=True, color=SUB, Visible="locInstalar"),
        dropdown("drpEqPosicao", F(f'''
            With(
                {{wMalha: Substitute(gblEquip.Descricao, "Peneira ", "Malha ")}},
                With(
                    {{wMesma: Filter(colPontos, Ensaio = "PEN" && wMalha in Parametro)}},
                    ForAll(If(CountRows(wMesma) > 0, wMesma, Filter(colPontos, Ensaio = "PEN")), {pos_item})
                )
            )'''), 44, 494, "Parent.Width - 88", 54, Visible="locInstalar"),
        label("lblEqInsNota", '"Mostrando as posições da mesma malha. O histórico de verificações continua ligado a cada peneira."',
              44, 556, "Parent.Width - 88", 60, size=11, color=SUB, Visible="locInstalar"),
        button("btnEqInsCancelar", '"Cancelar"', 44, 680, "(Parent.Width - 100) / 2", 66, "UpdateContext({locInstalar: false})",
               fill=WHITE, color=NAVY, BorderColor=NAVY, BorderThickness=2, Visible="locInstalar"),
        button("btnEqInsConfirmar", '"Confirmar"', "56 + (Parent.Width - 100) / 2", 680, "(Parent.Width - 100) / 2", 66, instalar,
               Visible="locInstalar"),
    ]
    return {"scrEquipamento": {"Properties": {"Fill": f"={BG}", "OnVisible": onvisible}, "Children": filhos}}



# ------------------------------------------------------------------ TELA 3 · Lançamento por medição
def tela_ensaio():
    p = "En"
    onvisible = F(f'''
        UpdateContext({{locRevisar: false, locSalvando: false}});
        Reset(dtp{p}Data); Reset(drp{p}Turno); Reset(drp{p}Causa); Reset(txt{p}Just); Reset(txt{p}Obs);
        Reset(txt{p}Estufa); Reset(txt{p}Tara); Reset(txt{p}PIni); Reset(txt{p}PFin);
        // Recriar a coleção zera os campos digitados da galeria
        ClearCollect(
            colForm,
            AddColumns(
                Sort(Filter(colPontos, Ensaio = gblEnsaio.Title && (IsBlank(gblDestacar) || IsBlank(TAG) || TAG = gblDestacar)), Ordem),
                Nom, LookUp(colLimites, Title = ChaveLimite).Nominal,
                LI, LookUp(colLimites, Title = ChaveLimite).LimInferior,
                LS, LookUp(colLimites, Title = ChaveLimite).LimSuperior,
                Tol, LookUp(colLimites, Title = ChaveLimite).Tolerancia
            )
        )''')
    umi = f'gblEnsaio.Title = "UMI"'
    y_gal = F(f'If({umi}, 384, 268)')
    criterio = '''
        If(
            ThisItem.EntradaTxt = "Massa", ThisItem.RotuloA & " e " & ThisItem.RotuloB,
            !IsBlank(ThisItem.LI) || !IsBlank(ThisItem.LS),
                "Limites " & Coalesce(Text(ThisItem.LI), "—") & " a " & Coalesce(Text(ThisItem.LS), "—") & " " & ThisItem.Unidade,
            !IsBlank(ThisItem.Tol), "Diferença até ± " & Text(ThisItem.Tol) & " " & ThisItem.Unidade,
            "Informativo"
        )'''
    filhos = cabecalho(p, "gblEnsaio.Nome", F('gblEnsaio.Frequencia & " · lançamento"'),
                       voltar="Navigate(scrEquipamento, ScreenTransition.UnCover)") + topo_lancamento(p) + [
        # bloco da estufa (umidade)
        label(f"lbl{p}EstufaRot", '"ESTUFA 100 °C · UMIDADE (%) OU PESAGENS (g)"', 20, 266, "Parent.Width - 40", 26, size=10, bold=True, color=SUB, Visible=umi),
        textinput(f"txt{p}Estufa", 20, 292, "(Parent.Width - 70) / 4", 52, '"Estufa %"', Visible=umi),
        textinput(f"txt{p}Tara", "30 + (Parent.Width - 70) / 4", 292, "(Parent.Width - 70) / 4", 52, '"Tara"', Visible=umi),
        textinput(f"txt{p}PIni", "40 + 2 * (Parent.Width - 70) / 4", 292, "(Parent.Width - 70) / 4", 52, '"P. inicial"', Visible=umi),
        textinput(f"txt{p}PFin", "50 + 3 * (Parent.Width - 70) / 4", 292, "(Parent.Width - 70) / 4", 52, '"P. final"', Visible=umi),
        label(f"lbl{p}EstufaValor", F(f'''
            With(
                {{e: Value(txt{p}Estufa.Text), t: Value(txt{p}Tara.Text), pi: Value(txt{p}PIni.Text), pf: Value(txt{p}PFin.Text)}},
                With(
                    {{x: If(!IsBlank(e), e, !IsBlank(t) && !IsBlank(pi) && !IsBlank(pf) && pi > t, (pi - pf) / (pi - t) * 100)}},
                    If(IsBlank(x), "", Text(x, "0.##########", "en-US"))
                )
            )'''), 0, 0, 10, 10, Visible=False),
        label(f"lbl{p}EstufaInfo", F(f'If(IsBlank(lbl{p}EstufaValor.Text), "Informe a estufa ou as três pesagens", "Estufa usada: " & Text(Value(lbl{p}EstufaValor.Text, "en-US"), "0.00") & " %")'),
              20, 346, "Parent.Width - 40", 30, size=12, bold=True, color=NAVY, Visible=umi),
        gallery(f"gal{p}Pontos", "colForm", 20, y_gal, "Parent.Width - 40", F(f'Parent.Height - ({y_gal[1:]}) - 112'), 176, [
            rect("recPtCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 10", fill=WHITE,
                 BorderColor=F(f'If(!IsBlank(gblDestacar) && ThisItem.TAG = gblDestacar, {ACCENT}, {LINE})'),
                 BorderThickness=F('If(!IsBlank(gblDestacar) && ThisItem.TAG = gblDestacar, 3, 1)')),
            label("lblPtTitulo", F('''
                Switch(
                    ThisItem.EntradaTxt,
                    "Massa", "Fração " & ThisItem.Parametro,
                    "Comparativo", ThisItem.Parametro,
                    Coalesce(ThisItem.TAG, ThisItem.EquipamentoBD) & If(IsBlank(ThisItem.Secao), "", " · " & ThisItem.Secao)
                )'''), 14, 8, "Parent.TemplateWidth - 200", 32, size=14, bold=True, color=NAVY),
            label("lblPtSub", F(f'''
                If(ThisItem.EntradaTxt in ["Massa", "Comparativo"], ThisItem.EquipamentoBD, ThisItem.Parametro) &
                If(ThisItem.EntradaTxt = "Tempo", " · " & ThisItem.Fator & " voltas", "") & " · " &
                {textwrap.indent(textwrap.dedent(criterio).strip(), "                ").strip()}'''),
                  14, 40, "Parent.TemplateWidth - 28", 30, size=11, color=SUB),
            textinput("txtPtA", 14, 76, F('If(ThisItem.EntradaTxt in ["Tempo", "Comparativo", "Massa"], (Parent.TemplateWidth - 38) / 2, Parent.TemplateWidth - 28)'), 54,
                      F('Coalesce(ThisItem.RotuloA, "Valor")'), size=17, Format="TextFormat.Text", Align="Align.Center", FontWeight="FontWeight.Bold"),
            textinput("txtPtB", F('24 + (Parent.TemplateWidth - 38) / 2'), 76, F('(Parent.TemplateWidth - 38) / 2'), 54,
                      F('Coalesce(ThisItem.RotuloB, "Valor")'), size=17, Align="Align.Center", FontWeight="FontWeight.Bold",
                      Visible=F('ThisItem.EntradaTxt in ["Tempo", "Comparativo", "Massa"]')),
            label("lblPtV", F('''
                With(
                    {a: Value(txtPtA.Text), b: Value(txtPtB.Text)},
                    With(
                        {
                            v: Switch(
                                ThisItem.EntradaTxt,
                                "Tempo", If(Coalesce(a, 0) * 60 + Coalesce(b, 0) > 0 && !IsBlank(ThisItem.Fator), ThisItem.Fator * 60 / (Coalesce(a, 0) * 60 + Coalesce(b, 0))),
                                "Massa", Blank(),
                                a
                            )
                        },
                        If(IsBlank(v), "", Text(v, "0.##########", "en-US"))
                    )
                )'''), 0, 0, 10, 10, Visible=False),
            label("lblPtRef", F(f'''
                With(
                    {{r: Switch(ThisItem.EntradaTxt, "Comparativo", Value(txtPtB.Text), "ComparativoReferencia", Value(lbl{p}EstufaValor.Text, "en-US"), ThisItem.Nom)}},
                    If(IsBlank(r), "", Text(r, "0.##########", "en-US"))
                )'''), 0, 0, 10, 10, Visible=False),
            label("lblPtRes", F('''
                With(
                    {v: Value(lblPtV.Text, "en-US"), r: Value(lblPtRef.Text, "en-US")},
                    If(
                        IsBlank(v), "",
                        !IsBlank(ThisItem.LI) || !IsBlank(ThisItem.LS),
                            If((IsBlank(ThisItem.LI) || v >= ThisItem.LI) && (IsBlank(ThisItem.LS) || v <= ThisItem.LS), "Conforme", "Não conforme"),
                        !IsBlank(ThisItem.Tol),
                            If(IsBlank(r), "", If(Round(Abs(v - r), 6) <= Round(ThisItem.Tol, 6), "Conforme", "Não conforme")),
                        "Informativo"
                    )
                )'''), "Parent.TemplateWidth - 186", 10, 172, 30, size=10, bold=True, Align="Align.Center",
                  Fill=res_fill("Self.Text"), Color=res_color("Self.Text"), Visible=F('!IsBlank(Self.Text)')),
            label("lblPtCalc", F('''
                With(
                    {v: Value(lblPtV.Text, "en-US"), r: Value(lblPtRef.Text, "en-US")},
                    If(
                        IsBlank(v), "",
                        If(ThisItem.EntradaTxt = "Tempo", "Rotação " & Text(v, "0.00") & " rpm   ", "") &
                        If(IsBlank(r), "", "Diferença " & Text(v - r, "0.###"))
                    )
                )'''), 14, 134, "Parent.TemplateWidth - 28", 30, size=12, bold=True, color=TXT),
        ]),
    ] + barra_e_revisao(p, revisar_ensaio(p), F(f'''
        CountRows(Filter(gal{p}Pontos.AllItems, !IsBlank(txtPtA.Text))) & " de " & CountRows(colForm) & " campos preenchidos" & Char(10) &
        CountRows(Filter(gal{p}Pontos.AllItems, lblPtRes.Text = "Não conforme")) & " não conforme(s) até aqui"'''))
    return {"scrEnsaio": {"Properties": {"Fill": f"={BG}", "OnVisible": onvisible}, "Children": filhos}}


def revisar_ensaio(p):
    return F(f'''
        ClearCollect(
            colEntradas,
            ForAll(
                gal{p}Pontos.AllItems,
                {{
                    Ordem: Ordem, TAG: TAG, EquipamentoBD: EquipamentoBD, Parametro: Parametro, Unidade: Unidade,
                    Entrada: EntradaTxt, Tipo: TipoTxt, Fator: Fator, LI: LI, LS: LS, Tol: Tol, Nom: Nom,
                    A: Value(txtPtA.Text), B: Value(txtPtB.Text),
                    V: Value(lblPtV.Text, "en-US"), R: Value(lblPtRef.Text, "en-US"), Res: lblPtRes.Text
                }}
            )
        );
        ClearCollect(
            colResultados,
            ForAll(
                Filter(colEntradas, Entrada <> "Massa" && !IsBlank(V)),
                {{
                    Equipamento: EquipamentoBD, TAG: TAG, Parametro: Parametro, Unidade: Unidade, Tipo: Tipo,
                    Valor: V, Referencia: R, Diferenca: If(!IsBlank(R), V - R), LI: LI, LS: LS, Tol: Tol,
                    Resultado: Res, ObsItem: "", ItensNaoOK: ""
                }}
            )
        );
        // Granulometria LTF x LCE: indicadores calculados a partir das massas (mesmas fórmulas da planilha)
        If(
            gblEnsaio.Title = "GRA" && Sum(colEntradas, Coalesce(A, 0)) > 0 && Sum(colEntradas, Coalesce(B, 0)) > 0,
            With(
                {{tA: Sum(colEntradas, Coalesce(A, 0)), tB: Sum(colEntradas, Coalesce(B, 0)), r0: First(colEntradas)}},
                With(
                    {{
                        pA: AddColumns(colEntradas, P, Coalesce(A, 0) / tA * 100),
                        pB: AddColumns(colEntradas, P, Coalesce(B, 0) / tB * 100)
                    }},
                    With(
                        {{
                            a63: Sum(Filter(pA, Ordem >= 8), P), b63: Sum(Filter(pB, Ordem >= 8), P),
                            a34: Sum(Filter(pA, Ordem = 3 || Ordem = 4), P), b34: Sum(Filter(pB, Ordem = 3 || Ordem = 4), P),
                            a5: LookUp(pA, Ordem = 5).P, b5: LookUp(pB, Ordem = 5).P,
                            a6: LookUp(pA, Ordem = 6).P, b6: LookUp(pB, Ordem = 6).P,
                            aDm: Sum(pA, P * Coalesce(Fator, 0)) / 100, bDm: Sum(pB, P * Coalesce(Fator, 0)) / 100,
                            L63: LookUp(colLimites, Title = "G63"), LRG: LookUp(colLimites, Title = "GRG")
                        }},
                        Collect(
                            colResultados,
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "-6,3 mm", Unidade: "%",
                              Valor: a63, Referencia: b63, Diferenca: a63 - b63, Tol: L63.Tolerancia,
                              Resultado: If(Round(Abs(a63 - b63), 6) <= Round(L63.Tolerancia, 6), "Conforme", "Não conforme"), ObsItem: "", ItensNaoOK: ""}},
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "RG (Relação)", Unidade: "-",
                              Valor: a34 / a5, Referencia: b34 / b5, Diferenca: a34 / a5 - b34 / b5, Tol: LRG.Tolerancia,
                              Resultado: If(Round(Abs(a34 / a5 - b34 / b5), 6) <= Round(LRG.Tolerancia, 6), "Conforme", "Não conforme"), ObsItem: "", ItensNaoOK: ""}},
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "-16,0 +12,5 mm", Unidade: "%",
                              Valor: a34, Referencia: b34, Diferenca: a34 - b34, Resultado: "Informativo", ObsItem: "", ItensNaoOK: ""}},
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "-16,0 +9,0 mm", Unidade: "%",
                              Valor: a34 + a5, Referencia: b34 + b5, Diferenca: a34 + a5 - b34 - b5, Resultado: "Informativo", ObsItem: "", ItensNaoOK: ""}},
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "-16,0 +8,0 mm", Unidade: "%",
                              Valor: a34 + a5 + a6, Referencia: b34 + b5 + b6, Diferenca: a34 + a5 + a6 - b34 - b5 - b6, Resultado: "Informativo", ObsItem: "", ItensNaoOK: ""}},
                            {{Equipamento: r0.EquipamentoBD, TAG: r0.TAG, Tipo: r0.Tipo, Parametro: "Diâmetro médio", Unidade: "mm",
                              Valor: aDm, Referencia: bDm, Diferenca: aDm - bDm, Resultado: "Informativo", ObsItem: "", ItensNaoOK: ""}}
                        )
                    )
                )
            )
        );
        // Tamboramento 5 kg x 15 kg: percentuais pelas massas de ensaio da lista Configuracoes
        If(
            gblEnsaio.Title = "T515",
            With(
                {{
                    mA: Value(LookUp(colConfig, Title = "T515MassaLTF").Valor), mB: Value(LookUp(colConfig, Title = "T515MassaLCE").Valor),
                    r1: LookUp(colEntradas, Ordem = 1), r2: LookUp(colEntradas, Ordem = 2), L: LookUp(colLimites, Title = "T515")
                }},
                If(
                    !IsBlank(r1.A) && !IsBlank(r1.B),
                    With(
                        {{a63: r1.A / (mA / 100), b63: r1.B / (mB / 100)}},
                        Collect(
                            colResultados,
                            {{Equipamento: r1.EquipamentoBD, TAG: r1.TAG, Tipo: r1.Tipo, Parametro: "+6,3 mm", Unidade: "%",
                              Valor: a63, Referencia: b63, Diferenca: a63 - b63, Tol: L.Tolerancia,
                              Resultado: If(Round(Abs(a63 - b63), 6) <= Round(L.Tolerancia, 6), "Conforme", "Não conforme"), ObsItem: "", ItensNaoOK: ""}}
                        );
                        If(
                            !IsBlank(r2.A) && !IsBlank(r2.B),
                            With(
                                {{a05: 100 - (a63 + r2.A / (mA / 100)), b05: 100 - (b63 + r2.B / (mB / 100))}},
                                Collect(
                                    colResultados,
                                    {{Equipamento: r1.EquipamentoBD, TAG: r1.TAG, Tipo: r1.Tipo, Parametro: "-0,5 mm", Unidade: "%",
                                      Valor: a05, Referencia: b05, Diferenca: a05 - b05, Resultado: "Informativo", ObsItem: "", ItensNaoOK: ""}}
                                )
                            )
                        )
                    )
                )
            )
        );
        Set(gblNC, CountRows(Filter(colResultados, Resultado = "Não conforme")));
        If(
            CountRows(colResultados) = 0,
            Notify("Preencha pelo menos um resultado antes de enviar.", NotificationType.Error),
            {umi_check(p)}
            UpdateContext({{locRevisar: true}})
        )''')


def umi_check(p):
    return f'''gblEnsaio.Title = "UMI" && IsBlank(lbl{p}EstufaValor.Text),
            Notify("Informe a umidade da estufa (ou as três pesagens).", NotificationType.Error),'''


# ------------------------------------------------------------------ TELA 4 · Verificação dos peneiradores (checklist)
ITENS_PEN = [("I1", "Telas", "Condições das telas"), ("I2", "Inclinação", "Inclinação das telas"),
             ("I3", "Vibração", "Vibração do peneirador"), ("I4", "Travamento", "Travamento das peneiras"),
             ("I5", "Calibração", "Calibração")]
NC_PEN = '(I1 = "NÃO OK" || I2 = "NÃO OK" || I3 = "NÃO OK" || I4 = "NÃO OK" || I5 = "NÃO OK")'


def tela_verificacao():
    p = "Ve"
    onvisible = F(f'''
        UpdateContext({{
            locRevisar: false, locSalvando: false, locSoNC: false,
            locGrupo: "Todos",
            locConjunto: gblConjuntoFiltro
        }});
        Reset(dtp{p}Data); Reset(drp{p}Turno); Reset(drp{p}Causa); Reset(txt{p}Just); Reset(txt{p}Obs);
        If(CountRows(colEquip) = 0, {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "    ").strip()});
        // Todas as peneiras começam OK; a calibração já vem NÃO OK quando a peneira está vencida no cadastro
        ClearCollect(
            colFormVe,
            AddColumns(
                Sort(Filter(colPontos, Ensaio = gblEnsaio.Title && !IsBlank(TAG) && (IsBlank(gblConjuntoFiltro) || Conjunto = gblConjuntoFiltro)), Ordem),
                Vencida,
                With(
                    {{e: LookUp(colEquip, Title = TAG)}},
                    !IsBlank(e) && e.ControlaCalibracao && !IsBlank(e.ProximaCalibracao) && e.ProximaCalibracao < Today()
                ),
                I1, "OK", I2, "OK", I3, "OK", I4, "OK", I5, "OK"
            )
        );
        UpdateIf(colFormVe, Vencida, {{I5: "NÃO OK"}});
        Set(gblConjuntoFiltro, "")''')
    grupo_filtro = '(locGrupo = "Todos" || Secao = locGrupo)'
    chips = []
    larg = "(Parent.TemplateWidth - 28 - 4 * 8) / 5"
    for i, (campo, rot, _) in enumerate(ITENS_PEN):
        ciclo = ('"OK", "NÃO OK", "NÃO OK", "SEM TAG", "SEM TAG", "N/A", "OK"' if campo == "I5"
                 else '"OK", "NÃO OK", "NÃO OK", "N/A", "OK"')
        chips.append(button(f"btnVg{campo}", F(f'"{rot.upper()}" & Char(10) & ThisItem.{campo}'),
                            f"14 + {i} * ({larg} + 8)", 72, larg, 56,
                            F(f'Patch(colFormVe, ThisItem, {{{campo}: Switch(ThisItem.{campo}, {ciclo})}})'),
                            fill=F(f'Switch(ThisItem.{campo}, "OK", {OK_S}, "NÃO OK", {BAD}, {MUT_S})'),
                            color=F(f'Switch(ThisItem.{campo}, "OK", {OK}, "NÃO OK", {WHITE}, {MUT})'),
                            size=9, **radius(8)))
    filhos_galeria = [
        rect("recVgCard", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight - 8", fill=WHITE,
             BorderColor=F(f'If(!IsBlank(gblDestacar) && ThisItem.TAG = gblDestacar, {ACCENT}, {LINE})'),
             BorderThickness=F('If(!IsBlank(gblDestacar) && ThisItem.TAG = gblDestacar, 3, 1)')),
        label("lblVgTag", "ThisItem.TAG", 14, 6, "Parent.TemplateWidth - 210", 32, size=15, bold=True, color=NAVY),
        label("lblVgSub", F('Trim(First(Split(ThisItem.Parametro, "(")).Value) & " · " & ThisItem.Secao & If(ThisItem.Vencida, "  ·  calibração vencida", "")'),
              14, 36, "Parent.TemplateWidth - 28", 28, size=11, color=F(f'If(ThisItem.Vencida, {BAD}, {SUB})')),
        label("lblVgRes", F(f'If({NC_PEN.replace("I", "ThisItem.I")}, "Não conforme", "Conforme")'),
              "Parent.TemplateWidth - 186", 8, 172, 28, size=10, bold=True, Align="Align.Center",
              Fill=res_fill("Self.Text"), Color=res_color("Self.Text")),
    ] + chips

    revisar = F(f'''
        ClearCollect(
            colResultados,
            ForAll(
                Filter(colFormVe, {grupo_filtro}) As pe,
                With(
                    {{
                        itens: Table(
                            {', '.join(f'{{n: "{nome}", v: pe.{campo}}}' for campo, _, nome in ITENS_PEN)}
                        )
                    }},
                    {{
                        Equipamento: pe.EquipamentoBD, TAG: pe.TAG, Parametro: pe.Parametro, Unidade: "itens OK", Tipo: pe.TipoTxt,
                        Valor: CountRows(Filter(itens, v = "OK")), Referencia: CountRows(itens),
                        Diferenca: If(false, 0), LI: If(false, 0), LS: If(false, 0), Tol: If(false, 0),
                        Resultado: If(CountRows(Filter(itens, v = "NÃO OK")) > 0, "Não conforme", "Conforme"),
                        ItensNaoOK: Concat(Filter(itens, v = "NÃO OK"), n, "; "),
                        ObsItem: Concat(
                            Filter(
                                Table(
                                    {{t: Concat(Filter(itens, v = "NÃO OK"), n, "; ")}},
                                    {{t: If(CountRows(Filter(itens, v = "SEM TAG")) > 0, "Sem TAG de calibração", "")}}
                                ),
                                !IsBlank(t)
                            ),
                            t,
                            "; "
                        )
                    }}
                )
            )
        );
        Set(gblNC, CountRows(Filter(colResultados, Resultado = "Não conforme")));
        UpdateContext({{locRevisar: true}})''')

    filhos = cabecalho(p, "gblEnsaio.Nome", F('gblEnsaio.Frequencia & " · " & CountRows(colFormVe) & " peneiras"'),
                       voltar="Navigate(scrEquipamento, ScreenTransition.UnCover)") + topo_lancamento(p) + [
        # peneirador lido no QR
        label(f"lbl{p}Conjunto", F(f'''
            "PENEIRADOR  ·  " & Upper(Coalesce(LookUp(colEquip, Title = locConjunto).Descricao, First(colFormVe).Secao, locConjunto)) &
            "  ·  " & CountRows(colFormVe) & " peneiras" &
            With({{n: CountRows(Filter(colFormVe, {NC_PEN}))}}, If(n > 0, "  ·  " & n & " NC", ""))'''),
              20, 266, "Parent.Width - 40", 54, size=12, bold=True, color=NAVY, Fill=LAB_S, PaddingLeft=14),
        button(f"btn{p}TodosOK", F('"✔  Marcar todos OK" & If(locGrupo = "Todos", "", " (" & locGrupo & ")")'), 20, 332,
               "(Parent.Width - 50) * 0.62", 54, F(f'''
            UpdateIf(colFormVe, {grupo_filtro}, {{I1: "OK", I2: "OK", I3: "OK", I4: "OK", I5: If(Vencida, "NÃO OK", "OK")}});
            Notify("Itens marcados como OK. Peneiras com calibração vencida continuam com Calibração = NÃO OK.", NotificationType.Information)'''),
               fill=OK, color=WHITE, size=13),
        button(f"btn{p}SoNC", F('If(locSoNC, "Mostrar todas", "Só não conformes")'), "30 + (Parent.Width - 50) * 0.62", 332,
               "(Parent.Width - 50) * 0.38", 54, "UpdateContext({locSoNC: !locSoNC})",
               fill=WHITE, color=NAVY, size=12, BorderColor=NAVY, BorderThickness=1),
        label(f"lbl{p}Legenda", '"Toque no item para alternar: OK → NÃO OK → N/A.  Calibração: OK → NÃO OK → SEM TAG → N/A."',
              20, 390, "Parent.Width - 40", 40, size=10, color=SUB),
        gallery(f"gal{p}Peneiras", F(f'Filter(colFormVe, {grupo_filtro} && (!locSoNC || {NC_PEN}))'),
                20, 432, "Parent.Width - 40", "Parent.Height - 432 - 112", 138, filhos_galeria),
        label(f"lbl{p}Vazio", F('If(locSoNC, "Nenhuma peneira não conforme neste filtro.", "Nenhuma peneira cadastrada para este ensaio.")'),
              20, 460, "Parent.Width - 40", 60, size=12, color=SUB, Align="Align.Center", Visible=F(f'CountRows(gal{p}Peneiras.AllItems) = 0')),
    ] + barra_e_revisao(p, revisar, F(f'''
        CountRows(Filter(colFormVe, {grupo_filtro})) & " peneiras neste envio" & If(locGrupo = "Todos", "", " (" & locGrupo & ")") & Char(10) &
        CountRows(Filter(colFormVe, {grupo_filtro} && {NC_PEN})) & " não conforme(s)"'''))
    return {"scrVerificacao": {"Properties": {"Fill": f"={BG}", "OnVisible": onvisible}, "Children": filhos}}


# ------------------------------------------------------------------ TELA 5 · Calibração (instrumentação)
def tela_calibrar():
    p = "Ca"
    onvisible = F(f'''
        UpdateContext({{locResultado: "", locSalvando: false}});
        Reset(dtp{p}Data); Reset(txt{p}Resp); Reset(txt{p}Local); Reset(drp{p}Motivo); Reset(txt{p}MotivoOutro);
        Reset(chk{p}Manut); Reset(txt{p}Cert); Reset(txt{p}Obs)''')
    rot = lambda n, t, y, vis=None: label(n, f'"{t}"', 20, y, "Parent.Width - 40", 26, size=10, bold=True, color=SUB, Visible=vis)
    rec = 'locResultado = "Recusada"'
    salvar_cal = F(f'''
        If(
            gblPerfil <> "Instrumentação",
                Notify("Somente a instrumentação registra calibrações.", NotificationType.Error),
            IsBlank(dtp{p}Data.SelectedDate) || dtp{p}Data.SelectedDate > Today(),
                Notify("Informe uma data de calibração válida (não pode ser futura).", NotificationType.Error),
            IsBlank(Trim(txt{p}Resp.Text)),
                Notify("Informe o responsável pela calibração.", NotificationType.Error),
            IsBlank(locResultado),
                Notify("Escolha o resultado: Aprovada ou Recusada.", NotificationType.Error),
            {rec} && IsBlank(drp{p}Motivo.Selected.Value),
                Notify("Selecione o motivo da recusa.", NotificationType.Error),
            {rec} && drp{p}Motivo.Selected.Value = "Outro" && IsBlank(Trim(txt{p}MotivoOutro.Text)),
                Notify("Descreva o motivo da recusa.", NotificationType.Error),

            UpdateContext({{locSalvando: true}});
            With(
                {{
                    wProx: DateAdd(dtp{p}Data.SelectedDate, Coalesce(gblEquip.PeriodicidadeMeses, 12), TimeUnit.Months),
                    wMotivo: If({rec}, If(drp{p}Motivo.Selected.Value = "Outro", Trim(txt{p}MotivoOutro.Text), drp{p}Motivo.Selected.Value), ""),
                    wManut: {rec} && chk{p}Manut.Value
                }},
                With(
                    {{
                        wStatus: If(
                            wManut, "Bloqueado · Em manutenção",
                            {rec}, "Bloqueado · Calibração recusada",
                            wProx < Today(), "Bloqueado · Calibração vencida",
                            wProx <= DateAdd(Today(), gblDiasAviso, TimeUnit.Days), "Liberado · Calibração vence em breve",
                            "Liberado"
                        )
                    }},
                    IfError(
                        Patch(
                            Calibracoes,
                            Defaults(Calibracoes),
                            {{
                                Title: gblEquip.Title,
                                EquipamentoID: gblEquip.ID,
                                DataCalibracao: dtp{p}Data.SelectedDate,
                                Responsavel: Trim(txt{p}Resp.Text),
                                Local: Trim(txt{p}Local.Text),
                                Resultado: locResultado,
                                MotivoRecusa: wMotivo,
                                ManutencaoSolicitada: wManut,
                                ProximaCalibracao: wProx,
                                Certificado: Trim(txt{p}Cert.Text),
                                Observacoes: Trim(txt{p}Obs.Text),
                                ModoLeitura: gblModoLeitura
                            }}
                        ),
                        Notify("Erro ao gravar a calibração: " & FirstError.Message, NotificationType.Error),

                        Patch(
                            Equipamentos,
                            gblEquip,
                            {{
                                UltimaCalibracao: dtp{p}Data.SelectedDate,
                                ProximaCalibracao: wProx,
                                UltimoResultado: locResultado,
                                UltimoResponsavel: Trim(txt{p}Resp.Text),
                                UltimoCertificado: Trim(txt{p}Cert.Text),
                                MotivoRecusa: wMotivo,
                                ManutencaoSolicitada: wManut,
                                Local: Coalesce(Trim(txt{p}Local.Text), gblEquip.Local),
                                Status: wStatus
                            }}
                        ),
                        Notify("Calibração gravada, mas o equipamento não foi atualizado: " & FirstError.Message, NotificationType.Error),

                        Set(gblEquip, LookUp(Equipamentos, ID = gblEquip.ID));
                        {textwrap.indent(textwrap.dedent(CARREGAR_EQUIP).strip(), "                        ").strip()};
                        Notify("Calibração registrada. Próxima em " & Text(wProx, "dd/mm/yyyy") & ".", NotificationType.Success);
                        Navigate(scrEquipamento, ScreenTransition.UnCover)
                    )
                )
            );
            UpdateContext({{locSalvando: false}})
        )''')
    filhos = cabecalho(p, '"Registrar calibração"', F('gblEquip.Title & " · " & gblEquip.Descricao'),
                       voltar="Navigate(scrEquipamento, ScreenTransition.UnCover)") + [
        rot(f"lbl{p}DataRot", "DATA DA CALIBRAÇÃO *", 124),
        ctl(f"dtp{p}Data", "Classic/DatePicker@2.6.0", DefaultDate="Today()", Format="DateTimeFormat.ShortDate", X=20, Y=150,
            Width="Parent.Width - 40", Height=52, Size=14),
        rot(f"lbl{p}RespRot", "RESPONSÁVEL *", 210),
        textinput(f"txt{p}Resp", 20, 236, "Parent.Width - 40", 52, '"Nome do responsável"', default="gblNome"),
        rot(f"lbl{p}LocalRot", "LOCAL DE INSTALAÇÃO", 296),
        textinput(f"txt{p}Local", 20, 322, "Parent.Width - 40", 52, '"Área / setor"', default="gblEquip.Local"),
        rot(f"lbl{p}ResRot", "RESULTADO DA CALIBRAÇÃO *", 382),
        button(f"btn{p}Aprovada", '"✔  APROVADA"', 20, 408, "(Parent.Width - 50) / 2", 60,
               'UpdateContext({locResultado: "Aprovada"}); Reset(drpCaMotivo); Reset(txtCaMotivoOutro); Reset(chkCaManut)',
               fill=F(f'If(locResultado = "Aprovada", {OK}, {WHITE})'), color=F(f'If(locResultado = "Aprovada", {WHITE}, {OK})'),
               BorderColor=OK, BorderThickness=2),
        button(f"btn{p}Recusada", '"✖  RECUSADA"', "30 + (Parent.Width - 50) / 2", 408, "(Parent.Width - 50) / 2", 60,
               'UpdateContext({locResultado: "Recusada"})',
               fill=F(f'If(locResultado = "Recusada", {BAD}, {WHITE})'), color=F(f'If(locResultado = "Recusada", {WHITE}, {BAD})'),
               BorderColor=BAD, BorderThickness=2),
        rot(f"lbl{p}MotivoRot", "MOTIVO DA RECUSA *", 478, rec),
        dropdown(f"drp{p}Motivo", '["Fora da tolerância", "Desvio na linearidade", "Repetibilidade insatisfatória", "Tela danificada ou desgastada", "Falha mecânica", "Falha eletrônica", "Outro"]',
                 20, 504, "Parent.Width - 40", 52, default='""', AllowEmptySelection=True, Visible=rec),
        textinput(f"txt{p}MotivoOutro", 20, 562, "Parent.Width - 40", 52, '"Descreva o motivo"',
                  Visible=F(f'{rec} && drp{p}Motivo.Selected.Value = "Outro"')),
        ctl(f"chk{p}Manut", "Classic/CheckBox@2.1.0", Text='"Enviar para manutenção"', Default=False, X=20, Y=620,
            Width="Parent.Width - 40", Height=48, Size=14, Visible=rec),
        rot(f"lbl{p}CertRot", "Nº DO CERTIFICADO · VAI NA IMAGEM DE CONSULTA", 676),
        textinput(f"txt{p}Cert", 20, 702, "Parent.Width - 40", 52, '"Ex.: CERT-2026-0154"'),
        rot(f"lbl{p}ObsRot", "OBSERVAÇÕES · O PRÓXIMO TÉCNICO VÊ NA FICHA", 762),
        textinput(f"txt{p}Obs", 20, 788, "Parent.Width - 40", 84, '"Ex.: calibrada, mas o pé de nivelamento está solto"', Mode="TextMode.MultiLine", size=13),
        label(f"lbl{p}Prox", F(f'"Próxima calibração: " & Text(DateAdd(dtp{p}Data.SelectedDate, Coalesce(gblEquip.PeriodicidadeMeses, 12), TimeUnit.Months), "dd/mm/yyyy") & "  (a cada " & Coalesce(gblEquip.PeriodicidadeMeses, 12) & " meses)"'),
              20, 880, "Parent.Width - 40", 34, size=12, bold=True, color=NAVY),
        button(f"btn{p}Salvar", F('If(locSalvando, "Salvando...", "Salvar calibração")'), 20, 924, "Parent.Width - 40", 70, salvar_cal,
               DisplayMode=F('If(locSalvando, DisplayMode.Disabled, DisplayMode.Edit)')),
    ]
    return {"scrCalibrar": {"Properties": {"Fill": f"={BG}", "OnVisible": onvisible}, "Children": filhos}}


# ------------------------------------------------------------------ escrita do YAML
def precisa_bloco(v: str) -> bool:
    return "\n" in v or ": " in v or " #" in v or v.endswith(":")


def emitir(no, ind=0):
    pad = " " * ind
    out = []
    for k, v in no.items():
        if isinstance(v, dict):
            out.append(f"{pad}{k}:")
            out += emitir(v, ind + 2)
        elif isinstance(v, list):
            out.append(f"{pad}{k}:")
            for item in v:
                (nome, corpo), = item.items()
                out.append(f"{pad}  - {nome}:")
                out += emitir(corpo, ind + 6)
        else:
            if precisa_bloco(v):
                out.append(f"{pad}{k}: |-")
                out += [f"{pad}  {linha}" if linha else "" for linha in v.split("\n")]
            else:
                out.append(f"{pad}{k}: {v}")
    return out


def para_ptbr(formula: str) -> str:
    """Converte uma fórmula do formato invariável (vírgula) para o Studio em português (ponto e vírgula)."""
    out, i, em_texto = [], 0, False
    while i < len(formula):
        c = formula[i]
        if c == '"':
            em_texto = not em_texto
            out.append(c)
        elif em_texto:
            out.append(c)
        elif c == ";":
            out.append(";;")
        elif c == ",":
            out.append(";")
        elif c == "." and i > 0 and formula[i - 1].isdigit() and i + 1 < len(formula) and formula[i + 1].isdigit():
            j = i - 1
            while j >= 0 and formula[j].isdigit():
                j -= 1
            out.append("," if j < 0 or not (formula[j].isalnum() or formula[j] == "_") else c)
        else:
            out.append(c)
        i += 1
    return "".join(out)


def escrever_leitor_qr():
    """Texto do OnScan do leitor de QR (controle inserido à mão na tela Início)."""
    f = F(ONSCAN)[1:]
    props = [
        ("X", "20"), ("Y", 'If(gblPerfil = "Laboratorista", 206, 176)'), ("Width", "Parent.Width - 40"), ("Height", "104"),
        ("Visible", 'gblPerfil <> "Colaborador"'), ("Text", '""'), ("Fill", "RGBA(0, 0, 0, 0)"), ("HoverFill", "RGBA(255, 255, 255, 0.08)"),
        ("PressedFill", "RGBA(255, 255, 255, 0.16)"), ("BorderThickness", "0"),
    ]
    tab = "\n".join(f"  {k:<16} {para_ptbr(v)}" for k, v in props)
    tab_en = "\n".join(f"  {k:<16} {v}" for k, v in props)
    texto = f"""LEITOR DE QR CODE DA TELA INÍCIO (scrInicio) · gerado por ferramentas/gerar_telas.py
Guia do Power Apps, seção 2.5.

1. Selecione a tela scrInicio. Menu Inserir > Mídia > Leitor de código de barras
   (em inglês: Insert > Media > Barcode reader).
2. Renomeie o controle para:  bcrIniScanner
3. Propriedades (o controle fica transparente, por cima do cartão azul "Ler QR Code"):
   Se alguma propriedade de cor não existir na sua versão, deixe como está:
   o botão continua funcionando.

   --- Studio em português ---
{tab}

   --- Studio em inglês ---
{tab_en}

4. Propriedade OnScan: cole a fórmula abaixo (escolha o idioma do seu Studio).

=== OnScan · Studio em PORTUGUÊS (o mais comum) ===
{para_ptbr(f)}

=== OnScan · Studio em INGLÊS ===
{f}
"""
    (DESTINO.parent / "leitor-qr-OnScan.txt").write_text(texto, encoding="utf-8")


def main():
    DESTINO.mkdir(parents=True, exist_ok=True)
    DESTINO_CTRL.mkdir(parents=True, exist_ok=True)
    telas = [("tela-1-scrInicio", tela_inicio()), ("tela-2-scrEquipamento", tela_equipamento()),
             ("tela-3-scrEnsaio", tela_ensaio()), ("tela-4-scrVerificacao", tela_verificacao()),
             ("tela-5-scrCalibrar", tela_calibrar())]
    nomes = set()
    for arq, tela in telas:
        texto = "\n".join(emitir({"Screens": tela})) + "\n"
        doc = yaml.safe_load(texto)  # garante YAML válido
        (nome_tela, corpo), = doc["Screens"].items()

        def checar(filhos):
            for item in filhos:
                (n, c), = item.items()
                assert n not in nomes, f"nome de controle repetido: {n}"
                nomes.add(n)
                for pk, pv in c["Properties"].items():
                    assert isinstance(pv, str) and pv.startswith("="), (n, pk, pv)
                checar(c.get("Children", []))
        checar(corpo["Children"])
        (DESTINO / f"{arq}.yaml").write_text(texto, encoding="utf-8")
        ctrl = "\n".join(emitir({"_": corpo["Children"]})[1:])
        ctrl = "\n".join(l[2:] if l.startswith("  ") else l for l in ctrl.split("\n")) + "\n"
        yaml.safe_load(ctrl)
        (DESTINO_CTRL / f"{arq}-controles.yaml").write_text(ctrl, encoding="utf-8")
        onv = corpo["Properties"]["OnVisible"][1:]
        (DESTINO_CTRL / f"{arq}-OnVisible.txt").write_text(
            "=== Studio em PORTUGUÊS ===\n" + para_ptbr(onv) + "\n\n=== Studio em INGLÊS ===\n" + onv + "\n", encoding="utf-8")
        print(f"{arq}: {len(corpo['Children'])} controles, {len(texto.splitlines())} linhas")
    print(f"total de controles com nome único: {len(nomes)}")
    escrever_leitor_qr()


if __name__ == "__main__":
    main()
