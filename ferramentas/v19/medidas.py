# -*- coding: utf-8 -*-
"""Medidas da v15 (tabela _Painel). Cada item: (nome, expressão DAX, formato, pasta, oculta)."""

def PER(e):
    """Sem filtro de período na página -> mês corrente. Com filtro (fatia de mês, eixo de
    semana, filtro de visual em Mês Offset) -> respeita o filtro."""
    return ("IF (\n    ISCROSSFILTERED ( DimCalendario ),\n    %s,\n    CALCULATE ( %s, KEEPFILTERS ( DimCalendario[Mês Atual] = TRUE () ) )\n)" % (e, e))

AZUL, OCRE, LARANJA, CINZA = '"#1B6FB0"', '"#C99400"', '"#E0620F"', '"#8C98A2"'

def nivel_txt(n):
    return ('SWITCH ( %s, 1, "● Na meta", 2, "▲ Atenção", 3, "◆ Crítico", "○ Sem dados" )' % n)

def nivel_cor(n):
    return 'SWITCH ( %s, 1, %s, 2, %s, 3, %s, %s )' % (n, AZUL, OCRE, LARANJA, CINZA)

def nivel_razao(v, m, lim=0.9):
    return ('VAR _v = %s\nVAR _m = %s\nVAR _r = DIVIDE ( _v, _m )\nRETURN\n    IF ( ISBLANK ( _v ) || ISBLANK ( _m ), BLANK (), IF ( _r >= 1, 1, IF ( _r >= %s, 2, 3 ) ) )' % (v, m, lim))

def cfg(nome, padrao, pct=True):
    v = 'VALUE ( _v )'
    return ('VAR _v = LOOKUPVALUE ( tbl_Config[Valor], tbl_Config[Parâmetro], "%s" )\nRETURN\n    IF ( ISBLANK ( _v ) || _v = "", %s, %s )'
            % (nome, padrao, "DIVIDE ( %s, 100 )" % v if pct else v))

PERIODO = """VAR _sel = ISFILTERED ( DimCalendario[Ano Mês Nome] )
VAR _ini = IF ( _sel, CALCULATE ( MIN ( DimCalendario[Data] ), ALLSELECTED ( DimCalendario ) ), DATE ( YEAR ( TODAY () ), MONTH ( TODAY () ), 1 ) )
VAR _fimMes = IF ( _sel, CALCULATE ( MAX ( DimCalendario[Data] ), ALLSELECTED ( DimCalendario ) ), EOMONTH ( TODAY (), 0 ) )
VAR _fim = MIN ( _fimMes, TODAY () - 1 )"""

PROG_LINHA = """IF (
                _fim < _ini, 0,
                IF (
                    _sem,
                    CALCULATE ( DISTINCTCOUNT ( DimCalendario[Semana Início] ), ALL ( DimCalendario ), DimCalendario[Data] >= _ini, DimCalendario[Data] <= _fim ),
                    _por
                        * CALCULATE (
                            COUNTROWS ( DimCalendario ),
                            ALL ( DimCalendario ),
                            DimCalendario[Data] >= _ini,
                            DimCalendario[Data] <= _fim,
                            CONTAINSSTRING ( _dias, FORMAT ( DimCalendario[Dia Semana Número], "0" ) )
                        )
                )
            )"""

# equipamentos com NC no período cuja calibração externa está vencida
CRUZ = """CONCATENATEX (
    FILTER (
        ADDCOLUMNS ( ALL ( tbl_Afericoes[TagKey] ), "@nc", [NC Aferição] ),
        [@nc] > 0
            && CALCULATE (
                COUNTROWS ( tbl_Calibracao ),
                FILTER ( ALL ( tbl_Calibracao[TagKey] ), SUBSTITUTE ( tbl_Calibracao[TagKey], "-", "" ) = tbl_Afericoes[TagKey] ),
                KEEPFILTERS ( tbl_Calibracao[StatusAtual] = "Vencido" )
            ) > 0
    ),
    tbl_Afericoes[TagKey],
    ", ",
    tbl_Afericoes[TagKey], ASC
)"""

M = []
def add(nome, expr, fmt=None, pasta="", oculta=False):
    M.append((nome, expr, fmt, pasta, oculta))

# ── 10 período e metas ─────────────────────────────────────────────────────
add("Mês de Referência", """VAR _d = IF ( ISFILTERED ( DimCalendario[Ano Mês Nome] ), MAX ( DimCalendario[Data] ), TODAY () )
VAR _t = FORMAT ( _d, "MMMM yyyy", "pt-BR" )
RETURN
    UPPER ( LEFT ( _t, 1 ) ) & MID ( _t, 2, 100 )""", None, "10 Período e metas")
add("Meta Calibração em Dia", cfg("Meta Calibração em Dia (%)", "0.9"), "0%", "10 Período e metas")
add("Meta Aferições Conformes", cfg("Meta Aferições Conformes (%)", "0.95"), "0%", "10 Período e metas")
add("Meta Aderência à Rotina", cfg("Meta Aderência Rotina (%)", "0.9"), "0%", "10 Período e metas")
add("Meta Inspeções no Mês", cfg("Meta Inspeções/Mês", "100", pct=False), "#,0", "10 Período e metas")

# ── 11 disponibilidade e paradas ───────────────────────────────────────────
P = "11 Disponibilidade e paradas"
add("Disponibilidade no Mês", """VAR _v = CALCULATE ( [% Disponibilidade], KEEPFILTERS ( DimCalendario[Mês Atual] = TRUE () ) )
RETURN
    IF ( ISBLANK ( _v ), BLANK (), MAX ( _v, 0 ) )""", "0.0%;-0.0%;0.0%", P)
add("Horas Paradas no Mês", "CALCULATE ( [Horas Paradas], KEEPFILTERS ( DimCalendario[Mês Atual] = TRUE () ) )", "#,0.0", P)
add("Horas Paradas (mês)", "COALESCE ( [Horas Paradas no Mês], 0 )", "#,0.0", P)
add("Paradas no Mês", "COALESCE ( CALCULATE ( COUNTROWS ( Tbl_Paradas ), KEEPFILTERS ( DimCalendario[Mês Atual] = TRUE () ) ), 0 )", "#,0", P)
add("Paradas em Aberto", """COALESCE (
    CALCULATE (
        COUNTROWS ( Tbl_Paradas ),
        KEEPFILTERS ( NOT ( Tbl_Paradas[Situação] IN { "Resolvido", "Resolvida", "Concluído", "Concluída" } ) ),
        KEEPFILTERS ( NOT ISBLANK ( Tbl_Paradas[Situação] ) )
    ),
    0
)""", "#,0", P)
add("Disponibilidade Semana", """VAR _dias = CALCULATE ( COUNTROWS ( DimCalendario ), KEEPFILTERS ( DimCalendario[Data Passada] = TRUE () ) )
VAR _cap = [Total Equipamentos] * _dias * 24
RETURN
    IF ( _cap > 0, MAX ( DIVIDE ( _cap - [Horas Paradas], _cap ), 0 ) )""", "0.0%;-0.0%;0.0%", P)
add("Nível Disponibilidade", nivel_razao("[Disponibilidade no Mês]", "[Param Meta Disponibilidade]"), "0", P, True)
add("Situação Disponibilidade", nivel_txt("[Nível Disponibilidade]"), None, P)
add("Cor Disponibilidade", nivel_cor("[Nível Disponibilidade]"), None, P, True)
add("Disponibilidade · Meta", """VAR _n = [Nível Disponibilidade]
RETURN
    "meta " & FORMAT ( [Param Meta Disponibilidade], "0%" ) & " · "
        & SWITCH ( _n, 1, "na meta", 2, "um pouco abaixo da meta", 3, "abaixo da meta", "sem dados no mês" )""", None, P)
add("Disponibilidade · Fatos", """VAR _a = [Paradas em Aberto]
RETURN
    FORMAT ( [Horas Paradas (mês)], "#,0.0" ) & " h paradas no mês  ·  " & _a
        & IF ( _a = 1, " parada em aberto", " paradas em aberto" )""", None, P)
add("Barra Disponibilidade", "MIN ( COALESCE ( [Disponibilidade no Mês], 0 ), 1 )", "0%", P, True)
add("Barra Disponibilidade Resto", "1 - [Barra Disponibilidade]", "0%", P, True)
add("Cor Barra Disponibilidade Grupo", """VAR _m = COALESCE ( MAX ( tbl_Grupos[Meta Disponibilidade] ), [Param Meta Disponibilidade] )
VAR _v = [Disponibilidade no Mês]
RETURN
    IF ( ISBLANK ( _v ), %s, IF ( _v >= _m, %s, IF ( _v >= _m * 0.9, %s, %s ) ) )""" % (CINZA, AZUL, OCRE, LARANJA), None, P, True)
add("Cor Horas Paradas", """IF ( [Paradas em Aberto] > 0, %s, %s )""" % (LARANJA, AZUL), None, P, True)
add("Última Parada", """VAR _d = CALCULATE ( MAX ( Tbl_Paradas[Data] ) )
RETURN
    IF (
        NOT ISBLANK ( _d ),
        FORMAT ( _d, "dd/MM" ) & " · " & CALCULATE ( MAX ( Tbl_Paradas[Ocorrência] ), Tbl_Paradas[Data] = _d )
    )""", None, P)
add("Situação da Última Parada", """VAR _d = CALCULATE ( MAX ( Tbl_Paradas[Data] ) )
RETURN
    IF ( NOT ISBLANK ( _d ), CALCULATE ( MAX ( Tbl_Paradas[Situação] ), Tbl_Paradas[Data] = _d ) )""", None, P)
add("Cor Situação da Parada", """VAR _s = [Situação da Última Parada]
RETURN
    IF ( ISBLANK ( _s ) || _s IN { "Resolvido", "Resolvida", "Concluído", "Concluída" }, "#1D2B36", %s )""" % LARANJA, None, P, True)

# ── 12 calibração ──────────────────────────────────────────────────────────
C = "12 Calibração (gerencial)"
add("% Calibrações em Dia", "DIVIDE ( [Calibracoes Em Dia] + 0, [Total Instrumentos] )", "0%", C)
add("Nível Calibração", nivel_razao("[% Calibrações em Dia]", "[Meta Calibração em Dia]"), "0", C, True)
add("Situação Calibração", nivel_txt("[Nível Calibração]"), None, C)
add("Cor Calibração", nivel_cor("[Nível Calibração]"), None, C, True)
add("Calibração · Meta", """"meta " & FORMAT ( [Meta Calibração em Dia], "0%" ) & " · "
    & FORMAT ( [Calibracoes Em Dia] + 0, "#,0" ) & " de " & FORMAT ( [Total Instrumentos] + 0, "#,0" ) & " instrumentos em dia\"""", None, C)
add("Calibração · Fatos", """VAR _v = [Calibracoes Vencidas] + 0
RETURN
    FORMAT ( _v, "#,0" ) & " vencidas (" & FORMAT ( DIVIDE ( _v, [Total Instrumentos] ), "0%" ) & " do cadastro)  ·  "
        & FORMAT ( [Calibracoes Sem Data] + 0, "#,0" ) & " sem data de vencimento\"""", None, C)
add("Vencem em até 30 dias", "CALCULATE ( COUNTROWS ( tbl_Calibracao ), KEEPFILTERS ( tbl_Calibracao[DiasParaVencer] >= 0 && tbl_Calibracao[DiasParaVencer] <= 30 ) ) + 0", "#,0", C)
add("Vencem em 31 a 60 dias", "CALCULATE ( COUNTROWS ( tbl_Calibracao ), KEEPFILTERS ( tbl_Calibracao[DiasParaVencer] >= 31 && tbl_Calibracao[DiasParaVencer] <= 60 ) ) + 0", "#,0", C)
add("Vencem em 61 a 90 dias", "CALCULATE ( COUNTROWS ( tbl_Calibracao ), KEEPFILTERS ( tbl_Calibracao[DiasParaVencer] >= 61 && tbl_Calibracao[DiasParaVencer] <= 90 ) ) + 0", "#,0", C)
add("Barra Calibração", "MIN ( COALESCE ( [% Calibrações em Dia], 0 ), 1 )", "0%", C, True)
add("Barra Calibração Resto", "1 - [Barra Calibração]", "0%", C, True)

# ── 13 aferições ───────────────────────────────────────────────────────────
A = "13 Aferições"
add("_Afer Avaliadas", 'CALCULATE ( COUNTROWS ( tbl_Afericoes ), KEEPFILTERS ( tbl_Afericoes[Resultado] IN { "Conforme", "Não conforme" } ) )', "#,0", A, True)
add("_Afer Conformes", 'CALCULATE ( COUNTROWS ( tbl_Afericoes ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Conforme" ) )', "#,0", A, True)
add("_Afer NC", 'CALCULATE ( COUNTROWS ( tbl_Afericoes ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Não conforme" ) )', "#,0", A, True)
add("Aferições Avaliadas", PER("[_Afer Avaliadas]"), "#,0", A)
add("Aferições Conformes", PER("[_Afer Conformes]"), "#,0", A)
add("NC Aferição", PER("[_Afer NC]"), "#,0", A)
add("NC Aferição (cartão)", "COALESCE ( [NC Aferição], 0 )", "#,0", A)
add("% Aferições Conformes", "DIVIDE ( [Aferições Conformes] + 0, [Aferições Avaliadas] )", "0.0%;-0.0%;0.0%", A)
add("Linha no Período", "IF ( NOT ISBLANK ( " + PER("COUNTROWS ( tbl_Afericoes )") + " ), 1 )", "0", A, True)
add("Equipamentos com NC", PER('CALCULATE ( DISTINCTCOUNT ( tbl_Afericoes[Equipamento] ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Não conforme" ) )') + " + 0", "#,0", A)
add("Reincidentes", 'COUNTROWS ( FILTER ( VALUES ( tbl_Afericoes[Equipamento] ), [NC Aferição] >= 2 ) ) + 0', "#,0", A)
add("Reincidentes · Lista", """VAR _t = FILTER ( ADDCOLUMNS ( VALUES ( tbl_Afericoes[Equipamento] ), "@nc", [NC Aferição] ), [@nc] >= 2 )
RETURN
    IF (
        COUNTROWS ( _t ) > 0,
        CONCATENATEX ( TOPN ( 4, _t, [@nc], DESC, tbl_Afericoes[Equipamento], ASC ), tbl_Afericoes[Equipamento] & " (" & [@nc] & "×)", " · ", [@nc], DESC ),
        "nenhum equipamento reincidente"
    )""", None, A)
# v19 · regra do DIA CERTO. Nível de um ensaio num dia (contexto: 1 data e 1 ensaio):
#   1 feito no dia certo · 2 feito com atraso (depois do dia certo e antes da próxima data da rotina)
#   3 não feito · 4 incompleto (Blaine: 1 dos 2 turnos) · 5 registro fora do dia programado
#   6 programado (hoje ou futuro)
add("_Mapa Nível Data", """VAR _d = SELECTEDVALUE ( DimCalendario[Data] )
VAR _hoje = TODAY ()
VAR _dias = SELECTEDVALUE ( DimEnsaio[Dias Semana] )
VAR _por = COALESCE ( SELECTEDVALUE ( DimEnsaio[Por Dia] ), 1 )
VAR _rot = SELECTEDVALUE ( DimEnsaio[Tem Rotina] ) = TRUE ()
VAR _primeiro = CALCULATE ( MIN ( tbl_Afericoes[Data] ), REMOVEFILTERS () )
VAR _inicio = IF ( NOT ISBLANK ( _primeiro ), DATE ( YEAR ( _primeiro ), MONTH ( _primeiro ), 1 ) )
VAR _sess =
    IF (
        _por >= 2,
        CALCULATE ( COUNTROWS ( tbl_Afericoes ), ALL ( DimCalendario ), tbl_Afericoes[Data] = _d, tbl_Afericoes[Parâmetro] = "Blaine Manual" ),
        IF ( CALCULATE ( COUNTROWS ( tbl_Afericoes ), ALL ( DimCalendario ), tbl_Afericoes[Data] = _d ) > 0, 1, 0 )
    )
VAR _s = COALESCE ( _sess, 0 )
VAR _prog = _rot && NOT ISBLANK ( _d ) && CONTAINSSTRING ( _dias, FORMAT ( WEEKDAY ( _d, 2 ), "0" ) )
-- atraso vale até a véspera da PRÓXIMA data da rotina (e nunca depois de hoje)
VAR _prox =
    CALCULATE (
        MIN ( DimCalendario[Data] ),
        ALL ( DimCalendario ),
        DimCalendario[Data] > _d,
        CONTAINSSTRING ( _dias, FORMAT ( DimCalendario[Dia Semana Número], "0" ) )
    )
VAR _limite = MIN ( COALESCE ( _prox, DATE ( 2100, 1, 1 ) ) - 1, _hoje )
VAR _depois = CALCULATE ( COUNTROWS ( tbl_Afericoes ), ALL ( DimCalendario ), tbl_Afericoes[Data] > _d, tbl_Afericoes[Data] <= _limite )
RETURN
    IF (
        ISBLANK ( _d ) || ISBLANK ( _inicio ) || _d < _inicio, BLANK (),
        IF (
            NOT _prog,
            IF ( _s > 0 && _d <= _hoje, 5 ),
            IF (
                _s >= _por, 1,
                IF ( _d >= _hoje, 6, IF ( _s > 0, 4, IF ( COALESCE ( _depois, 0 ) > 0, 2, 3 ) ) )
            )
        )
    )""", "0", A, True)

def conta(niveis):
    return PERIODO + """
VAR _datas = CALCULATETABLE ( VALUES ( DimCalendario[Data] ), ALL ( DimCalendario ), DimCalendario[Data] >= _ini, DimCalendario[Data] <= MIN ( _fimMes, TODAY () ) )
RETURN
    IF (
        ISEMPTY ( ALL ( tbl_Afericoes ) ), BLANK (),
        SUMX (
            FILTER ( VALUES ( DimEnsaio[Ensaio] ), CALCULATE ( SELECTEDVALUE ( DimEnsaio[Tem Rotina] ) ) = TRUE () ),
            SUMX (
                _datas,
                VAR _x = DimCalendario[Data]
                RETURN
                    IF ( CALCULATE ( [_Mapa Nível Data], ALL ( DimCalendario ), DimCalendario[Data] = _x ) IN { %s }, 1 )
            )
        ) + 0
    )""" % niveis

add("Rotina · No Dia Certo", conta("1"), "#,0", A)
add("Rotina · Com Atraso", conta("2, 4"), "#,0", A)
add("Rotina · Não Feitas", conta("3"), "#,0", A)
add("Rotina · Programadas", conta("1, 2, 3, 4"), "#,0", A)
# nomes antigos mantidos (usados em textos e na Visão Geral): agora seguem a regra do dia certo
add("Aferições Programadas", "[Rotina · Programadas]", "#,0", A)
add("Aferições Realizadas", "[Rotina · No Dia Certo]", "#,0", A)
add("Aderência à Rotina", "IF ( ISEMPTY ( ALL ( tbl_Afericoes ) ), BLANK (), DIVIDE ( [Aferições Realizadas], [Aferições Programadas] ) )", "0%", A)
add("Ensaio Aparece", "IF ( NOT ISEMPTY ( ALL ( tbl_Afericoes ) ) && ( NOT ISBLANK ( [Aferições Avaliadas] ) || [Aferições Programadas] > 0 ), 1, 0 )", "0", A, True)
add("Nível Aferições", """VAR _v = [% Aferições Conformes]
VAR _m = [Meta Aferições Conformes]
RETURN
    IF ( ISBLANK ( _v ), BLANK (), IF ( _v >= _m, 1, IF ( _v >= _m - 0.05, 2, 3 ) ) )""", "0", A, True)
add("Situação Aferições", nivel_txt("[Nível Aferições]"), None, A)
add("Cor Aferições", nivel_cor("[Nível Aferições]"), None, A, True)
add("Aferições · Meta", """"meta " & FORMAT ( [Meta Aferições Conformes], "0%" ) & " · "
    & FORMAT ( [Aferições Conformes] + 0, "#,0" ) & " de " & FORMAT ( [Aferições Avaliadas] + 0, "#,0" ) & " registros\"""", None, A)
add("Aferições · Fatos", """VAR _nc = [NC Aferição (cartão)]
RETURN
    _nc & IF ( _nc = 1, " não conformidade em ", " não conformidades em " ) & [Equipamentos com NC] & " equip.  ·  rotina "
        & IF ( ISBLANK ( [Aderência à Rotina] ), "—", FORMAT ( [Aderência à Rotina], "0%" ) )""", None, A)
add("Aderência · Meta", """"meta " & FORMAT ( [Meta Aderência à Rotina], "0%" ) & " · "
    & FORMAT ( [Aferições Realizadas] + 0, "#,0" ) & " de " & FORMAT ( [Aferições Programadas] + 0, "#,0" ) & " no dia certo\"""", None, A)
add("Barra Aferições", "MIN ( COALESCE ( [% Aferições Conformes], 0 ), 1 )", "0%", A, True)
add("Barra Aferições Resto", "1 - [Barra Aferições]", "0%", A, True)
add("Barra Aderência", "MIN ( COALESCE ( [Aderência à Rotina], 0 ), 1 )", "0%", A, True)
add("Barra Aderência Resto", "1 - [Barra Aderência]", "0%", A, True)
add("Nível Ensaio", """VAR _pc = [% Aferições Conformes]
VAR _ad = [Aderência à Rotina]
VAR _m = [Meta Aferições Conformes]
VAR _ma = [Meta Aderência à Rotina]
RETURN
    IF (
        ISBLANK ( _pc ) && ISBLANK ( _ad ), BLANK (),
        IF (
            NOT ISBLANK ( _pc ) && _pc < _m - 0.05, 3,
            IF ( ( NOT ISBLANK ( _pc ) && _pc < _m ) || ( NOT ISBLANK ( _ad ) && _ad < _ma ), 2, 1 )
        )
    )""", "0", A, True)
add("Situação Ensaio", nivel_txt("[Nível Ensaio]"), None, A)
add("Cor Situação Ensaio", nivel_cor("[Nível Ensaio]"), None, A, True)
add("Último Registro", "VAR _d = " + PER("MAX ( tbl_Afericoes[Data] )") + '\nRETURN\n    IF ( NOT ISBLANK ( _d ), FORMAT ( _d, "dd/MM" ) )', None, A)
add("NC com Calibração Vencida", CRUZ, None, A)
add("Ensaio da NC", PER('CALCULATE ( MAX ( tbl_Afericoes[Ensaio] ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Não conforme" ) )'), None, A)
add("O que Falhou", PER('CALCULATE ( CONCATENATEX ( VALUES ( tbl_Afericoes[Parâmetro] ), tbl_Afericoes[Parâmetro], "; ", tbl_Afericoes[Parâmetro], ASC ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Não conforme" ) )'), None, A)
add("Datas das NC", PER('CALCULATE ( CONCATENATEX ( VALUES ( tbl_Afericoes[Data] ), FORMAT ( tbl_Afericoes[Data], "dd/MM" ), ", ", tbl_Afericoes[Data], ASC ), KEEPFILTERS ( tbl_Afericoes[Resultado] = "Não conforme" ) )'), None, A)
add("Calibração Externa", """VAR _k = SELECTEDVALUE ( tbl_Afericoes[TagKey] )
VAR _st =
    CALCULATE (
        MIN ( tbl_Calibracao[StatusAtual] ),
        FILTER ( ALL ( tbl_Calibracao[TagKey] ), SUBSTITUTE ( tbl_Calibracao[TagKey], "-", "" ) = _k )
    )
RETURN
    IF ( ISBLANK ( _k ), BLANK (), IF ( ISBLANK ( _st ), "sem cadastro", _st ) )""", None, A)
add("Gravidade da NC", """VAR _nc = [NC Aferição]
RETURN
    IF ( ISBLANK ( _nc ), BLANK (), IF ( _nc >= 3 || [Calibração Externa] = "Vencido", "◆ Crítico", "▲ Atenção" ) )""", None, A)
add("Cor Gravidade da NC", """IF ( LEFT ( [Gravidade da NC], 1 ) = "◆", %s, %s )""" % (LARANJA, OCRE), None, A, True)
add("Cor Resultado", """SWITCH ( SELECTEDVALUE ( tbl_Afericoes[Resultado] ), "Não conforme", %s, "Conforme", %s, "#5B6B78" )""" % (LARANJA, AZUL), None, A, True)
add("% da Tolerância (último)", PER("""CALCULATE (
        VAR _d = MAX ( tbl_Afericoes[Data] )
        RETURN CALCULATE ( MAX ( tbl_Afericoes[% da Tolerância] ), tbl_Afericoes[Data] = _d )
    )"""), "0%", A)
add("Cor % da Tolerância", """VAR _p = [%% da Tolerância (último)]
RETURN
    IF ( ISBLANK ( _p ), %s, IF ( _p < 0.8, %s, IF ( _p <= 1, %s, %s ) ) )""" % (CINZA, AZUL, OCRE, LARANJA), None, A, True)
# carta de controle: UM equipamento e UM parâmetro escolhidos; sem nada escolhido abre no Blaine Manual
CARTA_VARS = """VAR _livre =
    NOT ISFILTERED ( tbl_Afericoes[Equipamento] ) && NOT ISFILTERED ( tbl_Afericoes[Parâmetro] )
        && NOT ISFILTERED ( DimEnsaio[Ensaio Curto] ) && NOT ISFILTERED ( DimEnsaio[Ensaio] )
VAR _ok =
    CALCULATE ( DISTINCTCOUNT ( tbl_Afericoes[Equipamento] ), ALLSELECTED ( tbl_Afericoes[Data] ) ) = 1
        && CALCULATE ( DISTINCTCOUNT ( tbl_Afericoes[Parâmetro] ), ALLSELECTED ( tbl_Afericoes[Data] ) ) = 1
VAR _tl = NOT ISBLANK ( CALCULATE ( MAX ( tbl_Afericoes[Lim. Inferior] ), ALLSELECTED ( tbl_Afericoes[Data] ) ) )"""
def carta(nome, esc_lim, esc_tol, padrao, cond_valor=True):
    corpo = CARTA_VARS + """
VAR _v =
    IF (
        _ok, """ + PER("IF ( _tl, %s, %s )" % (esc_lim, esc_tol)) + """,
        IF ( _livre, CALCULATE ( """ + PER(padrao) + """, tbl_Afericoes[Parâmetro] = "Blaine Manual" ) )
    )
RETURN
    """ + ("_v" if not cond_valor else "IF ( NOT ISBLANK ( [Carta · Valor] ), _v )")
    add(nome, corpo, "#,0.00", A)
carta("Carta · Valor", "AVERAGE ( tbl_Afericoes[Valor] )", "AVERAGE ( tbl_Afericoes[Diferença] )", "AVERAGE ( tbl_Afericoes[Valor] )", cond_valor=False)
carta("Carta · Lim. Inferior", "MIN ( tbl_Afericoes[Lim. Inferior] )", "- MIN ( tbl_Afericoes[Tolerância] )", "MIN ( tbl_Afericoes[Lim. Inferior] )")
carta("Carta · Lim. Superior", "MAX ( tbl_Afericoes[Lim. Superior] )", "MAX ( tbl_Afericoes[Tolerância] )", "MAX ( tbl_Afericoes[Lim. Superior] )")
carta("Carta · Nominal", "AVERAGE ( tbl_Afericoes[Referência] )", "0", "AVERAGE ( tbl_Afericoes[Referência] )")
add("Carta · Título", CARTA_VARS + """
VAR _eqB = CALCULATE ( MAX ( tbl_Afericoes[Equipamento] ), tbl_Afericoes[Parâmetro] = "Blaine Manual" )
RETURN
    SWITCH (
        TRUE (),
        _ok,
            "Carta de controle · " & SELECTEDVALUE ( tbl_Afericoes[Equipamento] ) & " · " & SELECTEDVALUE ( tbl_Afericoes[Parâmetro] )
                & IF ( _tl, "   (valor × limites)", "   (diferença × ± tolerância)" ),
        _livre,
            "Carta de controle · Blaine Manual · " & _eqB & "   (padrão; escolha outro ensaio e equipamento na lateral)",
        "Carta de controle · escolha UM equipamento e UM parâmetro na lateral"
    )""", None, A)

# mapa (matriz tbl_MapaLinhas × DimCalendario[Dia]) · v19: regra do dia certo
# linha 0 = dia da semana (10) ou hoje (11) · demais linhas = _Mapa Nível Data do ensaio (0 = vazio)
add("_Mapa Data", PERIODO.replace("VAR _fim = MIN ( _fimMes, TODAY () - 1 )", "") + """
VAR _dia = SELECTEDVALUE ( DimCalendario[Dia] )
RETURN
    IF (
        NOT ISBLANK ( _dia ),
        CALCULATE ( MAX ( DimCalendario[Data] ), ALL ( DimCalendario ), DimCalendario[Data] >= _ini, DimCalendario[Data] <= _fimMes, DimCalendario[Dia] = _dia )
    )""", "dd/mm/yyyy", A, True)
add("Mapa · Nível", """VAR _ordem = SELECTEDVALUE ( tbl_MapaLinhas[Ordem] )
VAR _ens = SELECTEDVALUE ( tbl_MapaLinhas[Ensaio] )
VAR _d = [_Mapa Data]
VAR _visivel = NOT ISBLANK ( _ens ) && CALCULATE ( COUNTROWS ( DimEnsaio ), KEEPFILTERS ( DimEnsaio[Ensaio] = _ens ) ) > 0
RETURN
    IF (
        ISBLANK ( _d ) || ISBLANK ( _ordem ) || ISEMPTY ( ALL ( tbl_Afericoes ) ), BLANK (),
        IF (
            _ordem = 0, IF ( _d = TODAY (), 11, 10 ),
            IF (
                _visivel,
                COALESCE ( CALCULATE ( [_Mapa Nível Data], ALL ( DimCalendario ), DimCalendario[Data] = _d, DimEnsaio[Ensaio] = _ens ), 0 )
            )
        )
    )""", "0", A, True)
add("Mapa", """VAR _n = [Mapa · Nível]
VAR _d = [_Mapa Data]
RETURN
    SWITCH (
        _n,
        1, "✔",
        2, " ",
        3, "✖",
        4, "½",
        5, "✔",
        6, "○",
        0, " ",
        10, SWITCH ( WEEKDAY ( _d, 2 ), 1, "seg", 2, "ter", 3, "qua", 4, "qui", 5, "sex", 6, "sáb", 7, "dom" ),
        11, "hoje"
    )""", None, A)
add("Mapa · Fundo", """SWITCH ( [Mapa · Nível],
    1, "#1B6FB0", 2, "#F3CF5B", 3, "#E0620F", 4, "#F8D9C4", 5, "#DCEAF6", 6, "#FFFFFF", 0, "#F1F4F7",
    11, "#FFC000", "#FFFFFF" )""", None, A, True)
add("Mapa · Fonte", """SWITCH ( [Mapa · Nível],
    1, "#FFFFFF", 2, "#F3CF5B", 3, "#FFFFFF", 4, "#9A3F05", 5, "#1B6FB0", 6, "#8C98A2", 0, "#F1F4F7",
    10, IF ( WEEKDAY ( [_Mapa Data], 2 ) >= 6, "#9AA8B4", "#5B6B78" ), 11, "#1D2B36", "#5B6B78" )""", None, A, True)

# cola de regras (tabela tbl_MapaLinhas embaixo do mapa)
add("Regra · Dias Certos", PERIODO.replace("VAR _fim = MIN ( _fimMes, TODAY () - 1 )", "") + """
VAR _ens = SELECTEDVALUE ( tbl_MapaLinhas[Ensaio] )
VAR _dias = LOOKUPVALUE ( DimEnsaio[Dias Semana], DimEnsaio[Ensaio], _ens )
VAR _rot = LOOKUPVALUE ( DimEnsaio[Tem Rotina], DimEnsaio[Ensaio], _ens )
VAR _t =
    CALCULATETABLE (
        VALUES ( DimCalendario[Data] ),
        ALL ( DimCalendario ),
        DimCalendario[Data] >= _ini,
        DimCalendario[Data] <= _fimMes,
        CONTAINSSTRING ( _dias, FORMAT ( DimCalendario[Dia Semana Número], "0" ) )
    )
VAR _n = COUNTROWS ( _t )
RETURN
    IF (
        ISBLANK ( _ens ) || _ens = "", BLANK (),
        IF (
            _rot <> TRUE () || COALESCE ( _n, 0 ) = 0, "sem dia programado",
            IF ( _n >= 28, "todos os dias do mês", CONCATENATEX ( _t, FORMAT ( DimCalendario[Data], "dd" ), ", ", DimCalendario[Data], ASC ) )
        )
    )""", None, A)
add("Regra · No Dia Certo", """VAR _ens = SELECTEDVALUE ( tbl_MapaLinhas[Ensaio] )
VAR _p = CALCULATE ( [Rotina · Programadas], KEEPFILTERS ( TREATAS ( { _ens }, DimEnsaio[Ensaio] ) ) )
VAR _ok = CALCULATE ( [Rotina · No Dia Certo], KEEPFILTERS ( TREATAS ( { _ens }, DimEnsaio[Ensaio] ) ) )
VAR _at = CALCULATE ( [Rotina · Com Atraso], KEEPFILTERS ( TREATAS ( { _ens }, DimEnsaio[Ensaio] ) ) )
VAR _nf = CALCULATE ( [Rotina · Não Feitas], KEEPFILTERS ( TREATAS ( { _ens }, DimEnsaio[Ensaio] ) ) )
RETURN
    IF (
        ISBLANK ( _ens ) || _ens = "", BLANK (),
        IF (
            COALESCE ( _p, 0 ) = 0, "—",
            _ok & " de " & _p & "  ·  " & _at & " com atraso  ·  " & _nf & " não feitas"
        )
    )""", None, A)
add("Regra Aparece", "IF ( SELECTEDVALUE ( tbl_MapaLinhas[Ordem] ) > 0, 1, 0 )", "0", A, True)
add("Última Aferição", "VAR _d = " + PER("MAX ( tbl_Afericoes[Data] )") + '\nRETURN\n    IF ( NOT ISBLANK ( _d ), FORMAT ( _d, "dd/MM" ) )', None, A)
add("Último Valor", PER("""CALCULATE (
        VAR _d = MAX ( tbl_Afericoes[Data] )
        RETURN CALCULATE ( AVERAGE ( tbl_Afericoes[Valor] ), tbl_Afericoes[Data] = _d )
    )"""), "#,0.00", A)
add("Média do Período", PER("AVERAGE ( tbl_Afericoes[Valor] )"), "#,0.00", A)
add("Referência Média", PER("AVERAGE ( tbl_Afericoes[Referência] )"), "#,0.00", A)
add("Faixa Aceita", """VAR _li = MIN ( tbl_Afericoes[Lim. Inferior] )
VAR _ls = MAX ( tbl_Afericoes[Lim. Superior] )
VAR _t = MAX ( tbl_Afericoes[Tolerância] )
RETURN
    IF (
        NOT ISBLANK ( _li ), FORMAT ( _li, "#,0.0#" ) & " a " & FORMAT ( _ls, "#,0.0#" ),
        IF ( NOT ISBLANK ( _t ), "± " & FORMAT ( _t, "#,0.0#" ) & " da referência" )
    )""", None, A)
add("Último Resultado", PER("""CALCULATE (
        VAR _d = MAX ( tbl_Afericoes[Data] )
        RETURN CALCULATE ( MAX ( tbl_Afericoes[Resultado] ), tbl_Afericoes[Data] = _d )
    )"""), None, A)
add("Cor Último Resultado", """SWITCH ( [Último Resultado], "Não conforme", "#E0620F", "Conforme", "#1B6FB0", "#5B6B78" )""", None, A, True)
add("Registros no Período", PER("COUNTROWS ( tbl_Afericoes )"), "#,0", A)
add("Mapa · Dias Não Feitos", "[Rotina · Não Feitas]", "0", A)
add("Cor Aderência", """VAR _v = [Aderência à Rotina]
VAR _m = [Meta Aderência à Rotina]
RETURN
    IF ( ISBLANK ( _v ), "#8C98A2", IF ( _v >= _m, "#1B6FB0", IF ( _v >= _m * 0.9, "#C99400", "#E0620F" ) ) )""", None, A, True)

# ── 14 inspeções ───────────────────────────────────────────────────────────
I = "14 Inspeções (gerencial)"
add("Inspeções no Mês", "COALESCE ( " + PER("[Inspecoes Realizadas]") + ", 0 )", "#,0", I)
add("Inspeções Esperadas até Hoje", "ROUNDUP ( [Meta Inspeções no Mês] * DIVIDE ( DAY ( TODAY () ), DAY ( EOMONTH ( TODAY (), 0 ) ) ), 0 )", "#,0", I)
add("Nível Inspeções", nivel_razao("[Inspeções no Mês]", "[Inspeções Esperadas até Hoje]"), "0", I, True)
add("Situação Inspeções", nivel_txt("[Nível Inspeções]"), None, I)
add("Cor Inspeções", nivel_cor("[Nível Inspeções]"), None, I, True)
add("Inspeções · Meta", """"de " & FORMAT ( [Meta Inspeções no Mês], "#,0" ) & " no mês · proporcional até hoje: " & FORMAT ( [Inspeções Esperadas até Hoje], "#,0" )""", None, I)
add("Inspeções · Fatos", """VAR _pl = COALESCE ( """ + PER('CALCULATE ( [Inspecoes Realizadas], KEEPFILTERS ( tbl_Inspecoes[Origem] = "Planilha" ) )') + """, 0 )
VAR _fo = COALESCE ( """ + PER('CALCULATE ( [Inspecoes Realizadas], KEEPFILTERS ( tbl_Inspecoes[Origem] = "Forms" ) )') + """, 0 )
VAR _nc = COALESCE ( """ + PER('CALCULATE ( [Inspecoes Realizadas], KEEPFILTERS ( tbl_Inspecoes[Resultado_Inspecao] = "Nao Conforme" ) )') + """, 0 )
RETURN
    _pl & " pela planilha + " & _fo & " pelo Forms  ·  " & _nc & IF ( _nc = 1, " não conforme", " não conformes" )""", None, I)
add("Barra Inspeções", "MIN ( DIVIDE ( [Inspeções no Mês], [Inspeções Esperadas até Hoje] ), 1 ) + 0", "0%", I, True)
add("Barra Inspeções Resto", "1 - [Barra Inspeções]", "0%", I, True)

# ── 15 visão geral: frase, pontos de decisão, saúde por grupo ─────────────
V = "15 Visão Geral"
add("Resumo Geral", """VAR _t =
    {
        ( "Disponibilidade", [Nível Disponibilidade] ),
        ( "Calibração", [Nível Calibração] ),
        ( "Aferições", [Nível Aferições] ),
        ( "Inspeções", [Nível Inspeções] )
    }
VAR _ok = COALESCE ( COUNTROWS ( FILTER ( _t, [Value2] = 1 ) ), 0 )
VAR _com = COALESCE ( COUNTROWS ( FILTER ( _t, NOT ISBLANK ( [Value2] ) ) ), 0 )
VAR _crit = CONCATENATEX ( FILTER ( _t, [Value2] = 3 ), [Value1], ", " )
VAR _aten = CONCATENATEX ( FILTER ( _t, [Value2] = 2 ), [Value1], ", " )
VAR _p = [Pontos que Pedem Decisão]
RETURN
    [Mês de Referência] & ":  " & _ok & " de " & _com & " frentes na meta."
        & IF ( _crit <> "", "  Críticas: " & _crit & "." )
        & IF ( _aten <> "", "  Em atenção: " & _aten & "." )
        & IF ( _p > 0, "  " & _p & IF ( _p = 1, " ponto pede decisão.", " pontos pedem decisão." ) )""", None, V)
add("_Ponto Ensaio", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
VAR _n = IF ( _k = "AF1", 1, IF ( _k = "AF2", 2, 0 ) )
VAR _t =
    FILTER (
        ADDCOLUMNS (
            ALL ( DimEnsaio[Ensaio] ),
            "@nc", [NC Aferição],
            "@re", CALCULATE ( COUNTROWS ( FILTER ( VALUES ( tbl_Afericoes[Equipamento] ), [NC Aferição] >= 2 ) ) )
        ),
        [@re] > 0
    )
VAR _top = TOPN ( _n, _t, [@nc], DESC, DimEnsaio[Ensaio], ASC )
RETURN
    IF ( _n > 0 && COUNTROWS ( _top ) >= _n, CONCATENATEX ( TOPN ( 1, _top, [@nc], ASC, DimEnsaio[Ensaio], DESC ), DimEnsaio[Ensaio] ) )""", None, V, True)
add("Ponto Ativo", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
RETURN
    SWITCH (
        _k,
        "CAL", IF ( [Calibracoes Vencidas] > 0, 1, 0 ),
        "AF1", IF ( NOT ISBLANK ( [_Ponto Ensaio] ), 1, 0 ),
        "AF2", IF ( NOT ISBLANK ( [_Ponto Ensaio] ), 1, 0 ),
        "INS", IF ( [Nível Inspeções] >= 2, 1, 0 ),
        "PAR", IF ( [Paradas em Aberto] > 0, 1, 0 ),
        "DISP", IF ( [Nível Disponibilidade] >= 2, 1, 0 ),
        0
    )""", "0", V, True)
add("Pontos que Pedem Decisão", "COALESCE ( COUNTROWS ( FILTER ( ALL ( tbl_Pontos ), [Ponto Ativo] = 1 ) ), 0 )", "0", V)
add("Ponto Nº", """IF (
    [Ponto Ativo] = 1,
    RANKX ( FILTER ( ALL ( tbl_Pontos ), [Ponto Ativo] = 1 ), CALCULATE ( MIN ( tbl_Pontos[Ordem] ) ), , ASC )
)""", "0", V)
REINC_ENSAIO = """VAR _e = [_Ponto Ensaio]
VAR _re =
    CALCULATETABLE (
        FILTER ( ADDCOLUMNS ( VALUES ( tbl_Afericoes[Equipamento] ), "@nc", [NC Aferição] ), [@nc] >= 2 ),
        DimEnsaio[Ensaio] = _e
    )"""
add("Ponto Título", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
""" + REINC_ENSAIO + """
VAR _curto = LOOKUPVALUE ( DimEnsaio[Ensaio Curto], DimEnsaio[Ensaio], _e )
RETURN
    SWITCH (
        _k,
        "CAL", FORMAT ( [Calibracoes Vencidas], "#,0" ) & " instrumentos com calibração vencida",
        "AF1", _curto & ": " & CONCATENATEX ( _re, tbl_Afericoes[Equipamento], ", ", [@nc], DESC ) & " com falha repetida",
        "AF2", _curto & ": " & CONCATENATEX ( _re, tbl_Afericoes[Equipamento], ", ", [@nc], DESC ) & " com falha repetida",
        "INS", "Inspeções de amostradores abaixo do ritmo",
        "PAR", [Paradas em Aberto] & IF ( [Paradas em Aberto] = 1, " parada em aberto", " paradas em aberto" ),
        "DISP", "Disponibilidade abaixo da meta"
    )""", None, V)
add("Ponto Detalhe", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
""" + REINC_ENSAIO + """
VAR _ncE = CALCULATE ( [NC Aferição], DimEnsaio[Ensaio] = _e )
VAR _labs =
    CONCATENATEX (
        FILTER ( ADDCOLUMNS ( VALUES ( DimLaboratorio[Laboratorio] ), "@v", [Calibracoes Vencidas] ), [@v] > 0 ),
        DimLaboratorio[Laboratorio] & " " & [@v], ", ", [@v], DESC
    )
VAR _uteis =
    COUNTROWS (
        FILTER ( ALL ( DimCalendario ), DimCalendario[Data] >= TODAY () && DimCalendario[Data] <= EOMONTH ( TODAY (), 0 ) && DimCalendario[Tipo de Dia] = "Dia útil" )
    )
VAR _falta = MAX ( [Meta Inspeções no Mês] - [Inspeções no Mês], 0 )
RETURN
    SWITCH (
        _k,
        "CAL", FORMAT ( DIVIDE ( [Calibracoes Vencidas], [Total Instrumentos] ), "0%" ) & " do cadastro · " & _labs & ".",
        "AF1", _ncE & " não conformidades no mês: " & CONCATENATEX ( _re, tbl_Afericoes[Equipamento] & " " & [@nc] & "×", ", ", [@nc], DESC ) & ".",
        "AF2", _ncE & " não conformidades no mês: " & CONCATENATEX ( _re, tbl_Afericoes[Equipamento] & " " & [@nc] & "×", ", ", [@nc], DESC ) & ".",
        "INS", [Inspeções no Mês] & " de " & [Meta Inspeções no Mês] & " no mês; faltam " & _falta & " em " & _uteis & " dias úteis até " & FORMAT ( EOMONTH ( TODAY (), 0 ), "dd/MM" ) & ".",
        "PAR",
            VAR _ab =
                CALCULATETABLE (
                    ADDCOLUMNS (
                        VALUES ( Tbl_Paradas[Equip. (Tag)] ),
                        "@n", CALCULATE ( COUNTROWS ( Tbl_Paradas ) ),
                        "@nome",
                            VAR _t = Tbl_Paradas[Equip. (Tag)]
                            VAR _e = CALCULATE ( MAX ( Tbl_Paradas[Equipamento (auto)] ) )
                            RETURN IF ( ISBLANK ( _e ) || _e = "", _t, _e & " (" & _t & ")" )
                    ),
                    KEEPFILTERS ( NOT ( Tbl_Paradas[Situação] IN { "Resolvido", "Resolvida", "Concluído", "Concluída" } ) ),
                    KEEPFILTERS ( NOT ISBLANK ( Tbl_Paradas[Situação] ) )
                )
            VAR _neq = COUNTROWS ( _ab )
            RETURN
                _neq & IF ( _neq = 1, " equipamento: ", " equipamentos: " )
                    & CONCATENATEX ( TOPN ( 4, _ab, [@n], DESC, Tbl_Paradas[Equip. (Tag)], ASC ),
                                     [@nome] & IF ( [@n] > 1, " · " & [@n] & " paradas", "" ), "; ", [@n], DESC )
                    & IF ( _neq > 4, "; e mais " & ( _neq - 4 ), "" ) & ".",
        "DISP", FORMAT ( [Disponibilidade no Mês], "0.0%" ) & " no mês × meta " & FORMAT ( [Param Meta Disponibilidade], "0%" ) & "."
    )""", None, V)
add("Ponto Ação", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
VAR _e = [_Ponto Ensaio]
VAR _cruz = [NC com Calibração Vencida]
RETURN
    SWITCH (
        _k,
        "CAL", "Plano de recuperação com a calibradora" & IF ( _cruz <> "", ", começando por " & _cruz & " (também reprovaram na aferição).", ", começando pelos instrumentos críticos." ),
        "AF1", SWITCH ( _e,
                   "Tambor de Abrasão", "Ajustar a velocidade do tambor e reavaliar os ensaios feitos nele no período.",
                   "Verificação Peneiradores", "Substituir ou calibrar as telas reprovadas; conferir a calibração externa.",
                   "Umidade", "Repetir o comparativo e verificar o analisador.",
                   "Compressão", "Repetir o comparativo e verificar a prensa.",
                   "Investigar a causa, corrigir e repetir a aferição." ),
        "AF2", SWITCH ( _e,
                   "Tambor de Abrasão", "Ajustar a velocidade do tambor e reavaliar os ensaios feitos nele no período.",
                   "Verificação Peneiradores", "Substituir ou calibrar as telas reprovadas; conferir a calibração externa.",
                   "Umidade", "Repetir o comparativo e verificar o analisador.",
                   "Compressão", "Repetir o comparativo e verificar a prensa.",
                   "Investigar a causa, corrigir e repetir a aferição." ),
        "INS", "Rever a meta ou redistribuir as inspeções por turno.",
        "PAR", "Cobrar a peça ou o serviço pendente e registrar a previsão de retorno.",
        "DISP", "Atacar as maiores causas de parada do mês (página Equipamentos & Paradas)."
    )""", None, V)
add("Ponto Gravidade", """VAR _k = SELECTEDVALUE ( tbl_Pontos[Chave] )
VAR _e = [_Ponto Ensaio]
VAR _max = MAXX ( FILTER ( ADDCOLUMNS ( VALUES ( tbl_Afericoes[Equipamento] ), "@nc", CALCULATE ( [NC Aferição], DimEnsaio[Ensaio] = _e ) ), NOT ISBLANK ( [@nc] ) ), [@nc] )
VAR _n =
    SWITCH (
        _k,
        "CAL", IF ( [Nível Calibração] = 3, 3, 2 ),
        "AF1", IF ( _max >= 3, 3, 2 ),
        "AF2", IF ( _max >= 3, 3, 2 ),
        "INS", [Nível Inspeções],
        "PAR", 2,
        "DISP", [Nível Disponibilidade]
    )
RETURN
    IF ( [Ponto Ativo] = 1, """ + nivel_txt("_n") + ")", None, V)
add("Cor Ponto Gravidade", """SWITCH ( LEFT ( [Ponto Gravidade], 1 ), "◆", %s, "▲", %s, "●", %s, %s )""" % (LARANJA, OCRE, AZUL, CINZA), None, V, True)

G = "16 Saúde por grupo"
add("Disponibilidade do Grupo", """IF ( SELECTEDVALUE ( tbl_Grupos[Domínio] ) = "Equipamentos de amostragem", [Disponibilidade no Mês] )""", "0.0%;-0.0%;0.0%", G)
add("Calibração do Grupo", """IF ( SELECTEDVALUE ( tbl_Grupos[Domínio] ) = "Instrumentos de laboratório", [% Calibrações em Dia] )""", "0%", G)
add("Aferição do Grupo", "[% Aferições Conformes]", "0.0%;-0.0%;0.0%", G)
add("Equip. no Grupo", "MAX ( tbl_Grupos[Nº Equipamentos] )", "#,0", G)
add("Grupo Aparece", """IF (
    NOT ISBLANK ( SELECTEDVALUE ( tbl_Grupos[Grupo Operacional] ) )
        && ( NOT ISBLANK ( [Disponibilidade do Grupo] ) || NOT ISBLANK ( [Calibração do Grupo] ) || NOT ISBLANK ( [Aferição do Grupo] ) ),
    1, 0
)""", "0", G, True)
add("Nível do Grupo", """VAR _md = COALESCE ( MAX ( tbl_Grupos[Meta Disponibilidade] ), [Param Meta Disponibilidade] )
VAR _d = [Disponibilidade do Grupo]
VAR _c = [Calibração do Grupo]
VAR _a = [Aferição do Grupo]
VAR _mc = [Meta Calibração em Dia]
VAR _ma = [Meta Aferições Conformes]
VAR _nd = IF ( ISBLANK ( _d ), BLANK (), IF ( _d >= _md, 1, IF ( _d >= _md * 0.9, 2, 3 ) ) )
VAR _nc = IF ( ISBLANK ( _c ), BLANK (), IF ( _c >= _mc, 1, IF ( _c >= _mc * 0.9, 2, 3 ) ) )
VAR _na = IF ( ISBLANK ( _a ), BLANK (), IF ( _a >= _ma, 1, IF ( _a >= _ma - 0.05, 2, 3 ) ) )
RETURN
    MAX ( MAX ( _nd, _nc ), _na )""", "0", G, True)
add("Situação do Grupo", nivel_txt("[Nível do Grupo]"), None, G)
add("Cor Situação do Grupo", nivel_cor("[Nível do Grupo]"), None, G, True)
add("Observação do Grupo", """VAR _dom = SELECTEDVALUE ( tbl_Grupos[Domínio] )
VAR _ab = [Paradas em Aberto]
VAR _hp = [Horas Paradas (mês)]
VAR _venc = CALCULATETABLE ( VALUES ( tbl_Calibracao[Tag] ), KEEPFILTERS ( tbl_Calibracao[StatusAtual] = "Vencido" ) )
VAR _nv = COUNTROWS ( _venc )
VAR _tv = IF ( _nv > 0, IF ( _nv <= 3, CONCATENATEX ( _venc, tbl_Calibracao[Tag], ", " ), _nv & " instrumentos" ) & IF ( _nv = 1, " com calibração vencida", " com calibração vencida" ) )
VAR _pior = TOPN ( 1, FILTER ( ADDCOLUMNS ( VALUES ( tbl_Afericoes[Equipamento] ), "@nc", [NC Aferição] ), [@nc] > 0 ), [@nc], DESC, tbl_Afericoes[Equipamento], ASC )
VAR _ta = IF ( COUNTROWS ( _pior ) > 0, CONCATENATEX ( _pior, tbl_Afericoes[Equipamento] & ": " & [@nc] & " NC na aferição" ) )
RETURN
    IF (
        _dom = "Equipamentos de amostragem",
        IF ( _ab > 0, _ab & IF ( _ab = 1, " parada em aberto", " paradas em aberto" ), IF ( _hp > 0, FORMAT ( _hp, "#,0.0" ) & " h paradas no mês", "sem pendências" ) ),
        VAR _txt = _tv & IF ( _tv <> "" && _ta <> "", " · " ) & _ta
        RETURN IF ( _txt = "", "sem pendências", _txt )
    )""", None, G)

add("Fonte Aferições", """VAR _st = MAX ( tbl_Afericoes_Status[Status] )
VAR _pl = MAX ( tbl_Afericoes_Status[Planilhas na Pasta] ) + 0
VAR _com = COALESCE ( COUNTROWS ( ALL ( tbl_Afericoes[Arquivo] ) ), 0 )
RETURN
    IF (
        _st <> "ok", _st,
        "Pasta lida: " & _pl & IF ( _pl = 1, " planilha .xlsm, ", " planilhas .xlsm, " ) & _com
            & " com a aba oculta BD_Afericoes."
            & IF ( _com < _pl, " As demais ainda são a versão antiga (sem a aba oculta) e são ignoradas.", "" )
    )""", None, A)

# resumo automático (Aferições)
add("Resumo Aferições", """VAR _l = SELECTEDVALUE ( tbl_Linhas[Linha] )
VAR _av = [Aferições Avaliadas]
VAR _nc = [NC Aferição (cartão)]
VAR _pior =
    TOPN (
        1,
        FILTER ( ADDCOLUMNS ( FILTER ( ALL ( DimEnsaio ), DimEnsaio[Tem Rotina] ), "@a", [Aderência à Rotina] ), NOT ISBLANK ( [@a] ) && [@a] < 1 ),
        [@a], ASC, DimEnsaio[Ordem], ASC
    )
VAR _tp = IF ( COUNTROWS ( _pior ) > 0, CONCATENATEX ( _pior, DimEnsaio[Ensaio Curto] & " (" & FORMAT ( [@a], "0%" ) & ")" ) )
VAR _cruz = [NC com Calibração Vencida]
RETURN
    IF (
        ISBLANK ( _av ),
        SWITCH ( _l, 1, "Sem aferições no período.", 2, [Fonte Aferições] ),
        SWITCH (
            _l,
            1, FORMAT ( [% Aferições Conformes], "0.0%" ) & " dos " & FORMAT ( _av, "#,0" ) & " registros de " & LOWER ( [Mês de Referência] ) & " estão conformes (meta " & FORMAT ( [Meta Aferições Conformes], "0%" ) & ").",
            2, IF ( _nc = 0, "Nenhuma não conformidade no período.",
                   "As " & _nc & " não conformidades vêm de " & [Equipamentos com NC] & " equipamentos. Reincidentes: " & [Reincidentes · Lista] & "." ),
            3, "Rotina: " & [Aferições Realizadas] & " de " & [Aferições Programadas] & " aferições programadas feitas no dia certo (" & FORMAT ( [Aderência à Rotina], "0%" ) & "); "
                   & [Rotina · Com Atraso] & " com atraso e " & [Rotina · Não Feitas] & " não feitas"
                   & IF ( _tp <> "", "; pior pontualidade em " & _tp & ".", "." ),
            4, IF ( _cruz <> "", "Calibração externa vencida em equipamento com não conformidade: " & _cruz & ".",
                   "Nenhum equipamento com não conformidade está com a calibração externa vencida." )
        )
    )""", None, A)

# ── 17 cabeçalho e placar por laboratório (v16) ───────────────────────────
H = "17 Cabeçalho e laboratório"
add("Cabeçalho · Período", """"Período:  " & [Mês de Referência]""", None, H)
add("Cabeçalho · Atualização", """"Atualizado em  " & [Última Atualização]""", None, H)
add("Nível Laboratório", """VAR _d = [Disponibilidade no Mês]
VAR _md = [Param Meta Disponibilidade]
VAR _c = [% Calibrações em Dia]
VAR _mc = [Meta Calibração em Dia]
VAR _nd = IF ( ISBLANK ( _d ), BLANK (), IF ( _d >= _md, 1, IF ( _d >= _md * 0.9, 2, 3 ) ) )
VAR _nc = IF ( ISBLANK ( _c ), BLANK (), IF ( _c >= _mc, 1, IF ( _c >= _mc * 0.9, 2, 3 ) ) )
RETURN
    MAX ( _nd, _nc )""", "0", H, True)
add("Situação Laboratório", nivel_txt("[Nível Laboratório]"), None, H)
add("Cor Situação Laboratório", nivel_cor("[Nível Laboratório]"), None, H, True)
add("Laboratório Aparece", "IF ( [Total Instrumentos] > 0 || [Total Equipamentos] > 0, 1, 0 )", "0", H, True)

# barra de progresso em texto (a barra em gráfico não aparecia: ficava um retângulo branco)
def barra_txt(nome, base):
    add(nome, """VAR _p = COALESCE ( [%s], 0 )
VAR _n = ROUND ( _p * 24, 0 )
RETURN
    REPT ( "█", _n ) & REPT ( "░", 24 - _n )""" % base, None, "18 Barras", True)
for b in ("Disponibilidade", "Calibração", "Aferições", "Inspeções", "Aderência"):
    barra_txt("Barra Texto " + b, "Barra " + b)
