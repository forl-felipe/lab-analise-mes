# -*- coding: utf-8 -*-
"""v16 — páginas no desenho do mockup aprovado, páginas mais altas (1920×1440, ajuste à largura)."""
import os, sys, json, glob, shutil, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pbir import *

R = "/home/user/lab-analise-mes/Painel Samarco/Painel.Report"
P = R + "/definition/pages"
ALT = 1440

def ler(f): return json.load(open(f, encoding="utf-8"))
def gravar(f, o): json.dump(o, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
def nome_pag(p): return ler("%s/%s/page.json" % (P, p))["displayName"]
paginas = {nome_pag(os.path.basename(d.rstrip("/"))): os.path.basename(d.rstrip("/")) for d in glob.glob(P + "/*/") if os.path.exists(d + "page.json")}
VISAO, EQUIP, CALIB = paginas["VISÃO GERAL"], paginas["EQUIPAMENTOS & PARADAS"], paginas["CALIBRAÇÕES"]
INTER, INSP, NOTAS, SOBRE = paginas["INTERVENÇÕES"], paginas["INSPEÇÕES"], paginas["NOTAS / OMs"], paginas["SOBRESSALENTES"]
AFER, AFERD = paginas["AFERIÇÕES"], paginas["AFERIÇÕES · DETALHE"]

def visuais(pid):
    for f in glob.glob("%s/%s/visuals/*/visual.json" % (P, pid)):
        yield f, ler(f)
def pos(v): p = v["position"]; return (round(p["x"]), round(p["y"]), round(p["width"]), round(p["height"]))

# ── moldura-modelo (da Visão Geral atual) ─────────────────────────────────
MOLDURA = {}
for f, v in visuais(VISAO):
    t = v["visual"]["visualType"]; x, y, w, h = pos(v)
    if (t, x) in (("textbox", 16), ("textbox", 248)) and y in (ALT - 40, ALT - 36): y = 1040 if x == 16 else 1044
    chave = {("textbox", 0, 0): "lateral", ("image", 14, 22): "logo", ("textbox", 16, 70): "subtitulo",
             ("textbox", 224, 0): "cabecalho", ("textbox", 224, 88): "faixa",
             ("textbox", 16, 492): "filtros", ("textbox", 16, 1040): "powerbi", ("textbox", 248, 1044): "rodape"}.get((t, x, y))
    if chave: MOLDURA[chave] = v
MOLDURA.pop("subtitulo", None)
assert {"lateral", "logo", "cabecalho", "faixa", "filtros", "powerbi", "rodape"} <= set(MOLDURA), sorted(MOLDURA)

def texto_cab(v, titulo=None, sub=None):
    """v18: cabeçalho azul-marinho com o título em branco e sem a frase de subtítulo"""
    ps = v["visual"]["objects"]["general"][0]["properties"]["paragraphs"]
    if titulo: ps[0]["textRuns"][0]["value"] = titulo
    ps[0]["textRuns"][0]["textStyle"].update({"fontSize": "20pt", "color": "#FFFFFF"})
    del ps[1:]
    ps.insert(0, {"textRuns": [{"value": " ", "textStyle": {"fontSize": "6pt", "color": "#00335A", "fontFamily": "Segoe UI"}}]})
    v["visual"].setdefault("visualContainerObjects", {})["background"] = [{"properties": {
        "show": {"expr": {"Literal": {"Value": "true"}}},
        "color": {"solid": {"color": {"expr": {"Literal": {"Value": "'#00335A'"}}}}}}}]

def limpar(pid):
    shutil.rmtree("%s/%s/visuals" % (P, pid), ignore_errors=True); os.makedirs("%s/%s/visuals" % (P, pid))

def pagina_alta(pid):
    pj = ler("%s/%s/page.json" % (P, pid)); pj["height"] = ALT; pj["displayOption"] = "FitToWidth"
    gravar("%s/%s/page.json" % (P, pid), pj)

def moldura(pid, titulo, sub, rodape=None):
    for chave, v in MOLDURA.items():
        o = copy.deepcopy(v); o["name"] = vid(pid + "|moldura|" + chave)
        if chave == "lateral": o["position"]["height"] = ALT
        if chave == "cabecalho": texto_cab(o, titulo, sub)
        if chave == "powerbi": o["position"].update({"y": ALT - 40, "z": 7400})
        if chave == "rodape":
            o["position"]["y"] = ALT - 36
            if rodape: o["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"][0]["value"] = rodape
        d = "%s/%s/visuals/%s" % (P, pid, o["name"]); os.makedirs(d, exist_ok=True)
        gravar(d + "/visual.json", o)

NAVY, TINTA, SUAVE, AZUL, OCRE, LARANJA, CINZA = "#00335A", "#1D2B36", "#5B6B78", "#1B6FB0", "#C99400", "#E0620F", "#B5BEC4"
X0, LARG = 244, 1656

def cabecalho_cartoes(pg, periodo=True):
    if periodo:
        cartao(pg, "cab-periodo", (1150, 20, 360, 52, 400), "Cabeçalho · Período", tam=12, cor="#FFFFFF")
    cartao(pg, "cab-atualizacao", (1520, 20, 380, 52, 400), "Cabeçalho · Atualização", tam=11, cor="#D6DEE6", fonte="Segoe UI")

# ── menu: pílula (imagem) por cima, com o link; sem botão por cima ───────
MENU = [("visao", VISAO), ("equip", EQUIP), ("calib", CALIB), ("afer", AFER),
        ("inter", INTER), ("insp", INSP), ("notas", NOTAS), ("sobre", SOBRE)]
def menu(pid):
    pg = Pagina("%s/%s" % (P, pid)); y = 108
    em_afer = pid in (AFER, AFERD)
    for chave, destino in MENU:
        ativo = destino == pid or (chave == "afer" and em_afer)
        link = None if (destino == pid) else destino
        imagem(pg, "menu-img-" + chave, (10, y, 204, 36, 7100), "menu-%s-%s.png" % (chave, "on" if ativo else "off"), link=link)
        y += 40
        if chave == "afer" and em_afer:
            for sub, dest in (("afer-resumo", AFER), ("afer-detalhe", AFERD)):
                on = dest == pid
                imagem(pg, "menu-img-" + sub, (34, y, 180, 28, 7100), "menu-%s-%s.png" % (sub, "on" if on else "off"),
                       link=None if on else dest)
                y += 32

def limpar_menu_e_cabecalho(pid):
    for f, v in visuais(pid):
        t = v["visual"]["visualType"]; x, y, w, h = pos(v)
        if t in ("actionButton", "image") and x == 10 and w == 204 and 100 <= y <= 470:
            shutil.rmtree(os.path.dirname(f))
        elif t == "cardVisual" and (x, y) == (1596, 20):          # "Última Atualização" cortado
            shutil.rmtree(os.path.dirname(f))

# páginas mantidas: só menu e cabeçalho
for pid in (CALIB, INTER, INSP, NOTAS, SOBRE):
    limpar_menu_e_cabecalho(pid)
    menu(pid)
    cabecalho_cartoes(Pagina("%s/%s" % (P, pid)), periodo=False)
    for f, v in visuais(pid):
        x, y, w, h = pos(v); t = v["visual"]["visualType"]
        if t == "textbox" and (x, y) == (224, 0): texto_cab(v); gravar(f, v)
        elif t == "textbox" and (x, y) == (16, 70): shutil.rmtree(os.path.dirname(f))
    # todas as páginas com o mesmo ajuste: largura da tela (antes umas ajustavam à altura)
    pj = ler("%s/%s/page.json" % (P, pid)); pj["displayOption"] = "FitToWidth"; gravar("%s/%s/page.json" % (P, pid), pj)

# CALIBRAÇÕES: TAG na tabela "Última calibração, vencimento e situação"
for f, v in visuais(CALIB):
    if v["visual"]["visualType"] != "tableEx": continue
    pr = v["visual"]["query"]["queryState"]["Values"]["projections"]
    props = [p["field"].get("Column", {}).get("Property") for p in pr]
    if "Equipamento" in props and "Tag" not in props and any("Vencimento" in (p.get("displayName") or p["nativeQueryRef"]) or p["field"].get("Column", {}).get("Property") == "DataVencimento" for p in pr):
        pr.insert(0, proj(COL("Tag", "tbl_Calibracao"), "TAG"))
        gravar(f, v); print("TAG incluída em", v["name"])

MES0 = [f_em("mes0", "DimCalendario", "Mês Offset", ["0L"])]

# =============================================================================
# VISÃO GERAL
# =============================================================================
limpar(VISAO); pagina_alta(VISAO)
moldura(VISAO, "GESTÃO DE EQUIPAMENTOS E CALIBRAÇÕES", "Situação do mês em 4 frentes e o que precisa de decisão")
menu(VISAO)
pg = Pagina("%s/%s" % (P, VISAO))
cabecalho_cartoes(pg)
fatiador(pg, "slc-lab", (12, 516, 200, 150), COL("Laboratorio", "DimLaboratorio"))


PILARES = [
    ("disp", "DISPONIBILIDADE", "Disponibilidade no Mês", "Situação Disponibilidade", "Cor Disponibilidade",
     "Disponibilidade · Meta", "Barra Disponibilidade", "Disponibilidade · Fatos", "POR SEMANA", "card-disp.png"),
    ("calib", "CALIBRAÇÃO", "% Calibrações em Dia", "Situação Calibração", "Cor Calibração",
     "Calibração · Meta", "Barra Calibração", "Calibração · Fatos", "VENCEM NOS PRÓXIMOS", "card-emdia.png"),
    ("afer", "AFERIÇÕES", "% Aferições Conformes", "Situação Aferições", "Cor Aferições",
     "Aferições · Meta", "Barra Aferições", "Aferições · Fatos", "POR SEMANA", "card-ok.png"),
    ("insp", "INSPEÇÕES", "Inspeções no Mês", "Situação Inspeções", "Cor Inspeções",
     "Inspeções · Meta", "Barra Inspeções", "Inspeções · Fatos", "REALIZADO × ESPERADO ATÉ HOJE", "card-insp.png"),
]
YP, HP, WP, GP = 108, 450, 405, 12
for i, (k, tit, valor, sit, cor, meta, barra, fatos, ttl, ic) in enumerate(PILARES):
    x = X0 + i * (WP + GP)
    textbox(pg, "pilar-%s" % k, (x, YP, WP, HP, 1000), [[(" ", 8, "#FFFFFF", False)]], fundo="#FFFFFF", borda="#D6DEE6", sombra=True)
    imagem(pg, "pilar-ic-%s" % k, (x + 12, YP + 14, 28, 28, 1100), ic)
    textbox(pg, "pilar-tit-%s" % k, (x + 44, YP + 8, WP - 210, 40, 1100), [[(tit, 11, NAVY, True)]])
    cartao(pg, "pilar-sit-%s" % k, (x + WP - 166, YP + 6, 160, 46, 1100), sit, tam=11, cor_medida=cor)
    cartao(pg, "pilar-val-%s" % k, (x + 10, YP + 52, WP - 20, 88, 1100), valor, tam=38, cor=NAVY)
    cartao(pg, "pilar-meta-%s" % k, (x + 10, YP + 140, WP - 20, 42, 1100), meta, tam=10, cor=SUAVE, fonte="Segoe UI")
    cartao(pg, "pilar-barra-%s" % k, (x + 10, YP + 184, WP - 20, 38, 1100), barra.replace("Barra ", "Barra Texto "), tam=12, cor_medida=cor, fonte="Segoe UI", quebra=False)
    cartao(pg, "pilar-fatos-%s" % k, (x + 10, YP + 224, WP - 20, 50, 1100), fatos, tam=10, cor=TINTA, fonte="Segoe UI")
    textbox(pg, "pilar-linha-%s" % k, (x + 16, YP + 280, WP - 32, 2, 1100), [[(" ", 1, "#E6EBF0", False)]], fundo="#E6EBF0")
    cz = (x + 10, YP + 286, WP - 20, HP - 294, 1100)
    if k == "disp":
        colunas(pg, "pilar-graf-disp", cz, ttl, COL("Semana Rótulo", "DimCalendario"),
                [(MEAS("Disponibilidade Semana"), "Disponibilidade")], cores=[AZUL], filtros=MES0, fundo=False, borda=False, tam_titulo=9)
    elif k == "calib":
        colunas(pg, "pilar-graf-calib", cz, ttl, None,
                [(MEAS("Vencem em até 30 dias"), "até 30 dias"), (MEAS("Vencem em 31 a 60 dias"), "31 a 60 dias"),
                 (MEAS("Vencem em 61 a 90 dias"), "61 a 90 dias")], cores=[OCRE, AZUL, "#36A9C9"], legenda=True, fundo=False, borda=False, tam_titulo=9)
    elif k == "afer":
        colunas(pg, "pilar-graf-afer", cz, ttl, COL("Semana Rótulo", "DimCalendario"),
                [(MEAS("% Aferições Conformes"), "Conformes")], cores=[AZUL], filtros=MES0, fundo=False, borda=False, tam_titulo=9)
    else:
        colunas(pg, "pilar-graf-insp", cz, ttl, None,
                [(MEAS("Inspeções no Mês"), "Realizadas"), (MEAS("Inspeções Esperadas até Hoje"), "Esperado até hoje")],
                cores=[LARANJA, CINZA], legenda=True, fundo=False, borda=False, tam_titulo=9)

YS = YP + HP + 12            # 642
G = "tbl_Grupos"
tabela(pg, "saude-grupos", (X0, YS, 1060, 400, 1000), "SAÚDE POR GRUPO OPERACIONAL",
       [(COL("Grupo Operacional", G), "Grupo"), (MEAS("Situação do Grupo"), "Situação"),
        (MEAS("Equip. no Grupo"), "Equip."), (MEAS("Disponibilidade do Grupo"), "Disponibilidade"),
        (MEAS("Calibração do Grupo"), "Calibração em dia"), (MEAS("Aferição do Grupo"), "Aferição conforme"),
        (MEAS("Observação do Grupo"), "Observação")],
       cores={1: "Cor Situação do Grupo"}, larguras={0: 190, 1: 110, 2: 55, 3: 105, 4: 115, 5: 115, 6: 330},
       filtros=[f_igual_medida("grupo-aparece", "Grupo Aparece")],
       subtitulo="Cada grupo mostra só o que se aplica a ele: disponibilidade para amostragem; calibração e aferição para laboratório",
       ordem=(COL("Grupo Operacional", G), "Ascending"), fonte=11)
tabela(pg, "placar-lab", (X0 + 1072, YS, LARG - 1072, 400, 1000), "POR LABORATÓRIO",
       [(COL("Laboratorio", "DimLaboratorio"), "Lab."), (MEAS("Situação Laboratório"), "Situação"),
        (MEAS("Disponibilidade no Mês"), "Disponib."), (MEAS("% Calibrações em Dia"), "Calib. em dia"),
        (MEAS("Calibracoes Vencidas", "_Medidas"), "Vencidas"), (MEAS("Vencem em até 30 dias"), "Vencem 30 d")],
       cores={1: "Cor Situação Laboratório"}, larguras={0: 60, 1: 110, 2: 85, 3: 95, 4: 75, 5: 90},
       filtros=[f_igual_medida("lab-aparece", "Laboratório Aparece")],
       subtitulo="As mesmas frentes por laboratório; clique para filtrar a página", ordem=(COL("Laboratorio", "DimLaboratorio"), "Ascending"), fonte=12)
YD = YS + 412                # 1054
tabela(pg, "pontos", (X0, YD, LARG, ALT - 56 - YD, 1000), "PONTOS QUE PEDEM DECISÃO",
       [(MEAS("Ponto Nº"), "Nº"), (COL("Tema", "tbl_Pontos"), "Tema"), (MEAS("Ponto Gravidade"), "Gravidade"),
        (MEAS("Ponto Título"), "O quê"), (MEAS("Ponto Detalhe"), "Por quê"), (MEAS("Ponto Ação"), "Ação sugerida")],
       cores={2: "Cor Ponto Gravidade"}, larguras={0: 40, 1: 170, 2: 110, 3: 360, 4: 460, 5: 470},
       filtros=[f_igual_medida("ponto-ativo", "Ponto Ativo")],
       subtitulo="Só o que está fora da meta e tem ação; o detalhe fica na página de cada tema",
       ordem=(MEAS("Ponto Nº"), "Ascending"), fonte=12)

# =============================================================================
# EQUIPAMENTOS & PARADAS
# =============================================================================
limpar(EQUIP); pagina_alta(EQUIP)
moldura(EQUIP, "EQUIPAMENTOS & PARADAS", "Disponibilidade, paradas e situação de cada equipamento em uma página só")
menu(EQUIP)
pg = Pagina("%s/%s" % (P, EQUIP))
cabecalho_cartoes(pg)
fatiador(pg, "slc-lab", (12, 516, 200, 130), COL("Laboratorio", "DimLaboratorio"))
fatiador(pg, "slc-grupo", (12, 654, 200, 200), COL("Grupo Operacional", "tbl_Equipamentos"))
fatiador(pg, "slc-status", (12, 862, 200, 130), COL("Status Operacional", "tbl_Equipamentos"))
KP = [("Total Equipamentos", "Equipamentos", "_Medidas", "card-equip.png"),
      ("Equipamentos Operantes", "Operantes", "_Medidas", "card-ok.png"),
      ("Disponibilidade no Mês", "Disponibilidade no mês", M, "card-disp.png"),
      ("Horas Paradas (mês)", "Horas paradas no mês", M, "card-horas.png"),
      ("Paradas no Mês", "Paradas no mês", M, "card-parada.png"),
      ("Paradas em Aberto", "Paradas em aberto", M, "card-critico.png")]
WK = (LARG - 5 * 12) / 6
for i, (m, n, t, ic) in enumerate(KP):
    x = round(X0 + i * (WK + 12))
    kpi(pg, "kpi-%d" % i, (x, 104, round(WK), 110, 3000), m, n, t)
    imagem(pg, "kpi-ic-%d" % i, (x + round(WK) - 48, 116, 36, 36, 3100), ic)
colunas(pg, "disp-grupo", (X0, 226, 820, 420, 1000), "DISPONIBILIDADE POR GRUPO × META",
        COL("Grupo Operacional", "tbl_Grupos"), [(MEAS("Disponibilidade no Mês"), "Disponibilidade")],
        tipo="clusteredBarChart", cor_medida="Cor Barra Disponibilidade Grupo",
        subtitulo="Mês corrente · azul na meta, ocre até 10% abaixo, laranja abaixo disso",
        ordem=(MEAS("Disponibilidade no Mês"), "Ascending"))
colunas(pg, "horas-equip", (X0 + 832, 226, LARG - 832, 420, 1000), "HORAS PARADAS POR EQUIPAMENTO",
        COL("Equipamento", "tbl_Equipamentos"), [(MEAS("Horas Paradas no Mês"), "Horas paradas")],
        tipo="clusteredBarChart", cor_medida="Cor Horas Paradas",
        subtitulo="Mês corrente · laranja = equipamento com parada ainda em aberto",
        filtros=[f_maior_medida("horas>0", "Horas Paradas no Mês")],
        ordem=(MEAS("Horas Paradas no Mês"), "Descending"))
colunas(pg, "paradas-dia", (X0, 658, LARG, 290, 1000), "LINHA DO TEMPO DAS PARADAS · HORAS POR DIA",
        COL("Data", "DimCalendario"), [(MEAS("Horas Paradas", "_Medidas"), "Horas paradas")],
        tipo="columnChart", serie=COL("Situação", "Tbl_Paradas"), legenda=True, filtros=MES0,
        subtitulo="Cada coluna é um dia do mês; a cor separa paradas resolvidas das que seguem em aberto")
tabela(pg, "tab-equip", (X0, 960, LARG, ALT - 56 - 960, 1000), "EQUIPAMENTOS",
       [(COL("Tag", "tbl_Equipamentos"), "Tag"), (COL("Equipamento", "tbl_Equipamentos"), "Equipamento"),
        (COL("Grupo Operacional", "tbl_Equipamentos"), "Grupo"), (COL("Status Operacional", "tbl_Equipamentos"), "Status"),
        (MEAS("Disponibilidade no Mês"), "Disponib. no mês"), (MEAS("Horas Paradas no Mês"), "Horas paradas no mês"),
        (MEAS("Última Parada"), "Última parada"), (MEAS("Situação da Última Parada"), "Situação")],
       cores={7: "Cor Situação da Parada"}, larguras={0: 130, 1: 300, 2: 190, 3: 120, 4: 130, 5: 160, 6: 330, 7: 180},
       subtitulo="Clique num equipamento para filtrar os gráficos acima",
       ordem=(MEAS("Horas Paradas no Mês"), "Descending"), fonte=11)

# =============================================================================
# AFERIÇÕES · RESUMO DO MÊS
# =============================================================================
TA = "tbl_Afericoes"
ROD_AF = "Aferições: planilhas mensais CALIBRAÇÃO (aba oculta BD_Afericoes) na pasta do SharePoint  ·  Base_PowerBI_Gestao_Equipamentos.xlsx   ·   v18"
def lateral_afer(pg, extra):
    fatiador(pg, "slc-mes", (12, 532, 200, 150), COL("Ano Mês Nome", "DimCalendario"),
             filtros=[f_entre("mes-12", "DimCalendario", "Mês Offset", "-12L", "0L")])
    y = 690
    for seed, campo, h in extra:
        fatiador(pg, seed, (12, y, 200, h), campo); y += h + 8

def mini(pg, k, x, w, titulo, valor, meta=None, barra=None, cor=None, fatos=None):
    y0 = 104
    textbox(pg, "mini-%s" % k, (x, y0, w, 190, 1000), [[(" ", 8, "#FFFFFF", False)]], fundo="#FFFFFF", borda="#D6DEE6", sombra=True)
    textbox(pg, "mini-tit-%s" % k, (x + 8, y0 + 6, w - 16, 34, 1100), [[(titulo, 10, NAVY, True)]])
    cartao(pg, "mini-val-%s" % k, (x + 10, y0 + 40, w - 20, 72, 1100), valor, tam=30, cor=NAVY)
    if meta: cartao(pg, "mini-meta-%s" % k, (x + 10, y0 + 112, w - 20, 40, 1100), meta, tam=10, cor=SUAVE, fonte="Segoe UI")
    if barra: cartao(pg, "mini-barra-%s" % k, (x + 10, y0 + 154, w - 20, 34, 1100), barra.replace("Barra ", "Barra Texto "), tam=11, cor_medida=cor, fonte="Segoe UI", quebra=False)
    if fatos: cartao(pg, "mini-fatos-%s" % k, (x + 10, y0 + 112, w - 20, 70, 1100), fatos, tam=10, cor=TINTA, fonte="Segoe UI")

limpar(AFER); pagina_alta(AFER)
moldura(AFER, "AFERIÇÕES E COMPARATIVOS", "Como estão as aferições do mês e onde agir", rodape=ROD_AF)
menu(AFER)
pg = Pagina("%s/%s" % (P, AFER))
cabecalho_cartoes(pg)
lateral_afer(pg, [("slc-tipo", COL("Tipo", TA), 120), ("slc-ensaio", COL("Ensaio Curto", "DimEnsaio"), 240)])
W4 = (LARG - 3 * 12) / 4
mini(pg, "conf", X0, round(W4), "CONFORMIDADE", "% Aferições Conformes", "Aferições · Meta", "Barra Aferições", "Cor Aferições")
mini(pg, "ader", round(X0 + W4 + 12), round(W4), "ADERÊNCIA À ROTINA", "Aderência à Rotina", "Aderência · Meta", "Barra Aderência", "Cor Aderência")
mini(pg, "nc", round(X0 + 2 * (W4 + 12)), round(W4), "NÃO CONFORMIDADES", "NC Aferição (cartão)", fatos="Aferições · Fatos")
mini(pg, "reinc", round(X0 + 3 * (W4 + 12)), round(W4), "EQUIPAMENTOS REINCIDENTES", "Reincidentes", fatos="Reincidentes · Lista")
tabela(pg, "ensaios", (X0, 306, 1000, 380, 1000), "SITUAÇÃO POR ENSAIO",
       [(COL("Ensaio", "DimEnsaio"), "Ensaio"), (MEAS("Situação Ensaio"), "Situação"),
        (MEAS("% Aferições Conformes"), "Conformes"), (MEAS("Aderência à Rotina"), "Rotina"),
        (COL("Frequência", "DimEnsaio"), "Rotina (frequência)"), (MEAS("Mapa · Dias Não Feitos"), "Não feitos"),
        (MEAS("NC Aferição"), "NC"), (MEAS("Último Registro"), "Último")],
       cores={1: "Cor Situação Ensaio"}, larguras={0: 210, 1: 110, 2: 100, 3: 80, 4: 180, 5: 90, 6: 60, 7: 80},
       filtros=[f_igual_medida("ens-aparece", "Ensaio Aparece")],
       subtitulo="Ordenado pela ordem das abas; clique num ensaio para filtrar a página",
       ordem=(COL("Ensaio", "DimEnsaio"), "Ascending"), fonte=12)
tabela(pg, "resumo", (X0 + 1012, 306, LARG - 1012, 380, 1000), "RESUMO AUTOMÁTICO DO MÊS",
       [(COL("Linha", "tbl_Linhas"), "#"), (MEAS("Resumo Aferições"), "O que o mês mostra")],
       larguras={0: 30, 1: 560}, subtitulo="Texto gerado pelo painel a cada atualização",
       ordem=(COL("Linha", "tbl_Linhas"), "Ascending"), fonte=12)
matriz(pg, "mapa", (X0, 698, LARG, ALT - 56 - 698, 1000), "MAPA DE AFERIÇÕES · ENSAIO × DIA",
       COL("Ensaio Curto", "DimEnsaio"), COL("Dia", "DimCalendario"), MEAS("Mapa"), "Mapa · Fundo", "Mapa · Fonte",
       subtitulo="✔ azul = aferição registrada no dia   ·   ✖ laranja = dia da rotina sem registro até a próxima data da rotina (prazo vencido)   ·   em branco = sem rotina, ou feito dentro do prazo em outro dia", largura_col=46)

# =============================================================================
# AFERIÇÕES · DETALHE POR ENSAIO
# =============================================================================
limpar(AFERD); pagina_alta(AFERD)
moldura(AFERD, "AFERIÇÕES · DETALHE POR ENSAIO", "Carta de controle, comparativos e registros do ensaio escolhido na lateral", rodape=ROD_AF)
menu(AFERD)
pg = Pagina("%s/%s" % (P, AFERD))
cabecalho_cartoes(pg)
lateral_afer(pg, [("slc-ensaio", COL("Ensaio Curto", "DimEnsaio"), 170), ("slc-equip", COL("Equipamento", TA), 190),
                  ("slc-param", COL("Parâmetro", TA), 170)])
textbox(pg, "carta-fundo", (X0, 104, LARG, 580, 900), [[(" ", 8, "#FFFFFF", False)]], fundo="#FFFFFF", borda="#D6DEE6", sombra=True)
cartao(pg, "carta-titulo", (X0 + 16, 108, LARG - 32, 52, 1000), "Carta · Título", tam=13, cor=NAVY)
linhas(pg, "carta", (X0 + 10, 162, LARG - 20, 514, 1000), None, COL("Data", TA),
       [(MEAS("Carta · Valor"), "Valor medido"), (MEAS("Carta · Lim. Inferior"), "Limite inferior"),
        (MEAS("Carta · Lim. Superior"), "Limite superior"), (MEAS("Carta · Nominal"), "Nominal / zero")],
       [AZUL, LARANJA, LARANJA, CINZA])
tabela(pg, "resultados", (X0, 696, 780, ALT - 56 - 696, 1000), "RESULTADOS POR EQUIPAMENTO",
       [(COL("Equipamento", TA), "Equipamento"), (COL("Parâmetro", TA), "Parâmetro"),
        (MEAS("Última Aferição"), "Última"), (MEAS("Último Valor"), "Último valor"),
        (MEAS("Média do Período"), "Média"), (MEAS("Faixa Aceita"), "Faixa aceita"),
        (MEAS("Último Resultado"), "Último resultado"), (MEAS("Registros no Período"), "Nº")],
       cores={6: "Cor Último Resultado"}, larguras={0: 150, 1: 150, 2: 55, 3: 75, 4: 65, 5: 105, 6: 95, 7: 35},
       filtros=[f_igual_medida("res-periodo", "Linha no Período"), f_em("res-tipo", TA, "Tipo", ["'Calibração'", "'Comparativo'"])],
       subtitulo="Última aferição, último valor e média do período de cada equipamento",
       ordem=(COL("Equipamento", TA), "Ascending"), fonte=11)
tabela(pg, "registros", (X0 + 792, 696, LARG - 792, ALT - 56 - 696, 1000), "REGISTROS DO PERÍODO",
       [(COL("Data", TA), "Data"), (COL("Equipamento", TA), "Equipamento"), (COL("Parâmetro", TA), "Parâmetro"),
        (COL("Valor", TA), "Valor"), (COL("Referência", TA), "Referência"), (COL("Diferença", TA), "Diferença"),
        (COL("Resultado", TA), "Resultado"), (COL("Responsável", TA), "Responsável")],
       cores={6: "Cor Resultado"}, larguras={0: 85, 1: 180, 2: 190, 3: 70, 4: 80, 5: 75, 6: 100, 7: 100},
       filtros=[f_igual_medida("linha-periodo", "Linha no Período")],
       ordem=(COL("Data", TA), "Descending"), fonte=11)

# rodapé v16 nas páginas mantidas
for pid in (CALIB, INTER, INSP, NOTAS, SOBRE, VISAO, EQUIP):
    for f, v in visuais(pid):
        if v["visual"]["visualType"] == "textbox":
            s = json.dumps(v, ensure_ascii=False)
            if "v15" in s or "v16" in s or "v17" in s: gravar(f, json.loads(s.replace("v15", "v18").replace("v16", "v18").replace("v17", "v18")))
print("v18 montada")
