# -*- coding: utf-8 -*-
"""v19 — página AFERIÇÕES: mapa pela regra do DIA CERTO, legenda e cola de regras.
Mexe só na página Aferições (e no número da versão no rodapé)."""
import os, sys, json, glob, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pbir import *

R = "/home/user/lab-analise-mes/Painel Samarco/Painel.Report"
P = R + "/definition/pages"
AFER = "8871b58324bf213b8c74"
ALT = 1900
X0, LARG = 244, 1656
NAVY, TINTA, SUAVE, AZUL, OCRE, LARANJA = "#00335A", "#1D2B36", "#5B6B78", "#1B6FB0", "#C99400", "#E0620F"

def ler(f): return json.load(open(f, encoding="utf-8"))
def gravar(f, o): json.dump(o, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
def visuais(pid):
    for f in sorted(glob.glob("%s/%s/visuals/*/visual.json" % (P, pid))):
        yield f, ler(f)
def pos(v): p = v["position"]; return (round(p["x"]), round(p["y"]), round(p["width"]), round(p["height"]))

pasta = "%s/%s" % (P, AFER)
assert ler(pasta + "/page.json")["displayName"] == "AFERIÇÕES"
pg = Pagina(pasta)

# ── 1. sai o mapa antigo; página mais alta; moldura acompanha ─────────────
for f, v in visuais(AFER):
    t = v["visual"]["visualType"]; x, y, w, h = pos(v)
    if t == "pivotTable":
        shutil.rmtree(os.path.dirname(f)); continue
    if t == "textbox" and (x, y) == (0, 0):                 # lateral azul
        v["position"]["height"] = ALT; gravar(f, v)
    elif t == "textbox" and x == 16 and y >= 1300:          # "Power BI" no pé da lateral
        v["position"]["y"] = ALT - 40; gravar(f, v)
    elif t == "textbox" and x == 248 and y >= 1300:         # rodapé
        v["position"]["y"] = ALT - 36; gravar(f, v)
pj = ler(pasta + "/page.json"); pj["height"] = ALT; pj["displayOption"] = "FitToWidth"; gravar(pasta + "/page.json", pj)

# ── 2. cartão de rotina e situação por ensaio pela nova regra ─────────────
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "_mini.py"), encoding="utf-8").read())
W4 = (LARG - 3 * 12) / 4
mini(pg, "ader", round(X0 + W4 + 12), round(W4), "ROTINA · FEITAS NO DIA CERTO", "Aderência à Rotina", "Aderência · Meta",
     "Barra Aderência", "Cor Aderência")
tabela(pg, "ensaios", (X0, 306, 1000, 380, 1000), "SITUAÇÃO POR ENSAIO",
       [(COL("Ensaio", "DimEnsaio"), "Ensaio"), (MEAS("Situação Ensaio"), "Situação"),
        (MEAS("% Aferições Conformes"), "Conformes"), (MEAS("Aderência à Rotina"), "No dia certo"),
        (MEAS("Rotina · Com Atraso"), "Com atraso"), (MEAS("Mapa · Dias Não Feitos"), "Não feitos"),
        (MEAS("NC Aferição"), "NC"), (MEAS("Último Registro"), "Último")],
       cores={1: "Cor Situação Ensaio"}, larguras={0: 220, 1: 120, 2: 100, 3: 110, 4: 100, 5: 100, 6: 70, 7: 90},
       filtros=[f_igual_medida("ens-aparece", "Ensaio Aparece")],
       subtitulo="Ordenado pela ordem das abas; clique num ensaio para filtrar a página",
       ordem=(COL("Ensaio", "DimEnsaio"), "Ascending"), fonte=12)

# ── 3. mapa ───────────────────────────────────────────────────────────────
Y_MAPA, H_MAPA = 698, 620
nome_mapa = matriz(pg, "mapa19", (X0, Y_MAPA, LARG, H_MAPA, 1000), "MAPA DE AFERIÇÕES · DIA CERTO",
       COL("Linha", "tbl_MapaLinhas"), COL("Dia", "DimCalendario"), MEAS("Mapa"), "Mapa · Fundo", "Mapa · Fonte",
       subtitulo="Cada coluna é um dia do mês e a marca fica no dia certo de cada ensaio. Clique no nome de um ensaio para ver a regra dele na tabela abaixo.",
       largura_col=46)
f = "%s/visuals/%s/visual.json" % (pasta, nome_mapa); v = ler(f)
o = v["visual"]["objects"]
o["values"][0]["properties"]["fontSize"] = D(15)
o["rowHeaders"][0]["properties"].update({"fontSize": D(14), "fontColor": COR(TINTA)})
o["columnHeaders"][0]["properties"].update({"fontSize": D(13), "fontColor": COR(TINTA), "backColor": COR("#FFFFFF")})
g = o["grid"][0]["properties"]; g["rowPadding"] = D(14); g["gridVerticalWeight"] = D(3); g["gridHorizontalWeight"] = D(3)
gravar(f, v)

# ── 4. legenda ────────────────────────────────────────────────────────────
Y_LEG = Y_MAPA + H_MAPA + 10
textbox(pg, "leg-fundo", (X0, Y_LEG, LARG, 72, 900), [[(" ", 8, "#FFFFFF", False)]], fundo="#FFFFFF", borda="#D6DEE6")
LEG = [("✔", AZUL, "#FFFFFF", "Feito no dia certo", "registro no próprio dia da rotina"),
       (" ", "#F3CF5B", "#F3CF5B", "Feito com atraso", "depois do dia certo, antes da próxima data"),
       ("✔", "#DCEAF6", AZUL, "Feito fora do dia programado", "dia em que a atrasada foi feita, ou extra"),
       ("✖", LARANJA, "#FFFFFF", "Não feito", "passou o dia certo e não há registro"),
       ("○", "#FFFFFF", "#8C98A2", "Programado", "dia certo que ainda vai chegar (ou hoje)"),
       ("½", "#F8D9C4", "#9A3F05", "Incompleto", "Blaine: só 1 dos 2 turnos")]
wl = (LARG - 24) / len(LEG)
for i, (sim, fundo, cor, tit, desc) in enumerate(LEG):
    x = round(X0 + 12 + i * wl)
    n = textbox(pg, "leg-cor-%d" % i, (x, Y_LEG + 14, 44, 44, 1000), [[(sim, 14, cor, True)]], fundo=fundo,
                borda="#D6DEE6" if fundo == "#FFFFFF" else None)
    fv = "%s/visuals/%s/visual.json" % (pasta, n); vv = ler(fv)
    for par in vv["visual"]["objects"]["general"][0]["properties"]["paragraphs"]: par["horizontalTextAlignment"] = "center"
    if fundo == "#FFFFFF": vv["visual"]["visualContainerObjects"]["border"][0]["properties"]["radius"] = D(4)
    gravar(fv, vv)
    textbox(pg, "leg-txt-%d" % i, (x + 50, Y_LEG + 8, round(wl) - 58, 58, 1000),
            [[(tit, 11, TINTA, True)], [(desc, 9, SUAVE, False)]])

# ── 5. cola de regras (filtrada pelo clique no mapa) ──────────────────────
Y_REG = Y_LEG + 84
nome_regras = tabela(pg, "regras", (X0, Y_REG, LARG, ALT - 56 - Y_REG, 1000), "REGRAS DE CADA ENSAIO · O DIA CERTO",
       [(COL("Linha", "tbl_MapaLinhas"), "Ensaio"), (COL("Regra", "tbl_MapaLinhas"), "Quando deve ser feito"),
        (MEAS("Regra · Dias Certos"), "Dias certos no mês"), (MEAS("Regra · No Dia Certo"), "Resultado no mês")],
       larguras={0: 200, 1: 520, 2: 400, 3: 500},
       filtros=[f_igual_medida("regra-aparece", "Regra Aparece")],
       subtitulo="Feito depois do dia certo conta como atraso (amarelo) até a véspera da próxima data da rotina; depois disso o dia certo fica ✖ não feito.",
       ordem=(COL("Linha", "tbl_MapaLinhas"), "Ascending"), fonte=12)

# ── 6. clique no mapa filtra só a cola de regras ──────────────────────────
alvos = [ler(f)["name"] for f, _ in visuais(AFER)]
pj = ler(pasta + "/page.json")
pj["visualInteractions"] = [{"source": nome_mapa, "target": t, "type": "NoFilter"}
                            for t in alvos if t not in (nome_mapa, nome_regras)]
gravar(pasta + "/page.json", pj)

# ── 7. versão no rodapé ───────────────────────────────────────────────────
for f in glob.glob(P + "/*/visuals/*/visual.json"):
    s = open(f, encoding="utf-8").read()
    if '"textbox"' in s and "v18" in s:
        open(f, "w", encoding="utf-8").write(s.replace("v18", "v19"))
print("v19 montada · mapa", nome_mapa, "· regras", nome_regras)
