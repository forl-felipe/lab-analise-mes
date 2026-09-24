# -*- coding: utf-8 -*-
"""v19.1 — só dimensionamento do mapa e legenda em uma caixa de texto (sem mexer em medidas)."""
import os, sys, json, glob, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pbir import *

P = "/home/user/lab-analise-mes/Painel Samarco/Painel.Report/definition/pages"
AFER = "8871b58324bf213b8c74"
pasta = "%s/%s" % (P, AFER)
X0, LARG = 244, 1656
TINTA, SUAVE, AZUL, LARANJA = "#1D2B36", "#5B6B78", "#1B6FB0", "#E0620F"
def ler(f): return json.load(open(f, encoding="utf-8"))
def gravar(f, o): json.dump(o, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

# ── mapa: sem quebra de texto, largura fixa da coluna dos ensaios ─────────
fm = glob.glob(pasta + "/visuals/*/visual.json")
mapa = [f for f in fm if ler(f)["visual"]["visualType"] == "pivotTable"]
assert len(mapa) == 1
f = mapa[0]; v = ler(f); vis = v["visual"]; o = vis["objects"]
linha = vis["query"]["queryState"]["Rows"]["projections"][0]
linha["displayName"] = "Ensaio"; linha["nativeQueryRef"] = "Ensaio"
o["values"][0]["properties"].update({"fontSize": D(12), "wordWrap": B(False)})
o["rowHeaders"] = [{"properties": {"fontSize": D(13), "fontColor": COR(TINTA), "wordWrap": B(False)}}]
o["columnHeaders"] = [{"properties": {"fontSize": D(12), "fontColor": COR(TINTA), "backColor": COR("#FFFFFF"),
                                      "autoSizeColumnWidth": B(False), "alignment": S("Center"), "wordWrap": B(False)}}]
g = o["grid"][0]["properties"]; g["rowPadding"] = D(12)
o["columnWidth"] = [{"properties": {"value": D(46)}, "selector": {"metadata": qref(vis["query"]["queryState"]["Values"]["projections"][0]["field"])}},
                    {"properties": {"value": D(170)}, "selector": {"metadata": qref(linha["field"])}}]
gravar(f, v)
Y_LEG = round(v["position"]["y"] + v["position"]["height"]) + 10

# ── legenda: uma caixa de texto só (antes eram 12 caixinhas) ──────────────
for f in fm:
    if not os.path.exists(f): continue
    w = ler(f); p = w["position"]
    if w["visual"]["visualType"] == "textbox" and Y_LEG - 2 <= p["y"] <= Y_LEG + 20:
        shutil.rmtree(os.path.dirname(f))
pg = Pagina(pasta)
ITENS = [("■", AZUL, "Feito no dia certo"), ("■", "#F3CF5B", "Feito com atraso"),
         ("■", "#A9CBEA", "Feito fora do dia programado"), ("■", LARANJA, "Não feito"),
         ("○", "#8C98A2", "Programado"), ("■", "#F2B48C", "Incompleto (Blaine: 1 dos 2 turnos)")]
p1 = []
for sim, cor, txt in ITENS:
    p1 += [(sim + " ", 20, cor, False), (txt + "        ", 12, TINTA, True)]
p2 = [("Amarelo: ", 11, "#8A6A00", True),
      ("feito depois do dia certo, antes da próxima data da rotina (o dia em que foi feito fica azul-claro).     ", 11, SUAVE, False),
      ("Laranja ✖: ", 11, LARANJA, True),
      ("passou o dia certo e não há registro até a próxima data.", 11, SUAVE, False)]
n = textbox(pg, "legenda-unica", (X0, Y_LEG, LARG, 72, 1000), [p1, p2], fundo="#FFFFFF", borda="#D6DEE6")
print("legenda", n, "em y", Y_LEG)

# interações: o clique no mapa só filtra a cola de regras (tira alvos que não existem mais)
pj = ler(pasta + "/page.json")
nomes = {ler(f)["name"]: ler(f)["visual"]["visualType"] for f in glob.glob(pasta + "/visuals/*/visual.json")}
fonte = [k for k, t in nomes.items() if t == "pivotTable"][0]
regras = [k for k, t in nomes.items() if t == "tableEx" and "REGRAS" in json.dumps(ler(glob.glob(pasta + "/visuals/%s/visual.json" % k)[0]), ensure_ascii=False)][0]
pj["visualInteractions"] = [{"source": fonte, "target": k, "type": "NoFilter"} for k in sorted(nomes) if k not in (fonte, regras)]
gravar(pasta + "/page.json", pj)
print("interações:", len(pj["visualInteractions"]), "alvos; regras =", regras)
