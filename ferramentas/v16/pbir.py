# -*- coding: utf-8 -*-
"""Construtores de visual.json (PBIR) usados na v15 — só padrões já presentes no relatório
ou formatos consagrados do PBIR (literal / medida em fx)."""
import json, os, hashlib

SCH = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.12.0/schema.json"
M = "_Painel"            # tabela de medidas da v15

def vid(seed): return hashlib.md5(seed.encode("utf-8")).hexdigest()[:20]
def LIT(v): return {"expr": {"Literal": {"Value": v}}}
def S(t): return LIT("'%s'" % t.replace("'", "''"))
def B(b): return LIT("true" if b else "false")
def D(n): return LIT("%sD" % n)
def COR(c): return {"solid": {"color": S(c)}}
def MEAS(p, t=M): return {"Measure": {"Expression": {"SourceRef": {"Entity": t}}, "Property": p}}
def COL(p, t): return {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": p}}
def CORM(p, t=M): return {"solid": {"color": {"expr": MEAS(p, t)}}}

def proj(f, nome=None, ativo=False):
    k = "Measure" if "Measure" in f else "Column"
    t = f[k]["Expression"]["SourceRef"]["Entity"]; p = f[k]["Property"]
    d = {"field": f, "queryRef": "%s.%s" % (t, p), "nativeQueryRef": nome or p}
    if nome: d["displayName"] = nome
    if ativo: d["active"] = True
    return d

def qref(f):
    k = "Measure" if "Measure" in f else "Column"
    return "%s.%s" % (f[k]["Expression"]["SourceRef"]["Entity"], f[k]["Property"])

class Pagina:
    def __init__(self, pasta):
        self.pasta = pasta; self.z = 1000; self.n = 0
    def grava(self, seed, pos, visual, filtros=None, how=None):
        self.n += 1; self.z += 10
        nome = vid(self.pasta + "|" + seed)
        o = {"$schema": SCH, "name": nome,
             "position": {"x": pos[0], "y": pos[1], "z": pos[4] if len(pos) > 4 else self.z,
                          "width": pos[2], "height": pos[3], "tabOrder": self.z},
             "visual": visual}
        if filtros: o["filterConfig"] = {"filters": filtros}
        if how: o["howCreated"] = how
        d = os.path.join(self.pasta, "visuals", nome); os.makedirs(d, exist_ok=True)
        json.dump(o, open(d + "/visual.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        return nome

# ── objetos de contêiner ────────────────────────────────────────────────
def vc(titulo=None, fundo=None, borda=None, sombra=False, subtitulo=None, tam=13):
    o = {}
    if titulo is None:
        o["title"] = [{"properties": {"show": B(False)}}]
    else:
        o["title"] = [{"properties": {"show": B(True), "text": S(titulo), "heading": S("Normal"),
                                      "bold": B(True), "fontSize": D(tam), "fontColor": COR("#00335A")}}]
    if subtitulo:
        o["subTitle"] = [{"properties": {"show": B(True), "text": S(subtitulo), "fontSize": D(10),
                                         "fontColor": COR("#5B6B78")}}]
    else:
        o["subTitle"] = [{"properties": {"show": B(False)}}]
    if fundo is False:
        o["background"] = [{"properties": {"show": B(False)}}]
    elif fundo:
        o["background"] = [{"properties": {"show": B(True), "color": COR(fundo), "transparency": D(0)}}]
    if borda is False:
        o["border"] = [{"properties": {"show": B(False)}}]
    elif borda:
        o["border"] = [{"properties": {"show": B(True), "color": COR(borda), "radius": D(10)}}]
    o["dropShadow"] = [{"properties": {"show": B(sombra)}}]
    return o

# ── visuais ─────────────────────────────────────────────────────────────
def textbox(pg, seed, pos, paragrafos, fundo=None, borda=None, sombra=False):
    """paragrafos: lista de listas de (texto, tamanho_pt, cor, negrito)"""
    ps = []
    for par in paragrafos:
        runs = []
        for (t, tam, cor, neg) in par:
            st = {"fontSize": "%dpt" % tam, "color": cor, "fontFamily": "Segoe UI"}
            if neg: st["fontWeight"] = "bold"
            runs.append({"value": t, "textStyle": st})
        ps.append({"textRuns": runs})
    v = {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": ps}}]},
         "drillFilterOtherVisuals": True}
    o = {}
    if fundo:
        o["background"] = [{"properties": {"show": B(True), "color": COR(fundo), "transparency": D(0)}}]
    if borda:
        o["border"] = [{"properties": {"show": B(True), "color": COR(borda), "radius": D(10)}}]
    if sombra:
        o["dropShadow"] = [{"properties": {"show": B(True)}}]
    if o: v["visualContainerObjects"] = o
    return pg.grava(seed, pos, v)

def cartao(pg, seed, pos, medida, tam=12, cor="#1D2B36", cor_medida=None, fonte="Segoe UI Semibold",
           fundo=False, tabela=M, quebra=True):
    """cartão clássico (visualType card): um valor, sem rótulo de categoria"""
    cor_v = CORM(cor_medida) if cor_medida else COR(cor)
    v = {"visualType": "card",
         "query": {"queryState": {"Values": {"projections": [proj(MEAS(medida, tabela))]}}},
         "objects": {
             "labels": [{"properties": {"fontSize": D(tam), "color": cor_v, "fontFamily": S(fonte)}}],
             "categoryLabels": [{"properties": {"show": B(False)}}],
             "wordWrap": [{"properties": {"show": B(quebra)}}]},
         "visualContainerObjects": vc(None, fundo=fundo if fundo else False, borda=False),
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v)

def kpi(pg, seed, pos, medida, nome, tabela="_Medidas"):
    """cartão novo (cardVisual), igual aos KPIs que já existem no relatório"""
    f = MEAS(medida, tabela)
    v = {"visualType": "cardVisual",
         "query": {"queryState": {"Data": {"projections": [proj(f, nome)]}}},
         "visualContainerObjects": {
             "background": [{"properties": {"show": B(True), "color": {"solid": {"color": {"expr": {"ThemeDataColor": {"ColorId": 0, "Percent": 0}}}}}}}],
             "dropShadow": [{"properties": {"show": B(True)}}]},
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v)

def imagem(pg, seed, pos, arquivo, link=None):
    v = {"visualType": "image",
         "objects": {"image": [{"properties": {"sourceFile": {"image": {
             "name": S(arquivo),
             "url": {"expr": {"ResourcePackageItem": {"PackageName": "RegisteredResources", "PackageType": 1, "ItemName": arquivo}}},
             "scaling": S("Fit")}}}}]},
         "visualContainerObjects": {"background": [{"properties": {"show": B(False)}}]},
         "drillFilterOtherVisuals": True}
    if link:
        v["visualContainerObjects"]["visualLink"] = [{"properties": {"show": B(True), "type": S("PageNavigation"), "navigationSection": S(link)}}]
    return pg.grava(seed, pos, v)

def botao(pg, seed, pos, destino):
    v = {"visualType": "actionButton",
         "objects": {"icon": [{"properties": {"shapeType": S("blank")}, "selector": {"id": "default"}}]},
         "visualContainerObjects": {
             "background": [{"properties": {"show": B(False)}}],
             "visualLink": [{"properties": {"show": B(True), "type": S("PageNavigation"), "navigationSection": S(destino)}}]},
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v, how="InsertVisualButton")

def fatiador(pg, seed, pos, campo, filtros=None):
    v = {"visualType": "listSlicer",
         "query": {"queryState": {"Values": {"projections": [proj(campo, ativo=True)]}}},
         "drillFilterOtherVisuals": True}
    pos = tuple(pos[:4]) + (7500 + pg.n,)          # acima do fundo da lateral (z 6000)
    return pg.grava(seed, pos, v, filtros=filtros)

def barra_meta(pg, seed, pos, m_valor, m_resto, cor="#1B6FB0"):
    """barra de progresso: 100% empilhada com valor + restante"""
    fv, fr = MEAS(m_valor), MEAS(m_resto)
    v = {"visualType": "hundredPercentStackedBarChart",
         "query": {"queryState": {"Y": {"projections": [proj(fv), proj(fr)]}}},
         "objects": {
             "categoryAxis": [{"properties": {"show": B(False)}}],
             "valueAxis": [{"properties": {"show": B(False), "gridlineShow": B(False)}}],
             "legend": [{"properties": {"show": B(False)}}],
             "labels": [{"properties": {"show": B(False)}}],
             "dataPoint": [{"properties": {"fill": COR(cor)}, "selector": {"metadata": qref(fv)}},
                           {"properties": {"fill": COR("#E6EBF0")}, "selector": {"metadata": qref(fr)}}]},
         "visualContainerObjects": vc(None, fundo=False, borda=False),
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v)

def colunas(pg, seed, pos, titulo, categoria=None, medidas=(), cores=None, filtros=None, rotulos=True,
            legenda=False, subtitulo=None, eixo_valor=False, tipo="clusteredColumnChart", serie=None,
            cor_medida=None, ordem=None, fundo=None, borda=None, tam_titulo=13):
    q = {}
    if categoria is not None: q["Category"] = {"projections": [proj(categoria, ativo=True)]}
    if serie is not None: q["Series"] = {"projections": [proj(serie)]}
    q["Y"] = {"projections": [proj(f, n) for f, n in medidas]}
    obj = {"legend": [{"properties": {"show": B(legenda), "position": S("Bottom")}}],
           "labels": [{"properties": {"show": B(rotulos), "fontSize": D(10)}}],
           "valueAxis": [{"properties": {"show": B(eixo_valor), "gridlineShow": B(eixo_valor)}}]}
    dp = []
    if cores:
        for (f, _), c in zip(medidas, cores):
            dp.append({"properties": {"fill": COR(c)}, "selector": {"metadata": qref(f)}})
    if cor_medida:
        dp.append({"properties": {"fill": CORM(cor_medida)},
                   "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}})
    if dp: obj["dataPoint"] = dp
    v = {"visualType": tipo, "query": {"queryState": q}, "objects": obj,
         "visualContainerObjects": vc(titulo, subtitulo=subtitulo, fundo=fundo, borda=borda, tam=tam_titulo),
         "drillFilterOtherVisuals": True}
    if ordem:
        v["query"]["sortDefinition"] = {"sort": [{"field": ordem[0], "direction": ordem[1]}], "isDefaultSort": False}
    return pg.grava(seed, pos, v, filtros=filtros)

def tabela(pg, seed, pos, titulo, campos, cores=None, larguras=None, filtros=None, ordem=None,
           subtitulo=None, cabecalho=True, fonte=10):
    """campos: lista de (field, nome) · cores: {índice: medida de cor} · larguras: {índice: px}"""
    q = {"Values": {"projections": [proj(f, n) for f, n in campos]}}
    obj = {"values": [{"properties": {"fontSize": D(fonte)}}],
           "columnHeaders": [{"properties": {"fontSize": D(fonte)}}],
           "total": [{"properties": {"totals": B(False)}}]}
    if not cabecalho:
        obj["columnHeaders"] = [{"properties": {"fontColor": COR("#FFFFFF"), "backColor": COR("#FFFFFF"), "fontSize": D(1)}}]
    for i, m in (cores or {}).items():
        obj["values"].append({"properties": {"fontColor": CORM(m)},
                              "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                           "metadata": qref(campos[i][0])}})
    if larguras:
        obj["columnWidth"] = [{"properties": {"value": D(w)}, "selector": {"metadata": qref(campos[i][0])}}
                              for i, w in larguras.items()]
    v = {"visualType": "tableEx", "query": {"queryState": q}, "objects": obj,
         "visualContainerObjects": vc(titulo, subtitulo=subtitulo), "drillFilterOtherVisuals": True}
    if ordem:
        v["query"]["sortDefinition"] = {"sort": [{"field": ordem[0], "direction": ordem[1]}], "isDefaultSort": False}
    return pg.grava(seed, pos, v, filtros=filtros)

def linhas(pg, seed, pos, titulo, eixo, medidas, cores, subtitulo=None):
    q = {"Category": {"projections": [proj(eixo, ativo=True)]},
         "Y": {"projections": [proj(f, n) for f, n in medidas]}}
    obj = {"legend": [{"properties": {"show": B(True), "position": S("Top")}}],
           "dataPoint": [{"properties": {"fill": COR(c)}, "selector": {"metadata": qref(f)}} for (f, _), c in zip(medidas, cores)],
           "valueAxis": [{"properties": {"show": B(True), "gridlineShow": B(True)}}],
           "lineStyles": [{"properties": {"showMarker": B(True), "strokeWidth": D(2)}}]
                         + [{"properties": {"lineStyle": S("dashed"), "showMarker": B(False)}, "selector": {"metadata": qref(f)}}
                            for (f, _) in medidas[1:]]}
    v = {"visualType": "lineChart", "query": {"queryState": q}, "objects": obj,
         "visualContainerObjects": vc(titulo, subtitulo=subtitulo), "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v)

# ── filtros de visual ───────────────────────────────────────────────────
def f_igual_medida(seed, medida, valor="1L", tabela=M):
    return {"name": vid("f" + seed), "field": MEAS(medida, tabela), "type": "Advanced",
            "filter": {"Version": 2, "From": [{"Name": "m", "Entity": tabela, "Type": 0}],
                       "Where": [{"Condition": {"Comparison": {"ComparisonKind": 0,
                                  "Left": {"Measure": {"Expression": {"SourceRef": {"Source": "m"}}, "Property": medida}},
                                  "Right": {"Literal": {"Value": valor}}}}}]}}

def f_maior_medida(seed, medida, valor="0D", tabela=M):
    return {"name": vid("f" + seed), "field": MEAS(medida, tabela), "type": "Advanced",
            "filter": {"Version": 2, "From": [{"Name": "m", "Entity": tabela, "Type": 0}],
                       "Where": [{"Condition": {"Comparison": {"ComparisonKind": 1,
                                  "Left": {"Measure": {"Expression": {"SourceRef": {"Source": "m"}}, "Property": medida}},
                                  "Right": {"Literal": {"Value": valor}}}}}]}}

def f_em(seed, tabela, coluna, valores):
    return {"name": vid("f" + seed), "field": COL(coluna, tabela), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": tabela, "Type": 0}],
                       "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": coluna}}],
                                                       "Values": [[{"Literal": {"Value": v}}] for v in valores]}}}]}}

def f_entre(seed, tabela, coluna, de, ate):
    c = {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": coluna}}
    return {"name": vid("f" + seed), "field": COL(coluna, tabela), "type": "Advanced",
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": tabela, "Type": 0}],
                       "Where": [{"Condition": {"And": {
                           "Left": {"Comparison": {"ComparisonKind": 2, "Left": c, "Right": {"Literal": {"Value": de}}}},
                           "Right": {"Comparison": {"ComparisonKind": 4, "Left": c, "Right": {"Literal": {"Value": ate}}}}}}}]}}

# ── v16 ─────────────────────────────────────────────────────────────────
def botao_invisivel(pg, seed, pos, destino):
    """botão de navegação SEM preenchimento (na v15 o preenchimento padrão cobria a pílula)"""
    v = {"visualType": "actionButton",
         "objects": {"icon": [{"properties": {"shapeType": S("blank")}, "selector": {"id": "default"}}],
                     "fill": [{"properties": {"show": B(False)}, "selector": {"id": "default"}}],
                     "outline": [{"properties": {"show": B(False)}, "selector": {"id": "default"}}]},
         "visualContainerObjects": {
             "background": [{"properties": {"show": B(False)}}],
             "visualLink": [{"properties": {"show": B(True), "type": S("PageNavigation"), "navigationSection": S(destino)}}]},
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v, how="InsertVisualButton")

def medidor(pg, seed, pos, medida, cor_medida):
    """barra de progresso: barra única de 0 a 100%, cor pela situação"""
    f = MEAS(medida)
    v = {"visualType": "clusteredBarChart",
         "query": {"queryState": {"Y": {"projections": [proj(f)]}}},
         "objects": {
             "categoryAxis": [{"properties": {"show": B(False)}}],
             "valueAxis": [{"properties": {"show": B(False), "gridlineShow": B(False), "start": D(0), "end": D(1)}}],
             "legend": [{"properties": {"show": B(False)}}],
             "labels": [{"properties": {"show": B(False)}}],
             "dataPoint": [{"properties": {"fill": CORM(cor_medida)},
                            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]},
         "visualContainerObjects": vc(None, fundo="#EEF2F6", borda=False),
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v)

def sem_total(visual_json_path):
    import json
    o = json.load(open(visual_json_path, encoding="utf-8"))
    o["visual"].setdefault("objects", {})["total"] = [{"properties": {"totals": B(False)}}]
    json.dump(o, open(visual_json_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def matriz(pg, seed, pos, titulo, linhas_, colunas_, valor, fundo_m, fonte_m, subtitulo=None, filtros=None):
    fv = valor
    v = {"visualType": "pivotTable",
         "query": {"queryState": {"Rows": {"projections": [proj(linhas_, ativo=True)]},
                                  "Columns": {"projections": [proj(colunas_, ativo=True)]},
                                  "Values": {"projections": [proj(fv, " ")]}}},
         "objects": {
             "subTotals": [{"properties": {"rowSubtotals": B(False), "columnSubtotals": B(False)}}],
             "values": [{"properties": {"fontSize": D(12)}},
                        {"properties": {"backColor": CORM(fundo_m), "fontColor": CORM(fonte_m)},
                         "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": qref(fv)}}],
             "columnHeaders": [{"properties": {"fontSize": D(10)}}],
             "rowHeaders": [{"properties": {"fontSize": D(11)}}],
             "grid": [{"properties": {"gridVertical": B(True), "gridVerticalColor": COR("#FFFFFF"), "gridVerticalWeight": D(2),
                                      "gridHorizontal": B(True), "gridHorizontalColor": COR("#FFFFFF"), "gridHorizontalWeight": D(2),
                                      "rowPadding": D(6)}}]},
         "visualContainerObjects": vc(titulo, subtitulo=subtitulo),
         "drillFilterOtherVisuals": True}
    return pg.grava(seed, pos, v, filtros=filtros)
