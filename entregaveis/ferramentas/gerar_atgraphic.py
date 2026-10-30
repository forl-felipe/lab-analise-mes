"""Gera o arquivo .atgraphic da tela de compressão, pronto para abrir no Aspen Process Graphics Editor V14.2.

Uso (na pasta ferramentas):
    node gerar_imagens.js            (gera ../03_aspen/objetos_tela.json a partir do protótipo)
    python gerar_atgraphic.py

Saída: ../03_aspen/LCE_Compressao_Online_Usina4.atgraphic

Tudo é objeto nativo do Aspen (vetorial), como na tela PIMS PQ US04: painéis e linhas viram polígonos,
textos viram objetos de texto, os valores viram campos de dados e as tendências já vêm ligadas às
tags. Assim a tela não perde qualidade ao maximizar. Só o logotipo é inserido no editor (ver roteiro).

O formato foi obtido por engenharia reversa da tela PIMS_PQ_US04_atual_rev_1.atgraphic:
arquivo composto OLE com um fluxo "Contents" serializado no padrão MFC (CArchive). Tela de
50 x 35 unidades = 5000 x 3500 unidades lógicas; o protótipo tem 1600 x 1120 px, logo 1 px = 3,125.
"""
import json
import math
import os
import re
import struct
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from montar_pacote_aspen import SPEC  # noqa: E402  (tags, mapas e formatos dos campos)
from logo_poligonos import poligonos  # noqa: E402

SAIDA = os.path.join(AQUI, "..", "03_aspen")
ARQUIVO = "LCE_Compressao_Online_Usina4.atgraphic"
ESCALA = 5000 / 1600
CLSID_ASPEN = "A272B18A-DA53-11CF-A640-00A0240279F4"  # mesmo identificador da tela PQ US04
SERVIDOR = "UBU"
MAPA_DISCRETO = "IP_DISCRETEMAP"


# ---------------------------------------------------------------- serialização no padrão MFC
def u(v):
    return int(round(v * ESCALA))


def wstr(s):
    b = s.encode("utf-16-le")
    n = len(s)
    if n < 0xFF:
        return b"\xff\xfe\xff" + bytes([n]) + b
    return b"\xff\xfe\xff\xff" + struct.pack("<H", n) + b


def cor(rgb):
    r, g, b = rgb
    return bytes([r, g, b, 2])


def rgb(css):
    return tuple(int(x) for x in re.findall(r"\d+", css)[:3])


def logfont(altura, peso):
    nome = "Arial".encode("utf-16-le").ljust(64, b"\0")
    return struct.pack("<5i", altura, 0, 0, 0, peso) + bytes([0, 0, 0, 1, 0, 0, 2, 0x20]) + nome


def comum(ret, c1, f1, c2, f2, c3, f3, tag="", atributo="VAL", mapa="", servidor="GERMANO"):
    x1, y1, x2, y2 = ret
    return (struct.pack("<4i", x1, y1, x2, y2) + bytes(12)
            + cor(c1) + struct.pack("<I", f1) + cor(c2) + struct.pack("<I", f2) + cor(c3) + struct.pack("<I", f3)
            + wstr(tag) + wstr(atributo) + wstr(mapa) + wstr(servidor))


def poligono(pontos, preenchimento, linha, fundo):
    xs = [p[0] for p in pontos]
    ys = [p[1] for p in pontos]
    ret = (min(xs), min(ys), max(xs), max(ys))
    if preenchimento:
        corpo = comum(ret, preenchimento, 0, linha or preenchimento, 0, preenchimento, 1)
    else:
        corpo = comum(ret, linha, 1, linha, 0, fundo, 1)
    return "AtBasicShape", corpo + bytes(10) + struct.pack("<H", len(pontos)) + b"".join(struct.pack("<2i", *p) for p in pontos)


def retangulo_arredondado(x, y, w, h, r):
    if r <= 0:
        p = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
    else:
        p = []
        cantos = [(x + w - r, y + r, -90), (x + w - r, y + h - r, 0), (x + r, y + h - r, 90), (x + r, y + r, 180)]
        for cx, cy, a0 in cantos:
            for k in range(5):
                a = math.radians(a0 + k * 22.5)
                p.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    p = [(u(a), u(b)) for a, b in p]
    return p + [p[0]]


def texto(t):
    pad = 3
    ret = (u(t["x"] - pad), u(t["y"]), u(t["x"] + t["w"] + pad), u(t["y"] + t["h"]))
    altura = int(round(t["tamanho"] * 1.117 * ESCALA))
    peso = 700 if t["peso"] >= 600 else 400
    c = rgb(t["cor"])
    corpo = comum(ret, c, 0, c, 0, rgb(t["fundo"]), 1)
    return "AtTextShape", corpo + bytes(10) + wstr(t["texto"]) + logfont(altura, peso) + bytes(4)


def campo(c, spec):
    ret = (u(c["x"]), u(c["y"]), u(c["x"] + c["w"]), u(c["y"] + c["h"]))
    mapa = spec["mapa"] if spec["mapa"].startswith("IP_") else MAPA_DISCRETO
    formato = spec["formato"]
    decimais = 1 if formato.startswith("1 casa") else 0
    largura = 19 if "aaaa" in formato else (7 if "casa" in formato else 12)
    altura = int(round(c["tamanho"] * 1.117 * ESCALA))
    peso = 700 if c["peso"] >= 600 else 400
    texto_cor = tuple(int(v) for v in spec["cor"].split(","))
    corpo = comum(ret, texto_cor, 0, texto_cor, 0, rgb(c["fundo"]), 1,
                  tag=spec["tag"], atributo="VAL", mapa=mapa, servidor=SERVIDOR)
    corpo += bytes(10) + bytes(4) + struct.pack("<2I", largura, decimais) + logfont(altura, peso) + wstr("") + wstr("") + bytes(12)
    return "AtDataField", corpo


# Penas das tendências: (tag, mapa, mínimo, máximo, cor RGB). Escala fixa, igual para todas as penas.
PENAS = {
    3: [("307-GI-001", "IP_ANALOGMAP", 0, 500, (89, 198, 242))],
    7: [("307-GI-002", "IP_ANALOGMAP", 0, 150, (89, 198, 242))],
    11: [("M4710050020-0031-HHLFU", "IP_MESVALOR", 0, 500, (89, 198, 242))],
    13: [("M4710050020-0031-HHLFU", "IP_MESVALOR", 50, 450, (255, 255, 255)),     # referência do laboratório
         ("307-GI-001", "IP_ANALOGMAP", 50, 450, (89, 198, 242))],                # compressão online, cada ensaio
    # Cortadores: escala 8 a 40 para Automático (32) e Manual (16) ficarem nas linhas de referência da tela
    17: [("U04-07CR001", "IP_PIDMAP03", 8, 40, (89, 198, 242))],
    18: [("U04-07CR002", "IP_PIDMAP03", 8, 40, (89, 198, 242))],
}
FUNDO_TENDENCIA = {13: (1, 41, 74), 17: (1, 41, 74), 18: (1, 41, 74)}   # fundo da área dos gráficos


def tendencia(c):
    ret = (u(c["x"]), u(c["y"]), u(c["x"] + c["w"]), u(c["y"] + c["h"]))
    fundo = FUNDO_TENDENCIA.get(c["n"], rgb(c["fundo"]))
    corpo = comum(ret, fundo, 0, (22, 77, 120), 0, fundo, 1) + bytes(10)
    penas = PENAS[c["n"]]
    corpo += struct.pack("<H", len(penas))
    for tag, mapa, mn, mx, cor_pena in penas:
        corpo += wstr(tag) + wstr("VAL") + wstr(mapa) + wstr(SERVIDOR) + struct.pack("<2f", mn, mx) + cor(cor_pena) + b"\xff\xff"
    return "AtTrend", corpo


def lampada(l):
    """Lâmpada animada: cor apagada fora da faixa e cor acesa quando o valor está entre lo e hi
    (mesma animação dos indicadores Manual/Automático da tela PQ US04)."""
    cx, cy, r = l["x"] + l["w"] / 2, l["y"] + l["h"] / 2, l["w"] / 2
    pts = [(u(cx + r * math.cos(math.radians(a))), u(cy + r * math.sin(math.radians(a)))) for a in range(0, 360, 20)]
    pts.append(pts[0])
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    base, acesa = rgb(l["base"]), rgb(l["ativa"])
    mapa = "IP_PIDMAP03" if l["tag"].startswith("U04-") else MAPA_DISCRETO
    corpo = comum((min(xs), min(ys), max(xs), max(ys)), base, 0, base, 0, base, 1,
                  tag=l["tag"], atributo="VAL", mapa=mapa, servidor=SERVIDOR)
    corpo += bytes.fromhex("01000000010001000000") + struct.pack("<2f", l["lo"], l["hi"]) + wstr("") + wstr("") + cor(acesa) + bytes(4)
    corpo += struct.pack("<H", len(pts)) + b"".join(struct.pack("<2i", *p) for p in pts)
    return "AtBasicShape", corpo


def logotipo(lg):
    """Logotipo oficial (kit visual Samarco) convertido em polígonos, na caixa do protótipo."""
    with open(os.path.join(AQUI, "..", "assets", "samarco-logo-branca.svg"), encoding="utf-8") as fh:
        svg = fh.read()
    w, h, polis = poligonos(svg, rgb(lg["fundo"]))
    esc = min(lg["w"] / w, lg["h"] / h)
    objs = []
    for pts, c in polis:
        p = []
        for x, y in pts:   # pontos repetidos após o arredondamento deixam o contorno serrilhado
            q = (u(lg["x"] + x * esc), u(lg["y"] + y * esc))
            if not p or q != p[-1]:
                p.append(q)
        objs.append(poligono(p + [p[0]], c, c, c))
    return objs


def simbolo(s):
    x, y, w, h = s["x"], s["y"], s["w"], s["h"]
    if s["forma"] == "tri":
        p = [(x + w / 2, y), (x + w, y + h), (x, y + h)]
    elif s["forma"] == "dia":
        p = [(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)]
    elif s["forma"] == "circ":
        p = [(x + w / 2 + w / 2 * math.cos(math.radians(a)), y + h / 2 + h / 2 * math.sin(math.radians(a))) for a in range(0, 360, 30)]
    else:
        p = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
    p = [(u(a), u(b)) for a, b in p]
    c = rgb(s["cor"])
    return poligono(p + [p[0]], c, c, rgb(s["fundo"]))


def conteudo(dados):
    def forma(f):
        h = max(f["h"], 1.3)  # linhas de 1 px ficam visíveis em qualquer zoom
        pts = retangulo_arredondado(f["x"], f["y"], f["w"], h, f["raio"])
        return poligono(pts, rgb(f["preenchimento"]) if f["preenchimento"] else None,
                        rgb(f["linha"]) if f["linha"] else None, rgb(f.get("fundo") or dados["fundoTela"]))

    # Ordem de desenho: painéis, tendências, grade e referências por cima das tendências, símbolos, textos e campos.
    objetos = [forma(f) for f in dados["formas"] if not f.get("sobre")]
    if dados.get("logo"):
        objetos += logotipo(dados["logo"])
    objetos += [tendencia(c) for c in dados["campos"] if c["n"] in PENAS]
    objetos += [forma(f) for f in dados["formas"] if f.get("sobre")]
    objetos += [simbolo(s) for s in dados["simbolos"]]
    objetos += [lampada(l) for l in dados.get("lampadas", [])]
    objetos += [texto(t) for t in dados["textos"]]
    for c in dados["campos"]:
        spec = SPEC[c["n"]]
        if spec["ferramenta"] == "Campo de dados":
            objetos.append(campo(c, spec))

    cab = (b"AzPn" + struct.pack("<H", 5) + struct.pack("<2i", 5000, 3500)
           + cor(rgb(dados["fundoTela"])) + cor((192, 192, 192)) + struct.pack("<2i", 100, 100) + wstr("")
           + struct.pack("<5i", 120, 0, 0, 0, 0) + bytes([0, 0, 0, 1, 0, 0, 2, 0x20]) + "Arial".encode("utf-16-le").ljust(64, b"\0")
           + struct.pack("<H", len(objetos)))
    corpo = bytearray(cab)
    indice, classes = 1, {}
    for nome, dados_obj in objetos:
        if nome in classes:
            corpo += struct.pack("<H", 0x8000 | classes[nome])
        else:
            classes[nome] = indice
            indice += 1
            corpo += b"\xff\xff" + struct.pack("<HH", 5, len(nome)) + nome.encode("ascii")
        indice += 1
        corpo += dados_obj
    return bytes(corpo), objetos


# ---------------------------------------------------------------- arquivo composto OLE (CFB versão 3)
def guid_bytes(g):
    a, b, c, d, e = g.split("-")
    return struct.pack("<IHH", int(a, 16), int(b, 16), int(c, 16)) + bytes.fromhex(d + e)


def entrada_dir(nome, tipo, filho, inicio, tamanho, clsid=b"\0" * 16):
    n = (nome + "\0").encode("utf-16-le") if nome else b""
    return (n.ljust(64, b"\0") + struct.pack("<HBB", len(n), tipo, 1 if tipo else 0)
            + struct.pack("<3I", 0xFFFFFFFF, 0xFFFFFFFF, filho) + clsid + bytes(4) + bytes(16)
            + struct.pack("<IIi", inicio, tamanho, 0))


def cfb(fluxo):
    SETOR, MINI = 512, 64
    FREE, END, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD

    def cadeia(inicio, n):
        return [inicio + k + 1 for k in range(n - 1)] + [END]

    if len(fluxo) >= 4096:                    # fluxo em setores normais
        n_dados = -(-len(fluxo) // SETOR)
        n_fat = 1
        while n_fat * 128 < n_fat + 1 + n_dados:
            n_fat += 1
        assert n_fat <= 109, "fluxo grande demais"
        fat = [FATSECT] * n_fat + [END] + cadeia(n_fat + 1, n_dados)
        dados = fluxo.ljust(n_dados * SETOR, b"\0")
        dir_ini, minifat_ini, n_minifat, raiz_ini, raiz_tam, cont_ini = n_fat, END, 0, END, 0, n_fat + 1
    else:                                     # fluxo pequeno vai no mini fluxo, como manda o formato
        n_mini = -(-len(fluxo) // MINI)
        mini = fluxo.ljust(n_mini * MINI, b"\0")
        n_cont = -(-len(mini) // SETOR)
        n_fat = 1
        fat = [FATSECT, END, END] + cadeia(3, n_cont)
        minifat = cadeia(0, n_mini) + [FREE] * (128 - n_mini)
        dados = struct.pack("<128I", *minifat) + mini.ljust(n_cont * SETOR, b"\0")
        dir_ini, minifat_ini, n_minifat, raiz_ini, raiz_tam, cont_ini = 1, 2, 1, 3, len(mini), 0
    fat += [FREE] * (n_fat * 128 - len(fat))
    difat = list(range(n_fat)) + [FREE] * (109 - n_fat)
    cab = (bytes.fromhex("D0CF11E0A1B11AE1") + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
           + struct.pack("<IIIIIIIII", 0, n_fat, dir_ini, 0, 4096, minifat_ini, n_minifat, END, 0)
           + struct.pack("<109I", *difat))
    assert len(cab) == 512
    diretorio = (entrada_dir("Root Entry", 5, 1, raiz_ini, raiz_tam, guid_bytes(CLSID_ASPEN))
                 + entrada_dir("Contents", 2, FREE, cont_ini, len(fluxo)))
    vazio = b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<3I", FREE, FREE, FREE) + bytes(16 + 4 + 16 + 12)
    diretorio += vazio * 2
    return cab + struct.pack(f"<{len(fat)}I", *fat) + diretorio + dados


def main():
    with open(os.path.join(SAIDA, "objetos_tela.json"), encoding="utf-8") as fh:
        dados = json.load(fh)
    fluxo, objetos = conteudo(dados)
    caminho = os.path.join(SAIDA, ARQUIVO)
    with open(caminho, "wb") as fh:
        fh.write(cfb(fluxo))
    tipos = {}
    for nome, _ in objetos:
        tipos[nome] = tipos.get(nome, 0) + 1
    print(f"{ARQUIVO}: {len(objetos)} objetos {tipos}, fluxo de {len(fluxo)} bytes")

    # Arquivo mínimo para diagnóstico: um painel, um texto, um campo e uma tendência.
    minimo = dict(dados, formas=[f for f in dados["formas"] if f["preenchimento"] and f["linha"]][:1],
                  textos=dados["textos"][:1], simbolos=[], lampadas=dados.get("lampadas", [])[:1], logo=None,
                  campos=[c for c in dados["campos"] if c["n"] in (2, 13)])
    fluxo_min, objs_min = conteudo(minimo)
    with open(os.path.join(SAIDA, "teste_minimo.atgraphic"), "wb") as fh:
        fh.write(cfb(fluxo_min))
    print(f"teste_minimo.atgraphic: {len(objs_min)} objetos")


if __name__ == "__main__":
    main()
