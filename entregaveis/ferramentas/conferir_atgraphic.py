"""Confere um .atgraphic gerado por gerar_atgraphic.py: relê o arquivo objeto a objeto, no mesmo
formato que o Aspen usa, e redesenha a tela em SVG para comparação visual com o protótipo.

Uso: python conferir_atgraphic.py arquivo.atgraphic saida.svg
"""
import html
import struct
import sys

import olefile


def ler(caminho):
    with olefile.OleFileIO(caminho) as ole:
        d = ole.openstream("Contents").read()
    p = 0

    def rd(fmt):
        nonlocal p
        v = struct.unpack_from(fmt, d, p)
        p += struct.calcsize(fmt)
        return v

    def rstr():
        nonlocal p
        assert d[p:p + 3] == b"\xff\xfe\xff", f"texto esperado na posição {p}"
        n = d[p + 3]
        p += 4
        if n == 0xFF:
            n = struct.unpack_from("<H", d, p)[0]
            p += 2
        s = d[p:p + 2 * n].decode("utf-16-le")
        p += 2 * n
        return s

    def cor():
        r, g, b, _ = rd("<4B")
        return f"rgb({r},{g},{b})"

    assert d[:4] == b"AzPn"
    p = 4
    (versao,) = rd("<H")
    larg, alt = rd("<2i")
    fundo = cor()
    cor()
    rd("<2i")
    rstr()
    p += 92
    (n,) = rd("<H")
    classes, indice, objetos = {}, 1, []
    for _ in range(n):
        (tag,) = rd("<H")
        if tag == 0xFFFF:
            _, tam = rd("<HH")
            nome = d[p:p + tam].decode("ascii")
            p += tam
            classes[indice] = nome
            indice += 1
        else:
            nome = classes[tag & 0x7FFF]
        indice += 1
        ret = rd("<4i")
        p += 12
        c1 = cor(); f1, = rd("<I"); c2 = cor(); rd("<I"); c3 = cor(); rd("<I")
        tagname, atributo, mapa, servidor = rstr(), rstr(), rstr(), rstr()
        (animado,) = struct.unpack_from("<I", d, p)
        p += 10
        ob = dict(classe=nome, ret=ret, c1=c1, f1=f1, c2=c2, c3=c3, tag=tagname, mapa=mapa)
        if animado:
            ob["faixa"] = rd("<2f")
            rstr(); rstr()
            ob["acesa"] = cor()
            p += 4
        if nome == "AtTextShape":
            ob["texto"] = rstr()
            ob["fonte"], = struct.unpack_from("<i", d, p)
            ob["peso"], = struct.unpack_from("<i", d, p + 16)
            p += 92 + 4
        elif nome == "AtBasicShape":
            (k,) = rd("<H")
            ob["pontos"] = [rd("<2i") for _ in range(k)]
        elif nome == "AtTrend":
            (k,) = rd("<H")
            ob["penas"] = []
            for _ in range(k):
                t, _a, m, _s = rstr(), rstr(), rstr(), rstr()
                mn, mx = rd("<2f")
                c = cor()
                p += 2
                ob["penas"].append((t, m, mn, mx, c))
        elif nome == "AtDataField":
            p += 4
            ob["largura"], ob["decimais"] = rd("<2I")
            ob["fonte"], = struct.unpack_from("<i", d, p)
            ob["peso"], = struct.unpack_from("<i", d, p + 16)
            p += 92
            rstr(); rstr()
            p += 12
        else:
            raise ValueError(f"classe não prevista: {nome}")
        objetos.append(ob)
    assert p == len(d), f"sobraram {len(d) - p} bytes no fim do fluxo"
    return dict(versao=versao, largura=larg, altura=alt, fundo=fundo, objetos=objetos)


def svg(tela, caminho):
    W, H = tela["largura"], tela["altura"]
    s = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="1400" height="980">',
         f'<rect width="{W}" height="{H}" fill="{tela["fundo"]}"/>']
    for o in tela["objetos"]:
        x1, y1, x2, y2 = o["ret"]
        if o["classe"] == "AtBasicShape":
            pts = " ".join(f"{a},{b}" for a, b in o["pontos"])
            fill = "none" if o["f1"] else o.get("acesa", o["c1"])
            s.append(f'<polygon points="{pts}" fill="{fill}" stroke="{o["c2"]}" stroke-width="3"/>')
        elif o["classe"] == "AtTrend":
            s.append(f'<rect x="{x1}" y="{y1}" width="{x2 - x1}" height="{y2 - y1}" fill="{o["c1"]}" stroke="#ff4fd8" stroke-width="4" stroke-dasharray="12 8"/>')
            for i, (t, m, mn, mx, c) in enumerate(o["penas"]):
                s.append(f'<text x="{x1 + 20}" y="{y1 + 60 + 50 * i}" font-family="Arial" font-size="40" fill="{c}">Tendência: {html.escape(t)} ({mn:g} a {mx:g})</text>')
        else:
            tam = o["fonte"] / 1.117
            peso = "bold" if o["peso"] >= 600 else "normal"
            txt = o["texto"] if o["classe"] == "AtTextShape" else f'[{o["tag"]}]'
            cor = o["c1"] if o["classe"] == "AtTextShape" else "#ff4fd8"
            if o["classe"] == "AtDataField":
                s.append(f'<rect x="{x1}" y="{y1}" width="{x2 - x1}" height="{y2 - y1}" fill="none" stroke="#ff4fd8" stroke-width="4" stroke-dasharray="12 8"/>')
                tam = min(tam, 40)
            s.append(f'<text x="{(x1 + x2) / 2}" y="{(y1 + y2) / 2}" font-family="Arial" font-size="{tam:.1f}" font-weight="{peso}" '
                     f'fill="{cor}" text-anchor="middle" dominant-baseline="central">{html.escape(txt)}</text>')
    s.append("</svg>")
    with open(caminho, "w", encoding="utf-8") as fh:
        fh.write("\n".join(s))


if __name__ == "__main__":
    tela = ler(sys.argv[1])
    tipos = {}
    for o in tela["objetos"]:
        tipos[o["classe"]] = tipos.get(o["classe"], 0) + 1
    print(f"Arquivo íntegro: tela {tela['largura']} x {tela['altura']}, {len(tela['objetos'])} objetos {tipos}")
    if len(sys.argv) > 2:
        svg(tela, sys.argv[2])
