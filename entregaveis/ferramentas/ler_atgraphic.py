"""Lê um arquivo .atgraphic do Aspen Process Graphics Editor e lista objetos, tags e posições.

Uso:
    python ler_atgraphic.py tela.atgraphic [pasta_saida]

Saídas na pasta de saída (padrão: pasta atual):
    <nome>_objetos.csv   um objeto por linha: tipo, posição, textos e tags (separador ponto e vírgula)
    <nome>_tags.csv      tags distintas, com campo/mapa e servidor
    <nome>_layout.svg    desenho aproximado da tela, para conferência

Requer: pip install olefile

Leitura feita por engenharia reversa do formato (arquivo composto OLE com um fluxo
"Contents" serializado). Não altera o arquivo original. Os nomes de tipo são aproximados:
classes declaradas no arquivo aparecem pelo nome (AtTextShape, AtTrend, AtDataField...);
objetos que reutilizam uma classe aparecem como "refN".
"""
import csv
import html
import os
import re
import struct
import sys

import olefile

IGNORAR = {"", "VAL", "GERMANO"}  # atributos e fonte de dados padrão de objetos sem tag


def ler_conteudo(caminho):
    with olefile.OleFileIO(caminho) as ole:
        return ole.openstream("Contents").read()


def strings(d):
    out, j = [], 0
    while j < len(d):
        if d[j:j + 3] == b"\xff\xfe\xff":
            n, k = d[j + 3], j + 4
            if n == 0xFF:
                n, k = int.from_bytes(d[j + 4:j + 6], "little"), j + 6
            out.append((j, d[k:k + 2 * n].decode("utf-16-le", "replace")))
            j = k + 2 * n
        else:
            j += 1
    return out


def retangulo(d, o):
    return struct.unpack("<4i", d[o:o + 16])


def objetos(d):
    cabecas = []
    for m in re.finditer(rb"\xff\xff\x05\x00(.)\x00", d, re.S):
        n = m.group(1)[0]
        nome = d[m.end():m.end() + n]
        if re.fullmatch(rb"At[A-Za-z]+", nome):
            cabecas.append((m.end() + n, nome.decode()))
    for o in range(len(d) - 18):
        if d[o + 1] == 0x80 and 1 <= d[o] <= 40:
            x1, y1, x2, y2 = retangulo(d, o + 2)
            if 0 <= x1 < x2 < 20000 and 0 <= y1 < y2 < 20000 and x2 - x1 < 15000:
                cabecas.append((o + 2, "ref%d" % d[o]))
    cabecas.sort()
    filtradas, ultimo = [], -100
    for e, nome in cabecas:
        if e - ultimo >= 18:
            filtradas.append((e, nome))
            ultimo = e
    ss = strings(d)
    res = []
    for i, (e, nome) in enumerate(filtradas):
        fim = filtradas[i + 1][0] if i + 1 < len(filtradas) else len(d)
        textos = [t for o, t in ss if e < o < fim and t not in IGNORAR]
        res.append({"tipo": nome, "ret": retangulo(d, e), "textos": textos})
    return res


def tags(objs):
    """Tags vinculadas: trios (tag, mapa, servidor), em que o mapa começa por IP_."""
    vistos = {}
    for ob in objs:
        t = ob["textos"]
        for i in range(len(t) - 2):
            if t[i + 1].startswith("IP_") and re.fullmatch(r"[A-Za-z0-9_.\-]+", t[i]):
                vistos.setdefault((t[i], t[i + 1], t[i + 2]), 0)
                vistos[(t[i], t[i + 1], t[i + 2])] += 1
    return vistos


def svg(objs, caminho):
    xs = [ob["ret"][2] for ob in objs] + [5000]
    ys = [ob["ret"][3] for ob in objs] + [3500]
    W, H = max(xs) + 50, max(ys) + 50
    partes = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W // 3}" height="{H // 3}">',
              f'<rect width="{W}" height="{H}" fill="#c0c0c0"/>']
    for ob in objs:
        x1, y1, x2, y2 = ob["ret"]
        tags_ob = [t for i, t in enumerate(ob["textos"]) if i + 1 < len(ob["textos"]) and ob["textos"][i + 1].startswith("IP_")]
        rotulo = " ".join(t for t in ob["textos"] if not t.startswith("IP_") and t != "UBU" and t not in tags_ob and "\\" not in t)
        cor = "#1f6fb5" if tags_ob else "#555"
        partes.append(f'<rect x="{x1}" y="{y1}" width="{x2 - x1}" height="{y2 - y1}" fill="none" stroke="{cor}" stroke-width="6"/>')
        if tags_ob:
            partes.append(f'<text x="{x1 + 10}" y="{y2 - 12}" font-size="44" font-family="Arial" fill="#0b3d6e">{html.escape(", ".join(dict.fromkeys(tags_ob)))}</text>')
        if rotulo.strip():
            partes.append(f'<text x="{x1 + 10}" y="{y1 + 60}" font-size="64" font-family="Arial" fill="#000">{html.escape(rotulo.strip()[:60])}</text>')
    partes.append("</svg>")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write("\n".join(partes))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    arq = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else "."
    os.makedirs(saida, exist_ok=True)
    base = os.path.splitext(os.path.basename(arq))[0]
    objs = objetos(ler_conteudo(arq))

    with open(os.path.join(saida, base + "_objetos.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["n", "tipo", "x1", "y1", "x2", "y2", "textos_e_tags"])
        for i, ob in enumerate(objs, 1):
            w.writerow([i, ob["tipo"], *ob["ret"], " | ".join(ob["textos"])])

    with open(os.path.join(saida, base + "_tags.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["tag", "mapa", "servidor", "ocorrencias"])
        for (t, m, s), n in sorted(tags(objs).items()):
            w.writerow([t, m, s, n])

    svg(objs, os.path.join(saida, base + "_layout.svg"))
    print(f"{len(objs)} objetos e {len(tags(objs))} vínculos de tag em {saida}")


if __name__ == "__main__":
    main()
