"""Converte o logotipo vetorial (SVG do kit visual Samarco) em polígonos simples para o Aspen.

O Aspen desenha polígonos preenchidos, sem furos. Por isso cada contorno do SVG vira um polígono, e os
contornos internos (o miolo das letras A, R e O, por exemplo) são desenhados depois, na cor do fundo
onde o logotipo fica. O resultado é igual à regra de preenchimento "evenodd" do SVG original.
"""
import re


def _numeros(txt):
    return [float(v) for v in re.findall(r"-?\d+(?:\.\d+)?", txt)]


def contornos(d, passos=6):
    """Lista de contornos (listas de pontos) de um atributo d com comandos absolutos M, L, C e Z."""
    saida, atual, pos = [], [], (0.0, 0.0)
    for cmd, args in re.findall(r"([MLCZ])([^MLCZ]*)", d):
        v = _numeros(args)
        if cmd == "M":
            if atual:
                saida.append(atual)
            pos = (v[0], v[1])
            atual = [pos]
            for i in range(2, len(v), 2):
                pos = (v[i], v[i + 1])
                atual.append(pos)
        elif cmd == "L":
            for i in range(0, len(v), 2):
                pos = (v[i], v[i + 1])
                atual.append(pos)
        elif cmd == "C":
            for i in range(0, len(v), 6):
                x0, y0 = pos
                x1, y1, x2, y2, x3, y3 = v[i:i + 6]
                for k in range(1, passos + 1):
                    t = k / passos
                    a, b, c, e = (1 - t) ** 3, 3 * (1 - t) ** 2 * t, 3 * (1 - t) * t ** 2, t ** 3
                    atual.append((a * x0 + b * x1 + c * x2 + e * x3, a * y0 + b * y1 + c * y2 + e * y3))
                pos = (x3, y3)
        elif cmd == "Z":
            if atual:
                saida.append(atual)
            atual = []
    if atual:
        saida.append(atual)
    return [c for c in saida if len(c) >= 3]


def _dentro(p, poli):
    x, y = p
    dentro = False
    for i in range(len(poli)):
        x1, y1 = poli[i]
        x2, y2 = poli[i - 1]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            dentro = not dentro
    return dentro


def poligonos(svg_txt, fundo):
    """Retorna (largura, altura, [(pontos, cor_rgb)]) na ordem de desenho."""
    vb = _numeros(re.search(r'viewBox="([^"]+)"', svg_txt).group(1))
    lista = []
    for cor_hex, d in re.findall(r'<path fill="(#[0-9A-Fa-f]{6})"[^>]*?d="([^"]+)"', svg_txt):
        cor = tuple(int(cor_hex[i:i + 2], 16) for i in (1, 3, 5))
        cs = contornos(d)
        for c in cs:
            prof = sum(1 for o in cs if o is not c and _dentro(c[0], o))
            lista.append((prof, c, cor if prof % 2 == 0 else fundo))
    lista.sort(key=lambda t: t[0])   # por camada: contornos externos primeiro, furos depois
    return vb[2], vb[3], [(c, cor) for _, c, cor in lista]
