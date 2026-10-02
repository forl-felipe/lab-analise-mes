# -*- coding: utf-8 -*-
"""Células de PREENCHIMENTO de cada aba (layout da planilha de setembro, com Alpine até a linha 24).
É o que se apaga para fazer a planilha padrão (rótulos, limites e fórmulas ficam)."""
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

def faixa(cols, r0, r1):
    return ["%s%d" % (c, r) for r in range(r0, r1 + 1) for c in cols]
def cols(a, b):
    return [L(i) for i in range(CI(a), CI(b) + 1)]

ENTRADAS = {}
ENTRADAS["Tambor de Abrasão"] = (faixa("A C D H I M N Q R".split(), 6, 14) + faixa("A C D H I L M O".split(), 25, 34)
                                 + ["C20", "C21", "H20", "H21", "C41", "C42", "H41", "H42"])
ENTRADAS["Gran. Fina Alpine"] = faixa("A B D F H".split(), 7, 24)
ENTRADAS["Calibração Blaine"] = faixa("A C H L N".split(), 6, 71) + [("J%d" % r, "num") for r in range(6, 72)]
ENTRADAS["Comparativo Fisher"] = faixa("A B C D E".split(), 8, 12)
pen = []
for k in range(13):
    r = 3 + 10 * k
    pen += faixa(cols("B", "AB"), r + 3, r + 8) + faixa(["AC", "AD", "AE"], r, r + 9)
ENTRADAS["Verificação Peneiradores"] = pen
ENTRADAS["Compressão"] = faixa(cols("B", "P") + ["W", "X"], 6, 9) + faixa(cols("B", "P") + ["W", "X"], 20, 23)
gra = []
for r in (7, 31, 55):
    gra += ["C%d" % r, "G%d" % r, "J%d" % r, "N%d" % r] + faixa(["C", "F", "J", "M"], r + 5, r + 14)
ENTRADAS["Granulometria"] = gra
ENTRADAS["Umidade"] = faixa("A B C D E K L M N".split(), 5, 18)
t5 = []
for r in (7, 16, 25):
    t5 += ["D%d" % r, "G%d" % r, "L%d" % r, "O%d" % r] + faixa(["C", "G", "K", "O"], r + 4, r + 6)
ENTRADAS["Tamb Kg x Tamb 15,0Kg"] = t5 + faixa(["R"], 10, 14)

# Lançamentos de outubro já feitos (copiados do arquivo de outubro como estão)
OUTUBRO = {
    "Calibração Blaine": faixa(cols("A", "N"), 7, 7),
    "Verificação Peneiradores": faixa(cols("B", "AE"), 3, 12),
}
