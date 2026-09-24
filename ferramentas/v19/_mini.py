def cartao_mini_dummy(): pass
def mini(pg, k, x, w, titulo, valor, meta=None, barra=None, cor=None, fatos=None):
    y0 = 104
    textbox(pg, "mini-%s" % k, (x, y0, w, 190, 1000), [[(" ", 8, "#FFFFFF", False)]], fundo="#FFFFFF", borda="#D6DEE6", sombra=True)
    textbox(pg, "mini-tit-%s" % k, (x + 8, y0 + 6, w - 16, 34, 1100), [[(titulo, 10, NAVY, True)]])
    cartao(pg, "mini-val-%s" % k, (x + 10, y0 + 40, w - 20, 72, 1100), valor, tam=30, cor=NAVY)
    if meta: cartao(pg, "mini-meta-%s" % k, (x + 10, y0 + 112, w - 20, 40, 1100), meta, tam=10, cor=SUAVE, fonte="Segoe UI")
    if barra: cartao(pg, "mini-barra-%s" % k, (x + 10, y0 + 154, w - 20, 34, 1100), barra.replace("Barra ", "Barra Texto "), tam=11, cor_medida=cor, fonte="Segoe UI", quebra=False)
    if fatos: cartao(pg, "mini-fatos-%s" % k, (x + 10, y0 + 112, w - 20, 70, 1100), fatos, tam=10, cor=TINTA, fonte="Segoe UI")

