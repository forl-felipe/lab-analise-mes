# -*- coding: utf-8 -*-
"""Confere referências das medidas: Tabela[Coluna] existe; [Medida] existe (ou é coluna de ADDCOLUMNS/VAR)."""
import re, glob, io, sys
D = "Painel Samarco/Painel.SemanticModel/definition/tables"
tabs, meas = {}, set()
for f in glob.glob(D + "/*.tmdl"):
    s = io.open(f, encoding="utf-8").read().replace("\r\n", "\n")
    t = re.search(r"^table ('?)(.+?)\1$", s, re.M).group(2)
    tabs[t] = set(m.group(2) for m in re.finditer(r"^\tcolumn ('?)(.+?)\1$", s, re.M))
    meas |= set(m.group(2) for m in re.finditer(r"^\tmeasure ('?)(.+?)\1 =", s, re.M))
alvo = sys.argv[1] if len(sys.argv) > 1 else D + "/_Painel.tmdl"
s = io.open(alvo, encoding="utf-8").read().replace("\r\n", "\n")
blocos = re.split(r"\n\tmeasure ", s)[1:]
erros = 0
for b in blocos:
    nome = re.match(r"('?)(.+?)\1 =", b).group(2)
    corpo = "\n".join(l for l in b.split("\n")[1:] if l.startswith("\t\t\t"))
    corpo = re.sub(r'"[^"]*"', '""', corpo)                     # tira strings
    locais = set(re.findall(r'"(@\w+)"', b)) | {"Value1", "Value2"}
    for m in re.finditer(r"(\b[A-Za-z_][\w]*|'[^']+')\[([^\]]+)\]", corpo):
        t = m.group(1).strip("'"); c = m.group(2)
        if t not in tabs: print("%-32s tabela desconhecida: %s" % (nome, t)); erros += 1
        elif c not in tabs[t]: print("%-32s coluna desconhecida: %s[%s]" % (nome, t, c)); erros += 1
    for m in re.finditer(r"(?<![\w'\]])\[([^\]]+)\]", corpo):
        r = m.group(1)
        if r.startswith("@") or r in ("Value1", "Value2", "Value"): continue
        if r not in meas: print("%-32s medida desconhecida: [%s]" % (nome, r)); erros += 1
    bal = corpo.count("(") - corpo.count(")")
    if bal: print("%-32s parênteses desbalanceados (%+d)" % (nome, bal)); erros += 1
print("medidas conferidas:", len(blocos), "| problemas:", erros)
sys.exit(1 if erros else 0)
