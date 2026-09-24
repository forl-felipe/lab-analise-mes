# -*- coding: utf-8 -*-
"""Valida visual.json/page.json contra os esquemas públicos do PBIR (microsoft/json-schemas)."""
import json, glob, os, sys, urllib.request, collections
from jsonschema import Draft7Validator
from referencing import Registry, Resource
BASE = "https://raw.githubusercontent.com/microsoft/json-schemas/main/fabric/item/report/definition/"
CACHE = "/tmp/pbir_schema/cache"; os.makedirs(CACHE, exist_ok=True)
def baixa(url):
    f = os.path.join(CACHE, url.replace("/", "_").replace(":", ""))
    if not os.path.exists(f):
        u = url.replace("https://developer.microsoft.com/json-schemas/fabric/item/report/definition/", BASE)
        open(f, "wb").write(urllib.request.urlopen(u, timeout=30).read())
    return json.load(open(f, encoding="utf-8"))
def retrieve(uri):
    return Resource.from_contents(baixa(uri.split("#")[0]))
reg = Registry(retrieve=retrieve)
ID = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/"
vs = baixa(ID + "visualContainer/2.9.0/schema.json")
ps = baixa(ID + "page/2.1.0/schema.json")
VV = Draft7Validator(vs, registry=reg); PV = Draft7Validator(ps, registry=reg)
R = "/home/user/lab-analise-mes/Painel Samarco/Painel.Report/definition/pages"
novos = set(sys.argv[1:])
cont = collections.Counter(); exemplos = {}
for f in glob.glob(R + "/*/visuals/*/visual.json") + glob.glob(R + "/*/page.json"):
    d = json.load(open(f, encoding="utf-8")); d.pop("$schema", None)
    v = PV if f.endswith("page.json") else VV
    for e in v.iter_errors(d):
        chave = (e.message[:110], "/".join(str(p) for p in e.absolute_path)[:80])
        pag = f.split("/pages/")[1].split("/")[0]
        grupo = "NOVO" if pag in novos or any(n in f for n in novos) else "antigo"
        cont[(grupo,) + chave] += 1; exemplos.setdefault((grupo,) + chave, f)
for k, n in sorted(cont.items(), key=lambda kv: -kv[1])[:40]:
    print(n, k[0], "|", k[2], "|", k[1])
print("total de erros:", sum(cont.values()))
