import sys, zipfile, re, collections, datetime, openpyxl
from lxml import etree
f = sys.argv[1]
z = zipfile.ZipFile(f)
assert z.testzip() is None
bad = []
for n in z.namelist():
    if n.endswith(".xml") or n.endswith(".rels"):
        try: etree.fromstring(z.read(n))
        except Exception as e: bad.append((n, str(e)[:80]))
print("XML inválidos:", bad or "nenhum", "| vbaProject:", "xl/vbaProject.bin" in z.namelist(), "| calcChain:", "xl/calcChain.xml" in z.namelist())
for n in z.namelist():
    if n.startswith("xl/tables/"): print(n, re.search(r'name="([^"]+)"[^>]*ref="([^"]+)"', z.read(n).decode()).groups())
wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
print("abas:", wb.sheetnames)
ws = wb["BD_Afericoes"]; rows = list(ws.iter_rows(values_only=True)); H = rows[0]
regs = [dict(zip(H, r)) for r in rows[1:]]
val = [r for r in regs if r["Data"] not in (None, "") and r["Resultado"] not in (None, "")]
print("linhas:", len(regs), "| registros p/ Power BI:", len(val))
c = collections.Counter((r["Ensaio"], r["Resultado"]) for r in val)
for k in sorted(c): print("   ", k, c[k])
meses = collections.Counter((r["Data"].strftime("%Y-%m") if isinstance(r["Data"], datetime.datetime) else str(r["Data"])) for r in val)
print("por mês:", dict(meses))
bad = [ (i+2, r) for i, r in enumerate(regs) if any(isinstance(v, str) and v.startswith("#") for v in r.values())]
print("células com erro (#):", len(bad))
lim = list(wb["BD_Limites"].iter_rows(min_row=1, max_row=24, max_col=9, values_only=True))
for r in lim:
    if any(v is not None for v in r): print("   LIM", [str(v)[:40] if v is not None else "" for v in r[:8]])
