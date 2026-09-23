# -*- coding: utf-8 -*-
"""Insere as abas ocultas BD_Afericoes e BD_Limites no .xlsm por cirurgia de XML
(sem abrir/salvar com openpyxl, para preservar VBA, gráficos, desenhos e modos de exibição)."""
import zipfile, re, sys, json
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter as L
import spec

SRC, DST = sys.argv[1], sys.argv[2]
CACHE = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else {}

z = zipfile.ZipFile(SRC)
files = {n: z.read(n) for n in z.namelist()}
infos = {i.filename: i for i in z.infolist()}

# ---------- shared strings
sst = files["xl/sharedStrings.xml"].decode("utf8")
sis = re.findall(r"<si>.*?</si>", sst, re.S)
idx = {}
for i, si in enumerate(sis):
    m = re.fullmatch(r'<si><t(?: xml:space="preserve")?>(.*?)</t></si>', si, re.S)
    if m: idx.setdefault(m.group(1), i)
novos, nrefs = [], 0
def s_id(txt):
    global nrefs
    nrefs += 1
    k = escape(txt)
    if k not in idx:
        idx[k] = len(sis) + len(novos)
        sp = ' xml:space="preserve"' if txt != txt.strip() else ""
        novos.append("<si><t%s>%s</t></si>" % (sp, k))
    return idx[k]

# ---------- estilos: um xf de data (dd/mm/aaaa) e um de número 0,00
sty = files["xl/styles.xml"].decode("utf8")
nx = int(re.search(r'<cellXfs count="(\d+)">', sty).group(1))
XF_DATE, XF_HEAD = nx, nx + 1
sty = sty.replace('<numFmts count="4">', '<numFmts count="5">').replace(
    "</numFmts>", '<numFmt numFmtId="168" formatCode="dd/mm/yyyy"/></numFmts>')
nfont = int(re.search(r'<fonts count="(\d+)"', sty).group(1))
sty = re.sub(r'<fonts count="\d+"', '<fonts count="%d"' % (nfont + 1), sty, 1)
sty = sty.replace("</fonts>", '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>', 1)
sty = sty.replace('<cellXfs count="%d">' % nx, '<cellXfs count="%d">' % (nx + 2))
cx_end = sty.index("</cellXfs>")
sty = sty[:cx_end] + ('<xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                      '<xf numFmtId="0" fontId="%d" fillId="0" borderId="0" xfId="0" applyFont="1"/>' % nfont) + sty[cx_end:]
files["xl/styles.xml"] = sty.encode("utf8")

# ---------- geração de célula
def cell(ref, val, sheet, style=None):
    s = ' s="%d"' % style if style is not None else ""
    if val is None or val == "": return ""
    if isinstance(val, tuple):                       # fórmula
        f = escape(val[1])
        cv = CACHE.get(sheet, {}).get(ref)
        if isinstance(cv, (int, float)) and not isinstance(cv, bool):
            return '<c r="%s"%s><f>%s</f><v>%r</v></c>' % (ref, s, f, float(cv) if isinstance(cv, float) else cv)
        return '<c r="%s"%s t="str"><f>%s</f><v>%s</v></c>' % (ref, s, f, escape(cv if isinstance(cv, str) else ""))
    if isinstance(val, (int, float)):
        return '<c r="%s"%s><v>%r</v></c>' % (ref, s, val)
    return '<c r="%s"%s t="s"><v>%d</v></c>' % (ref, s, s_id(val))

def sheet_xml(sheet, header, rows, widths, extra_rows=()):
    ncol = len(header)
    out = ['<row r="1">' + "".join(cell("%s1" % L(i + 1), h, sheet, XF_HEAD) for i, h in enumerate(header)) + "</row>"]
    for n, row in enumerate(rows, start=2):
        cs = "".join(cell("%s%d" % (L(i + 1), n), row[i], sheet, XF_DATE if header[i] == "Data" else None) for i in range(ncol))
        out.append('<row r="%d">%s</row>' % (n, cs))
    for n, texto, st in extra_rows:
        out.append('<row r="%d">%s</row>' % (n, cell("A%d" % n, texto, sheet, st)))
    last = max([len(rows) + 1] + [n for n, _, _ in extra_rows])
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(widths))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetPr><tabColor rgb="FF003A70"/></sheetPr><dimension ref="A1:%s%d"/>'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>%s</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
            % (L(ncol), last, cols, "".join(out))).encode("utf8")

def table_xml(tid, name, header, nrows):
    ref = "A1:%s%d" % (L(len(header)), nrows + 1)
    tc = "".join('<tableColumn id="%d" name="%s"/>' % (i + 1, escape(h)) for i, h in enumerate(header))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%d" name="%s" displayName="%s" ref="%s" totalsRowShown="0">'
            '<autoFilter ref="%s"/><tableColumns count="%d">%s</tableColumns>'
            '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
            % (tid, name, name, ref, ref, len(header), tc)).encode("utf8")

REL_T = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%d.xml"/></Relationships>')

# ---------- BD_Afericoes
rows = spec.resolve()
H = spec.COLS
grid = [[r.get(L(i + 1)) for i in range(len(H))] for r in rows]
files["xl/worksheets/sheet15.xml"] = sheet_xml("BD_Afericoes", H, grid,
    [22, 13, 12, 26, 30, 9, 11, 11, 11, 11, 11, 11, 14, 16, 7, 34, 34])
files["xl/tables/table1.xml"] = table_xml(1, "tbl_Afericoes", H, len(grid))
files["xl/worksheets/_rels/sheet15.xml.rels"] = (REL_T % 1).encode("utf8")

# ---------- BD_Limites (parâmetros + legenda)
HL = ["Chave", "Ensaio", "Parâmetro", "Unidade", "Nominal", "Lim. Inferior", "Lim. Superior", "Tolerância", "Fonte"]
lgrid = [list(t) for t in spec.LIM]
n0 = len(lgrid) + 3
legenda = [
    (n0, "COMO FUNCIONA", XF_HEAD),
    (n0 + 1, "Esta aba e a BD_Afericoes são geradas para o Power BI. Os operadores continuam preenchendo as abas de sempre; nada aqui precisa ser digitado.", None),
    (n0 + 2, "BD_Afericoes (tabela tbl_Afericoes): 1 linha por equipamento x ensaio x data, com fórmulas que apontam para as células das abas de preenchimento (coluna Origem).", None),
    (n0 + 3, "Resultado é recalculado aqui de forma padronizada: Calibração = valor dentro de Lim. Inferior/Superior; Comparativo = |Diferença| <= Tolerância; Peneiradores = nenhum item 'NÃO OK'.", None),
    (n0 + 4, "Para mudar um limite ou tolerância, altere SOMENTE as colunas Nominal / Lim. Inferior / Lim. Superior / Tolerância desta tabela (tbl_Limites). As fórmulas da BD_Afericoes leem daqui.", None),
    (n0 + 5, "Granulometria e Tamb 5x15 são recalculados a partir das massas (g) digitadas, e não das colunas de % das abas, para não herdar fórmulas quebradas.", None),
    (n0 + 6, "Não inserir/excluir linhas ou colunas nas abas dos operadores: as fórmulas apontam para posições fixas. Para um novo mês, copie este arquivo (Salvar como) e limpe só os valores digitados.", None),
]
files["xl/worksheets/sheet16.xml"] = sheet_xml("BD_Limites", HL, lgrid, [8, 20, 22, 10, 10, 13, 13, 12, 70], legenda)
files["xl/tables/table2.xml"] = table_xml(2, "tbl_Limites", HL, len(lgrid))
files["xl/worksheets/_rels/sheet16.xml.rels"] = (REL_T % 2).encode("utf8")

# ---------- shared strings final
sst_new = re.sub(r'count="\d+" uniqueCount="\d+"', 'count="%d" uniqueCount="%d"' % (
    int(re.search(r'count="(\d+)"', sst).group(1)) + nrefs, len(sis) + len(novos)), sst, 1)
sst_new = sst_new.replace("</sst>", "".join(novos) + "</sst>")
files["xl/sharedStrings.xml"] = sst_new.encode("utf8")

# ---------- workbook.xml, rels, content types
wb = files["xl/workbook.xml"].decode("utf8")
assert "BD_Afericoes" not in wb
wb = wb.replace("</sheets>", '<sheet name="BD_Afericoes" sheetId="21" state="hidden" r:id="rId25"/>'
                             '<sheet name="BD_Limites" sheetId="22" state="hidden" r:id="rId26"/></sheets>')
wb = wb.replace('<calcPr calcId="191028"/>', '<calcPr calcId="191028" fullCalcOnLoad="1"/>')
assert 'fullCalcOnLoad="1"' in wb
files["xl/workbook.xml"] = wb.encode("utf8")
rl = files["xl/_rels/workbook.xml.rels"].decode("utf8")
assert 'Id="rId25"' not in rl
W = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
rl = rl.replace("</Relationships>", '<Relationship Id="rId25" Type="%s" Target="worksheets/sheet15.xml"/>'
                '<Relationship Id="rId26" Type="%s" Target="worksheets/sheet16.xml"/></Relationships>' % (W, W))
files["xl/_rels/workbook.xml.rels"] = rl.encode("utf8")
ct = files["[Content_Types].xml"].decode("utf8")
WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
TB = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
ct = ct.replace("</Types>", "".join('<Override PartName="/%s" ContentType="%s"/>' % (p, t) for p, t in [
    ("xl/worksheets/sheet15.xml", WS), ("xl/worksheets/sheet16.xml", WS),
    ("xl/tables/table1.xml", TB), ("xl/tables/table2.xml", TB)]) + "</Types>")
files["[Content_Types].xml"] = ct.encode("utf8")

# ---------- grava preservando ordem e compressão originais
with zipfile.ZipFile(DST, "w") as out:
    for n in z.namelist():
        i = infos[n]
        zi = zipfile.ZipInfo(n, date_time=i.date_time); zi.compress_type = i.compress_type; zi.external_attr = i.external_attr
        out.writestr(zi, files[n])
    for n in ["xl/worksheets/sheet15.xml", "xl/worksheets/sheet16.xml", "xl/worksheets/_rels/sheet15.xml.rels",
              "xl/worksheets/_rels/sheet16.xml.rels", "xl/tables/table1.xml", "xl/tables/table2.xml"]:
        out.writestr(n, files[n], compress_type=zipfile.ZIP_DEFLATED)
print("ok", DST, "linhas:", len(grid), "strings novas:", len(novos))
