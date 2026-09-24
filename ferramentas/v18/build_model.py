# -*- coding: utf-8 -*-
"""v15 — alterações no modelo semântico (TMDL)."""
import os, sys, re, uuid
sys.path.insert(0, "/tmp"); sys.path.insert(0, os.path.dirname(__file__))
from build_lib import read, write, quote, indent_m
import medidas

ROOT = "/home/user/lab-analise-mes/Painel Samarco/Painel.SemanticModel/definition"
T = ROOT + "/tables"
HERE = os.path.dirname(os.path.abspath(__file__))
NS = uuid.UUID("7b3c1f10-0000-4000-8000-000000000015")
def tag(seed): return str(uuid.uuid5(NS, seed))

def col(table, name, dtype, fmt=None, hidden=False, summ="none", sort_by=None, calc=False, date=False):
    ls = ["\tcolumn %s" % quote(name), "\t\tdataType: %s" % dtype]
    if fmt: ls.append("\t\tformatString: %s" % fmt)
    if hidden: ls.append("\t\tisHidden")
    ls.append("\t\tlineageTag: %s" % tag(table + "|" + name))
    ls.append("\t\tsummarizeBy: %s" % summ)
    if calc: ls.append("\t\tisNameInferred")
    ls.append("\t\tsourceColumn: %s" % (("[%s]" % name) if calc else name))
    if sort_by: ls.append("\t\tsortByColumn: %s" % quote(sort_by))
    ls += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
    if date: ls += ["\t\tannotation UnderlyingDateTimeDataType = Date", ""]
    return "\n".join(ls)

def m_table(name, cols, m_code, hidden=False):
    head = ["table %s" % quote(name)]
    if hidden: head.append("\tisHidden")
    head += ["\tlineageTag: %s" % tag("tabela|" + name), ""]
    body = "\n".join(cols)
    part = ["\tpartition %s = m" % quote(name), "\t\tmode: import", "\t\tsource =", indent_m(m_code), "",
            "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(head) + "\n" + body + "\n" + "\n".join(part)

def M(f): return open(os.path.join(HERE, f), encoding="utf-8").read().rstrip()

# ── tbl_Afericoes ─────────────────────────────────────────────────────────
TA = "tbl_Afericoes"
cols = [col(TA, "Arquivo", "string"), col(TA, "Ensaio", "string"), col(TA, "Tipo", "string"),
        col(TA, "Data", "dateTime", "dd/mm/yyyy", date=True), col(TA, "Equipamento", "string"),
        col(TA, "Parâmetro", "string"), col(TA, "Unidade", "string"),
        col(TA, "Valor", "double", "#,0.00"), col(TA, "Referência", "double", "#,0.00"),
        col(TA, "Diferença", "double", "#,0.00"), col(TA, "Lim. Inferior", "double", "#,0.00"),
        col(TA, "Lim. Superior", "double", "#,0.00"), col(TA, "Tolerância", "double", "#,0.00"),
        col(TA, "Resultado", "string"), col(TA, "Responsável", "string"), col(TA, "Letra", "string"),
        col(TA, "Observação", "string"), col(TA, "Origem", "string"), col(TA, "Chave", "string", hidden=True),
        col(TA, "Conforme", "int64", "0", summ="sum"), col(TA, "TagKey", "string"),
        col(TA, "Grupo Operacional", "string"), col(TA, "% da Tolerância", "double", "0%"),
        col(TA, "Comparativo", "string"), col(TA, "Semana Início", "dateTime", "dd/mm/yyyy", date=True)]
write(T + "/tbl_Afericoes.tmdl", m_table(TA, cols, M("m_afericoes.m")))

TS = "tbl_Afericoes_Status"
cols = [col(TS, "Status", "string"), col(TS, "Planilhas na Pasta", "int64", "0", summ="sum"),
        col(TS, "Pastas de Ano", "int64", "0", summ="sum")]
write(T + "/tbl_Afericoes_Status.tmdl", m_table(TS, cols, M("m_status.m")))

TE = "DimEnsaio"
cols = [col(TE, "Ensaio", "string", sort_by="Ordem"), col(TE, "Ensaio Curto", "string", sort_by="Ordem"),
        col(TE, "Ordem", "int64", "0", hidden=True), col(TE, "Frequência", "string"),
        col(TE, "Tem Rotina", "boolean", hidden=True), col(TE, "Semanal", "boolean", hidden=True),
        col(TE, "Dias Semana", "string", hidden=True), col(TE, "Por Dia", "int64", "0", hidden=True)]
write(T + "/DimEnsaio.tmdl", m_table(TE, cols, M("m_dimensaio.m")))

TP = "tbl_Pontos"
cols = [col(TP, "Chave", "string"), col(TP, "Ordem", "int64", "0"), col(TP, "Tema", "string")]
write(T + "/tbl_Pontos.tmdl", m_table(TP, cols, M("m_pontos.m"), hidden=False))

TL = "tbl_Linhas"
cols = [col(TL, "Linha", "int64", "0")]
write(T + "/tbl_Linhas.tmdl", m_table(TL, cols, M("m_linhas.m")))

# ── _Painel (medidas) ─────────────────────────────────────────────────────
out = ["table _Painel", "\tlineageTag: %s" % tag("tabela|_Painel"), ""]
for nome, expr, fmt, pasta, oculta in medidas.M:
    out.append("\tmeasure %s =" % quote(nome))
    out.append("\t\t\t")
    for l in expr.split("\n"):
        out.append("\t\t\t" + l if l.strip() else "\t\t\t")
    if fmt: out.append("\t\tformatString: %s" % fmt)
    if oculta: out.append("\t\tisHidden")
    out.append("\t\tdisplayFolder: %s" % pasta)
    out.append("\t\tlineageTag: %s" % tag("medida|" + nome))
    out.append("")
out += ["\tcolumn Value", "\t\tisHidden", "\t\tformatString: 0", "\t\tlineageTag: %s" % tag("_Painel|Value"),
        "\t\tsummarizeBy: sum", "\t\tisNameInferred", "\t\tsourceColumn: [Value]", "",
        "\t\tannotation SummarizationSetBy = Automatic", "",
        "\tpartition _Painel = calculated", "\t\tmode: import", "\t\tsource = {1}", ""]
write(T + "/_Painel.tmdl", "\n".join(out))

# ── tbl_Calibracao: família de instrumento -> tbl_Grupos ──────────────────
p = T + "/tbl_Calibracao.tmdl"; s = read(p)
if "'Grupo Operacional'" not in s:
    old = "\t\t\t\tin\n\t\t\t\t    Lab\n"
    new = ("\t\t\t\t    ,\n" if False else "")
    assert old in s, "fim do M de tbl_Calibracao não encontrado"
    passo = [
        "",
        "    // v15: família de instrumento de laboratório (liga a calibração à saúde por grupo)",
        "    Grupo = Table.AddColumn( Lab, \"Grupo Operacional\",",
        "        each let k = if [TagKey] = null then \"\" else [TagKey]",
        "             in  if      Text.StartsWith( k, \"66TA\" ) then \"Tambores de Abrasão\"",
        "                 else if Text.StartsWith( k, \"66AG\" ) then \"Peneiradores Alpine\"",
        "                 else if Text.StartsWith( k, \"66PS\" ) then \"Prensas de Compressão\"",
        "                 else if Text.StartsWith( k, \"66AN\" ) then \"Analisadores de Umidade\"",
        "                 else null, type nullable text )",
        "in",
        "    Grupo"]
    # o passo anterior (Lab) é o último: precisa de vírgula
    lab_ini = s.index("\t\t\t\t    Lab = Table.TransformColumns( Tipo,")
    lab_fim = s.index(old)
    bloco = s[lab_ini:lab_fim].rstrip("\n")
    assert not bloco.endswith(","), bloco[-80:]
    s = s[:lab_ini] + bloco + ",\n" + "\n".join(("\t\t\t\t" + l) if l else "\t\t\t\t" for l in passo) + "\n" + s[lab_fim + len(old):]
    i = s.index("\tpartition tbl_Calibracao")
    s = s[:i] + col("tbl_Calibracao", "Grupo Operacional", "string") + "\n" + s[i:]
    write(p, s)

# ── DimCalendario: semana (início e rótulo) ───────────────────────────────
p = T + "/DimCalendario.tmdl"; s = read(p)
if "Semana Início" not in s:
    old = '\t\t\t\t    "Mês Atual",         EOMONTH ( [Data], 0 ) = EOMONTH ( _Hoje, 0 )\n'
    assert old in s
    new = ('\t\t\t\t    "Mês Atual",         EOMONTH ( [Data], 0 ) = EOMONTH ( _Hoje, 0 ),\n'
           '\t\t\t\t    "Semana Início",     [Data] - WEEKDAY ( [Data], 2 ) + 1,\n'
           '\t\t\t\t    "Semana Rótulo",     "sem. " & FORMAT ( [Data] - WEEKDAY ( [Data], 2 ) + 1, "dd/MM/yy" )\n')
    s = s.replace(old, new)
    i = s.index("\tpartition DimCalendario")
    s = s[:i] + col("DimCalendario", "Semana Início", "dateTime", "dd/mm/yyyy", calc=True, date=True) + "\n" + \
        col("DimCalendario", "Semana Rótulo", "string", calc=True, sort_by="Semana Início") + "\n" + s[i:]
    write(p, s)

# ── relacionamentos ───────────────────────────────────────────────────────
p = ROOT + "/relationships.tmdl"; s = read(p)
novos = [("tbl_Afericoes", "Data", "DimCalendario", "Data"),
         ("tbl_Afericoes", "Ensaio", "DimEnsaio", "Ensaio"),
         ("tbl_Afericoes", "Grupo Operacional", "tbl_Grupos", "Grupo Operacional"),
         ("tbl_Calibracao", "Grupo Operacional", "tbl_Grupos", "Grupo Operacional")]
for a, ca, b, cb in novos:
    rid = tag("rel|%s|%s" % (a, ca))
    if rid in s: continue
    s = s.rstrip("\n") + "\n\nrelationship %s\n\tfromColumn: %s.%s\n\ttoColumn: %s.%s\n" % (rid, a, quote(ca), b, quote(cb))
write(p, s)

# ── model.tmdl ────────────────────────────────────────────────────────────
p = ROOT + "/model.tmdl"; s = read(p)
for t in ["tbl_Afericoes", "tbl_Afericoes_Status", "DimEnsaio", "tbl_Pontos", "tbl_Linhas", "_Painel"]:
    if "ref table %s\n" % t not in s:
        s = s.replace("ref table _Medidas_Inspecoes\n", "ref table _Medidas_Inspecoes\nref table %s\n" % t)
s = s.replace('"DePara_Inspecao_Equipamento"]', '"DePara_Inspecao_Equipamento","tbl_Afericoes","DimEnsaio","tbl_Pontos","tbl_Linhas"]')
if '"tbl_Afericoes_Status"' not in s:
    s = s.replace('"tbl_Afericoes","DimEnsaio"', '"tbl_Afericoes","tbl_Afericoes_Status","DimEnsaio"')
write(p, s)
print("modelo v15 aplicado:", len(medidas.M), "medidas")
