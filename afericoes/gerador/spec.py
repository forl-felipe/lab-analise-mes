# -*- coding: utf-8 -*-
"""Gera a lista de linhas (fórmulas) da aba oculta BD_Afericoes."""
from openpyxl.utils import get_column_letter as L

TAMB = "'Tambor de Abrasão'"; ALP = "'Gran. Fina Alpine'"; BLA = "'Calibração Blaine'"
FIS = "'Comparativo Fisher'"; PEN = "'Verificação Peneiradores'"; COM = "Compressão"
GRA = "Granulometria"; UMI = "Umidade"; T515 = "'Tamb Kg x Tamb 15,0Kg'"

# Colunas da tabela tbl_Afericoes
COLS = ["Ensaio", "Tipo", "Data", "Equipamento", "Parâmetro", "Unidade", "Valor",
        "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância",
        "Resultado", "Responsável", "Letra", "Observação", "Origem"]
# A B C D E F G H I J K L M N O P Q

# Linhas de BD_Limites (linha 2 em diante): chave -> (Ensaio, Parâmetro, Unidade, Nominal, LimInf, LimSup, Tol, Fonte)
LIM = [
 ("TAMB",  "Tambor de Abrasão", "Rotação", "rpm", 25, 24, 26, None, "Aba Tambor de Abrasão, A4/A23: 25 ± 1 rpm (limites X7/Z7)"),
 ("ALP",   "Gran. Fina Alpine", "Passante", "%", 88.4, 87.4, 89.4, None, "Aba Gran. Fina Alpine, linha 5 e colunas I:K"),
 ("BLAM",  "Blaine", "Blaine Manual", "cm²/g", 2030, 1986, 2074, None, "Aba Calibração Blaine, colunas E:G e fórmula I6"),
 ("BLAS",  "Blaine", "Calibração Star", "cm²/g", 2120, 2076, 2164, None, "Aba Calibração Blaine, fórmula K6"),
 ("BLAA",  "Blaine", "Blaine Automático", "cm²/g", 2030, 1986, 2074, None, "Aba Calibração Blaine, fórmula M6"),
 ("FIS",   "Comparativo Fisher", "Superfície específica", "cm²/g", None, None, None, 300, "Aba Comparativo Fisher, F6: diferença aceitável 300 cm²/g"),
 ("COM16", "Compressão", "FX -16,0 +12,5 mm", "kgf/pel", 355, None, None, 20, "Aba Compressão, B1: nominal 355 kgf/pel, diferença aceitável 20"),
 ("COM12", "Compressão", "FX -12,5 +10,0 mm", "kgf/pel", 310, None, None, 20, "Aba Compressão, B16: nominal 310 kgf/pel, diferença aceitável 20"),
 ("VEL",   "Compressão", "Velocidade", "mm/min", 15, 13, 17, None, "Aba Compressão, B1: velocidade 15 ± 2 mm/min"),
 ("G63",   "Granulometria", "-6,3 mm", "%", None, None, None, 0.30, "Aba Granulometria, R11: tolerância 0,30"),
 ("GRG",   "Granulometria", "RG (Relação)", "-", None, None, None, 0.15, "Aba Granulometria, R10: tolerância 0,15"),
 ("UMI",   "Umidade", "Umidade", "%", None, None, None, 0.05, "Aba Umidade, coluna I: diferença aceitável 0,05"),
 ("T515",  "Tamb 5 kg x 15 kg", "+6,3 mm", "%", None, None, None, 0.5, "Aba Tamb Kg x Tamb 15,0Kg, V10: diferença aceitável 0,5"),
]
LROW = {k: i + 2 for i, (k, *_) in enumerate(LIM)}
# BD_Limites: A Chave, B Ensaio, C Parâmetro, D Unidade, E Nominal, F Lim. Inferior, G Lim. Superior, H Tolerância, I Fonte
_LC = {"D": "E", "E": "F", "F": "G", "G": "H"}   # apelidos usados abaixo: D=Nominal E=LimInf F=LimSup G=Tol
def lim(k, col):
    return "BD_Limites!$%s$%d" % (_LC[col], LROW[k])

def num(ref): return 'IF(ISNUMBER(%s),%s,"")' % (ref, ref)
def dt(ref):  return 'IF(N(%s)>0,%s,"")' % (ref, ref)
def tx(ref):  return 'IF(%s&""="","",TRIM(%s&""))' % (ref, ref)

# Regras de resultado (r = linha da própria BD_Afericoes)
def res_lim(r):  return 'IF(G{0}="","",IF(AND(G{0}>=J{0},G{0}<=K{0}),"Conforme","Não conforme"))'.format(r)
def res_tol(r):  return 'IF(I{0}="","",IF(ROUND(ABS(I{0}),6)<=L{0},"Conforme","Não conforme"))'.format(r)
def dif(r):      return 'IF(OR(G{0}="",H{0}=""),"",G{0}-H{0})'.format(r)

rows = []   # cada linha: dict coluna->("f", formula) ou valor constante
def add(**kw): rows.append(kw)

def F(s): return ("f", s)

# ---------------- Tambor de Abrasão (calibração, limites 24–26 rpm)
for tag, col, r0, r1, nome in [("66TA05","E",6,14,"Q"),("66TA08","J",6,14,"Q"),("66TA09","O",6,14,"Q"),
                               ("66TA06","E",25,34,"L"),("66TA07","J",25,34,"L")]:
    for r in range(r0, r1 + 1):
        add(A="Tambor de Abrasão", B="Calibração", C=F(dt("%s!$A$%d" % (TAMB, r))), D=tag,
            E="Rotação", F="rpm", G=F(num("%s!$%s$%d" % (TAMB, col, r))), H=F(lim("TAMB","D")),
            I="DIF", J=F(lim("TAMB","E")), K=F(lim("TAMB","F")), M="LIM",
            N=F(tx("%s!$%s$%d" % (TAMB, nome, r))), Q="Tambor de Abrasão!%s%d" % (col, r))

# ---------------- Alpine (calibração, 87,4–89,4 %)
for tag, col, pen in [("66AG09","B","66PN670"),("66AG10","D","66PN671"),("66AG11","F","66PN671")]:
    for r in range(7, 22):
        add(A="Gran. Fina Alpine", B="Calibração", C=F(dt("%s!$A$%d" % (ALP, r))), D=tag,
            E="Passante (peneira %s)" % pen, F="%", G=F(num("%s!$%s$%d" % (ALP, col, r))), H=F(lim("ALP","D")),
            I="DIF", J=F(lim("ALP","E")), K=F(lim("ALP","F")), M="LIM",
            N=F(tx("%s!$H$%d" % (ALP, r))), Q="Gran. Fina Alpine!%s%d" % (col, r))

# ---------------- Blaine (manual, Star, automático)
for k, eq, col in [("BLAM","Blaine Manual","H"),("BLAS","Blaine Star","J"),("BLAA","Blaine Automático","L")]:
    for r in range(6, 72):
        eqf = F('IF(%s!$D$%d&""="","%s",TRIM(%s!$D$%d&""))' % (BLA, r, "66PB08", BLA, r)) if k == "BLAM" else eq
        add(A="Blaine", B="Calibração", C=F(dt("%s!$A$%d" % (BLA, r))), D=eqf,
            E=LIM[LROW[k]-2][2], F="cm²/g", G=F(num("%s!$%s$%d" % (BLA, col, r))), H=F(lim(k,"D")),
            I="DIF", J=F(lim(k,"E")), K=F(lim(k,"F")), M="LIM",
            N=F(tx("%s!$N$%d" % (BLA, r))), O=F(tx("%s!$C$%d" % (BLA, r))),
            P=F(tx("%s!$B$%d" % (BLA, r))), Q="Calibração Blaine!%s%d" % (col, r))

# ---------------- Fisher (comparativo 66PB01 x 66PB09, tolerância 300)
for r in range(8, 13):
    add(A="Comparativo Fisher", B="Comparativo", C=F(dt("%s!$A$%d" % (FIS, r))), D="66PB01 x 66PB09",
        E="Superfície específica", F="cm²/g", G=F(num("%s!$D$%d" % (FIS, r))), H=F(num("%s!$E$%d" % (FIS, r))),
        I="DIF", L=F(lim("FIS","G")), M="TOL", N=F(tx("%s!$B$%d" % (FIS, r))), Q="Comparativo Fisher!D%d:E%d" % (r, r))

# ---------------- Compressão (valor x nominal, tolerância 20; velocidade 13–17)
for k, r0 in [("COM16", 6), ("COM12", 20)]:
    fx = LIM[LROW[k]-2][2]
    for r in range(r0, r0 + 4):
        for tag, col, vcol in [("66PS04","C","E"),("66PS07","F","H"),("66PS08","I","K"),("66PS10","L",None),("66PS05","N","P")]:
            src = "%s!$%s$%d" % (COM, col, r)
            add(A="Compressão", B="Comparativo", C=F(dt("%s!$B$%d" % (COM, r))), D=tag,
                E="Resistência %s" % fx, F="kgf/pel", G=F(num(src)), H=F(lim(k,"D")), I="DIF",
                L=F(lim(k,"G")), M="TOL", N=F(tx("%s!$W$%d" % (COM, r))),
                P=F('IF(ISTEXT(%s),TRIM(%s),"")&IF(%s!$X$%d&""="","",IF(ISTEXT(%s)," - ","")&TRIM(%s!$X$%d&""))' % (src, src, COM, r, src, COM, r)),
                Q="Compressão!%s%d" % (col, r))
            if vcol:
                add(A="Compressão", B="Verificação", C=F(dt("%s!$B$%d" % (COM, r))), D=tag,
                    E="Velocidade %s" % fx, F="mm/min", G=F(num("%s!$%s$%d" % (COM, vcol, r))), H=F(lim("VEL","D")),
                    I="DIF", J=F(lim("VEL","E")), K=F(lim("VEL","F")), M="LIM",
                    N=F(tx("%s!$W$%d" % (COM, r))), Q="Compressão!%s%d" % (vcol, r))

# ---------------- Granulometria (LTF 66PN09 x LCE 66PN08) — recalculado a partir das massas (g)
def gpct(c, r, a, b):   # % das faixas das linhas r+a..r+b sobre o total (linhas r+5..r+14)
    return 'IF(N(SUM({s}!${c}${t0}:${c}${t1}))=0,"",SUM({s}!${c}${a0}:${c}${a1})/SUM({s}!${c}${t0}:${c}${t1})*100)'.format(
        s=GRA, c=c, t0=r+5, t1=r+14, a0=r+a, a1=r+b)
def grg(c, r):
    return 'IF(N({s}!${c}${d})=0,"",({s}!${c}${a}+{s}!${c}${b})/{s}!${c}${d})'.format(s=GRA, c=c, a=r+7, b=r+8, d=r+9)
for r in (7, 31, 55):
    for lado, cl, cr, cd, cn in [("esq", "C", "F", "G", "C"), ("dir", "J", "M", "N", "J")]:
        for par, un, kind in [("-6,3 mm", "%", ("P", 12, 14, "G63")), ("RG (Relação)", "-", ("RG", 0, 0, "GRG")),
                              ("-16,0 +12,5 mm", "%", ("P", 7, 8, None)), ("-16,0 +9,0 mm", "%", ("P", 7, 9, None)),
                              ("-16,0 +8,0 mm", "%", ("P", 7, 10, None))]:
            t, a, b, k = kind
            gv = gpct(cl, r, a, b) if t == "P" else grg(cl, r)
            gh = gpct(cr, r, a, b) if t == "P" else grg(cr, r)
            add(A="Granulometria", B="Comparativo", C=F(dt("%s!$%s$%d" % (GRA, cd, r))), D="66PN09 (LTF) x 66PN08 (LCE)",
                E=par, F=un, G=F(gv), H=F(gh), I="DIF", L=F(lim(k, "G")) if k else None,
                M="TOL" if k else "INFO", N=F(tx("%s!$%s$%d" % (GRA, cn, r))),
                Q="Granulometria!%s%d:%s%d" % (cl, r+5, cr, r+14))

# ---------------- Umidade (analisadores x estufa, tolerância 0,05)
for tag, col in [("66AN10","B"),("66AN11","C"),("66AN12","D")]:
    for r in range(5, 19):
        add(A="Umidade", B="Comparativo", C=F(dt("%s!$A$%d" % (UMI, r))), D=tag + " x Estufa",
            E="Umidade", F="%", G=F(num("%s!$%s$%d" % (UMI, col, r))), H=F(num("%s!$E$%d" % (UMI, r))),
            I="DIF", L=F(lim("UMI","G")), M="TOL", N=F(tx("%s!$L$%d" % (UMI, r))), O=F(tx("%s!$K$%d" % (UMI, r))),
            P=F(tx("%s!$N$%d" % (UMI, r))), Q="Umidade!%s%d" % (col, r))

# ---------------- Tamboramento 5 kg (LTF) x 15 kg (T03), tolerância 0,5
for r in (7, 16, 25):
    for cd, cn, c5, c15 in [("D","G","C","G"), ("L","O","K","O")]:
        g5 = "%s!$%s$%d" % (T515, c5, r+4); g15 = "%s!$%s$%d" % (T515, c15, r+4)
        add(A="Tamb 5 kg x 15 kg", B="Comparativo", C=F(dt("%s!$%s$%d" % (T515, cd, r))), D="66TA05 (5 kg) x T03 (15 kg)",
            E="+6,3 mm", F="%", G=F('IF(N(%s)=0,"",%s/50)' % (g5, g5)), H=F('IF(N(%s)=0,"",%s/150)' % (g15, g15)),
            I="DIF", L=F(lim("T515","G")), M="TOL", N=F(tx("%s!$%s$%d" % (T515, cn, r))),
            Q="Tamb Kg x Tamb 15,0Kg!%s%d:%s%d" % (c5, r+4, c15, r+4))

# ---------------- Verificação dos peneiradores (13 blocos x 27 peneiras)
ITENS = ["Condições das telas", "Inclinação das telas", "Vibração do peneirador", "Travamento das peneiras", "Calibração"]
def grupo(ci):
    return "Circulares" if ci <= 12 else "Peneirador 09" if ci <= 19 else "Peneirador 05" if ci <= 26 else "Peneirador 06"
for k in range(13):
    r = 3 + 10 * k
    for ci in range(2, 29):          # B..AB
        c = L(ci); rng = "%s!$%s$%d:$%s$%d" % (PEN, c, r+3, c, r+7)
        obs = "&".join('IF(%s!$%s$%d="NÃO OK","%s; ","")' % (PEN, c, r+3+i, it) for i, it in enumerate(ITENS))
        add(A="Verificação Peneiradores", B="Verificação",
            C=F('IF(MAX(%s!$AC$%d:$AC$%d)=0,"",MAX(%s!$AC$%d:$AC$%d))' % (PEN, r, r+9, PEN, r, r+9)),
            D=F('IF(%s!$%s$%d&""="","%s",TRIM(%s!$%s$%d&""))' % (PEN, c, r+1, grupo(ci), PEN, c, r+1)),
            E=F('"Malha "&%s!$%s$%d&" mm ("&"%s"&")"' % (PEN, c, r, grupo(ci))), F="itens OK",
            G=F('IF(COUNTA(%s)=0,"",COUNTIF(%s,"OK"))' % (rng, rng)), H=F('IF(COUNTA(%s)=0,"",COUNTA(%s))' % (rng, rng)),
            M=F('IF(COUNTA({0})=0,"",IF(COUNTIF({0},"NÃO OK")>0,"Não conforme","Conforme"))'.format(rng)),
            N=F('IFERROR(TRIM(INDEX(%s!$AE$%d:$AE$%d,MATCH("*",%s!$AE$%d:$AE$%d,0))),"")' % (PEN, r, r+9, PEN, r, r+9)),
            O=F('IFERROR(TRIM(INDEX(%s!$AD$%d:$AD$%d,MATCH("*",%s!$AD$%d:$AD$%d,0))),"")' % (PEN, r, r+9, PEN, r, r+9)),
            P=F('IF(COUNTIF(%s,"SEM TAG")>0,"Sem TAG de calibração; ","")&%s' % (rng, obs)),
            Q="Verificação Peneiradores!%s%d:%s%d" % (c, r+3, c, r+7))

def resolve():
    """Troca os marcadores DIF/LIM/TOL/INFO pelas fórmulas da própria linha."""
    out = []
    for i, row in enumerate(rows):
        n = i + 2
        row = dict(row)
        if row.get("I") == "DIF": row["I"] = F(dif(n))
        m = row.get("M")
        if m == "LIM": row["M"] = F(res_lim(n))
        elif m == "TOL": row["M"] = F(res_tol(n))
        elif m == "INFO": row["M"] = F('IF(G{0}="","","Informativo")'.format(n))
        out.append(row)
    return out
