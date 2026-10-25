"""
Confere se as regras do sistema (ferramentas/regras.py, iguais às do app) dão os
mesmos resultados que a planilha de aferição.

Uso:
    python validar_regras.py "planilha.xlsm" pasta_da_carga
        pasta_da_carga = saída do gerar_carga.py (usa csv/Limites.csv e csv/PontosAfericao.csv)

Testes:
  A. Cada resultado da BD_Afericoes recalculado com o critério gravado na própria linha.
  B. Cada resultado recalculado com os critérios ATUAIS da lista Limites
     (o que o app faria hoje com os mesmos valores).
  C. Os cálculos da aba Lançamento (rpm, umidade, granulometria, tamboramento)
     refeitos a partir das entradas e comparados com as células calculadas pela planilha.
  D. Aderência à agenda (previstos x realizados por ensaio) comparada com a aba Painel.
Gera relatorio-validacao.md na pasta da carga.
"""

import csv
import sys
from collections import Counter
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "ferramentas"))
from regras import Limite, aderencia, avaliar, granulometria, rpm, tamboramento, umidade_estufa  # noqa: E402


def f(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main(planilha: Path, pasta: Path):
    wb = openpyxl.load_workbook(planilha, data_only=True, read_only=True)
    bd = [r for r in wb["BD_Afericoes"].iter_rows(min_row=2, max_col=21, values_only=True) if r[0]]
    lan = {(r, c): v for r, row in enumerate(wb["Lançamento"].iter_rows(max_row=160, max_col=20, values_only=True), 1)
           for c, v in enumerate(row, 1) if v not in (None, "")}
    limites = {r["Title"]: r for r in csv.DictReader((pasta / "csv" / "Limites.csv").open(encoding="utf-8-sig"))}
    pontos = list(csv.DictReader((pasta / "csv" / "PontosAfericao.csv").open(encoding="utf-8-sig")))
    ensaios = {r["Title"]: r for r in csv.DictReader((pasta / "csv" / "Ensaios.csv").open(encoding="utf-8-sig"))}

    def lim(chave):
        r = limites.get(chave)
        return None if not r else Limite(f(r["Nominal"]), f(r["LimInferior"]), f(r["LimSuperior"]), f(r["Tolerancia"]))

    # (NomeBD, Parametro) -> chave de limite, a partir dos pontos e dos indicadores calculados
    mapa = {}
    for p in pontos:
        nomebd = ensaios[p["Ensaio"]]["NomeBD"]
        if p["ChaveLimite"]:
            mapa[(nomebd, p["Parametro"])] = p["ChaveLimite"]
    mapa[("Granulometria", "-6,3 mm")] = "G63"
    mapa[("Granulometria", "RG (Relação)")] = "GRG"
    mapa[("Tamb 5 kg x 15 kg", "+6,3 mm")] = "T515"

    linhas = ["# Relatório de validação das regras", "", f"Planilha: `{planilha.name}` · {len(bd)} resultados na BD_Afericoes", ""]

    # ---------------- A ----------------
    okA, difA, pulados = 0, [], 0
    for r in bd:
        ens, par, val, ref, li, ls, tol, res = r[0], r[4], f(r[6]), f(r[7]), f(r[9]), f(r[10]), f(r[11]), r[12]
        if r[5] == "itens OK":
            pulados += 1
            continue
        crit = None if li is None and ls is None and tol is None else Limite(ref, li, ls, tol)
        _, calc = avaliar(val, ref, crit)
        if calc == res:
            okA += 1
        else:
            difA.append((r[2].date(), ens, r[3], par, val, ref, res, calc))
    linhas += ["## A. Resultado recalculado com o critério gravado na linha", "",
               f"- Conferem: **{okA} de {len(bd) - pulados}**",
               f"- Verificação dos peneiradores ({pulados} linhas): a planilha grava só a contagem de itens, não os itens; a regra foi conferida pela fórmula (qualquer NÃO OK reprova; SEM TAG não reprova).", ""]
    if difA:
        linhas += ["| Data | Ensaio | Equipamento | Parâmetro | Valor | Ref. | Planilha | Regra |", "|---|---|---|---|---|---|---|---|"]
        linhas += [f"| {d} | {e} | {q} | {p} | {v} | {rf} | {a} | {b} |" for d, e, q, p, v, rf, a, b in difA]
        linhas.append("")

    # ---------------- B ----------------
    okB, difB, semchave = 0, [], Counter()
    for r in bd:
        ens, par, val, ref, res = r[0], r[4], f(r[6]), f(r[7]), r[12]
        if r[5] == "itens OK":
            continue
        chave = mapa.get((ens, par))
        if not chave:
            semchave[(ens, par)] += 1
            if res == "Informativo":
                okB += 1
            else:
                difB.append((r[2].date(), ens, r[3], par, val, ref, res, "sem critério"))
            continue
        L = lim(chave)
        referencia = ref if ens in ("Umidade", "Granulometria", "Tamb 5 kg x 15 kg", "Comparativo Fisher") else L.nominal
        _, calc = avaliar(val, referencia, L)
        if calc == res:
            okB += 1
        else:
            difB.append((r[2].date(), ens, r[3], par, val, referencia, res, calc))
    total_b = len(bd) - pulados
    linhas += ["## B. Resultado com os critérios atuais da lista Limites", "",
               f"- Conferem: **{okB} de {total_b}**", ""]
    if difB:
        linhas += ["Diferenças (normalmente indicam critério alterado durante o mês):", "",
                   "| Data | Ensaio | Equipamento | Parâmetro | Valor | Ref. | Planilha | App hoje |", "|---|---|---|---|---|---|---|---|"]
        linhas += [f"| {d} | {e} | {q} | {p} | {v} | {rf} | {a} | {b} |" for d, e, q, p, v, rf, a, b in difB]
        linhas.append("")
    if semchave:
        linhas += ["Parâmetros sem critério (gravados como Informativo): " + ", ".join(f"{e} / {p} ({n})" for (e, p), n in semchave.items()), ""]

    # ---------------- C ----------------
    def c(ref):
        import re
        m = re.match(r"([A-Z]+)(\d+)", ref)
        col = 0
        for ch in m.group(1):
            col = col * 26 + ord(ch) - 64
        return lan.get((int(m.group(2)), col))

    testes = []

    def comparar(nome, esperado, obtido, tol=1e-6):
        e, o = f(esperado), f(obtido)
        if e is None and o is None:
            return
        ok = e is not None and o is not None and abs(e - o) <= tol * max(1, abs(e))
        testes.append((nome, esperado, obtido, ok))

    for r in range(33, 38):
        v = f(c(f"E{r}"))
        if c(f"F{r}") is not None or c(f"G{r}") is not None:
            comparar(f"Tambor {c(f'C{r}')} · rpm", c(f"I{r}"), rpm(v, f(c(f"F{r}")), f(c(f"G{r}"))))
    est = f(c("F50"))
    if est is None:
        est = umidade_estufa(f(c("G51")), f(c("I51")), f(c("K51")))
    if est is not None:
        comparar("Estufa (%)", c("N51"), est)
    ma = [f(c(f"F{r}")) for r in range(78, 88)]
    mb = [f(c(f"I{r}")) for r in range(78, 88)]
    dm = [float(p["Fator"]) for p in pontos if p["Ensaio"] == "GRA"]
    if any(ma) and any(mb):
        g = granulometria(ma, mb, dm)
        for nome, linha in (("-6,3 mm", 91), ("RG (Relação)", 92), ("-16,0 +12,5 mm", 93), ("-16,0 +9,0 mm", 94), ("-16,0 +8,0 mm", 95), ("Diâmetro médio", 96)):
            comparar(f"Granulometria {nome} · LTF", c(f"F{linha}"), g[nome][0])
            comparar(f"Granulometria {nome} · LCE", c(f"H{linha}"), g[nome][1])
    if c("F102") is not None and c("F103") is not None:
        t = tamboramento(f(c("F102")), f(c("H102")), 5000, f(c("F103")), f(c("H103")), 15000)
        comparar("Tamboramento +6,3 · LTF", c("J102"), t["+6,3 mm"][0])
        comparar("Tamboramento +6,3 · LCE", c("J103"), t["+6,3 mm"][1])
    okC = sum(1 for t in testes if t[3])
    linhas += ["## C. Cálculos da aba Lançamento refeitos pelas regras do app", "",
               f"- Conferem: **{okC} de {len(testes)}**", "",
               "| Cálculo | Planilha | Regras do app | OK |", "|---|---|---|---|"]
    linhas += [f"| {n} | {e} | {o} | {'✔' if ok else '✖'} |" for n, e, o, ok in testes]
    linhas.append("")

    # ---------------- D ----------------
    import datetime as dt
    from collections import Counter as _C, defaultdict
    pref = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
    agenda = {e["NomeBD"]: {(i, s) for i, p in enumerate(pref) for s in "DN" if e[p + s] == "1"} for e in ensaios.values()}
    painel = [r for r in wb["Painel"].iter_rows(min_row=11, max_row=20, values_only=True) if r[1] and isinstance(r[5], (int, float))]
    ini_p = next(v for r in wb["Painel"].iter_rows(min_row=6, max_row=6, values_only=True) for v in r if isinstance(v, dt.datetime))
    ultimo = max(r[2] for r in bd).date()
    regs = [(r[0], r[2].date(), r[17]) for r in bd]
    sit = aderencia(agenda, regs, ini_p.date(), ultimo, dt.datetime.combine(ultimo + dt.timedelta(days=1), dt.time(8, 0)))
    cont = defaultdict(_C)
    for _, _, e, s in sit:
        cont[e][s] += 1
    nome_para_bd = {e["Nome"]: e["NomeBD"] for e in ensaios.values()}
    linhasD, okD = [], 0
    for r in painel:
        bdn = nome_para_bd.get(r[1])
        if not bdn:
            continue
        k = cont[bdn]
        prev = sum(v for s, v in k.items() if s not in ("Futuro", "Em andamento"))
        real = k["No turno"] + k["No dia"] + k["Com atraso"]
        ok = prev == r[5] and real == r[7]
        okD += ok
        linhasD.append(f"| {r[1]} | {r[5]} / {r[7]} | {prev} / {real} | {'✔' if ok else '✖'} |")
    linhas += ["## D. Aderência à agenda (aba Painel)", "", f"Período {ini_p:%d/%m} a {ultimo:%d/%m}. Conferem: **{okD} de {len(linhasD)}** ensaios", "",
               "| Ensaio | Planilha (previstos / realizados) | Regra do sistema | OK |", "|---|---|---|---|"] + linhasD + [""]
    print(f"D: {okD}/{len(linhasD)}")

    (pasta / "relatorio-validacao.md").write_text("\n".join(linhas), encoding="utf-8")
    print(f"A: {okA}/{len(bd) - pulados}  B: {okB}/{total_b}  C: {okC}/{len(testes)}")
    for d in difA[:10]:
        print("  A≠", d)
    for d in difB[:20]:
        print("  B≠", d)
    for t in testes:
        if not t[3]:
            print("  C≠", t)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    main(Path(sys.argv[1]), Path(sys.argv[2]))
