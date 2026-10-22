"""Lê a planilha de paradas dos filtros prensa e embute os eventos no dashboard HTML.

Uso:
    python scripts/atualizar_dados_filtros.py caminho/da/planilha.xlsx

O cabeçalho é localizado pelo nome das colunas (pode começar em qualquer linha/coluna).
Colunas usadas: Equipamento, Início, Término, Horas Paradas, Especialidade, Motivo, Descrição.
As demais colunas são ignoradas. O mesmo processo existe na aba "Importar dados" do HTML.
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

HTML = Path(__file__).resolve().parent.parent / "dashboards" / "filtros-prensa" / "index.html"
INICIO, FIM = "/*DADOS_INICIO*/", "/*DADOS_FIM*/"
COLUNAS = {
    "eq": ["equipamento"], "ini": ["inicio"], "fim": ["termino", "fim"], "h": ["horas paradas"],
    "esp": ["especialidade"], "mot": ["motivo"], "desc": ["descricao"],
}


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "").lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", s)).strip()


def ler_eventos(xlsx):
    rows = list(openpyxl.load_workbook(xlsx, data_only=True).worksheets[0].iter_rows(values_only=True))
    for n, row in enumerate(rows[:30]):
        cab = [norm(c) for c in row]
        idx = {k: next((i for i, c in enumerate(cab) if c in nomes), None) for k, nomes in COLUNAS.items()}
        if None not in (idx["eq"], idx["ini"], idx["fim"]):
            break
    else:
        sys.exit("Cabeçalho não encontrado (Equipamento, Início, Término).")

    def val(r, k):
        return r[idx[k]] if idx[k] is not None and idx[k] < len(r) else None

    eventos = []
    for r in rows[n + 1:]:
        eq, ini, fim = val(r, "eq"), val(r, "ini"), val(r, "fim")
        if not eq or not hasattr(ini, "strftime") or not hasattr(fim, "strftime"):
            continue
        h = val(r, "h")
        eventos.append({
            "eq": str(eq).strip(),
            "ini": ini.strftime("%Y-%m-%dT%H:%M"),
            "fim": fim.strftime("%Y-%m-%dT%H:%M"),
            "h": round(float(h) if h is not None else (fim - ini).total_seconds() / 3600, 2),
            "esp": str(val(r, "esp") or "").strip().upper() or "N/I",
            "mot": str(val(r, "mot") or "").strip() or "Não informado",
            "desc": re.sub(r"\s+", " ", str(val(r, "desc") or "")).strip(),
        })
    return sorted(eventos, key=lambda e: (e["eq"], e["ini"]))


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    eventos = ler_eventos(sys.argv[1])
    dados = json.dumps(eventos, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    html = HTML.read_text(encoding="utf-8")
    a, b = html.index(INICIO) + len(INICIO), html.index(FIM)
    HTML.write_text(html[:a] + dados + html[b:], encoding="utf-8")
    print(f"{len(eventos)} eventos gravados em {HTML}")


if __name__ == "__main__":
    main()
