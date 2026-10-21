"""Lê a planilha de disponibilidade dos filtros prensa e injeta os eventos no dashboard HTML.

Uso:
    python scripts/atualizar_dados_filtros.py caminho/da/planilha.xlsx

Colunas esperadas (linha 1): Equipamento, Início, Término, Horas Paradas,
Especialidade, Motivo, Descrição.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

HTML = Path(__file__).resolve().parent.parent / "dashboards" / "filtros-prensa" / "index.html"
INICIO, FIM = "/*DADOS_INICIO*/", "/*DADOS_FIM*/"


def ler_eventos(xlsx):
    ws = openpyxl.load_workbook(xlsx, data_only=True).active
    eventos = []
    for eq, ini, fim, horas, esp, motivo, desc in ws.iter_rows(min_row=2, max_col=7, values_only=True):
        if not eq or not ini or not fim:
            continue
        eventos.append({
            "eq": str(eq).strip(),
            "ini": ini.strftime("%Y-%m-%dT%H:%M"),
            "fim": fim.strftime("%Y-%m-%dT%H:%M"),
            "h": round(float(horas or 0), 2),
            "esp": str(esp or "").strip().upper(),
            "mot": str(motivo or "").strip(),
            "desc": re.sub(r"\s+", " ", str(desc or "")).strip(),
        })
    return eventos


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    dados = json.dumps(ler_eventos(sys.argv[1]), ensure_ascii=False, separators=(",", ":"))
    html = HTML.read_text(encoding="utf-8")
    a, b = html.index(INICIO) + len(INICIO), html.index(FIM)
    HTML.write_text(html[:a] + dados + html[b:], encoding="utf-8")
    print(f"{len(json.loads(dados))} eventos gravados em {HTML}")


if __name__ == "__main__":
    main()
