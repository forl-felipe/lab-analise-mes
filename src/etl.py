"""ETL – consolida os históricos de TML 2025/2026 numa base única e auditável.

Fonte primária: abas "Historico PFN ..." e "Histórico PSC" (registro de ensaios).
Abas derivadas ("Regressão ...", "Confirmações", "Dados preditor") são usadas para:
  * trazer a Superfície Específica (SE, Blaine cm²/g), que só existe nas abas de regressão;
  * acrescentar as réplicas T02 da aba "Confirmações";
  * auditar inconsistências de transcrição (ver data_quality_report).

Uso:  python src/etl.py
Saída: data/processed/tml_dataset.csv  e  data/processed/data_quality.csv
"""
from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"

F25 = RAW / "Historico_TML_2025.xlsx"
F26 = RAW / "Historico_TML_2026.xlsx"

COLMAP = {
    "Material": "material", "PFUF": "pfuf", "PM": "pm", "PFSA": "pfsa_col", "Data": "date",
    "SiO2": "sio2", "P": "p", "PPC": "ppc", "Densidade (g/cm³)": "rho_s", "-325#": "m325",
    "Hematita Especular": "hem_esp", "Hematita Porosa": "hem_por", "Magnetita": "mag",
    "Goethita": "goe", "TML %": "tml", "Umidade Embarque": "moist_ship", "SE": "se_blaine",
}
NUM = ["pfuf", "pm", "pfsa_col", "sio2", "p", "ppc", "rho_s", "m325", "hem_esp", "hem_por",
       "mag", "goe", "tml", "moist_ship", "se_blaine"]


def to_num(v):
    """Converte textos como '3.38', '18,77', '0.065' em float."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def read_block(path, sheet, header_row, first_col="B"):
    """Lê uma tabela cujo cabeçalho está em `header_row` (1-based)."""
    df = pd.read_excel(path, sheet_name=sheet, header=header_row - 1)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    df["_excel_row"] = df.index + header_row + 1
    return df


def standardise(df, source):
    rename = {c: COLMAP[c] for c in df.columns if c in COLMAP}
    df = df.rename(columns=rename)
    keep = [c for c in list(COLMAP.values()) + ["_excel_row"] if c in df.columns]
    df = df[keep].copy()
    for c in NUM:
        if c in df.columns:
            df[c] = df[c].map(to_num)
    df["source"] = source
    return df


def product_of(name: str, sio2: float) -> str:
    n = str(name).upper()
    if "PSC" in n:
        return "PSC"
    if re.search(r"PF\s*/?\s*S\s*/?\s*A", n) or "PFSA" in n:
        return "PFSA"
    return "PFN"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    frames = []

    h25 = standardise(read_block(F25, "Historico PFN 2025 x Gráfico", 4), "2025:Historico PFN")
    h26 = standardise(read_block(F26, "Historico PFN com gráfico", 4), "2026:Historico PFN")
    p25 = standardise(read_block(F25, "Histórico PSC", 4), "2025:Historico PSC")
    p26 = standardise(read_block(F26, "Histórico PSC", 4), "2026:Historico PSC")
    conf = read_block(F25, "Confirmações", 4)
    conf = conf.rename(columns={"Produto Embarque": "Material"})
    conf = standardise(conf, "2025:Confirmações")
    frames = [h25, h26, p25, p26]

    base = pd.concat(frames, ignore_index=True)
    # remove linhas de resumo (ex.: "MÉDIA" da aba PSC) – só ensaios com material e data
    base = base[base["tml"].notna() & base["material"].notna() & base["date"].notna()].copy()

    # Réplicas da aba Confirmações: só as T02 (as T01 já estão no histórico)
    conf_t02 = conf[conf["material"].astype(str).str.contains("T02")].copy()
    base = pd.concat([base, conf_t02], ignore_index=True)

    # ---- Superfície específica (SE) vinda das abas de regressão 2026 --------
    se_tabs = []
    for sh in ["Regressão Embarques", "Regressão Filtrado"]:
        r = read_block(F26, sh, 2)
        r = r.rename(columns={"HISTÓRICO TML 2026": "Material"})
        se_tabs.append(standardise(r, f"2026:{sh}"))
    se = pd.concat(se_tabs, ignore_index=True)
    key = ["ppc", "rho_s", "tml"]
    se_k = se.dropna(subset=["se_blaine"]).groupby(key)["se_blaine"].first()
    base = base.merge(se_k.rename("se_blaine").reset_index(), on=key, how="left")

    base["material"] = base["material"].astype(str).str.strip()
    base["product"] = [product_of(m, s) for m, s in zip(base["material"], base["sio2"])]
    base["year"] = base["source"].str[:4].astype(int)
    base["date"] = pd.to_datetime(base["date"], errors="coerce")
    base["is_shipment"] = base["material"].str.contains("Navio|Embarcado|Emb\\.", case=False)
    base["has_desiccant"] = base["material"].str.contains("Desidratante", case=False) & \
        ~base["material"].str.contains("Sem Desidratante", case=False)
    base["is_replicate_label"] = base["material"].str.contains(r"T\s*-?\s*0?2|T02", regex=True)

    # Mineralogia: 2025 e 2026 parecem vir de protocolos diferentes (ver relatório)
    base["mineral_protocol"] = np.where(base["hem_esp"].notna(), "M26 (4 fases)",
                                np.where(base["hem_por"].notna(), "M25 (2 fases)", "sem"))

    # ---- Identificador de amostra física (réplicas compartilham química) ----
    chem_key = base[["product", "sio2", "p", "ppc", "rho_s"]].round(4).astype(str).agg("|".join, axis=1)
    base["sample_id"] = pd.factorize(chem_key)[0]
    base["se_blaine"] = base.groupby("sample_id")["se_blaine"].transform(lambda x: x.ffill().bfill())
    base["n_tests_sample"] = base.groupby("sample_id")["tml"].transform("size")

    # ---- Margem de segurança de embarque (TML - umidade) -------------------
    base["margin_pp"] = base["tml"] - base["moist_ship"]
    base["moist_over_tml"] = base["moist_ship"] > base["tml"]

    base = base.sort_values(["date", "source"]).reset_index(drop=True)
    cols = ["sample_id", "source", "_excel_row", "date", "year", "product", "material", "is_shipment",
            "has_desiccant", "sio2", "p", "ppc", "rho_s", "m325", "se_blaine", "hem_esp", "hem_por",
            "mag", "goe", "mineral_protocol", "pfuf", "pm", "pfsa_col", "tml", "moist_ship",
            "margin_pp", "moist_over_tml", "n_tests_sample"]
    base = base[cols]
    base.to_csv(OUT / "tml_dataset.csv", index=False)

    dq = data_quality_report(base)
    dq.to_csv(OUT / "data_quality.csv", index=False)
    print(base.groupby(["year", "product"]).size())
    print(f"\nEnsaios: {len(base)} | amostras físicas: {base.sample_id.nunique()}")
    print(f"Problemas de qualidade registrados: {len(dq)}")


def data_quality_report(base):
    """Auditoria das abas derivadas contra o histórico primário."""
    issues = []

    def add(sheet, cell, problem, detail):
        issues.append(dict(planilha=sheet, celula=cell, problema=problem, detalhe=detail))

    # 1) números gravados como texto
    import openpyxl
    for f, tag in [(F25, "2025"), (F26, "2026")]:
        wb = openpyxl.load_workbook(f)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and re.fullmatch(r"\s*-?\d+[.,]\d+\s*", c.value):
                        add(f"{tag}:{ws.title}", c.coordinate, "número armazenado como texto",
                            f"'{c.value}' – fórmulas/gráficos do Excel ignoram esta célula")
            # fórmulas de média com intervalos que não cobrem os dados
            if ws.title == "Dados preditor":
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("=AVERAGE"):
                            rng = re.findall(r"\(([A-Z])(\d+):[A-Z](\d+)\)", c.value)
                            if rng:
                                col, a, b = rng[0]
                                filled = [r for r in range(max(1, int(a) - 5), int(c.row))
                                          if ws[f"{col}{r}"].value is not None
                                          and not isinstance(ws[f"{col}{r}"].value, str)]
                                outside = [r for r in filled if not (int(a) <= r <= int(b))]
                                if outside:
                                    add(f"{tag}:{ws.title}", c.coordinate,
                                        "média com intervalo incompleto",
                                        f"{c.value} deixa de fora as linhas {outside}")
                for r in range(4, 58):
                    i, j = ws[f"I{r}"].value, ws[f"J{r}"].value
                    if i is not None or j is not None:
                        add(f"{tag}:{ws.title}", f"I{r}:J{r}", "valor fora da coluna",
                            f"TML={i} e PPC={j} digitados nas colunas I/J – não entram nas médias")
            if ws.title.startswith("Historico PFN"):
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and re.fullmatch(r"=F\d+\+(\d+)", c.value):
                            d = int(re.findall(r"\+(\d+)", c.value)[0])
                            if d != 179:
                                add(f"{tag}:{ws.title}", c.coordinate, "validade divergente",
                                    f"{c.value} (padrão da planilha é +179 dias)")
    # 2) Dados preditor x histórico (2025)
    dp = pd.read_excel(F25, sheet_name="Dados preditor", header=None)
    hist = base[(base.year == 2025) & base.source.str.contains("Historico")]
    for r in range(dp.shape[0]):
        d = dp.iat[r, 1]
        if isinstance(d, pd.Timestamp) or hasattr(d, "year"):
            sio2, ppc, tml = to_num(dp.iat[r, 2]), to_num(dp.iat[r, 3]), to_num(dp.iat[r, 7])
            cand = hist[hist.date == pd.Timestamp(d)]
            if len(cand) and not np.isnan(tml):
                m = cand.iloc[(cand.tml - tml).abs().argsort()[:1]].iloc[0]
                for nm, v, ref in [("SiO2", sio2, m.sio2), ("PPC", ppc, m.ppc)]:
                    if not np.isnan(v) and not np.isnan(ref) and abs(v - ref) > 0.02:
                        add("2025:Dados preditor", f"linha {r + 1}", f"{nm} diverge do histórico",
                            f"{d:%d/%m/%Y}: preditor={v} x histórico={ref}")
            por = to_num(dp.iat[r, 4])
            if not np.isnan(por) and not np.isnan(tml) and abs(por - tml) < 1e-9:
                add("2025:Dados preditor", f"E{r + 1}", "valor na coluna errada",
                    f"Porosidade = {por} é o próprio TML")
    # 3) ordem cronológica e embarques acima do TML
    for _, r in base[base.moist_over_tml].iterrows():
        add(r.source, f"linha {r._excel_row}", "umidade de embarque > TML",
            f"{r.material}: umidade {r.moist_ship:.2f}% x TML {r.tml:.2f}%")
    # 4) mesma química, TML muito diferente
    g = base.groupby("sample_id")["tml"].agg(["min", "max", "size"])
    for sid, row in g[(g["size"] > 1) & (g["max"] - g["min"] > 0.2)].iterrows():
        mats = "; ".join(base.loc[base.sample_id == sid, "material"].tolist())
        add("ambas", f"amostra {sid}", "réplicas discrepantes",
            f"ΔTML={row['max'] - row['min']:.2f} pp entre: {mats}")
    return pd.DataFrame(issues)


if __name__ == "__main__":
    main()
