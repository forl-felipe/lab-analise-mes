"""Raio-X da base de TML + benchmark de modelos. Gera figuras e reports/metrics.json.

Uso:  python src/analysis.py   (rodar depois de src/etl.py)
"""
from __future__ import annotations

import json
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats
from sklearn.ensemble import GradientBoostingRegressor, RandomForestRegressor
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, ConstantKernel, DotProduct, WhiteKernel
from sklearn.linear_model import LinearRegression, RidgeCV
from sklearn.model_selection import LeaveOneGroupOut
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

from predictor import SPECS, TMLPredictor, load_training_frame
from tml_physics import tml_from_void_ratio

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "reports" / "figures"
FIG.mkdir(parents=True, exist_ok=True)

# ---- paleta (dataviz reference palette, modo claro) ----
SURF, INK, INK2, GRID = "#fcfcfb", "#0b0b0b", "#52514e", "#e4e3df"
S1, S2, S3 = "#2a78d6", "#eb6834", "#1baf7a"          # categóricas 1-3
CRIT, GOOD = "#d03b3b", "#0ca30c"                      # status
plt.rcParams.update({
    "figure.facecolor": SURF, "axes.facecolor": SURF, "savefig.facecolor": SURF,
    "axes.edgecolor": GRID, "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "axes.spines.top": False,
    "axes.spines.right": False, "font.size": 10, "axes.titlesize": 12, "axes.titleweight": "bold",
    "axes.titlecolor": INK, "axes.titlelocation": "left", "legend.frameon": False, "lines.linewidth": 2,
})
TYPE_COLORS = {"Pátio / produção": S1, "Embarque": S2}


def savefig(fig, name):
    fig.tight_layout()
    fig.savefig(FIG / name, dpi=160)
    plt.close(fig)


def grouped_cv(df, feats, make, target="tml"):
    df = df.dropna(subset=feats).reset_index(drop=True)
    pred = np.empty(len(df))
    for tr, te in LeaveOneGroupOut().split(df, groups=df["sample_id"]):
        m = make().fit(df.loc[tr, feats].values, df.loc[tr, target].values)
        p = m.predict(df.loc[te, feats].values)
        pred[te] = tml_from_void_ratio(p, df.loc[te, "rho_s"].values) if target == "e80" else p
    return df, pred


def main():
    d = load_training_frame()
    d["tipo"] = np.where(d["ship"] == 1, "Embarque", "Pátio / produção")
    full = pd.read_csv(ROOT / "data" / "processed" / "tml_dataset.csv", parse_dates=["date"])
    M = {}

    # ============ 1. Descritiva e repetibilidade ============
    M["n_ensaios"] = int(len(full)); M["n_amostras"] = int(full.sample_id.nunique())
    M["por_produto"] = full.groupby("product").tml.agg(["count", "mean", "std", "min", "max"]).round(3).to_dict("index")
    reps = full[full.n_tests_sample > 1].groupby("sample_id")
    pairs = [(g.material.iloc[-1], abs(g.tml.iloc[0] - g.tml.iloc[1])) for _, g in reps]
    intra = [dl for m, dl in pairs if "ofc" not in m]
    s_r = float(np.sqrt(np.mean(np.square(intra)) / 2))
    M["repetibilidade"] = {"pares": len(intra), "s_r_pp": round(s_r, 3), "r95_pp": round(2.77 * s_r, 3),
                           "azimuth_delta_pp": round([dl for m, dl in pairs if "ofc" in m][0], 2)}

    # ============ 2. Correlações (Spearman, PFN+PFSA) ============
    names = {"ppc": "PPC", "rho_s": "Densidade real", "goe": "Goethita", "hem_por": "Hematita porosa",
             "mag": "Magnetita", "hem_esp": "Hematita especular", "m325": "-325#", "se_blaine": "Sup. específica",
             "sio2": "SiO₂", "p": "P", "ship": "Amostra de embarque"}
    corr = []
    for c, lab in names.items():
        t = d[[c, "tml"]].dropna()
        if c in ("goe", "hem_por"):          # protocolo mineralógico só comparável dentro de 2026
            t = d[d.mineral_protocol == "M26 (4 fases)"][[c, "tml"]].dropna(); lab += " (2026)"
        if c in ("mag", "hem_esp"):
            lab += " (2026)"
        rho, p = stats.spearmanr(t[c], t["tml"])
        lo, hi = np.tanh(np.arctanh(rho) + np.array([-1, 1]) * 1.96 / np.sqrt(max(len(t) - 3, 1)))
        corr.append(dict(var=lab, rho=rho, lo=lo, hi=hi, p=p, n=len(t)))
    corr = pd.DataFrame(corr).sort_values("rho")
    M["correlacoes"] = corr.round(3).to_dict("records")
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    y = np.arange(len(corr))
    sig = corr.p < 0.05
    ax.hlines(y, corr.lo, corr.hi, color=INK2, lw=1.2)
    ax.scatter(corr.rho[sig], y[sig], s=60, color=S1, zorder=3, label="significativo (p < 0,05)", edgecolor=SURF, lw=2)
    ax.scatter(corr.rho[~sig], y[~sig], s=60, facecolor=SURF, edgecolor=S1, lw=1.6, zorder=3, label="não significativo")
    ax.axvline(0, color=INK2, lw=0.8)
    ax.set_yticks(y, [f"{v}  (n={n})" for v, n in zip(corr["var"], corr["n"])])
    ax.set_xlabel("Correlação de Spearman com o TML (IC 95 %)"); ax.set_xlim(-1, 1)
    ax.set_title("O que se move junto com o TML")
    ax.legend(loc="lower right", fontsize=8)
    savefig(fig, "01_correlacoes.png")

    # ============ 3. Efeito embarque × pátio ============
    fit = smf.ols("tml ~ ppc + ship", d).fit(cov_type="cluster", cov_kwds={"groups": d.sample_id})
    fit26 = smf.ols("tml ~ ppc + ship", d[d.year == 2026]).fit()
    fit25 = smf.ols("tml ~ ppc + ship", d[d.year == 2025]).fit()
    M["efeito_embarque"] = {
        "coef_pp": round(fit.params["ship"], 3), "ic95": [round(v, 3) for v in fit.conf_int().loc["ship"]],
        "p": float(fit.pvalues["ship"]), "coef_ppc": round(fit.params["ppc"], 3), "R2": round(fit.rsquared, 3),
        "coef_2025": round(fit25.params["ship"], 3), "p_2025": float(fit25.pvalues["ship"]),
        "coef_2026": round(fit26.params["ship"], 3), "p_2026": float(fit26.pvalues["ship"])}
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    xs = np.linspace(d.ppc.min() - .05, d.ppc.max() + .05, 50)
    for tipo, col in TYPE_COLORS.items():
        s = d[d.tipo == tipo]
        ax.scatter(s.ppc, s.tml, s=42, color=col, edgecolor=SURF, lw=1.5, label=f"{tipo} (n={len(s)})", zorder=3)
        ax.plot(xs, fit.params["Intercept"] + fit.params["ppc"] * xs + fit.params["ship"] * (tipo == "Embarque"), color=col)
    ax.annotate(f"Δ = {fit.params['ship']:+.2f} pp no mesmo PPC\n(p = {fit.pvalues['ship']:.0e})",
                xy=(4.15, fit.params['Intercept'] + fit.params['ppc'] * 4.15 - 0.22), fontsize=9, color=INK)
    ax.set_xlabel("PPC (%)"); ax.set_ylabel("TML (%)")
    ax.set_title("Amostras de embarque têm TML menor que as de pátio/produção")
    ax.legend(loc="upper left", fontsize=9)
    savefig(fig, "02_efeito_embarque.png")

    # ============ 4. Segurança: umidade de embarque × TML ============
    sh = full[full.moist_ship.notna() & full["product"].isin(["PFN", "PFSA", "PSC"])].copy()
    over = sh.moist_ship > sh.tml
    M["seguranca"] = {"embarques_com_umidade": int(len(sh)), "umidade_acima_TML": int(over.sum()),
                      "margem_mediana_pp": round(float((sh.tml - sh.moist_ship).median()), 3),
                      "margem_min_pp": round(float((sh.tml - sh.moist_ship).min()), 3)}
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    lim = [10.0, 13.0]
    ax.plot(lim, lim, color=INK2, lw=1, ls="--")
    ax.text(11.12, 11.38, "umidade = TML", color=INK2, fontsize=8, rotation=28)
    ax.scatter(sh.tml[~over], sh.moist_ship[~over], s=46, color=GOOD, edgecolor=SURF, lw=1.5,
               label=f"umidade < TML ({(~over).sum()} registros)", zorder=3)
    ax.scatter(sh.tml[over], sh.moist_ship[over], s=52, marker="^", color=CRIT, edgecolor=SURF, lw=1.5,
               label=f"umidade > TML ({over.sum()} registros)", zorder=3)
    ax.set_xlim(10.2, 12.8); ax.set_ylim(10.0, 11.5)
    ax.set_xlabel("TML do embarque (%)"); ax.set_ylabel("Umidade de embarque registrada (%)")
    ax.set_title("Umidade de embarque registrada vs. TML")
    ax.legend(loc="lower right", fontsize=9)
    savefig(fig, "03_umidade_vs_tml.png")

    # ============ 5. Linha do tempo ============
    fig, ax = plt.subplots(figsize=(9, 4.2))
    for tipo, col in TYPE_COLORS.items():
        s = d[d.tipo == tipo]
        ax.scatter(s.date, s.tml, s=36, color=col, edgecolor=SURF, lw=1.2, label=tipo, zorder=3)
    ax.set_ylabel("TML (%)"); ax.set_title("Histórico de TML – PFN/PFSA, 2025–2026")
    ax.legend(loc="upper left", fontsize=9)
    savefig(fig, "04_linha_do_tempo.png")

    # ============ 6. Equações atuais fora da amostra ============
    eqs = {"Hematita porosa\ny = −0,0644x + 13,986\n(R² Excel = 0,83)": ("hem_por", lambda v: -0.0644 * v + 13.986),
           "Goethita\ny = 0,0635x + 9,6894\n(R² Excel = 0,81)": ("goe", lambda v: 0.0635 * v + 9.6894),
           "PPC\ny = 0,8999x + 8,296\n(R² Excel = 0,45)": ("ppc", lambda v: 0.8999 * v + 8.296)}
    M["equacoes_atuais"] = {}
    fig, axs = plt.subplots(1, 3, figsize=(11, 4), sharey=True)
    m26 = d[d.mineral_protocol == "M26 (4 fases)"]
    for ax, (lab, (c, f)) in zip(axs, eqs.items()):
        s = m26.dropna(subset=[c])
        pr = f(s[c]); err = pr - s.tml
        r2 = 1 - np.sum(err ** 2) / np.sum((s.tml - s.tml.mean()) ** 2)
        M["equacoes_atuais"][c] = {"n": len(s), "rmse": round(float(np.sqrt(np.mean(err ** 2))), 3),
                                   "vies": round(float(err.mean()), 3), "R2_real": round(float(r2), 2)}
        ax.plot([10.3, 12.8], [10.3, 12.8], color=INK2, lw=1, ls="--")
        ax.scatter(s.tml, pr, s=36, color=S1, edgecolor=SURF, lw=1.2, zorder=3)
        ax.set_title(lab, fontsize=9, fontweight="normal")
        ax.text(0.04, 0.96, f"Todos os 23 ensaios 2026:\nR² real = {r2:.2f}\nRMSE = {np.sqrt(np.mean(err**2)):.2f} pp",
                transform=ax.transAxes, va="top", fontsize=8.5, color=INK)
        ax.set_xlabel("TML medido (%)")
    axs[0].set_ylabel("TML previsto pela equação (%)")
    fig.suptitle("Equações atuais aplicadas a todos os ensaios de 2026 (não só aos 10–11 filtrados)",
                 x=0.01, ha="left", fontweight="bold", fontsize=12)
    savefig(fig, "05_equacoes_atuais.png")

    # ============ 7. Benchmark de modelos (LOSO-CV) ============
    lin = lambda: LinearRegression()
    rid = lambda: make_pipeline(StandardScaler(), RidgeCV(alphas=np.logspace(-3, 3, 30)))
    gpr = lambda: make_pipeline(StandardScaler(), GaussianProcessRegressor(
        ConstantKernel() * RBF([1.0] * 4) + DotProduct() + WhiteKernel(.05), normalize_y=True, n_restarts_optimizer=2, random_state=0))
    rf = lambda: RandomForestRegressor(500, min_samples_leaf=3, random_state=0)
    gbm = lambda: GradientBoostingRegressor(n_estimators=200, max_depth=2, learning_rate=.05, subsample=.8, random_state=0)
    chem = ["ppc", "rho_s", "sio2", "p"]
    bench = [("Média histórica (sem modelo)", None, None, "tml"),
             ("Linear: PPC", ["ppc"], lin, "tml"),
             ("Linear: PPC + ρs + SiO₂ + P", chem, lin, "tml"),
             ("Random Forest: química", chem, rf, "tml"),
             ("Gradient Boosting: química", chem, gbm, "tml"),
             ("Processo Gaussiano: química", chem, gpr, "tml"),
             ("Ridge: química + embarque", chem + ["ship"], rid, "tml"),
             ("Híbrido físico v0_simples", SPECS["v0_simples"], lin, "e80"),
             ("Híbrido físico v0_completo", SPECS["v0_completo"], lin, "e80")]
    rows = []
    for nm, feats, mk, tgt in bench:
        if feats is None:
            pred = np.empty(len(d))
            for tr, te in LeaveOneGroupOut().split(d, groups=d.sample_id):
                pred[te] = d.tml.values[tr].mean()
            s = d
        else:
            s, pred = grouped_cv(d, feats, mk, tgt)
        e = pred - s.tml.values
        rows.append(dict(modelo=nm, rmse=np.sqrt(np.mean(e ** 2)), mae=np.mean(np.abs(e)),
                         dentro_03=np.mean(np.abs(e) <= 0.3)))
        if nm == "Híbrido físico v0_completo":
            best_s, best_pred = s, pred
    bench = pd.DataFrame(rows)
    M["benchmark"] = bench.round(3).to_dict("records")
    fig, ax = plt.subplots(figsize=(8, 4.6))
    b = bench.iloc[::-1]
    cols = [S2 if "Híbrido" in m else (INK2 if "Média" in m else S1) for m in b.modelo]
    ax.barh(b.modelo, b.rmse, color=cols, height=0.6)
    for yy, v in enumerate(b.rmse):
        ax.text(v + 0.005, yy, f"{v:.3f}", va="center", fontsize=8.5, color=INK)
    ax.axvline(s_r, color=GOOD, lw=1.5, ls=":")
    ax.text(s_r + 0.005, -0.45, f"repetibilidade do lab (s_r = {s_r:.2f} pp)", fontsize=8, color=INK2, va="center",
            bbox=dict(facecolor=SURF, edgecolor="none", pad=1))
    ax.set_xlabel("RMSE na validação cruzada por amostra (pp de umidade) – menor é melhor")
    ax.set_title("Erro fora da amostra: modelos candidatos")
    ax.grid(axis="y", visible=False)
    savefig(fig, "06_benchmark_modelos.png")

    # ============ 8. Observado × previsto (modelo v0) ============
    m = TMLPredictor(features=SPECS["v0_completo"], name="v0_completo").fit(d)
    r = np.asarray(m.cv_residuals_)
    q05 = np.quantile(r, 0.05)
    fig, ax = plt.subplots(figsize=(6.2, 5.4))
    ax.plot([10.3, 12.8], [10.3, 12.8], color=INK2, lw=1, ls="--")
    for tipo, col in TYPE_COLORS.items():
        k = (best_s.ship == (tipo == "Embarque")).values
        ax.scatter(best_s.tml[k], best_pred[k], s=42, color=col, edgecolor=SURF, lw=1.3, label=tipo, zorder=3)
    ax.set_xlabel("TML medido (%)"); ax.set_ylabel("TML previsto – validação cruzada (%)")
    ax.set_title("Modelo híbrido v0: previsto × medido")
    ax.text(0.04, 0.96, f"RMSE = {np.sqrt(np.mean(r**2)):.2f} pp\nMAE = {np.mean(np.abs(r)):.2f} pp\n"
            f"limite inferior 95 % = previsto {q05:+.2f} pp", transform=ax.transAxes, va="top", fontsize=9)
    ax.legend(loc="lower right", fontsize=9)
    savefig(fig, "07_v0_previsto_vs_medido.png")
    M["modelo_v0"] = {"equacao": m.equation(), **{k: round(v, 3) if isinstance(v, float) else v for k, v in m.cv_metrics().items()},
                      "lim_inf_95_offset_pp": round(float(q05), 3)}

    # ============ 9. Física: índice de vazios ============
    fig, ax = plt.subplots(figsize=(7.5, 4.6))
    for tipo, col in TYPE_COLORS.items():
        s = d[d.tipo == tipo]
        ax.scatter(s.ppc, s.e80, s=40, color=col, edgecolor=SURF, lw=1.2, label=tipo, zorder=3)
    ax.set_xlabel("PPC (%)"); ax.set_ylabel("Índice de vazios no TML, e80 (–)")
    ax.set_title("A grandeza física por trás do TML: o empacotamento (e80)")
    ax.legend(fontsize=9, loc="upper left")
    savefig(fig, "08_indice_vazios.png")
    M["fisica"] = {"e80_media": round(float(d.e80.mean()), 4), "e80_cv": round(float(d.e80.std() / d.e80.mean()), 4),
                   "porosidade_media": round(float((d.e80 / (1 + d.e80)).mean()), 4),
                   "rmse_e_constante": round(float(np.sqrt(np.mean((tml_from_void_ratio(d.e80.mean(), d.rho_s) - d.tml) ** 2))), 3)}

    (ROOT / "reports" / "metrics.json").write_text(json.dumps(M, indent=2, ensure_ascii=False, default=float))
    print(json.dumps(M, indent=1, ensure_ascii=False, default=float))


if __name__ == "__main__":
    main()
