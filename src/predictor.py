"""Preditor de TML v0 – modelo híbrido física + estatística (PFD80).

Estrutura
---------
1. Camada física (tml_physics): TML = f(e80, ρs) com Sr = 80 %.
2. Camada estatística: o índice de vazios e80 é previsto por regressão linear sobre
   variáveis medidas rotineiramente (PPC, ρs, SiO2, P) + o tipo de amostra
   (embarque vs. pátio/produção).
3. Incerteza: resíduos de validação cruzada leave-one-sample-out (réplicas ficam juntas),
   usados como no jackknife+ para dar um limite inferior unilateral (lado seguro).

Uso
---
    python src/predictor.py train                 # treina e salva models/tml_v0.json
    python src/predictor.py predict --ppc 3.6 --rho 4.69 --sio2 2.1 --p 0.061 --embarque 1
"""
from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import LeaveOneGroupOut

from tml_physics import tml_from_void_ratio, void_ratio_from_tml

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "processed" / "tml_dataset.csv"
MODEL_DIR = ROOT / "models"

SPECS = {
    # nome: features   (todas medidas na rotina do laboratório)
    "v0_simples": ["ppc", "ship"],
    "v0_completo": ["ppc", "rho_s", "sio2", "p", "ship"],
}


def load_training_frame(path=DATA) -> pd.DataFrame:
    d = pd.read_csv(path, parse_dates=["date"])
    d = d[d["product"].isin(["PFN", "PFSA"])].copy()   # PSC tem só 3 ensaios: fora do domínio
    d["ship"] = d["is_shipment"].astype(float)
    d["e80"] = void_ratio_from_tml(d["tml"], d["rho_s"])
    return d.reset_index(drop=True)


@dataclass
class TMLPredictor:
    features: list[str]
    name: str = "v0"
    coef_: dict = field(default_factory=dict)
    intercept_: float = 0.0
    cv_residuals_: list = field(default_factory=list)
    domain_: dict = field(default_factory=dict)
    n_train_: int = 0

    # ---------------- treino ----------------
    def _fit_core(self, X, e):
        m = LinearRegression().fit(X, e)
        return m.intercept_, m.coef_

    def fit(self, df: pd.DataFrame) -> "TMLPredictor":
        df = df.dropna(subset=self.features + ["e80", "tml"]).reset_index(drop=True)
        X, e = df[self.features].values, df["e80"].values
        b0, b = self._fit_core(X, e)
        self.intercept_, self.coef_ = float(b0), dict(zip(self.features, map(float, b)))
        # resíduos de CV agrupada (réplicas da mesma amostra nunca se separam)
        res = np.empty(len(df))
        for tr, te in LeaveOneGroupOut().split(X, groups=df["sample_id"]):
            c0, c = self._fit_core(X[tr], e[tr])
            e_hat = c0 + X[te] @ c
            res[te] = df["tml"].values[te] - tml_from_void_ratio(e_hat, df["rho_s"].values[te])
        self.cv_residuals_ = res.tolist()
        self.domain_ = {f: [float(df[f].min()), float(df[f].max())] for f in self.features + ["rho_s"]}
        self.n_train_ = len(df)
        return self

    # ---------------- previsão ----------------
    def predict_void_ratio(self, df: pd.DataFrame):
        return self.intercept_ + sum(self.coef_[f] * df[f].values for f in self.features)

    def predict(self, df: pd.DataFrame, conf: float = 0.95) -> pd.DataFrame:
        e_hat = self.predict_void_ratio(df)
        tml_hat = tml_from_void_ratio(e_hat, df["rho_s"].values)
        r = np.asarray(self.cv_residuals_)
        lo = tml_hat + np.quantile(r, 1 - conf)              # limite inferior unilateral
        lo2, hi2 = tml_hat + np.quantile(r, (1 - conf) / 2), tml_hat + np.quantile(r, 1 - (1 - conf) / 2)
        out = pd.DataFrame({"e80_prev": e_hat, "tml_prev": tml_hat,
                            f"tml_lim_inf_{int(conf*100)}": lo,
                            "pi_inf": lo2, "pi_sup": hi2})
        out["fora_do_dominio"] = [
            ", ".join(f for f in self.domain_ if f in df and not (self.domain_[f][0] <= row[f] <= self.domain_[f][1]))
            for _, row in df.iterrows()]
        return out

    def cv_metrics(self) -> dict:
        r = np.asarray(self.cv_residuals_)
        return {"n": self.n_train_, "rmse": float(np.sqrt(np.mean(r ** 2))), "mae": float(np.mean(np.abs(r))),
                "vies": float(np.mean(r)), "p95_abs": float(np.quantile(np.abs(r), 0.95)),
                "frac_|erro|<=0.3pp": float(np.mean(np.abs(r) <= 0.3))}

    def equation(self) -> str:
        terms = " ".join(f"{v:+.4f}·{k}" for k, v in self.coef_.items())
        return (f"e80 = {self.intercept_:.4f} {terms}\n"
                f"w = 0.80·e80/ρs ;  TML(%) = 100·w/(1+w)")

    # ---------------- persistência ----------------
    def save(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = {k: getattr(self, k) for k in
                   ["name", "features", "coef_", "intercept_", "cv_residuals_", "domain_", "n_train_"]}
        payload["cv_metrics"] = self.cv_metrics()
        payload["equation"] = self.equation()
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False))

    @classmethod
    def load(cls, path: Path) -> "TMLPredictor":
        p = json.loads(Path(path).read_text())
        p.pop("cv_metrics", None), p.pop("equation", None)
        return cls(**p)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("train")
    pr = sub.add_parser("predict")
    pr.add_argument("--modelo", default="v0_completo", choices=list(SPECS))
    pr.add_argument("--ppc", type=float, required=True)
    pr.add_argument("--rho", type=float, required=True, help="densidade real (g/cm³)")
    pr.add_argument("--sio2", type=float, default=np.nan)
    pr.add_argument("--p", type=float, default=np.nan)
    pr.add_argument("--embarque", type=int, default=1, help="1 = amostra de embarque; 0 = pátio/produção")
    a = ap.parse_args()

    if a.cmd == "train":
        df = load_training_frame()
        for name, feats in SPECS.items():
            m = TMLPredictor(features=feats, name=name).fit(df)
            m.save(MODEL_DIR / f"tml_{name}.json")
            print(f"[{name}]  {m.cv_metrics()}\n{m.equation()}\n")
    else:
        m = TMLPredictor.load(MODEL_DIR / f"tml_{a.modelo}.json")
        x = pd.DataFrame([{"ppc": a.ppc, "rho_s": a.rho, "sio2": a.sio2, "p": a.p, "ship": float(a.embarque)}])
        r = m.predict(x).iloc[0]
        print(f"TML previsto            : {r.tml_prev:.2f} %")
        print(f"Limite inferior (95 %)  : {r.tml_lim_inf_95:.2f} %   <- usar para decisão")
        print(f"Intervalo 95 %          : {r.pi_inf:.2f} – {r.pi_sup:.2f} %")
        if r.fora_do_dominio:
            print(f"ATENÇÃO: fora da faixa de treino em: {r.fora_do_dominio}")


if __name__ == "__main__":
    main()
