"""Física do ensaio Proctor/Fagerberg (PFD80) usada no modelo híbrido.

No PFD80 (IMSBC Code, App. 2 – método modificado para finos de minério de ferro) o TML
é a umidade em que a curva de compactação cruza a linha de grau de saturação Sr = 80 %.
Num ponto qualquer de saturação Sr:

    w (base seca)      = Sr · e · ρw / ρs
    TML (base úmida %) = 100 · w / (1 + w)

onde e = índice de vazios do material compactado e ρs = densidade real dos sólidos.
Invertendo, cada ensaio histórico revela o índice de vazios "efetivo" e80 da amostra,
uma grandeza física de empacotamento — é ela que o modelo aprende.
"""
import numpy as np

SR_PFD80 = 0.80
RHO_W = 1.0  # g/cm³


def void_ratio_from_tml(tml_pct, rho_s, sr=SR_PFD80):
    """Índice de vazios no TML (e80) a partir do TML (% base úmida) e ρs (g/cm³)."""
    tml = np.asarray(tml_pct, float) / 100.0
    w = tml / (1.0 - tml)
    return w * np.asarray(rho_s, float) / (sr * RHO_W)


def tml_from_void_ratio(e, rho_s, sr=SR_PFD80):
    """TML (% base úmida) a partir do índice de vazios e da densidade real."""
    w = sr * np.asarray(e, float) * RHO_W / np.asarray(rho_s, float)
    return 100.0 * w / (1.0 + w)


def porosity(e):
    e = np.asarray(e, float)
    return e / (1.0 + e)
