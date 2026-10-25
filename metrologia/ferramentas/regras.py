"""
Regras de cálculo do sistema Metrologia, idênticas às do app (Power Fx) e às da
planilha de aferição. Servem de referência e de teste automático.

Regra geral de aceitação (mesma da planilha, aba BD_Limites):
  - Se o critério tem limite inferior e/ou superior:
        Conforme quando (LimInferior vazio ou valor >= LimInferior)
                    e   (LimSuperior vazio ou valor <= LimSuperior)
  - Senão, se tem tolerância:
        Conforme quando |valor - referência| <= tolerância   (arredondado a 6 casas)
  - Senão: Informativo

A referência é o Nominal do critério, ou o segundo valor informado (comparativos),
ou a umidade da estufa (analisadores).
"""

from dataclasses import dataclass


@dataclass
class Limite:
    nominal: float | None = None
    li: float | None = None
    ls: float | None = None
    tol: float | None = None


def avaliar(valor, referencia, lim: Limite | None):
    """Devolve (diferenca, resultado). resultado: 'Conforme', 'Não conforme', 'Informativo' ou ''."""
    if valor is None:
        return None, ""
    dif = None if referencia is None else valor - referencia
    if lim is None:
        return dif, "Informativo"
    if lim.li is not None or lim.ls is not None:
        ok = (lim.li is None or valor >= lim.li) and (lim.ls is None or valor <= lim.ls)
        return dif, "Conforme" if ok else "Não conforme"
    if lim.tol is not None:
        if dif is None:
            return None, ""
        return dif, "Conforme" if round(abs(dif), 6) <= round(lim.tol, 6) else "Não conforme"
    return dif, "Informativo"


def rpm(voltas, minutos, segundos):
    t = (minutos or 0) * 60 + (segundos or 0)
    return None if t <= 0 or not voltas else voltas * 60 / t


def umidade_estufa(tara, p_inicial, p_final):
    if None in (tara, p_inicial, p_final) or p_inicial <= tara:
        return None
    return (p_inicial - p_final) / (p_inicial - tara) * 100


def checklist(itens: list[str]):
    """itens: 'OK', 'NÃO OK', 'SEM TAG', 'N/A'. Devolve (itens_ok, respondidos, resultado)."""
    resp = [i for i in itens if i]
    ok = sum(1 for i in resp if i == "OK")
    if not resp:
        return 0, 0, ""
    if any(i == "NÃO OK" for i in resp):
        return ok, len(resp), "Não conforme"
    return ok, len(resp), "Conforme"


def granulometria(massas_a: list[float], massas_b: list[float], diametros: list[float]):
    """10 frações na ordem +19,0 ... -0,5. Devolve dict indicador -> (A, B)."""
    def pct(m):
        tot = sum(x or 0 for x in m)
        return [((x or 0) / tot * 100) if tot else None for x in m], tot

    pa, ta = pct(massas_a)
    pb, tb = pct(massas_b)

    def ind(p, tot):
        if not tot:
            return {}
        return {
            "-6,3 mm": p[7] + p[8] + p[9],
            "RG (Relação)": (p[2] + p[3]) / p[4] if p[4] else None,
            "-16,0 +12,5 mm": p[2] + p[3],
            "-16,0 +9,0 mm": p[2] + p[3] + p[4],
            "-16,0 +8,0 mm": p[2] + p[3] + p[4] + p[5],
            "Diâmetro médio": sum(x * d for x, d in zip(p, diametros)) / 100,
        }

    ia, ib = ind(pa, ta), ind(pb, tb)
    return {k: (ia.get(k), ib.get(k)) for k in ia}


def tamboramento(m63_a, m05_a, massa_a, m63_b, m05_b, massa_b):
    """Devolve {'+6,3 mm': (A, B), '-0,5 mm': (A, B)} em %."""
    def lado(m63, m05, massa):
        if m63 is None or m05 is None:
            return None, None
        p63, p05 = m63 / (massa / 100), m05 / (massa / 100)
        return p63, 100 - (p63 + p05)

    a63, a05 = lado(m63_a, m05_a, massa_a)
    b63, b05 = lado(m63_b, m05_b, massa_b)
    return {"+6,3 mm": (a63, b63), "-0,5 mm": (a05, b05)}


def letra_escala(data, turno, data_ref, letras07, letras19):
    pos = (data - data_ref).days % len(letras07)
    return letras07[pos] if turno == "07x19" else letras19[pos]


def status_equipamento(e: dict, hoje, dias_aviso=30, ultima_afericao=None):
    """Mesma regra do app, do fluxo diário e do Power BI.

    e: ControlaCalibracao, ProximaCalibracao, UltimaCalibracao, UltimoResultado, ManutencaoSolicitada
    ultima_afericao: (data, resultado) da aferição mais recente do equipamento
    """
    if e.get("ControlaCalibracao"):
        prox = e.get("ProximaCalibracao")
        if e.get("UltimoResultado") == "Recusada":
            return "Bloqueado · Calibração recusada"
        if e.get("ManutencaoSolicitada"):
            return "Bloqueado · Em manutenção"
        if prox and prox < hoje:
            return "Bloqueado · Calibração vencida"
    if ultima_afericao and ultima_afericao[1] == "Não conforme":
        ult_cal = e.get("UltimaCalibracao")
        if not (ult_cal and e.get("UltimoResultado") == "Aprovada" and ult_cal >= ultima_afericao[0]):
            return "Bloqueado · Aferição não conforme"
    if e.get("ControlaCalibracao"):
        prox = e.get("ProximaCalibracao")
        if not prox:
            return "Liberado · Sem data de calibração"
        if (prox - hoje).days <= dias_aviso:
            return "Liberado · Calibração vence em breve"
    return "Liberado"


def aderencia(agenda: dict, registros: list, inicio, fim, agora):
    """Mesma regra da macro ModPainel da planilha (aba Painel e Gestão por Letras).

    agenda: {ensaio: set de (dia_semana 0=seg..6=dom, 'D'|'N')}
    registros: [(ensaio, data, turno)] (um por resultado ou por lançamento; tanto faz)
    Devolve [(data, turno, ensaio, situacao)] com situacao em
    'No turno', 'No dia', 'Com atraso', 'Não feito', 'Em andamento', 'Futuro'.
    """
    import datetime as _dt

    def slot(d, t):
        return d.toordinal() * 2 + (1 if t == "19x07" else 0)

    def previsto(ens, s):
        d = _dt.date.fromordinal(s // 2)
        return (d.weekday(), "N" if s % 2 else "D") in agenda.get(ens, set())

    regs = [(e, slot(d, t), d) for e, d, t in registros]
    out = []
    d = inicio
    while d <= fim:
        for tt, t in ((0, "07x19"), (1, "19x07")):
            s = slot(d, t)
            for ens in agenda:
                if not previsto(ens, s):
                    continue
                if any(e == ens and r == s for e, r, _ in regs):
                    sit = "No turno"
                elif not previsto(ens, d.toordinal() * 2 + (1 - tt)) and any(e == ens and dd == d for e, _, dd in regs):
                    sit = "No dia"
                else:
                    prox = next((x for x in range(s + 1, s + 122) if previsto(ens, x)), s + 121)
                    if any(e == ens and s < r < prox for e, r, _ in regs):
                        sit = "Com atraso"
                    else:
                        ini_turno = _dt.datetime.combine(d, _dt.time(7 if tt == 0 else 19))
                        fim_turno = ini_turno + _dt.timedelta(hours=12)
                        sit = "Futuro" if agora < ini_turno else "Em andamento" if agora < fim_turno else "Não feito"
                out.append((d, t, ens, sit))
        d += _dt.timedelta(days=1)
    return out
