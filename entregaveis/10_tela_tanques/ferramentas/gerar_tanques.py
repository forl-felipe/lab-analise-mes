"""Gera a tela Tanques e Fluxos (padrão visual Samarco) para o Aspen Process Graphics Editor V14.2.

Refaz a tela Tanques.atgraphic com as mesmas tags, fórmulas e escalas, no padrão das telas Compressão
Online Usina 4 e Qualidade de Embarque.

Uso (nesta pasta):
    python gerar_tanques.py prototipo          (escreve ../prototipo/index.html, com dados simulados)
    node ../../ferramentas/gerar_imagens.js ../prototipo/index.html ../aspen normal 2240
    python gerar_tanques.py aspen              (gera ../aspen/Tanques_e_Fluxos.atgraphic e o passo a passo)

Tela de 70 x 35 unidades (como a original), protótipo de 2240 x 1120 px (32 px por unidade).
"""
import json
import math
import os
import random
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "..", "ferramentas"))

SAIDA = os.path.join(AQUI, "..", "aspen")
PROTO = os.path.join(AQUI, "..", "prototipo", "index.html")
ARQUIVO = "Tanques_e_Fluxos.atgraphic"
A, MES, PID3, PID8 = "IP_ANALOGMAP", "IP_MESVALOR", "IP_PIDMAP03", "IP_PIDMAP08"
LARGURA, ALTURA = 2240, 1120

# ----------------------------------------------------------------------------- tanques
# Estoque em tms = (nível - 1,5 m) x 400, zero abaixo de 1,5 m: mesma fórmula da tela original.
PREP1 = [("TQ 01", "lt-803"), ("TQ 02", "lt-805"), ("TQ 03", "lt-807"), ("TQ 04", "li-3802")]
PREP2 = [("TQ 01", "301HG001-LIT001-R"), ("TQ 02", "301HG002-LIT002-R"), ("TQ 08", "401-LI-001"),
         ("TQ 09", "401-LI-002"), ("TQ 10", "401-LI-004")]


def termo(tag):
    return f"( IF {{{tag}}} -1.5 < 0 THEN 0 ELSE   {{{tag}}} - 1.5)*400"


def formula(tags):
    return "= " + " + ".join(termo(t) for t in tags)


F_P1 = formula([t for _, t in PREP1])
F_P2 = formula([t for _, t in PREP2])
F_TOT = formula([t for _, t in PREP1 + PREP2])

# ----------------------------------------------------------------------------- qualidade (MES)
# (nome, unidade, tag, casas decimais, escala mín, escala máx): escalas da tela original
QUALIDADE = [("PPC pellet feed", "%", "M4650030010-0013-HHLQU", 2, 0, 5),
             ("SiO2", "%", "M4710050020-0004-HHLQU", 2, 0, 5),
             ("Carbono (%C)", "%", "M4710050020-0493-HHLQU", 2, 0, 5),
             ("Tamboramento", "%", "M4710050020-0029-HHLFU", 2, 91, 95),
             ("Compressão FX maior", "kgf/pelota", "M4710050020-0031-HHLFU", 0, 150, 350),
             ("Compressão FX menor", "kgf/pelota", "M4710050020-0035-HHLFU", 0, 150, 350)]

# ----------------------------------------------------------------------------- fluxos
# título, penas [(tag, mapa, mín, máx)], legenda da escala, campos [(tag, mapa, decimais, rótulo, unidade)]
FLUXOS = [
    ("Pelotas Usina 4", [("407-WI-001_10", A, 0, 2000)], "Vazão, 0 a 2000 t/h",
     [("407-WIQ-001_1 ", PID3, 0, "Acumulado do dia", "t")]),
    ("Pelotas Usina 3", [("307-WI-001", A, 0, 2000)], "Vazão, 0 a 2000 t/h",
     [("307-WI-001_10_Dia", A, 0, "Acumulado do dia", "t")]),
    ("Ritmo (MES)", [("M4710050031-0150-HHCC", MES, 0, 30000)], "Ritmo, 0 a 30000 t/dia",
     [("M4710050031-0150-HHCC", MES, 0, "Ritmo informado no MES", "t/dia")]),
    ("Repolpagem", [("001-WI-1000", A, 0, 400)], "Vazão, 0 a 400",
     [("U00-001WIT-1000", A, 0, "Valor atual", "")]),
    ("Material para emergência 3", [("307TP008-WIT003-R", A, 0, 3000)], "Vazão, 0 a 3000 t/h",
     [("307TP008-WIT003-R", A, 0, "Vazão atual", "t/h")]),
    ("Material para emergência 4", [("407-FIQ-707", A, 0, 3000)], "Vazão, 0 a 3000 t/h",
     [("407-WIQ-007", PID3, 0, "Acumulado do dia", "t")]),
    ("Emergência 4", [("406-FIQ-610", A, 0, 2000)], "Vazão, 0 a 2000 t/h",
     [("406-FIQ-610_1-Dia", A, 0, "Acumulado do dia", "t")]),
    ("Emergência 2", [("008-WI-010_1", A, 0, 2000)], "Vazão, 0 a 2000 t/h",
     [("008-WIQ-010_10", PID3, 0, "Acumulado do dia", "t")]),
    ("PSC pátio", [("407-WI-003_10", A, 0, 100)], "Vazão, 0 a 100",
     [("407-WIQ-003_10", PID3, 0, "Acumulado do dia", "t")]),
    ("PSC MI", [("407-WI-003_2", A, 0, 100)], "Vazão, 0 a 100",
     [("407-WIQ-003_2_D", A, 0, "Acumulado do dia", "t")]),
    ("Retorno de pellet feed", [("U04-03TP003", PID8, 0, 100)], "Escala 0 a 100",
     [("U04-03TP003", PID8, 1, "Valor atual", "")]),
    ("Pellet feed Usina 3", [("303TP005-WI001.1-R", A, 0, 5000), ("303TP005-IT001-R", A, 0, 100)],
     "Vazão 0 a 5000 t/h (azul) e 0 a 100 (branco)",
     [("008-WIQ-010_30", PID3, 0, "Retorno", "t"), ("008-WIQ-010_20", PID3, 0, "Expedição", "t")]),
]

AZUL_CLARO, BRANCO, AMARELO = (89, 198, 242), (255, 255, 255), (255, 192, 0)
PLOT = (1, 41, 74)


# ============================================================================= protótipo
def prototipo():
    random.seed(3)
    spec = {}          # objeto -> descrição para o gerador (gravada junto do protótipo)
    n = [0]

    def slot(desc):
        n[0] += 1
        spec[n[0]] = desc
        return n[0]

    def serie(k=40, base=0.6, amp=0.2):
        return [max(0.02, min(0.98, base + amp * math.sin(i / 4 + random.random()) + random.uniform(-0.05, 0.05))) for i in range(k)]

    def linha(vals, w, h, cor, larg=2.5):
        pts = " ".join(f"{2 + i * (w - 4) / (len(vals) - 1):.1f},{(1 - v) * h:.1f}" for i, v in enumerate(vals))
        return f'<polyline points="{pts}" fill="none" stroke="{cor}" stroke-width="{larg}" stroke-linejoin="round"/>'

    def tanques(titulo, lista, f_total, max_total, top):
        cols = []
        for nome, tag in lista:
            v = random.uniform(1500, 3400)
            nb = slot(dict(tipo="barra", tag="= " + termo(tag), mapa="Fixed", min=0, max=5000, cor="89, 198, 242",
                           nome=f"{titulo}: {nome}, nível em tms"))
            nv = slot(dict(tipo="campo", tag="= " + termo(tag), mapa="Fixed", dec=0, nome=f"{titulo}: {nome}, estoque (tms)"))
            cols.append((nome, v, 5000, nb, nv, "var(--sam-azul-claro)"))
        tot = sum(c[1] for c in cols)
        nb = slot(dict(tipo="barra", tag=f_total, mapa="Fixed", min=0, max=max_total, cor="255, 192, 0",
                       nome=f"{titulo}: total, nível em tms"))
        nv = slot(dict(tipo="campo", tag=f_total, mapa="Fixed", dec=0, nome=f"{titulo}: total (tms)"))
        cols.append(("Total", tot, max_total, nb, nv, "var(--sam-amarelo)"))
        html = []
        for nome, v, mx, nb, nv, cor in cols:
            html.append(f'''      <div class="tq{' total' if nome == 'Total' else ''}"><div class="tq-nome">{nome}</div>
        <div class="slot barra" data-slot="{nb}"><div class="enche" style="height:{v / mx * 100:.0f}%; background:{cor}"></div></div>
        <div class="tq-val"><span class="slot" data-slot="{nv}">{v:.0f}</span><span class="un">tms</span></div></div>''')
        return f'''  <section class="card" style="left:32px; top:{top}px; width:1000px; height:300px;">
    <div class="topo"><h2><span class="acento"></span>{titulo}</h2><span class="rot" style="margin-left:auto">Estoque em tms = (nível - 1,5 m) x 400</span></div>
    <div class="tanques">
{chr(10).join(html)}
    </div>
  </section>'''

    corpo = [tanques("Tanques: Preparação 1", PREP1, F_P1, 20000, 117),
             tanques("Tanques: Preparação 2", PREP2, F_P2, 25000, 433)]

    # Estoque total
    ntot = slot(dict(tipo="campo", tag=F_TOT, mapa="Fixed", dec=0, nome="Estoque total: Preparação 1 + 2 (tms)"))
    np1 = slot(dict(tipo="campo", tag=F_P1, mapa="Fixed", dec=0, nome="Estoque total: Preparação 1 (tms)"))
    np2 = slot(dict(tipo="campo", tag=F_P2, mapa="Fixed", dec=0, nome="Estoque total: Preparação 2 (tms)"))
    ntend = slot(dict(tipo="tendencia", nome="Estoque total: tendência",
                      penas=[(F_P2, "Fixed", 0, 45000, BRANCO), (F_P1, "Fixed", 0, 45000, AMARELO), (F_TOT, "Fixed", 0, 45000, AZUL_CLARO)]))
    PW, PH = 600, 210
    grade = "".join(f'<div class="gl sobre" style="top:{(1 - v / 45000) * PH:.0f}px"></div><div class="ylab" style="top:{(1 - v / 45000) * PH - 8:.0f}px">{v}</div>'
                    for v in (15000, 30000))
    corpo.append(f'''  <section class="card" style="left:32px; top:749px; width:1000px; height:339px;">
    <h2><span class="acento"></span>Estoque total nos tanques</h2>
    <div class="tot">
      <div class="tot-esq">
        <div class="rot">Preparação 1 + Preparação 2</div>
        <div style="display:flex; align-items:baseline"><span class="slot v-grande" data-slot="{ntot}" style="width:200px">26318</span><span class="unid">tms</span></div>
        <div class="tot-sub"><span class="traco" style="background:var(--sam-amarelo)"></span><span class="rot">Preparação 1</span><span class="slot v-med" data-slot="{np1}" style="width:90px">10433</span><span class="rot">tms</span></div>
        <div class="tot-sub"><span class="traco" style="background:#FFFFFF"></span><span class="rot">Preparação 2</span><span class="slot v-med" data-slot="{np2}" style="width:90px">15885</span><span class="rot">tms</span></div>
      </div>
      <div class="graf" style="width:{PW + 48}px; height:{PH}px">
        <div class="plot-bg" style="width:{PW}px; height:{PH}px"></div>{grade}
        <div class="slot" data-slot="{ntend}" style="position:absolute; left:48px; top:0; width:{PW}px; height:{PH}px"><svg width="{PW}" height="{PH}">
          {linha(serie(base=0.58, amp=0.05), PW, PH, "var(--sam-azul-claro)", 3)}{linha(serie(base=0.35, amp=0.04), PW, PH, "#FFFFFF", 2)}{linha(serie(base=0.23, amp=0.04), PW, PH, "var(--sam-amarelo)", 2)}</svg></div>
        <div class="leg"><span class="traco" style="background:var(--sam-azul-claro)"></span>Total<span class="traco" style="background:var(--sam-amarelo)"></span>Preparação 1<span class="traco" style="background:#FFFFFF"></span>Preparação 2<span class="rot">escala 0 a 45000 tms</span></div>
      </div>
    </div>
  </section>''')

    # Qualidade
    tiles = []
    for nome, un, tag, dec, mn, mx in QUALIDADE:
        nv = slot(dict(tipo="campo", tag=tag, mapa=MES, dec=dec, nome=f"Qualidade: {nome} ({un})"))
        nt = slot(dict(tipo="tendencia", nome=f"Qualidade: {nome}, minitendência", penas=[(tag, MES, mn, mx, AZUL_CLARO)]))
        v = random.uniform(mn + (mx - mn) * 0.4, mn + (mx - mn) * 0.7)
        tiles.append(f'''      <div class="ql"><div class="ql-nome">{nome}</div><div class="ql-un">{un}, escala {mn} a {mx}</div>
        <span class="slot ql-val" data-slot="{nv}">{f"{v:.{dec}f}".replace(".", ",")}</span>
        <div class="slot" data-slot="{nt}" style="width:146px; height:44px"><svg width="146" height="44">{linha(serie(20, 0.55, 0.15), 146, 44, "var(--sam-azul-claro)", 2)}</svg></div></div>''')
    corpo.append(f'''  <section class="card" style="left:1052px; top:117px; width:1156px; height:236px;">
    <h2><span class="acento"></span>Qualidade (MES)</h2>
    <div class="qls">
{chr(10).join(tiles)}
    </div>
  </section>''')

    # Fluxos
    for k, (titulo, penas, escala, campos) in enumerate(FLUXOS):
        lin, col = divmod(k, 4)
        x, y = 1052 + col * 292, 369 + lin * 245
        vals = []
        for tag, mapa, dec, rot, un in campos:
            nv = slot(dict(tipo="campo", tag=tag, mapa=mapa, dec=dec, nome=f"{titulo}: {rot.lower()}" + (f" ({un})" if un else "")))
            v = random.uniform(100, 18000) if un in ("t", "t/dia") else random.uniform(10, 1500)
            vals.append((nv, f"{v:.{dec}f}".replace(".", ","), rot, un))
        cores = [AZUL_CLARO, BRANCO]
        nt = slot(dict(tipo="tendencia", nome=f"{titulo}: tendência",
                       penas=[(t, m, mn, mx, cores[i]) for i, (t, m, mn, mx) in enumerate(penas)][::-1]))
        traco = "".join(linha(serie(30, 0.55 - 0.25 * i, 0.15), 244, 92, "var(--sam-azul-claro)" if i == 0 else "#FFFFFF", 2.5) for i in range(len(penas)))
        if len(vals) == 1:
            nv, v, rot, un = vals[0]
            valores = f'''<div class="fx-val"><span class="slot fx-num" data-slot="{nv}">{v}</span><span class="un">{un}</span></div>
      <div class="rot fx-rot">{rot}</div>'''
        else:
            valores = '<div class="fx-dois">' + "".join(
                f'<div><div class="rot">{rot}</div><span class="slot fx-num2" data-slot="{nv}">{v}</span><span class="un">{un}</span></div>'
                for nv, v, rot, un in vals) + "</div>"
        corpo.append(f'''  <section class="card fx" style="left:{x}px; top:{y}px; width:280px; height:233px;">
    <h3><span class="acento"></span>{titulo}</h3>
      {valores}
    <div class="rot fx-esc">{escala}</div>
    <div class="fx-graf"><div class="plot-bg" style="left:0; width:244px; height:92px"></div>
      <div class="slot" data-slot="{nt}" style="position:absolute; left:0; top:0; width:244px; height:92px"><svg width="244" height="92">{traco}</svg></div></div>
  </section>''')

    html = MODELO.replace("{{CORPO}}", "\n\n".join(corpo))
    with open(PROTO, "w", encoding="utf-8") as fh:
        fh.write(html)
    with open(os.path.join(SAIDA, "especificacao.json"), "w", encoding="utf-8") as fh:
        json.dump(spec, fh, ensure_ascii=False, indent=1)
    print(f"{PROTO}: {len(spec)} objetos dinâmicos")


MODELO = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tanques e Fluxos</title>
<!--
  Protótipo da tela Tanques e Fluxos (padrão visual Samarco). Dados simulados.
  Gerado por ../ferramentas/gerar_tanques.py; não editar à mão.
  Parâmetros na URL: ?modo=demo | fundo | mapa
-->
<style>
  :root {
    --sam-azul: #00335A; --sam-profundo: #002643; --sam-amarelo: #FFC000; --sam-azul-claro: #59C6F2; --sam-cinza: #B5BEC4;
    --bg: var(--sam-profundo); --card: var(--sam-azul); --card-line: #0B4F82; --tile: #013E6B; --plot: #01294A;
    --grade: #164D78; --text: #FFFFFF; --muted: var(--sam-cinza); --ui: Arial, "Liberation Sans", sans-serif;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { background: #0d1117; color: var(--text); font-family: var(--ui); overflow: hidden; height: 100%; }
  #palco { position: absolute; left: 0; top: 0; width: 2240px; height: 1120px; background: var(--bg); transform-origin: 0 0; overflow: hidden; }
  .cab { position: absolute; left: 0; top: 0; width: 2240px; height: 92px; background: var(--sam-azul); display: flex; align-items: center; padding: 0 32px; }
  .filete { position: absolute; left: 0; top: 92px; width: 2240px; height: 5px; background: var(--sam-amarelo); }
  .logo, .logo img { width: 220px; height: 48px; display: block; }
  .sep { width: 1px; height: 44px; background: #2C5C85; margin: 0 28px; }
  .cab h1 { font-size: 27px; font-weight: 700; }
  .card { position: absolute; background: var(--card); border: 1px solid var(--card-line); border-radius: 10px; padding: 16px 24px; }
  .card h2 { font-size: 20px; font-weight: 700; display: flex; align-items: center; gap: 10px; height: 26px; }
  .card h3 { font-size: 16px; font-weight: 700; display: flex; align-items: center; gap: 8px; height: 22px; white-space: nowrap; }
  .topo { display: flex; align-items: center; }
  .acento { width: 5px; height: 22px; background: var(--sam-amarelo); border-radius: 2px; }
  h3 .acento { height: 18px; }
  .rot { font-size: 13px; color: var(--muted); }
  .un, .unid { font-size: 13px; color: var(--muted); margin-left: 6px; }
  .unid { font-size: 18px; margin-left: 10px; }
  .slot { display: inline-block; white-space: nowrap; overflow: hidden; color: var(--text); font-family: var(--ui); }
  .v-grande { font-size: 64px; font-weight: 700; line-height: 72px; height: 72px; text-align: right; }
  .v-med { font-size: 18px; font-weight: 700; line-height: 24px; height: 24px; text-align: right; }

  .tanques { display: flex; gap: 18px; margin-top: 14px; }
  .tq { width: 140px; background: var(--tile); border: 1px solid var(--card-line); border-radius: 8px; padding: 10px 12px; height: 226px; }
  .tq.total { border-color: var(--sam-amarelo); }
  .tq-nome { font-size: 17px; font-weight: 700; text-align: center; height: 22px; }
  .barra { display: block; width: 56px; height: 140px; margin: 6px auto 0; background: var(--plot); position: relative; }
  .enche { position: absolute; left: 0; right: 0; bottom: 0; }
  .tq-val { display: flex; align-items: baseline; justify-content: center; margin-top: 8px; }
  .tq-val .slot { font-size: 24px; font-weight: 700; line-height: 30px; height: 30px; width: 76px; text-align: right; }

  .tot { display: flex; gap: 32px; margin-top: 16px; align-items: flex-start; }
  .tot-esq { width: 280px; }
  .tot-sub { display: flex; align-items: center; gap: 8px; margin-top: 18px; }
  .traco { width: 18px; height: 4px; border-radius: 2px; display: inline-block; }
  .graf { position: relative; }
  .plot-bg { position: absolute; left: 48px; top: 0; background: var(--plot); }
  .gl { position: absolute; left: 48px; width: 600px; height: 1px; background: var(--grade); }
  .ylab { position: absolute; left: 0; width: 42px; text-align: right; font-size: 12px; color: var(--muted); line-height: 16px; }
  .leg { position: absolute; left: 48px; top: 220px; display: flex; align-items: center; gap: 8px; font-size: 13px; color: var(--muted); white-space: nowrap; }
  .leg .traco { margin-left: 10px; }

  .qls { display: flex; gap: 12px; margin-top: 12px; }
  .ql { width: 174px; height: 160px; background: var(--tile); border: 1px solid var(--card-line); border-radius: 8px; padding: 10px 12px; }
  .ql-nome { font-size: 14px; font-weight: 700; height: 19px; white-space: nowrap; }
  .ql-un { font-size: 11px; color: var(--muted); height: 15px; white-space: nowrap; }
  .ql-val { display: block; font-size: 32px; font-weight: 700; line-height: 40px; height: 40px; text-align: right; margin: 4px 0 6px; }

  .fx { padding: 14px 18px; }
  .fx-val { display: flex; align-items: baseline; margin-top: 10px; height: 36px; }
  .fx-num { font-size: 30px; font-weight: 700; line-height: 36px; height: 36px; width: 150px; text-align: right; }
  .fx-rot { height: 16px; margin-top: 2px; }
  .fx-dois { display: flex; gap: 18px; margin-top: 8px; height: 54px; }
  .fx-num2 { font-size: 22px; font-weight: 700; line-height: 28px; height: 28px; width: 80px; text-align: right; }
  .fx-esc { font-size: 11px; height: 15px; margin-top: 6px; white-space: nowrap; }
  .fx-graf { position: relative; width: 244px; height: 92px; margin-top: 4px; }

  body.fundo .slot { visibility: hidden; }
  body.mapa .slot { visibility: visible !important; color: transparent !important; outline: 2px dashed #ff4fd8; outline-offset: -2px; position: relative; background: rgba(255,79,216,.12); }
  body.mapa .slot > * { visibility: hidden !important; }
  body.mapa .slot::after { content: attr(data-slot); position: absolute; left: 4px; top: 2px; font: 700 14px/1 Arial, sans-serif; color: #fff; background: #c2189e; padding: 2px 5px; border-radius: 3px; }
</style>
</head>
<body>
<div id="palco">
  <header class="cab">
    <div class="logo" id="logo"><img src="../../assets/samarco-logo-branca.svg" alt="Samarco"></div>
    <div class="sep"></div>
    <h1>Tanques e Fluxos</h1>
  </header>
  <div class="filete"></div>

{{CORPO}}
</div>
<script>
(function () {
  const params = new URLSearchParams(location.search);
  const modo = params.get('modo') || 'demo';
  if (modo === 'fundo' || modo === 'mapa') document.body.classList.add(modo);
  const palco = document.getElementById('palco');
  function escala() {
    if (params.get('captura')) return;
    const s = Math.min(innerWidth / 2240, innerHeight / 1120);
    palco.style.transform = `scale(${s})`;
    palco.style.left = ((innerWidth - 2240 * s) / 2) + 'px';
    palco.style.top = ((innerHeight - 1120 * s) / 2) + 'px';
  }
  addEventListener('resize', escala); escala();
})();
</script>
</body>
</html>
"""


# ============================================================================= arquivo do Aspen
def aspen():
    import gerar_atgraphic as g
    import montar_pacote_aspen as mp

    with open(os.path.join(SAIDA, "especificacao.json"), encoding="utf-8") as fh:
        esp = {int(k): v for k, v in json.load(fh).items()}
    SPEC, PENAS = {}, {}
    FMT = {0: "Sem casas decimais (0)", 1: "1 casa decimal (0,0)", 2: "2 casas decimais (0,00)"}
    for n, e in esp.items():
        if e["tipo"] == "campo":
            SPEC[n] = mp.campo(e["nome"], e["tag"], e["mapa"], "VAL", 14, True, "Direita", "1, 62, 107", FMT[e["dec"]],
                               obs="Campo calculado (fórmula da tela original)." if e["mapa"] == "Fixed" else "")
        elif e["tipo"] == "barra":
            SPEC[n] = dict(ferramenta="Barra", nome=e["nome"], tag=e["tag"], mapa="Fixed", min=e["min"], max=e["max"],
                           cor=e["cor"], fundo_barra="1, 41, 74", obs=f"Barra de nível, escala {e['min']} a {e['max']} tms.")
        else:
            penas = [tuple(p[:4]) + (tuple(p[4]),) for p in e["penas"]]
            SPEC[n] = mp.tendencia(e["nome"], "Já vem pronta no arquivo: " + "; ".join(
                f"{p[0][:40] if p[1] != 'Fixed' else 'fórmula'} de {p[2]:g} a {p[3]:g}" for p in penas) + ".",
                ", ".join(p[0] if p[1] != "Fixed" else "fórmula" for p in penas))
            PENAS[n] = penas
    g.SPEC, g.PENAS, g.FUNDO_TENDENCIA = SPEC, PENAS, {n: PLOT for n in PENAS}
    mp.SPEC = SPEC

    with open(os.path.join(SAIDA, "objetos_tela.json"), encoding="utf-8") as fh:
        dados = json.load(fh)
    with open(os.path.join(SAIDA, "posicoes.json"), encoding="utf-8") as fh:
        posicoes = json.load(fh)
    assert {p["n"] for p in posicoes} == set(SPEC), "objetos do protótipo e da especificação não conferem"

    fluxo, objetos = g.conteudo(dados)
    with open(os.path.join(SAIDA, ARQUIVO), "wb") as fh:
        fh.write(g.cfb(fluxo))
    tipos = {}
    for nome, _ in objetos:
        tipos[nome] = tipos.get(nome, 0) + 1
    print(f"{ARQUIVO}: {len(objetos)} objetos {tipos}, fluxo de {len(fluxo)} bytes")
    minimo = dict(dados, formas=[f for f in dados["formas"] if f["preenchimento"] and f["linha"]][:1], textos=dados["textos"][:1],
                  simbolos=[], lampadas=[], logo=None, campos=[c for c in dados["campos"] if c["n"] in (1, 2, len(esp))])
    with open(os.path.join(SAIDA, "teste_minimo.atgraphic"), "wb") as fh:
        fh.write(g.cfb(g.conteudo(minimo)[0]))

    mp.planilha(posicoes, os.path.join(SAIDA, "posicoes_objetos.xlsx"))
    with open(os.path.join(AQUI, "roteiro_modelo.md"), encoding="utf-8") as fh:
        modelo = fh.read()
    fichas = "\n".join(ficha(p["n"], SPEC[p["n"]], p, mp) for p in sorted(posicoes, key=lambda p: p["n"]))
    texto = (modelo.replace("{{FICHAS}}", fichas)
             .replace("{{N_CAMPOS}}", str(sum(1 for s in SPEC.values() if s["ferramenta"] == "Campo de dados")))
             .replace("{{N_BARRAS}}", str(sum(1 for s in SPEC.values() if s["ferramenta"] == "Barra")))
             .replace("{{N_TEND}}", str(len(PENAS)))
             .replace("{{N_TEXTOS}}", str(len(dados["textos"]))))
    mp.publicar_roteiro(texto, SAIDA)
    print(f"Pacote montado em {os.path.abspath(SAIDA)}: {len(SPEC)} objetos dinâmicos")


def ficha(n, s, pos, mp):
    if s["ferramenta"] != "Barra":
        t = mp.ficha(n, s, pos)
        return t.replace(f"| Tag | `{s.get('tag', '')}` |", "| Tag | Fórmula da tela original (ver Parte B) |") if s.get("mapa") == "Fixed" else t
    return "\n".join([f"### Objeto {n}: {s['nome']}", "", "| Propriedade | Valor |", "|---|---|", "| Ferramenta | Barra (bargraph) |",
                      f"| Posição em unidades (X ; Y) | {mp.un(pos['x'])} ; {mp.un(pos['y'])} |",
                      f"| Tamanho em unidades (largura ; altura) | {mp.un(pos['w'])} ; {mp.un(pos['h'])} |",
                      "| Tag | Fórmula da tela original (ver Parte B) |", f"| Escala | {s['min']} a {s['max']} tms |"]) + "\n"


if __name__ == "__main__":
    {"prototipo": prototipo, "aspen": aspen}[sys.argv[1] if len(sys.argv) > 1 else "prototipo"]()
