"""Gera a tela Qualidade de Embarque para o Aspen Process Graphics Editor V14.2.

Uso (nesta pasta):
    node ../../ferramentas/gerar_imagens.js ../prototipo/index.html ../aspen normal
    python gerar_embarque.py

Saídas em ../aspen/: Qualidade_de_Embarque.atgraphic, teste_minimo.atgraphic, posições dos objetos
e o passo a passo (roteiro_aspen_graphics.md / .html / .pdf).

Reaproveita o gerador da tela Compressão Online Usina 4 (../../ferramentas), que já abriu no PIMS:
mesmo formato de arquivo, só mudam os objetos, as tags e as penas das tendências.
"""
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "..", "ferramentas"))
import gerar_atgraphic as g  # noqa: E402
import montar_pacote_aspen as mp  # noqa: E402

SAIDA = os.path.join(AQUI, "..", "aspen")
ARQUIVO = "Qualidade_de_Embarque.atgraphic"
A, MES = "IP_ANALOGMAP", "IP_MESVALOR"
CARD, TILE, CAIXA = "0, 51, 90", "1, 62, 107", "1, 41, 74"
HORA = mp.HORA
FMT = {0: "Sem casas decimais (0)", 1: "1 casa decimal (0,0)", 2: "2 casas decimais (0,00)", 3: "3 casas decimais (0,000)"}

# Elementos de qualidade: (objeto, nome, tag, casas decimais)
FISICO = [(24, "Compressão 16,0 mm (kgf/pelota)", "M0620010000-0031-TNLFU", 0),
          (25, "Compressão 12,5 mm (kgf/pelota)", "M0620010000-0035-TNLFU", 0),
          (26, "Tamboramento +6,3 mm (%)", "M0620010000-0029-TNLFU", 1),
          (27, "Finos -6,3 mm (%)", "M0620010000-0026-TNLFU", 1),
          (28, "Relação granulométrica (%)", "M0620010000-2021-TNLFU", 2),
          (29, "Umidade H2O (%)", "M0620010000-0001-TNLFU", 2)]
QUIMICO = [(31, "Fe (%)", "M0620010000-0002-TNLQU", 2), (32, "SiO2 (%)", "M0620010000-0004-TNLQU", 2),
           (33, "Al2O3 (%)", "M0620010000-0005-TNLQU", 2), (34, "CaO (%)", "M0620010000-0006-TNLQU", 2),
           (35, "MgO (%)", "M0620010000-0007-TNLQU", 2), (36, "B2 (%)", "M0620010000-0018-TNLQU", 2),
           (37, "P (%)", "M0620010000-0008-TNLQU", 3), (38, "Mn (%)", "M0620010000-0115-TNLQU", 3),
           (39, "PPC (%)", "M0620010000-0013-TNLQU", 2)]

SPEC = {
    1: mp.campo("Ritmo de carregamento: balança principal (t/h)", "008-WI-010", A, "VAL", 44, True, "Direita", TILE, FMT[0]),
    2: mp.campo("Ritmo de carregamento: balança redundante (t/h)", "008-WI-010_1", A, "VAL", 17, True, "Direita", TILE, FMT[0]),
    3: mp.campo("Berço Oeste: total carregado (t)", "309-WIT-004.TOT_NAV_OE", A, "VAL", 44, True, "Direita", TILE, FMT[0],
                obs="A lâmpada ao lado do título acende em verde quando o total é maior que zero (já vem pronta)."),
    15: mp.campo("Berço Leste: total carregado (t)", "309-WIT-004.TOT_NAV_LE", A, "VAL", 44, True, "Direita", TILE, FMT[0],
                 obs="A lâmpada ao lado do título acende em verde quando o total é maior que zero (já vem pronta)."),
    23: mp.campo("Físico: último resultado", "M0620010000-0031-TNLFU", MES, HORA, 13, False, "Esquerda", CARD, mp.DH),
    30: mp.campo("Químico: último resultado", "M0620010000-0002-TNLQU", MES, HORA, 13, False, "Esquerda", CARD, mp.DH),
    40: mp.tendencia("Tendência: ritmo de carregamento", "Já vem pronta no arquivo, escala 0 a 2500 t/h. Ver Parte C.", "008-WI-010"),
    41: mp.tendencia("Tendência: compressão", "Já vem pronta no arquivo, escala 150 a 450 kgf/pelota. Ver Parte C.",
                     "M0620010000-0031-TNLFU e M0620010000-0035-TNLFU"),
    42: mp.tendencia("Tendência: Fe e SiO2", "Já vem pronta no arquivo, Fe de 62 a 68% e SiO2 de 0 a 6%. Ver Parte C.",
                     "M0620010000-0002-TNLQU e M0620010000-0004-TNLQU"),
}
for k in range(11):
    SPEC[4 + k] = mp.campo(f"Berço Oeste: porão P{k + 1} (t)", f"309-WIT-004.TOT_POR_OE{k + 1}", A, "VAL", 15, True, "Direita", CAIXA, FMT[0])
for k in range(7):
    SPEC[16 + k] = mp.campo(f"Berço Leste: porão P{k + 1} (t)", f"309-WIT-004.TOT_POR_LE{k + 1:02d}", A, "VAL", 15, True, "Direita", CAIXA, FMT[0])
for n, nome, tag, dec in FISICO:
    SPEC[n] = mp.campo(f"Físico: {nome}", tag, MES, "VAL", 37, True, "Direita", TILE, FMT[dec])
for n, nome, tag, dec in QUIMICO:
    SPEC[n] = mp.campo(f"Químico: {nome}", tag, MES, "VAL", 29, True, "Direita", TILE, FMT[dec])

AZUL_CLARO, BRANCO, PLOT = (89, 198, 242), (255, 255, 255), (1, 41, 74)
PENAS = {
    40: [("008-WI-010", A, 0, 2500, AZUL_CLARO)],
    41: [("M0620010000-0035-TNLFU", MES, 150, 450, BRANCO), ("M0620010000-0031-TNLFU", MES, 150, 450, AZUL_CLARO)],
    42: [("M0620010000-0004-TNLQU", MES, 0, 6, BRANCO), ("M0620010000-0002-TNLQU", MES, 62, 68, AZUL_CLARO)],
}


def main():
    g.SPEC, g.PENAS, g.FUNDO_TENDENCIA = SPEC, PENAS, {n: PLOT for n in PENAS}
    mp.SPEC = SPEC
    with open(os.path.join(SAIDA, "objetos_tela.json"), encoding="utf-8") as fh:
        dados = json.load(fh)
    with open(os.path.join(SAIDA, "posicoes.json"), encoding="utf-8") as fh:
        posicoes = json.load(fh)
    assert {p["n"] for p in posicoes} == set(SPEC), "Objetos do protótipo e da especificação não conferem"

    fluxo, objetos = g.conteudo(dados)
    with open(os.path.join(SAIDA, ARQUIVO), "wb") as fh:
        fh.write(g.cfb(fluxo))
    tipos = {}
    for nome, _ in objetos:
        tipos[nome] = tipos.get(nome, 0) + 1
    print(f"{ARQUIVO}: {len(objetos)} objetos {tipos}, fluxo de {len(fluxo)} bytes")
    minimo = dict(dados, formas=[f for f in dados["formas"] if f["preenchimento"] and f["linha"]][:1],
                  textos=dados["textos"][:1], simbolos=[], lampadas=dados["lampadas"][:1], logo=None,
                  campos=[c for c in dados["campos"] if c["n"] in (1, 40)])
    with open(os.path.join(SAIDA, "teste_minimo.atgraphic"), "wb") as fh:
        fh.write(g.cfb(g.conteudo(minimo)[0]))

    linhas = mp.planilha(posicoes, os.path.join(SAIDA, "posicoes_objetos.xlsx"))
    with open(os.path.join(AQUI, "roteiro_modelo.md"), encoding="utf-8") as fh:
        modelo = fh.read()
    fichas = "\n".join(mp.ficha(p["n"], SPEC[p["n"]], p) for p in sorted(posicoes, key=lambda p: p["n"]))
    texto = (modelo.replace("{{FICHAS}}", fichas)
             .replace("{{N_CAMPOS}}", str(sum(1 for s in SPEC.values() if s["ferramenta"] == "Campo de dados")))
             .replace("{{N_TEND}}", str(len(PENAS)))
             .replace("{{N_FORMAS}}", str(sum(1 for n, _ in objetos if n == "AtBasicShape") - len(dados["lampadas"])))
             .replace("{{N_TEXTOS}}", str(len(dados["textos"]))))
    mp.publicar_roteiro(texto, SAIDA)
    print(f"Pacote montado em {os.path.abspath(SAIDA)}: {len(SPEC)} objetos dinâmicos, {len(linhas) - 1} posições")


if __name__ == "__main__":
    main()
