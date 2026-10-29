"""Gera a planilha com macros de correlação entre a prensa automática (307-GI-001) e o laboratório (MES).

Uso: python gerar_planilha_aspen.py [saida.xlsm] [--teste dados.json]

A busca dos dados segue o Relatório Físico: a macro (vba/ModPIMS.bas) escreve a fórmula
GetCalculationValues do suplemento Aspen Process Data na aba Dados_PIMS, espera o retorno e grava
os resultados nas abas Prensa e MES. Correlacao e Painel são fórmulas.

Regra (definida pelo LCE):
  - O resultado da prensa entra no PIMS 20 min depois do corte: 13:20 é o corte das 13:00.
  - Os cortes de 12:00, 12:30, 13:00 e 13:30 formam um composto de 2 h, comparado com o resultado
    do MES das 13:30; depois 14:00 a 15:30 com o MES das 15:30, e assim por diante.
  - O composto da prensa é a média dos cortes com resultado (1 a 4).

A planilha sai SEM DADOS. Com --teste, grava dados de um JSON (só para conferência das fórmulas;
gera .xlsx, sem macros).
"""
import json
import os
import sys
from datetime import datetime, timedelta

import xlsxwriter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_projeto import vba_project_bin  # noqa: E402

AQUI = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(AQUI, "..", "assets", "samarco-logo-branca.png")
NP = 5001       # última linha da aba Prensa (5.000 resultados, cerca de 100 dias)
NM = 1501       # última linha das abas MES e Correlacao (1.500 compostos, cerca de 125 dias)
INICIO_PRENSA = datetime(2026, 10, 28, 0, 0)

AZUL, AMARELO, AZUL_MEDIO, LARANJA = "#00335A", "#FFC000", "#0091D5", "#F37021"
F = "Arial"

args = [a for a in sys.argv[1:]]
teste = None
if "--teste" in args:
    i = args.index("--teste")
    teste = json.load(open(args[i + 1]))
    del args[i:i + 2]
SAIDA = args[0] if args else "correlacao_prensa_laboratorio.xlsm"
COM_VBA = SAIDA.lower().endswith(".xlsm")

wb = xlsxwriter.Workbook(SAIDA)
wb.set_properties({"title": "Correlação prensa automática x laboratório", "company": "LCE"})


def fmt(**k):
    base = {"font_name": F, "font_size": 10, "valign": "vcenter"}
    base.update(k)
    return wb.add_format(base)


f_cab = fmt(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, border=1, border_color="#C9D3DB")
f_ent = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB")
f_ent_dh = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB", num_format="dd/mm/yyyy hh:mm")
f_dado_dh = fmt(num_format="dd/mm/yyyy hh:mm")
f_dado_n = fmt(num_format="0.0")
f_calc = fmt(bg_color="#F2F2F2")
f_calc_dh = fmt(bg_color="#F2F2F2", num_format="dd/mm/yyyy hh:mm")
f_calc_1 = fmt(bg_color="#F2F2F2", num_format="0.0")
f_txt = fmt(text_wrap=True, valign="top")
f_sec = fmt(bold=True, font_size=12, font_color=AZUL)
f_tit = fmt(bold=True, font_size=15, font_color=AZUL)
f_alerta = fmt(bold=True, font_color="#C00000", text_wrap=True)
f_banda = fmt(bg_color=AZUL)
f_banda_tit = fmt(bold=True, font_size=18, font_color="white", bg_color=AZUL, indent=1)
f_amarelo = fmt(bg_color=AMARELO)
f_cartao = fmt(bg_color="#EEF3F7")
f_cartao_tit = fmt(bold=True, font_size=12, font_color=AZUL, bottom=2, bottom_color=AMARELO)
f_status = fmt(italic=True, font_color="#5B6B77")
f_rotulo = fmt(bold=True, font_color=AZUL)


def valor_cartao(num):
    return fmt(bold=True, font_size=12, font_color=AZUL, bg_color="#EEF3F7", align="right", num_format=num)


# ======================================================================= abas (na ordem de exibição)
wp = wb.add_worksheet("Painel")
wl = wb.add_worksheet("Leia-me")
wcfg = wb.add_worksheet("Configurações")
wpr = wb.add_worksheet("Prensa")
wm = wb.add_worksheet("MES")
wc = wb.add_worksheet("Correlacao")
wd = wb.add_worksheet("Dados_PIMS")
CODENOMES = {wp: "shPainel", wl: "shLeiame", wcfg: "shConfig", wpr: "shPrensa", wm: "shMES", wc: "shCorrelacao",
             wd: "shDados"}


def nome(n, ws, cel):
    wb.define_name(n, f"='{ws.get_name()}'!{cel}")


# ======================================================================= Configurações
wcfg.hide_gridlines(2)
wcfg.set_column("A:A", 50)
wcfg.set_column("B:B", 28)
wcfg.set_column("C:C", 100)
wcfg.write_row(0, 0, ["Parâmetro", "Valor", "Observação"], f_cab)
wcfg.set_row(0, 24)
linhas_cfg = [
    ("Conexão com o PIMS (Aspen Process Data no Excel)", None, None, None, None),
    ("Servidor do PIMS", "UBU", "cfgServidor", "Mesmo servidor do Relatório Físico.", None),
    ("Formato da data para o Aspen", "Texto dd/mm/aaaa", "cfgFormatoData",
     "Padrão do Relatório Físico. Opções: Texto dd/mm/aaaa, Texto mm/dd/aaaa ou Data do Excel.", "lista"),
    ("Espera máxima pela resposta do PIMS (s)", 120, "cfgTimeout", "Tempo máximo de cada consulta.", None),
    ("Início da fórmula do suplemento", "=GetCalculationValues(", "cfgPrefixo",
     "Como o Excel mostra a função do Aspen. Se o Excel não reconhecer, a macro tenta o nome completo do suplemento.",
     "texto"),
    ("", None, None, None, None),
    ("Prensa automática", None, None, None, None),
    ("Tag", "307-GI-001", "cfgTagPrensa", "Média de compressão do ensaio (kgf/pelota).", None),
    ("Tipo (mapa)", "IP_ANALOGMAP", "cfgMapaPrensa", "", None),
    ("Intervalo da consulta", "1m", "cfgPeriodoPrensa",
     "Leitura a cada 1 minuto. Cada resultado novo da tag vira uma linha da aba Prensa, com o horário em que apareceu no PIMS.",
     None),
    ("Tipo de cálculo do suplemento", "0", "cfgCalculoPrensa",
     "0 = valor no instante (mesmo código da consulta de ritmo de processo do Relatório Físico).", "texto"),
    ("Início de operação da prensa", INICIO_PRENSA, "cfgInicioPrensa",
     "Usado pelo botão Desde o início da prensa. Antes desta data a tag não tem resultados válidos.", "dh"),
    ("", None, None, None, None),
    ("Laboratório (MES)", None, None, None, None),
    ("Tag", "M4710050020-0031-HHLFU", "cfgTagLab",
     "Resistência à Compressão, pelota queimada US4, em kgf/pelota (Relatório Físico, aba Configurações, linha 40).",
     None),
    ("Tipo (mapa)", "IP_MESVALOR", "cfgMapaLab", "", None),
    ("Intervalo da consulta", "30m", "cfgPeriodoLab",
     "Intervalos de 30 min: o horário do resultado no MES fica identificado (13:30, 15:30 e assim por diante).", None),
    ("Tipo de cálculo do suplemento", "1", "cfgCalculoLab", "1 = mesmo código da consulta de qualidade do Relatório Físico.",
     "texto"),
    ("", None, None, None, None),
    ("Regra de comparação", None, None, None, None),
    ("Defasagem entre o corte e a entrada no PIMS (min)", 20, "cfgDefasagem",
     "Informado pelo LCE: o corte das 13:00 entra no PIMS às 13:20.", None),
    ("Deslocamento das janelas (meias horas)", 0, "cfgDeslocamento",
     "0 = janelas de 12:00 a 13:30, 14:00 a 15:30 e assim por diante. Use 1 para começar nas meias horas, 2 nas horas ímpares.",
     None),
    ("Mínimo de cortes para a janela ser válida", 1, "cfgMinCortes",
     "1 = usa a média mesmo com um só corte. Use 4 para considerar só janelas completas.", None),
    ("Ajuste do horário do MES (min)", 0, "cfgAjusteMES",
     "Use se o horário do resultado no MES não for o do último corte (13:30). Exemplo: -30 se o MES gravar 14:00.", None),
    ("Tolerância da diferença prensa x laboratório (kgf/pelota)", 20, "cfgTolerancia",
     "A confirmar pelo LCE. Usada no indicador de janelas dentro da tolerância.", None),
]
for i, (rot, v, n, obs, tipo) in enumerate(linhas_cfg, 1):
    if v is None:
        wcfg.write(i, 0, rot, f_rotulo if rot else None)
        continue
    wcfg.write(i, 0, rot, fmt())
    if tipo == "dh":
        wcfg.write_datetime(i, 1, v, f_ent_dh)
    elif tipo == "texto" or isinstance(v, str):
        wcfg.write_string(i, 1, str(v), f_ent)
    else:
        wcfg.write_number(i, 1, v, f_ent)
    if tipo == "lista":
        wcfg.data_validation(i, 1, i, 1, {"validate": "list",
                                          "source": ["Texto dd/mm/aaaa", "Texto mm/dd/aaaa", "Data do Excel"]})
    wcfg.write(i, 2, obs, fmt(text_wrap=True))
    nome(n, wcfg, f"$B${i + 1}")
for n in ("cfgDefasagem", "cfgDeslocamento", "cfgMinCortes", "cfgAjusteMES", "cfgTolerancia"):
    pass
P_DEF, P_DESL, P_MIN, P_AJ, P_TOL = "cfgDefasagem", "cfgDeslocamento", "cfgMinCortes", "cfgAjusteMES", "cfgTolerancia"

# ======================================================================= Dados_PIMS (oculta)
wd.set_column("A:A", 19)
wd.set_column("B:B", 21)
wd.set_column("E:E", 19)
wd.set_column("F:F", 21)
wd.write("A1", "Consultas ao PIMS (Aspen Process Data). Preenchida pela macro, não editar.", f_rotulo)
wd.write("A3", "Início")
wd.write_string("B3", "")
wd.write("A4", "Fim")
wd.write_string("B4", "")
wd.write("A6", "Consulta: prensa automática", f_rotulo)
wd.write("A7", "Sem consulta")
wd.write_row("A8", ["Data/hora", "Valor"], f_cab)
wd.write("E6", "Consulta: laboratório (MES)", f_rotulo)
wd.write("E7", "Sem consulta")
wd.write_row("E8", ["Data/hora", "Valor"], f_cab)
wd.hide()

# ======================================================================= Prensa
cab_pr = ["Data e hora no PIMS", "Média prensa 307-GI-001 (kgf/pelota)", "Hora do corte", "Índice de meia hora", "Janela",
          "Posição do corte na janela (1 a 4)", "Válido (1 = entra na média)", "Horário do MES correspondente"]
wpr.write_row(0, 0, cab_pr, f_cab)
wpr.set_row(0, 42)
for col, w in zip("ABCDEFGH", [19, 16, 17, 11, 10, 12, 12, 17]):
    wpr.set_column(f"{col}:{col}", w)
wpr.freeze_panes(1, 0)
wpr.write_comment("A1", "Preenchido pela macro (botões do Painel): um resultado por linha, com o horário em que entrou no PIMS.")
prensa = teste["prensa"] if teste else []
for r in range(2, NP + 1):
    i = r - 2
    if i < len(prensa):
        wpr.write_datetime(r - 1, 0, datetime.fromisoformat(prensa[i][0]), f_dado_dh)
        wpr.write_number(r - 1, 1, prensa[i][1], f_dado_n)
    else:
        wpr.write_blank(r - 1, 0, None, f_dado_dh)
        wpr.write_blank(r - 1, 1, None, f_dado_n)
    wpr.write_formula(r - 1, 3, f'=IF(A{r}="","",ROUND((A{r}-{P_DEF}/1440)*48,0))', f_calc)
    wpr.write_formula(r - 1, 2, f'=IF(D{r}="","",D{r}/48)', f_calc_dh)
    wpr.write_formula(r - 1, 4, f'=IF(D{r}="","",INT((D{r}-{P_DESL})/4))', f_calc)
    wpr.write_formula(r - 1, 5, f'=IF(D{r}="","",MOD(D{r}-{P_DESL},4)+1)', f_calc)
    wpr.write_formula(r - 1, 6, f'=IF(A{r}="","",IF(AND(ISNUMBER(B{r}),B{r}<>0),1,0))', f_calc)
    wpr.write_formula(r - 1, 7, f'=IF(E{r}="","",(E{r}*4+3+{P_DESL})/48)', f_calc_dh)

# ======================================================================= MES
wm.write_row(0, 0, ["Data e hora do resultado no MES", "Compressão laboratório (kgf/pelota)", "Índice de meia hora",
                    "Janela", "Posição (esperado 4)"], f_cab)
wm.set_row(0, 42)
for col, w in zip("ABCDE", [21, 16, 11, 10, 12]):
    wm.set_column(f"{col}:{col}", w)
wm.freeze_panes(1, 0)
wm.write_comment("A1", "Preenchido pela macro (botões do Painel): um resultado por composto de 2 h.")
mes = teste["mes"] if teste else []
for r in range(2, NM + 1):
    i = r - 2
    if i < len(mes):
        wm.write_datetime(r - 1, 0, datetime.fromisoformat(mes[i][0]), f_dado_dh)
        wm.write_number(r - 1, 1, mes[i][1], f_dado_n)
    else:
        wm.write_blank(r - 1, 0, None, f_dado_dh)
        wm.write_blank(r - 1, 1, None, f_dado_n)
    wm.write_formula(r - 1, 2, f'=IF(A{r}="","",ROUND((A{r}+{P_AJ}/1440)*48,0))', f_calc)
    wm.write_formula(r - 1, 3, f'=IF(C{r}="","",INT((C{r}-{P_DESL})/4))', f_calc)
    wm.write_formula(r - 1, 4, f'=IF(C{r}="","",MOD(C{r}-{P_DESL},4)+1)', f_calc)

# ======================================================================= Correlacao
tit = ["Horário do MES", "Janela de cortes", "Laboratório (kgf/pelota)", "Corte 1", "Corte 2", "Corte 3", "Corte 4",
       "Cortes com resultado", "Composto prensa (kgf/pelota)", "Diferença prensa menos laboratório", "Situação",
       "Diferença absoluta", "Dentro da tolerância",
       "id janela", "X gráfico", "Lab gráfico", "Prensa gráfico", "Dif gráfico", "Lab válido", "Prensa válido", "Dif válida"]
wc.write_row(0, 0, tit, f_cab)
wc.set_row(0, 42)
for k, w in enumerate([17, 15, 13, 10, 10, 10, 10, 11, 14, 14, 26, 11, 12] + [8] * 8):
    wc.set_column(k, k, w, None, {"hidden": k >= 13})
wc.freeze_panes(1, 0)
PE, PF, PG, PB = (f"Prensa!${c}$2:${c}${NP}" for c in "EFGB")
for r in range(2, NM + 1):
    f = {
        "A": (f'=IF(MES!A{r}="","",MES!C{r}/48)', f_calc_dh),
        "N": (f'=IF(A{r}="","",MES!D{r})', f_calc),
        "B": (f'=IF(A{r}="","",TEXT((N{r}*4+{P_DESL})/48,"hh:mm")&" a "&TEXT(A{r},"hh:mm"))', f_calc),
        "C": (f'=IF(A{r}="","",MES!B{r})', f_calc_1),
        "H": (f'=IF(A{r}="","",COUNT(D{r}:G{r}))', f_calc),
        "I": (f'=IF(A{r}="","",IF(H{r}=0,"",AVERAGE(D{r}:G{r})))', f_calc_1),
        "J": (f'=IF(OR(I{r}="",C{r}=""),"",I{r}-C{r})', f_calc_1),
        "K": (f'=IF(A{r}="","",IF(MES!E{r}<>4,"Horário do MES fora do padrão",'
              f'IF(H{r}<{P_MIN},"Cortes insuficientes","Válida")))', f_calc),
        "L": (f'=IF(K{r}="Válida",ABS(J{r}),"")', f_calc_1),
        "M": (f'=IF(K{r}="Válida",IF(L{r}<={P_TOL},"Sim","Não"),"")', f_calc),
        "O": (f'=IF(K{r}="Válida",A{r},NA())', f_calc_dh),
        "P": (f'=IF(K{r}="Válida",C{r},NA())', f_calc_1),
        "Q": (f'=IF(K{r}="Válida",I{r},NA())', f_calc_1),
        "R": (f'=IF(K{r}="Válida",J{r},NA())', f_calc_1),
        "S": (f'=IF(K{r}="Válida",C{r},"")', f_calc_1),
        "T": (f'=IF(K{r}="Válida",I{r},"")', f_calc_1),
        "U": (f'=IF(K{r}="Válida",J{r},"")', f_calc_1),
    }
    for k, col in enumerate("DEFG", 1):
        f[col] = (f'=IF(A{r}="","",IFERROR(AVERAGEIFS({PB},{PE},N{r},{PF},{k},{PG},1),""))', f_calc_1)
    for col, (formula, fm) in f.items():
        wc.write_formula(f"{col}{r}", formula, fm)

# ======================================================================= Painel
J = lambda col: f"Correlacao!${col}$2:${col}${NM}"  # noqa: E731
wp.hide_gridlines(2)
wp.set_zoom(90)
wp.set_landscape()
wp.set_paper(9)
wp.fit_to_pages(1, 0)
for col, w in zip("ABCDEFGHIJ", [2, 34, 20, 2, 42, 13, 2, 36, 13, 2]):
    wp.set_column(f"{col}:{col}", w)
wp.set_row(0, 46)
for c in range(10):
    wp.write_blank(0, c, None, f_banda)
    wp.write_blank(1, c, None, f_amarelo)
wp.write(0, 1, "Correlação prensa automática x laboratório", f_banda_tit)
wp.set_row(1, 4)
if os.path.exists(LOGO):
    wp.insert_image("H1", LOGO, {"x_scale": 0.105, "y_scale": 0.105, "x_offset": 170, "y_offset": 14,
                                 "object_position": 3})
wp.write("B3", "Composto de 2 h da prensa (média dos cortes de 12:00 a 13:30) comparado com o resultado do MES das 13:30, "
         "e assim por diante. Dados buscados no PIMS pelo suplemento Aspen, como no Relatório Físico.", f_status)
wp.write("B5", "Início do período", f_rotulo)
wp.write_datetime("C5", INICIO_PRENSA, f_ent_dh)
nome("pnlInicio", wp, "$C$5")
wp.write("B6", "Fim do período (em branco = agora)", f_rotulo)
wp.write_blank("C6", None, f_ent_dh)
nome("pnlFim", wp, "$C$6")
wp.set_row(4, 24)
wp.set_row(5, 24)
botoes = [("E5", 0, "AtualizarPeriodo", "Atualizar período"),
          ("E5", 160, "AtualizarDesdeInicio", "Desde o início da prensa"),
          ("H5", 40, "LimparDados", "Limpar dados")]
for cel, dx, macro, rotulo in botoes:
    wp.insert_button(cel, {"macro": macro, "caption": rotulo, "width": 150, "height": 44, "x_offset": dx, "y_offset": 2})
wp.merge_range("B8:I8", "Sem dados. Use um dos botões de atualização.", f_status)
nome("pnlStatus", wp, "$B$8")

L0 = 10   # linha (1-based) do título dos cartões
indicadores = [
    ("B", "Período analisado", [
        ("Primeiro composto", f'=IF(COUNT({J("A")})=0,"",MIN({J("A")}))', "dd/mm/yyyy hh:mm", "ini"),
        ("Último composto", f'=IF(COUNT({J("A")})=0,"",MAX({J("A")}))', "dd/mm/yyyy hh:mm", "fim"),
        ("Resultados do MES", f'=COUNT({J("A")})', "0", "nmes"),
        ("Janelas válidas", f'=COUNTIF({J("K")},"Válida")', "0", "val"),
        ("Janelas com os 4 cortes", f'=COUNTIFS({J("K")},"Válida",{J("H")},4)', "0", None),
        ("Resultados da prensa", f'=COUNT(Prensa!$B$2:$B${NP})', "0", None),
    ]),
    ("E", "Concordância", [
        ("Média do laboratório (kgf/pelota)", f'=IFERROR(AVERAGEIFS({J("C")},{J("K")},"Válida"),"")', "0.0", None),
        ("Média do composto da prensa (kgf/pelota)", f'=IFERROR(AVERAGEIFS({J("I")},{J("K")},"Válida"),"")', "0.0", None),
        ("Viés médio, prensa menos laboratório", f'=IFERROR(AVERAGE({J("U")}),"")', "0.0", "vies"),
        ("Desvio padrão das diferenças", f'=IFERROR(STDEV({J("U")}),"")', "0.0", "dp"),
        ("Erro médio absoluto", f'=IFERROR(AVERAGE({J("L")}),"")', "0.0", None),
        ("Janelas dentro da tolerância", f'=IFERROR(COUNTIF({J("M")},"Sim")/COUNTIF({J("K")},"Válida"),"")', "0%", None),
    ]),
    ("H", "Correlação", [
        ("Coeficiente de correlação (r)", f'=IFERROR(CORREL({J("S")},{J("T")}),"")', "0.00", "r"),
        ("Coeficiente de determinação (R²)", "R2", "0.00", "r2"),
        ("Inclinação (prensa = a + b x lab)", f'=IFERROR(SLOPE({J("T")},{J("S")}),"")', "0.00", None),
        ("Intercepto a", f'=IFERROR(INTERCEPT({J("T")},{J("S")}),"")', "0.0", None),
        ("Limite inferior de concordância", "LI", "0.0", "li"),
        ("Limite superior de concordância", "LS", "0.0", "ls"),
    ]),
]
ref = {}
for col, _, linhas in indicadores:
    colv = chr(ord(col) + 1)
    for i, (_, _, _, chave) in enumerate(linhas):
        if chave:
            ref[chave] = f"{colv}{L0 + 1 + i}"
especiais = {"R2": f'=IF({ref["r"]}="","",{ref["r"]}^2)',
             "LI": f'=IF({ref["dp"]}="","",{ref["vies"]}-1.96*{ref["dp"]})',
             "LS": f'=IF({ref["dp"]}="","",{ref["vies"]}+1.96*{ref["dp"]})'}
for col, titulo, linhas in indicadores:
    colv = chr(ord(col) + 1)
    wp.write(f"{col}{L0}", titulo, f_cartao_tit)
    wp.write_blank(f"{colv}{L0}", None, f_cartao_tit)
    for i, (rot, form, num, _) in enumerate(linhas):
        lin = L0 + 1 + i
        wp.write(f"{col}{lin}", rot, f_cartao)
        wp.write_formula(f"{colv}{lin}", especiais.get(form, form), valor_cartao(num))
        wp.set_row(lin - 1, 20)

LL = L0 + 8
wp.write(f"B{LL}", "Leitura", f_sec)
leitura = (f'=IF({ref["val"]}=0,"Sem janelas válidas no período. Atualize os dados pelos botões acima.",'
           f'IF({ref["val"]}<10,"Poucas janelas válidas para concluir (mínimo recomendado: 30). Até agora: "&{ref["val"]}&".",'
           f'IF({ref["r2"]}>=0.7,"Correlação forte",IF({ref["r2"]}>=0.4,"Correlação moderada","Correlação fraca"))'
           f'&" (R² = "&FIXED({ref["r2"]},2)&"). Em média a prensa mede "&FIXED(ABS({ref["vies"]}),1)&" kgf/pelota "'
           f'&IF({ref["vies"]}>=0,"acima","abaixo")&" do laboratório; 95% das janelas ficam entre "&FIXED({ref["li"]},1)'
           f'&" e "&FIXED({ref["ls"]},1)&"."))')
wp.merge_range(f"B{LL + 1}:I{LL + 1}", "", fmt(text_wrap=True, valign="top"))
wp.write_formula(f"B{LL + 1}", leitura, fmt(text_wrap=True, valign="top"))
wp.set_row(LL, 34)


def grafico(titulo, xcol, series, eixo_x, eixo_y, data=False, so_marcador=False, tendencia=False, faixa_y=None,
            faixa_x=None, legenda=True):
    ch = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers" if not so_marcador else "marker_only"})
    for col, rot, cor in series:
        s = {"name": rot,
             "categories": f"=Correlacao!${xcol}$2:${xcol}${NM}",
             "values": f"=Correlacao!${col}$2:${col}${NM}",
             "marker": {"type": "circle", "size": 5, "fill": {"color": cor}, "border": {"color": cor}}}
        if so_marcador:
            s["line"] = {"none": True}
        else:
            s["line"] = {"color": cor, "width": 1.75}
        if tendencia:
            s["trendline"] = {"type": "linear",
                              "line": {"color": LARANJA, "width": 1.25, "dash_type": "dash"}}
        ch.add_series(s)
    ch.set_title({"name": titulo, "name_font": {"name": F, "size": 12, "bold": True, "color": AZUL}})
    xa = {"name": eixo_x, "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 9},
          "major_gridlines": {"visible": False}}
    ya = {"name": eixo_y, "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 9},
          "major_gridlines": {"visible": True, "line": {"color": "#E3E8EC"}}}
    xa["label_position"] = "low"
    if data:
        xa["num_format"] = "dd/mm hh:mm"
        xa["date_axis"] = True
    if faixa_y:
        ya["min"], ya["max"] = faixa_y
    if faixa_x:
        xa["min"], xa["max"] = faixa_x
    ch.set_x_axis(xa)
    ch.set_y_axis(ya)
    ch.set_legend({"position": "bottom", "font": {"name": F, "size": 9}} if legenda else {"none": True})
    ch.show_hidden_data()     # as colunas de apoio ficam ocultas na aba Correlacao
    ch.set_chartarea({"border": {"color": "#C9D3DB"}})
    return ch


G0 = LL + 3
c1 = grafico("Composto de 2 h: prensa automática x laboratório", "O",
             [("Q", "Prensa (composto)", AZUL_MEDIO), ("P", "Laboratório (MES)", AZUL)],
             "Horário do resultado do MES", "kgf/pelota", data=True, faixa_y=(150, 450))
wp.insert_chart(f"B{G0}", c1, {"x_scale": 2.4, "y_scale": 1.1})
c2 = grafico("Correlação: laboratório x prensa", "P", [("Q", "Janelas válidas", AZUL_MEDIO)],
             "Laboratório (kgf/pelota)", "Prensa, composto (kgf/pelota)", so_marcador=True, tendencia=True,
             faixa_y=(150, 450), faixa_x=(150, 450), legenda=False)
wp.insert_chart(f"B{G0 + 18}", c2, {"x_scale": 1.1, "y_scale": 1.25})
c3 = grafico("Diferença prensa menos laboratório por janela", "O", [("R", "Diferença", LARANJA)],
             "Horário do resultado do MES", "kgf/pelota", data=True, so_marcador=True, legenda=False)
wp.insert_chart(f"B{G0 + 18}", c3, {"x_scale": 1.25, "y_scale": 1.25, "x_offset": 550})

# Resumo diário
R0 = G0 + 37
wp.write(f"B{R0}", "Resumo por dia", f_sec)
wp.write_row(f"B{R0 + 1}", ["Dia", "Janelas válidas"], f_cab)
wp.write_row(f"E{R0 + 1}", ["Média laboratório x média prensa (kgf/pelota)", "Viés médio"], f_cab)
for k in range(31):
    r = R0 + 2 + k
    dia = f'B{r}'
    wp.write_formula(dia, f'=IF(OR({ref["ini"]}="",INT({ref["ini"]})+{k}>{ref["fim"]}),"",INT({ref["ini"]})+{k})',
                     fmt(num_format="dd/mm/yyyy"))
    crit = f'{J("A")},">="&{dia},{J("A")},"<"&{dia}+1'
    wp.write_formula(f"C{r}", f'=IF({dia}="","",COUNTIFS({J("K")},"Válida",{crit}))', fmt())
    wp.write_formula(f"E{r}", f'=IF(OR({dia}="",C{r}=0),"",FIXED(AVERAGEIFS({J("C")},{J("K")},"Válida",{crit}),1)'
                              f'&" x "&FIXED(AVERAGEIFS({J("I")},{J("K")},"Válida",{crit}),1))',
                     fmt())
    wp.write_formula(f"F{r}", f'=IF(OR({dia}="",C{r}=0),"",AVERAGEIFS({J("U")},{crit}))',
                     fmt(num_format="0.0"))
wp.conditional_format(f"B{R0 + 2}:F{R0 + 32}", {"type": "formula", "criteria": f"=$B{R0 + 2}<>\"\"",
                                                 "format": wb.add_format({"bg_color": "#EEF3F7"})})
wp.activate()

# ======================================================================= Leia-me
wl.hide_gridlines(2)
wl.set_column("A:A", 150)
textos = [
    ("Correlação prensa automática x laboratório (MES)", f_tit),
    ("", None),
    ("A planilha sai sem dados. Na aba Painel, use o botão Desde o início da prensa ou informe o período e use Atualizar período.",
     f_alerta),
    ("", None),
    ("Como funciona a busca", f_sec),
    ("1. Mesmo mecanismo do Relatório Físico: a macro escreve a fórmula GetCalculationValues do suplemento Aspen Process Data "
     "na aba oculta Dados_PIMS, espera o PIMS responder, lê os valores e grava nas abas Prensa e MES.", None),
    ("2. Pré-requisito: o suplemento Aspen Process Data ativo no Excel (o mesmo usado pelo Relatório Físico) e a planilha aberta "
     "com as macros habilitadas.", None),
    ("3. Prensa automática (307-GI-001): leitura a cada 1 minuto. Cada resultado novo da tag vira uma linha da aba Prensa, "
     "com o horário em que apareceu no PIMS. A busca começa 1 h antes do período, para completar a primeira janela.", None),
    ("4. Laboratório (M4710050020-0031-HHLFU, Resistência à Compressão US4): consulta em intervalos de 30 min, um resultado por "
     "composto de 2 h.", None),
    ("5. Períodos longos são divididos em várias consultas, automaticamente. Cada atualização substitui os dados das abas Prensa e MES.",
     None),
    ("6. Tags, servidor, formato da data e tempo de espera ficam na aba Configurações. Se o PIMS recusar uma tag, a macro avisa "
     "qual foi.", None),
    ("", None),
    ("Regra de comparação (aba Configurações)", f_sec),
    ("1. O resultado da prensa entra no PIMS 20 min depois do corte. A planilha desconta essa defasagem: 13:20 vira corte das 13:00, "
     "13:50 vira corte das 13:30. Depois o horário é arredondado para a meia hora mais próxima.", None),
    ("2. Os cortes de 12:00, 12:30, 13:00 e 13:30 formam um composto de 2 h, comparado com o resultado do MES das 13:30. Depois "
     "14:00 a 15:30 com o MES das 15:30, e assim por diante.", None),
    ("3. O composto da prensa é a média dos cortes que tiveram resultado (1 a 4). Valores zero não entram.", None),
    ("4. Resultado do MES fora do horário do 4º corte aparece como Horário do MES fora do padrão na aba Correlacao. Nesse caso, "
     "ajuste o parâmetro Ajuste do horário do MES.", None),
    ("", None),
    ("Abas", f_sec),
    ("Painel: período, botões de atualização, indicadores, gráficos e resumo por dia.", None),
    ("Configurações: conexão com o PIMS, tags e regra de comparação (células amarelas).", None),
    ("Prensa e MES: dados buscados no PIMS (colunas A e B) e cálculo do corte e da janela (cinza).", None),
    ("Correlacao: uma linha por resultado do MES, com os 4 cortes da prensa, o composto, a diferença e a situação.", None),
    ("Dados_PIMS (oculta): área de trabalho das consultas ao Aspen. Para ver, use a macro MostrarDadosPIMS (Alt+F8).", None),
    ("", None),
    ("Se a busca não funcionar", f_sec),
    ("1. Abra o Relatório Físico na mesma máquina e confirme que a atualização do MES funciona (suplemento ativo).", None),
    ("2. Confira na aba Configurações o servidor (UBU), as tags e o formato da data. Use o mesmo formato configurado no Relatório Físico.",
     None),
    ("3. Se a função aparecer como #NOME?, o suplemento não está ativo: Arquivo > Opções > Suplementos > Suplementos COM.", None),
    ("4. Se o Excel bloquear as macros, clique com o botão direito no arquivo > Propriedades > Desbloquear, e abra de novo.", None),
    ("", None),
    ("Capacidade: 5.000 resultados da prensa e 1.500 compostos do MES (cerca de 3 meses).", None),
]
for i, (t, fm) in enumerate(textos):
    wl.write(i, 0, t, fm or f_txt)

# ======================================================================= macros
if COM_VBA:
    modulo = open(os.path.join(AQUI, "vba", "ModPIMS.bas"), encoding="utf-8").read().replace("\r\n", "\n").replace("\n", "\r\n")
    if os.environ.get("VBA_TESTE"):    # rotina extra só para a conferência automática da macro
        modulo += open(os.environ["VBA_TESTE"], encoding="utf-8").read().replace("\n", "\r\n")
    docs = {"ThisWorkbook": "Option Explicit\r\n\r\nPrivate Sub Workbook_Open()\r\n    shPainel.Activate\r\nEnd Sub\r\n"}
    docs.update({cn: "" for cn in CODENOMES.values()})
    binario = os.path.join(AQUI, "vbaProject.bin")
    open(binario, "wb").write(vba_project_bin({"ModPIMS": modulo}, docs))
    wb.add_vba_project(binario)
    wb.set_vba_name("ThisWorkbook")
    for ws, cn in CODENOMES.items():
        ws.set_vba_name(cn)
wb.close()
if COM_VBA:
    os.remove(binario)
print(SAIDA, "com macros" if COM_VBA else "sem macros", f"| prensa {len(prensa)} | mes {len(mes)}")
