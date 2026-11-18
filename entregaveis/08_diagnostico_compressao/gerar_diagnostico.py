"""Gera a planilha com macros de diagnóstico da compressão online x laboratório.

Uso: python gerar_diagnostico.py [saida.xlsm] [--teste dados.json]

Tags (as mesmas da tela Compressão Online Usina 4):
  307-GI-001               compressão online, média de cada ensaio (a cada 30 min)
  307-GI-002               desvio padrão das pelotas de cada ensaio
  M4710050020-0031-HHLFU   referência do laboratório (composto de 2 h, MES)
  U04-07CR001 e 07CR002    modo dos cortadores da amostragem (Automático = 32)

Regra (definida pelo LCE): o resultado do laboratório das 13:30 corresponde aos cortes da compressão
online de 12:00, 12:30, 13:00 e 13:30; o das 15:30 aos cortes de 14:00 a 15:30, e assim por diante.
O composto online é a média dos cortes com resultado. O resultado da prensa entra no PIMS cerca de
20 min depois do corte (13:20 é o corte das 13:00); a defasagem fica na aba Configurações.

A busca segue o Relatório Físico (fórmula GetCalculationValues do suplemento Aspen, escrita pela macro
vba/ModPIMS.bas). Testes, Diagnóstico e Painel são fórmulas. A planilha sai SEM DADOS. Com --teste,
grava dados de um JSON (só para conferir as fórmulas; gera .xlsx, sem macros).
"""
import json
import os
import sys
from datetime import datetime

import xlsxwriter

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "ferramentas"))
from vba_projeto import vba_project_bin  # noqa: E402

LOGO = os.path.join(AQUI, "..", "assets", "samarco-logo-branca.png")
NP = 5001       # última linha da aba Prensa (5.000 resultados, cerca de 100 dias)
NM = 1501       # última linha das abas MES e Testes (1.500 compostos, cerca de 125 dias)
NC = 10001      # última linha da aba Cortadores (10.000 leituras de cada cortador)
INICIO_PRENSA = datetime(2026, 10, 28, 0, 0)

AZUL, AMARELO, AZUL_MEDIO, LARANJA, VERDE = "#00335A", "#FFC000", "#0091D5", "#F37021", "#2E9E5B"
F = "Arial"

args = list(sys.argv[1:])
teste = None
if "--teste" in args:
    i = args.index("--teste")
    teste = json.load(open(args[i + 1]))
    del args[i:i + 2]
SAIDA = args[0] if args else os.path.join(AQUI, "diagnostico_compressao_online.xlsm")
COM_VBA = SAIDA.lower().endswith(".xlsm")

wb = xlsxwriter.Workbook(SAIDA)
wb.set_properties({"title": "Diagnóstico compressão online x laboratório", "company": "LCE"})


def fmt(**k):
    base = {"font_name": F, "font_size": 10, "valign": "vcenter"}
    base.update(k)
    return wb.add_format(base)


f_cab = fmt(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, border=1, border_color="#C9D3DB")
f_ent = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB")
f_ent_dh = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB", num_format="dd/mm/yyyy hh:mm")
f_dado_dh = fmt(num_format="dd/mm/yyyy hh:mm")
f_dado_n = fmt(num_format="0.0")
f_dado_0 = fmt(num_format="0")
f_calc = fmt(bg_color="#F2F2F2")
f_calc_dh = fmt(bg_color="#F2F2F2", num_format="dd/mm/yyyy hh:mm")
f_calc_1 = fmt(bg_color="#F2F2F2", num_format="0.0")
f_calc_pct = fmt(bg_color="#F2F2F2", num_format="0.0%")
f_txt = fmt(text_wrap=True, valign="top")
f_sec = fmt(bold=True, font_size=12, font_color=AZUL)
f_tit = fmt(bold=True, font_size=15, font_color=AZUL)
f_alerta = fmt(bold=True, font_color="#C00000", text_wrap=True)
f_banda = fmt(bg_color=AZUL)
f_banda_tit = fmt(bold=True, font_size=18, font_color="white", bg_color=AZUL, indent=1)
f_amarelo = fmt(bg_color=AMARELO)
f_cartao = fmt(bg_color="#EEF3F7")
f_status = fmt(italic=True, font_color="#5B6B77")
f_rotulo = fmt(bold=True, font_color=AZUL)

# ======================================================================= abas (na ordem de exibição)
wp = wb.add_worksheet("Painel")
wt = wb.add_worksheet("Testes")
wdg = wb.add_worksheet("Diagnóstico")
wl = wb.add_worksheet("Leia-me")
wcfg = wb.add_worksheet("Configurações")
wpr = wb.add_worksheet("Prensa")
wm = wb.add_worksheet("MES")
wco = wb.add_worksheet("Cortadores")
wlog = wb.add_worksheet("Log")
wd = wb.add_worksheet("Dados_PIMS")
CODENOMES = {wp: "shPainel", wt: "shTestes", wdg: "shDiagnostico", wl: "shLeiame", wcfg: "shConfig", wpr: "shPrensa",
             wm: "shMES", wco: "shCortadores", wlog: "shLog", wd: "shDados"}


def nome(n, ws, cel):
    wb.define_name(n, f"='{ws.get_name()}'!{cel}")


# ======================================================================= Configurações
wcfg.hide_gridlines(2)
wcfg.set_column("A:A", 52)
wcfg.set_column("B:B", 28)
wcfg.set_column("C:C", 100)
wcfg.write_row(0, 0, ["Parâmetro", "Valor", "Observação"], f_cab)
wcfg.set_row(0, 24)


def consulta(rotulo, sufixo, tag, mapa, periodo, calculo, obs_tag, obs_periodo, obs_calc):
    return [(rotulo, None, None, None, None),
            ("Tag", tag, "cfgTag" + sufixo, obs_tag, None),
            ("Tipo (mapa)", mapa, "cfgMapa" + sufixo, "", None),
            ("Intervalo da consulta", periodo, "cfgPeriodo" + sufixo, obs_periodo, None),
            ("Tipo de cálculo do suplemento", calculo, "cfgCalculo" + sufixo, obs_calc, "texto"),
            ("", None, None, None, None)]


CALC0 = "0 = valor no instante (mesmo código da consulta de ritmo de processo do Relatório Físico)."
linhas_cfg = [
    ("Conexão com o PIMS (Aspen Process Data no Excel)", None, None, None, None),
    ("Servidor do PIMS", "UBU", "cfgServidor", "Mesmo servidor do Relatório Físico.", None),
    ("Formato da data para o Aspen", "Texto dd/mm/aaaa", "cfgFormatoData",
     "Padrão do Relatório Físico. Opções: Texto dd/mm/aaaa, Texto mm/dd/aaaa ou Data do Excel.", "lista"),
    ("Espera máxima pela resposta do PIMS (s)", 300, "cfgTimeout", "Tempo máximo de cada consulta (o Relatório Físico leva cerca de 2 min).", None),
    ("Início da fórmula do suplemento", "=GetCalculationValues(", "cfgPrefixo",
     "Como o Excel mostra a função do Aspen. Se o Excel não reconhecer, a macro tenta o nome completo do suplemento.", "texto"),
    ("Início de operação da prensa", INICIO_PRENSA, "cfgInicioPrensa",
     "Usado pelo botão Desde o início da prensa. Antes desta data a tag não tem resultados válidos.", "dh"),
    ("", None, None, None, None),
] + consulta("Compressão online (obrigatória)", "Prensa", "307-GI-001", "IP_ANALOGMAP", "1m", "0",
             "Média de compressão de cada ensaio (kgf/pelota).",
             "Leitura a cada 1 minuto. Cada resultado novo vira uma linha da aba Prensa, com o horário em que entrou no PIMS.", CALC0) \
  + consulta("Referência do laboratório (obrigatória)", "Lab", "M4710050020-0031-HHLFU", "IP_MESVALOR", "30m", "1",
             "Resistência à Compressão, pelota queimada US4, em kgf/pelota (mesma tag da tela).",
             "Intervalos de 30 min: o horário do resultado no MES fica identificado (13:30, 15:30 e assim por diante).",
             "1 = mesmo código da consulta de qualidade do Relatório Físico.") \
  + consulta("Desvio padrão da compressão online (complementar)", "Desvio", "307-GI-002", "IP_ANALOGMAP", "1m", "0",
             "Desvio padrão das pelotas de cada ensaio. Se a consulta falhar, a análise sai sem ele.", "Mesma leitura da média.", CALC0) \
  + consulta("Cortador 07CR001 (complementar)", "Cort1", "U04-07CR001", "IP_PIDMAP03", "10m", "0",
             "Modo do cortador, como nas lâmpadas da tela. Se a consulta falhar, a análise sai sem a situação dos cortadores.",
             "Leitura a cada 10 minutos.", CALC0) \
  + consulta("Cortador 07CR002 (complementar)", "Cort2", "U04-07CR002", "IP_PIDMAP03", "10m", "0",
             "Modo do cortador, como nas lâmpadas da tela.", "Leitura a cada 10 minutos.", CALC0) \
  + [
    ("Regra de comparação", None, None, None, None),
    ("Defasagem entre o corte e a entrada no PIMS (min)", 20, "cfgDefasagem",
     "Informado pelo LCE: o corte das 13:00 entra no PIMS às 13:20. Use 0 se o horário da tag já for o do corte.", None),
    ("Deslocamento das janelas (meias horas)", 0, "cfgDeslocamento",
     "0 = cortes de 12:00 a 13:30 com o laboratório das 13:30, 14:00 a 15:30 com o das 15:30 e assim por diante.", None),
    ("Mínimo de cortes para a janela ser válida", 1, "cfgMinCortes",
     "1 = usa a média mesmo com um só corte. Use 4 para considerar só compostos completos.", None),
    ("Ajuste do horário do MES (min)", 0, "cfgAjusteMES",
     "Use se o horário do resultado no MES não for o do último corte (13:30). Exemplo: -30 se o MES gravar 14:00.", None),
    ("Valor do modo Automático dos cortadores", 32, "cfgModoAuto", "Mesmo valor usado nas lâmpadas da tela (Manual = 16).", None),
    ("Excluir janelas com cortador fora do automático (1 = sim, 0 = não)", 1, "cfgExcluirManual",
     "Com 1, a janela em que algum cortador saiu do automático não entra na estatística (a amostra da prensa pode não "
     "representar o composto). Ela continua listada na aba Testes.", None),
    ("Tolerância da diferença online x laboratório (kgf/pelota)", 20, "cfgTolerancia",
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
        wcfg.data_validation(i, 1, i, 1, {"validate": "list", "source": ["Texto dd/mm/aaaa", "Texto mm/dd/aaaa", "Data do Excel"]})
    wcfg.write(i, 2, obs, fmt(text_wrap=True))
    nome(n, wcfg, f"$B${i + 1}")
P_DEF, P_DESL, P_MIN, P_AJ, P_TOL = "cfgDefasagem", "cfgDeslocamento", "cfgMinCortes", "cfgAjusteMES", "cfgTolerancia"

# ======================================================================= Dados_PIMS (oculta) e Log
wd.write("A1", "Consultas ao PIMS (Aspen Process Data). Preenchida pela macro, não editar.", f_rotulo)
wd.write("A3", "Início")
wd.write_string("B3", "")
wd.write("A4", "Fim")
wd.write_string("B4", "")
for k, (col, titulo) in enumerate(zip("AEIMQ", ["compressão online", "laboratório (MES)", "desvio padrão",
                                                 "cortador 07CR001", "cortador 07CR002"])):
    c = 4 * k
    wd.set_column(c, c, 19)
    wd.set_column(c + 1, c + 1, 14)
    wd.write(5, c, "Consulta: " + titulo, f_rotulo)
    wd.write(6, c, "Sem consulta")
    wd.write_row(7, c, ["Data/hora", "Valor"], f_cab)
wd.hide()

wlog.set_column("A:A", 19)
wlog.set_column("B:B", 22)
wlog.set_column("C:C", 160)
wlog.write("A1", "Registro da última busca ou do último teste de conexão (preenchido pela macro). Se algo não funcionar, envie esta aba.", f_rotulo)
wlog.write_row("A2", ["Data e hora", "Etapa", "Detalhe"], f_cab)
for r in range(2, 1000):
    wlog.set_row(r, None, fmt(num_format="dd/mm/yyyy hh:mm:ss"))
wlog.freeze_panes(2, 0)

# ======================================================================= Prensa
cab_pr = ["Data e hora no PIMS", "Média 307-GI-001 (kgf/pelota)", "Desvio padrão 307-GI-002 (kgf/pelota)", "Hora do corte",
          "Índice de meia hora", "Janela", "Posição do corte na janela (1 a 4)", "Válido (1 = entra na média)",
          "Horário do laboratório correspondente"]
wpr.write_row(0, 0, cab_pr, f_cab)
wpr.set_row(0, 42)
for k, w in enumerate([19, 14, 14, 17, 11, 10, 12, 12, 17]):
    wpr.set_column(k, k, w)
wpr.freeze_panes(1, 0)
wpr.write_comment("A1", "Colunas A a C: preenchidas pela macro (um resultado por linha, com o horário em que entrou no PIMS). "
                        "Também podem ser coladas à mão.")
prensa = teste["prensa"] if teste else []
for r in range(2, NP + 1):
    i = r - 2
    if i < len(prensa):
        wpr.write_datetime(r - 1, 0, datetime.fromisoformat(prensa[i][0]), f_dado_dh)
        wpr.write_number(r - 1, 1, prensa[i][1], f_dado_n)
        if prensa[i][2] is not None:
            wpr.write_number(r - 1, 2, prensa[i][2], f_dado_n)
    else:
        wpr.write_blank(r - 1, 0, None, f_dado_dh)
    wpr.write_formula(r - 1, 4, f'=IF(A{r}="","",ROUND((A{r}-{P_DEF}/1440)*48,0))', f_calc)
    wpr.write_formula(r - 1, 3, f'=IF(E{r}="","",E{r}/48)', f_calc_dh)
    wpr.write_formula(r - 1, 5, f'=IF(E{r}="","",INT((E{r}-{P_DESL})/4))', f_calc)
    wpr.write_formula(r - 1, 6, f'=IF(E{r}="","",MOD(E{r}-{P_DESL},4)+1)', f_calc)
    wpr.write_formula(r - 1, 7, f'=IF(A{r}="","",IF(AND(ISNUMBER(B{r}),B{r}<>0),1,0))', f_calc)
    wpr.write_formula(r - 1, 8, f'=IF(F{r}="","",(F{r}*4+3+{P_DESL})/48)', f_calc_dh)

# ======================================================================= MES
wm.write_row(0, 0, ["Data e hora do resultado no MES", "Compressão laboratório (kgf/pelota)", "Índice de meia hora",
                    "Janela", "Posição (esperado 4)"], f_cab)
wm.set_row(0, 42)
for k, w in enumerate([21, 16, 11, 10, 12]):
    wm.set_column(k, k, w)
wm.freeze_panes(1, 0)
wm.write_comment("A1", "Colunas A e B: preenchidas pela macro, um resultado por composto de 2 h. Também podem ser coladas à mão.")
mes = teste["mes"] if teste else []
for r in range(2, NM + 1):
    i = r - 2
    if i < len(mes):
        wm.write_datetime(r - 1, 0, datetime.fromisoformat(mes[i][0]), f_dado_dh)
        wm.write_number(r - 1, 1, mes[i][1], f_dado_0)
    else:
        wm.write_blank(r - 1, 0, None, f_dado_dh)
    wm.write_formula(r - 1, 2, f'=IF(A{r}="","",ROUND((A{r}+{P_AJ}/1440)*48,0))', f_calc)
    wm.write_formula(r - 1, 3, f'=IF(C{r}="","",INT((C{r}-{P_DESL})/4))', f_calc)
    wm.write_formula(r - 1, 4, f'=IF(C{r}="","",MOD(C{r}-{P_DESL},4)+1)', f_calc)

# ======================================================================= Cortadores
wco.write_row(0, 0, ["Data e hora 07CR001", "Modo 07CR001", "Janela", "Fora do automático (1 = sim)",
                     "Data e hora 07CR002", "Modo 07CR002", "Janela", "Fora do automático (1 = sim)"], f_cab)
wco.set_row(0, 42)
for k, w in enumerate([19, 10, 10, 12, 19, 10, 10, 12]):
    wco.set_column(k, k, w)
wco.freeze_panes(1, 0)
wco.write_comment("A1", "Colunas A, B, E e F: preenchidas pela macro (modo do cortador; Automático = 32).")
cort = teste.get("cortadores", [[], []]) if teste else [[], []]
for r in range(2, NC + 1):
    i = r - 2
    for base, serie in ((0, cort[0]), (4, cort[1])):
        ct, cv, cj, cf = (chr(65 + base + k) for k in range(4))
        if i < len(serie):
            wco.write_datetime(r - 1, base, datetime.fromisoformat(serie[i][0]), f_dado_dh)
            wco.write_number(r - 1, base + 1, serie[i][1], f_dado_0)
        else:
            wco.write_blank(r - 1, base, None, f_dado_dh)
        wco.write_formula(r - 1, base + 2, f'=IF({ct}{r}="","",INT((ROUND({ct}{r}*1440,0)/30-{P_DESL})/4))', f_calc)
        wco.write_formula(r - 1, base + 3, f'=IF(OR({ct}{r}="",NOT(ISNUMBER({cv}{r}))),"",IF({cv}{r}=cfgModoAuto,0,1))', f_calc)

# ======================================================================= Testes (uma linha por resultado do laboratório)
tit = ["Horário do laboratório", "Cortes da compressão online", "Laboratório (kgf/pelota)", "Corte 1", "Corte 2", "Corte 3",
       "Corte 4", "Cortes com resultado", "Composto online (kgf/pelota)", "Desvio entre os cortes (kgf/pelota)",
       "Desvio padrão médio dos ensaios (307-GI-002)", "Diferença online menos laboratório", "Diferença (%)",
       "Cortadores", "Situação", "Diferença absoluta", "Dentro da tolerância",
       "id janela", "X gráfico", "Lab gráfico", "Online gráfico", "Dif gráfico", "Média gráfico",
       "Lab válido", "Online válido", "Dif válida", "Média válida", "Leituras fora do automático", "Leituras dos cortadores",
       "Dif % válida"]
wt.write_row(0, 0, tit, f_cab)
wt.set_row(0, 54)
for k, w in enumerate([17, 15, 13, 9, 9, 9, 9, 10, 14, 12, 14, 14, 10, 17, 26, 11, 11] + [8] * 13):
    wt.set_column(k, k, w, None, {"hidden": k >= 17})
wt.freeze_panes(1, 3)
PF, PG, PH, PB, PC = (f"Prensa!${c}$2:${c}${NP}" for c in "FGHBC")
CC1, CF1, CC2, CF2 = (f"Cortadores!${c}$2:${c}${NC}" for c in "CDGH")
for r in range(2, NM + 1):
    f = {
        "A": (f'=IF(MES!A{r}="","",MES!C{r}/48)', f_calc_dh),
        "R": (f'=IF(A{r}="","",MES!D{r})', f_calc),
        "B": (f'=IF(A{r}="","",TEXT((R{r}*4+{P_DESL})/48,"hh:mm")&" a "&TEXT(A{r},"hh:mm"))', f_calc),
        "C": (f'=IF(A{r}="","",MES!B{r})', f_calc_1),
        "H": (f'=IF(A{r}="","",COUNT(D{r}:G{r}))', f_calc),
        "I": (f'=IF(A{r}="","",IF(H{r}=0,"",AVERAGE(D{r}:G{r})))', f_calc_1),
        "J": (f'=IF(A{r}="","",IF(H{r}<2,"",STDEV(D{r}:G{r})))', f_calc_1),
        "K": (f'=IF(A{r}="","",IFERROR(AVERAGEIFS({PC},{PF},R{r},{PH},1),""))', f_calc_1),
        "L": (f'=IF(OR(I{r}="",C{r}=""),"",I{r}-C{r})', f_calc_1),
        "M": (f'=IF(L{r}="","",L{r}/C{r})', f_calc_pct),
        "AB": (f'=IF(A{r}="","",COUNTIFS({CC1},R{r},{CF1},1)+COUNTIFS({CC2},R{r},{CF2},1))', f_calc),
        "AC": (f'=IF(A{r}="","",COUNTIF({CC1},R{r})+COUNTIF({CC2},R{r}))', f_calc),
        "N": (f'=IF(A{r}="","",IF(AC{r}=0,"Sem dado",IF(AB{r}>0,"Fora do automático","Automático")))', f_calc),
        "O": (f'=IF(A{r}="","",IF(MES!E{r}<>4,"Horário do MES fora do padrão",IF(H{r}<{P_MIN},"Cortes insuficientes",'
              f'IF(AND(cfgExcluirManual=1,N{r}="Fora do automático"),"Cortador fora do automático","Válida"))))', f_calc),
        "P": (f'=IF(O{r}="Válida",ABS(L{r}),"")', f_calc_1),
        "Q": (f'=IF(O{r}="Válida",IF(P{r}<={P_TOL},"Sim","Não"),"")', f_calc),
        "S": (f'=IF(O{r}="Válida",A{r},NA())', f_calc_dh),
        "T": (f'=IF(O{r}="Válida",C{r},NA())', f_calc_1),
        "U": (f'=IF(O{r}="Válida",I{r},NA())', f_calc_1),
        "V": (f'=IF(O{r}="Válida",L{r},NA())', f_calc_1),
        "W": (f'=IF(O{r}="Válida",(C{r}+I{r})/2,NA())', f_calc_1),
        "X": (f'=IF(O{r}="Válida",C{r},"")', f_calc_1),
        "Y": (f'=IF(O{r}="Válida",I{r},"")', f_calc_1),
        "Z": (f'=IF(O{r}="Válida",L{r},"")', f_calc_1),
        "AA": (f'=IF(O{r}="Válida",(C{r}+I{r})/2,"")', f_calc_1),
        "AD": (f'=IF(O{r}="Válida",M{r},"")', f_calc_pct),
    }
    for k, col in enumerate("DEFG", 1):
        f[col] = (f'=IF(A{r}="","",IFERROR(AVERAGEIFS({PB},{PF},R{r},{PG},{k},{PH},1),""))', f_calc_1)
    for col, (formula, fm) in f.items():
        wt.write_formula(f"{col}{r}", formula, fm)
wt.conditional_format(f"O2:O{NM}", {"type": "text", "criteria": "begins with", "value": "Válida",
                                    "format": wb.add_format({"font_color": VERDE, "bold": True})})
wt.conditional_format(f"N2:N{NM}", {"type": "text", "criteria": "begins with", "value": "Fora",
                                    "format": wb.add_format({"font_color": LARANJA, "bold": True})})
for n_, col in (("tA", "A"), ("tSit", "O"), ("tLab", "X"), ("tOn", "Y"), ("tDif", "Z"), ("tMed", "AA"), ("tAbs", "P"),
                ("tTol", "Q"), ("tDifP", "AD"), ("tDPc", "J"), ("tDP2", "K"), ("tCort", "N")):
    wb.define_name(n_, f"=Testes!${col}$2:${col}${NM}")

# ======================================================================= Diagnóstico
wdg.hide_gridlines(2)
for k, w in enumerate([2, 44, 16, 70, 2, 16, 16, 16, 16, 16]):
    wdg.set_column(k, k, w)
wdg.write("B1", "Diagnóstico estatístico: compressão online x laboratório", f_tit)
wdg.write("B2", "Só entram as janelas com situação Válida na aba Testes (composto online e laboratório da mesma janela de 2 h). "
          "Diferença = composto online menos laboratório.", f_status)
wdg.write_row("B4", ["Indicador", "Valor", "Leitura"], f_cab)
N_ = "C6"
ind = [
    ("n", "Janelas válidas (pares online x laboratório)", "=COUNT(tDif)", "0",
     f'=IF({N_}<30,"Amostra ainda pequena: conclusões preliminares (recomendado 30 ou mais).","Amostra suficiente para conclusões.")'),
    ("tot", "Resultados do laboratório no período", "=COUNT(tA)", "0", '=""'),
    ("exc", "Excluídas por cortador fora do automático", '=COUNTIF(tSit,"Cortador fora do automático")', "0",
     '=IF(cfgExcluirManual=1,"Fora da estatística (aba Configurações).","Exclusão desligada: entram na estatística.")'),
    ("ins", "Excluídas por cortes insuficientes", '=COUNTIF(tSit,"Cortes insuficientes")', "0", '=""'),
    ("fpd", "Horário do MES fora do padrão", '=COUNTIF(tSit,"Horário do MES fora do padrão")', "0",
     '=IF(C10>0,"Conferir o Ajuste do horário do MES (aba Configurações).","")'),
    (None, "Médias", None, None, None),
    ("mlab", "Média do laboratório (kgf/pelota)", "=IF(C6=0,\"\",AVERAGE(tLab))", "0.0", '=""'),
    ("mon", "Média do composto online (kgf/pelota)", "=IF(C6=0,\"\",AVERAGE(tOn))", "0.0", '=""'),
    ("dpc", "Desvio médio entre os 4 cortes (kgf/pelota)", '=IFERROR(AVERAGEIFS(tDPc,tSit,"Válida"),"")', "0.0",
     '=IF(C14="","","Variação típica da compressão online dentro de uma mesma janela de 2 h.")'),
    ("dp2", "Desvio padrão médio dos ensaios, 307-GI-002 (kgf/pelota)", '=IFERROR(AVERAGEIFS(tDP2,tSit,"Válida"),"")', "0.0",
     '=IF(C15="","","Dispersão média entre as pelotas de cada ensaio da prensa.")'),
    (None, "Viés (diferença média)", None, None, None),
    ("vies", "Viés médio, online menos laboratório (kgf/pelota)", "=IF(C6<2,\"\",AVERAGE(tDif))", "0.0",
     '=IF(C17="","",IF(C17>=0,"A compressão online mede em média "&FIXED(C17,1)&" kgf/pelota acima do laboratório.",'
     '"A compressão online mede em média "&FIXED(-C17,1)&" kgf/pelota abaixo do laboratório."))'),
    ("viesp", "Viés médio relativo (%)", "=IF(C6<2,\"\",AVERAGE(tDifP))", "0.0%", '=""'),
    ("dpd", "Desvio padrão das diferenças (kgf/pelota)", "=IF(C6<2,\"\",STDEV(tDif))", "0.0", '=""'),
    ("icl", "IC 95% do viés: limite inferior", "=IF(C6<2,\"\",C17-TINV(0.05,C6-1)*C19/SQRT(C6))", "0.0", '=""'),
    ("ich", "IC 95% do viés: limite superior", "=IF(C6<2,\"\",C17+TINV(0.05,C6-1)*C19/SQRT(C6))", "0.0", '=""'),
    ("p", "Teste t pareado: valor-p", "=IF(OR(C6<2,C19=0),\"\",TDIST(ABS(C17/(C19/SQRT(C6))),C6-1,2))", "0.000",
     '=IF(C22="","",IF(C22<0.05,"Viés significativo (valor-p < 0,05): a diferença não é acaso.",'
     '"Viés não significativo (valor-p >= 0,05): a diferença média pode ser acaso."))'),
    (None, "Concordância (Bland-Altman)", None, None, None),
    ("li", "Limite inferior de concordância (viés - 1,96 DP)", "=IF(C6<2,\"\",C17-1.96*C19)", "0.0",
     '=IF(C24="","","95% das diferenças devem ficar entre "&FIXED(C24,1)&" e "&FIXED(C25,1)&" kgf/pelota.")'),
    ("ls", "Limite superior de concordância (viés + 1,96 DP)", "=IF(C6<2,\"\",C17+1.96*C19)", "0.0", '=""'),
    ("mae", "Erro médio absoluto (kgf/pelota)", '=IFERROR(AVERAGEIFS(tAbs,tSit,"Válida"),"")', "0.0", '=""'),
    ("rmse", "Raiz do erro quadrático médio (kgf/pelota)", "=IF(C6<2,\"\",SQRT(SUMSQ(tDif)/C6))", "0.0", '=""'),
    ("tol", "Janelas dentro da tolerância", '=IF(C6=0,"",COUNTIF(tTol,"Sim")/C6)', "0%",
     '=IF(C28="","",FIXED(C28*100,0)&"% das janelas com diferença de até "&cfgTolerancia&" kgf/pelota.")'),
    ("ccc", "Coeficiente de concordância de Lin (CCC)",
     "=IF(C6<3,\"\",2*COVAR(tLab,tOn)/(VARP(tLab)+VARP(tOn)+(C12-C13)^2))", "0.000",
     '=IF(C29="","",IF(C29>0.99,"Concordância quase perfeita (acima de 0,99).",IF(C29>=0.95,"Concordância substancial (0,95 a 0,99).",'
     'IF(C29>=0.9,"Concordância moderada (0,90 a 0,95).","Concordância fraca (abaixo de 0,90)."))))'),
    (None, "Correlação e regressão (online = a + b x laboratório)", None, None, None),
    ("r", "Coeficiente de correlação (r)", "=IF(C6<3,\"\",CORREL(tLab,tOn))", "0.000",
     '=IF(C31="","",IF(ABS(C31)>=0.7,"Correlação forte.",IF(ABS(C31)>=0.4,"Correlação moderada.","Correlação fraca.")))'),
    ("r2", "Coeficiente de determinação (R²)", "=IF(C31=\"\",\"\",C31^2)", "0.000",
     '=IF(C32="","",FIXED(C32*100,0)&"% da variação do online acompanha a variação do laboratório.")'),
    ("b", "Inclinação b da regressão simples (informativa)", "=IF(C6<3,\"\",SLOPE(tOn,tLab))", "0.000",
     '=IF(C33="","","Fica abaixo de 1 quando o laboratório também tem erro de medição; não usar sozinha para concluir.")'),
    ("bd", "Inclinação de Deming (erro nos dois métodos)",
     "=IF(C6<3,\"\",(DEVSQ(tOn)-DEVSQ(tLab)+SQRT((DEVSQ(tOn)-DEVSQ(tLab))^2+4*(COVAR(tLab,tOn)*C6)^2))/(2*COVAR(tLab,tOn)*C6))",
     "0.000", '=IF(C34="","","Perto de 1: online e laboratório variam na mesma proporção.")'),
    ("bp", "Viés proporcional: inclinação das diferenças sobre a média", "=d_bas", "0.000",
     '=IF(C35="","","IC 95% de "&FIXED(d_bal,3)&" a "&FIXED(d_bah,3)&IF(AND(d_bal<=0,d_bah>=0),'
     '": inclui zero, sem indício de viés proporcional.",": não inclui zero, a diferença muda com o nível de compressão."))'),
    ("a", "Intercepto a da regressão simples (kgf/pelota)", "=IF(C6<3,\"\",INTERCEPT(tOn,tLab))", "0.0", '=""'),
]
linha_ref = {}
for k, (chave, rot, form, num, leitura) in enumerate(ind):
    r = 5 + k   # índice 0-based da linha => linha Excel r+1
    if chave is None:
        wdg.write(r, 1, rot, f_sec)
        continue
    linha_ref[chave] = f"C{r + 1}"
    wdg.write(r, 1, rot, f_cartao)
    wdg.write_formula(r, 2, form, fmt(bold=True, font_color=AZUL, bg_color="#EEF3F7", align="right", num_format=num))
    wdg.write_formula(r, 3, leitura, fmt(text_wrap=True, bg_color="#EEF3F7"))
# as fórmulas acima citam as células pela posição; confere que a tabela não mudou de lugar
assert linha_ref == {"n": "C6", "tot": "C7", "exc": "C8", "ins": "C9", "fpd": "C10", "mlab": "C12", "mon": "C13", "dpc": "C14",
                     "dp2": "C15", "vies": "C17", "viesp": "C18", "dpd": "C19", "icl": "C20", "ich": "C21", "p": "C22",
                     "li": "C24", "ls": "C25", "mae": "C26", "rmse": "C27", "tol": "C28", "ccc": "C29", "r": "C31",
                     "r2": "C32", "b": "C33", "bd": "C34", "bp": "C35", "a": "C36"}, linha_ref
R = linha_ref
for n_, cel in R.items():
    wb.define_name("d_" + n_, f"='Diagnóstico'!${cel[0]}${cel[1:]}")

# Conclusão automática
wdg.write("B38", "Conclusão", f_sec)
conclusao = [
    '=IF(d_n<3,"Ainda não há janelas válidas suficientes para o diagnóstico (mínimo 3). Atualize os dados no Painel.",'
    '"Foram comparadas "&d_n&" janelas de 2 h (composto online x laboratório)"&IF(d_n<30,", amostra ainda pequena: resultado preliminar.","."))',
    '=IF(d_n<3,"",IF(d_p<0.05,"Viés: a compressão online mede "&FIXED(ABS(d_vies),1)&" kgf/pelota "&IF(d_vies>=0,"acima","abaixo")'
    '&" do laboratório, diferença significativa (valor-p "&FIXED(d_p,3)&"; IC 95% de "&FIXED(d_icl,1)&" a "&FIXED(d_ich,1)&").",'
    '"Viés: diferença média de "&FIXED(d_vies,1)&" kgf/pelota, sem significância estatística (valor-p "&FIXED(d_p,3)&")."))',
    '=IF(d_n<3,"",IF(AND(d_bal<=0,d_bah>=0),"Sem viés proporcional: a diferença é parecida em compressões altas e baixas.",'
    '"Viés proporcional: a diferença muda com o nível de compressão ("&FIXED(d_bas,2)&" kgf/pelota por kgf/pelota de média)."))',
    '=IF(d_n<3,"","Concordância: 95% das janelas ficam entre "&FIXED(d_li,1)&" e "&FIXED(d_ls,1)&" kgf/pelota de diferença; "'
    '&FIXED(d_tol*100,0)&"% dentro da tolerância de "&cfgTolerancia&" kgf/pelota; CCC de Lin "&FIXED(d_ccc,3)&".")',
    '=IF(d_n<3,"",IF(AND(d_p>=0.05,d_ccc>=0.9),"Diagnóstico: a compressão online acompanha o laboratório no período.",'
    'IF(AND(d_p<0.05,d_bal<=0,d_bah>=0),"Diagnóstico: deslocamento constante entre os métodos; avaliar com o LCE um ajuste de "'
    '&FIXED(-d_vies,1)&" kgf/pelota na compressão online e acompanhar.",'
    '"Diagnóstico: a compressão online ainda não reproduz o laboratório; investigar a amostragem, a calibração da prensa e os pontos fora dos limites.")))',
]
for k, c in enumerate(conclusao):
    wdg.merge_range(38 + k, 1, 38 + k, 3, "", fmt(text_wrap=True, valign="top"))
    wdg.write_formula(38 + k, 1, c, fmt(text_wrap=True, valign="top"))
    wdg.set_row(38 + k, 30)

# Dados auxiliares dos gráficos (colunas F a J)
wdg.write("F4", "Dados dos gráficos (não editar)", f_rotulo)
wdg.write_row("F5", ["Reta de igualdade X", "Reta de igualdade Y"], f_cab)
wdg.write_row("F6", [150, 150])
wdg.write_row("F7", [450, 450])
wdg.write_row("F9", ["Bland-Altman X", "Viés", "Limite inferior", "Limite superior"], f_cab)
for k, ext in enumerate(("MIN", "MAX")):
    r = 10 + k
    wdg.write_formula(f"F{r}", f'=IF(d_n<2,NA(),{ext}(tMed))')
    wdg.write_formula(f"G{r}", '=IF(d_n<2,NA(),d_vies)')
    wdg.write_formula(f"H{r}", '=IF(d_n<2,NA(),d_li)')
    wdg.write_formula(f"I{r}", '=IF(d_n<2,NA(),d_ls)')
wdg.write_row("F13", ["Histograma: de", "até", "Faixa", "Janelas"], f_cab)
# Viés proporcional (Bland-Altman): regressão das diferenças sobre a média dos dois métodos
wdg.write_row("F30", ["Inclinação dif x média", "IC 95% inferior", "IC 95% superior"], f_cab)
wdg.write_formula("F31", '=IF(d_n<3,"",SLOPE(tDif,tMed))', fmt(num_format="0.000"))
wdg.write_formula("G31", '=IF(d_n<3,"",F31-TINV(0.05,d_n-2)*STEYX(tDif,tMed)/SQRT(DEVSQ(tMed)))', fmt(num_format="0.000"))
wdg.write_formula("H31", '=IF(d_n<3,"",F31+TINV(0.05,d_n-2)*STEYX(tDif,tMed)/SQRT(DEVSQ(tMed)))', fmt(num_format="0.000"))
for n_, cel in (("d_bas", "$F$31"), ("d_bal", "$G$31"), ("d_bah", "$H$31")):
    wb.define_name(n_, f"='Diagnóstico'!{cel}")
faixas = [(-1000, -60)] + [(x, x + 10) for x in range(-60, 60, 10)] + [(60, 1000)]
for k, (lo, hi) in enumerate(faixas):
    r = 14 + k
    rot = "abaixo de -60" if lo == -1000 else ("60 ou mais" if hi == 1000 else f"{lo} a {hi}")
    wdg.write_number(f"F{r}", lo)
    wdg.write_number(f"G{r}", hi)
    wdg.write_string(f"H{r}", rot)
    wdg.write_formula(f"I{r}", f'=COUNTIFS(tDif,">="&F{r},tDif,"<"&G{r})', f_dado_0)
H0, H1 = 14, 14 + len(faixas) - 1

# Resumo por dia
wdg.write("B46", "Resumo por dia", f_sec)
wdg.write_row("B47", ["Dia", "Janelas válidas", "Viés médio (kgf/pelota)"], f_cab)
wdg.write_row("F47", ["DP das diferenças", "Média laboratório", "Média online"], f_cab)
for k in range(62):
    r = 48 + k
    dia = f"B{r}"
    wdg.write_formula(dia, f'=IF(COUNT(tA)=0,"",IF(INT(MIN(tA))+{k}>MAX(tA),"",INT(MIN(tA))+{k}))', fmt(num_format="dd/mm/yyyy"))
    crit = f'tSit,"Válida",tA,">="&{dia},tA,"<"&{dia}+1'
    wdg.write_formula(f"C{r}", f'=IF({dia}="","",COUNTIFS({crit}))', fmt())
    wdg.write_formula(f"D{r}", f'=IF(OR({dia}="",C{r}=0),"",AVERAGEIFS(tDif,{crit}))', fmt(num_format="0.0"))
    wdg.write_formula(f"G{r}", f'=IF(OR({dia}="",C{r}=0),"",AVERAGEIFS(tLab,{crit}))', fmt(num_format="0.0"))
    wdg.write_formula(f"H{r}", f'=IF(OR({dia}="",C{r}=0),"",AVERAGEIFS(tOn,{crit}))', fmt(num_format="0.0"))
# STDEV condicional é fórmula de matriz
for k in range(62):
    r = 48 + k
    dia = f"B{r}"
    wdg.write_array_formula(f"F{r}:F{r}", f'{{=IF(OR({dia}="",C{r}<2),"",STDEV(IF(tSit="Válida",IF(tA>={dia},IF(tA<{dia}+1,tDif)))))}}',
                            fmt(num_format="0.0"))
wdg.conditional_format("B48:H109", {"type": "formula", "criteria": "=$B48<>\"\"", "format": wb.add_format({"bg_color": "#EEF3F7"})})

# ======================================================================= Painel
J = lambda col: f"Testes!${col}$2:${col}${NM}"  # noqa: E731
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
wp.write(0, 1, "Diagnóstico: compressão online x laboratório", f_banda_tit)
wp.set_row(1, 4)
if os.path.exists(LOGO):
    wp.insert_image("H1", LOGO, {"x_scale": 0.105, "y_scale": 0.105, "x_offset": 170, "y_offset": 14, "object_position": 3})
wp.write("B3", "Laboratório das 13:30 comparado com a média dos cortes online de 12:00, 12:30, 13:00 e 13:30; o das 15:30 com os cortes "
         "de 14:00 a 15:30, e assim por diante. Dados buscados no PIMS pelo suplemento Aspen, como no Relatório Físico.", f_status)
wp.write("B5", "Início do período", f_rotulo)
wp.write_datetime("C5", INICIO_PRENSA, f_ent_dh)
nome("pnlInicio", wp, "$C$5")
wp.write("B6", "Fim do período (em branco = agora)", f_rotulo)
wp.write_blank("C6", None, f_ent_dh)
nome("pnlFim", wp, "$C$6")
wp.set_row(4, 24)
wp.set_row(5, 24)
for cel, dx, macro, rotulo in [("E5", 0, "AtualizarPeriodo", "Atualizar período"),
                               ("E5", 160, "AtualizarDesdeInicio", "Desde o início da prensa"),
                               ("H5", 0, "TestarConexao", "Testar conexão"),
                               ("H5", 160, "LimparDados", "Limpar dados")]:
    wp.insert_button(cel, {"macro": macro, "caption": rotulo, "width": 150, "height": 44, "x_offset": dx, "y_offset": 2})
wp.merge_range("B8:I8", "Sem dados. Use um dos botões de atualização.", f_status)
nome("pnlStatus", wp, "$B$8")

f_ct = fmt(bold=True, font_size=12, font_color=AZUL, bottom=2, bottom_color=AMARELO)
cartoes = [
    ("B", "Amostra", [("Janelas válidas", "=d_n", "0"), ("Resultados do laboratório", "=d_tot", "0"),
                      ("Excluídas (cortador fora do automático)", "=d_exc", "0")]),
    ("E", "Viés e concordância", [("Viés médio, online menos lab (kgf/pelota)", "=d_vies", "0.0"),
                                  ("Valor-p do viés (teste t pareado)", "=d_p", "0.000"),
                                  ("Limites de concordância (kgf/pelota)", '=IF(d_li="","",FIXED(d_li,1)&"  a  "&FIXED(d_ls,1))', "@")]),
    ("H", "Correlação", [("Correlação r", "=d_r", "0.000"), ("Concordância de Lin (CCC)", "=d_ccc", "0.000"),
                         ("Dentro da tolerância", "=d_tol", "0%")]),
]
for col, titulo, linhas in cartoes:
    colv = chr(ord(col) + 1)
    wp.write(f"{col}10", titulo, f_ct)
    wp.write_blank(f"{colv}10", None, f_ct)
    for i, (rot, form, num) in enumerate(linhas):
        lin = 11 + i
        wp.write(f"{col}{lin}", rot, f_cartao)
        wp.write_formula(f"{colv}{lin}", form, fmt(bold=True, font_size=12, font_color=AZUL, bg_color="#EEF3F7", align="right", num_format=num))
        wp.set_row(lin - 1, 20)
wp.write("B15", "Diagnóstico", f_sec)
for k in range(len(conclusao)):
    r = 16 + k
    wp.merge_range(f"B{r}:I{r}", "", fmt(text_wrap=True, valign="top"))
    wp.write_formula(f"B{r}", f"='Diagnóstico'!B{39 + k}", fmt(text_wrap=True, valign="top"))
    wp.set_row(r - 1, 30)


def grafico(titulo, series, eixo_x, eixo_y, data=False, faixa_y=None, faixa_x=None, legenda=True):
    ch = wb.add_chart({"type": "scatter"})
    for s in series:
        ch.add_series(s)
    ch.set_title({"name": titulo, "name_font": {"name": F, "size": 12, "bold": True, "color": AZUL}})
    xa = {"name": eixo_x, "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 9},
          "major_gridlines": {"visible": False}, "label_position": "low"}
    ya = {"name": eixo_y, "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 9},
          "major_gridlines": {"visible": True, "line": {"color": "#E3E8EC"}}}
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
    ch.show_hidden_data()     # as colunas de apoio ficam ocultas na aba Testes
    ch.set_chartarea({"border": {"color": "#C9D3DB"}})
    return ch


def pontos(nome_, x, y, cor, linha=False, tamanho=5):
    s = {"name": nome_, "categories": x, "values": y,
         "marker": {"type": "circle", "size": tamanho, "fill": {"color": cor}, "border": {"color": cor}}}
    s["line"] = {"color": cor, "width": 1.75} if linha else {"none": True}
    return s


def reta(nome_, x, y, cor, tracejada=True):
    return {"name": nome_, "categories": x, "values": y, "marker": {"type": "none"},
            "line": {"color": cor, "width": 1.5, "dash_type": "dash" if tracejada else "solid"}}


G0 = 22
c1 = grafico("Composto online x laboratório ao longo do tempo",
             [pontos("Composto online", f"={J('S')}", f"={J('U')}", AZUL_MEDIO, linha=True),
              pontos("Laboratório", f"={J('S')}", f"={J('T')}", AZUL, linha=True)],
             "Horário do resultado do laboratório", "kgf/pelota", data=True, faixa_y=(150, 450))
wp.insert_chart(f"B{G0}", c1, {"x_scale": 2.4, "y_scale": 1.1})
DG = "'Diagnóstico'"
c2 = grafico("Laboratório x composto online",
             [pontos("Janelas válidas", f"={J('T')}", f"={J('U')}", AZUL_MEDIO),
              reta("Igualdade (online = laboratório)", f"={DG}!$F$6:$F$7", f"={DG}!$G$6:$G$7", "#7F8C99")],
             "Laboratório (kgf/pelota)", "Composto online (kgf/pelota)", faixa_y=(150, 450), faixa_x=(150, 450))
wp.insert_chart(f"B{G0 + 18}", c2, {"x_scale": 1.1, "y_scale": 1.3})
c3 = grafico("Bland-Altman: diferença x média dos métodos",
             [pontos("Janelas válidas", f"={J('W')}", f"={J('V')}", AZUL_MEDIO),
              reta("Viés", f"={DG}!$F$10:$F$11", f"={DG}!$G$10:$G$11", AZUL, tracejada=False),
              reta("Limite inferior", f"={DG}!$F$10:$F$11", f"={DG}!$H$10:$H$11", LARANJA),
              reta("Limite superior", f"={DG}!$F$10:$F$11", f"={DG}!$I$10:$I$11", LARANJA)],
             "Média de online e laboratório (kgf/pelota)", "Online menos laboratório (kgf/pelota)")
wp.insert_chart(f"B{G0 + 18}", c3, {"x_scale": 1.25, "y_scale": 1.3, "x_offset": 550})
c4 = wb.add_chart({"type": "column"})
c4.add_series({"name": "Janelas", "categories": f"={DG}!$H${H0}:$H${H1}", "values": f"={DG}!$I${H0}:$I${H1}",
               "fill": {"color": AZUL_MEDIO}, "gap": 40})
c4.set_title({"name": "Distribuição das diferenças (online menos laboratório)", "name_font": {"name": F, "size": 12, "bold": True, "color": AZUL}})
c4.set_x_axis({"name": "kgf/pelota", "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 8, "rotation": -45}})
c4.set_y_axis({"name": "Janelas", "name_font": {"name": F, "size": 9}, "num_font": {"name": F, "size": 9},
               "major_gridlines": {"visible": True, "line": {"color": "#E3E8EC"}}})
c4.set_legend({"none": True})
c4.set_chartarea({"border": {"color": "#C9D3DB"}})
wp.insert_chart(f"B{G0 + 39}", c4, {"x_scale": 2.4, "y_scale": 1.0})
wp.activate()

# ======================================================================= Leia-me
wl.hide_gridlines(2)
wl.set_column("A:A", 150)
textos = [
    ("Diagnóstico: compressão online x laboratório", f_tit),
    ("", None),
    ("A planilha sai sem dados. Na aba Painel, use o botão Testar conexão e depois Desde o início da prensa ou Atualizar período.", f_alerta),
    ("", None),
    ("Regra de comparação", f_sec),
    ("1. O laboratório analisa um composto de 2 h a partir de 4 cortes de 30 min. O resultado das 13:30 é comparado com a média dos "
     "cortes da compressão online de 12:00, 12:30, 13:00 e 13:30; o das 15:30 com os de 14:00 a 15:30, e assim por diante.", None),
    ("2. O resultado da prensa entra no PIMS cerca de 20 min depois do corte (13:20 vira corte das 13:00). A planilha desconta essa "
     "defasagem e arredonda para a meia hora mais próxima (aba Configurações).", None),
    ("3. O composto online é a média dos cortes que tiveram resultado. Valores zero não entram.", None),
    ("4. Se algum cortador da amostragem sair do automático durante a janela, a janela fica marcada e, por padrão, sai da estatística.", None),
    ("", None),
    ("Como funciona a busca", f_sec),
    ("1. Mesmo mecanismo do Relatório Físico: a macro escreve a fórmula GetCalculationValues do suplemento Aspen Process Data na aba "
     "oculta Dados_PIMS, espera o PIMS responder, lê os valores e grava nas abas Prensa, MES e Cortadores.", None),
    ("2. Pré-requisito: o suplemento Aspen Process Data ativo no Excel (o mesmo do Relatório Físico) e a planilha aberta com as macros habilitadas.", None),
    ("3. Tags: 307-GI-001 (compressão online), M4710050020-0031-HHLFU (laboratório), 307-GI-002 (desvio padrão) e U04-07CR001 e "
     "U04-07CR002 (cortadores), as mesmas da tela Compressão Online Usina 4.", None),
    ("4. Cada busca e cada teste registram as etapas na aba Log. Se algo não funcionar, envie a aba Log.", None),
    ("5. Alternativa sem macro: colar data/hora e valor nas colunas A e B das abas Prensa e MES (por exemplo, exportados do Process Explorer).", None),
    ("", None),
    ("Abas", f_sec),
    ("Painel: período, botões, indicadores, diagnóstico e gráficos.", None),
    ("Testes: uma linha por resultado do laboratório, com os 4 cortes online, o composto, a diferença, os cortadores e a situação.", None),
    ("Diagnóstico: estatística completa com a leitura de cada indicador, conclusão automática e resumo por dia.", None),
    ("Configurações: conexão com o PIMS, tags e regra de comparação (células amarelas).", None),
    ("Prensa, MES e Cortadores: dados buscados no PIMS e cálculo da janela de cada registro (cinza).", None),
    ("", None),
    ("Como ler o diagnóstico", f_sec),
    ("Viés: diferença média entre o composto online e o laboratório. O teste t pareado diz se ela é estatisticamente diferente de zero "
     "(valor-p menor que 0,05).", None),
    ("Limites de concordância (Bland-Altman): faixa em que se esperam 95% das diferenças de uma janela. Quanto mais estreita, mais o "
     "online substitui o laboratório janela a janela.", None),
    ("CCC de Lin: concordância entre os dois métodos (1 = idênticos). Diferente da correlação, cai quando há viés.", None),
    ("Correlação r: se os dois sobem e descem juntos. Pode ser alta mesmo com viés.", None),
    ("Viés proporcional: inclinação das diferenças sobre a média dos dois métodos (Bland-Altman). Se o intervalo de 95% não inclui "
     "zero, a diferença muda com o nível de compressão. A inclinação da regressão simples é só informativa: fica abaixo de 1 quando o "
     "laboratório também tem erro de medição; a de Deming considera o erro nos dois métodos.", None),
    ("Desvio entre os cortes: variação da compressão online dentro da mesma janela de 2 h, para comparar com a diferença online x laboratório.", None),
    ("", None),
    ("Capacidade: 5.000 resultados da prensa, 1.500 compostos do laboratório e 10.000 leituras de cada cortador (cerca de 2 meses com "
     "leitura a cada 10 min).", None),
]
for i, (t, fm) in enumerate(textos):
    wl.write(i, 0, t, fm or f_txt)

# ======================================================================= macros
if COM_VBA:
    modulo = open(os.path.join(AQUI, "vba", "ModPIMS.bas"), encoding="utf-8").read().replace("\r\n", "\n").replace("\n", "\r\n")
    if os.environ.get("VBA_TESTE"):    # rotina extra só para a conferência automática da macro
        modulo += open(os.environ["VBA_TESTE"], encoding="utf-8").read().replace("\n", "\r\n")
    docs = {"ThisWorkbook": "Option Explicit\r\n\r\nPrivate Sub Workbook_Open()\r\n    On Error Resume Next\r\n"
                            "    Me.Worksheets(\"Painel\").Activate\r\nEnd Sub\r\n"}
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
print(SAIDA, "com macros" if COM_VBA else "sem macros", f"| prensa {len(prensa)} | mes {len(mes)} | cortadores {len(cort[0])}+{len(cort[1])}")
