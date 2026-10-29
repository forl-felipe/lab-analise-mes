"""Gera a planilha de correlação entre a prensa automática (307-GI-001) e o laboratório (MES).

Uso: python gerar_planilha_correlacao.py [saida.xlsx]

Regra (definida pelo LCE):
  - O resultado da prensa entra no PIMS 20 min depois do corte: 13:20 corresponde ao corte das 13:00,
    13:50 ao corte das 13:30.
  - Os cortes de 12:00, 12:30, 13:00 e 13:30 formam um composto de 2 h, comparado com o resultado do MES
    das 13:30. Depois 14:00 a 15:30 com o MES das 15:30, e assim por diante.
  - O composto da prensa é a média dos cortes com resultado (1 a 4). Valores zero não entram.

Toda a lógica fica em fórmulas: basta trocar os dados das abas Prensa e MES. A planilha sai com DADOS
SIMULADOS para demonstração.
"""
import math
import random
import sys
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SAIDA = sys.argv[1] if len(sys.argv) > 1 else "correlacao_prensa_mes.xlsx"
NP = 20001      # última linha da aba Prensa
NM = 5001       # última linha das abas MES e Correlacao
F = "Arial"
AZUL, AMARELO, AZUL_CLARO = "00335A", "FFC000", "59C6F2"
CAB = PatternFill("solid", fgColor=AZUL)
ENTRADA = PatternFill("solid", fgColor="FFF2CC")
CALC = PatternFill("solid", fgColor="F2F2F2")
CARTAO = PatternFill("solid", fgColor="EEF3F7")
FINO = Side(style="thin", color="C9D3DB")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
DH = "dd/mm/yyyy hh:mm"

wb = Workbook()


def cabecalho(ws, titulos, larguras, linha=1):
    for c, (t, w) in enumerate(zip(titulos, larguras), 1):
        cel = ws.cell(row=linha, column=c, value=t)
        cel.font = Font(name=F, bold=True, color="FFFFFF")
        cel.fill = CAB
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cel.column_letter].width = w
    ws.row_dimensions[linha].height = 42
    ws.freeze_panes = ws.cell(row=linha + 1, column=1)


def fonte_arial(ws):
    for linha in ws.iter_rows():
        for c in linha:
            f = c.font
            if f is None or f.name != F:
                c.font = Font(name=F, bold=f.bold if f else False, color=f.color if f else None,
                              size=f.size if f else 11, italic=f.italic if f else False)


# ======================================================================= Painel (primeira aba)
wp = wb.active
wp.title = "Painel"

# ======================================================================= Leia-me
wl = wb.create_sheet("Leia-me")
textos = [
    ("Correlação prensa automática x laboratório (MES)", "titulo"),
    ("", None),
    ("ATENÇÃO: a planilha sai com DADOS SIMULADOS. Use o botão de busca (macro) ou cole os dados reais nas abas Prensa e MES.", "alerta"),
    ("", None),
    ("Regra de agrupamento (aba Parametros)", "secao"),
    ("1. O resultado da prensa entra no PIMS 20 min depois do corte. A planilha desconta essa defasagem: 13:20 vira corte das 13:00, 13:50 vira corte das 13:30.", None),
    ("2. Depois de descontar, o horário é arredondado para a meia hora mais próxima, para absorver segundos e pequenos atrasos.", None),
    ("3. Os cortes de 12:00, 12:30, 13:00 e 13:30 formam um composto de 2 h, comparado com o resultado do MES das 13:30. Depois 14:00 a 15:30 com o MES das 15:30, e assim por diante.", None),
    ("4. O composto da prensa é a média dos cortes que tiveram resultado (1 a 4). Valores zero, que a tag registra entre os ensaios, não entram.", None),
    ("5. Se um corte tiver mais de um registro, a planilha usa a média dele, para que cada corte conte uma vez só.", None),
    ("", None),
    ("Como atualizar os dados", "secao"),
    ("Opção A (busca automática): importe o módulo BuscarPIMS.bas (Alt+F11 > Arquivo > Importar arquivo), salve como .xlsm, preencha período e conexão na aba Parametros e execute a macro BuscarDadosPIMS.", None),
    ("Opção B (Aspen Process Data no Excel): busque os dados reais (actual) das tags no período e cole data/hora e valor nas colunas amarelas das abas Prensa e MES, a partir da linha 2.", None),
    ("Opção C (exportação do Process Explorer): exporte a tendência das duas tags para Excel ou CSV e cole da mesma forma.", None),
    ("", None),
    ("Abas", "secao"),
    ("Painel: indicadores, gráficos e resumo mensal.", None),
    ("Parametros: defasagem, deslocamento das janelas, mínimo de cortes, tolerância, período e conexão (células amarelas).", None),
    ("Prensa: dados da tag 307-GI-001 (amarelo) e o cálculo do corte e da janela de cada registro (cinza).", None),
    ("MES: dados da tag M4710050020-0031-HHLFU (amarelo).", None),
    ("Correlacao: uma linha por resultado do MES, com os 4 cortes da prensa, o composto, a diferença e a situação.", None),
    ("", None),
    ("Legenda", "secao"),
    ("Amarelo: entrada de dados ou parâmetros. Cinza: fórmulas, não editar. As colunas ocultas da aba Correlacao alimentam os gráficos; o #N/D nelas é proposital, para o Excel não desenhar pontos vazios.", None),
    ("Capacidade: 20.000 registros da prensa e 5.000 resultados do MES (cerca de 1 ano). Para mais, copie as fórmulas das últimas linhas para baixo.", None),
]
for i, (t, tipo) in enumerate(textos, 1):
    c = wl.cell(row=i, column=1, value=t)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if tipo == "titulo":
        c.font = Font(name=F, bold=True, size=15, color=AZUL)
    elif tipo == "secao":
        c.font = Font(name=F, bold=True, size=12, color=AZUL)
    elif tipo == "alerta":
        c.font = Font(name=F, bold=True, color="C00000")
wl.column_dimensions["A"].width = 150

# ======================================================================= Parametros
wpar = wb.create_sheet("Parametros")
cabecalho(wpar, ["Parâmetro", "Valor", "Observação"], [56, 26, 90])
params = [
    ("Defasagem entre o corte e a entrada no PIMS (min)", 20, "Informado pelo LCE: corte das 13:00 entra no PIMS às 13:20.", "0"),
    ("Deslocamento das janelas (meias horas)", 0, "0 = janelas de 12:00 a 13:30, 14:00 a 15:30... Use 1 para começar nas meias horas ímpares, e assim por diante.", "0"),
    ("Mínimo de cortes para a janela ser válida", 1, "1 = usa a média mesmo com um só corte. Use 4 para considerar só janelas completas.", "0"),
    ("Ignorar valores zero (1 = sim, 0 = não)", 1, "A tag da prensa volta a zero entre os ensaios.", "0"),
    ("Ajuste do horário do MES (min)", 0, "Use se o horário gravado no MES não for o do último corte (13:30). Exemplo: -30 se o MES gravar 14:00.", "0"),
    ("Tolerância da diferença prensa x laboratório (kgf/pelota)", 20, "A confirmar pelo LCE. Usada no indicador de janelas dentro da tolerância.", "0"),
    ("", None, "", None),
    ("Busca automática (macro BuscarDadosPIMS)", None, "", None),
    ("Início do período", datetime(2026, 9, 1), "Data e hora inicial da busca.", DH),
    ("Fim do período", datetime(2026, 10, 15, 23, 59), "Data e hora final da busca.", DH),
    ("Fonte ODBC (DSN) do Aspen SQLplus", "IP21", "Nome da conexão ODBC do InfoPlus.21 na sua máquina (confirmar com a TI).", None),
    ("Tag da prensa automática", "307-GI-001", "", None),
    ("Tag do laboratório (MES)", "M4710050020-0031-HHLFU", "", None),
]
for r, (n, v, o, fmt) in enumerate(params, 2):
    wpar.cell(row=r, column=1, value=n).font = Font(name=F, bold=(v is None and n != ""))
    if v is not None:
        c = wpar.cell(row=r, column=2, value=v)
        c.fill = ENTRADA
        c.border = BORDA
        if fmt:
            c.number_format = fmt
    wpar.cell(row=r, column=3, value=o).alignment = Alignment(wrap_text=True)
P_DEF, P_DESL, P_MIN, P_ZERO, P_AJ, P_TOL = (f"Parametros!$B${i}" for i in range(2, 8))

# ======================================================================= dados simulados
random.seed(7)
inicio = datetime(2026, 9, 1)
dias = 45


def qualidade(t):
    h = (t - inicio).total_seconds() / 3600
    return 300 + 22 * math.sin(h / 40) + 9 * math.sin(h / 5.3)


cortes = []
t = inicio
while t < inicio + timedelta(days=dias):
    cortes.append(t)
    t += timedelta(minutes=30)

prensa, mes = [], []
parada = (inicio + timedelta(days=17, hours=6), inicio + timedelta(days=18, hours=2))
for c in cortes:
    if parada[0] <= c < parada[1] or random.random() < 0.05:
        continue
    entrada = c + timedelta(minutes=20, seconds=random.randint(0, 90))
    valor = round(qualidade(c) + 8 + random.gauss(0, 12), 1)
    prensa.append((entrada, valor))
    prensa.append((entrada + timedelta(minutes=random.randint(4, 8)), 0))   # a tag volta a zero
for k in range(0, len(cortes), 4):
    jan = cortes[k:k + 4]
    if len(jan) < 4 or (parada[0] <= jan[0] < parada[1]):
        continue
    lab = round(sum(qualidade(c) for c in jan) / 4 + random.gauss(0, 8))
    mes.append((jan[-1], lab))

# ======================================================================= Prensa
wpr = wb.create_sheet("Prensa")
cabecalho(wpr, ["Data e hora no PIMS", "Média prensa 307-GI-001 (kgf/pelota)", "Hora do corte", "Índice de meia hora",
                "Janela", "Posição do corte na janela (1 a 4)", "Válido (1 = entra na média)", "Horário do MES correspondente"],
          [19, 16, 17, 11, 10, 12, 12, 17])
wpr["A1"].comment = Comment("Dados simulados. Substituir pelos dados reais (a macro faz isso automaticamente).", "LCE")
for r in range(2, NP + 1):
    if r - 2 < len(prensa):
        a, b = prensa[r - 2]
        wpr.cell(row=r, column=1, value=a)
        wpr.cell(row=r, column=2, value=b)
    wpr.cell(row=r, column=1).fill = ENTRADA
    wpr.cell(row=r, column=2).fill = ENTRADA
    wpr.cell(row=r, column=1).number_format = DH
    wpr.cell(row=r, column=2).number_format = "0.0"
    f = {
        "D": f'=IF(A{r}="","",ROUND((A{r}-{P_DEF}/1440)*48,0))',
        "C": f'=IF(D{r}="","",D{r}/48)',
        "E": f'=IF(D{r}="","",INT((D{r}-{P_DESL})/4))',
        "F": f'=IF(D{r}="","",MOD(D{r}-{P_DESL},4)+1)',
        "G": f'=IF(A{r}="","",IF(AND(ISNUMBER(B{r}),OR({P_ZERO}=0,B{r}<>0)),1,0))',
        "H": f'=IF(E{r}="","",(E{r}*4+3+{P_DESL})/48)',
    }
    for col, formula in f.items():
        c = wpr[f"{col}{r}"]
        c.value = formula
        c.fill = CALC
    wpr[f"C{r}"].number_format = DH
    wpr[f"H{r}"].number_format = DH

# ======================================================================= MES
wm = wb.create_sheet("MES")
cabecalho(wm, ["Data e hora do resultado no MES", "Compressão laboratório (kgf/pelota)", "Índice de meia hora", "Janela",
               "Posição (esperado 4)"], [21, 16, 11, 10, 12])
wm["A1"].comment = Comment("Dados simulados. Substituir pelos dados reais.", "LCE")
for r in range(2, NM + 1):
    if r - 2 < len(mes):
        a, b = mes[r - 2]
        wm.cell(row=r, column=1, value=a)
        wm.cell(row=r, column=2, value=b)
    wm.cell(row=r, column=1).fill = ENTRADA
    wm.cell(row=r, column=2).fill = ENTRADA
    wm.cell(row=r, column=1).number_format = DH
    f = {
        "C": f'=IF(A{r}="","",ROUND((A{r}+{P_AJ}/1440)*48,0))',
        "D": f'=IF(C{r}="","",INT((C{r}-{P_DESL})/4))',
        "E": f'=IF(C{r}="","",MOD(C{r}-{P_DESL},4)+1)',
    }
    for col, formula in f.items():
        c = wm[f"{col}{r}"]
        c.value = formula
        c.fill = CALC

# ======================================================================= Correlacao
wc = wb.create_sheet("Correlacao")
tit = ["Horário do MES", "Janela de cortes", "Laboratório (kgf/pelota)", "Corte 1", "Corte 2", "Corte 3", "Corte 4",
       "Cortes com resultado", "Composto prensa (kgf/pelota)", "Diferença prensa menos laboratório", "Situação",
       "Diferença absoluta", "Dentro da tolerância",
       "id janela", "X gráfico", "Lab gráfico", "Prensa gráfico", "Dif gráfico", "Lab válido", "Prensa válido", "Dif válida"]
cabecalho(wc, tit, [17, 15, 13, 10, 10, 10, 10, 11, 14, 14, 26, 11, 12, 8, 8, 8, 8, 8, 8, 8, 8])
PE, PF, PG, PB = (f"Prensa!${c}$2:${c}${NP}" for c in "EFGB")
for r in range(2, NM + 1):
    f = {
        "A": f'=IF(MES!A{r}="","",MES!C{r}/48)',
        "N": f'=IF(A{r}="","",MES!D{r})',
        "B": f'=IF(A{r}="","",TEXT((N{r}*4+{P_DESL})/48,"hh:mm")&" a "&TEXT(A{r},"hh:mm"))',
        "C": f'=IF(A{r}="","",MES!B{r})',
        "H": f'=IF(A{r}="","",COUNT(D{r}:G{r}))',
        "I": f'=IF(A{r}="","",IF(H{r}=0,"",AVERAGE(D{r}:G{r})))',
        "J": f'=IF(OR(I{r}="",C{r}=""),"",I{r}-C{r})',
        "K": (f'=IF(A{r}="","",IF(MES!E{r}<>4,"Horário do MES fora do padrão",'
              f'IF(H{r}<{P_MIN},"Cortes insuficientes","Válida")))'),
        "L": f'=IF(K{r}="Válida",ABS(J{r}),"")',
        "M": f'=IF(K{r}="Válida",IF(L{r}<={P_TOL},"Sim","Não"),"")',
        "O": f'=IF(K{r}="Válida",A{r},NA())',
        "P": f'=IF(K{r}="Válida",C{r},NA())',
        "Q": f'=IF(K{r}="Válida",I{r},NA())',
        "R": f'=IF(K{r}="Válida",J{r},NA())',
        "S": f'=IF(K{r}="Válida",C{r},"")',
        "T": f'=IF(K{r}="Válida",I{r},"")',
        "U": f'=IF(K{r}="Válida",J{r},"")',
    }
    for k, col in enumerate("DEFG", 1):
        f[col] = f'=IF(A{r}="","",IFERROR(AVERAGEIFS({PB},{PE},N{r},{PF},{k},{PG},1),""))'
    for col, formula in f.items():
        c = wc[f"{col}{r}"]
        c.value = formula
        c.fill = CALC
    wc[f"A{r}"].number_format = DH
    wc[f"O{r}"].number_format = DH
    for col in "CDEFGIJLPQRSTU":
        wc[f"{col}{r}"].number_format = "0.0"
for col in "NOPQRSTU":
    wc.column_dimensions[col].hidden = True

# ======================================================================= Painel
J = lambda col: f"Correlacao!${col}$2:${col}${NM}"
wp.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJ", [3, 30, 21, 3, 42, 12, 3, 36, 14, 14]):
    wp.column_dimensions[col].width = w
wp.merge_cells("B1:J1")
wp["B1"] = "Correlação prensa automática x laboratório (MES)"
wp["B1"].font = Font(name=F, bold=True, size=18, color="FFFFFF")
wp["B1"].fill = CAB
wp["B1"].alignment = Alignment(vertical="center", indent=1)
wp.row_dimensions[1].height = 40
for col in "ABCDEFGHIJ":
    wp[f"{col}1"].fill = CAB
    wp[f"{col}2"].fill = PatternFill("solid", fgColor=AMARELO)
wp.row_dimensions[2].height = 5
wp["B3"] = "Composto de 2 h da prensa (média dos cortes 12:00 a 13:30) comparado com o resultado do MES das 13:30, e assim por diante."
wp["B3"].font = Font(name=F, italic=True, color="5B6B77")

indicadores = [
    ("B5", "Período analisado", [
        ("Início", f"=IFERROR(MIN({J('A')}),\"\")", DH),
        ("Fim", f"=IFERROR(MAX({J('A')}),\"\")", DH),
        ("Resultados do MES", f"=COUNT({J('A')})", "0"),
        ("Janelas válidas", f'=COUNTIF({J("K")},"Válida")', "0"),
        ("Janelas com os 4 cortes", f'=COUNTIFS({J("K")},"Válida",{J("H")},4)', "0"),
        ("Janelas com 1 a 3 cortes", f'=COUNTIFS({J("K")},"Válida",{J("H")},"<4")', "0"),
    ]),
    ("E5", "Concordância", [
        ("Média do laboratório (kgf/pelota)", f'=IFERROR(AVERAGEIFS({J("C")},{J("K")},"Válida"),"")', "0.0"),
        ("Média do composto da prensa (kgf/pelota)", f'=IFERROR(AVERAGEIFS({J("I")},{J("K")},"Válida"),"")', "0.0"),
        ("Viés médio, prensa menos laboratório", f'=IFERROR(AVERAGE({J("U")}),"")', "0.0"),
        ("Desvio padrão das diferenças", f'=IFERROR(STDEV({J("U")}),"")', "0.0"),
        ("Erro médio absoluto", f'=IFERROR(AVERAGE({J("L")}),"")', "0.0"),
        ("Janelas dentro da tolerância", f'=IFERROR(COUNTIF({J("M")},"Sim")/COUNTIF({J("K")},"Válida"),"")', "0%"),
    ]),
    ("H5", "Correlação", [
        ("Coeficiente de correlação (r)", f'=IFERROR(CORREL({J("S")},{J("T")}),"")', "0.00"),
        ("Coeficiente de determinação (R²)", '=IFERROR(I6^2,"")', "0.00"),
        ("Inclinação (prensa = a + b x lab)", f'=IFERROR(SLOPE({J("T")},{J("S")}),"")', "0.00"),
        ("Intercepto a", f'=IFERROR(INTERCEPT({J("T")},{J("S")}),"")', "0.0"),
        ("Limite inferior de concordância", '=IFERROR(F8-1.96*F9,"")', "0.0"),
        ("Limite superior de concordância", '=IFERROR(F8+1.96*F9,"")', "0.0"),
    ]),
]
for ancora, titulo, linhas in indicadores:
    col = ancora[0]
    colv = chr(ord(col) + 1)
    wp[f"{col}5"] = titulo
    wp[f"{col}5"].font = Font(name=F, bold=True, size=12, color=AZUL)
    wp[f"{col}5"].border = Border(bottom=Side(style="medium", color=AMARELO))
    wp[f"{colv}5"].border = Border(bottom=Side(style="medium", color=AMARELO))
    for i, (rot, form, fmt) in enumerate(linhas, 6):
        a = wp[f"{col}{i}"]
        a.value = rot
        a.fill = CARTAO
        v = wp[f"{colv}{i}"]
        v.value = form
        v.number_format = fmt
        v.font = Font(name=F, bold=True, size=12, color=AZUL)
        v.fill = CARTAO
        v.alignment = Alignment(horizontal="right")
wp["B13"] = "Leitura"
wp["B13"].font = Font(name=F, bold=True, size=12, color=AZUL)
wp["B14"] = ('=IF(C9=0,"",IF(C9<10,"Poucas janelas válidas para concluir (mínimo recomendado: 30).",'
             'IF(I7>=0.7,"Correlação forte",IF(I7>=0.4,"Correlação moderada","Correlação fraca"))'
             '&" (R² = "&FIXED(I7,2)&"). Em média a prensa mede "&FIXED(ABS(F8),1)&" kgf/pelota "'
             '&IF(F8>=0,"acima","abaixo")&" do laboratório; 95% das janelas ficam entre "&FIXED(I10,1)&" e "&FIXED(I11,1)&"."))')
wp.merge_cells("B14:J14")
wp["B14"].alignment = Alignment(wrap_text=True, vertical="top")
wp.row_dimensions[14].height = 34

# Gráficos
def disperso(titulo, x, ys, eixo_x, eixo_y, ancora, larg=24, alt=9.5, data=False, tendencia=False, so_marcador=False, faixa_y=None, faixa_x=None):
    ch = ScatterChart()
    ch.title = titulo
    ch.style = 13
    ch.height, ch.width = alt, larg
    ch.x_axis.title = eixo_x
    ch.y_axis.title = eixo_y
    xref = Reference(wc, min_col=x, min_row=2, max_row=NM)
    for col, nome, cor in ys:
        s = Series(Reference(wc, min_col=col, min_row=2, max_row=NM), xref, title=nome)
        s.marker.symbol = "circle"
        s.marker.size = 5
        s.marker.graphicalProperties.solidFill = cor
        s.marker.graphicalProperties.line.solidFill = cor
        if so_marcador:
            s.graphicalProperties.line.noFill = True
        else:
            s.graphicalProperties.line.solidFill = cor
            s.graphicalProperties.line.width = 22000
        if tendencia:
            s.trendline = Trendline(trendlineType="linear", dispRSqr=True, dispEq=True)
        ch.series.append(s)
    if data:
        ch.x_axis.number_format = "dd/mm"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.visible_cells_only = False   # as colunas de apoio ficam ocultas na aba Correlacao
    if len(ys) == 1:
        ch.legend = None
    if faixa_y:
        ch.y_axis.scaling.min, ch.y_axis.scaling.max = faixa_y
    if faixa_x:
        ch.x_axis.scaling.min, ch.x_axis.scaling.max = faixa_x
    wp.add_chart(ch, ancora)


disperso("Composto de 2 h: prensa automática x laboratório", 15, [(17, "Prensa (composto)", "0091D5"), (16, "Laboratório (MES)", AZUL)],
         "Data", "kgf/pelota", "B16", larg=29, alt=9, data=True, faixa_y=(150, 450))
disperso("Correlação: laboratório x prensa", 19, [(20, "Janelas válidas", "0091D5")], "Laboratório (kgf/pelota)",
         "Prensa, composto (kgf/pelota)", "B36", larg=11, alt=10, tendencia=True, so_marcador=True, faixa_y=(150, 450), faixa_x=(150, 450))
disperso("Diferença prensa menos laboratório por janela", 15, [(18, "Diferença", "F37021")], "Data", "kgf/pelota", "E36",
         larg=18, alt=10, data=True, so_marcador=True)

# Resumo mensal
wp["B57"] = "Resumo mensal"
wp["B57"].font = Font(name=F, bold=True, size=12, color=AZUL)
for i, t in zip((0, 1, 3), ["Mês", "Janelas válidas", "Viés médio, prensa menos laboratório (kgf/pelota)"]):
    c = wp.cell(row=58, column=2 + i, value=t)
    c.font = Font(name=F, bold=True, color="FFFFFF")
    c.fill = CAB
for k in range(12):
    r = 59 + k
    wp[f"B{r}"] = f'=IF(OR($C$6="",DATE(YEAR($C$6),MONTH($C$6)+{k},1)>$C$7),"",DATE(YEAR($C$6),MONTH($C$6)+{k},1))'
    wp[f"B{r}"].number_format = "mmm/yyyy"
    wp[f"C{r}"] = f'=IF(B{r}="","",COUNTIFS({J("K")},"Válida",{J("A")},">="&B{r},{J("A")},"<"&EDATE(B{r},1)))'
    wp[f"D{r}"] = ""
    wp[f"E{r}"] = f'=IF(OR(B{r}="",C{r}=0),"",AVERAGEIFS({J("U")},{J("A")},">="&B{r},{J("A")},"<"&EDATE(B{r},1)))'
    wp[f"E{r}"].number_format = "0.0"
    for col in "BCE":
        wp[f"{col}{r}"].fill = CARTAO

for w in wb.worksheets:
    fonte_arial(w)
wp["B1"].font = Font(name=F, bold=True, size=18, color="FFFFFF")
wb.save(SAIDA)
print(f"{SAIDA}: {len(prensa)} registros da prensa, {len(mes)} resultados do MES (simulados)")
