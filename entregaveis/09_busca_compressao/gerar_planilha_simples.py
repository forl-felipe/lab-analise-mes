"""Gera a planilha simples compressão online x laboratório (sem macro dentro do arquivo).

Uso: python gerar_planilha_simples.py [saida.xlsx] [--teste dados.json]

A macro de busca fica em ModBusca.bas (Windows-1252 com CRLF, como o editor do VBA espera) e é importada pelo próprio Excel do usuário
(Alt+F11 > Arquivo > Importar arquivo), que a compila como no Relatório Físico.
Com saída .xlsm, o gerador embute a macro (usado só na conferência automática no LibreOffice).

Regra: o laboratório das 13:30 corresponde aos cortes da compressão online de 12:00, 12:30, 13:00 e
13:30; o das 15:30 aos de 14:00 a 15:30, e assim por diante. O composto online é a média dos cortes
com resultado. O resultado entra no PIMS cerca de 20 min depois do corte (aba Configuracoes).
"""
import json
import os
import sys
from datetime import datetime

import xlsxwriter

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "ferramentas"))

NP = 5001       # última linha da aba Online (5.000 resultados)
NM = 1501       # última linha das abas Laboratorio e Comparacao (1.500 resultados)
INICIO = datetime(2026, 11, 16, 0, 0)
AZUL, AMARELO = "#00335A", "#FFC000"
F = "Arial"

args = list(sys.argv[1:])
teste = None
if "--teste" in args:
    i = args.index("--teste")
    teste = json.load(open(args[i + 1]))
    del args[i:i + 2]
SAIDA = args[0] if args else os.path.join(AQUI, "compressao_online_x_laboratorio.xlsx")
COM_VBA = SAIDA.lower().endswith(".xlsm")

wb = xlsxwriter.Workbook(SAIDA)


def fmt(**k):
    base = {"font_name": F, "font_size": 10, "valign": "vcenter"}
    base.update(k)
    return wb.add_format(base)


f_cab = fmt(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, border=1, border_color="#C9D3DB")
f_ent = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB")
f_ent_dh = fmt(bg_color="#FFF2CC", border=1, border_color="#C9D3DB", num_format="dd/mm/yyyy hh:mm")
f_dh = fmt(num_format="dd/mm/yyyy hh:mm")
f_1 = fmt(num_format="0.0")
f_calc = fmt(bg_color="#F2F2F2")
f_calc_dh = fmt(bg_color="#F2F2F2", num_format="dd/mm/yyyy hh:mm")
f_calc_1 = fmt(bg_color="#F2F2F2", num_format="0.0")
f_rot = fmt(bold=True, font_color=AZUL)
f_sec = fmt(bold=True, font_size=12, font_color=AZUL)
f_txt = fmt(text_wrap=True, valign="top")
f_status = fmt(italic=True, font_color="#5B6B77")

wp = wb.add_worksheet("Painel")
wc = wb.add_worksheet("Comparacao")
wo = wb.add_worksheet("Online")
wl = wb.add_worksheet("Laboratorio")
wcfg = wb.add_worksheet("Configuracoes")
wlog = wb.add_worksheet("Log")
wa = wb.add_worksheet("Consulta_Aspen")


def nome(n, ws, cel):
    wb.define_name(n, f"='{ws.get_name()}'!{cel}")


# ======================================================================= Configuracoes
wcfg.set_column("A:A", 48)
wcfg.set_column("B:B", 28)
wcfg.set_column("C:C", 90)
wcfg.write_row(0, 0, ["Parâmetro", "Valor", "Observação"], f_cab)
cfg = [
    ("Conexão com o PIMS (igual ao Relatório Físico)", None, None, None),
    ("Servidor do PIMS", "UBU", "cfgServidor", ""),
    ("Formato da data para o Aspen", "Texto dd/mm/aaaa", "cfgFormatoData", "Opções: Texto dd/mm/aaaa, Texto mm/dd/aaaa ou Data do Excel."),
    ("Espera máxima pela resposta do PIMS (s)", 300, "cfgTimeout", ""),
    ("Início da fórmula do suplemento", "=GetCalculationValues(", "cfgPrefixo", "Como o Excel mostra a função do Aspen."),
    ("", None, None, None),
    ("Compressão online", None, None, None),
    ("Tag", "307-GI-001", "cfgTagPrensa", "Média de compressão de cada ensaio (kgf/pelota)."),
    ("Tipo (mapa)", "IP_ANALOGMAP", "cfgMapaPrensa", ""),
    ("Intervalo da consulta", "1m", "cfgPeriodoPrensa", "Leitura a cada 1 min; cada resultado novo vira uma linha da aba Online."),
    ("Tipo de cálculo do suplemento", "0", "cfgCalculoPrensa", "0 = valor no instante."),
    ("", None, None, None),
    ("Laboratório", None, None, None),
    ("Tag", "M4710050020-0031-HHLFU", "cfgTagLab", "Resistência à Compressão US4 (kgf/pelota)."),
    ("Tipo (mapa)", "IP_MESVALOR", "cfgMapaLab", ""),
    ("Intervalo da consulta", "30m", "cfgPeriodoLab", ""),
    ("Tipo de cálculo do suplemento", "1", "cfgCalculoLab", "1 = mesmo código da consulta de qualidade do Relatório Físico."),
    ("", None, None, None),
    ("Regra de comparação", None, None, None),
    ("Defasagem entre o corte e a entrada no PIMS (min)", 20, "cfgDefasagem", "O corte das 13:00 entra no PIMS às 13:20. Use 0 se a hora da tag já for a do corte."),
    ("Deslocamento das janelas (meias horas)", 0, "cfgDeslocamento", "0 = cortes de 12:00 a 13:30 com o laboratório das 13:30."),
    ("Ajuste do horário do laboratório (min)", 0, "cfgAjusteMES", "Use se o resultado do laboratório não for gravado no horário do último corte."),
]
for i, (rot, v, n, obs) in enumerate(cfg, 1):
    if v is None:
        wcfg.write(i, 0, rot, f_rot if rot else None)
        continue
    wcfg.write(i, 0, rot)
    if isinstance(v, str):
        wcfg.write_string(i, 1, v, f_ent)
    else:
        wcfg.write_number(i, 1, v, f_ent)
    wcfg.write(i, 2, obs)
    nome(n, wcfg, f"$B${i + 1}")

# ======================================================================= Consulta_Aspen e Log
wa.write("A1", "Área de trabalho da macro (consultas ao Aspen). Não editar.", f_rot)
wa.write("A3", "Início")
wa.write_string("B3", "")
wa.write("A4", "Fim")
wa.write_string("B4", "")
for c, titulo in ((0, "compressão online"), (4, "laboratório")):
    wa.set_column(c, c, 19)
    wa.set_column(c + 1, c + 1, 14)
    wa.write(5, c, "Consulta: " + titulo, f_rot)
    wa.write(6, c, "Sem consulta")
    wa.write_row(7, c, ["Data/hora", "Valor"], f_cab)
wlog.set_column("A:A", 19)
wlog.set_column("B:B", 22)
wlog.set_column("C:C", 160)
wlog.write("A1", "Registro da última busca ou teste (preenchido pela macro). Se algo não funcionar, envie esta aba.", f_rot)
wlog.write_row("A2", ["Data e hora", "Etapa", "Detalhe"], f_cab)
for r in range(2, 1000):
    wlog.set_row(r, None, fmt(num_format="dd/mm/yyyy hh:mm:ss"))

# ======================================================================= Online e Laboratorio
D, DESL, AJ = "cfgDefasagem", "cfgDeslocamento", "cfgAjusteMES"
wo.write_row(0, 0, ["Data e hora no PIMS", "Compressão online 307-GI-001 (kgf/pelota)", "Hora do corte", "Índice de meia hora",
                    "Janela", "Posição na janela (1 a 4)", "Válido (1 = entra na média)"], f_cab)
wo.set_row(0, 42)
for k, w in enumerate([19, 16, 17, 11, 10, 11, 11]):
    wo.set_column(k, k, w)
wo.freeze_panes(1, 0)
online = teste["online"] if teste else []
for r in range(2, NP + 1):
    i = r - 2
    if i < len(online):
        wo.write_datetime(r - 1, 0, datetime.fromisoformat(online[i][0]), f_dh)
        wo.write_number(r - 1, 1, online[i][1], f_1)
    else:
        wo.write_blank(r - 1, 0, None, f_dh)
    wo.write_formula(r - 1, 3, f'=IF(A{r}="","",ROUND((A{r}-{D}/1440)*48,0))', f_calc)
    wo.write_formula(r - 1, 2, f'=IF(D{r}="","",D{r}/48)', f_calc_dh)
    wo.write_formula(r - 1, 4, f'=IF(D{r}="","",INT((D{r}-{DESL})/4))', f_calc)
    wo.write_formula(r - 1, 5, f'=IF(D{r}="","",MOD(D{r}-{DESL},4)+1)', f_calc)
    wo.write_formula(r - 1, 6, f'=IF(A{r}="","",IF(AND(ISNUMBER(B{r}),B{r}<>0),1,0))', f_calc)

wl.write_row(0, 0, ["Data e hora do resultado", "Laboratório M4710050020-0031-HHLFU (kgf/pelota)", "Índice de meia hora",
                    "Janela", "Posição (esperado 4)"], f_cab)
wl.set_row(0, 42)
for k, w in enumerate([19, 18, 11, 10, 11]):
    wl.set_column(k, k, w)
wl.freeze_panes(1, 0)
lab = teste["lab"] if teste else []
for r in range(2, NM + 1):
    i = r - 2
    if i < len(lab):
        wl.write_datetime(r - 1, 0, datetime.fromisoformat(lab[i][0]), f_dh)
        wl.write_number(r - 1, 1, lab[i][1], f_1)
    else:
        wl.write_blank(r - 1, 0, None, f_dh)
    wl.write_formula(r - 1, 2, f'=IF(A{r}="","",ROUND((A{r}+{AJ}/1440)*48,0))', f_calc)
    wl.write_formula(r - 1, 3, f'=IF(C{r}="","",INT((C{r}-{DESL})/4))', f_calc)
    wl.write_formula(r - 1, 4, f'=IF(C{r}="","",MOD(C{r}-{DESL},4)+1)', f_calc)

# ======================================================================= Comparacao
wc.write_row(0, 0, ["Horário do laboratório", "Cortes da compressão online", "Laboratório (kgf/pelota)", "Corte 1", "Corte 2",
                    "Corte 3", "Corte 4", "Cortes com resultado", "Composto online (kgf/pelota)",
                    "Diferença online menos laboratório (kgf/pelota)", "Observação", "id"], f_cab)
wc.set_row(0, 42)
for k, w in enumerate([17, 15, 13, 9, 9, 9, 9, 10, 14, 16, 30, 6]):
    wc.set_column(k, k, w, None, {"hidden": k == 11})
wc.freeze_panes(1, 0)
OB, OE, OF, OG = (f"Online!${c}$2:${c}${NP}" for c in "BEFG")
for r in range(2, NM + 1):
    f = {
        "A": (f'=IF(Laboratorio!A{r}="","",Laboratorio!C{r}/48)', f_calc_dh),
        "L": (f'=IF(A{r}="","",Laboratorio!D{r})', f_calc),
        "B": (f'=IF(A{r}="","",TEXT((L{r}*4+{DESL})/48,"hh:mm")&" a "&TEXT(A{r},"hh:mm"))', f_calc),
        "C": (f'=IF(A{r}="","",Laboratorio!B{r})', f_calc_1),
        "H": (f'=IF(A{r}="","",COUNT(D{r}:G{r}))', f_calc),
        "I": (f'=IF(A{r}="","",IF(H{r}=0,"",AVERAGE(D{r}:G{r})))', f_calc_1),
        "J": (f'=IF(OR(I{r}="",C{r}=""),"",I{r}-C{r})', f_calc_1),
        "K": (f'=IF(A{r}="","",IF(Laboratorio!E{r}<>4,"Horário fora do padrão",IF(H{r}=0,"Sem resultado online",'
              f'IF(H{r}<4,"Composto com "&H{r}&" cortes",""))))', f_calc),
    }
    for k, col in enumerate("DEFG", 1):
        f[col] = (f'=IF(A{r}="","",IFERROR(AVERAGEIFS({OB},{OE},L{r},{OF},{k},{OG},1),""))', f_calc_1)
    for col, (formula, fm) in f.items():
        wc.write_formula(f"{col}{r}", formula, fm)

# ======================================================================= Painel
wp.hide_gridlines(2)
wp.set_landscape()
wp.fit_to_pages(1, 0)
for col, w in zip("ABCDEF", [2, 48, 20, 4, 40, 20]):
    wp.set_column(f"{col}:{col}", w)
wp.set_row(0, 40)
for c in range(7):
    wp.write_blank(0, c, None, fmt(bg_color=AZUL))
    wp.write_blank(1, c, None, fmt(bg_color=AMARELO))
wp.set_row(1, 4)
wp.write(0, 1, "Compressão online x laboratório", fmt(bold=True, font_size=16, font_color="white", bg_color=AZUL))
wp.write("B3", "Primeira vez: instalar a macro (2 minutos)", f_sec)
passos = ["1. Salve este arquivo na sua máquina e abra no Excel com o suplemento Aspen ativo (o mesmo do Relatório Físico).",
          "2. Pressione Alt+F11. No editor, menu Arquivo > Importar arquivo, escolha ModBusca.bas e feche o editor.",
          "3. Menu Arquivo > Salvar como, tipo Pasta de Trabalho Habilitada para Macro do Excel (.xlsm). Use sempre o .xlsm daqui em diante.",
          "4. Clique em Testar conexão. Se aparecer alguma mensagem de erro, tire uma foto dela e da aba Log."]
for k, t in enumerate(passos):
    wp.merge_range(3 + k, 1, 3 + k, 5, t, f_txt)
wp.write("B9", "Período da busca", f_sec)
wp.write("B10", "Início", f_rot)
wp.write_datetime("C10", INICIO, f_ent_dh)
nome("pnlInicio", wp, "$C$10")
wp.write("B11", "Fim (em branco = agora)", f_rot)
wp.write_blank("C11", None, f_ent_dh)
nome("pnlFim", wp, "$C$11")
for k, (macro, rotulo) in enumerate([("TestarConexao", "Testar conexão"), ("BuscarDados", "Buscar dados"), ("LimparDados", "Limpar dados")]):
    wp.insert_button("E9", {"macro": macro, "caption": rotulo, "width": 120, "height": 40, "x_offset": k * 128, "y_offset": 4})
wp.merge_range("B13:F13", "Sem dados.", f_status)
nome("pnlStatus", wp, "$B$13")
wp.write("B15", "Resumo", f_sec)
resumo = [("Resultados da compressão online", f"=COUNT(Online!$B$2:$B${NP})", "0"),
          ("Resultados do laboratório", f"=COUNT(Laboratorio!$B$2:$B${NM})", "0"),
          ("Resultados do laboratório com composto online", f"=COUNT(Comparacao!$J$2:$J${NM})", "0"),
          ("Média do laboratório (kgf/pelota)", f'=IFERROR(AVERAGEIFS(Comparacao!$C$2:$C${NM},Comparacao!$J$2:$J${NM},"<>"),"")', "0.0"),
          ("Média do composto online (kgf/pelota)", f'=IFERROR(AVERAGEIFS(Comparacao!$I$2:$I${NM},Comparacao!$J$2:$J${NM},"<>"),"")', "0.0"),
          ("Diferença média, online menos laboratório (kgf/pelota)", f"=IFERROR(AVERAGE(Comparacao!$J$2:$J${NM}),\"\")", "0.0")]
for k, (rot, form, num) in enumerate(resumo):
    wp.write(15 + k, 1, rot, fmt(bg_color="#EEF3F7"))
    wp.write_formula(15 + k, 2, form, fmt(bold=True, font_color=AZUL, bg_color="#EEF3F7", align="right", num_format=num))
wp.write("B23", "O detalhe de cada resultado do laboratório, com os 4 cortes online, está na aba Comparacao. "
         "Os gráficos e a estatística entram na próxima etapa, depois que a busca estiver funcionando.", f_status)
wp.activate()

if COM_VBA:   # só para a conferência automática
    from vba_projeto import vba_project_bin
    modulo = open(os.path.join(AQUI, "ModBusca.bas"), encoding="cp1252").read().replace("\r\n", "\n").replace("\n", "\r\n")
    if os.environ.get("VBA_TESTE"):
        modulo += open(os.environ["VBA_TESTE"], encoding="utf-8").read().replace("\n", "\r\n")
    binario = os.path.join(AQUI, "vbaProject.bin")
    open(binario, "wb").write(vba_project_bin({"ModBusca": modulo}, {"ThisWorkbook": ""}))
    wb.add_vba_project(binario)
    wb.set_vba_name("ThisWorkbook")
wb.close()
if COM_VBA:
    os.remove(binario)
print(SAIDA, f"| online {len(online)} | laboratório {len(lab)}")
