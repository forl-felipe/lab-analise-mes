"""Gera Monitoramento_SE_Espessador_Batch.xlsm (Excel com macros).

Planilha automática do report "S.E. - Espessador x Batch" (antes: Monitoramento SE - Após fator ESP 6.2.xlsx):
o botão Buscar no MES traz os resultados do dia (Aspen IP.21, GetCalculationValues), grava no histórico
e o Painel (mesmo padrão visual do Comparativo Germano x Ubu) mostra os últimos N dias.

Uso:  python3 build_monitor.py [--teste]
  --teste  acrescenta um modulo de teste automatico (validacao com LibreOffice)
"""
import os
import re
import sys
import json
import zipfile
import datetime as dtm

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as rc

from vba_project import build_vba_project

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
VBA_DIR = os.path.join(RAIZ, 'vba')
TESTE = '--teste' in sys.argv
SAIDA = os.path.join(RAIZ, 'Monitoramento_SE_Espessador_Batch%s.xlsm' % ('_TESTE' if TESTE else ''))
LOGO = os.path.join(AQUI, 'assets', 'logo_samarco.png')   # kit visual Samarco (fundo escuro)
HISTORICO = os.path.join(AQUI, 'historico.json')

# ---------------------------------------------------------------- paleta Samarco (igual ao Comparativo)
NAVY = '#00335A'         # Azul Samarco: fundo principal, cabecalhos e barras de titulo
NAVY2 = '#004070'        # Azul medio: paineis secundarios e subtitulos
AZUL_TIT = '#003B6C'     # Azul de titulo: titulos sobre fundo branco
AZUL_PROF = '#002643'    # Azul profundo: contraste e rodapes
AMARELO = '#FFC000'      # Amarelo: destaques e linhas de apoio
AZUL_CLARO = '#59C6F2'   # Azul claro: setas, fluxos e icones
LARANJA = '#F37021'      # Laranja: alertas e valores em evidencia
CINZA_LOGO = '#B5BEC4'   # Cinza da marca
AZUL_CINZA = '#61889B'   # Azul acinzentado: contornos e linhas de graficos
INK = '#1D2B36'
INK2 = '#4A5A66'
MUTED = '#5B6B77'
LINE = '#C9D1D7'
LINE2 = '#DDE3E8'
BG = '#F3F5F7'
SEC = '#F3F5F7'
TH = '#E6EBEF'
OK = '#0A7C50'
NOK = '#C62828'
NOK_ROW = '#FDF3F3'
COR_A = '#A3D9F3'        # barras da serie A (espessador ponderado): azul claro suave
COR_A_MEDIA = NAVY2
COR_B = '#D5DBDF'        # barras da serie B (batch EB6): cinza da marca suave
COR_B_MEDIA = AZUL_CINZA
COR_C = '#B9C3CB'        # 3a serie (TK 02)
COR_C_MEDIA = '#7D8C98'
COR_DIF = LARANJA
COR_DIF_TXT = '#B04F12'
INPUT = '#FFF8E1'
BRANCO = '#FFFFFF'
FONTE = 'Segoe UI'
SENHA_CFG = 'Samarco@Lab'

# ----------------------------------------------------------------------------
# Itens buscados no MES (ordem da aba Dia e da aba Configuracoes)
#   chave, rotulo, tag, mapa, janela, agregacao, faixa valida, decimais, de onde vem no MES
#   janela E = espessador D-1 23:00 a D 21:00 · D = dia 00:00 a 23:59
#   tag e mapa: os que o Diagnostico do MES achou com valor em 29/10/2026 (12 de 12); se o MES mudar, o Diagnostico
#   testa as outras formas da tag e os dois mapas e grava em Configuracoes o que trouxer valor
#   tags e horarios conferidos pelo laboratorio no MES (relatorio_alessandra.xlsx, 26/10/2026)
# ----------------------------------------------------------------------------
ITENS = [
    ('prodC3', 'Produção C3 (t)', 'M4239010000-596-HHCC', 'IP_MESVALOR', 'E', 'Média', 0, 200000, 0,
     '15-G2-HH INDICES PRODUCAO - Ritmo Produção Concentrador III (conferência 26/10)'),
    ('prodC2', 'Produção C2 (t)', 'M0239010000-0596-HHCC', 'IP_AnalogMap', 'E', 'Média', 0, 200000, 0,
     '15-G2-HH INDICES PRODUCAO - Ritmo Produção Concentrador II'),
    ('espC3', 'S.E. alimentação espessador C3', 'M4238010000-2888-HHLQ', 'IP_AnalogMap', 'E', 'Média', 300, 8000,
     0, '40-G3-HH ALIMENTACAO ESPESSADO - SE'),
    ('espC2', 'S.E. alimentação espessador C2', 'M0238010000-0017-HHLQ', 'IP_AnalogMap', 'E', 'Média', 300, 8000,
     0, '40-G2-HH ALIMENTACAO ESPESSADO - SE'),
    ('eb4', 'S.E. batch EB4', 'M250010000-17-ABLQ', 'IP_MESVALOR', 'D', 'Média', 300, 8000, 0,
     '50-G2-HH AMOSTRAGEM BATCH - SE'),
    ('eb6', 'S.E. batch EB6', 'M4250010000-17-ABLQ', 'IP_MESVALOR', 'D', 'Média', 300, 8000, 0,
     '50-G3-HH AMOSTRAGEM BATCH - SE'),
    ('tk01', 'S.E. TK 01 - EB6', 'M239040000-17-TBLQ', 'IP_MESVALOR', 'D', 'Média', 300, 8000, 0,
     'Tk1 U2 - SE (239040000-17-TB-LQ, cm²/g)'),
    ('tk02', 'S.E. TK 02 - EB6', 'M239050000-17-TBLQ', 'IP_MESVALOR', 'D', 'Média', 300, 8000, 0,
     'Tk2 U2 - SE (239050000-17-TB-LQ, cm²/g)'),
    ('lamaC3', 'Lama incorporada C3 (tms)', 'M4239000000-0781-DDDVR', 'IP_AnalogMap', 'D', 'Último', 0, 20000, 0,
     'BOLETIM DVR - Lama incorporada no concentrado'),
    ('lamaC2', 'Lama incorporada C2 (tms)', 'M0239000000-0781-DDDVR', 'IP_AnalogMap', 'D', 'Último', 0, 20000, 0,
     '00-G2-DD BALANCO DVR - Lama incorporada (soma com a do C3)'),
    ('concC3', 'Concentrado C3 (tms)', 'M4238010000-1445-DDDVR', 'IP_AnalogMap', 'D', 'Último', 0, 200000, 0,
     'BOLETIM DVR - Beneficiamento - Concentrado C3 (tms) (conferência 26/10)'),
    ('concC2', 'Concentrado C2 (tms)', 'M0238010000-1445-DDDVR', 'IP_AnalogMap', 'D', 'Último', 0, 200000, 0,
     'BOLETIM DVR - Esp. Conc DVR (tms)'),
]
NI = len(ITENS)
IK = [i[0] for i in ITENS]
SERVIDOR_PADRAO = 'UBU'
# janelas de 1 h: 1o horario (horas a partir de D 00:00) e quantos horarios
#   E: D-1 23:00 ... D 21:00 (23 horarios; caderno e Word) - defasagem de 2 h em relacao ao batch
#   D: D 01:00 ... D 23:00 e o fechamento 23:59 (24 janelas de 1 h a partir de 00:30)
JANELAS = {'E': (-1, 23, 'Espessador: D-1 23:00 a D 21:00'), 'D': (1, 24, 'Dia: D 00:00 a D 23:59')}
FATOR_PADRAO = 160
TOL_PADRAO = 150        # tolerancia da diferenca de S.E. (mesma do Comparativo; confirmar com o laboratorio)
DIAS_REPORT = 15
NMAXD = 60              # maximo de dias no Painel

# ----------------------------------------------------------------------------
# Historico (aba Histórico): cabecalho na linha H_ROW_HDR (0-based), dados a partir da linha seguinte
# ----------------------------------------------------------------------------
H_ROW_HDR = 5
H_NLIN = 3000
H_COLS = [  # chave, titulo, largura, formato, formula (None = valor)
    ('data', 'Data', 11, 'data', None),
    ('prodC3', 'Produção C3 (t)', 11, 'int', None),
    ('prodC2', 'Produção C2 (t)', 11, 'int', None),
    ('espC3', 'S.E. ESP C3', 9, 'int', None),
    ('espC2', 'S.E. ESP C2', 9, 'int', None),
    ('espP', 'S.E. ESP C2 + C3 (ponderada)', 12, 'int',
     '=IF(OR({espC3}="",{espC2}="",{prodC3}="",{prodC2}=""),"",IF({prodC3}+{prodC2}=0,"",'
     '({espC3}*{prodC3}+{espC2}*{prodC2})/({prodC3}+{prodC2})))'),
    ('eb4', 'S.E. batch EB4', 9, 'int', None),
    ('eb4P', 'S.E. batch EB4 + ESP C3 (ponderada)', 12, 'int',
     '=IF(OR({eb4}="",{espC3}="",{prodC3}="",{prodC2}=""),"",IF({prodC3}+{prodC2}=0,"",'
     '({eb4}*{prodC2}+{espC3}*{prodC3})/({prodC3}+{prodC2})))'),
    ('eb6', 'S.E. batch EB6', 9, 'int', None),
    ('difJ', 'Diferença EB6 − (EB4 + ESP C3)', 12, 'sinal', '=IF(OR({eb6}="",{eb4P}=""),"",{eb6}-{eb4P})'),
    ('difK', 'Diferença EB6 − (ESP C2 + C3)', 12, 'sinal', '=IF(OR({eb6}="",{espP}=""),"",{eb6}-{espP})'),
    ('tk01', 'TK 01 - EB6', 9, 'int', None),
    ('tk02', 'TK 02 - EB6', 9, 'int', None),
    ('fator', 'Fator', 7, 'int', None),
    ('lamaC3', 'Lama incorp. C3 (tms)', 10, 'int', None),
    ('lamaC2', 'Lama incorp. C2 (tms)', 10, 'int', None),
    ('lama', 'Lama incorporada (tms)', 11, 'int', '=IF(AND({lamaC3}="",{lamaC2}=""),"",N({lamaC3})+N({lamaC2}))'),
    ('concC3', 'Concentrado C3 (tms)', 11, 'int', None),
    ('concC2', 'Concentrado C2 (tms)', 11, 'int', None),
    ('pctLama', '% lama', 8, 'pct',
     '=IF(OR({lama}="",{concC3}="",{concC2}=""),"",IF({concC3}+{concC2}=0,"",{lama}/({concC3}+{concC2})))'),
    ('atual', 'Atualizado em', 15, 'datahora', None),
    ('obs', 'Observação', 30, 'txt', None),
]
H_COL1 = 1                       # coluna B
HC = {k: H_COL1 + i for i, (k, *_r) in enumerate(H_COLS)}

# Aba Dia (0-based) - mesmo padrao da aba Batch do Comparativo
D_ROW_DATA = 4
D_ROW_HDR = 9
D_ROW_IT0 = D_ROW_HDR + 1        # 1o item
D_ROW_FATOR = D_ROW_IT0 + NI
D_ROW_OBS = D_ROW_FATOR + 1
D_ROW_LEG = D_ROW_OBS + 2        # legenda D / D-1
D_ROW_RES = D_ROW_OBS + 5        # resultado do dia (calculado)
D_ROW_COM = D_ROW_RES + 8        # comentarios do painel e tolerancia
D_ROW_ST = D_ROW_COM + 10        # ultima busca
D_COL = dict(ITEM=1, VAL=2, NRES=3, JAN=4)

# Aba Diagnostico (0-based)
X_ROW_HDR = 7
X_ROW0 = X_ROW_HDR + 1
X_COLS = ['Item', 'Tag', 'Mapa', 'Resultado', 'Valor do dia', 'Horários com valor (D-1 12:00 a D+1 12:00)']
X_LARG = [32, 26, 14, 34, 13, 95]
X_COL = dict(TAG=2, MAPA=3, RES=4, VAL=5, LISTA=6)     # 0-based

# Aba Configuracoes (0-based)
C_ROW_IT0 = 7
C_COL = dict(CHAVE=1, ROT=2, TAG=3, SERV=4, MAPA=5, JAN=6, AGR=7, VMIN=8, VMAX=9, DEC=10, ORIG=11)

# Aba _Graf (0-based): linha 0 = titulos; k = 1..NMAXD+1 nas linhas 1..NMAXD+1
#   k <= rDias: dia; k = rDias + 1: "Média" (barras escuras e ◆ diferenca media, como no Comparativo)
G_COLS = ['data', 'rot', 'espP', 'eb6', 'difK', 'difKmed', 'tolP', 'tolN', 'eb4P', 'difJ', 'difJmed',
          'tk01', 'tk02', 'pctLama', 'pctMed', 'pct0']
GC = {k: i for i, k in enumerate(G_COLS)}
G_DIA = ['espP', 'eb6', 'difK', 'eb4P', 'difJ', 'tk01', 'tk02', 'pctLama']   # valor do dia (historico)
G_MEDIA = ['espP', 'eb6', 'eb4P', 'tk01', 'tk02']                            # barras: media na linha Média
G_ROW_ST = NMAXD + 4             # 1a linha das estatisticas (0-based)
ESTAT = ['espP', 'eb6', 'difK', 'eb4P', 'difJ', 'tk01', 'tk02', 'pctLama', 'lama', 'prodC3', 'prodC2', 'concC3',
         'concC2', 'fator', 'obs']
# series de cada grafico (ordem do SeriesCollection: barras e depois linhas) e cores das barras (base, media)
GRAFICOS = [
    ('g1', 'Espessador C2 + C3 × Batch EB6', ['espP', 'eb6'], ['difK', 'difKmed', 'tolP', 'tolN'], 'difK'),
    ('g2', 'Batch EB4 + Espessador C3 × Batch EB6', ['eb4P', 'eb6'], ['difJ', 'difJmed', 'tolP', 'tolN'], 'difJ'),
    ('g3', '% de lama no batch × S.E. do batch EB6 e dos tanques', ['eb6', 'tk01', 'tk02'], ['pctLama', 'pctMed'],
     None),
]
NOMES_SERIE = {'espP': 'ESP C2 + C3 (ponderada)', 'eb6': 'Batch EB6', 'eb4P': 'EB4 + ESP C3 (ponderada)',
               'tk01': 'TK 01 - EB6', 'tk02': 'TK 02 - EB6', 'difK': 'Diferença (EB6 − ESP C2 + C3)',
               'difJ': 'Diferença (EB6 − EB4 + ESP C3)', 'difKmed': 'Diferença média', 'difJmed': 'Diferença média',
               'tolP': 'Tolerância (±)', 'tolN': 'Tolerância (±)', 'pctLama': '% lama', 'pctMed': '% lama média'}

# Painel: colunas em pixels (as mesmas do Comparativo) e 1a linha do report
P_COLS = [10, 84, 70, 70, 92, 62, 62, 12, 122, 12, 246, 12, 120, 196, 10]
P_R0 = 5


def fmt_sinal(milhar=True):
    b = '#,##0' if milhar else '0'
    return '+%s;"−"%s;%s' % (b, b, b)


def dm(x):
    """dd/mm em formula, sem codigo de formato do TEXT (que muda com o idioma do Excel: yyyy x aaaa)."""
    return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")' % (x, x)


def dmy(x):
    return '%s&"/"&YEAR(%s)' % (dm(x), x)


# textos de numeros em formulas (FIXED segue o idioma do Excel: 2.210 / 0,83%)
def tx0(x):
    return 'IF(ISNUMBER(%s),FIXED(%s,0),"—")' % (x, x)


def txs(x):
    return 'IF(ISNUMBER(%s),IF(%s>0,"+","")&FIXED(%s,0),"—")' % (x, x, x)


def txp(x):
    return 'IF(ISNUMBER(%s),FIXED(%s*100,2)&"%%","—")' % (x, x)


def quote(ws):
    return "'%s'" % ws.name


class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_size(1500, 950)
        self.wb.set_calc_mode('auto')
        self.fmts = {}
        self.wb.set_properties({'title': 'Monitoramento S.E. - Espessador x Batch', 'company': 'Samarco - Laboratório'})
        self.hoje = dtm.date.today()
        with open(HISTORICO, encoding='utf-8') as fh:
            self.hist = sorted(json.load(fh), key=lambda d: d['data'])
        if TESTE:       # o teste automatico usa a base fixa ate 15/10/2026 (dias seguintes sao simulados)
            self.hist = [d for d in self.hist if d['data'] <= '2026-10-15']
        self.dia0 = self.hist[-1] if self.hist else {}      # a aba Dia abre com o ultimo dia do historico
        self.ultimo_dia = (dtm.datetime.strptime(self.hist[-1]['data'], '%Y-%m-%d') if self.hist
                           else dtm.datetime.combine(self.hoje, dtm.time()))

    # ----------------------------------------------------------- formatos
    def f(self, **kw):
        base = {'font_name': FONTE, 'font_size': 10, 'font_color': INK, 'valign': 'vcenter'}
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self.fmts:
            self.fmts[chave] = self.wb.add_format(base)
        return self.fmts[chave]

    def fi(self, **kw):
        base = dict(bg_color=INPUT, border=1, border_color=LINE, locked=False)
        base.update(kw)
        return self.f(**base)

    def nome(self, nome, ws, row, col):
        self.wb.define_name(nome, '=%s!%s' % (quote(ws), rc(row, col, True, True)))

    def botao(self, ws, row, col, texto, macro, largura=150, altura=30, estilo='p', x=0, y=0):
        estilos = {'a': (AMARELO, AZUL_PROF, None), 'p': (NAVY2, BRANCO, None), 'g': (OK, BRANCO, None),
                   's': ('#E8EBF0', INK2, None), 'n': (NAVY, BRANCO, None)}
        fundo, cor, borda = estilos[estilo]
        ws.insert_textbox(row, col, texto, {
            'width': largura, 'height': altura, 'x_offset': x, 'y_offset': y,
            'font': {'name': FONTE, 'size': 10, 'bold': True, 'color': cor},
            'align': {'vertical': 'middle', 'horizontal': 'center'},
            'fill': {'color': fundo},
            'line': {'color': borda, 'width': 1} if borda else {'none': True},
            'description': 'macro:' + macro,
            'object_position': 3,
        })

    def faixa(self, ws, r, c1, c2, titulo, sub=None):
        """Faixa azul-marinho no topo da aba (igual as abas do Comparativo)."""
        ws.set_row(r, 34)
        ws.merge_range(r, c1, r, c2, titulo, self.f(bold=True, font_size=14, font_color=BRANCO, bg_color=NAVY,
                                                  indent=1, bottom=5, bottom_color=AMARELO))
        if sub is not None:
            ws.set_row(r + 1, 22)
            ws.merge_range(r + 1, c1, r + 1, c2, sub, self.f(font_size=9.5, bold=True, font_color='#D5DEEB',
                                                           bg_color=NAVY2, indent=1))

    def secao(self, ws, r, c1, c2, texto, altura=22):
        ws.set_row(r, altura)
        fm = self.f(bold=True, font_size=10, font_color=AZUL_TIT, valign='bottom', bottom=2, bottom_color=AMARELO)
        if c1 == c2:
            ws.write(r, c1, texto, fm)
        else:
            ws.merge_range(r, c1, r, c2, texto, fm)

    def rotulo(self, ws, r, c1, c2, texto):
        fm = self.f(bold=True, font_size=8, font_color=MUTED, valign='bottom')
        if c2 > c1:
            ws.merge_range(r, c1, r, c2, texto.upper(), fm)
        else:
            ws.write(r, c1, texto.upper(), fm)

    def num_fmt(self, tipo):
        return {'data': 'dd/mm/yyyy', 'int': '#,##0', 'sinal': fmt_sinal(), 'pct': '0.00%',
                'datahora': 'dd/mm/yyyy hh:mm', 'txt': '@'}[tipo]

    # ------------------------------------------------------------------------
    def construir(self):
        wb = self.wb
        self.ws_pai = wb.add_worksheet('Painel')
        self.ws_dia = wb.add_worksheet('Dia')
        self.ws_his = wb.add_worksheet('Histórico')
        self.ws_dia_x = wb.add_worksheet('Diagnóstico')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_gra = wb.add_worksheet('_Graf')
        for ws, cn in [(self.ws_pai, 'shPainel'), (self.ws_dia, 'shDia'), (self.ws_his, 'shHist'),
                       (self.ws_dia_x, 'shDiag'), (self.ws_cfg, 'shConfig'), (self.ws_mes, 'shDadosMES'),
                       (self.ws_gra, 'shGraf')]:
            ws.set_vba_name(cn)
            if ws not in (self.ws_gra, self.ws_mes):
                ws.hide_gridlines(2)
        self.ws_pai.set_tab_color(NAVY2)
        self.ws_dia.set_tab_color('#FDB913')
        self.ws_his.set_tab_color(OK)
        self.ws_dia_x.set_tab_color(MUTED)
        self.ws_cfg.set_tab_color('#9AA3AD')
        self.aba_config()
        self.aba_historico()
        self.aba_dia()
        self.aba_graf()
        self.aba_painel()
        self.aba_diagnostico()
        self.aba_dados_mes()
        self.ws_mes.hide()
        self.ws_gra.hide()
        self.ws_cfg.hide()
        self.ws_dia.activate()

    # ------------------------------------------------------------------------ Configuracoes
    def aba_config(self):
        ws = self.ws_cfg
        ws.set_column(0, 0, 1.5)
        larg = {C_COL['CHAVE']: 9, C_COL['ROT']: 30, C_COL['TAG']: 28, C_COL['SERV']: 8, C_COL['MAPA']: 14,
                C_COL['JAN']: 7, C_COL['AGR']: 9, C_COL['VMIN']: 8, C_COL['VMAX']: 9, C_COL['DEC']: 8,
                C_COL['ORIG']: 62}
        for c, w in larg.items():
            ws.set_column(c, c, w)
        self.faixa(ws, 1, 1, C_COL['ORIG'], 'CONFIGURAÇÕES', 'Tags do MES, janelas de horário e padrões do painel')
        self.botao(ws, 3, C_COL['ORIG'], 'Voltar', 'FecharConfiguracoes', largura=110, altura=26, estilo='s', x=300)
        self.botao(ws, 3, C_COL['ORIG'], 'Ver Dados_MES', 'MostrarDadosMES', largura=130, altura=26, estilo='s',
                   x=150)
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center', font_size=9)
        cel = self.fi()
        celc = self.fi(align='center')
        fixo = self.f(border=1, border_color=LINE, bg_color=BG, font_color=MUTED)
        r0 = C_ROW_IT0 - 2
        self.secao(ws, r0, 1, C_COL['ORIG'], 'ITENS BUSCADOS NO MES  (janela E = espessador D−1 23:00 a D 21:00 · '
                                              'D = dia 00:00 a 23:59 · item sem tag = lançar à mão na aba Dia)')
        tit = {'CHAVE': 'Chave', 'ROT': 'Item', 'TAG': 'Tag', 'SERV': 'Servidor', 'MAPA': 'Mapa', 'JAN': 'Janela',
               'AGR': 'Agregação', 'VMIN': 'Mín. válido', 'VMAX': 'Máx. válido', 'DEC': 'Decimais',
               'ORIG': 'Origem no MES (tela / receita / item)'}
        ws.set_row(r0 + 1, 30)
        for k, c in C_COL.items():
            ws.write(r0 + 1, c, tit[k], th)
        for i, it in enumerate(ITENS):
            r = C_ROW_IT0 + i
            ws.write(r, C_COL['CHAVE'], it[0], fixo)
            ws.write(r, C_COL['ROT'], it[1], cel)
            ws.write_string(r, C_COL['TAG'], it[2], cel)
            ws.write(r, C_COL['SERV'], SERVIDOR_PADRAO, celc)
            ws.write(r, C_COL['MAPA'], it[3], celc)
            ws.write(r, C_COL['JAN'], it[4], celc)
            ws.write(r, C_COL['AGR'], it[5], celc)
            ws.write(r, C_COL['VMIN'], it[6], celc)
            ws.write(r, C_COL['VMAX'], it[7], celc)
            ws.write(r, C_COL['DEC'], it[8], celc)
            ws.write(r, C_COL['ORIG'], it[9], cel)
        r1 = C_ROW_IT0 + NI
        ws.data_validation(C_ROW_IT0, C_COL['JAN'], r1 - 1, C_COL['JAN'], {'validate': 'list', 'source': ['E', 'D']})
        ws.data_validation(C_ROW_IT0, C_COL['AGR'], r1 - 1, C_COL['AGR'],
                           {'validate': 'list', 'source': ['Soma', 'Média', 'Último']})
        ws.data_validation(C_ROW_IT0, C_COL['MAPA'], r1 - 1, C_COL['MAPA'],
                           {'validate': 'list', 'source': ['IP_AnalogMap', 'IP_MESVALOR', 'IP_DiscreteMap']})
        ws.write(r1, 1, 'Item sem valor: aba Diagnóstico → Rodar diagnóstico. Ele testa as formas da tag (-0596- ↔ -596-, M0… ↔ M…) '
                        'com IP_MESVALOR e IP_AnalogMap e grava aqui a combinação que trouxe valor.',
                 self.f(font_size=8.5, italic=True, font_color=MUTED))

        # janelas
        r = r1 + 2
        self.secao(ws, r, 1, 5, 'JANELAS DE HORÁRIO  (consultas de 1 h: o resultado das 01:00 cai em 00:30–01:30)')
        ws.write(r + 1, 1, 'Janela', th)
        ws.write(r + 1, 2, 'Descrição', th)
        ws.write(r + 1, 3, '1º horário (h a partir de D 00:00)', th)
        ws.write(r + 1, 4, 'Nº de horários', th)
        ws.set_row(r + 1, 30)
        for j, (k, (h1, n, desc)) in enumerate(JANELAS.items()):
            ws.write(r + 2 + j, 1, k, fixo)
            ws.write(r + 2 + j, 2, desc, fixo)
            ws.write(r + 2 + j, 3, h1, celc)
            ws.write(r + 2 + j, 4, n, celc)
            self.nome('cfgJan%sH1' % k, ws, r + 2 + j, 3)
            self.nome('cfgJan%sN' % k, ws, r + 2 + j, 4)

        # parametros gerais
        r = r + 6
        self.secao(ws, r, 1, 5, 'PARÂMETROS')
        params = [
            ('cfgFonte', 'Fonte dos dados (MES ou SIMULAÇÃO para treinar sem MES)', 'MES'),
            ('cfgFatorPadrao', 'Fator padrão de um dia novo (se não houver dia anterior no histórico)', FATOR_PADRAO),
            ('cfgDiasPadrao', 'Dias no painel (padrão ao abrir)', DIAS_REPORT),
            ('cfgTimeout', 'Tempo máximo de espera de cada consulta (s)', 120),
            ('cfgCalc', 'Tipo de cálculo do GetCalculationValues (1 = média da janela, igual aos relatórios)', '1'),
            ('cfgFormatoData', 'Formato da data na consulta (DATA, TEXTO ou MM/DD)', 'DATA'),
            ('mesPrefixo', 'Forma da função do Aspen reconhecida pelo Excel (preenchida sozinha)', ''),
        ]
        for j, (nm, desc, v) in enumerate(params):
            ws.merge_range(r + 1 + j, 1, r + 1 + j, 2, desc, fixo)
            if isinstance(v, str):
                ws.write_string(r + 1 + j, 3, v, celc)
            else:
                ws.write(r + 1 + j, 3, v, celc)
            self.nome(nm, ws, r + 1 + j, 3)
        ws.data_validation(r + 1, 3, r + 1, 3, {'validate': 'list', 'source': ['MES', 'SIMULAÇÃO']})
        ws.data_validation(r + 6, 3, r + 6, 3, {'validate': 'list', 'source': ['DATA', 'TEXTO', 'MM/DD']})
        ws.protect(SENHA_CFG, {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------------------ Historico
    def aba_historico(self):
        ws = self.ws_his
        ws.set_column(0, 0, 1.5)
        for i, (k, t, w, tipo, form) in enumerate(H_COLS):
            ws.set_column(H_COL1 + i, H_COL1 + i, w)
        ultc = H_COL1 + len(H_COLS) - 1
        self.faixa(ws, 1, 1, ultc, 'HISTÓRICO DIÁRIO',
                   'Um dia por linha. Colunas azuis: MES ou lançamento · cinza: calculadas (não editar)')
        ws.write(3, 1, 'Gravado pela aba Dia (Buscar no MES / Salvar). Pode corrigir à mão as colunas de valor '
                       '(Fator, Observação...). Base de 28/01/2026 em diante: planilha anterior + buscas no MES. Dá para lançar um dia à mão aqui (data + valores).',
                 self.f(font_color=MUTED, italic=True, font_size=9))
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center', font_size=9)
        thc = self.f(bold=True, font_color=BRANCO, bg_color=AZUL_CINZA, border=1, border_color=LINE,
                     text_wrap=True, align='center', font_size=9)
        ws.set_row(H_ROW_HDR, 42)
        for i, (k, t, w, tipo, form) in enumerate(H_COLS):
            ws.write(H_ROW_HDR, H_COL1 + i, t, thc if form else th)
        val = {tipo: self.f(num_format=self.num_fmt(tipo), border=1, border_color=LINE2,
                            align='left' if tipo == 'txt' else 'center', font_size=9.5)
               for tipo in ('data', 'int', 'sinal', 'pct', 'datahora', 'txt')}
        calc = {tipo: self.f(num_format=self.num_fmt(tipo), border=1, border_color=LINE2, bg_color=BG,
                             align='center', font_color=INK2, font_size=9.5) for tipo in ('int', 'sinal', 'pct')}
        for n in range(H_NLIN):
            r = H_ROW_HDR + 1 + n
            reg = self.hist[n] if n < len(self.hist) else None
            refs = {k: rc(r, H_COL1 + i) for i, (k, *_x) in enumerate(H_COLS)}
            for i, (k, t, w, tipo, form) in enumerate(H_COLS):
                c = H_COL1 + i
                if form:
                    ws.write_formula(r, c, form.format(**refs), calc[tipo], '')
                elif reg is not None:
                    if k == 'data':
                        ws.write_datetime(r, c, dtm.datetime.strptime(reg['data'], '%Y-%m-%d'), val['data'])
                    elif k == 'atual' and reg.get('atual'):
                        ws.write_datetime(r, c, dtm.datetime.fromisoformat(reg['atual']), val['datahora'])
                    elif k == 'obs' and reg.get('obs'):
                        ws.write_string(r, c, str(reg['obs']), val['txt'])
                    elif k in reg and isinstance(reg[k], (int, float)):
                        ws.write_number(r, c, reg[k], val[tipo])
                    else:
                        ws.write_blank(r, c, None, val[tipo])
                else:
                    ws.write_blank(r, c, None, val[tipo])
        ws.freeze_panes(H_ROW_HDR + 1, H_COL1 + 1)
        ws.autofilter(H_ROW_HDR, H_COL1, H_ROW_HDR + len(self.hist), ultc)
        for k, c in HC.items():
            self.wb.define_name('h_' + k, "=%s!%s:%s" % (quote(ws), rc(H_ROW_HDR + 1, c, True, True),
                                                        rc(H_ROW_HDR + H_NLIN, c, True, True)))

    # ------------------------------------------------------------------------ Dia
    def aba_dia(self):
        ws = self.ws_dia
        ws.set_column(0, 0, 1.5)
        larg = {'ITEM': 40, 'VAL': 14, 'NRES': 11, 'JAN': 46}
        for k, w in larg.items():
            ws.set_column(D_COL[k], D_COL[k], w)
        ultc = D_COL['JAN']
        ws.set_row(0, 8)
        ws.set_row(1, 34)
        ws.merge_range(1, 1, 1, ultc, 'DADOS DO DIA', self.f(bold=True, font_size=14, font_color=BRANCO,
                                                             bg_color=NAVY, indent=1, bottom=5,
                                                             bottom_color=AMARELO))
        self.botao(ws, 1, D_COL['NRES'], 'Ver painel →', 'IrPainel', largura=110, altura=24, estilo='p', x=0, y=5)
        self.botao(ws, 1, D_COL['JAN'], 'Diagnóstico do MES', 'IrDiagnostico', largura=140, altura=24, estilo='s',
                   x=40, y=5)
        self.botao(ws, 1, D_COL['JAN'], 'Configurações', 'AbrirConfiguracoes', largura=120, altura=24, estilo='s',
                   x=190, y=5)

        # (1) data + botoes
        self.secao(ws, 3, 1, 2, '①  DIA DOS DADOS')
        ws.set_row(D_ROW_DATA, 30)
        ws.write(D_ROW_DATA, D_COL['ITEM'], 'Dia dos dados (D) ▸', self.f(bold=True, font_color=AZUL_TIT,
                                                                             align='right'))
        ent = self.fi(bold=True, font_size=13, align='center', num_format='dd/mm/yyyy', border=2,
                      border_color=AMARELO)
        ws.write_datetime(D_ROW_DATA, D_COL['VAL'], self.ultimo_dia, ent)
        self.nome('dData', ws, D_ROW_DATA, D_COL['VAL'])
        self.botao(ws, 3, D_COL['NRES'], 'Buscar no MES', 'BuscarMES', largura=170, altura=34, estilo='a', x=4, y=6)
        self.botao(ws, 3, D_COL['JAN'], 'Salvar no histórico', 'SalvarDia', largura=180, altura=34, estilo='g',
                   x=104, y=6)
        ws.write(D_ROW_DATA + 1, D_COL['ITEM'], 'Buscar no MES já grava o dia no histórico e no painel. Salvar no '
                                                 'histórico: depois de corrigir ou lançar um valor à mão.',
                 self.f(font_color=MUTED, italic=True, font_size=9))

        # (2) resultados
        self.secao(ws, D_ROW_HDR - 1, 1, ultc, '②  RESULTADOS DO MES  (células amarelas: dá para corrigir · o Fator '
                                               'é lançado à mão)')
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center', font_size=9)
        ws.set_row(D_ROW_HDR, 30)
        for c, t in [(D_COL['ITEM'], 'Item'), (D_COL['VAL'], 'Valor'), (D_COL['NRES'], 'Nº de resultados'),
                     (D_COL['JAN'], 'Janela no MES')]:
            ws.write(D_ROW_HDR, c, t, th)
        itf = self.f(bold=True, border=1, border_color=LINE2, indent=1)
        valf = self.fi(bold=True, font_size=11, align='center', num_format='#,##0')
        valf1 = self.fi(bold=True, font_size=11, align='center', num_format='#,##0.0')
        info = self.f(border=1, border_color=LINE2, font_color=MUTED, font_size=9, indent=1)
        infoc = self.f(border=1, border_color=LINE2, font_color=MUTED, font_size=9, align='center')
        for i, it in enumerate(ITENS):
            r = D_ROW_IT0 + i
            ws.set_row(r, 21)
            cfg = lambda cc, i=i: "%s!%s" % (quote(self.ws_cfg), rc(C_ROW_IT0 + i, C_COL[cc], True, True))
            ws.write_formula(r, D_COL['ITEM'], '=' + cfg('ROT'), itf, it[1])
            vf = valf1 if it[0] in ('lamaC3', 'lamaC2') else valf
            if it[0] in self.dia0:
                ws.write_number(r, D_COL['VAL'], self.dia0[it[0]], vf)
            else:
                ws.write_blank(r, D_COL['VAL'], None, vf)
            ws.write_blank(r, D_COL['NRES'], None, infoc)
            ws.write_formula(r, D_COL['JAN'],
                             '=IF(%s="","lançar à mão",IF(%s="E","D-1 23:00 a D 21:00","D 00:00 a D 23:59")&"  ·  "&'
                             'LOWER(%s))' % (cfg('TAG'), cfg('JAN'), cfg('AGR')), info, '')
            self.nome('d_' + it[0], ws, r, D_COL['VAL'])
        # numero de resultados zero em vermelho (item sem valor no MES)
        ws.conditional_format(D_ROW_IT0, D_COL['NRES'], D_ROW_IT0 + NI - 1, D_COL['NRES'],
                              {'type': 'cell', 'criteria': '==', 'value': 0,
                               'format': self.wb.add_format({'font_color': NOK, 'bold': True})})
        r = D_ROW_FATOR
        ws.set_row(r, 21)
        ws.write(r, D_COL['ITEM'], 'Fator', itf)
        ws.write(r, D_COL['VAL'], self.dia0.get('fator', FATOR_PADRAO), valf)
        ws.merge_range(r, D_COL['NRES'], r, D_COL['JAN'], 'lançado à mão (dia novo: repete o do dia anterior)', info)
        self.nome('d_fator', ws, r, D_COL['VAL'])
        r = D_ROW_OBS
        ws.set_row(r, 21)
        ws.write(r, D_COL['ITEM'], 'Observação do dia (vai para o painel)', itf)
        ws.merge_range(r, D_COL['VAL'], r, D_COL['JAN'], '', self.fi(indent=1))
        self.nome('d_obs', ws, r, D_COL['VAL'])

        # legenda das janelas, com as datas do dia escolhido
        leg = self.f(font_color=INK2, font_size=9, bg_color=BG, indent=1)
        legb = self.f(font_color=INK2, font_size=9, bg_color=BG, indent=1, bold=True)
        r = D_ROW_LEG
        ws.merge_range(r, 1, r, ultc, '', legb)
        ws.write_formula(r, 1, '="Legenda:   D = dia dos dados ("&%s&")   ·   D-1 = dia anterior ("&%s&")   ·   '
                               'report enviado em D+1 ("&%s&")"' % (dmy('dData'), dm('(dData-1)'), dm('(dData+1)')),
                         legb, '')
        ws.merge_range(r + 1, 1, r + 1, ultc, '', leg)
        ws.write_formula(r + 1, 1, '="D-1 23:00 a D 21:00 = "&%s&" 23:00 a "&%s&" 21:00 (produção e espessadores)   ·   '
                                   'D 00:00 a D 23:59 = "&%s&" (batch, tanques e DVR)"'
                         % (dm('(dData-1)'), dm('dData'), dm('dData')), leg, '')

        # (3) resultado do dia (mesmas formulas do historico)
        r = D_ROW_RES
        self.secao(ws, r, 1, 2, '③  RESULTADO DO DIA')
        refs = {k: 'd_' + k for k in IK}
        refs.update({'espP': '_espP', 'eb4P': '_eb4P', 'lama': '_lama'})
        res = [
            ('S.E. ESP C2 + C3 (ponderada pela produção)', 'espP', '#,##0'),
            ('S.E. batch EB4 + ESP C3 (ponderada)', 'eb4P', '#,##0'),
            ('Diferença EB6 − (ESP C2 + C3)', 'difK', fmt_sinal()),
            ('Diferença EB6 − (EB4 + ESP C3)', 'difJ', fmt_sinal()),
            ('Lama incorporada total (tms)', 'lama', '#,##0'),
            ('% lama', 'pctLama', '0.00%'),
        ]
        form = {k: f for k, _t, _w, _tp, f in H_COLS if f}
        for j, (t, k, nf) in enumerate(res):
            rr = r + 1 + j
            ws.write(rr, D_COL['ITEM'], t, self.f(border=1, border_color=LINE2, indent=1))
            ws.write_formula(rr, D_COL['VAL'], form[k].format(**refs),
                             self.f(bold=True, border=1, border_color=LINE2, bg_color=BG, align='center',
                                    num_format=nf), '')
            self.nome('_' + k, ws, rr, D_COL['VAL'])

        # (4) comentarios do painel (editaveis) e tolerancia
        r = D_ROW_COM
        self.secao(ws, r, 1, ultc, '④  COMENTÁRIOS DO PAINEL E TOLERÂNCIA  (texto automático: dá para editar · '
                                   'Sim/Não = entra no painel)')
        self.botao(ws, r, D_COL['JAN'], 'Restaurar automáticos', 'RestaurarComentarios', largura=150, altura=22,
                   estilo='s', x=176, y=0)
        lab = self.f(bold=True, border=1, border_color=LINE2, indent=1, font_size=9.5)
        tog = self.fi(align='center', bold=True)
        txt = self.fi(text_wrap=True, valign='top', font_size=9.5, indent=1)
        nomes_auto = ['Automático: último dia', 'Automático: período', 'Automático: % lama']
        for j in range(3):
            rr = r + 1 + j
            ws.set_row(rr, 32)
            ws.write(rr, D_COL['ITEM'], nomes_auto[j], lab)
            ws.write(rr, D_COL['VAL'], 'Sim', tog)
            ws.data_validation(rr, D_COL['VAL'], rr, D_COL['VAL'], {'validate': 'list', 'source': ['Sim', 'Não']})
            ws.merge_range(rr, D_COL['NRES'], rr, D_COL['JAN'], '', txt)
            ws.write_formula(rr, D_COL['NRES'], '=gCom%d' % (j + 1), txt, '')
            self.nome('dComOn%d' % (j + 1), ws, rr, D_COL['VAL'])
            self.nome('dCom%d' % (j + 1), ws, rr, D_COL['NRES'])
        for j in range(3):
            rr = r + 4 + j
            ws.set_row(rr, 24)
            ws.write(rr, D_COL['ITEM'], 'Tópico %d (livre)' % (j + 1), lab)
            ws.merge_range(rr, D_COL['VAL'], rr, D_COL['JAN'], '', txt)
            self.nome('dTop%d' % (j + 1), ws, rr, D_COL['VAL'])
        rr = r + 8
        ws.set_row(rr, 24)
        ws.write(rr, D_COL['ITEM'], 'Tolerância da diferença de S.E. (±)', lab)
        ws.write(rr, D_COL['VAL'], TOL_PADRAO, self.fi(align='center', bold=True, font_size=11, num_format='#,##0'))
        self.nome('cfgTol', ws, rr, D_COL['VAL'])
        ws.merge_range(rr, D_COL['NRES'], rr, D_COL['JAN'], 'linha verde tracejada e OK/NOK do painel',
                       self.f(border=1, border_color=LINE2, font_color=MUTED, font_size=9, indent=1))

        r = D_ROW_ST
        self.rotulo(ws, r, 1, 1, 'Última busca')
        ws.merge_range(r + 1, 1, r + 3, ultc, '',
                       self.f(font_color=INK2, font_size=9, text_wrap=True, valign='top', border=1,
                              border_color=LINE, bg_color=BG))
        self.nome('dStatus', ws, r + 1, 1)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------------------ _Graf
    def aba_graf(self):
        ws = self.ws_gra
        tit = ['Data', 'Rótulo'] + [NOMES_SERIE.get(k, k) for k in G_COLS[2:]]
        for c, t in enumerate(tit):
            ws.write(0, c, t)
        for k in range(1, NMAXD + 2):
            r = k
            a = rc(r, 0)
            ws.write_formula(r, 0, '=IF(%d>rDias,"",rFim-rDias+%d)' % (k, k), self.f(num_format='dd/mm/yyyy'))
            ws.write_formula(r, 1, '=IF(%d<=rDias,%s,IF(%d=rDias+1,"Média",""))' % (k, dm(a), k))
            for key in G_DIA:
                col = 'h_' + key
                dia = ('IFERROR(IF(INDEX(%s,MATCH(%s,h_data,0))="",NA(),INDEX(%s,MATCH(%s,h_data,0))),NA())'
                       % (col, a, col, a))
                media = 'IF(ISNUMBER(gMed_%s),gMed_%s,NA())' % (key, key) if key in G_MEDIA else 'NA()'
                ws.write_formula(r, GC[key], '=IF(%d<=rDias,%s,IF(%d=rDias+1,%s,NA()))' % (k, dia, k, media))
            for key, base in [('difKmed', 'difK'), ('difJmed', 'difJ'), ('pctMed', 'pctLama')]:
                ws.write_formula(r, GC[key], '=IF(AND(%d=rDias+1,ISNUMBER(gMed_%s)),gMed_%s,NA())' % (k, base, base))
            ws.write_formula(r, GC['tolP'], '=IF(%d<=rDias+1,cfgTol,NA())' % k)
            ws.write_formula(r, GC['tolN'], '=IF(%d<=rDias+1,-cfgTol,NA())' % k)
            ws.write_formula(r, GC['pct0'], '=IF(%d<=rDias,IFERROR(%s/1,-1),-1)' % (k, rc(r, GC['pctLama'])))
        # estatisticas do periodo (direto do historico): media, ultimo dia (rFim) e dias com valor
        r0 = G_ROW_ST
        ws.write(r0 - 1, 0, 'Estatísticas do período')
        for j, key in enumerate(ESTAT):
            r = r0 + j
            hk = 'h_' + key
            ws.write(r, 0, key)
            ws.write_formula(r, 1, '=IFERROR(INDEX(%s,MATCH(rFim,h_data,0))&"","")' % hk if key == 'obs' else
                             '=IFERROR(INDEX(%s,MATCH(rFim,h_data,0))/1,"")' % hk)
            self.nome('gUlt_' + key, ws, r, 1)
            if key == 'obs':
                continue
            ws.write_formula(r, 3, '=SUMPRODUCT((h_data>=rFim-rDias+1)*(h_data<=rFim)*ISNUMBER(%s))' % hk)
            ws.write_formula(r, 2, '=IFERROR(SUMIFS(%s,h_data,">="&(rFim-rDias+1),h_data,"<="&rFim)/%s,"")'
                             % (hk, rc(r, 3)))
            self.nome('gMed_' + key, ws, r, 2)
            self.nome('gN_' + key, ws, r, 3)
            if key in ('difK', 'difJ'):
                ws.write_formula(r, 4, '=COUNTIFS(h_data,">="&(rFim-rDias+1),h_data,"<="&rFim,%s,">"&cfgTol)+'
                                       'COUNTIFS(h_data,">="&(rFim-rDias+1),h_data,"<="&rFim,%s,"<"&-cfgTol)'
                                       % (hk, hk))
                self.nome('gFora_' + key, ws, r, 4)
        ws.write(r0 - 1, 1, 'Último dia')
        ws.write(r0 - 1, 2, 'Média')
        ws.write(r0 - 1, 3, 'Dias com valor')
        ws.write(r0 - 1, 4, 'Dias fora da tol.')
        # % lama maxima do periodo e o dia
        r = r0 + len(ESTAT) + 1
        p0 = '%s:%s' % (rc(1, GC['pct0'], True, True), rc(NMAXD, GC['pct0'], True, True))
        ws.write(r, 0, 'pctMax')
        ws.write_formula(r, 1, '=IF(MAX(%s)<0,"",MAX(%s))' % (p0, p0))
        ws.write_formula(r, 2, '=IFERROR(INDEX(%s:%s,MATCH(MAX(%s),%s,0)),"")'
                         % (rc(1, 0, True, True), rc(NMAXD, 0, True, True), p0, p0))
        self.nome('gPctMax', ws, r, 1)
        self.nome('gPctMaxDia', ws, r, 2)
        # status do ultimo dia
        r += 1
        for j, key in enumerate(('difK', 'difJ')):
            ws.write(r + j, 0, 'status ' + key)
            ws.write_formula(r + j, 1, '=IF(ISNUMBER(gUlt_%s),IF(ABS(gUlt_%s)<=cfgTol,"OK","NOK"),"—")' % (key, key))
            self.nome('pSt_' + key, ws, r + j, 1)
        # textos do painel
        r += 3
        ini = '(rFim-rDias+1)'
        ws.write(r, 0, 'Comentários')
        auto = [
            ('="Último dia ("&%s&"): batch EB6 "&%s&"; diferença para ESP C2 + C3 "&%s&" e para EB4 + ESP C3 "&%s&"."'
             % (dm('rFim'), tx0('gUlt_eb6'), txs('gUlt_difK'), txs('gUlt_difJ'))),
            ('="Período de "&rDias&" dias: diferença média "&%s&" e "&%s&"; fora da tolerância (±"&cfgTol&") em "'
             '&gFora_difK&" e "&gFora_difJ&" dia(s)."' % (txs('gMed_difK'), txs('gMed_difJ'))),
            ('="%% lama: "&%s&" no último dia; média "&%s&IF(ISNUMBER(gPctMax),", máxima "&%s&" em "&%s,"")&"."'
             % (txp('gUlt_pctLama'), txp('gMed_pctLama'), txp('gPctMax'), dm('gPctMaxDia'))),
        ]
        for j, f in enumerate(auto):
            ws.write_formula(r, 2 + j, f)
            self.nome('gCom%d' % (j + 1), ws, r, 2 + j)
        # comentarios do painel = automaticos marcados "Sim" (texto da aba Dia, editavel) + topicos + observacao
        partes = ['IF(AND(dComOn%d="Sim",TRIM(dCom%d&"")<>""),"•  "&dCom%d&CHAR(10),"")' % (j, j, j) for j in (1, 2, 3)]
        partes += ['IF(TRIM(dTop%d&"")<>"","•  "&dTop%d&CHAR(10),"")' % (j, j) for j in (1, 2, 3)]
        partes += ['IF(TRIM(gUlt_obs)<>"","•  "&gUlt_obs,"")']
        com = '=' + '&'.join(partes)
        ws.write_formula(r, 1, com)
        self.nome('gComentarios', ws, r, 1)
        r += 1
        ws.write(r, 0, 'Resumo')
        resumo = ('="Monitoramento S.E. — Espessador × Batch · "&%s&CHAR(10)&"Batch EB6 "&%s&" · ESP C2 + C3 "&%s&'
                  '" (dif. "&%s&", "&pSt_difK&") · EB4 + ESP C3 "&%s&" (dif. "&%s&", "&pSt_difJ&") · %% lama "&%s&'
                  'CHAR(10)&"Período de "&rDias&" dias ("&%s&" a "&%s&"): diferença média "&%s&" e "&%s&'
                  '"; fora da tolerância (±"&cfgTol&") em "&gFora_difK&" e "&gFora_difJ&" dia(s)."'
                  % (dmy('rFim'), tx0('gUlt_eb6'), tx0('gUlt_espP'), txs('gUlt_difK'), tx0('gUlt_eb4P'),
                     txs('gUlt_difJ'), txp('gUlt_pctLama'), dm(ini), dm('rFim'), txs('gMed_difK'), txs('gMed_difJ')))
        ws.write_formula(r, 1, resumo)
        self.nome('gResumo', ws, r, 1)
        r += 2
        ws.write(r, 0, 'Teste')
        self.nome('gTeste', ws, r, 1)

    # ------------------------------------------------------------------------ Painel (padrao do Comparativo)
    def aba_painel(self):
        ws = self.ws_pai
        for c, px in enumerate(P_COLS):
            ws.set_column_pixels(c, c, px)
        ws.set_column_pixels(len(P_COLS), len(P_COLS), 20)
        ws.set_column(len(P_COLS) + 1, 16383, None, None, {'hidden': True})
        ws.set_default_row(15, True)   # linhas sem uso ocultas
        bgw = self.f(bg_color=BRANCO)
        # ---- barra de ferramentas (fora da imagem)
        tb = self.f(bg_color=BG)
        for r in range(0, P_R0 - 1):
            for c in range(0, len(P_COLS) + 1):
                ws.write_blank(r, c, None, tb)
        ws.set_row(0, 8)
        ws.set_row(1, 14)
        ws.set_row(2, 26)
        ws.set_row(3, 10)
        lb = self.f(bold=True, font_size=7.5, font_color=MUTED, bg_color=BG)
        ws.write(1, 1, 'DATA FINAL', lb)
        ws.write(1, 3, 'DIAS', lb)
        ws.merge_range(2, 1, 2, 2, '', self.fi(align='center', bold=True, font_size=11, num_format='dd/mm/yyyy'))
        ws.write_datetime(2, 1, self.ultimo_dia, self.fi(align='center', bold=True, font_size=11,
                                                         num_format='dd/mm/yyyy'))
        self.nome('rFim', ws, 2, 1)
        ws.write(2, 3, DIAS_REPORT, self.fi(align='center', bold=True, font_size=11))
        self.nome('rDias', ws, 2, 3)
        ws.data_validation(2, 3, 2, 3, {'validate': 'integer', 'criteria': 'between', 'minimum': 2,
                                        'maximum': NMAXD, 'error_message': 'De 2 a %d dias.' % NMAXD})
        self.botao(ws, 1, 4, 'Último dia', 'IrUltimoDia', 86, 30, 's', x=4, y=4)
        self.botao(ws, 1, 5, '✎ Ir para o Dia', 'IrDia', 112, 30, 's', x=6, y=4)
        self.botao(ws, 1, 8, 'Atualizar painel', 'AtualizarPainel', 116, 30, 's', x=4, y=4)
        self.botao(ws, 1, 10, 'Copiar imagem', 'CopiarImagemPainel', 118, 30, 'p', x=2, y=4)
        self.botao(ws, 1, 10, 'Copiar resumo', 'CopiarResumo', 118, 30, 'n', x=126, y=4)

        R = P_R0
        # ---- cabecalho
        ws.set_row(R, 38)
        hdr = self.f(bg_color=NAVY, font_color=BRANCO, bold=True, font_size=14)
        ws.write_blank(R, 0, None, self.f(bg_color=NAVY))
        ws.merge_range(R, 1, R, 6, 'MONITORAMENTO S.E. — ESPESSADOR × BATCH',
                       self.f(bg_color=NAVY, font_color=BRANCO, bold=True, font_size=12, indent=1))
        ws.merge_range(R, 7, R, 11, '|    S.E. · MÉDIA PONDERADA PELA PRODUÇÃO C2/C3',
                       self.f(bg_color=NAVY, font_color=CINZA_LOGO, bold=True, font_size=10))
        ws.merge_range(R, 12, R, 13, '', self.f(bg_color=NAVY))
        ws.write_blank(R, 14, None, self.f(bg_color=NAVY))
        ws.insert_image(R, 13, LOGO, {'x_scale': 0.14, 'y_scale': 0.14, 'x_offset': 52, 'y_offset': 12,
                                      'object_position': 1, 'description': 'Samarco'})
        ws.set_row(R + 1, 22)
        sub = self.f(bg_color=NAVY2, font_color=BRANCO, bold=True, font_size=10, bottom=5, bottom_color=AMARELO)
        ws.write_blank(R + 1, 0, None, self.f(bg_color=NAVY2, bottom=5, bottom_color=AMARELO))
        ws.merge_range(R + 1, 1, R + 1, 5, 'REPORT GERENCIAL', self.f(bg_color=NAVY2, font_color=BRANCO, bold=True,
                                                                    font_size=10, indent=1, bottom=5,
                                                                    bottom_color=AMARELO))
        ws.merge_range(R + 1, 6, R + 1, 13, '', sub)
        ws.write_formula(R + 1, 6, '="PERÍODO  "&%s&" A "&%s&"   |   "&rDias&" DIAS"'
                         % (dmy('(rFim-rDias+1)'), dmy('rFim')),
                         self.f(bg_color=NAVY2, font_color=BRANCO, bold=True, font_size=9.5, align='right',
                                indent=1, bottom=5, bottom_color=AMARELO), '')
        ws.write_blank(R + 1, 14, None, self.f(bg_color=NAVY2, bottom=5, bottom_color=AMARELO))
        ws.set_row(R + 2, 10)
        for c in range(0, 15):
            ws.write_blank(R + 2, c, None, bgw)

        # ---- quadros do topo
        r = R + 3
        ws.set_row(r, 20)
        secf = self.f(bold=True, font_size=8, font_color=INK, bg_color=SEC, top=1, bottom=1, left=1, right=1,
                      border_color=LINE, indent=1)
        ws.merge_range(r, 1, r, 6, '', secf)
        ws.write_formula(r, 1, '="ÚLTIMO DIA · "&%s&"   |   PERÍODO DE "&rDias&" DIAS"' % dmy('rFim'), secf, '')
        ws.merge_range(r, 8, r, 10, 'COMENTÁRIOS', secf)
        ws.merge_range(r, 12, r, 13, 'PRODUÇÃO E DVR', secf)
        # farol: cabecalho
        rh = r + 1
        ws.set_row(rh, 28)
        fth = self.f(bold=True, font_size=7.5, font_color=INK2, bg_color=TH, align='center', border=1,
                     border_color=LINE, text_wrap=True)
        ws.merge_range(rh, 1, rh, 2, '', fth)
        ws.write_formula(rh, 3, '="ÚLTIMO DIA"&CHAR(10)&"("&%s&")"' % dm('rFim'), fth, '')
        ws.write(rh, 4, 'MÉDIA DO\nPERÍODO', fth)
        ws.write(rh, 5, 'TOLER.\n(±)', fth)
        ws.write(rh, 6, 'STATUS', fth)
        ftd = dict(font_size=10, align='center', border=1, border_color='#D5D9E2')
        linhas = [('S.E. batch EB6', 'eb6', '#,##0', False), ('ESP C2 + C3 (pond.)', 'espP', '#,##0', False),
                  ('EB4 + ESP C3 (pond.)', 'eb4P', '#,##0', False), ('EB6 − (ESP C2 + C3)', 'difK', fmt_sinal(), True),
                  ('EB6 − (EB4 + ESP C3)', 'difJ', fmt_sinal(), True), ('% lama no batch', 'pctLama', '0.00%', False)]
        for j, (t, k, nf, dif) in enumerate(linhas):
            rr = rh + 1 + j
            ws.set_row(rr, 21)
            ws.merge_range(rr, 1, rr, 2, t, self.f(bold=True, font_size=10, border=1, border_color='#D5D9E2',
                                                    indent=1))
            ws.write_formula(rr, 3, '=IF(ISNUMBER(gUlt_%s),gUlt_%s,"—")' % (k, k),
                             self.f(bold=dif, num_format=nf, **ftd), '')
            ws.write_formula(rr, 4, '=IF(ISNUMBER(gMed_%s),gMed_%s,"—")' % (k, k), self.f(num_format=nf, **ftd), '')
            if dif:
                ws.write_formula(rr, 5, '=cfgTol', self.f(num_format='#,##0', **ftd), TOL_PADRAO)
                ws.write_formula(rr, 6, '=pSt_%s' % k, self.f(bold=True, **ftd), '')
            else:
                ws.write(rr, 5, '—', self.f(font_color=MUTED, **ftd))
                ws.write(rr, 6, '—', self.f(font_color=MUTED, **ftd))
        f1, f2 = rh + 1, rh + len(linhas)
        ws.conditional_format(f1, 6, f2, 6, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                             'format': self.wb.add_format({'font_color': OK, 'bold': True})})
        ws.conditional_format(f1, 6, f2, 6, {'type': 'cell', 'criteria': '==', 'value': '"NOK"',
                                             'format': self.wb.add_format({'font_color': NOK, 'bold': True,
                                                                           'bg_color': NOK_ROW})})
        ws.conditional_format(f1, 1, f2, 5, {'type': 'formula', 'criteria': '=$G%d="NOK"' % (f1 + 1),
                                             'format': self.wb.add_format({'bg_color': NOK_ROW})})
        rl = f2 + 1
        ws.set_row(rl, 16)
        ws.merge_range(rl, 1, rl, 3, 'OK — DIFERENÇA DENTRO DA TOLERÂNCIA',
                       self.f(font_size=7, font_color=MUTED, bottom=1, left=1, border_color=LINE, indent=1))
        ws.merge_range(rl, 4, rl, 6, 'NOK — FORA DA TOLERÂNCIA',
                       self.f(font_size=7, font_color=MUTED, bottom=1, right=1, border_color=LINE, align='right',
                              indent=1))
        # comentarios
        fcm = self.f(font_size=9.5, font_color=INK2, text_wrap=True, valign='top', border=1, border_color=LINE,
                     indent=1)
        ws.merge_range(rh, 8, rl, 10, '', fcm)
        ws.write_formula(rh, 8, '=gComentarios', fcm, '')
        # producao e DVR do ultimo dia
        fcap = self.f(bold=True, font_size=7.5, font_color=INK2, bg_color=BG, align='center', left=1, right=1,
                      border_color=LINE)
        ws.merge_range(rh, 12, rh, 13, '', fcap)
        ws.write_formula(rh, 12, '="ÚLTIMO DIA · "&TEXT(DAY(rFim),"00")&"/"&TEXT(MONTH(rFim),"00")&"/"&'
                                 'RIGHT(YEAR(rFim),2)', fcap, '')
        caixa = [('Produção C3 (t)', 'prodC3'), ('Produção C2 (t)', 'prodC2'), ('Conc. C3 (tms)', 'concC3'),
                 ('Conc. C2 (tms)', 'concC2'), ('Lama incorp. (tms)', 'lama'), ('Fator', 'fator')]
        for j, (t, k) in enumerate(caixa):
            rr = rh + 1 + j
            ws.write(rr, 12, t, self.f(bold=True, font_size=9.5, border=1, border_color='#D5D9E2', indent=1))
            ws.write_formula(rr, 13, '=IF(ISNUMBER(gUlt_%s),gUlt_%s,"—")' % (k, k),
                             self.f(font_size=10, align='center', border=1, border_color='#D5D9E2',
                                    num_format='#,##0'), '')
        ws.write_blank(rl, 12, None, self.f(left=1, bottom=1, border_color=LINE))
        ws.write_blank(rl, 13, None, self.f(right=1, bottom=1, border_color=LINE))
        ws.set_row(rl + 1, 10)

        # ---- graficos (largura total, um por linha)
        r = rl + 2
        H_GRAF = 13          # linhas de 15 pt
        largura = sum(P_COLS[1:14])
        badges = {'g1': '=IF(pSt_difK="—","—  ","●  ÚLTIMO DIA "&pSt_difK&"  ")',
                  'g2': '=IF(pSt_difJ="—","—  ","●  ÚLTIMO DIA "&pSt_difJ&"  ")',
                  'g3': '="●  ÚLTIMO DIA "&%s&"  "' % txp('gUlt_pctLama')}
        for gk, titulo, barras, linhas_g, dif in GRAFICOS:
            ws.set_row(r, 19)
            fh = self.f(bold=True, font_size=8, font_color=INK, bg_color='#F8F9FB', top=1, left=1, bottom=1,
                        border_color=LINE, indent=1)
            ws.merge_range(r, 1, r, 12, titulo.upper(), fh)
            ws.write_formula(r, 13, badges[gk], self.f(bold=True, font_size=8, font_color='#888888',
                                                       bg_color='#F8F9FB', top=1, right=1, bottom=1,
                                                       border_color=LINE, align='right'), '')
            if dif:
                st = 'pSt_' + dif
                ws.conditional_format(r, 13, r, 13, {'type': 'formula', 'criteria': '=%s="OK"' % st,
                                                     'format': self.wb.add_format({'font_color': OK})})
                ws.conditional_format(r, 13, r, 13, {'type': 'formula', 'criteria': '=%s="NOK"' % st,
                                                     'format': self.wb.add_format({'font_color': NOK})})
            for rr in range(r + 1, r + 1 + H_GRAF):
                ws.set_row(rr, 15)
            self.grafico(ws, r + 1, 1, gk, barras, linhas_g, largura, int(H_GRAF * 20) - 2)
            r = r + 1 + H_GRAF
            ws.set_row(r, 10)
            r += 1
        # rodape
        ws.set_row(r, 16)
        ws.merge_range(r, 1, r, 13, '', self.f(font_size=7, font_color=MUTED, align='center'))
        ws.write_formula(r, 1, '="Diferença = batch EB6 − média ponderada pela produção (C2/C3)   ·   faixa verde '
                               'tracejada = tolerância (±"&cfgTol&")   ·   ◆ = diferença média do período   ·   barras '
                               'escuras = média do período"', self.f(font_size=7, font_color=MUTED, align='center'), '')
        self.wb.define_name('pArea', '=%s!%s:%s' % (quote(ws), rc(R, 0, True, True), rc(r, 14, True, True)))
        ws.print_area(R, 0, r, 14)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.3, 0.3, 0.3, 0.3)
        ws.center_horizontally()
        ws.protect('', {'format_columns': True, 'format_rows': True})

    def grafico(self, ws, r, c, gk, barras, linhas_g, largura, altura):
        """Mesmo desenho do grafico do Comparativo: barras por dia + barra escura da Média, linha laranja da
        diferenca (eixo 2) com rotulos, ◆ diferenca media e tolerancia verde tracejada."""
        g = quote(self.ws_gra)
        n = DIAS_REPORT
        ult = n + 1                     # linha da "Média"

        def rng(key):
            return '=%s!%s:%s' % (g, rc(1, GC[key], True, True), rc(ult, GC[key], True, True))
        cats = rng('rot')
        cores = [(COR_A, COR_A_MEDIA), (COR_B, COR_B_MEDIA), (COR_C, COR_C_MEDIA)]
        col = self.wb.add_chart({'type': 'column'})
        for j, key in enumerate(barras):
            base, media = cores[j]
            pontos = [None] * n + [{'fill': {'color': media}, 'border': {'none': True}}]
            op = {'name': NOMES_SERIE[key], 'categories': cats, 'values': rng(key), 'fill': {'color': base},
                  'border': {'none': True}, 'points': pontos}
            if j == 0:
                op.update({'gap': 55, 'overlap': -4})
            col.add_series(op)
        lin = self.wb.add_chart({'type': 'line'})
        pct = gk == 'g3'
        fs = '0.00%' if pct else fmt_sinal()
        lin.add_series({'name': NOMES_SERIE[linhas_g[0]], 'categories': cats, 'values': rng(linhas_g[0]),
                        'y2_axis': True, 'smooth': True,
                        'line': {'color': COR_DIF, 'width': 2.5},
                        'marker': {'type': 'circle', 'size': 6, 'fill': {'color': COR_DIF},
                                   'border': {'color': BRANCO, 'width': 1}},
                        'data_labels': {'value': True, 'position': 'above', 'num_format': fs,
                                        'font': {'name': FONTE, 'bold': True, 'size': 8, 'color': COR_DIF_TXT}}})
        lin.add_series({'name': NOMES_SERIE[linhas_g[1]], 'categories': cats, 'values': rng(linhas_g[1]),
                        'y2_axis': True, 'line': {'none': True},
                        'marker': {'type': 'diamond', 'size': 9, 'fill': {'color': NAVY},
                                   'border': {'color': BRANCO, 'width': 1}},
                        'data_labels': {'value': True, 'position': 'above', 'num_format': fs,
                                        'font': {'name': FONTE, 'bold': True, 'size': 8, 'color': NAVY}}})
        for key in linhas_g[2:]:
            lin.add_series({'name': NOMES_SERIE[key], 'categories': cats, 'values': rng(key), 'y2_axis': True,
                            'line': {'color': OK, 'width': 1, 'dash_type': 'dash', 'transparency': 35},
                            'marker': {'type': 'none'}})
        col.combine(lin)
        # escala das barras: a partir de um pouco abaixo do menor valor (o VBA reajusta ao mudar o periodo)
        vals = [d[k] for d in self.hist[-n:] for k in ('eb6', 'espC3', 'espC2', 'eb4', 'tk01', 'tk02')
                if isinstance(d.get(k), (int, float))]
        y = {'num_font': {'name': FONTE, 'size': 7.5, 'color': '#5A6A82'}, 'num_format': '#,##0',
             'major_gridlines': {'visible': True, 'line': {'color': '#EDEFF3', 'width': 0.75}},
             'line': {'none': True}, 'major_tick_mark': 'none'}
        if vals:
            y['min'] = max(0, int((min(vals) - 150) // 100 * 100))
        col.set_x_axis({'num_font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                        'line': {'color': LINE}, 'major_tick_mark': 'none', 'label_position': 'low'})
        col.set_y_axis(y)
        lin.set_y2_axis({'num_font': {'name': FONTE, 'size': 7.5, 'color': '#A06D0A'},
                         'num_format': '0.0%' if pct else '#,##0', 'line': {'none': True}, 'major_tick_mark': 'none'})
        apagar = [len(barras) + 1] + ([len(barras) + 3] if not pct else [])
        col.set_legend({'position': 'bottom', 'font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                        'delete_series': apagar})
        col.set_chartarea({'border': {'color': LINE, 'width': 0.75}, 'fill': {'color': BRANCO}})
        col.set_plotarea({'fill': {'none': True}, 'border': {'none': True},
                          'layout': {'x': 0.05, 'y': 0.08, 'width': 0.88, 'height': 0.62}})
        col.set_size({'width': largura, 'height': altura})
        col.show_blanks_as('gap')
        col.show_na_as_empty_cell()
        ws.insert_chart(r, c, col, {'x_offset': 0, 'y_offset': 0, 'object_position': 1,
                                    'description': 'Gráfico %s' % gk})

    # ------------------------------------------------------------------------ Diagnostico
    def aba_diagnostico(self):
        ws = self.ws_dia_x
        ws.set_column(0, 0, 1.5)
        for j, w in enumerate(X_LARG):
            ws.set_column(1 + j, 1 + j, w)
        ultc = len(X_COLS)
        self.faixa(ws, 1, 1, ultc, 'DIAGNÓSTICO DO MES',
                   'Use se algum item vier sem valor: testa cada tag no MES e grava em Configurações a forma que '
                   'funciona')
        self.secao(ws, 4, 1, 1, '①  DIA DO TESTE (já fechado)')
        ws.set_row(5, 28)
        ws.write_datetime(5, 1, self.ultimo_dia, self.fi(bold=True, font_size=12, align='center',
                                                         num_format='dd/mm/yyyy', border=2, border_color=AMARELO))
        self.nome('xData', ws, 5, 1)
        self.botao(ws, 4, 2, 'Rodar diagnóstico', 'DiagnosticoMES', largura=180, altura=32, estilo='a', x=10, y=10)
        self.botao(ws, 4, 4, 'Voltar ao Dia', 'IrDia', largura=130, altura=32, estilo='s', x=20, y=10)
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center', font_size=9)
        ws.set_row(X_ROW_HDR, 30)
        for j, t in enumerate(X_COLS):
            ws.write(X_ROW_HDR, 1 + j, t, th)
        cel = self.f(border=1, border_color=LINE2, font_size=9, text_wrap=True)
        for i, it in enumerate(ITENS):
            ws.set_row(X_ROW0 + i, 30)
            ws.write(X_ROW0 + i, 1, it[1], self.f(bold=True, border=1, border_color=LINE2, font_size=9, indent=1))
            for j in range(1, len(X_COLS)):
                ws.write_blank(X_ROW0 + i, 1 + j, None, cel)
        cr = X_COL['RES']
        ws.conditional_format(X_ROW0, cr, X_ROW0 + NI - 1, cr,
                              {'type': 'text', 'criteria': 'begins with', 'value': 'OK',
                               'format': self.wb.add_format({'font_color': OK, 'bold': True})})
        ws.conditional_format(X_ROW0, cr, X_ROW0 + NI - 1, cr,
                              {'type': 'text', 'criteria': 'not containing', 'value': 'OK',
                               'format': self.wb.add_format({'font_color': NOK, 'bold': True})})
        ws.write(X_ROW0 + NI + 1, 1, 'Última execução:', self.f(bold=True, font_color=MUTED, font_size=9))
        self.nome('xStatus', ws, X_ROW0 + NI + 1, 2)

    # ------------------------------------------------------------------------ Dados_MES
    def aba_dados_mes(self):
        ws = self.ws_mes
        ws.set_column(0, 0, 20)
        ws.set_column(1, 14, 14)
        ws.write(0, 0, 'Consultas ao MES (Aspen) - preenchido pelo botão Buscar no MES',
                 self.f(bold=True, font_color=AZUL_TIT))
        datas = self.f(num_format='dd/mm/yyyy hh:mm:ss')
        for k, r in [('E', 2), ('D', 5)]:
            ws.write(r, 0, 'Janela %s - início' % k)
            ws.write(r + 1, 0, 'Janela %s - fim' % k)
            ws.write_blank(r, 1, None, datas)
            ws.write_blank(r + 1, 1, None, datas)
            self.nome('mesIni%s' % k, ws, r, 1)
            self.nome('mesFim%s' % k, ws, r + 1, 1)
        ws.write(8, 0, 'Última consulta')
        self.nome('mesUltima', ws, 8, 1)
        ws.write(9, 0, 'Resposta E')
        self.nome('mesRespE', ws, 9, 1)
        ws.write(10, 0, 'Resposta D')
        self.nome('mesRespD', ws, 10, 1)
        ws.write(11, 0, 'Tags recusadas')
        self.nome('mesInvalidas', ws, 11, 1)
        ws.write(13, 0, 'Consulta janela E (A15, saída A17)', self.f(bold=True))
        ws.write(73, 0, 'Consulta janela D e diagnóstico (A75, saída A77)', self.f(bold=True))

    # ------------------------------------------------------------------------ constantes do VBA
    def modulo_layout(self):
        L = ['Option Explicit', '',
             "' Gerado por build_monitor.py - posicoes fixas das abas (nao editar a mao)",
             'Public Const NI As Long = %d' % NI,
             'Public Const NMAXD As Long = %d' % NMAXD,
             "' aba Dia (1-based)",
             'Public Const D_ROW_IT0 As Long = %d' % (D_ROW_IT0 + 1)]
        for k, c in D_COL.items():
            L.append('Public Const D_COL_%s As Long = %d' % (k, c + 1))
        L += ["' aba Diagnostico (1-based)", 'Public Const X_ROW0 As Long = %d' % (X_ROW0 + 1)]
        for k, c in X_COL.items():
            L.append('Public Const X_COL_%s As Long = %d' % (k, c + 1))
        L += ["' aba Configuracoes (1-based)", 'Public Const C_ROW_IT0 As Long = %d' % (C_ROW_IT0 + 1)]
        for k, c in C_COL.items():
            L.append('Public Const C_COL_%s As Long = %d' % (k, c + 1))
        L += ["' aba Historico (1-based)", 'Public Const H_ROW_HDR As Long = %d' % (H_ROW_HDR + 1),
              'Public Const H_NLIN As Long = %d' % H_NLIN]
        for k, c in HC.items():
            L.append('Public Const H_COL_%s As Long = %d' % (k.upper(), c + 1))
        entradas = [k for k, _t, _w, _tp, f in H_COLS if not f]
        L.append('Public Const H_ENTRADAS As String = "%s"' % '|'.join(entradas))
        L.append("' itens do MES (ordem da aba Dia e de Configuracoes)")
        L.append('Public Const I_CHAVES As String = "%s"' % '|'.join(IK))
        L.append("' colunas da aba _Graf (1-based, nesta ordem); linha 1 = titulos, dia k na linha k + 1")
        L.append('Public Const G_CHAVES As String = "%s"' % '|'.join(G_COLS))
        L.append("' series de cada grafico do Painel (barras primeiro) e quantas sao barras")
        L.append('Public Const GRAF_SERIES As String = "%s"' % '|'.join(','.join(g[2] + g[3]) for g in GRAFICOS))
        L.append('Public Const GRAF_NBARRAS As String = "%s"' % '|'.join(str(len(g[2])) for g in GRAFICOS))
        hexrgb = lambda h: str(int(h[5:7], 16) * 65536 + int(h[3:5], 16) * 256 + int(h[1:3], 16))
        L.append('Public Const COR_BASE As String = "%s"' % '|'.join(hexrgb(c) for c in (COR_A, COR_B, COR_C)))
        L.append('Public Const COR_MEDIA As String = "%s"' % '|'.join(hexrgb(c) for c in
                                                                     (COR_A_MEDIA, COR_B_MEDIA, COR_C_MEDIA)))
        L.append('Public Const SENHA_CFG As String = "%s"' % SENHA_CFG)
        L.append('Public Const SENHA As String = ""')
        L.append('')
        return '\n'.join(L) + '\n'


# ============================================================================
def checar_declaracoes(nome, codigo):
    """O Excel exige declaracoes de modulo (Const, Dim, Private x As...) antes da 1a rotina."""
    em_rotina, viu_rotina = False, False
    for n, linha in enumerate(codigo.split('\n'), start=1):
        t = linha.strip()
        if re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property) ', t):
            em_rotina = viu_rotina = True
        elif re.match(r'^End (Sub|Function|Property)\b', t):
            em_rotina = False
        elif viu_rotina and not em_rotina and re.match(
                r'^(Public |Private |Global )?(Const |Dim |Declare |Type |Enum )|^(Public|Private|Global) \w+ As ', t):
            raise SystemExit('ERRO VBA em %s, linha %d: declaração depois de rotina: %s' % (nome, n, t))
    try:
        codigo.encode('cp1252')
    except UnicodeEncodeError as e:
        raise SystemExit('ERRO VBA em %s: caractere fora do cp1252 (%s) - use ChrW()' % (nome, e))
    return codigo


def ler_vba(nome):
    with open(os.path.join(VBA_DIR, nome), encoding='utf-8') as fh:
        return checar_declaracoes(nome, fh.read())


def montar_vba(layout_code):
    with open(os.path.join(VBA_DIR, 'ModLayout.bas'), 'w', encoding='utf-8') as fh:
        fh.write(layout_code)
    mods = [{'name': 'ThisWorkbook', 'kind': 'workbook', 'code': ler_vba('ThisWorkbook.cls')}]
    eventos = {
        'shDia': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                  '    On Error Resume Next\n    If gCarregando Then Exit Sub\n'
                  '    If Not Intersect(Target, Nm("dData")) Is Nothing Then\n'
                  '        CarregarDia\n        RestaurarComentarios\n    End If\n'
                  'End Sub\n'),
        'shPainel': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                     '    On Error Resume Next\n    If gCarregando Then Exit Sub\n'
                     '    If Not Intersect(Target, Nm("rFim")) Is Nothing Then AtualizarPainel\n'
                     '    If Not Intersect(Target, Nm("rDias")) Is Nothing Then AtualizarPainel\n'
                     'End Sub\n\nPrivate Sub Worksheet_Activate()\n    On Error Resume Next\n'
                     '    AjustarGraficos\nEnd Sub\n'),
    }
    for cn in ['shPainel', 'shDia', 'shHist', 'shDiag', 'shConfig', 'shDadosMES', 'shGraf']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': eventos.get(cn, '')})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModDia', 'ModPainel']:
        mods.append({'name': m, 'kind': 'module', 'code': ler_vba(m + '.bas')})
    if TESTE:
        mods.append({'name': 'ModTesteLO', 'kind': 'module', 'code': ler_vba('../build/ModTesteLO.bas')})
    return build_vba_project(mods)


def pos_processar(caminho):
    """Liga as macros as caixas de texto (botoes), cantos arredondados e 'nao imprimir'."""
    tmp = caminho + '.tmp'
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r'xl/drawings/drawing\d+\.xml$', item.filename):
                xml = data.decode('utf-8')
                xml = re.sub(r'<xdr:twoCellAnchor(.*?)</xdr:twoCellAnchor>', _patch_anchor, xml, flags=re.S)
                data = xml.encode('utf-8')
            zout.writestr(item, data)
    os.replace(tmp, caminho)


def _patch_anchor(m):
    bloco = m.group(0)
    achou = re.search(r'descr="macro:([A-Za-z0-9_]+)"', bloco)
    if not achou:
        return bloco
    macro = achou.group(1)
    bloco = bloco.replace(achou.group(0), 'descr="Botão: %s"' % macro)
    bloco = re.sub(r'<xdr:sp macro="[^"]*"', '<xdr:sp macro="[0]!%s"' % macro, bloco, count=1)
    bloco = bloco.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
                          '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/>'
                          '</a:avLst></a:prstGeom>', 1)
    bloco = bloco.replace('<xdr:clientData/>', '<xdr:clientData fPrintsWithSheet="0"/>')
    bloco = bloco.replace('<a:p><a:r>', '<a:p><a:pPr algn="ctr"/><a:r>')
    return bloco


def main():
    c = Construtor(SAIDA + '.base.xlsm')
    c.construir()
    vba_bin = montar_vba(c.modulo_layout())
    bin_path = os.path.join(AQUI, 'vbaProject.bin')
    with open(bin_path, 'wb') as fh:
        fh.write(vba_bin)
    c.wb.add_vba_project(bin_path)
    c.wb.close()
    os.replace(SAIDA + '.base.xlsm', SAIDA)
    pos_processar(SAIDA)
    os.remove(bin_path)
    print('Gerado:', SAIDA)


if __name__ == '__main__':
    main()
