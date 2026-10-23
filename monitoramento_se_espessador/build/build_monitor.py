"""Gera Monitoramento_SE_Espessador_Batch.xlsm (Excel com macros).

Planilha automática do report "S.E. - Espessador x Batch" (antes: Monitoramento SE - Após fator ESP 6.2.xlsx):
o botão Buscar no MES traz os resultados do dia (Aspen IP.21, GetCalculationValues), grava no histórico
e os gráficos do Report se atualizam pelos últimos N dias.

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
from xlsxwriter.utility import xl_rowcol_to_cell as rc, xl_col_to_name as colname

from vba_project import build_vba_project

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
VBA_DIR = os.path.join(RAIZ, 'vba')
TESTE = '--teste' in sys.argv
SAIDA = os.path.join(RAIZ, 'Monitoramento_SE_Espessador_Batch%s.xlsm' % ('_TESTE' if TESTE else ''))
LOGO = os.path.join(AQUI, 'assets', 'logo_samarco.png')   # kit visual Samarco (fundo escuro)
HISTORICO = os.path.join(AQUI, 'historico.json')

# Padrao visual Samarco (o mesmo do Comparativo Germano x Ubu)
NAVY = '#00335A'
NAVY2 = '#004070'
AZUL_TIT = '#003B6C'
AZUL_PROF = '#002643'
AMARELO = '#FFC000'
AZUL_CLARO = '#59C6F2'
LARANJA = '#F37021'
CINZA_LOGO = '#B5BEC4'
AZUL_CINZA = '#61889B'
INK = '#1D2B36'
INK2 = '#2B3A47'
MUTED = '#5B6B77'
LINE = '#C9D1D7'
BG = '#F3F5F7'
TH = '#E6EBEF'
OK = '#0A7C50'
NOK = '#C62828'
INPUT = '#FFF8E1'
BRANCO = '#FFFFFF'
FONTE = 'Segoe UI'
SENHA_CFG = 'Samarco@Lab'

# ----------------------------------------------------------------------------
# Itens buscados no MES (ordem da aba Dia e da aba Configuracoes)
#   chave, rotulo, tag, janela (E = espessador D-1 23:00 a D 21:00 / D = dia 00:00 a 23:59), agregacao,
#   faixa valida, decimais, de onde vem no MES (tela / receita / item)
# ----------------------------------------------------------------------------
ITENS = [
    ('prodC3', 'Produção C3 (t)', 'M4239010000-1445-DDDVRCC', 'D', 'Último', 0, 200000, 0,
     'DVR - Concen C3 transf Mineroduto 3'),
    ('prodC2', 'Produção C2 (t)', 'M0239010000-0596-HHCC', 'E', 'Soma', 0, 200000, 0,
     '15-G2-HH INDICES PRODUCAO - Ritmo Produção Concentrador II'),
    ('espC3', 'S.E. alimentação espessador C3', 'M4238010000-2888-HHLQ', 'E', 'Média', 300, 8000, 0,
     '40-G3-HH ALIMENTACAO ESPESSADO - SE'),
    ('espC2', 'S.E. alimentação espessador C2', 'M0238010000-0017-HHLQ', 'E', 'Média', 300, 8000, 0,
     '40-G2-HH ALIMENTACAO ESPESSADO - SE'),
    ('eb4', 'S.E. batch EB4', 'M250010000-17-ABLQ', 'D', 'Média', 300, 8000, 0,
     '50-G2-HH AMOSTRAGEM BATCH - SE'),
    ('eb6', 'S.E. batch EB6', 'M4250010000-0017-ABLQ', 'D', 'Média', 300, 8000, 0,
     '50-G3-HH AMOSTRAGEM BATCH - SE'),
    ('tk01', 'S.E. TK 01 - EB6', '', 'D', 'Média', 300, 8000, 0, 'sem tag: lançar à mão'),
    ('tk02', 'S.E. TK 02 - EB6', '', 'D', 'Média', 300, 8000, 0, 'sem tag: lançar à mão'),
    ('lamaC3', 'Lama incorporada C3 (tms)', 'M4239000000-0781-DDDVR', 'D', 'Último', 0, 20000, 0,
     'BOLETIM DVR - Lama incorporada no concentrado'),
    ('lamaC2', 'Lama incorporada C2 (tms)', 'M0239000000-0781-DDDVR', 'D', 'Último', 0, 20000, 0,
     '00-G2-DD BALANCO DVR - Lama incorporada (soma com a do C3)'),
    ('concC3', 'Concentrado C3 (tms)', 'M4224010000-0317-DDDVRCC', 'D', 'Último', 0, 200000, 0,
     'BOLETIM DVR - Beneficiamento - Concentrado C3 (no MES: Benef ROM Aliment C3)'),
    ('concC2', 'Concentrado C2 (tms)', 'M0238010000-1445-DDDVR', 'D', 'Último', 0, 200000, 0,
     'BOLETIM DVR - Esp. Conc DVR (tms)'),
]
NI = len(ITENS)
IK = [i[0] for i in ITENS]
SERVIDOR_PADRAO = 'UBU'
MAPA_PADRAO = 'IP_MESVALOR'
# janelas de 1 h: 1o horario (horas a partir de D 00:00) e quantos horarios
#   E: D-1 23:00 ... D 21:00 (23 horarios; caderno e Word) - defasagem de 2 h em relacao ao batch
#   D: D 01:00 ... D 23:00 e o fechamento 23:59 (24 janelas de 1 h a partir de 00:30)
JANELAS = {'E': (-1, 23, 'Espessador: D-1 23:00 a D 21:00'), 'D': (1, 24, 'Dia: D 00:00 a D 23:59')}
FATOR_PADRAO = 160
DIAS_REPORT = 15
NMAXD = 60          # maximo de dias no Report

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
    ('origem', 'Origem', 10, 'txt', None),
    ('atual', 'Atualizado em', 15, 'datahora', None),
    ('obs', 'Observação', 30, 'txt', None),
]
H_COL1 = 1                       # coluna B
HC = {k: H_COL1 + i for i, (k, *_r) in enumerate(H_COLS)}

# Aba Dia (0-based)
D_ROW_DATA = 4
D_ROW_HDR = 7
D_ROW_IT0 = 8                    # 1o item
D_ROW_FATOR = D_ROW_IT0 + NI
D_ROW_OBS = D_ROW_FATOR + 1
D_ROW_RES = D_ROW_OBS + 3        # resultado do dia (calculado)
D_ROW_PER = D_ROW_RES + 9        # buscar periodo
D_COL = dict(ITEM=1, VAL=2, NRES=3, JAN=4, TAG=5, SIT=6, ORIG=7)

# Aba Configuracoes (0-based)
C_ROW_IT0 = 7
C_COL = dict(CHAVE=1, ROT=2, TAG=3, SERV=4, MAPA=5, JAN=6, AGR=7, VMIN=8, VMAX=9, DEC=10, ORIG=11)

# Aba _Graf (0-based): linha 0 = titulos, dias a partir da linha 1
G_COLS = ['data', 'rot', 'espP', 'eb6', 'difK', 'eb4P', 'difJ', 'tk01', 'tk02', 'pctLama']
GC = {k: i for i, k in enumerate(G_COLS)}


def dm(x):
    """dd/mm em formula, sem codigo de formato do TEXT (que muda com o idioma do Excel: yyyy x aaaa)."""
    return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")' % (x, x)


def dmy(x):
    return '%s&"/"&YEAR(%s)' % (dm(x), x)


def quote(ws):
    return "'%s'" % ws.name


class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_size(1500, 950)
        self.fmts = {}
        self.wb.set_properties({'title': 'Monitoramento S.E. - Espessador x Batch', 'company': 'Samarco'})
        self.hoje = dtm.date.today()
        with open(HISTORICO, encoding='utf-8') as fh:
            self.hist = json.load(fh)

    def f(self, **kw):
        base = {'font_name': FONTE, 'font_size': 10, 'font_color': INK, 'valign': 'vcenter'}
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self.fmts:
            self.fmts[chave] = self.wb.add_format(base)
        return self.fmts[chave]

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
        ws.set_row(r, 34)
        ws.merge_range(r, c1, r, c2, titulo, self.f(bold=True, font_size=14, font_color=BRANCO, bg_color=NAVY,
                                                  indent=1, bottom=5, bottom_color=AMARELO))
        if sub is not None:
            ws.set_row(r + 1, 22)
            ws.merge_range(r + 1, c1, r + 1, c2, sub, self.f(font_size=9.5, bold=True, font_color='#D5DEEB',
                                                           bg_color=NAVY2, indent=1))

    def num_fmt(self, tipo):
        return {'data': 'dd/mm/yyyy', 'int': '#,##0', 'sinal': '+#,##0;−#,##0;0', 'pct': '0.00%',
                'datahora': 'dd/mm/yyyy hh:mm', 'txt': '@'}[tipo]

    # ------------------------------------------------------------------------
    def construir(self):
        wb = self.wb
        self.ws_rep = wb.add_worksheet('Report')
        self.ws_dia = wb.add_worksheet('Dia')
        self.ws_his = wb.add_worksheet('Histórico')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_gra = wb.add_worksheet('_Graf')
        for ws, cn in [(self.ws_rep, 'shReport'), (self.ws_dia, 'shDia'), (self.ws_his, 'shHist'),
                       (self.ws_cfg, 'shConfig'), (self.ws_mes, 'shDadosMES'), (self.ws_gra, 'shGraf')]:
            ws.set_vba_name(cn)
            ws.hide_gridlines(2)
        self.aba_config()
        self.aba_historico()
        self.aba_dia()
        self.aba_graf()
        self.aba_report()
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
        self.faixa(ws, 1, 1, C_COL['ORIG'], 'CONFIGURAÇÕES', 'Tags do MES, janelas de horário e padrões do report')
        self.botao(ws, 3, C_COL['ORIG'], 'Voltar', 'FecharConfiguracoes', largura=110, altura=26, estilo='s', x=300)
        self.botao(ws, 3, C_COL['ORIG'], 'Ver Dados_MES', 'MostrarDadosMES', largura=130, altura=26, estilo='s',
                   x=150)
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center')
        cel = self.f(border=1, border_color=LINE, bg_color=INPUT, locked=False)
        celc = self.f(border=1, border_color=LINE, bg_color=INPUT, locked=False, align='center')
        fixo = self.f(border=1, border_color=LINE, bg_color=BG, font_color=MUTED)
        r0 = C_ROW_IT0 - 2
        ws.write(r0, 1, 'Itens buscados no MES (janela E = espessador D−1 23:00 a D 21:00 · D = dia 00:00 a 23:59 · '
                        'agregação: Soma, Média ou Último · item sem tag = lançamento à mão na aba Dia)',
                 self.f(bold=True, font_color=AZUL_TIT))
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
            ws.write(r, C_COL['MAPA'], MAPA_PADRAO, celc)
            ws.write(r, C_COL['JAN'], it[3], celc)
            ws.write(r, C_COL['AGR'], it[4], celc)
            ws.write(r, C_COL['VMIN'], it[5], celc)
            ws.write(r, C_COL['VMAX'], it[6], celc)
            ws.write(r, C_COL['DEC'], it[7], celc)
            ws.write(r, C_COL['ORIG'], it[8], cel)
        r1 = C_ROW_IT0 + NI
        ws.data_validation(C_ROW_IT0, C_COL['JAN'], r1 - 1, C_COL['JAN'], {'validate': 'list', 'source': ['E', 'D']})
        ws.data_validation(C_ROW_IT0, C_COL['AGR'], r1 - 1, C_COL['AGR'],
                           {'validate': 'list', 'source': ['Soma', 'Média', 'Último']})
        self.wb.define_name('cfgItens', '=%s!%s:%s' % (quote(ws), rc(C_ROW_IT0, C_COL['CHAVE'], True, True),
                                                        rc(r1 - 1, C_COL['ORIG'], True, True)))

        # janelas
        r = r1 + 2
        ws.write(r, 1, 'Janelas de horário (consultas de 1 h; o resultado das 01:00 cai na janela 00:30–01:30)',
                 self.f(bold=True, font_color=AZUL_TIT))
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
        ws.write(r, 1, 'Parâmetros', self.f(bold=True, font_color=AZUL_TIT))
        params = [
            ('cfgFonte', 'Fonte dos dados (MES ou SIMULAÇÃO para treinar sem MES)', 'MES'),
            ('cfgFatorPadrao', 'Fator padrão de um dia novo (se não houver dia anterior no histórico)', FATOR_PADRAO),
            ('cfgDiasPadrao', 'Dias no Report (padrão ao abrir)', DIAS_REPORT),
            ('cfgTimeout', 'Tempo máximo de espera da consulta (s)', 120),
            ('cfgCalc', 'Tipo de cálculo do GetCalculationValues (igual aos relatórios do laboratório)', '1'),
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
        ws.protect(SENHA_CFG, {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------------------ Historico
    def aba_historico(self):
        ws = self.ws_his
        ws.set_column(0, 0, 1.5)
        for i, (k, t, w, tipo, form) in enumerate(H_COLS):
            ws.set_column(H_COL1 + i, H_COL1 + i, w)
        ultc = H_COL1 + len(H_COLS) - 1
        self.faixa(ws, 1, 1, ultc, 'HISTÓRICO DIÁRIO',
                   'Um dia por linha. Colunas brancas: MES ou lançamento · cinza: calculadas (não editar)')
        ws.write(3, 1, 'Gravado pela aba Dia (Buscar no MES / Salvar). Pode corrigir à mão as colunas brancas '
                       '(TK 01, TK 02, Fator, Observação...). A base de 28/01 a 15/10/2026 veio da planilha antiga.',
                 self.f(font_color=MUTED, italic=True))
        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center')
        thc = self.f(bold=True, font_color=BRANCO, bg_color=AZUL_CINZA, border=1, border_color=LINE,
                     text_wrap=True, align='center')
        ws.set_row(H_ROW_HDR, 42)
        for i, (k, t, w, tipo, form) in enumerate(H_COLS):
            ws.write(H_ROW_HDR, H_COL1 + i, t, thc if form else th)
        val = {tipo: self.f(num_format=self.num_fmt(tipo), border=1, border_color='#E3E7EA', align='center'
                            if tipo not in ('txt',) else 'left') for tipo in ('data', 'int', 'sinal', 'pct',
                                                                               'datahora', 'txt')}
        calc = {tipo: self.f(num_format=self.num_fmt(tipo), border=1, border_color='#E3E7EA', bg_color=BG,
                             align='center', font_color=INK2) for tipo in ('int', 'sinal', 'pct')}
        hist = sorted(self.hist, key=lambda d: d['data'])
        for n in range(H_NLIN):
            r = H_ROW_HDR + 1 + n
            reg = hist[n] if n < len(hist) else None
            refs = {k: rc(r, H_COL1 + i) for i, (k, *_x) in enumerate(H_COLS)}
            for i, (k, t, w, tipo, form) in enumerate(H_COLS):
                c = H_COL1 + i
                if form:
                    ws.write_formula(r, c, form.format(**refs), calc[tipo], '')
                elif reg is not None:
                    if k == 'data':
                        ws.write_datetime(r, c, dtm.datetime.strptime(reg['data'], '%Y-%m-%d'), val['data'])
                    elif k == 'origem':
                        ws.write_string(r, c, 'Importado', val['txt'])
                    elif k in reg:
                        ws.write_number(r, c, reg[k], val[tipo])
                    else:
                        ws.write_blank(r, c, None, val[tipo])
                else:
                    ws.write_blank(r, c, None, val[tipo])
        ws.freeze_panes(H_ROW_HDR + 1, H_COL1 + 1)
        ws.autofilter(H_ROW_HDR, H_COL1, H_ROW_HDR + len(hist), ultc)
        self.n_hist = len(hist)
        self.dia0 = hist[-1] if hist else {}       # a aba Dia abre com o ultimo dia do historico
        self.ultimo_dia = dtm.datetime.strptime(hist[-1]['data'], '%Y-%m-%d') if hist else None
        for k, c in HC.items():
            self.wb.define_name('h_' + k, "=%s!%s:%s" % (quote(ws), rc(H_ROW_HDR + 1, c, True, True),
                                                        rc(H_ROW_HDR + H_NLIN, c, True, True)))

    # ------------------------------------------------------------------------ Dia
    def aba_dia(self):
        ws = self.ws_dia
        ws.set_column(0, 0, 1.5)
        ws.set_column(D_COL['ITEM'], D_COL['ITEM'], 40)
        ws.set_column(D_COL['VAL'], D_COL['VAL'], 13)
        ws.set_column(D_COL['NRES'], D_COL['NRES'], 11)
        ws.set_column(D_COL['JAN'], D_COL['JAN'], 30)
        ws.set_column(D_COL['TAG'], D_COL['TAG'], 27)
        ws.set_column(D_COL['SIT'], D_COL['SIT'], 30)
        ws.set_column(D_COL['ORIG'], D_COL['ORIG'], 52)
        ultc = D_COL['ORIG']
        self.faixa(ws, 1, 1, ultc, 'MONITORAMENTO S.E. — ESPESSADOR × BATCH',
                   'Dia do report: buscar no MES, conferir, lançar o que não vem do MES e salvar no histórico')
        lab = self.f(bold=True, font_color=AZUL_TIT, align='right')
        ent = self.f(bold=True, font_size=12, bg_color=INPUT, border=2, border_color=AMARELO, align='center',
                     num_format='dd/mm/yyyy', locked=False)
        ws.set_row(D_ROW_DATA, 30)
        ws.write(D_ROW_DATA, D_COL['ITEM'], 'Data do report  ▸', lab)
        dia0 = self.ultimo_dia or dtm.datetime.combine(self.hoje, dtm.time())
        ws.write_datetime(D_ROW_DATA, D_COL['VAL'], dia0, ent)
        self.nome('dData', ws, D_ROW_DATA, D_COL['VAL'])
        self.botao(ws, D_ROW_DATA, D_COL['NRES'], 'Buscar no MES', 'BuscarMES', largura=150, estilo='a', y=0)
        self.botao(ws, D_ROW_DATA, D_COL['JAN'], 'Salvar no histórico', 'SalvarDia', largura=150, estilo='g',
                   x=120)
        self.botao(ws, D_ROW_DATA, D_COL['TAG'], 'Ver report', 'IrReport', largura=110, estilo='p', x=110)
        self.botao(ws, D_ROW_DATA, D_COL['SIT'], 'Configurações', 'AbrirConfiguracoes', largura=120, estilo='s',
                   x=120)
        ws.write_formula(D_ROW_DATA + 1, D_COL['ITEM'],
                         '="Espessador e produção C2: "&%s&" 23:00 a "&%s&" 21:00  ·  Batch, tanques e DVR: "&%s&'
                         '" 00:00 a 23:59"' % (dm('(dData-1)'), dm('dData'), dm('dData')),
                         self.f(font_color=MUTED, italic=True), '')
        ws.write(D_ROW_DATA + 2, D_COL['ITEM'], 'Ao trocar a data, o dia já salvo no histórico é aberto. '
                                                 'Células amarelas: podem ser corrigidas à mão.',
                 self.f(font_color=MUTED, italic=True))

        th = self.f(bold=True, font_color=BRANCO, bg_color=NAVY2, border=1, border_color=LINE, text_wrap=True,
                    align='center')
        ws.set_row(D_ROW_HDR, 30)
        for c, t in [(D_COL['ITEM'], 'Item'), (D_COL['VAL'], 'Valor'), (D_COL['NRES'], 'Nº de resultados'),
                     (D_COL['JAN'], 'Janela no MES'), (D_COL['TAG'], 'Tag'), (D_COL['SIT'], 'Situação'),
                     (D_COL['ORIG'], 'Origem no MES')]:
            ws.write(D_ROW_HDR, c, t, th)
        itf = self.f(bold=True, border=1, border_color=LINE, indent=1)
        valf = self.f(bold=True, font_size=11, border=1, border_color=AMARELO, bg_color=INPUT, align='center',
                      num_format='#,##0', locked=False)
        valf3 = self.f(bold=True, font_size=11, border=1, border_color=AMARELO, bg_color=INPUT, align='center',
                       num_format='#,##0.0', locked=False)
        info = self.f(border=1, border_color=LINE, font_color=MUTED, font_size=9)
        infoc = self.f(border=1, border_color=LINE, font_color=MUTED, font_size=9, align='center')
        cr = C_ROW_IT0
        for i, it in enumerate(ITENS):
            r = D_ROW_IT0 + i
            ws.set_row(r, 21)
            cfg = lambda cc: "%s!%s" % (quote(self.ws_cfg), rc(cr + i, C_COL[cc], True, True))
            ws.write_formula(r, D_COL['ITEM'], '=' + cfg('ROT'), itf, it[1])
            vf = valf if it[0] not in ('lamaC3', 'lamaC2') else valf3
            if it[0] in self.dia0:
                ws.write_number(r, D_COL['VAL'], self.dia0[it[0]], vf)
            else:
                ws.write_blank(r, D_COL['VAL'], None, vf)
            ws.write_blank(r, D_COL['NRES'], None, infoc)
            ws.write_formula(r, D_COL['JAN'],
                             '=IF(%s="","",IF(%s="E","D-1 23:00 a D 21:00","D 00:00 a 23:59")&" · "&%s)'
                             % (cfg('TAG'), cfg('JAN'), cfg('AGR')), info, '')
            ws.write_formula(r, D_COL['TAG'], '=IF(%s="","(sem tag)",%s)' % (cfg('TAG'), cfg('TAG')), info, '')
            ws.write_blank(r, D_COL['SIT'], None, info)
            ws.write_formula(r, D_COL['ORIG'], '=' + cfg('ORIG'), info, '')
            self.nome('d_' + it[0], ws, r, D_COL['VAL'])
        r = D_ROW_FATOR
        ws.set_row(r, 21)
        ws.write(r, D_COL['ITEM'], 'Fator', itf)
        ws.write(r, D_COL['VAL'], self.dia0.get('fator', FATOR_PADRAO), valf)
        ws.merge_range(r, D_COL['NRES'], r, D_COL['ORIG'], 'Lançado à mão (dia novo: repete o do dia anterior)',
                       info)
        self.nome('d_fator', ws, r, D_COL['VAL'])
        r = D_ROW_OBS
        ws.set_row(r, 21)
        ws.write(r, D_COL['ITEM'], 'Observação', itf)
        ws.merge_range(r, D_COL['VAL'], r, D_COL['ORIG'], '',
                       self.f(border=1, border_color=AMARELO, bg_color=INPUT, locked=False, indent=1))
        self.nome('d_obs', ws, r, D_COL['VAL'])
        self.nome('dOrigem', ws, r + 1, D_COL['VAL'])
        ws.write_string(r + 1, D_COL['VAL'], 'Importado' if self.dia0 else '', self.f(font_color=BRANCO, font_size=8))

        # resultado do dia (mesmas formulas do historico)
        r = D_ROW_RES
        ws.write(r, D_COL['ITEM'], 'Resultado do dia', self.f(bold=True, font_size=11, font_color=AZUL_TIT,
                                                               bottom=2, bottom_color=AMARELO))
        refs = {k: 'd_' + k for k in IK}
        refs.update({'espP': '_espP', 'eb4P': '_eb4P', 'lama': '_lama'})
        res = [
            ('S.E. ESP C2 + C3 (ponderada pela produção)', 'espP', '#,##0'),
            ('S.E. batch EB4 + ESP C3 (ponderada)', 'eb4P', '#,##0'),
            ('Diferença EB6 − (ESP C2 + C3)', 'difK', '+#,##0;−#,##0;0'),
            ('Diferença EB6 − (EB4 + ESP C3)', 'difJ', '+#,##0;−#,##0;0'),
            ('Lama incorporada total (tms)', 'lama', '#,##0'),
            ('% lama', 'pctLama', '0.00%'),
        ]
        form = {k: f for k, _t, _w, _tp, f in H_COLS if f}
        for j, (t, k, nf) in enumerate(res):
            rr = r + 1 + j
            ws.write(rr, D_COL['ITEM'], t, self.f(border=1, border_color=LINE, indent=1))
            ws.write_formula(rr, D_COL['VAL'], form[k].format(**refs),
                             self.f(bold=True, border=1, border_color=LINE, bg_color=BG, align='center',
                                    num_format=nf), '')
            self.nome('_' + k, ws, rr, D_COL['VAL'])
        ws.write(r + 1, D_COL['NRES'], '', None)

        # buscar periodo
        r = D_ROW_PER
        ws.write(r, D_COL['ITEM'], 'Buscar um período (preenche os dias atrasados)',
                 self.f(bold=True, font_size=11, font_color=AZUL_TIT, bottom=2, bottom_color=AMARELO))
        ent2 = self.f(bold=True, bg_color=INPUT, border=1, border_color=AMARELO, align='center',
                      num_format='dd/mm/yyyy', locked=False)
        ws.set_row(r + 1, 26)
        ws.write(r + 1, D_COL['ITEM'], 'De  ▸', lab)
        ws.write_blank(r + 1, D_COL['VAL'], None, ent2)
        ws.write(r + 2, D_COL['ITEM'], 'Até  ▸', lab)
        ws.write_blank(r + 2, D_COL['VAL'], None, ent2)
        ws.set_row(r + 2, 26)
        self.nome('dDe', ws, r + 1, D_COL['VAL'])
        self.nome('dAte', ws, r + 2, D_COL['VAL'])
        self.botao(ws, r + 1, D_COL['NRES'], 'Buscar período no MES', 'BuscarPeriodo', largura=190, altura=30,
                   estilo='a', y=10)
        ws.merge_range(r + 1, D_COL['JAN'] + 1, r + 2, D_COL['ORIG'],
                       'Cada dia é buscado e salvo no histórico. TK 01, TK 02, Fator e Observação já lançados '
                       'no histórico são mantidos.', self.f(font_color=MUTED, italic=True, text_wrap=True))
        r = r + 4
        ws.write(r, D_COL['ITEM'], 'Última busca', self.f(bold=True, font_color=MUTED))
        ws.merge_range(r, D_COL['VAL'], r + 2, D_COL['ORIG'], '',
                       self.f(font_color=INK2, font_size=9, text_wrap=True, valign='top', border=1,
                              border_color=LINE))
        self.nome('dStatus', ws, r, D_COL['VAL'])
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------------------ _Graf
    def aba_graf(self):
        ws = self.ws_gra
        his = quote(self.ws_his)
        tit = ['Data', 'Rótulo', 'S.E. ESP C2 + C3 (ponderada)', 'Batch EB6', 'Diferença EB6 − (ESP C2 + C3)',
               'Batch EB4 + ESP C3 (ponderada)', 'Diferença EB6 − (EB4 + ESP C3)', 'TK 01 - EB6', 'TK 02 - EB6',
               '% lama']
        for c, t in enumerate(tit):
            ws.write(0, c, t)
        for k in range(1, NMAXD + 1):
            r = k
            a = rc(r, 0)
            ws.write_formula(r, 0, '=IF(%d>rDias,"",rFim-rDias+%d)' % (k, k), self.f(num_format='dd/mm/yyyy'))
            ws.write_formula(r, 1, '=IF(%s="","",%s)' % (a, dm(a)))
            for key in G_COLS[2:]:
                col = 'h_' + key
                ws.write_formula(r, GC[key],
                                 '=IF(%s="",NA(),IFERROR(IF(INDEX(%s,MATCH(%s,h_data,0))="",NA(),'
                                 'INDEX(%s,MATCH(%s,h_data,0))),NA()))' % (a, col, a, col, a))
        # resumo do periodo (media e dias com valor, direto do historico) e do ultimo dia
        per = '(h_data>=rFim-rDias+1)*(h_data<=rFim)'
        for key in G_COLS[2:]:
            c = GC[key]
            hk = 'h_' + key
            ws.write_formula(NMAXD + 2, c, '=IFERROR(SUMIFS(%s,h_data,">="&(rFim-rDias+1),h_data,"<="&rFim)/%s,"")'
                             % (hk, rc(NMAXD + 4, c)))
            ws.write_formula(NMAXD + 3, c, '=IFERROR(INDEX(%s:%s,rDias)/1,"")' % (
                rc(1, c, True, True), rc(NMAXD, c, True, True)))
            ws.write_formula(NMAXD + 4, c, '=SUMPRODUCT(%s*ISNUMBER(%s))' % (per, hk))
            self.nome('gMed_' + key, ws, NMAXD + 2, c)
            self.nome('gUlt_' + key, ws, NMAXD + 3, c)
            self.nome('gN_' + key, ws, NMAXD + 4, c)
        ws.write(NMAXD + 2, 0, 'Média do período')
        ws.write(NMAXD + 3, 0, 'Último dia')
        ws.write(NMAXD + 4, 0, 'Dias com valor')
        ws.write(NMAXD + 6, 0, 'Teste')
        self.nome('gTeste', ws, NMAXD + 6, 1)

    # ------------------------------------------------------------------------ Report
    def aba_report(self):
        ws = self.ws_rep
        ws.set_column(0, 0, 1.5)
        ws.set_column(1, 15, 9.6)
        ws.set_column(16, 16, 1.5)
        ws.set_column(17, 20, 11)
        # controles (fora da area da imagem)
        lab = self.f(bold=True, font_color=AZUL_TIT, align='right')
        ent = self.f(bold=True, bg_color=INPUT, border=2, border_color=AMARELO, align='center', locked=False,
                     num_format='dd/mm/yyyy')
        ent2 = self.f(bold=True, bg_color=INPUT, border=2, border_color=AMARELO, align='center', locked=False)
        ws.set_row(1, 28)
        ws.merge_range(1, 1, 1, 2, 'Data final ▸', lab)
        fim = self.ultimo_dia or dtm.datetime.combine(self.hoje, dtm.time())
        ws.merge_range(1, 3, 1, 4, '', ent)
        ws.write_datetime(1, 3, fim, ent)
        self.nome('rFim', self.ws_rep, 1, 3)
        ws.write(1, 5, 'Dias ▸', lab)
        ws.write(1, 6, DIAS_REPORT, ent2)
        self.nome('rDias', self.ws_rep, 1, 6)
        ws.data_validation(1, 6, 1, 6, {'validate': 'integer', 'criteria': 'between', 'minimum': 2,
                                        'maximum': NMAXD})
        self.botao(ws, 1, 7, 'Atualizar', 'AtualizarReport', largura=100, altura=28, estilo='p', x=6)
        self.botao(ws, 1, 9, 'Copiar imagem', 'CopiarImagemReport', largura=125, altura=28, estilo='a', x=0)
        self.botao(ws, 1, 11, 'Último dia', 'IrUltimoDia', largura=100, altura=28, estilo='s', x=14)
        self.botao(ws, 1, 13, 'Ir para Dia', 'IrDia', largura=110, altura=28, estilo='s', x=8)

        R = 3                       # 1a linha da area da imagem
        ws.set_row(R, 40)
        ws.merge_range(R, 1, R, 15, 'Monitoramento S.E. — Espessador × Batch',
                       self.f(bold=True, font_size=16, font_color=BRANCO, bg_color=NAVY, indent=1))
        ws.insert_image(R, 13, LOGO, {'x_scale': 0.14, 'y_scale': 0.14, 'x_offset': 60, 'y_offset': 9,
                                      'object_position': 3})
        ws.set_row(R + 1, 22)
        ws.merge_range(R + 1, 1, R + 1, 15, '', self.f(font_size=9.5, bold=True, font_color='#D5DEEB',
                                                       bg_color=NAVY2, indent=1, bottom=5, bottom_color=AMARELO))
        ws.write_formula(R + 1, 1, '="Superfície específica (cm²/g) · "&%s&" a "&%s&" · "&rDias&" dias"' % (dmy('(rFim-rDias+1)'), dmy('rFim')),
                         self.f(font_size=9.5, bold=True, font_color='#D5DEEB', bg_color=NAVY2, indent=1,
                                bottom=5, bottom_color=AMARELO), '')

        # cartoes: ultimo dia e media do periodo
        r = R + 3
        cartoes = [('Batch EB6', 'eb6', '#,##0', 'FIXED(%s,0)'), ('ESP C2 + C3 (pond.)', 'espP', '#,##0', 'FIXED(%s,0)'),
                   ('EB6 − (ESP C2 + C3)', 'difK', '+#,##0;−#,##0;0', 'IF(%s>0,"+","")&FIXED(%s,0)'),
                   ('EB6 − (EB4 + ESP C3)', 'difJ', '+#,##0;−#,##0;0', 'IF(%s>0,"+","")&FIXED(%s,0)'),
                   ('% lama', 'pctLama', '0.00%', 'FIXED(%s*100,2)&"%%"')]
        ws.set_row(r, 18)
        ws.set_row(r + 1, 30)
        ws.set_row(r + 2, 18)
        for j, (t, k, nf, txt) in enumerate(cartoes):
            c1 = 1 + 3 * j
            ws.merge_range(r, c1, r, c1 + 2, t, self.f(bold=True, font_size=9, font_color=MUTED, bg_color=BG,
                                                     align='center', top=2, top_color=AMARELO if j else NAVY))
            ws.merge_range(r + 1, c1, r + 1, c1 + 2, '', self.f(bold=True, font_size=18, font_color=AZUL_TIT,
                                                               bg_color=BG, align='center', num_format=nf))
            ws.write_formula(r + 1, c1, '=gUlt_%s' % k, self.f(bold=True, font_size=18, font_color=AZUL_TIT,
                                                               bg_color=BG, align='center', num_format=nf), '')
            med = txt.replace('%s', 'gMed_' + k).replace('%%', '%')
            ws.merge_range(r + 2, c1, r + 2, c1 + 2, '', self.f(font_size=8.5, font_color=MUTED, bg_color=BG,
                                                               align='center'))
            ws.write_formula(r + 2, c1, '=IF(gN_%s=0,"sem valores no período","último dia · média do período: "&%s)' % (k, med),
                             self.f(font_size=8.5, font_color=MUTED, bg_color=BG, align='center'), '')
        ws.conditional_format(r + 1, 7, r + 1, 12, {'type': 'cell', 'criteria': '<', 'value': 0,
                                                    'format': self.wb.add_format({'font_color': NOK})})

        # graficos
        g = quote(self.ws_gra)
        n = DIAS_REPORT

        def rng(key):
            return '=%s!%s:%s' % (g, rc(1, GC[key], True, True), rc(n, GC[key], True, True))
        cats = rng('rot')
        largura, altura = 1060, 330
        self.graficos = []

        def eixos(col, lin, fmt2):
            col.set_x_axis({'num_font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                            'line': {'color': LINE}, 'major_tick_mark': 'none'})
            col.set_y_axis({'num_font': {'name': FONTE, 'size': 8, 'color': MUTED}, 'num_format': '#,##0',
                            'major_gridlines': {'visible': True, 'line': {'color': '#EDEFF3', 'width': 0.75}},
                            'line': {'none': True}, 'major_tick_mark': 'none'})
            lin.set_y2_axis({'num_font': {'name': FONTE, 'size': 8, 'color': '#B04F12'}, 'num_format': fmt2,
                             'line': {'none': True}, 'major_tick_mark': 'none'})
            col.set_legend({'position': 'bottom', 'font': {'name': FONTE, 'size': 8.5, 'bold': True, 'color': INK}})
            col.set_chartarea({'border': {'color': LINE, 'width': 0.75}, 'fill': {'color': BRANCO}})
            col.set_plotarea({'fill': {'none': True}, 'border': {'none': True}})
            col.set_size({'width': largura, 'height': altura})
            col.show_blanks_as('gap')
            col.show_na_as_empty_cell()

        def titulo(ch, t):
            ch.set_title({'name': t, 'name_font': {'name': FONTE, 'size': 11, 'bold': True, 'color': AZUL_TIT},
                          'overlay': False})

        def comparativo(t, ln_a, nome_a, cor_a, dif, nome_dif):
            # linhas = niveis (eixo principal); barras = diferenca (eixo secundario)
            lin = self.wb.add_chart({'type': 'line'})
            lin.add_series({'name': nome_a, 'categories': cats, 'values': rng(ln_a), 'y2_axis': True,
                            'line': {'color': cor_a, 'width': 2.25},
                            'marker': {'type': 'circle', 'size': 5, 'fill': {'color': cor_a},
                                       'border': {'color': BRANCO}}})
            lin.add_series({'name': 'Batch EB6', 'categories': cats, 'values': rng('eb6'), 'y2_axis': True,
                            'line': {'color': NAVY, 'width': 2.75},
                            'marker': {'type': 'square', 'size': 5, 'fill': {'color': NAVY},
                                       'border': {'color': BRANCO}}})
            col = self.wb.add_chart({'type': 'column'})
            col.add_series({'name': nome_dif, 'categories': cats, 'values': rng(dif),
                            'fill': {'color': '#F9C9A8'}, 'border': {'none': True}, 'gap': 70,
                            'invert_if_negative': False,
                            'data_labels': {'value': True, 'position': 'outside_end',
                                            'num_format': '+#,##0;−#,##0;0',
                                            'font': {'name': FONTE, 'size': 7.5, 'bold': True,
                                                     'color': '#B04F12'}}})
            col.combine(lin)
            # colunas (diferenca) no eixo da esquerda; linhas (S.E.) no da direita
            col.set_x_axis({'num_font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                            'line': {'color': LINE}, 'major_tick_mark': 'none', 'label_position': 'low'})
            col.set_y_axis({'num_font': {'name': FONTE, 'size': 8, 'color': '#B04F12'},
                            'num_format': '+#,##0;−#,##0;0', 'line': {'none': True}, 'major_tick_mark': 'none',
                            'name': 'Diferença', 'name_font': {'name': FONTE, 'size': 8, 'color': '#B04F12'}})
            lin.set_y2_axis({'num_font': {'name': FONTE, 'size': 8, 'color': MUTED}, 'num_format': '#,##0',
                             'line': {'none': True}, 'major_tick_mark': 'none',
                             'major_gridlines': {'visible': True, 'line': {'color': '#EDEFF3', 'width': 0.75}},
                             'name': 'S.E. (cm²/g)', 'name_font': {'name': FONTE, 'size': 8, 'color': MUTED}})
            col.set_legend({'position': 'bottom', 'font': {'name': FONTE, 'size': 8.5, 'bold': True, 'color': INK}})
            col.set_chartarea({'border': {'color': LINE, 'width': 0.75}, 'fill': {'color': BRANCO}})
            col.set_plotarea({'fill': {'none': True}, 'border': {'none': True}})
            col.set_size({'width': largura, 'height': altura})
            col.show_blanks_as('gap')
            col.show_na_as_empty_cell()
            titulo(col, t)
            return col

        r = R + 7
        ch1 = comparativo('S.E. — Espessador C2 + C3 × Batch EB6', 'espP', 'ESP C2 + C3 (ponderada)', AZUL_CLARO,
                          'difK', 'Diferença EB6 − (ESP C2 + C3)')
        ws.insert_chart(r, 1, ch1, {'object_position': 1, 'description': 'Gráfico 1'})
        ch2 = comparativo('S.E. — Batch EB4 + Espessador C3 × Batch EB6', 'eb4P', 'EB4 + ESP C3 (ponderada)',
                          AZUL_CINZA, 'difJ', 'Diferença EB6 − (EB4 + ESP C3)')
        ws.insert_chart(r + 17, 1, ch2, {'object_position': 1, 'description': 'Gráfico 2'})

        col = self.wb.add_chart({'type': 'column'})
        for key, nome, cor in [('eb6', 'Batch EB6', NAVY), ('tk01', 'TK 01 - EB6', AZUL_CLARO),
                               ('tk02', 'TK 02 - EB6', CINZA_LOGO)]:
            col.add_series({'name': nome, 'categories': cats, 'values': rng(key), 'fill': {'color': cor},
                            'border': {'none': True}, 'gap': 60, 'overlap': -5})
        lin = self.wb.add_chart({'type': 'line'})
        lin.add_series({'name': '% lama', 'categories': cats, 'values': rng('pctLama'), 'y2_axis': True,
                        'line': {'color': LARANJA, 'width': 2.25},
                        'marker': {'type': 'circle', 'size': 5, 'fill': {'color': LARANJA},
                                   'border': {'color': BRANCO}},
                        'data_labels': {'value': True, 'position': 'above', 'num_format': '0.00%',
                                        'font': {'name': FONTE, 'size': 7.5, 'bold': True, 'color': '#B04F12'}}})
        col.combine(lin)
        eixos(col, lin, '0.0%')
        col.set_y_axis({'num_font': {'name': FONTE, 'size': 8, 'color': MUTED}, 'num_format': '#,##0',
                        'min': 1000,
                        'major_gridlines': {'visible': True, 'line': {'color': '#EDEFF3', 'width': 0.75}},
                        'line': {'none': True}, 'major_tick_mark': 'none', 'name': 'S.E. (cm²/g)',
                        'name_font': {'name': FONTE, 'size': 8, 'color': MUTED}})
        titulo(col, '% de lama no batch × S.E. do batch EB6 e dos tanques')
        ws.insert_chart(r + 34, 1, col, {'object_position': 1, 'description': 'Gráfico 3'})
        ult = r + 34 + 17
        ws.merge_range(ult, 1, ult, 15,
                       'ESP C2 + C3 = (ESP C3 × produção C3 + ESP C2 × produção C2) / produção total  ·  '
                       'EB4 + ESP C3 = (EB4 × produção C2 + ESP C3 × produção C3) / produção total  ·  '
                       '% lama = lama incorporada (C3 + C2) / (concentrado C3 + C2)',
                       self.f(font_size=8, font_color=MUTED, italic=True, text_wrap=True))
        ws.set_row(ult, 26)
        self.wb.define_name('pArea', '=%s!%s:%s' % (quote(ws), rc(R, 1, True, True), rc(ult, 15, True, True)))
        ws.print_area(R, 1, ult, 15)
        ws.fit_to_pages(1, 1)
        ws.set_landscape()
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------------------ Dados_MES
    def aba_dados_mes(self):
        ws = self.ws_mes
        ws.set_column(0, 0, 20)
        ws.set_column(1, 14, 14)
        ws.write(0, 0, 'Consultas ao MES (Aspen) - preenchido pelo botão Buscar no MES',
                 self.f(bold=True, font_color=AZUL_TIT))
        datas = self.f(num_format='dd/mm/yyyy hh:mm:ss')
        for j, (k, r) in enumerate([('E', 2), ('D', 5)]):
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
        ws.write(73, 0, 'Consulta janela D (A75, saída A77)', self.f(bold=True))

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
        L.append("' colunas da aba _Graf (1-based, nesta ordem)")
        L.append('Public Const G_CHAVES As String = "%s"' % '|'.join(G_COLS))
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
                  '    If Not Intersect(Target, Nm("dData")) Is Nothing Then CarregarDia\n'
                  'End Sub\n'),
        'shReport': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                     '    On Error Resume Next\n    If gCarregando Then Exit Sub\n'
                     '    If Not Intersect(Target, Union(Nm("rFim"), Nm("rDias"))) Is Nothing Then AtualizarReport\n'
                     'End Sub\n'),
    }
    for cn in ['shReport', 'shDia', 'shHist', 'shConfig', 'shDadosMES', 'shGraf']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': eventos.get(cn, '')})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModDia', 'ModReport']:
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
