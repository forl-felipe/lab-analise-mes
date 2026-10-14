"""Gera a planilha Comparativo_Interlab_Germano_Ubu.xlsm (Comparativo de Medicoes de Qualidade).

Versao em Excel do painel HTML "Comparativo Interlaboratorial Germano x Ubu" (Report Gerencial):
  Painel          - o report: farol, comentarios, pellet feed ultra fino e os 5 graficos
                    (P, SiO2, PPC, <325# e S.E.); botoes de imagem, PNG, PDF e e-mail
  Batch           - cabecalho do batch, tolerancias, comentarios, PFUF e leituras por horario
                    (Germano x Ubu); botao Buscar no MES
  Enviar          - resumo pronto para o corpo do e-mail
  Histórico       - um batch por linha (tabela tbl_Batches, para o Power BI)
  BD_Leituras     - uma leitura por linha (tabela tbl_Leituras, para o Power BI)
  Configurações   - tags do MES, consulta, padroes e titulos
Ocultas: Dados_MES (consultas ao Aspen) e _Graf (series dos graficos e calculos do farol).

Uso:  python3 build_interlab.py [--teste]
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
SAIDA = os.path.join(RAIZ, 'Comparativo_Interlab_Germano_Ubu%s.xlsm' % ('_TESTE' if TESTE else ''))
LOGO = os.path.join(AQUI, 'assets', 'logo_branco.png')
HIST_USUARIO = os.path.join(AQUI, 'historico.json')

# ---------------------------------------------------------------- paleta do painel HTML
NAVY = '#14253D'
NAVY2 = '#1A3C6E'
BLUE = '#2F5C9E'
INK = '#1A2744'
INK2 = '#3A4A66'
MUTED = '#6A7A94'
LINE = '#C8CDD6'
LINE2 = '#E0E3EA'
BG = '#EEF1F5'
SEC = '#F0F2F5'
TH = '#E8EBF0'
OK = '#0A7C50'
OK_BG = '#E3F4EC'
NOK = '#C62828'
NOK_BG = '#FBE6E6'
NOK_ROW = '#FDF3F3'
COR_A = '#AECBEE'        # barras do laboratorio A (Germano)
COR_A_MEDIA = '#5F8AC8'
COR_B = '#E0E4EA'        # barras do laboratorio B (Ubu)
COR_B_MEDIA = '#A7B0BB'
COR_DIF = '#E08A00'
COR_DIF_TXT = '#9A5F00'
INPUT = '#FFF8E1'
BRANCO = '#FFFFFF'
FONTE = 'Segoe UI'

# ---------------------------------------------------------------- parametros
# (chave, rotulo, nome na frase NOK, titulo do grafico, decimais, tolerancia padrao,
#  faixa valida no MES (min, max), valor tipico A, valor tipico B (simulacao))
PARAMS = [
    ('SE', 'S.E.', 'Superfície Específica', 'Superfície Específica (S.E.)', 0, 150, (300, 5000), 2000, 1900),
    ('325', '<325#', 'Granulometria (<325#)', 'Granulometria <325#', 1, 1.0, (40, 100), 87.5, 87.0),
    ('SiO2', 'SiO₂', 'Sílica', 'Sílica (SiO₂)', 2, 0.19, (0, 20), 1.95, 1.93),
    ('P', 'P', 'Fósforo', 'Fósforo (P)', 3, 0.003, (0, 1), 0.058, 0.057),
    ('PPC', 'PPC', 'PPC', 'PPC', 2, 0.08, (0, 15), 3.60, 3.55),
]
NP = len(PARAMS)
PK = [p[0] for p in PARAMS]
# ordem dos graficos no painel (como no HTML)
ORDEM_GRAF = ['P', 'SiO2', 'PPC', '325', 'SE']
NMAX = 24            # horarios por batch (2 h cada = 48 h)
NHOR_PADRAO = 14
LABS = ('Germano', 'Ubu')
SERVIDOR_PADRAO = 'UBU'
# Tags informadas pelo laboratorio (aba Configuracoes da planilha devolvida em 14/10/2026), sem alteracao
# Correcoes (revisao de 14/10/2026):
#   S.E. Germano: M250010000-17-ABLQ (informada pelo laboratorio; as tags de Germano nao tem zeros a esquerda)
#   S.E. Ubu estava na US3 (M650060010); as demais da Alimentacao 04 estao na US4 (M4650060010)
#   P e PPC de Ubu: M650030010-0008/-0013-HHLQU (corrigidas pelo laboratorio na planilha de 14/10/2026)
#   Tags de Ubu (HHL...) usam o codigo com 4 digitos, como todas as tags ja validadas no MES (-0016-, -0008-, -0013-)
TAGS = {('SE', 0): 'M250010000-17-ABLQ', ('SE', 1): 'M4650060010-0017-HHLFU',
        ('325', 0): 'M250010000-16-ABLF', ('325', 1): 'M4650060010-0016-HHLFU',
        ('SiO2', 0): 'M250010000-4-ABLQ', ('SiO2', 1): 'M650030010-0004-HHLQU',
        ('P', 0): 'M250010000-8-ABLQ', ('P', 1): 'M650030010-0008-HHLQU',
        ('PPC', 0): 'M250010000-13-ABLQ', ('PPC', 1): 'M650030010-0013-HHLQU'}
MAPA_PADRAO = 'IP_MESVALOR'


def fmt_num(d, milhar=False):
    if d == 0:
        return '#,##0' if milhar else '0'
    return '0.' + '0' * d


def fmt_sinal(d, milhar=False):
    b = fmt_num(d, milhar)
    return '+%s;"−"%s;%s' % (b, b, b)


def dec_de(k):
    return PARAMS[PK.index(k)][4]


# ---------------------------------------------------------------- aba Batch: posicoes (0-based)
B_COL_N = 1               # B  Nº
B_COL_TU = 2              # C  data/hora Ubu
B_COL_TG = 3              # D  data/hora Germano
B_COL_V1 = 4              # E  1o valor (SE A); valor (p, lab) em B_COL_V1 + 2p + lab
B_COL_D1 = 15             # P  1a diferenca
B_COL_R1 = 20             # U  1o "rank" (par A/B para os graficos)
B_ROW_G0 = 16             # 1a linha de leituras
B_ROW_MED = B_ROW_G0 + NMAX
B_ROW_CNT = B_ROW_MED + 1

# ---------------------------------------------------------------- aba _Graf (0-based)
G_BLOCO = 9               # colunas por parametro
G_ROW_ST = 30             # tabela de calculos do farol (1a linha de parametro)
G_COL = dict(MA=1, MB=2, D=3, DR=4, ST=5, NP=6, TOL=7, NA=8, NB=9, DTXT=10, TTXT=11, NOKITEM=12, OKITEM=13,
             NOME=14)
G_ROW_X = G_ROW_ST + NP + 2   # linha dos totais / listas

# ---------------------------------------------------------------- Configurações (0-based)
C_ROW_TAG0 = 7            # 1a linha da tabela de tags
C_COL = dict(PARAM=1, LAB=2, TAG=3, SERV=4, MAPA=5, VMIN=6, VMAX=7, TIP=8)

# ---------------------------------------------------------------- Histórico
H_ROW_HDR = 6             # cabecalho da tabela tbl_Batches (0-based)
# Campos da aba Batch gravados no historico (nome da celula, coluna de tbl_Batches) - mesma ordem no VBA
CAMPOS = ([('bId', 'Batch'), ('bData', 'Data do report'), ('bMin', 'Mineroduto'),
           ('bIniU', 'Chegada Ubu (UBU3)'), ('bFimU', 'Término Ubu (UBU3)'), ('bNivel', 'Nível dos tanques'),
           ('bIniPMan', 'Início processo (manual)'),
           ('bIniG', 'Início Germano (EB06)'), ('bFimG', 'Término Germano (EB06)'), ('bNHor', 'Nº horários'),
           ('bLabA', 'Laboratório A'), ('bLabB', 'Laboratório B'), ('bRef', 'Referência'), ('bTit', 'Título'),
           ('bSub', 'Subtítulo')] +
          [('bTol_%s' % p[0], 'Tol %s' % p[1]) for p in PARAMS] +
          [('bIncFrase', 'Frase automática'), ('bIncAcomp', 'Acompanhamento no report')] +
          [('bCom%d' % i, 'Comentário %d' % i) for i in range(1, 7)])
H_COLS = ([h for _, h in CAMPOS] +
          sum([['%s média A' % p[1], '%s média B' % p[1], '%s diferença' % p[1], '%s status' % p[1]]
               for p in PARAMS], []) +
          ['Início no processo', 'Defasagem (h)', 'Nº NOK', 'Leituras', 'Salvo em', 'Salvo por'])
L_COLS = ['Batch', 'Ordem', 'Data/hora Ubu', 'Data/hora Germano', 'Parâmetro', 'Germano', 'Ubu',
          'Diferença', 'Tolerância', 'Status']


def quote(ws):
    return "'%s'" % ws.name


# ============================================================================
class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_properties({'title': 'Comparativo de Medições de Qualidade - Germano x Ubu',
                                'company': 'Samarco - Laboratório'})
        self.wb.set_calc_mode('auto')
        self._fmts = {}
        # batch aberto ao gerar: o mais completo do historico importado
        # a planilha abre com um batch em branco (o historico importado fica na aba Historico)
        self.atual = None

    # ----------------------------------------------------------- formatos
    def f(self, **kw):
        base = dict(font_name=FONTE, font_size=10, font_color=INK, valign='vcenter')
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self._fmts:
            self._fmts[chave] = self.wb.add_format(base)
        return self._fmts[chave]

    def fi(self, **kw):
        base = dict(bg_color=INPUT, border=1, border_color=LINE, locked=False)
        base.update(kw)
        return self.f(**base)

    def nome(self, nome, ws, r1, c1, r2=None, c2=None):
        ref = rc(r1, c1, True, True)
        if r2 is not None:
            ref += ':' + rc(r2, c2, True, True)
        self.wb.define_name(nome, '=%s!%s' % (quote(ws), ref))

    def ref(self, ws, r, c, absoluto=True):
        return '%s!%s' % (quote(ws), rc(r, c, absoluto, absoluto))

    def botao(self, ws, row, col, texto, macro, largura=180, altura=34, estilo='p', x=0, y=0):
        estilos = {'p': (NAVY2, BRANCO, None), 'g': (OK, BRANCO, None), 's': ('#E8EBF0', INK2, None),
                   'd': (BRANCO, NOK, '#F0C4C4'), 'n': (NAVY, BRANCO, None)}
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
        """Faixa azul-marinho no topo da aba (igual a barra do app HTML)."""
        ws.set_row(r, 34)
        ws.merge_range(r, c1, r, c2, titulo, self.f(bold=True, font_size=14, font_color=BRANCO, bg_color=NAVY,
                                                  indent=1))
        if sub is not None:
            ws.set_row(r + 1, 22)
            ws.merge_range(r + 1, c1, r + 1, c2, sub, self.f(font_size=9.5, bold=True, font_color='#D5DEEB',
                                                           bg_color=NAVY2, indent=1))

    def secao(self, ws, r, c1, c2, texto, altura=22):
        ws.set_row(r, altura)
        if c1 == c2:
            ws.write(r, c1, texto, self.f(bold=True, font_size=10, font_color=NAVY2, valign='bottom',
                                          bottom=2, bottom_color=NAVY2))
            return
        ws.merge_range(r, c1, r, c2, texto, self.f(bold=True, font_size=10, font_color=NAVY2, valign='bottom',
                                                  bottom=2, bottom_color=NAVY2))

    def rotulo(self, ws, r, c1, c2, texto):
        f = self.f(bold=True, font_size=8, font_color=MUTED, valign='bottom')
        if c2 > c1:
            ws.merge_range(r, c1, r, c2, texto.upper(), f)
        else:
            ws.write(r, c1, texto.upper(), f)

    def entrada(self, ws, r, c1, c2, valor, fmt, nome=None):
        if c2 > c1:
            ws.merge_range(r, c1, r, c2, '', fmt)
        if isinstance(valor, str) and valor.startswith('='):
            ws.write_formula(r, c1, valor, fmt)
        elif valor is None or valor == '':
            ws.write_blank(r, c1, None, fmt)
        else:
            ws.write(r, c1, valor, fmt)
        if nome:
            self.nome(nome, ws, r, c1)

    # ============================================================ abas
    def construir(self):
        wb = self.wb
        self.ws_pai = wb.add_worksheet('Painel')
        self.ws_bat = wb.add_worksheet('Batch')
        self.ws_env = wb.add_worksheet('Enviar')
        self.ws_his = wb.add_worksheet('Histórico')
        self.ws_bdl = wb.add_worksheet('BD_Leituras')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_gra = wb.add_worksheet('_Graf')
        for ws, cn in [(self.ws_pai, 'shPainel'), (self.ws_bat, 'shBatch'), (self.ws_env, 'shEnviar'),
                       (self.ws_his, 'shHist'), (self.ws_bdl, 'shLeituras'),
                       (self.ws_cfg, 'shConfig'),
                       (self.ws_mes, 'shDadosMES'), (self.ws_gra, 'shGraf')]:
            ws.set_vba_name(cn)
            if ws not in (self.ws_bdl, self.ws_gra, self.ws_mes):
                ws.hide_gridlines(2)
        self.ws_pai.set_tab_color(NAVY2)
        self.ws_bat.set_tab_color('#FDB913')
        self.ws_env.set_tab_color(BLUE)
        self.ws_his.set_tab_color(OK)
        self.ws_bdl.set_tab_color(MUTED)
        self.ws_cfg.set_tab_color('#9AA3AD')

        self.aba_config()
        self.aba_batch()
        self.aba_graf()
        self.aba_painel()
        self.aba_enviar()
        self.aba_historico()
        self.aba_dados_mes()
        self.ws_mes.hide()
        self.ws_gra.hide()
        self.ws_pai.activate()

    # ------------------------------------------------------------ Configurações
    def aba_config(self):
        ws = self.ws_cfg
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 30)
        ws.set_column(2, 2, 14)
        ws.set_column(3, 3, 30)
        ws.set_column(4, 5, 15)
        ws.set_column(6, 8, 13)
        ws.set_column(9, 9, 2)
        ws.set_column(10, 10, 60)
        self.faixa(ws, 1, 1, 10, 'CONFIGURAÇÕES',
                   'Tags do MES, consulta, padrões de um batch novo e títulos do report')
        self.botao(ws, 1, 8, 'Voltar ao Batch', 'IrBatch', 150, 26, 'n', x=80, y=4)

        self.secao(ws, 4, 1, 8, 'TAGS DO MES (UMA POR ANÁLISE E LABORATÓRIO)')
        ws.merge_range(5, 1, 5, 8, 'Tag em branco = análise fora da consulta (fica para lançamento manual). '
                       'Servidor e mapa como no MES (padrão UBU / IP_MESVALOR). Valores fora da faixa '
                       'válida são descartados.', self.f(font_size=8.5, font_color=MUTED, text_wrap=True))
        ws.set_row(5, 26)
        hdr = ['Análise', 'Laboratório', 'Tag', 'Servidor', 'Mapa', 'Mín. válido', 'Máx. válido',
               'Típico (simulação)']
        fh = self.f(bold=True, font_size=9, font_color=BRANCO, bg_color=NAVY2, align='center', border=1,
                    border_color=NAVY2)
        for i, h in enumerate(hdr):
            ws.write(C_ROW_TAG0 - 1, 1 + i, h, fh)
        fl = self.f(border=1, border_color=LINE2, bold=True)
        fc = self.f(border=1, border_color=LINE2, align='center')
        for p, pr in enumerate(PARAMS):
            for lab in range(2):
                r = C_ROW_TAG0 + p * 2 + lab
                ws.write(r, C_COL['PARAM'], pr[1], fl)
                ws.write(r, C_COL['LAB'], LABS[lab], fc)
                ws.write_string(r, C_COL['TAG'], TAGS.get((pr[0], lab), ''), self.fi(font_name='Consolas', font_size=9.5))
                ws.write(r, C_COL['SERV'], SERVIDOR_PADRAO, self.fi(align='center'))
                ws.write(r, C_COL['MAPA'], MAPA_PADRAO, self.fi(align='center'))
                ws.write(r, C_COL['VMIN'], pr[6][0], self.fi(align='center'))
                ws.write(r, C_COL['VMAX'], pr[6][1], self.fi(align='center'))
                ws.write(r, C_COL['TIP'], pr[7 + lab], self.fi(align='center'))
        self.nome('cfgTags', ws, C_ROW_TAG0, 1, C_ROW_TAG0 + 2 * NP - 1, 8)
        nota = ('As tags vêm das telas do MES (ex.: M650030010-0004-HHLQU). Germano e Ubu podem estar em '
                'servidores diferentes: informe o servidor de cada tag.')
        ws.merge_range(C_ROW_TAG0, 10, C_ROW_TAG0 + 3, 10, nota,
                       self.f(font_size=9, font_color=INK2, text_wrap=True, valign='top', bg_color='#F5F8FD',
                              left=5, left_color=NAVY2))

        r = C_ROW_TAG0 + 2 * NP + 1
        self.secao(ws, r, 1, 8, 'CONSULTA AO MES')
        fr = self.f(bold=True, font_size=9.5)
        fd = self.f(font_size=8.5, font_color=MUTED, text_wrap=True)
        itens = [
            ('cfgFonte', 'Fonte dos dados', 'MES', 'MES ou SIMULAÇÃO (dados fictícios, para treino)'),
            ('cfgTimeout', 'Tempo máximo de espera (s)', 120, 'Quanto esperar a resposta do Aspen'),
            ('cfgFormatoData', 'Formato da data para o MES', 'Texto dd/mm/aaaa',
             'Data do Excel, Texto dd/mm/aaaa ou Texto mm/dd/aaaa'),
            ('cfgMinAntes', 'Horário da linha após o início da janela (min)', 30,
             'Janela de 2 h que começa às 11:00 aparece como 11:30 (início + 30 min)'),
            ('cfgGradeU', 'Grade das janelas de Ubu (hora)', dtm.time(7, 0),
             'Janelas de 2 h a partir desta hora (07:00, 09:00, 11:00...). 1ª janela = 1ª completa após o início'),
            ('cfgGradeG', 'Grade das janelas de Germano (hora)', dtm.time(0, 0),
             'Horas pares (00:00, 02:00...): os resultados de Germano (01:00, 03:00...) caem no meio da janela'),
            ('cfgMinG', 'Horário de Germano após o início da janela (min)', 60,
             'Janela 00:00-02:00 aparece como 01:00 (horário do resultado em Germano; só referência)'),
            ('cfgCalc', 'Tipo de cálculo do Aspen', '1', '1 = média da janela (como nos relatórios do físico)'),
            ('cfgFaseG', 'Fase de Germano no evento do MES', 'EB06', 'Linha usada ao colar as Fases do Evento'),
            ('cfgFaseU', 'Fase de Ubu no evento do MES', 'UBU', 'Começo do código (UBU3, UBU1...)'),
        ]
        self.cfg_rows = {}
        for i, (nm, rot, val, desc) in enumerate(itens):
            rr = r + 1 + i
            ws.write(rr, 1, rot, fr)
            ws.merge_range(rr, 2, rr, 3, '', self.fi(align='center'))
            if isinstance(val, str):
                ws.write_string(rr, 2, val, self.fi(align='center'))
            elif isinstance(val, dtm.time):
                ws.write_datetime(rr, 2, dtm.datetime.combine(dtm.date(1899, 12, 31), val), self.fi(align='center', num_format='hh:mm'))
            else:
                ws.write(rr, 2, val, self.fi(align='center'))
            ws.merge_range(rr, 4, rr, 8, desc, fd)
            self.nome(nm, ws, rr, 2)
            self.cfg_rows[nm] = rr
        ws.data_validation(self.cfg_rows['cfgFonte'], 2, self.cfg_rows['cfgFonte'], 2,
                           {'validate': 'list', 'source': ['MES', 'SIMULAÇÃO']})
        ws.data_validation(self.cfg_rows['cfgFormatoData'], 2, self.cfg_rows['cfgFormatoData'], 2,
                           {'validate': 'list', 'source': ['Data do Excel', 'Texto dd/mm/aaaa', 'Texto mm/dd/aaaa']})
        self.botao(ws, r + 1, 10, 'Ver consultas (Dados_MES)', 'MostrarDadosMES', 200, 28, 's', x=10)

        r = r + 1 + len(itens) + 1
        self.secao(ws, r, 1, 8, 'NÍVEL DOS TANQUES EM UBU → ATRASO ATÉ O BATCH PREVALECER NO PROCESSO')
        ws.merge_range(r + 1, 1, r + 1, 8, 'Horas somadas à chegada em Ubu (UBU3) para achar o início da comparação. '
                       'Alto 8 h, Médio 6 h, Baixo 4 h (informado pelo laboratório).', fd)
        ws.set_row(r + 1, 24)
        # valores informados pelo laboratorio (14/10/2026): Alto 8 h, Medio 6 h, Baixo 4 h
        for i, (nivel, horas) in enumerate([('Alto', 8), ('Médio', 6), ('Baixo', 4)]):
            rr = r + 2 + i
            ws.write_string(rr, 1, nivel, fr)
            ws.write_number(rr, 2, horas, self.fi(align='center', num_format='0.0" h"'))
        self.nome('cfgAtrasoN', ws, r + 2, 1, r + 4, 1)
        self.nome('cfgAtrasoH', ws, r + 2, 2, r + 4, 2)
        self.wb.define_name('lstNivel', "='%s'!$B$%d:$B$%d" % (ws.name, r + 3, r + 5))
        r = r + 6
        self.secao(ws, r, 1, 8, 'PADRÕES DE UM BATCH NOVO')
        pad = [('padLabA', 'Laboratório A (barra azul)', 'Germano'),
               ('padLabB', 'Laboratório B (barra cinza)', 'Ubu'),
               ('padRef', 'Laboratório de referência', 'Germano'),
               ('padMin', 'Mineroduto', '03'),
               ('padNHor', 'Nº de horários', NHOR_PADRAO),
               ('padTit', 'Título principal', 'COMPARATIVO DE MEDIÇÕES DE QUALIDADE'),
               ('padSub', 'Subtítulo', 'Batch bombeado e pellet feed recebido'),
               ('padCom', 'Comentário padrão (1º tópico)', 'Ponto de amostragem do PF recebido na descarga '
                                                        'da filtragem.')]
        for k, pr in enumerate(PARAMS):
            pad.append(('padTol_%s' % pr[0], 'Tolerância %s (±)' % pr[1], pr[5]))
        for i, (nm, rot, val) in enumerate(pad):
            rr = r + 1 + i
            ws.write(rr, 1, rot, fr)
            ws.merge_range(rr, 2, rr, 5, '', self.fi())
            if isinstance(val, str):
                ws.write_string(rr, 2, val, self.fi())
            else:
                ws.write(rr, 2, val, self.fi(num_format='General', align='left'))
            self.nome(nm, ws, rr, 2)
        r = r + 1 + len(pad) + 1
        self.secao(ws, r, 1, 8, 'TÍTULOS DOS GRÁFICOS')
        for i, pr in enumerate(PARAMS):
            rr = r + 1 + i
            ws.write(rr, 1, pr[1], fr)
            ws.merge_range(rr, 2, rr, 5, pr[3], self.fi())
            self.nome('tit_%s' % pr[0], ws, rr, 2)
        r = r + 1 + NP + 1
        # listas (validacao de dados)
        self.secao(ws, r, 1, 8, 'LISTAS')
        ws.write(r + 1, 1, 'Mineroduto', fr)
        for i, v in enumerate(['01', '02', '03']):
            ws.write_string(r + 1, 2 + i, v, self.fi(align='center'))
        self.nome('lstMin', ws, r + 1, 2, r + 1, 4)
        ws.write(r + 2, 1, 'Sim / Não', fr)
        ws.write_string(r + 2, 2, 'Sim', self.f(align='center'))
        ws.write_string(r + 2, 3, 'Não', self.f(align='center'))
        self.nome('lstSimNao', ws, r + 2, 2, r + 2, 3)
        ws.write(r + 3, 1, 'Laboratórios do batch', fr)
        ws.write_formula(r + 3, 2, '=bLabA', self.f(align='center'))
        ws.write_formula(r + 3, 3, '=bLabB', self.f(align='center'))
        self.nome('lstLabs', ws, r + 3, 2, r + 3, 3)
        self.cfg_fim = r + 4
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Batch
    def aba_batch(self):
        ws = self.ws_bat
        b = self.atual
        ws.set_column(0, 0, 1.5)
        ws.set_column(1, 1, 5)
        ws.set_column(2, 2, 25)
        ws.set_column(3, 3, 16.5)
        ws.set_column(4, 13, 9.3)
        ws.set_column(14, 14, 1.5)
        ws.set_column(15, 19, 9.3)
        ws.set_column(20, 24, 6, None, {'hidden': True})
        ws.set_column(25, 25, 2)
        ws.set_default_row(18)
        ws.set_row(1, 34)
        ws.merge_range(1, 1, 1, 19, 'DADOS DO BATCH', self.f(bold=True, font_size=14, font_color=BRANCO, bg_color=NAVY,
                                                           indent=1))
        fdt = self.fi(align='center', num_format='dd/mm/yyyy hh:mm')
        fauto_dt = self.f(align='center', num_format='dd/mm/yy hh:mm', bg_color='#F3F7FC', border=1,
                          border_color=LINE2, locked=False, font_size=9.5)
        fauto_v = self.f(align='center', bg_color='#F3F7FC', border=1, border_color=LINE2, locked=False,
                         font_size=9.5)
        fcalc_h = self.f(align='center', bg_color=SEC, border=1, border_color=LINE2, num_format='0.0" h"',
                         font_color=INK2, font_size=9.5)
        # ① ② ③ lado a lado
        ws.set_row(2, 8)
        self.secao(ws, 3, 1, 2, '①  BATCH Nº')
        self.secao(ws, 3, 3, 7, '②  COLE AS FASES DO EVENTO (MES)')
        self.secao(ws, 3, 8, 11, '③  NÍVEL DOS TANQUES')
        ws.merge_range(4, 1, 5, 2, '', self.fi(align='center', bold=True, font_size=16))
        if b:
            ws.write_number(4, 1, b['id'], self.fi(align='center', bold=True, font_size=16))
        self.nome('bId', ws, 4, 1)
        fpaste = self.fi(font_size=8, num_format='dd/mm hh:mm', border=4, border_color='#C9A227')
        for r in range(4, 9):
            for c in range(3, 8):
                ws.write_blank(r, c, None, fpaste)
        self.nome('bFases', ws, 4, 3, 8, 7)
        ws.merge_range(9, 3, 9, 7, '', self.f(font_size=8.5, font_color=OK, bold=True, text_wrap=True))
        self.nome('bFasesMsg', ws, 9, 3)
        ws.merge_range(4, 8, 5, 11, '', self.fi(align='center', bold=True, font_size=13))
        if b and b.get('nivel'):
            ws.write_string(4, 8, b['nivel'].capitalize(), self.fi(align='center', bold=True, font_size=13))
        self.nome('bNivel', ws, 4, 8)
        ws.data_validation(4, 8, 4, 8, {'validate': 'list', 'source': '=lstNivel'})
        ws.merge_range(6, 8, 6, 9, 'Atraso no processo', self.f(font_size=8.5, font_color=MUTED, indent=1))
        ws.merge_range(6, 10, 6, 11, '', fcalc_h)
        ws.write_formula(6, 10, '=IF(AND(ISNUMBER(bIniP),ISNUMBER(bIniU)),ROUND((bIniP-bIniU)*24,2),'
                                'IF(bNivel<>"",gAtrasoNivel,""))', fcalc_h)
        self.nome('bAtraso', ws, 6, 10)
        for r in range(4, 9):
            ws.set_row(r, 17)
        ws.set_row(9, 16)
        # conferencia (preenchida sozinha; da para corrigir)
        ws.set_row(10, 6)
        fl = self.f(bold=True, font_size=7.5, font_color=MUTED, valign='bottom')
        conf = [(1, 2, 'Germano · início (EB06)', 'bIniG', fauto_dt), (3, 3, 'Germano · término', 'bFimG', fauto_dt),
                (4, 5, 'Ubu · chegada (UBU3)', 'bIniU', fauto_dt), (6, 7, 'Ubu · término', 'bFimU', fauto_dt)]
        for c1, c2, rot, nm, fmt in conf:
            if c2 > c1:
                ws.merge_range(11, c1, 11, c2, rot.upper(), fl)
            else:
                ws.write(11, c1, rot.upper(), fl)
            self.entrada(ws, 12, c1, c2, None, fmt, nm)
        ws.merge_range(11, 8, 11, 9, '► INÍCIO NO PROCESSO', self.f(bold=True, font_size=7.5, font_color=OK,
                                                                   valign='bottom'))
        fip = self.f(align='center', bold=True, font_color=OK, bg_color=OK_BG, border=1, border_color='#B9DEC9',
                     num_format='dd/mm/yy hh:mm', font_size=9.5)
        ws.merge_range(12, 8, 12, 9, '', fip)
        ws.write_formula(12, 8, '=IF(ISNUMBER(bIniPMan),bIniPMan,IF(ISNUMBER(bIniU),bIniU+N(gAtrasoNivel)/24,""))', fip)
        self.nome('bIniP', ws, 12, 8)
        ws.merge_range(11, 10, 11, 11, 'BI-HORÁRIOS', fl)
        self.entrada(ws, 12, 10, 11, len(b['rows']) if b else
                     '=IF(ISNUMBER(gNGer),gNGer,IF(ISNUMBER(gNHorFim),gNHorFim,padNHor))', fauto_v, 'bNHor')
        ws.data_validation(12, 10, 12, 10, {'validate': 'integer', 'criteria': 'between', 'minimum': 1,
                                            'maximum': NMAX, 'error_message': 'Entre 1 e %d horários.' % NMAX})
        ws.merge_range(11, 12, 11, 13, 'DEFASAGEM', fl)
        ws.merge_range(12, 12, 12, 13, '', fcalc_h)
        ws.write_formula(12, 12, '=IF(AND(ISNUMBER(gJanU),ISNUMBER(gJanG)),ROUND((gJanU-gJanG)*24,2),"")', fcalc_h)
        self.nome('bDef', ws, 12, 12)
        fdata = self.f(align='center', num_format='dd/mm/yyyy', bg_color='#F3F7FC', border=1, border_color=LINE2,
                       locked=False, font_size=9.5)
        ws.merge_range(11, 15, 11, 16, 'DATA DO REPORT', fl)
        ws.merge_range(12, 15, 12, 16, '', fdata)
        if b:
            ws.write_datetime(12, 15, dtm.datetime.strptime(b.get('chegada') or b['rows'][0]['data'], '%Y-%m-%d'), fdata)
        else:
            ws.write_formula(12, 15, '=TODAY()', fdata)
        self.nome('bData', ws, 12, 15)
        ws.merge_range(11, 17, 11, 18, 'MINERODUTO', fl)
        self.entrada(ws, 12, 17, 18, '03', fauto_v, 'bMin')
        ws.data_validation(12, 17, 12, 17, {'validate': 'list', 'source': '=lstMin'})
        ws.set_row(11, 14)
        ws.set_row(12, 22)
        ws.set_row(13, 8)

        # --- leituras
        self.secao(ws, B_ROW_G0 - 2, 1, 19, 'LEITURAS POR HORÁRIO  (em branco = sem análise no laboratório)')
        r0 = B_ROW_G0
        fgh = self.f(bold=True, font_size=8.5, font_color=BRANCO, bg_color=NAVY2, align='center', border=1,
                     border_color=BRANCO, text_wrap=True)
        fa = self.f(bold=True, font_size=8.5, font_color=BRANCO, bg_color=BLUE, align='center', border=1,
                    border_color=BRANCO, text_wrap=True)
        fb = self.f(bold=True, font_size=8.5, font_color=BRANCO, bg_color='#6B7480', align='center', border=1,
                    border_color=BRANCO, text_wrap=True)
        fdh = self.f(bold=True, font_size=8.5, font_color=BRANCO, bg_color='#B97A12', align='center', border=1,
                     border_color=BRANCO)
        ws.set_row(r0 - 1, 28)
        ws.write(r0 - 1, B_COL_N, 'Nº', fgh)
        ws.write(r0 - 1, B_COL_TU, 'Horário oficial\n(Ubu)', fgh)
        ws.write(r0 - 1, B_COL_TG, 'Origem em Germano\n(só referência)', fgh)
        for p, pr in enumerate(PARAMS):
            ws.write_formula(r0 - 1, B_COL_V1 + 2 * p, '="%s"&CHAR(10)&bLabA' % pr[1], fa)
            ws.write_formula(r0 - 1, B_COL_V1 + 2 * p + 1, '="%s"&CHAR(10)&bLabB' % pr[1], fb)
            ws.write(r0 - 1, B_COL_D1 + p, 'Δ ' + pr[1], fdh)
            ws.write(r0 - 1, B_COL_R1 + p, 'par ' + pr[0], fgh)
        fn = self.f(align='center', font_size=9, font_color=MUTED, border=1, border_color=LINE2)
        ftu = self.f(align='center', num_format='dd/mm hh:mm', border=1, border_color=LINE2, locked=False,
                     bg_color='#F7F9FC', font_size=9.5)
        ftg = self.f(align='center', num_format='dd/mm hh:mm', border=1, border_color=LINE2, font_size=9.5,
                     font_color=INK2)
        rows = b['rows'] if b else []
        for i in range(NMAX):
            r = r0 + i
            ws.write(r, B_COL_N, i + 1, fn)
            ws.write_formula(r, B_COL_TU, self.f_hora_u(i), ftu)
            ws.write_formula(r, B_COL_TG, self.f_hora_g(i, rc(r, B_COL_TU)), ftg)
            for p, pr in enumerate(PARAMS):
                for lab in range(2):
                    c = B_COL_V1 + 2 * p + lab
                    fmt = self.fi(align='center', num_format=fmt_num(pr[4]), font_size=9.5,
                                  bg_color='#FFFFFF' if lab == 0 else '#FAFBFC', border_color=LINE2)
                    v = rows[i].get('%s_%s' % (pr[0], 'GU'[lab])) if i < len(rows) else None
                    if v is None:
                        ws.write_blank(r, c, None, fmt)
                    else:
                        ws.write_number(r, c, v, fmt)
                ca, cb = rc(r, B_COL_V1 + 2 * p), rc(r, B_COL_V1 + 2 * p + 1)
                ws.write_formula(r, B_COL_D1 + p, '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),IF(bRef=bLabA,%s-%s,%s-%s),"")'
                                 % (ca, cb, cb, ca, ca, cb),
                                 self.f(align='center', num_format=fmt_sinal(pr[4]), font_size=9.5, border=1,
                                        border_color=LINE2, bg_color='#FFF7EA', font_color='#8A5A00'))
                a1 = rc(r0, B_COL_V1 + 2 * p, True, False)
                b1 = rc(r0, B_COL_V1 + 2 * p + 1, True, False)
                ws.write_formula(r, B_COL_R1 + p,
                                 '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),SUMPRODUCT(ISNUMBER(%s:%s)*ISNUMBER(%s:%s)),"")'
                                 % (ca, cb, a1, ca, b1, cb), self.f(font_size=8, align='center'))
        # media e contagem
        fm = self.f(bold=True, align='center', bg_color=TH, border=1, border_color=LINE, font_size=9.5)
        ws.merge_range(B_ROW_MED, B_COL_N, B_ROW_MED, B_COL_TG, 'MÉDIA', fm)
        ws.merge_range(B_ROW_CNT, B_COL_N, B_ROW_CNT, B_COL_TG, 'Nº de leituras',
                       self.f(align='center', font_size=8.5, font_color=MUTED, border=1, border_color=LINE2))
        g1, g2 = B_ROW_G0, B_ROW_G0 + NMAX - 1
        for p, pr in enumerate(PARAMS):
            for lab in range(2):
                c = B_COL_V1 + 2 * p + lab
                rng = '%s:%s' % (rc(g1, c), rc(g2, c))
                ws.write_formula(B_ROW_MED, c, '=IF(COUNT(%s)=0,"",AVERAGE(%s))' % (rng, rng),
                                 self.f(bold=True, align='center', bg_color=TH, border=1, border_color=LINE,
                                        font_size=9.5, num_format=fmt_num(pr[4])))
                ws.write_formula(B_ROW_CNT, c, '=COUNT(%s)' % rng,
                                 self.f(align='center', font_size=8.5, font_color=MUTED, border=1,
                                        border_color=LINE2))
            ma, mb = rc(B_ROW_MED, B_COL_V1 + 2 * p), rc(B_ROW_MED, B_COL_V1 + 2 * p + 1)
            ws.write_formula(B_ROW_MED, B_COL_D1 + p,
                             '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),IF(bRef=bLabA,%s-%s,%s-%s),"")'
                             % (ma, mb, mb, ma, ma, mb),
                             self.f(bold=True, align='center', bg_color='#FBE3BC', border=1, border_color=LINE,
                                    font_size=9.5, num_format=fmt_sinal(pr[4]), font_color='#7A4A00'))
            rngd = '%s:%s' % (rc(g1, B_COL_D1 + p), rc(g2, B_COL_D1 + p))
            ws.write_formula(B_ROW_CNT, B_COL_D1 + p, '=COUNT(%s)' % rngd,
                             self.f(align='center', font_size=8.5, font_color=MUTED, border=1, border_color=LINE2))
        # diferenca fora da tolerancia: em vermelho (por leitura e na media)
        for p, pr in enumerate(PARAMS):
            c = B_COL_D1 + p
            cel = rc(g1, c, False, False)
            ws.conditional_format(g1, c, B_ROW_MED, c, {
                'type': 'formula',
                'criteria': '=AND(ISNUMBER(%s),ABS(ROUND(%s,%d))>bTol_%s+1E-9)' % (cel, cel, pr[4], pr[0]),
                'format': self.wb.add_format({'font_color': NOK, 'bold': True, 'bg_color': NOK_BG})})
        # linhas alem do Nº de horarios: cinza
        ws.conditional_format(g1, B_COL_N, g2, B_COL_TG, {
            'type': 'formula', 'criteria': '=$B%d>bNHor' % (g1 + 1),
            'format': self.wb.add_format({'font_color': '#B0B8C4'})})
        self.nome('bGrade', ws, g1, B_COL_N, g2, B_COL_R1 + NP - 1)
        self.nome('bValores', ws, g1, B_COL_V1, g2, B_COL_V1 + 2 * NP - 1)
        self.nome('bHorasU', ws, g1, B_COL_TU, g2, B_COL_TU)

        # --- comentarios
        r = B_ROW_CNT + 2
        self.secao(ws, r, 1, 13, 'COMENTÁRIOS (CADA LINHA VIRA UM TÓPICO NO REPORT)')
        fauto = self.f(italic=True, font_size=9.5, font_color=INK2, bg_color=SEC, border=1, border_color=LINE2,
                       text_wrap=True)
        for rr, rot, ref, nm, nm_inc in [(r + 1, 'Frase automática', rc(G_ROW_X, 6, True, True), 'bFrase', 'bIncFrase'),
                                         (r + 2, 'Acompanhamento', None, 'bAcomp', 'bIncAcomp')]:
            ws.merge_range(rr, 1, rr, 2, rot, self.f(bold=True, font_size=9, font_color=MUTED))
            ws.merge_range(rr, 3, rr, 11, '', fauto)
            ws.write_formula(rr, 3, ("=_Graf!%s" % ref) if ref else '=gAcomp', fauto)
            self.nome(nm, ws, rr, 3)
            ws.write(rr, 12, 'No report?', self.f(bold=True, font_size=8, font_color=MUTED, align='right'))
            ws.write_string(rr, 13, 'Sim', self.fi(align='center'))
            self.nome(nm_inc, ws, rr, 13)
            ws.data_validation(rr, 13, rr, 13, {'validate': 'list', 'source': '=lstSimNao'})
            ws.set_row(rr, 28)
        com = self.comentarios(b)
        for i in range(6):
            rr = r + 3 + i
            ws.merge_range(rr, 1, rr, 2, 'Tópico %d' % (i + 1), self.f(bold=True, font_size=9, font_color=MUTED))
            ws.merge_range(rr, 3, rr, 13, '', self.fi(text_wrap=True))
            if i < len(com):
                ws.write_string(rr, 3, com[i], self.fi(text_wrap=True))
            self.nome('bCom%d' % (i + 1), ws, rr, 3)

        # --- ajustes do report (opcional)
        r = r + 10
        self.secao(ws, r, 1, 13, 'AJUSTES DO REPORT (OPCIONAL — JÁ VÊM DOS PADRÕES DA ABA CONFIGURAÇÕES)')
        fv = self.fi(align='center')
        campos = [
            (r + 1, 1, 2, 'Laboratório A (azul)', '=padLabA', fv, 'bLabA'),
            (r + 1, 3, 3, 'Laboratório B (cinza)', '=padLabB', fv, 'bLabB'),
            (r + 1, 4, 5, 'Referência', '=padRef', fv, 'bRef'),
            (r + 3, 1, 5, 'Título principal', '=padTit', self.fi(), 'bTit'),
            (r + 3, 6, 13, 'Subtítulo do cabeçalho', '=padSub', self.fi(), 'bSub'),
        ]
        for rl, c1, c2, rot, val, fmt, nm in campos:
            self.rotulo(ws, rl, c1, c2, rot)
            self.entrada(ws, rl + 1, c1, c2, val, fmt, nm)
        ws.data_validation(r + 2, 4, r + 2, 4, {'validate': 'list', 'source': '=lstLabs'})
        self.rotulo(ws, r + 1, 6, 7, 'Diferença')
        ws.merge_range(r + 2, 6, r + 2, 7, '', self.f(align='center', font_color=INK2, font_size=9, bold=True))
        ws.write_formula(r + 2, 6, '=IF(bRef=bLabA,bLabB&" − "&bLabA,bLabA&" − "&bLabB)',
                         self.f(align='center', font_color=INK2, font_size=9, bold=True))
        self.rotulo(ws, r + 1, 8, 10, 'Início no processo (ajuste manual)')
        self.entrada(ws, r + 2, 8, 10, None, self.fi(align='center', num_format='dd/mm/yyyy hh:mm'), 'bIniPMan')
        self.rotulo(ws, r + 5, 1, 13, 'Tolerâncias deste batch (|diferença| máxima aceita) — critério OK/NOK do farol')
        pos_tol = [(1, 2), (3, 3), (4, 5), (6, 7), (8, 9)]
        for p, pr in enumerate(PARAMS):
            c1, c2 = pos_tol[p]
            self.rotulo(ws, r + 6, c1, c2, pr[1] + ' (±)')
            self.entrada(ws, r + 7, c1, c2, '=padTol_%s' % pr[0], self.fi(align='center', num_format=fmt_num(pr[4])),
                         'bTol_%s' % pr[0])
        self.b_fim = r + 7

        # botoes (na faixa do topo e ao lado dos passos)
        self.botao(ws, 3, 13, 'Buscar no MES', 'BuscarMES', 165, 40, 'p', x=6, y=4)
        self.botao(ws, 3, 17, 'Salvar no histórico', 'SalvarHistorico', 170, 40, 'g', x=4, y=4)
        self.botao(ws, 6, 13, '＋ Novo batch', 'NovoBatch', 104, 26, 's', x=6, y=4)
        self.botao(ws, 6, 15, 'Limpar leituras', 'LimparLeituras', 118, 26, 'd', x=26, y=4)
        self.botao(ws, 6, 17, 'Ver painel →', 'IrPainel', 106, 26, 'n', x=68, y=4)
        ws.print_area(0, 0, self.b_fim, 19)
        ws.fit_to_pages(1, 0)
        ws.set_paper(9)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.protect('', {'format_columns': True, 'format_rows': True, 'select_locked_cells': True,
                        'select_unlocked_cells': True})

    @staticmethod
    def f_hora_u(i):
        """Horario da linha i (0-based) em Ubu: 1a janela + i x 2 h + 30 min (ate o Nº de horarios).
        (mesma formula em ModMES.RestaurarHorarios)"""
        return '=IF(AND(ISNUMBER(gJanU),%d<=bNHor),gJanU+%d/12+cfgMinAntes/1440,"")' % (i + 1, i)

    @staticmethod
    def f_hora_g(i, cel_u):
        """Horario da linha i em Germano: 1a janela de Germano + i x 2 h (ate o termino); sem inicio = Ubu."""
        return ('=IF(ISNUMBER(%s),IF(ISNUMBER(gJanG),IF(%d<=gNG,gJanG+%d/12+cfgMinG/1440,""),%s),"")'
                % (cel_u, i + 1, i, cel_u))

    # ------------------------------------------------------------ _Graf (series e farol)
    def aba_graf(self):
        ws = self.ws_gra
        bat = quote(self.ws_bat)
        g1, g2 = B_ROW_G0, B_ROW_G0 + NMAX - 1

        def bcol(c):
            return '%s!%s:%s' % (bat, rc(g1, c, True, True), rc(g2, c, True, True))

        # ---- calculos do farol: linha G_ROW_ST + p
        hdr = ['Parâmetro', 'Média A', 'Média B', 'Diferença', 'Dif. arred.', 'Status', 'Pares', 'Tol',
               'Leituras A', 'Leituras B', 'Dif. texto', 'Tol. texto', 'Item NOK', 'Item OK', 'Nome']
        for i, h in enumerate(hdr):
            ws.write(G_ROW_ST - 1, i, h)
        self.st = {}
        for p, pr in enumerate(PARAMS):
            r = G_ROW_ST + p
            k = pr[0]
            dec = pr[4]
            ca = bcol(B_COL_V1 + 2 * p)
            cb = bcol(B_COL_V1 + 2 * p + 1)
            cell = lambda key: rc(r, G_COL[key])
            ws.write(r, 0, pr[1])
            ws.write_formula(r, G_COL['MA'], '=IF(COUNT(%s)=0,"",AVERAGE(%s))' % (ca, ca))
            ws.write_formula(r, G_COL['MB'], '=IF(COUNT(%s)=0,"",AVERAGE(%s))' % (cb, cb))
            ma, mb = cell('MA'), cell('MB')
            ws.write_formula(r, G_COL['D'], '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),IF(bRef=bLabA,%s-%s,%s-%s),"")'
                             % (ma, mb, mb, ma, ma, mb))
            d = cell('D')
            ws.write_formula(r, G_COL['DR'], '=IF(ISNUMBER(%s),ROUND(%s,%d),"")' % (d, d, dec))
            ws.write_formula(r, G_COL['TOL'], '=bTol_%s' % k)
            ws.write_formula(r, G_COL['ST'], '=IF(ISNUMBER(%s),IF(ABS(%s)<=%s+1E-9,"OK","NOK"),"—")'
                             % (d, cell('DR'), cell('TOL')))
            ws.write_formula(r, G_COL['NP'], '=MAX(0,%s)' % bcol(B_COL_R1 + p))
            ws.write_formula(r, G_COL['NA'], '=COUNT(%s)' % ca)
            ws.write_formula(r, G_COL['NB'], '=COUNT(%s)' % cb)
            ws.write_formula(r, G_COL['DTXT'], '=IF(ISNUMBER(%s),IF(%s>0,"+",IF(%s<0,"−",""))&FIXED(ABS(%s),%d),"—")'
                             % (d, cell('DR'), cell('DR'), cell('DR'), dec))
            ws.write_formula(r, G_COL['TTXT'], '=FIXED(%s,%d)' % (cell('TOL'), dec))
            ws.write_formula(r, G_COL['NOKITEM'], '="%s (dif "&%s&", tol ±"&%s&")"' % (pr[1], cell('DTXT'), cell('TTXT')))
            ws.write_formula(r, G_COL['OKITEM'], '="%s"' % pr[1])
            ws.write(r, G_COL['NOME'], pr[2])
            self.st[k] = {key: self.ref(ws, r, G_COL[key]) for key in G_COL}
        # ---- totais e listas (linha G_ROW_X)
        rs = G_ROW_ST
        re_ = G_ROW_ST + NP - 1
        stc = '%s:%s' % (rc(rs, G_COL['ST'], True, True), rc(re_, G_COL['ST'], True, True))
        X = G_ROW_X
        ws.write(X - 1, 0, 'NOK')
        ws.write(X - 1, 1, 'OK')
        ws.write(X - 1, 2, 'Avaliados')
        ws.write(X - 1, 3, 'Lista NOK (nomes)')
        ws.write(X - 1, 4, 'Lista NOK (detalhe)')
        ws.write(X - 1, 5, 'Lista OK')
        ws.write(X - 1, 6, 'Frase automática')
        ws.write(X - 1, 7, 'Comentários (tópicos)')
        ws.write_formula(X, 0, '=COUNTIF(%s,"NOK")' % stc)
        ws.write_formula(X, 1, '=COUNTIF(%s,"OK")' % stc)
        ws.write_formula(X, 2, '=%s+%s' % (rc(X, 0), rc(X, 1)))

        def juntar(coluna, cond, sep_meio, sep_fim, n_ref):
            s0 = '&'.join('IF(%s="%s",%s&"|","")' % (rc(rs + p, G_COL['ST'], True, True), cond,
                                                     rc(rs + p, coluna, True, True)) for p in range(NP))
            s = 'LEFT(%s,LEN(%s)-1)' % ('(' + s0 + ')', '(' + s0 + ')')
            return ('=IF(%s=0,"",IF(%s=1,%s,SUBSTITUTE(SUBSTITUTE(%s,"|","%s",%s-1),"|","%s")))'
                    % (n_ref, n_ref, s, s, sep_fim, n_ref, sep_meio))
        ws.write_formula(X, 3, juntar(G_COL['NOME'], 'NOK', ', ', ' e ', rc(X, 0)))
        ws.write_formula(X, 4, juntar(G_COL['NOKITEM'], 'NOK', '; ', '; ', rc(X, 0)))
        ws.write_formula(X, 5, juntar(G_COL['OKITEM'], 'OK', ', ', ', ', rc(X, 1)))
        ws.write_formula(X, 6, '=IF(%s=0,"",IF(%s=1,"O parâmetro de "&%s&" apresentou resultado",'
                               '"Os parâmetros de "&%s&" apresentaram resultados")&" fora da tolerância estabelecida.")'
                         % (rc(X, 0), rc(X, 0), rc(X, 3), rc(X, 3)))
        topicos = ('IF(AND(bIncFrase="Sim",bFrase<>""),"•  "&bFrase&CHAR(10),"")'
                   '&IF(AND(bIncAcomp="Sim",gAcomp<>""),"•  "&gAcomp&CHAR(10),"")')
        for i in range(1, 7):
            topicos += '&IF(TRIM(bCom%d&"")<>"","•  "&TRIM(bCom%d&"")&CHAR(10),"")' % (i, i)
        ws.write_formula(X, 7, '=' + topicos)
        ws.write_formula(X, 8, '=IF(%s="","Sem observações.",LEFT(%s,LEN(%s)-1))' % (rc(X, 7), rc(X, 7), rc(X, 7)))
        self.nome('gNok', ws, X, 0)
        self.nome('gOk', ws, X, 1)
        self.nome('gAval', ws, X, 2)
        self.nome('gNokDet', ws, X, 4)
        self.nome('gOkLista', ws, X, 5)
        self.nome('gTopicos', ws, X, 7)
        self.nome('gComentarios', ws, X, 8)
        # textos auxiliares
        ws.write(X + 2, 0, 'Referência')
        ws.write_formula(X + 2, 1, '=IF(bRef=bLabB,bLabB,bLabA)')
        ws.write(X + 3, 0, 'Avaliado')
        ws.write_formula(X + 3, 1, '=IF(bRef=bLabB,bLabA,bLabB)')
        ws.write(X + 4, 0, 'Data report')
        ws.write_formula(X + 4, 1, '=IF(ISNUMBER(bData),TEXT(DAY(bData),"00")&"/"&TEXT(MONTH(bData),"00")&"/"&YEAR(bData),"")')
        ws.write(X + 5, 0, 'Legenda dif.')
        ws.write_formula(X + 5, 1, '="Diferença ("&%s&" − "&%s&")"' % (rc(X + 3, 1), rc(X + 2, 1)))
        self.nome('gRefNome', ws, X + 2, 1)
        self.nome('gAvalNome', ws, X + 3, 1)
        self.nome('gDataTxt', ws, X + 4, 1)
        self.nome('gLegDif', ws, X + 5, 1)
        # 1a janela completa de 2 h depois do inicio (grade da aba Configuracoes) e Nº de janelas de Germano
        ws.write(X + 8, 0, 'Janela Ubu')
        ws.write_formula(X + 8, 1, '=IF(ISNUMBER(bIniP),cfgGradeU+ROUNDUP((bIniP-cfgGradeU)*12-0.000001,0)/12,"")')
        ws.write(X + 9, 0, 'Janela Germano')
        ws.write_formula(X + 9, 1, '=IF(ISNUMBER(bIniG),cfgGradeG+ROUNDUP((bIniG-cfgGradeG)*12-0.000001,0)/12,"")')
        ws.write(X + 10, 0, 'Janelas Germano')
        ws.write_formula(X + 10, 1, '=IF(ISNUMBER(%s),MIN(bNHor,%s),bNHor)' % (rc(X + 19, 1), rc(X + 19, 1)))
        # bi-horarios de Germano (EB06): define o Nº de horarios do batch (Ubu tem os mesmos)
        ws.write(X + 19, 0, 'Bi-horários Germano')
        ws.write_formula(X + 19, 1, '=IF(AND(ISNUMBER(%s),ISNUMBER(bFimG)),MIN(%d,MAX(1,ROUNDUP((bFimG-%s)*12-0.000001,0))),"")'
                         % (rc(X + 9, 1), NMAX, rc(X + 9, 1)))
        self.nome('gNGer', ws, X + 19, 1)
        ws.write(X + 11, 0, 'Nº horários (término Ubu)')
        # o fim tambem anda o atraso do nivel dos tanques (inicio no processo - chegada)
        ws.write_formula(X + 11, 1, '=IF(AND(ISNUMBER(%s),ISNUMBER(bFimU)),MIN(%d,MAX(1,ROUNDUP((bFimU+IF(ISNUMBER(bIniU),'
                                    'bIniP-bIniU,0)-%s)*12-0.000001,0))),"")' % (rc(X + 8, 1), NMAX, rc(X + 8, 1)))
        ws.write(X + 12, 0, 'Atraso pelo nível (h)')
        ws.write_formula(X + 12, 1, '=IF(bNivel="",0,IFERROR(INDEX(cfgAtrasoH,MATCH(bNivel,cfgAtrasoN,0)),0))')
        self.nome('gAtrasoNivel', ws, X + 12, 1)

        def dt(x):
            return ('TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&" "&TEXT(HOUR(%s),"00")&":"&TEXT(MINUTE(%s),"00")'
                    % (x, x, x, x))
        partes = ['IF(ISNUMBER(bIniU),"chegada em Ubu "&%s&" · ","")' % dt('bIniU'),
                  'IF(bNivel<>"","nível dos tanques "&LOWER(bNivel)&" · ","")',
                  'IF(ISNUMBER(gJanU),"comparação a partir de "&%s&" · ","")' % dt('gJanU+cfgMinAntes/1440')]
        for k, f in enumerate(partes):
            ws.write(X + 13 + k, 0, 'Acomp. %d' % (k + 1))
            ws.write_formula(X + 13 + k, 1, '=' + f)
        junta = '&'.join(rc(X + 13 + k, 1) for k in range(len(partes)))
        ws.write(X + 17, 0, 'Acompanhamento')
        ws.write_formula(X + 17, 1, '=IF((%s)="","",UPPER(LEFT(%s,1))&MID(LEFT(%s,LEN(%s)-3),2,500)&".")'
                         % (junta, junta, junta, junta))
        self.nome('gAcomp', ws, X + 17, 1)
        ws.write(X + 18, 0, 'Resumo horários')
        ws.write_formula(X + 18, 1, '=IF(ISNUMBER(gJanU),"→ Ubu: "&bNHor&" horários a partir de "&%s&IF(ISNUMBER(gJanG),'
                                    '"   ·   Germano: "&gNG&" a partir de "&%s,"   ·   Germano: mesmos horários"),'
                                    '"→ Cole as fases do evento (passo ②) para calcular os horários.")'
                         % (dt('gJanU+cfgMinAntes/1440'), dt('gJanG+cfgMinG/1440')))
        self.nome('gResumoHor', ws, X + 18, 1)
        self.nome('gJanU', ws, X + 8, 1)
        self.nome('gJanG', ws, X + 9, 1)
        self.nome('gNG', ws, X + 10, 1)
        self.nome('gNHorFim', ws, X + 11, 1)
        ws.write_formula(X + 6, 0, '=bLabA&IF(bRef=bLabA," (ref.)","")')
        ws.write_formula(X + 6, 1, '=bLabB&IF(bRef=bLabB," (ref.)","")')
        self.nome('gSerieA', ws, X + 6, 0)
        self.nome('gSerieB', ws, X + 6, 1)

        # ---- series dos graficos: bloco de G_BLOCO colunas por parametro, linhas 1..NMAX+1
        for p, pr in enumerate(PARAMS):
            c0 = p * G_BLOCO
            st = lambda key: rc(G_ROW_ST + p, G_COL[key], True, True)
            for i, h in enumerate(['Rótulo', 'A', 'B', 'Dif', 'Dif média', '+Tol', '−Tol', 'Linha', 'k']):
                ws.write(0, c0 + i, '%s %s' % (pr[0], h))
            for kk in range(1, NMAX + 2):
                r = kk
                n = st('NP')
                lin = rc(r, c0 + 7)
                ws.write(r, c0 + 8, kk)
                ws.write_formula(r, c0 + 7, '=IF(%d<=%s,MATCH(%d,%s,0),"")' % (kk, n, kk, bcol(B_COL_R1 + p)))
                hora = 'INDEX(%s,%s)' % (bcol(B_COL_TU), lin)
                ws.write_formula(r, c0, '=IF(%d<=%s,TEXT(HOUR(%s),"00")&":"&TEXT(MINUTE(%s),"00")&CHAR(10)&'
                                        'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00"),IF(%d=%s+1,"Média",""))'
                                 % (kk, n, hora, hora, hora, hora, kk, n))
                for j, (cc, media) in enumerate([(B_COL_V1 + 2 * p, 'MA'), (B_COL_V1 + 2 * p + 1, 'MB')]):
                    ws.write_formula(r, c0 + 1 + j, '=IF(%d<=%s,INDEX(%s,%s),IF(AND(%d=%s+1,ISNUMBER(%s)),%s,NA()))'
                                     % (kk, n, bcol(cc), lin, kk, n, st(media), st(media)))
                ws.write_formula(r, c0 + 3, '=IF(%d<=%s,INDEX(%s,%s),NA())' % (kk, n, bcol(B_COL_D1 + p), lin))
                ws.write_formula(r, c0 + 4, '=IF(AND(%d=%s+1,ISNUMBER(%s)),%s,NA())' % (kk, n, st('D'), st('D')))
                ws.write_formula(r, c0 + 5, '=IF(%d<=%s+1,%s,NA())' % (kk, n, st('TOL')))
                ws.write_formula(r, c0 + 6, '=IF(%d<=%s+1,-%s,NA())' % (kk, n, st('TOL')))
        ws.write(NMAX + 3, 0, 'Nao editar: series dos graficos do Painel (o VBA ajusta o tamanho das series).')

    # ------------------------------------------------------------ Painel
    # colunas (pixels): A B C D E F G H I J K L M N O
    P_COLS = [10, 84, 70, 70, 92, 62, 62, 12, 122, 12, 246, 12, 120, 196, 10]
    P_R0 = 5            # 1a linha do report (0-based) - acima ficam os botoes

    def aba_painel(self):
        ws = self.ws_pai
        for c, px in enumerate(self.P_COLS):
            ws.set_column_pixels(c, c, px)
        ws.set_column_pixels(len(self.P_COLS), len(self.P_COLS), 20)
        ws.set_default_row(15)
        bgw = self.f(bg_color=BRANCO)
        # ---- barra de ferramentas (fora da imagem)
        tb = self.f(bg_color=BG)
        for r in range(0, self.P_R0 - 1):
            for c in range(0, len(self.P_COLS) + 1):
                ws.write_blank(r, c, None, tb)
        ws.set_row(0, 8)
        ws.set_row(1, 14)
        ws.set_row(2, 26)
        ws.set_row(3, 10)
        ws.write(1, 1, 'BATCH EM EXIBIÇÃO', self.f(bold=True, font_size=7.5, font_color=MUTED, bg_color=BG))
        ws.merge_range(2, 1, 2, 2, '', self.fi(align='center', bold=True, font_size=11, border_color=LINE))
        ws.write_blank(2, 1, None, self.fi(align='center', bold=True, font_size=11, border_color=LINE,
                                                         num_format='"Batch "0'))
        self.nome('pSel', ws, 2, 1)
        ws.data_validation(2, 1, 2, 1, {'validate': 'list', 'source': '=lstBatches',
                                        'error_type': 'information',
                                        'error_message': 'Batch não encontrado no histórico.'})
        ws.write(2, 3, '◂ escolha um batch salvo', self.f(font_size=8, italic=True, font_color=MUTED, bg_color=BG))
        self.botao(ws, 1, 4, '✎ Editar batch', 'IrBatch', 118, 30, 's', x=40, y=4)
        self.botao(ws, 1, 8, 'Atualizar painel', 'AtualizarPainel', 120, 30, 's', x=6, y=4)
        self.botao(ws, 1, 10, 'Copiar imagem', 'CopiarImagemPainel', 116, 30, 'p', x=10, y=4)
        self.botao(ws, 1, 10, 'Copiar resumo', 'CopiarResumo', 116, 30, 'n', x=132, y=4)

        R = self.P_R0
        last = len(self.P_COLS) - 2     # N
        # fundo branco da area do report
        # ---- cabecalho
        ws.set_row(R, 38)
        hdr = self.f(bg_color=NAVY, font_color=BRANCO, bold=True, font_size=14)
        ws.write_blank(R, 0, None, self.f(bg_color=NAVY))
        ws.merge_range(R, 1, R, 5, '', hdr)
        ws.write_formula(R, 1, '=IF(TRIM(bTit&"")="","COMPARATIVO DE MEDIÇÕES DE QUALIDADE",UPPER(bTit))',
                         self.f(bg_color=NAVY, font_color=BRANCO, bold=True, font_size=14, indent=1))
        ws.merge_range(R, 6, R, 11, '', hdr)
        ws.write_formula(R, 6, '="|    "&UPPER(bSub)',
                         self.f(bg_color=NAVY, font_color='#BFC8D6', bold=True, font_size=11))
        ws.merge_range(R, 12, R, 13, '', self.f(bg_color=NAVY))
        ws.write_blank(R, 14, None, self.f(bg_color=NAVY))
        ws.insert_image(R, 13, LOGO, {'x_scale': 0.26, 'y_scale': 0.26, 'x_offset': 70, 'y_offset': 12,
                                      'object_position': 1, 'description': 'Samarco'})
        ws.set_row(R + 1, 22)
        sub = self.f(bg_color=NAVY2, font_color=BRANCO, bold=True, font_size=10)
        ws.write_blank(R + 1, 0, None, self.f(bg_color=NAVY2))
        ws.merge_range(R + 1, 1, R + 1, 5, 'REPORT GERENCIAL', self.f(bg_color=NAVY2, font_color=BRANCO, bold=True,
                                                                    font_size=10, indent=1))
        ws.merge_range(R + 1, 6, R + 1, 13, '', sub)
        ws.write_formula(R + 1, 6, '="BATCH "&bId&"   |   "&gDataTxt&IF(bMin<>"","   |   Mineroduto "&bMin,"")',
                         self.f(bg_color=NAVY2, font_color='#E3E9F2', bold=True, font_size=9.5, align='right',
                                indent=1))
        ws.write_blank(R + 1, 14, None, self.f(bg_color=NAVY2))
        ws.set_row(R + 2, 10)
        for c in range(0, 15):
            ws.write_blank(R + 2, c, None, bgw)

        # ---- quadros do topo
        r = R + 3
        ws.set_row(r, 20)
        secf = self.f(bold=True, font_size=8, font_color=INK, bg_color=SEC, top=1, bottom=1, left=1, right=1,
                      border_color=LINE, indent=1)
        ws.merge_range(r, 1, r, 6, '', secf)
        ws.write_formula(r, 1, '=UPPER("Farol — Batch "&bId&IF(bMin<>""," · Mineroduto "&bMin,"")&" · ref. "&gRefNome)',
                         secf)
        ws.merge_range(r, 8, r, 13, 'COMENTÁRIOS', secf)
        # farol: cabecalho
        rh = r + 1
        ws.set_row(rh, 28)
        fth = self.f(bold=True, font_size=7.5, font_color=INK2, bg_color=TH, align='center', border=1,
                     border_color=LINE, text_wrap=True)
        ws.write_blank(rh, 1, None, fth)
        ws.write_formula(rh, 2, '=UPPER(bLabA)&IF(bRef=bLabA," ★","")&CHAR(10)&"(MÉDIA)"', fth)
        ws.write_formula(rh, 3, '=UPPER(bLabB)&IF(bRef=bLabB," ★","")&CHAR(10)&"(MÉDIA)"', fth)
        ws.write_formula(rh, 4, '="DIFERENÇA"&CHAR(10)&"("&UPPER(gAvalNome)&"−"&UPPER(gRefNome)&")"', fth)
        ws.write(rh, 5, 'TOLER.\n(±)', fth)
        ws.write(rh, 6, 'STATUS', fth)
        ftd = dict(font_size=10, align='center', border=1, border_color='#D5D9E2')
        for p, pr in enumerate(PARAMS):
            rr = rh + 1 + p
            ws.set_row(rr, 21)
            st = self.st[pr[0]]
            ws.write(rr, 1, pr[1], self.f(bold=True, font_size=10, border=1, border_color='#D5D9E2', indent=1))
            ws.write_formula(rr, 2, '=%s' % st['MA'], self.f(num_format=fmt_num(pr[4], True), **ftd))
            ws.write_formula(rr, 3, '=%s' % st['MB'], self.f(num_format=fmt_num(pr[4], True), **ftd))
            ws.write_formula(rr, 4, '=IF(ISNUMBER(%s),%s,"—")' % (st['D'], st['DR']),
                             self.f(bold=True, num_format=fmt_sinal(pr[4], True), **ftd))
            ws.write_formula(rr, 5, '=%s' % st['TOL'], self.f(num_format=fmt_num(pr[4], True), **ftd))
            ws.write_formula(rr, 6, '=%s' % st['ST'], self.f(bold=True, **ftd))
            self.nome('pSt_%s' % pr[0], ws, rr, 6)
        f1, f2 = rh + 1, rh + NP
        ws.conditional_format(f1, 6, f2, 6, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                             'format': self.wb.add_format({'font_color': OK, 'bold': True})})
        ws.conditional_format(f1, 6, f2, 6, {'type': 'cell', 'criteria': '==', 'value': '"NOK"',
                                             'format': self.wb.add_format({'font_color': NOK, 'bold': True,
                                                                           'bg_color': NOK_ROW})})
        ws.conditional_format(f1, 1, f2, 5, {'type': 'formula', 'criteria': '=$G%d="NOK"' % (f1 + 1),
                                             'format': self.wb.add_format({'bg_color': NOK_ROW})})
        rl = f2 + 1
        ws.set_row(rl, 16)
        fleg = self.f(font_size=7, font_color=MUTED, bottom=1, left=1, right=1, border_color=LINE)
        ws.merge_range(rl, 1, rl, 3, 'OK — DENTRO DA TOLERÂNCIA', self.f(font_size=7, font_color=MUTED, bottom=1,
                                                                      left=1, border_color=LINE, indent=1))
        ws.merge_range(rl, 4, rl, 6, 'NOK — FORA DA TOLERÂNCIA', self.f(font_size=7, font_color=MUTED, bottom=1,
                                                                     right=1, border_color=LINE, align='right',
                                                                     indent=1))
        # comentarios
        fcm = self.f(font_size=10, font_color=INK2, text_wrap=True, valign='top', border=1, border_color=LINE, indent=1)
        ws.merge_range(rh, 8, rl, 13, '', fcm)
        ws.write_formula(rh, 8, '=gComentarios', fcm)
        ws.set_row(rl + 1, 10)

        # ---- graficos: 2 por linha (P | SiO2), (PPC | <325#), S.E. inteiro
        self.graficos = []
        r = rl + 2
        linhas_graf = [('P', 'SiO2'), ('PPC', '325'), ('SE',)]
        H_GRAF = 13          # linhas de 15 pt
        for lin in linhas_graf:
            if len(lin) == 2:
                caixas = [(lin[0], 1, 7, 8), (lin[1], 10, 12, 13)]
            else:
                caixas = [(lin[0], 1, 12, 13)]
            ws.set_row(r, 19)
            for k, c1, c2, cb in caixas:
                fh = self.f(bold=True, font_size=8, font_color=INK, bg_color='#F8F9FB', top=1, left=1,
                            bottom=1, border_color=LINE, indent=1)
                ws.merge_range(r, c1, r, c2, '', fh)
                ws.write_formula(r, c1, '=UPPER(tit_%s)' % k, fh)
                ws.write_formula(r, cb, '=IF(pSt_%s="—","—  ","●  "&pSt_%s&"  ")' % (k, k),
                                 self.f(bold=True, font_size=8, font_color='#888888', bg_color='#F8F9FB', top=1,
                                        right=1, bottom=1, border_color=LINE, align='right'))
                self.badge(ws, r, cb, k)
            for rr in range(r + 1, r + 1 + H_GRAF):
                ws.set_row(rr, 15)
            for k, c1, c2, cb in caixas:
                largura = sum(self.P_COLS[c1:cb + 1])
                self.grafico(ws, r + 1, c1, k, largura, int(H_GRAF * 20) - 2)
            r = r + 1 + H_GRAF
            ws.set_row(r, 10)
            r += 1
        # rodape
        ws.set_row(r, 16)
        ws.merge_range(r, 1, r, 13, '', self.f(font_size=7, font_color=MUTED, align='center'))
        ws.write_formula(r, 1, '="Diferença = "&gAvalNome&" − "&gRefNome&" (referência: "&gRefNome&")   ·   faixa '
                               'verde tracejada = tolerância   ·   ◆ = diferença das médias"',
                         self.f(font_size=7, font_color=MUTED, align='center'))
        self.p_fim = r
        self.nome('pArea', ws, R, 0, r, 14)
        # fundo branco do report (celulas sem formato)
        ws.print_area(R, 0, r, 14)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.3, 0.3, 0.3, 0.3)
        ws.center_horizontally()
        ws.protect('', {'format_columns': True, 'format_rows': True})

    def badge(self, ws, r, c, k):
        st = 'pSt_%s' % k
        cel = rc(r, c, False, False)
        ws.conditional_format(r, c, r, c, {'type': 'formula', 'criteria': '=%s="OK"' % st,
                                           'format': self.wb.add_format({'font_color': OK})})
        ws.conditional_format(r, c, r, c, {'type': 'formula', 'criteria': '=%s="NOK"' % st,
                                           'format': self.wb.add_format({'font_color': NOK})})

    def n_pares(self, k):
        b = self.atual
        if not b:
            return 0
        return sum(1 for row in b['rows'] if ('%s_G' % k) in row and ('%s_U' % k) in row)

    def grafico(self, ws, r, c, k, largura, altura):
        p = PK.index(k)
        pr = PARAMS[p]
        g = quote(self.ws_gra)
        c0 = p * G_BLOCO
        n = self.n_pares(k)
        ult = n + 1                     # linha da "Média"

        def rng(cc):
            return '=%s!%s:%s' % (g, rc(1, c0 + cc, True, True), rc(ult, c0 + cc, True, True))
        cats = rng(0)
        col = self.wb.add_chart({'type': 'column'})
        pontos_a = [None] * n + [{'fill': {'color': COR_A_MEDIA}, 'border': {'none': True}}]
        pontos_b = [None] * n + [{'fill': {'color': COR_B_MEDIA}, 'border': {'none': True}}]
        col.add_series({'name': '=%s!%s' % (g, rc(G_ROW_X + 6, 0, True, True)), 'categories': cats,
                        'values': rng(1), 'fill': {'color': COR_A}, 'border': {'none': True}, 'gap': 55,
                        'overlap': -4, 'points': pontos_a})
        col.add_series({'name': '=%s!%s' % (g, rc(G_ROW_X + 6, 1, True, True)), 'categories': cats,
                        'values': rng(2), 'fill': {'color': COR_B}, 'border': {'none': True}, 'points': pontos_b})
        lin = self.wb.add_chart({'type': 'line'})
        fs = fmt_sinal(pr[4], True)
        lin.add_series({'name': '=%s!%s' % (g, rc(G_ROW_X + 5, 1, True, True)), 'categories': cats,
                        'values': rng(3), 'y2_axis': True, 'smooth': True,
                        'line': {'color': COR_DIF, 'width': 2.5},
                        'marker': {'type': 'circle', 'size': 6, 'fill': {'color': COR_DIF},
                                   'border': {'color': BRANCO, 'width': 1}},
                        'data_labels': {'value': True, 'position': 'above', 'num_format': fs,
                                        'font': {'name': FONTE, 'bold': True, 'size': 8, 'color': COR_DIF_TXT}}})
        lin.add_series({'name': 'Diferença das médias', 'categories': cats, 'values': rng(4), 'y2_axis': True,
                        'line': {'none': True},
                        'marker': {'type': 'diamond', 'size': 9, 'fill': {'color': NAVY},
                                   'border': {'color': BRANCO, 'width': 1}},
                        'data_labels': {'value': True, 'position': 'above', 'num_format': fs,
                                        'font': {'name': FONTE, 'bold': True, 'size': 8, 'color': NAVY}}})
        for j, cc in enumerate((5, 6)):
            lin.add_series({'name': 'Tolerância (±)', 'categories': cats, 'values': rng(cc), 'y2_axis': True,
                            'line': {'color': OK, 'width': 1, 'dash_type': 'dash', 'transparency': 35},
                            'marker': {'type': 'none'}})
        col.combine(lin)
        eixo_fmt = fmt_num(min(pr[4], 2), True) if k != 'SE' else '#,##0'
        col.set_x_axis({'num_font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                        'line': {'color': LINE}, 'major_tick_mark': 'none'})
        col.set_y_axis({'num_font': {'name': FONTE, 'size': 7.5, 'color': '#5A6A82'}, 'num_format': eixo_fmt,
                        'major_gridlines': {'visible': True, 'line': {'color': '#EDEFF3', 'width': 0.75}},
                        'line': {'none': True}, 'major_tick_mark': 'none'})
        lin.set_y2_axis({'num_font': {'name': FONTE, 'size': 7.5, 'color': '#A06D0A'},
                         'num_format': fmt_num(min(pr[4], 3), True) if k != 'SE' else '#,##0',
                         'line': {'none': True}, 'major_tick_mark': 'none'})
        col.set_legend({'position': 'bottom', 'font': {'name': FONTE, 'size': 8, 'bold': True, 'color': INK},
                        'delete_series': [3, 5]})
        col.set_chartarea({'border': {'color': LINE, 'width': 0.75}, 'fill': {'color': BRANCO}})
        col.set_plotarea({'fill': {'none': True}, 'border': {'none': True},
                          'layout': {'x': 0.06, 'y': 0.08, 'width': 0.86, 'height': 0.62}})
        col.set_size({'width': largura, 'height': altura})
        col.show_blanks_as('gap')
        col.show_na_as_empty_cell()
        ws.insert_chart(r, c, col, {'x_offset': 0, 'y_offset': 0, 'object_position': 1,
                                    'description': 'Gráfico %s' % k})
        self.graficos.append(k)

    # ------------------------------------------------------------ Enviar
    def aba_enviar(self):
        ws = self.ws_env
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 92)
        ws.set_column(2, 2, 2)
        ws.set_column(3, 3, 34)
        self.faixa(ws, 1, 1, 3, 'ENVIAR REPORTE')
        self.secao(ws, 4, 1, 1, 'RESUMO PARA O CORPO DO E-MAIL')
        coments = ''.join('&IF(TRIM(bCom%d&"")<>"",CHAR(10)&"•  "&TRIM(bCom%d&""),"")' % (i, i) for i in range(1, 7))
        f = ('="Comparativo "&bLabA&" × "&bLabB&" — Batch "&bId&IF(bMin<>""," (mineroduto "&bMin&")","")'
             '&" · "&gDataTxt&CHAR(10)'
             '&IF(gAval=0,"Sem pares "&bLabA&"/"&bLabB&" para avaliar.",'
             'IF(gNok=0,"Todos os parâmetros dentro da tolerância.",'
             'gNok&" de "&gAval&" fora da tolerância: "&gNokDet&"."))' + coments)
        fbox = self.f(font_size=10.5, font_color=INK2, text_wrap=True, valign='top', bg_color='#F8F9FB', border=1,
                      border_color=LINE)
        ws.merge_range(5, 1, 14, 1, '', fbox)
        ws.write_formula(5, 1, f, fbox)
        self.nome('eResumo', ws, 5, 1)
        ws.write(16, 1, 'ASSUNTO DO E-MAIL', self.f(bold=True, font_size=8, font_color=MUTED))
        ws.write_formula(17, 1, '="Comparativo Germano × Ubu — Batch "&bId&" — "&gDataTxt',
                         self.f(bold=True, font_color=INK2))
        self.nome('eAssunto', ws, 17, 1)
        self.botao(ws, 5, 3, 'Copiar resumo', 'CopiarResumo', 220, 40, 'n', x=6)
        self.botao(ws, 8, 3, 'Copiar painel (imagem)', 'CopiarImagemPainel', 220, 40, 'p', x=6)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Histórico + BD_Leituras
    def stats_batch(self, b):
        res = {}
        for p, pr in enumerate(PARAMS):
            k = pr[0]
            a = [r['%s_G' % k] for r in b['rows'] if ('%s_G' % k) in r]
            u = [r['%s_U' % k] for r in b['rows'] if ('%s_U' % k) in r]
            ma = sum(a) / len(a) if a else None
            mb = sum(u) / len(u) if u else None
            d = (mb - ma) if (ma is not None and mb is not None) else None
            st = '—' if d is None else ('OK' if abs(round(d, pr[4])) <= pr[5] + 1e-9 else 'NOK')
            res[k] = (ma, mb, d, st)
        return res

    @staticmethod
    def comentarios(b):
        return ['Ponto de amostragem do PF recebido na descarga da filtragem.']

    def aba_historico(self):
        ws = self.ws_his
        ws.set_column(0, 0, 2)
        ws.set_column(1, len(H_COLS), 13)
        ws.set_column(1, 1, 9)
        ws.set_column(10, 11, 30)
        ws.set_column(18, 23, 30)
        self.faixa(ws, 1, 1, 14, 'HISTÓRICO DE BATCHES',
                   'Um batch por linha (tabela tbl_Batches). As leituras ficam na aba BD_Leituras (tbl_Leituras). '
                   'As duas tabelas alimentam o Power BI.')
        self.botao(ws, 3, 1, 'Abrir batch selecionado', 'AbrirBatchSelecionado', 200, 30, 'p', y=4)
        self.botao(ws, 3, 3, 'Excluir batch selecionado', 'ExcluirBatchSelecionado', 200, 30, 'd', x=60, y=4)
        self.botao(ws, 3, 7, '← Painel', 'IrPainel', 120, 30, 's', x=20, y=4)
        ws.write(4, 10, 'Clique em qualquer célula da linha do batch e use os botões.',
                 self.f(font_size=8.5, italic=True, font_color=MUTED))
        ws.set_row(3, 22)
        ws.set_row(4, 18)
        # historico salvo pelo laboratorio (planilha de 14/10/2026), mapeado pelo nome das colunas
        with open(HIST_USUARIO, encoding='utf-8') as fh:
            hu = json.load(fh)

        def val(v):
            if isinstance(v, dict) and 'dt' in v:
                return dtm.datetime.fromisoformat(v['dt'])
            return v
        dados = [[val(b.get(h)) for h in H_COLS] for b in hu['batches']]
        lmap = [hu['colunas_leituras'].index(h) if h in hu['colunas_leituras'] else None for h in L_COLS]
        leituras = [[val(r[k]) if k is not None else None for k in lmap] for r in hu['leituras']]
        self.hist_dados = dados
        n = len(dados)
        fdate = self.wb.add_format({'num_format': 'dd/mm/yyyy', 'font_name': FONTE, 'font_size': 9.5,
                                    'align': 'center'})
        fdt = self.wb.add_format({'num_format': 'dd/mm/yyyy hh:mm', 'font_name': FONTE, 'font_size': 9.5,
                                  'align': 'center'})
        fnum = self.wb.add_format({'font_name': FONTE, 'font_size': 9.5, 'align': 'center', 'num_format': '0.000'})
        ftxt = self.wb.add_format({'font_name': FONTE, 'font_size': 9.5})
        cols = []
        for h in H_COLS:
            fmt = ftxt
            if h in ('Data do report',):
                fmt = fdate
            elif h in ('Chegada Ubu (UBU3)', 'Término Ubu (UBU3)', 'Início processo (manual)', 'Início Germano (EB06)',
                       'Término Germano (EB06)', 'Início no processo', 'Salvo em'):
                fmt = fdt
            elif 'média' in h or 'diferença' in h or h.startswith('Tol') or h.startswith('PFUF ') and h != 'PFUF data':
                fmt = fnum
            elif h == 'PFUF data':
                fmt = fdate
            cols.append({'header': h, 'format': fmt})
        ws.add_table(H_ROW_HDR, 1, H_ROW_HDR + max(n, 1), len(H_COLS),
                     {'name': 'tbl_Batches', 'columns': cols, 'data': dados, 'style': 'Table Style Light 9'})
        ws.set_row(H_ROW_HDR, 30)
        for i, h in enumerate(H_COLS):
            if h.endswith('status'):
                c = 1 + i
                ws.conditional_format(H_ROW_HDR + 1, c, H_ROW_HDR + 2000, c, {
                    'type': 'cell', 'criteria': '==', 'value': '"NOK"',
                    'format': self.wb.add_format({'font_color': NOK, 'bold': True, 'bg_color': NOK_BG})})
                ws.conditional_format(H_ROW_HDR + 1, c, H_ROW_HDR + 2000, c, {
                    'type': 'cell', 'criteria': '==', 'value': '"OK"',
                    'format': self.wb.add_format({'font_color': OK, 'bold': True, 'bg_color': OK_BG})})
        ws.freeze_panes(H_ROW_HDR + 1, 2)
        # lista de batches (validacao do Painel): coluna Batch da tabela
        self.wb.define_name('lstBatches', "=OFFSET('Histórico'!$B$%d,0,0,MAX(1,COUNT('Histórico'!$B$%d:$B$5000)),1)"
                            % (H_ROW_HDR + 2, H_ROW_HDR + 2))
        # BD_Leituras
        wl = self.ws_bdl
        wl.set_column(0, 1, 10)
        wl.set_column(2, 3, 18)
        wl.set_column(4, 4, 12)
        wl.set_column(5, 9, 12)
        lcols = []
        for h in L_COLS:
            fmt = ftxt
            if h.startswith('Data/hora'):
                fmt = fdt
            elif h in ('Germano', 'Ubu', 'Diferença', 'Tolerância'):
                fmt = fnum
            lcols.append({'header': h, 'format': fmt})
        wl.add_table(0, 0, max(len(leituras), 1), len(L_COLS) - 1,
                     {'name': 'tbl_Leituras', 'columns': lcols, 'data': leituras, 'style': 'Table Style Light 1'})
        wl.freeze_panes(1, 0)

    # ------------------------------------------------------------ Dados_MES (oculta)
    def aba_dados_mes(self):
        ws = self.ws_mes
        ws.set_column(0, 0, 20)
        ws.set_column(1, NP, 14)
        ws.set_column(7, 8, 30)
        ws.write(0, 0, 'Consultas ao MES (Aspen IP.21) - não editar', self.f(bold=True, font_size=13, font_color=NAVY2))
        fdt = self.f(num_format='dd/mm/yyyy hh:mm:ss', bg_color=SEC, align='left')
        rot = self.f(bold=True)
        for j, (lab, nm) in enumerate([('Germano', 'G'), ('Ubu', 'U')]):
            r = 2 + j * 3
            ws.write(r, 0, '%s - início' % lab, rot)
            ws.write_blank(r, 1, None, fdt)
            ws.write(r + 1, 0, '%s - fim' % lab, rot)
            ws.write_blank(r + 1, 1, None, fdt)
            self.nome('mesIni%s' % nm, ws, r, 1)
            self.nome('mesFim%s' % nm, ws, r + 1, 1)
        ws.write(2, 3, 'Prefixo Aspen', rot)
        ws.write_blank(2, 4, None)
        self.nome('mesPrefixo', ws, 2, 4)
        ws.write(3, 3, 'Última consulta', rot)
        ws.write_blank(3, 4, None)
        self.nome('mesTempo', ws, 3, 4)
        ws.write(4, 3, 'Tags inválidas', rot)
        ws.write_blank(4, 4, None)
        self.nome('mesInvalidas', ws, 4, 4)
        ws.write(5, 3, 'Resposta Germano', rot)
        ws.write_blank(5, 4, None)
        self.nome('mesRespG', ws, 5, 4)
        ws.write(6, 3, 'Resposta Ubu', rot)
        ws.write_blank(6, 4, None)
        self.nome('mesRespU', ws, 6, 4)
        # ancoras: Germano A10 (saida A12), Ubu A40 (saida A42)
        for lab, anc, out in [('Germano', 9, 11), ('Ubu', 39, 41)]:
            ws.write(anc - 1, 0, 'Consulta %s (a fórmula é escrita pelo botão Buscar no MES)' % lab,
                     self.f(bold=True, font_color=MUTED))
            ws.write(anc, 0, 'Aguardando a primeira consulta', self.f(font_size=8))
            ws.write(anc + 1, 0, 'Data/hora', self.f(bold=True, bg_color=NAVY2, font_color=BRANCO))
            for p, pr in enumerate(PARAMS):
                ws.write(anc + 1, 1 + p, pr[1], self.f(bold=True, bg_color=NAVY2, font_color=BRANCO))
        self.botao(ws, 0, 7, 'Voltar às Configurações', 'OcultarDadosMES', 200, 28, 'n', y=2)

    # ------------------------------------------------------------ VBA
    def modulo_layout(self):
        L = ["Option Explicit", "",
             "' Gerado por build_interlab.py - posicoes fixas das abas (nao editar a mao)",
             "Public Const NP As Long = %d" % NP,
             "Public Const NMAX As Long = %d" % NMAX,
             "Public Const HORAS_SLOT As Double = 2",
             "' aba Batch (1-based)",
             "Public Const B_ROW_G0 As Long = %d" % (B_ROW_G0 + 1),
             "Public Const B_COL_N As Long = %d" % (B_COL_N + 1),
             "Public Const B_COL_TU As Long = %d" % (B_COL_TU + 1),
             "Public Const B_COL_TG As Long = %d" % (B_COL_TG + 1),
             "Public Const B_COL_V1 As Long = %d" % (B_COL_V1 + 1),
             "Public Const B_COL_D1 As Long = %d" % (B_COL_D1 + 1),
             "' aba Configuracoes: tabela de tags (1-based)",
             "Public Const C_ROW_TAG0 As Long = %d" % (C_ROW_TAG0 + 1)]
        for k, v in C_COL.items():
            L.append("Public Const C_COL_%s As Long = %d" % (k, v + 1))
        L += ["' aba _Graf",
              "Public Const G_BLOCO As Long = %d" % G_BLOCO,
              "Public Const G_ROW_ST As Long = %d" % (G_ROW_ST + 1),
              "Public Const G_COL_NP As Long = %d" % (G_COL['NP'] + 1),
              "Public Const G_COL_MA As Long = %d" % (G_COL['MA'] + 1),
              "Public Const G_COL_MB As Long = %d" % (G_COL['MB'] + 1),
              "Public Const G_COL_D As Long = %d" % (G_COL['D'] + 1),
              "Public Const G_COL_ST As Long = %d" % (G_COL['ST'] + 1),
              "' Historico: colunas da tabela tbl_Batches (ordem)",
              "Public Const H_NCOL As Long = %d" % len(H_COLS),
              "Public Const H_ROW_HDR As Long = %d" % (H_ROW_HDR + 1),
              "Public Const H_COL1 As Long = 2",
              "' coluna (na tabela) da 1a estatistica (media A do 1o parametro)",
              "Public Const H_COL_STAT1 As Long = %d" % (H_COLS.index('%s média A' % PARAMS[0][1]) + 1),
              "Public Const H_NCAMPOS As Long = %d" % len(CAMPOS),
              "Public Const L_NCOL As Long = %d" % len(L_COLS),
              "' chaves e rotulos dos parametros (ordem das colunas da aba Batch)",
              'Public Const P_CHAVES As String = "%s"' % '|'.join(PK),
              'Public Const P_DECIMAIS As String = "%s"' % '|'.join(str(p[4]) for p in PARAMS),
              "' ordem dos graficos no Painel (ChartObjects 1..5)",
              'Public Const GRAF_ORDEM As String = "%s"' % '|'.join(self.graficos),
              "' Senha de protecao das abas (vazio = sem senha)",
              'Public Const SENHA As String = ""', "",
              "' campos da aba Batch gravados em tbl_Batches (mesma ordem das 1as colunas)",
              "Public Function CAMPOS_BATCH() As String",
              '    CAMPOS_BATCH = "%s"' % '|'.join(nm for nm, _ in CAMPOS),
              "End Function", "",
              "' rotulos dos parametros (com caracteres fora do cp1252 via ChrW)",
              "Public Function P_ROTULOS() As String",
              "    P_ROTULOS = %s" % vba_str('|'.join(p[1] for p in PARAMS)),
              "End Function", ""]
        return '\n'.join(L)


# ============================================================================
def vba_str(t):
    """Literal VBA; caracteres fora do cp1252 viram ChrW()."""
    partes, buf = [], ''
    for ch in t:
        try:
            ch.encode('cp1252')
            buf += ch.replace('"', '""')
        except UnicodeEncodeError:
            if buf:
                partes.append('"%s"' % buf)
                buf = ''
            partes.append('ChrW(%d)' % ord(ch))
    if buf or not partes:
        partes.append('"%s"' % buf)
    return ' & '.join(partes)


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
    return codigo


def ler_vba(nome):
    with open(os.path.join(VBA_DIR, nome), encoding='utf-8') as fh:
        return checar_declaracoes(nome, fh.read())


def montar_vba(layout_code):
    with open(os.path.join(VBA_DIR, 'ModLayout.bas'), 'w', encoding='utf-8') as fh:
        fh.write(layout_code)
    mods = [{'name': 'ThisWorkbook', 'kind': 'workbook', 'code': ler_vba('ThisWorkbook.cls')}]
    eventos = {
        'shPainel': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                     '    On Error Resume Next\n'
                     '    If Not Intersect(Target, Nm("pSel")) Is Nothing Then SelecionarBatch\n'
                     'End Sub\n\nPrivate Sub Worksheet_Activate()\n'
                     '    On Error Resume Next\n    AjustarGraficos\nEnd Sub\n'),
        'shBatch': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                    '    On Error Resume Next\n    MarcarAlterado\n'
                    '    If gCarregando Then Exit Sub\n'
                    '    If Not Intersect(Target, Nm("bFases")) Is Nothing Then LerFases\nEnd Sub\n'),
    }
    for cn in ['shPainel', 'shBatch', 'shEnviar', 'shHist', 'shLeituras', 'shConfig', 'shDadosMES',
               'shGraf']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': eventos.get(cn, '')})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModBatch', 'ModPainel']:
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
