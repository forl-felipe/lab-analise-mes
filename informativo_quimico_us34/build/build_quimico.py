"""Gera a planilha Relatorio_Quimico_US3_US4.xlsm (Laboratório Químico - Usinas 3 e 4).

Estrutura simples, 4 abas visiveis:
  Preenchimento      - Turno Dia em cima e Turno Noite embaixo (o tecnico so preenche aqui)
  Resumo Dia / Noite - tudo do turno migrado por FORMULA + resultados quimicos do MES;
                       dois botoes: Atualizar dados do MES e Copiar imagem
  Resultados gerais  - resultados do MES de qualquer periodo de ate 24 h
Consulta ao MES (Aspen IP.21) igual a do Informativo do Laboratorio Fisico (informativo_us34).

Uso:  python3 build_quimico.py [--teste]
  --teste  acrescenta um modulo de teste automatico (usado na validacao com LibreOffice)
"""
import os
import re
import sys
import zipfile
import datetime as dtm

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as rc, xl_col_to_name as colname

from vba_project import build_vba_project

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
VBA_DIR = os.path.join(RAIZ, 'vba')
TESTE = '--teste' in sys.argv
SAIDA = os.path.join(RAIZ, 'Relatorio_Quimico_US3_US4%s.xlsm' % ('_TESTE' if TESTE else ''))

# ---------------------------------------------------------------- paleta Samarco
AZUL = '#00335A'
AZUL_MEDIO = '#004070'
AZUL_TITULO = '#003B6C'
AZUL_PROFUNDO = '#002643'
AMARELO = '#FFC000'
OURO = '#FDB913'
AZUL_CLARO = '#59C6F2'
LARANJA = '#F37021'
CINZA = '#B5BEC4'
AZUL_ACINZ = '#61889B'
BRANCO = '#FFFFFF'
# neutros do kit visual Samarco (samarco-kit-visual.html)
TEXTO = '#1D2B36'
TEXTO_SEC = '#5B6B77'
BORDA = '#DDE3E8'
FUNDO_INPUT = '#FFF8E1'      # campo a preencher
FUNDO_CLARO = '#F3F5F7'      # fundo neutro do kit
FUNDO_GRUPO = '#E8EDF1'
FUNDO_AZUL_CLARO = '#EEF8FD'
LOGO_ESCURO = os.path.join(AQUI, 'assets', 'logo_fundo_escuro.png')
VERDE_TXT, VERDE_FUNDO = '#1E7B4A', '#E2F3E8'
LARANJA_FUNDO = '#FDE6D8'
FONTE = 'Segoe UI'

# ---------------------------------------------------------------- analises quimicas
# Tags conforme o Plano Amostral do Laboratorio Fisico/Quimico de Ubu (rev. 06):
#   tag = "M" & ponto & "-" & codigo (4 digitos) & "-" & fase (HH = bi-horario) & area
# Filtragem (Pellet Feed):  US3 U03-02TP004 = M650030010    US4 U04-02TP006 = M4650060010
# Pelota Queimada:          US3 U03-07TP002 = M710050020    US4 U04-07TP015 = M4710050020
# So os resultados bi-horarios (HH). Fe e Al2O3 da Filtragem sao diarios (DD) e ficam fora do turno.
# Pelota Queimada: no Plano Amostral as linhas de CaO (0115), MgO (0006), B2 (0122), Mn (0018) e PPC (0007)
# divergem dos codigos usados em todos os outros pontos do plano e nas planilhas em uso (CaO 0006, MgO 0007,
# B2 0018, Mn 0115, PPC 0013). Foram usados os codigos consistentes; as tags podem ser trocadas em Configuracoes.
PF3, PF4, PQ3, PQ4 = 'M650030010', 'M4650060010', 'M710050020', 'M4710050020'
MV, AM = 'IP_MESVALOR', 'IP_ANALOGMAP'
G_PF = 'Filtragem - Pellet Feed'
G_LM = 'Pelota Queimada'


def _us(grupo, p3, p4, nome, un, dec, cod, tip3, tip4, fase='HHLQU'):
    """Analise com uma tag por usina (US3/US4)."""
    t3 = '%s-%s-%s' % (p3, cod, fase) if cod else ''
    t4 = '%s-%s-%s' % (p4, cod, fase) if cod else ''
    return (grupo, nome, un, dec, 'Média', t3, MV, t4, MV, tip3, tip4, 'Q', False)


def _pf(nome, un, dec, cod, tip):
    """Pellet Feed: mesmo material para as duas usinas -> uma linha so (US3/4), tag do ponto M650030010."""
    tag = '%s-%s-HHLQU' % (PF3, cod)
    return (G_PF, nome, un, dec, 'Média', tag, MV, tag, MV, tip, tip, 'Q', True)


def _lm(nome, un, dec, cod, tip3, tip4):
    return _us(G_LM, PQ3, PQ4, nome, un, dec, cod, tip3, tip4)


# (grupo, analise, unidade, decimais, agregacao, tag US3, tipo, tag US4, tipo, tipico US3, tipico US4,
#  consulta, amostra_unica)
PARAMS = [
    _pf('SiO2', '%', 2, '0004', 1.29),
    _pf('P', '%', 3, '0008', 0.040),
    _pf('MgO', '%', 2, '0007', 0.05),
    _pf('CaO', '%', 2, '0006', 0.09),
    _pf('PPC', '%', 2, '0013', 3.66),
    _lm('Fe', '%', 2, '0002', 67.30, 65.60),
    _lm('SiO2', '%', 2, '0004', 1.82, 2.90),
    _lm('Al2O3', '%', 2, '0005', 0.42, 0.44),
    _lm('CaO', '%', 2, '0006', 0.82, 1.40),
    _lm('MgO', '%', 2, '0007', 0.10, 0.11),
    _lm('B2 (CaO/SiO2)', '-', 2, '0018', 0.45, 0.48),
    _lm('Mn', '%', 3, '0115', 0.070, 0.071),
]
NPARAM = len(PARAMS)
NSLOT = 7               # janelas do turno no Resumo: 07:30 ... 19:30 (Dia) e 19:30 ... 07:30 (Noite)
USINAS = ('US3', 'US4')


# ---------------------------------------------------------------- limites por produto (farol)
# Fonte: SMIN-POP-GEA-001 rev. 12 (Limites de Processo, aprovado em 24/03/2025).
# Pelota: Processo (GPU), limites Max./Min. para DADOS HORARIOS (itens 10.4.1 e 10.4.2).
#   Onde ha revisao com seta ("67,10 -> 67,19"), vale o valor novo.
# Pellet Feed: pelo concentrado da pelota (item 10.1/10.2): SiO2 max. bi-horario do concentrador III;
#   P e PPC = limite maximo diario do batch.
# Cada item: (chave, rotulo). Na aba Limites cada chave tem duas colunas: Min. e Max.
LIM_ITENS = [('LM_FE', 'Fe'), ('LM_SIO2', 'SiO2'), ('LM_P', 'P'), ('LM_CAO', 'CaO'), ('LM_B2', 'B2'),
             ('PF_SIO2', 'SiO2'), ('PF_P', 'P'), ('PF_PPC', 'PPC')]
PF_CONC = {'CLS': {'PF_SIO2': (None, 1.36), 'PF_P': (None, 0.050), 'PF_PPC': (None, 4.30)},
           'CNS': {'PF_SIO2': (None, 1.99), 'PF_P': (None, 0.075), 'PF_PPC': (None, 4.30)},
           'CHS': {'PF_SIO2': (None, 2.50), 'PF_P': (None, 0.075), 'PF_PPC': (None, 4.30)},
           'CSP': {'PF_SIO2': (None, 5.30), 'PF_P': (None, 0.110), 'PF_PPC': (None, 4.30)}}
# (produto, concentrado, {chave: (min, max)})
LIMITES = [
    ('PDR/MX', 'CLS', {'LM_FE': (67.39, None), 'LM_SIO2': (None, 1.54), 'LM_P': (None, 0.050), 'LM_CAO': (0.70, None)}),
    ('PDR/STD', 'CNS', {'LM_FE': (67.19, None), 'LM_SIO2': (None, 2.05), 'LM_P': (None, 0.074), 'LM_CAO': (0.65, None)}),
    ('PBF/MB45', 'CHS', {'LM_FE': (65.30, None), 'LM_SIO2': (None, 3.45), 'LM_P': (None, 0.074), 'LM_B2': (None, 0.55)}),
    ('PBF/STD', 'CHS', {'LM_FE': (65.10, None), 'LM_SIO2': (None, 3.20), 'LM_P': (None, 0.074), 'LM_B2': (0.75, None)}),
    ('PBF/HB', 'CNS', {'LM_FE': (65.30, None), 'LM_SIO2': (None, 2.80), 'LM_P': (None, 0.074), 'LM_B2': (0.95, None)}),
    ('PBF/SF', 'CNS', {'LM_FE': (65.00, None), 'LM_SIO2': (None, 3.00), 'LM_P': (None, 0.072), 'LM_B2': (1.10, None)}),
    ('PBF/SA', 'CSP', {'LM_FE': (63.50, None), 'LM_SIO2': (None, 5.30), 'LM_P': (None, 0.100), 'LM_B2': (0.40, None)}),
]
LIM_ROW1 = 8          # 1a linha de produtos na aba Limites (0-based)
LIM_NLIN = 25         # linhas disponiveis (produtos novos podem ser acrescentados)
LIM_COL1 = 3          # coluna do 1o limite (D), 0-based
# analise do relatorio -> chave de limite
PARAM_LIM = {(G_PF, 'SiO2'): 'PF_SIO2', (G_PF, 'P'): 'PF_P', (G_PF, 'PPC'): 'PF_PPC',
             (G_LM, 'Fe'): 'LM_FE', (G_LM, 'SiO2'): 'LM_SIO2', (G_LM, 'P'): 'LM_P', (G_LM, 'CaO'): 'LM_CAO',
             (G_LM, 'B2 (CaO/SiO2)'): 'LM_B2'}


def lim_col(chave, qual):
    """Letra da coluna do limite (qual = 0 min, 1 max) na aba Limites."""
    i = [k for k, _ in LIM_ITENS].index(chave)
    return colname(LIM_COL1 + 2 * i + qual)


# Consulta ao MES no mesmo formato da planilha de referencia (texto literal); uma consulta so,
# com todas as analises quimicas (tipo de calculo "1" = media da janela de 2 h)
BLOCOS = [('Q', 'Análises químicas', '1', 7)]
SERVIDOR_PADRAO = 'UBU'

# Configuracoes: colunas (1-based) da tabela de tags
CFG_ROW1 = 18
CFG_COL = dict(GRUPO=2, PARAM=3, UNID=4, DEC=5, AGG=6, TAG3=7, TIPO3=8, TAG4=9, TIPO4=10,
               LIE=11, LSE=12, VMIN=13, VMAX=14, TIP3=15, TIP4=16)
# Informativo: ate 12 janelas de 2 h (periodo livre de ate 24 h); o turno usa 6
NSLOT_INF = 12
INF_COL_H1 = 5                                # 1-based: 1a janela (E)
INF_COL_RES = INF_COL_H1 + NSLOT_INF           # 1-based: Resultado; +1 Min, +2 Max, +3 LIE, +4 LSE, +5 Status
INF_L_RES = colname(INF_COL_RES - 1)           # letra da coluna Resultado
INF_L_ST = colname(INF_COL_RES + 4)            # letra da coluna Status

SH_CFG = "'Configurações'"


def cfg_ref(p, key, absolute=True):
    """Referencia a celula da tabela de tags (p 1-based)."""
    cell = rc(CFG_ROW1 + p - 2, CFG_COL[key] - 1, True, True)
    return '%s!%s' % (SH_CFG, cell)


def fmt_dec(d):
    return '0' if d == 0 else '0.' + '0' * d


def literal(texto):
    """Texto literal para formula do Excel, quebrado em pedacos de 250 caracteres
    (limite de 255 por texto), unidos com &, como faz o proprio suplemento Aspen."""
    pedacos = [texto[i:i + 250] for i in range(0, len(texto), 250)] or ['']
    return '&'.join('"%s"' % p for p in pedacos)


def configurado(p):
    """Analise entra na consulta ao MES somente com as duas tags preenchidas."""
    pr = PARAMS[p - 1]
    return bool(pr[5].strip()) and bool(pr[7].strip())


def formula_consulta(ps, calc, ancora, saida):
    """Mesma assinatura das formulas H2/BN2/BT2 da planilha de referencia (so analises com tag)."""
    ps = [p for p in ps if configurado(p)]
    tags = ','.join('%s,%s' % (PARAMS[p - 1][5], PARAMS[p - 1][7]) for p in ps)
    mapas = ','.join('%s,%s' % (PARAMS[p - 1][6], PARAMS[p - 1][8]) for p in ps)
    servs = ','.join([SERVIDOR_PADRAO] * (2 * len(ps)))
    return ('=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("time,attribute",%s,%s,%s,'
            'Dados_MES!$B$3,Dados_MES!$B$4,"2h",0,"",0,"%s",0,1560,0,0,1,1,'
            'ADDRESS(ROW(Dados_MES!%s),COLUMN(Dados_MES!%s),1,,"Dados_MES"),"Dados_MES!%s",1)'
            % (literal(tags), literal(servs), literal(mapas), calc, ancora, ancora, saida))


# ============================================================================
class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_properties({'title': 'Relatório de Passagem de Turno - Laboratório Químico - US3/US4',
                                'company': 'Samarco - Laboratório Químico'})
        self._fmts = {}

    # ----------------------------------------------------------- formatos
    def f(self, **kw):
        base = dict(font_name=FONTE, font_size=10, font_color=TEXTO, valign='vcenter')
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self._fmts:
            self._fmts[chave] = self.wb.add_format(base)
        return self._fmts[chave]

    def f_input(self, **kw):
        base = dict(bg_color=FUNDO_INPUT, border=1, border_color=BORDA, locked=False)
        base.update(kw)
        return self.f(**base)

    # ----------------------------------------------------------- blocos visuais
    def cabecalho(self, ws, ultima_col, titulo, subtitulo, col_logo_fim=2, col_ini=1, r0=0):
        """Faixa Azul Samarco alta, com logo, titulo em caixa alta e filete amarelo."""
        ws.set_row(r0, 12)
        ws.set_row(r0 + 1, 34)
        ws.set_row(r0 + 2, 24)
        ws.set_row(r0 + 3, 6)
        faixa = self.f(bg_color=AZUL)
        for r in (r0, r0 + 1, r0 + 2):
            for c in range(col_ini, ultima_col + 1):
                ws.write_blank(r, c, None, faixa)
        ws.merge_range(r0 + 1, col_logo_fim + 1, r0 + 1, ultima_col, titulo,
                       self.f(bold=True, font_size=19, font_color=BRANCO, bg_color=AZUL, indent=1, valign='bottom'))
        ws.merge_range(r0 + 2, col_logo_fim + 1, r0 + 2, ultima_col, subtitulo,
                       self.f(bold=True, font_size=11, font_color=AMARELO, bg_color=AZUL, indent=1, valign='top'))
        for c in range(col_ini, ultima_col + 1):
            ws.write_blank(r0 + 3, c, None, self.f(bg_color=AMARELO))
        ws.insert_image(r0 + 1, col_ini, LOGO_ESCURO, {'x_scale': 0.21, 'y_scale': 0.21, 'x_offset': 12,
                                                      'y_offset': 8, 'object_position': 3,
                                                      'description': 'Samarco'})

    def secao(self, ws, row, c1, c2, texto, cor=None):
        """Titulo de secao: texto Azul Titulo sobre branco, com linha azul embaixo."""
        ws.set_row(row, 22)
        ws.merge_range(row, c1, row, c2, texto,
                       self.f(bold=True, font_size=11, font_color=AZUL_TITULO, valign='bottom',
                              bottom=2, bottom_color=AZUL))

    def botao(self, ws, row, col, texto, macro, largura=200, altura=40, estilo='primario',
              x=0, y=0, tamanho=None):
        # (fundo, cor do texto, borda, tamanho)
        estilos = {
            'primario': (AZUL, BRANCO, None, 11),
            'medio': (AZUL, BRANCO, None, 11),
            'titulo': (AZUL, BRANCO, None, 11),
            'destaque': (AMARELO, AZUL_PROFUNDO, None, 11),
            'claro': (BRANCO, AZUL, AZUL, 10),
            'pequeno': (BRANCO, AZUL, AZUL, 9),
        }
        fundo, cor, borda, tam = estilos[estilo]
        ws.insert_textbox(row, col, texto, {
            'width': largura, 'height': altura, 'x_offset': x, 'y_offset': y,
            'font': {'name': FONTE, 'size': tamanho or tam, 'bold': True, 'color': cor},
            'align': {'vertical': 'middle', 'horizontal': 'center'},
            'fill': {'color': fundo},
            'line': {'color': borda, 'width': 1} if borda else {'none': True},
            'description': 'macro:' + macro,
            'object_position': 3,
        })

    @staticmethod
    def _qn(ws):
        return "'%s'" % ws.name

    def nome(self, nome, ws, r1, c1, r2=None, c2=None):
        ref = rc(r1, c1, True, True)
        if r2 is not None:
            ref += ':' + rc(r2, c2, True, True)
        self.wb.define_name(nome, '=%s!%s' % (self._qn(ws), ref))

    # ============================================================ abas
    def construir(self):
        wb = self.wb
        self.ws_pre = wb.add_worksheet('Preenchimento')
        self.ws_res_d = wb.add_worksheet('Resumo Dia')
        self.ws_res_n = wb.add_worksheet('Resumo Noite')
        self.ws_ger = wb.add_worksheet('Resultados gerais')
        self.ws_lim = wb.add_worksheet('Limites')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_mapa = wb.add_worksheet('_Mapa')
        for ws, cn in [(self.ws_pre, 'shPreenchimento'), (self.ws_res_d, 'shResumoDia'),
                       (self.ws_res_n, 'shResumoNoite'), (self.ws_ger, 'shResultados'),
                       (self.ws_lim, 'shLimites'), (self.ws_cfg, 'shConfig'), (self.ws_mes, 'shDadosMES'), (self.ws_mapa, 'shMapa')]:
            ws.set_vba_name(cn)
            ws.hide_gridlines(2)
        self.ws_pre.set_tab_color(OURO)
        self.ws_res_d.set_tab_color(AZUL)
        self.ws_res_n.set_tab_color(AZUL)
        self.ws_ger.set_tab_color(AZUL_ACINZ)
        self.ws_cfg.set_tab_color(CINZA)

        self.aba_config()
        self.aba_limites()
        self.aba_dados_mes()
        self.aba_preenchimento()
        self.aba_resumo(self.ws_res_d, 'd', 'Dia', 'Dia  07:00 – 19:00', 7)
        self.aba_resumo(self.ws_res_n, 'n', 'Noite', 'Noite  19:00 – 07:00', 19)
        self.aba_resultados()
        self.aba_mapa()
        self.ws_mes.hide()
        self.ws_cfg.hide()
        self.ws_lim.hide()
        self.ws_mapa.very_hidden() if hasattr(self.ws_mapa, 'very_hidden') else self.ws_mapa.hide()
        self.ws_pre.activate()

    def link(self, ws, r, c, destino, texto, tamanho=9):
        """Hiperlink interno (navegacao sem macro)."""
        ws.write_url(r, c, 'internal:' + destino, self.f(font_size=tamanho, bold=True, font_color=AZUL_MEDIO,
                                                          underline=1), string=texto)

    # ------------------------------------------------------------ Configurações
    def aba_config(self):
        ws = self.ws_cfg
        larg = {0: 2, 1: 16, 2: 34, 3: 9, 4: 6, 5: 13, 6: 25, 7: 15, 8: 25, 9: 15,
                10: 8, 11: 8, 12: 9, 13: 9, 14: 9, 15: 9, 16: 2, 17: 2}
        for c, w in larg.items():
            ws.set_column(c, c, w)
        ws.set_column(18, 24, 17)
        self.cabecalho(ws, 15, 'CONFIGURAÇÕES', '')
        self.secao(ws, 4, 1, 15, 'Parâmetros gerais')
        gerais = [
            ('cfgFonte', 'Fonte dos dados', 'MES', 'MES ou SIMULAÇÃO (dados fictícios, para treino)'),
            ('cfgServidor', 'Servidor do MES', 'UBU', 'Servidor do suplemento Aspen'),
            ('cfgInicioDia', 'Início do turno Dia', dtm.time(7, 0), 'O turno Noite inicia 12 h depois'),
            ('cfgTimeout', 'Espera máxima pelo MES (s)', 60, 'Tempo máximo de resposta do Aspen'),
            ('cfgFormatoData', 'Formato da data para o MES', 'Texto dd/mm/aaaa',
             'Padrão: Texto dd/mm/aaaa'),
        ]
        for i, (nm, rot, val, desc) in enumerate(gerais):
            r = 5 + i
            ws.set_row(r, 20)
            ws.merge_range(r, 1, r, 2, rot, self.f(bold=True, indent=1, bg_color=FUNDO_CLARO,
                                                   border=1, border_color=BORDA))
            if isinstance(val, dtm.time):
                ws.merge_range(r, 3, r, 5, '', self.f_input(num_format='hh:mm', align='center'))
                ws.write_datetime(r, 3, dtm.datetime.combine(dtm.date(1899, 12, 31), val),
                                  self.f_input(num_format='hh:mm', align='center'))
            else:
                ws.merge_range(r, 3, r, 5, val, self.f_input(align='center'))
            ws.merge_range(r, 6, r, 9, desc, self.f(font_size=9, font_color=TEXTO_SEC, indent=1))
            self.nome(nm, ws, r, 3)
        ws.data_validation(5, 3, 5, 3, {'validate': 'list', 'source': '=lstFonte'})
        ws.data_validation(9, 3, 9, 3, {'validate': 'list', 'source': '=lstFormatoData'})
        # botoes
        self.botao(ws, 5, 11, 'Voltar ao preenchimento', 'IrPreenchimento', 230, 30, 'primario', x=5)
        self.botao(ws, 7, 11, 'Ver dados brutos do MES', 'MostrarDadosMES', 230, 30, 'claro', x=5)

        self.secao(ws, 15, 1, 15, 'Tags do MES')
        cab = ['Amostra', 'Análise', 'Unidade', 'Dec.', 'Resultado do turno', 'Tag US3', 'Tipo US3',
               'Tag US4', 'Tipo US4', 'LIE', 'LSE', 'Válido mín.', 'Válido máx.', 'Típico US3 (simul.)',
               'Típico US4 (simul.)']
        ws.set_row(16, 32)
        for i, t in enumerate(cab):
            ws.write(16, 1 + i, t, self.f(bold=True, font_color=BRANCO, bg_color=AZUL_MEDIO, text_wrap=True,
                                          align='center', border=1, border_color=BRANCO))
        for p, pr in enumerate(PARAMS, start=1):
            r = CFG_ROW1 - 1 + p - 1
            g, nome, un, dec, agg, t3, ty3, t4, ty4, tip3, tip4 = pr[:11]
            zebra = FUNDO_CLARO if p % 2 else BRANCO
            fl = self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1)
            ws.write(r, 1, g, self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1, font_color=AZUL_ACINZ))
            ws.write(r, 2, nome, self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1, bold=True))
            ws.write(r, 3, un, self.f_input(align='center'))
            ws.write(r, 4, dec, self.f_input(align='center'))
            ws.write(r, 5, agg, self.f_input(align='center'))
            ws.write(r, 6, t3, self.f_input(font_name='Consolas', font_size=9))
            ws.write(r, 7, ty3, self.f_input(font_size=9))
            if pr[12]:
                # amostra unica (Pellet Feed): a tag US4 repete a US3
                ws.write_formula(r, 8, '=IF(%s="","",%s)' % (rc(r, 6), rc(r, 6)),
                                 self.f(font_name='Consolas', font_size=9, font_color=AZUL_ACINZ, bg_color=FUNDO_CLARO,
                                        border=1, border_color=BORDA), t4)
            else:
                ws.write(r, 8, t4, self.f_input(font_name='Consolas', font_size=9))
            ws.write(r, 9, ty4, self.f_input(font_size=9))
            for c in (10, 11):
                ws.write_blank(r, c, None, self.f_input(align='center'))
            ws.write(r, 12, 0, self.f_input(align='center'))
            ws.write(r, 13, 50000, self.f_input(align='center'))
            ws.write(r, 14, tip3, self.f_input(align='center'))
            ws.write(r, 15, tip4, self.f_input(align='center'))
            del fl
        r1, r2 = CFG_ROW1 - 1, CFG_ROW1 - 1 + NPARAM - 1
        ws.data_validation(r1, 5, r2, 5, {'validate': 'list', 'source': '=lstAgreg'})
        ws.data_validation(r1, 7, r2, 7, {'validate': 'list', 'source': '=lstTipo'})
        ws.data_validation(r1, 9, r2, 9, {'validate': 'list', 'source': '=lstTipo'})
        self.nome('lstParam', ws, r1, 2, r2, 2)
        ws.freeze_panes(17, 3)
        semtag = self.wb.add_format({'bg_color': LARANJA_FUNDO})
        ws.conditional_format(r1, 6, r2, 6, {'type': 'blanks', 'format': semtag})
        ws.write(CFG_ROW1 - 1 + NPARAM + 1, 2, 'Laranja: análise sem tag.',
                 self.f(font_size=9, bold=True, font_color=LARANJA))
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.print_area(1, 1, CFG_ROW1 + NPARAM, 15)

        # listas
        listas = [('lstTurno', 'Turnos', ['Dia (07h às 19h)', 'Noite (19h às 07h)']),
                  ('lstTurma', 'Letras', ['A', 'B', 'C', 'D']),
                  ('lstAr', 'Sistema de ar', ['Compressor', 'Nitrogênio']),
                  ('lstLigado', 'Status', ['Ligado', 'Desligado']),
                  ('lstCond', 'Condição', ['Bom', 'Regular', 'Ruim']),
                  ('lstAgreg', 'Resultado do turno', ['Média', 'Soma', 'Último valor']),
                  ('lstFonte', 'Fonte', ['MES', 'SIMULAÇÃO']),
                  ('lstSN', 'Sim/Não', ['Sim', 'Não']),
                  ('lstTipo', 'Tipo de tag', [MV, AM]),
                  ('lstFormatoData', 'Formato data MES', ['Data do Excel', 'Texto dd/mm/aaaa', 'Texto mm/dd/aaaa'])]
        ws.merge_range(4, 18, 4, 27, 'Listas',
                       self.f(bold=True, font_color=BRANCO, bg_color=AZUL_ACINZ, indent=1))
        for i, (nm, tit, itens) in enumerate(listas):
            c = 18 + i
            ws.write(5 - 1 + 1, c, tit, self.f(bold=True, font_size=9, bg_color=FUNDO_GRUPO, text_wrap=True))
            for j, it in enumerate(itens):
                ws.write(6 + j, c, it, self.f(font_size=9, border=1, border_color=FUNDO_CLARO))
            self.nome(nm, ws, 6, c, 6 + len(itens) - 1, c)

    # ------------------------------------------------------------ Dados_MES
    def aba_dados_mes(self):
        """Tres consultas ao Aspen (Qualidade, Producao, Ritmo de processo) no mesmo formato da
        planilha de referencia. O VBA (ModMES.RegerarFormulasMES) reescreve estas formulas a partir
        da aba Configuracoes antes de cada atualizacao, sempre com as listas em texto literal."""
        ws = self.ws_mes
        ws.set_column(0, 0, 18, self.f(num_format='dd/mm/yyyy hh:mm'))
        ws.set_column(1, 2 * NPARAM, 14)
        ws.write(0, 0, 'Dados brutos do MES (Aspen IP.21) - não editar', self.f(bold=True, font_size=13, font_color=AZUL))
        txt = self.f(num_format='dd/mm/yyyy hh:mm:ss', bg_color=FUNDO_CLARO, align='left')
        rot = self.f(bold=True)
        ws.write(2, 0, 'Início', rot)
        ws.write_blank(2, 1, None, txt)
        ws.write(3, 0, 'Fim', rot)
        ws.write_blank(3, 1, None, txt)
        self.nome('mesInicio', ws, 2, 1)
        self.nome('mesFim', ws, 3, 1)
        ws.write(2, 3, 'Prefixo Aspen', rot)
        ws.write_blank(2, 4, None)
        self.nome('mesPrefixo', ws, 2, 4)
        ws.write(3, 3, 'Última consulta', rot)
        ws.write_blank(3, 4, None)
        self.nome('mesTempo', ws, 3, 4)
        self.blocos = []
        for cod, titulo, calc, lin in BLOCOS:
            ps = [i + 1 for i, pr in enumerate(PARAMS) if pr[11] == cod]
            psc = [p for p in ps if configurado(p)]    # so analises com tag entram na consulta
            r_anc = lin - 1                   # 0-based
            r_out = r_anc + 2
            ncol = 1 + 2 * len(psc)
            anc = rc(r_anc, 0)
            out = rc(r_out, 0)
            ws.write(r_anc - 1, 0, 'Consulta: %s  (tipo de cálculo "%s")' % (titulo, calc),
                     self.f(bold=True, font_color=AZUL_ACINZ))
            ws.write_formula(r_anc, 0, formula_consulta(ps, calc, anc, out), self.f(font_size=8), 'Success')
            ws.write(r_anc + 1, 0, 'Data/hora', self.f(bold=True, bg_color=AZUL, font_color=BRANCO))
            for i, pp in enumerate(psc):
                for k, us in enumerate(USINAS):
                    ws.write(r_anc + 1, 1 + i * 2 + k, 'P%02d %s · %s' % (pp, us, PARAMS[pp - 1][1]),
                             self.f(bold=True, bg_color=AZUL, font_color=BRANCO, font_size=8, text_wrap=True))
            ws.set_row(r_anc + 1, 45)
            f_show = ('{=_xll.AspenTech.PME.ProcessData.Functions.ShowCalculationValues('
                      'ADDRESS(ROW(Dados_MES!%s),COLUMN(Dados_MES!%s),1,,"Dados_MES"),Dados_MES!%s, 0)}' % (anc, anc, anc))
            ws.write_array_formula(r_out, 0, r_out + NSLOT - 1, ncol - 1, f_show, self.f(num_format='0.00'))
            self.blocos.append((cod, titulo, calc, anc, out, ps, ncol))
        self.botao(ws, 0, 7, 'Voltar às Configurações', 'OcultarDadosMES', 200, 28, 'primario', y=2)

    # ------------------------------------------------------------ Limites (oculta)
    def aba_limites(self):
        """Limites quimicos por produto (farol verde/vermelho dos Resumos e dos Resultados gerais).
        Aba oculta: botao direito numa guia -> Reexibir -> Limites. Pode ser editada/ampliada a mao."""
        ws = self.ws_lim
        n = len(LIM_ITENS)
        UC = LIM_COL1 + 2 * n - 1
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 14)
        ws.set_column(2, 2, 12)
        ws.set_column(LIM_COL1, UC, 8.5)
        ws.set_column(UC + 1, UC + 1, 3)
        self.cabecalho(ws, UC, 'LIMITES DE PROCESSO', '',
                       col_logo_fim=3)
        ws.set_row(4, 8)
        ws.set_row(5, 30)
        hdr = self.f(bold=True, font_size=9, font_color=BRANCO, bg_color=AZUL, align='center', text_wrap=True,
                     border=1, border_color=BRANCO)
        hdr2 = self.f(bold=True, font_size=8, font_color=AZUL_TITULO, bg_color=FUNDO_GRUPO, align='center',
                      border=1, border_color=BORDA)
        ws.set_row(6, 30)
        ws.merge_range(6, 1, 7, 1, 'Produto', hdr)
        ws.merge_range(6, 2, 7, 2, 'Concentrado', hdr)
        npel = sum(1 for k, _ in LIM_ITENS if k.startswith('LM'))
        ws.merge_range(6, LIM_COL1, 6, LIM_COL1 + 2 * npel - 1,
                       'PELOTA  |  PROCESSO (GPU), DADOS HORÁRIOS', hdr)
        ws.merge_range(6, LIM_COL1 + 2 * npel, 6, UC,
                       'PELLET FEED (FILTRAGEM)  |  POR CONCENTRADO', hdr)
        for i, (k, rotulo) in enumerate(LIM_ITENS):
            ws.write(7, LIM_COL1 + 2 * i, rotulo + ' mín.', hdr2)
            ws.write(7, LIM_COL1 + 2 * i + 1, rotulo + ' máx.', hdr2)
        inp = dict(bg_color=FUNDO_INPUT, border=1, border_color=BORDA, locked=False, align='center')
        for j in range(LIM_NLIN):
            r = LIM_ROW1 + j
            ws.set_row(r, 17)
            if j < len(LIMITES):
                prod, conc, lim = LIMITES[j]
                lim = dict(lim, **PF_CONC[conc])
            else:
                prod, conc, lim = '', '', {}
            ws.write(r, 1, prod, self.f(bold=True, indent=1, **inp))
            ws.write_blank(r, 2, None, self.f(**inp))
            for i, (k, _) in enumerate(LIM_ITENS):
                mn, mx = lim.get(k, (None, None))
                dec = 3 if k in ('LM_P', 'PF_P') else 2
                fm = self.f(num_format=fmt_dec(dec), **inp)
                for q, v in enumerate((mn, mx)):
                    if v is None:
                        ws.write_blank(r, LIM_COL1 + 2 * i + q, None, fm)
                    else:
                        ws.write_number(r, LIM_COL1 + 2 * i + q, v, fm)
        r = LIM_ROW1 + LIM_NLIN + 1
        # lista de produtos para a selecao (cresce sozinha com linhas novas)
        self.wb.define_name('lstProdutos', "=OFFSET('Limites'!$B$%d,0,0,MAX(1,COUNTA('Limites'!$B$%d:$B$%d)),1)"
                            % (LIM_ROW1 + 1, LIM_ROW1 + 1, LIM_ROW1 + LIM_NLIN))
        ws.freeze_panes(8, 3)
        ws.print_area(1, 1, LIM_ROW1 + LIM_NLIN + 3, UC)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Preenchimento
    def aba_preenchimento(self):
        """Uma aba so para o tecnico: Turno Dia em cima, Turno Noite embaixo, mesmos campos.
        Tudo o que e preenchido aqui aparece sozinho (por formula) no Resumo do turno."""
        ws = self.ws_pre
        UC = 10
        self.pre_uc = UC
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 5)
        ws.set_column(2, UC, 10.3)
        ws.set_column(UC + 1, UC + 1, 3)
        ws.set_column(UC + 2, UC + 2, 30)
        self.cabecalho(ws, UC, 'RELATÓRIO DE PASSAGEM DE TURNO', 'LABORATÓRIO QUÍMICO', col_logo_fim=4)
        rot = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                     border=1, border_color=BORDA, text_wrap=True)
        ws.set_row(4, 8)
        ws.set_row(5, 26)
        ws.merge_range(5, 1, 5, 3, 'Data do dia', rot)
        ws.merge_range(5, 4, 5, 5, '', self.f_input(num_format='dd/mm/yyyy', bold=True, font_size=12,
                                                    font_color=AZUL_TITULO, align='center'))
        self.nome('pData', ws, 5, 4)
        ws.data_validation(5, 4, 5, 4, {'validate': 'date', 'criteria': '>', 'value': dtm.date(2020, 1, 1),
                                        'error_message': 'Digite uma data válida (dd/mm/aaaa).'})
        ws.merge_range(5, 6, 5, UC, 'Turno Dia 07:00 às 19:00   |   Turno Noite 19:00 às 07:00',
                       self.f(font_size=9, font_color=TEXTO_SEC, indent=1))
        self.campos = {'d': {}, 'n': {}}
        r = 7
        ws.set_row(r, 18)
        self.link_pre_row = r
        r = self._bloco_turno(ws, 'd', 'TURNO DIA   07:00 – 19:00', r + 2)
        r = self._bloco_turno(ws, 'n', 'TURNO NOITE   19:00 – 07:00', r + 2)
        fim = r
        # atalhos (hiperlinks, sem macro)
        ws.write(self.link_pre_row, 1, 'Ir para:', self.f(font_size=9, bold=True, font_color=TEXTO_SEC))
        self.link(ws, self.link_pre_row, 2, "'Preenchimento'!B%d" % (self.bloco_row['n'] + 1), 'Turno Noite')
        self.link(ws, self.link_pre_row, 4, "'Resumo Dia'!A1", 'Resumo Dia')
        self.link(ws, self.link_pre_row, 6, "'Resumo Noite'!A1", 'Resumo Noite')
        self.link(ws, self.link_pre_row, 8, "'Resultados gerais'!A1", 'Resultados gerais')
        ws.print_area(1, 1, fim, UC)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_h_pagebreaks([self.bloco_row['n']])
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.freeze_panes(7, 0)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    def _bloco_turno(self, ws, t, titulo, r):
        """Campos de um turno (t = 'd' Dia, 'n' Noite). Devolve a ultima linha usada."""
        UC = self.pre_uc
        self.bloco_row = getattr(self, 'bloco_row', {})
        self.bloco_row[t] = r
        c = self.campos[t]
        num = self.f(font_size=8, font_color=TEXTO_SEC, align='center', bg_color=FUNDO_CLARO, border=1,
                     border_color=BORDA)
        # texto longo quebra a linha; a altura e ajustada pelo VBA (evento da aba Preenchimento)
        inp = self.f_input(indent=1, font_size=10, text_wrap=True)
        rot = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                     border=1, border_color=BORDA, text_wrap=True)
        # faixa do turno
        ws.set_row(r, 26)
        ws.merge_range(r, 1, r, UC, titulo, self.f(bold=True, font_size=13, font_color=BRANCO, bg_color=AZUL,
                                                   indent=1, bottom=3, bottom_color=AMARELO))
        self.link(ws, r, UC + 2, "'Resumo %s'!A1" % ('Dia' if t == 'd' else 'Noite'),
                  'Resumo %s' % ('Dia' if t == 'd' else 'Noite'), 10)
        r += 1
        ws.set_row(r, 22)
        ws.merge_range(r, 1, r, 2, 'Letra', rot)
        ws.write(r, 3, '', self.f_input(bold=True, align='center', font_size=11))
        ws.data_validation(r, 3, r, 3, {'validate': 'list', 'source': '=lstTurma'})
        self.nome(t + 'Letra', ws, r, 3)
        c['Letra'] = t + 'Letra'
        ws.write(r, 4, 'Técnico', rot)
        ws.merge_range(r, 5, r, UC, '', self.f_input(bold=True, indent=1, font_size=11))
        self.nome(t + 'Tecnico', ws, r, 5)
        r += 1
        ws.set_row(r, 22)
        prod = self.f_input(bold=True, align='center', font_size=11, font_color=AZUL_TITULO)
        for i, us in enumerate(('US3', 'US4')):
            c1 = 1 + 5 * i
            ws.merge_range(r, c1, r, c1 + 1, 'Produto ' + us, rot)
            ws.merge_range(r, c1 + 2, r, c1 + 4 if i == 0 else UC, '', prod)
            ws.data_validation(r, c1 + 2, r, c1 + 2, {'validate': 'list', 'source': '=lstProdutos',
                                                      'error_title': 'Produto',
                                                      'error_message': 'Escolha um produto da lista.'})
            self.nome(t + 'Prod' + us, ws, r, c1 + 2)
        r += 2

        def lista(r, titulo, nome, n, dica):
            ws.set_row(r, 20)
            ws.merge_range(r, 1, r, UC, titulo, self.f(bold=True, font_size=11, font_color=AZUL_TITULO,
                                                       valign='bottom', bottom=2, bottom_color=AZUL))
            for k in range(n):
                rr = r + 1 + k
                ws.set_row(rr, 19)
                ws.write(rr, 1, k + 1, num)
                ws.merge_range(rr, 2, rr, UC, '', inp)
            self.nome(t + nome, ws, r + 1, 2, r + n, 2)
            c[nome] = [rc(r + 1 + k, 2, True, True) for k in range(n)]
            return r + n + 2

        def campo(rr, c1, c2, texto, v1, v2, nome, padrao='', validar=None, fmt=None):
            fmt = fmt or inp
            if c1 == c2:
                ws.write(rr, c1, texto, rot)
            else:
                ws.merge_range(rr, c1, rr, c2, texto, rot)
            if v1 == v2:
                ws.write(rr, v1, padrao, fmt)
            else:
                ws.merge_range(rr, v1, rr, v2, padrao, fmt)
            if validar:
                ws.data_validation(rr, v1, rr, v1, {'validate': 'list', 'source': validar})
            self.nome(t + nome, ws, rr, v1)
            c[nome] = t + nome

        r = lista(r, '1. Tarefas realizadas', 'Real', 12, 'uma tarefa por linha')
        r = lista(r, '2. Solicitações', 'Sol', 4, 'nº da solicitação, amostra, pendência')
        r = lista(r, '3. Equipamentos', 'Equip', 4, 'falhas, vazamentos, manutenção')
        r = lista(r, '4. Tarefas a realizar (próximo turno)', 'AReal', 8, 'uma tarefa por linha')

        r = lista(r, '5. Dragas', 'Dragas', 3, '')

        # 6. Mineroduto 03: batch e teores (como no Fisico)
        ws.set_row(r, 20)
        ws.merge_range(r, 1, r, UC, '6. Mineroduto 03', self.f(bold=True, font_size=11, font_color=AZUL_TITULO,
                                                                 valign='bottom', bottom=2, bottom_color=AZUL))
        rr = r + 1
        cen = self.f_input(align='center', font_size=10)
        ws.set_row(rr, 20)
        campo(rr, 1, 2, 'Batch', 3, 4, 'Batch', '', fmt=self.f_input(align='center', bold=True, font_size=10))
        campo(rr, 5, 5, 'Fe', 6, 6, 'BatchFe', '', fmt=cen)
        campo(rr, 7, 7, 'SiO2', 8, 8, 'BatchSiO2', '', fmt=cen)
        campo(rr, 9, 9, 'Al2O3', 10, 10, 'BatchAl2O3', '', fmt=cen)
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 2, 'P', 3, 4, 'BatchP', '', fmt=cen)
        campo(rr, 5, 5, 'PPC', 6, 6, 'BatchPPC', '', fmt=cen)
        for k in range(2):
            rr += 1
            ws.set_row(rr, 19)
            ws.write(rr, 1, k + 1, num)
            ws.merge_range(rr, 2, rr, UC, '', inp)
        self.nome(t + 'Mineroduto', ws, rr - 1, 2, rr, 2)
        c['Mineroduto'] = [rc(rr - 1 + k, 2, True, True) for k in range(2)]
        r = rr + 2

        r = lista(r, '7. Filtragem', 'Filtragem', 6, '')

        self.secao(ws, r, 1, UC, '8. Comentários por usina')
        rr = r + 1
        for us in ('US3', 'US4'):
            ws.set_row(rr, 19)
            ws.set_row(rr + 1, 19)
            ws.set_row(rr + 2, 19)
            ws.merge_range(rr, 1, rr + 2, 3, 'Comentário ' + us, self.f(font_size=10, bold=True, font_color=AZUL_TITULO,
                                                                        bg_color=FUNDO_CLARO, indent=1, border=1,
                                                                        border_color=BORDA, valign='vcenter'))
            ws.merge_range(rr, 4, rr + 2, UC, '', self.f_input(indent=1, text_wrap=True, valign='top'))
            self.nome(t + 'Com' + us, ws, rr, 4)
            c['Com' + us] = t + 'Com' + us
            rr += 3
        r = rr + 1

        self.secao(ws, r, 1, UC, '9. Controle do laboratório')
        prog = ('OREGON (RX), Carbono (Leco CS-230), PCS (Calorímetro), mufla 1000 ºC, fotômetro, balanças, '
                'estufa e máquina de fusão')
        rr = r + 1
        ws.set_row(rr, 32)
        campo(rr, 1, 3, 'Programa em uso', 4, UC, 'Prog', '',
              fmt=self.f_input(indent=1, font_size=9, text_wrap=True))
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Foi necessário preparar padrões?', 4, 5, 'Padroes', '', '=lstSN')
        campo(rr, 6, 6, 'Quais', 7, UC, 'PadroesQuais')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Sistema de ar utilizado', 4, 5, 'Ar', '', '=lstAr')
        campo(rr, 6, 6, 'Status', 7, 8, 'ArStatus', '', '=lstLigado')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Compressor', 4, 5, 'Compressor', '', '=lstCond')
        campo(rr, 6, 6, 'Nitrogênio', 7, 8, 'Nitrogenio', '', '=lstCond')
        r = rr + 2

        self.secao(ws, r, 1, UC, '10. Cadinhos de platina')
        rr = r + 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Repassados para o turno', 4, UC, 'Cadinhos')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Retirado para reforma', 4, UC, 'Reforma')
        r = rr + 2

        self.secao(ws, r, 1, UC, '11. Observações gerais')
        rr = r + 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Hidrogênio no carvão', 4, 5, 'H2Carvao')
        campo(rr, 6, 7, 'Hidrogênio no coque', 8, UC, 'H2Coque')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Coque Planta 04', 4, UC, 'Coque04')
        rr += 1
        ws.merge_range(rr, 1, rr + 2, 3, 'Observações', rot)
        ws.merge_range(rr, 4, rr + 2, UC, '', self.f_input(indent=1, text_wrap=True, valign='top'))
        self.nome(t + 'Obs', ws, rr, 4)
        c['Obs'] = t + 'Obs'
        return rr + 2

    # ------------------------------------------------------------ tabela de resultados (Resumo e Resultados gerais)
    def tabela(self, ws, r, nslot, horas, nome_horas, tit_media, prod3, prod4, col_lim):
        """Tabela de analises: Amostra/analise | Un. | Usina | janelas de 2 h | Media | Min | Max.
        As janelas sao gravadas pelo VBA; media, minimo e maximo sao formulas.
        Chaves na coluna A (fonte branca): G1.. = grupo, P02U3 = analise/usina (usadas pelo VBA)."""
        CS = 4
        CM = CS + nslot
        UC = CM + 2
        chave_fmt = self.f(font_color=BRANCO, font_size=6)
        hdr = self.f(bold=True, font_size=9, font_color=BRANCO, bg_color=AZUL, align='center', text_wrap=True,
                     border=1, border_color=BRANCO)
        ws.set_row(r, 30)
        cabs = ['Amostra / análise', 'Un.', 'Usina'] + horas + [tit_media, 'Mín.', 'Máx.']
        for i, t in enumerate(cabs):
            ws.write(r, 1 + i, t, hdr)
        self.nome(nome_horas, ws, r, CS, r, CS + nslot - 1)
        r += 1
        grupo = None
        gi = 0
        verde = self.wb.add_format({'font_color': VERDE_TXT, 'bold': True, 'bg_color': VERDE_FUNDO})
        vermelho = self.wb.add_format({'font_color': '#B42318', 'bold': True, 'bg_color': '#FDE3E1'})
        # colunas auxiliares (ocultas) com o limite minimo e maximo do produto de cada linha
        ws.set_column(col_lim, col_lim + 1, 8, None, {'hidden': True})
        for p, pr in enumerate(PARAMS, start=1):
            g, nome, un, dec = pr[:4]
            unica = pr[12]
            if g != grupo:
                grupo = g
                gi += 1
                ws.set_row(r, 18)
                ws.write(r, 0, 'G%d' % gi, chave_fmt)
                ws.merge_range(r, 1, r, UC, g, self.f(bold=True, font_size=10, font_color=AZUL_TITULO,
                                                      bg_color=FUNDO_GRUPO, indent=1, border=1, border_color=BORDA))
                r += 1
            zebra = BRANCO if p % 2 else '#F8F9FA'
            b = dict(border=1, border_color=BORDA, bg_color=zebra)
            nf = fmt_dec(dec)
            nlin = 1 if unica else 2
            fnome = self.f(font_size=10, bold=True, font_color=TEXTO, indent=1, **b)
            fun = self.f(font_size=9, font_color=TEXTO_SEC, align='center', **b)
            if nlin == 2:
                ws.merge_range(r, 1, r + 1, 1, nome, fnome)
                ws.merge_range(r, 2, r + 1, 2, un, fun)
            else:
                ws.write(r, 1, nome, fnome)
                ws.write(r, 2, un, fun)
            chave_lim = PARAM_LIM.get((g, nome))
            for k in range(nlin):
                rr = r + k
                ws.set_row(rr, 17)
                ws.write(rr, 0, 'P%02dU%d' % (p, 3 + k), chave_fmt)
                ws.write(rr, 3, 'US3/4' if unica else USINAS[k],
                         self.f(bold=True, font_size=9, font_color=AZUL if k == 0 else AZUL_ACINZ, align='center', **b))
                vals = '%s:%s' % (rc(rr, CS), rc(rr, CM - 1))
                for cc in range(CS, CM):
                    ws.write_blank(rr, cc, None, self.f(font_size=10, font_color=TEXTO, align='center',
                                                        num_format=nf, **b))
                ws.write_formula(rr, CM, '=IF(COUNT(%s)=0,"",AVERAGE(%s))' % (vals, vals),
                                 self.f(font_size=10.5, bold=True, font_color=AZUL_TITULO, align='center',
                                        num_format=nf, **dict(b, bg_color=FUNDO_AZUL_CLARO)))
                ws.write_formula(rr, CM + 1, '=IF(COUNT(%s)=0,"",MIN(%s))' % (vals, vals),
                                 self.f(font_size=9, font_color=AZUL_ACINZ, align='center', num_format=nf, **b))
                ws.write_formula(rr, CM + 2, '=IF(COUNT(%s)=0,"",MAX(%s))' % (vals, vals),
                                 self.f(font_size=9, font_color=AZUL_ACINZ, align='center', num_format=nf, **b))
                # farol: limite do produto selecionado (aba Limites)
                if unica:
                    pr_ref = 'IF(%s="",%s,%s)' % (prod3, prod4, prod3)
                else:
                    pr_ref = prod3 if k == 0 else prod4
                for q in (0, 1):
                    if chave_lim:
                        col = lim_col(chave_lim, q)
                        idx = "INDEX('Limites'!$%s:$%s,MATCH(%s,'Limites'!$B:$B,0))" % (col, col, pr_ref)
                        ws.write_formula(rr, col_lim + q, '=IFERROR(IF(%s="","",%s),"")' % (idx, idx))
                    else:
                        ws.write_blank(rr, col_lim + q, None)
                if chave_lim:
                    c0 = rc(rr, CS, False, False)
                    v = 'ROUND(%s,%d)' % (c0, dec)
                    mn = '$%s%d' % (colname(col_lim), rr + 1)
                    mx = '$%s%d' % (colname(col_lim + 1), rr + 1)
                    fora = 'OR(AND(ISNUMBER({mn}),{v}<{mn}),AND(ISNUMBER({mx}),{v}>{mx}))'.format(v=v, mn=mn, mx=mx)
                    ws.conditional_format(rr, CS, rr, CM, {
                        'type': 'formula', 'format': vermelho,
                        'criteria': '=AND(ISNUMBER({c}),{f})'.format(c=c0, f=fora)})
                    ws.conditional_format(rr, CS, rr, CM, {
                        'type': 'formula', 'format': verde,
                        'criteria': '=AND(ISNUMBER({c}),OR(ISNUMBER({mn}),ISNUMBER({mx})),NOT({f}))'.format(
                            c=c0, mn=mn, mx=mx, f=fora)})
            r += nlin
        return r, UC

    # ------------------------------------------------------------ Resumo Dia / Resumo Noite
    def secao_faixa(self, ws, r, c1, c2, texto):
        ws.set_row(r, 21)
        ws.merge_range(r, c1, r, c2, texto, self.f(bold=True, font_size=11, font_color=BRANCO, bg_color=AZUL_TITULO,
                                                   indent=1))

    def aba_resumo(self, ws, t, nome_turno, sub_turno, hora_ini):
        """Resumo de UM turno (imagem para o e-mail), no mesmo formato do Preenchimento:
        identificacao, ocorrencias, comentarios por usina, controle do laboratorio, cadinhos, observacoes
        (tudo por FORMULA a partir do Preenchimento) e informativo de qualidade do MES."""
        c = self.campos[t]
        CS = 4
        UC = CS + NSLOT + 2          # M
        CB = UC + 2                  # O (botoes)
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 28)
        ws.set_column(2, 3, 9)
        ws.set_column(CS, CS + NSLOT - 1, 10.5)
        ws.set_column(CS + NSLOT, CS + NSLOT, 12)
        ws.set_column(CS + NSLOT + 1, UC, 10)
        ws.set_column(UC + 1, UC + 1, 3)
        ws.set_column(CB, CB, 30)
        chave_fmt = self.f(font_color=BRANCO, font_size=6)
        P = "'Preenchimento'!"
        self.cabecalho(ws, UC, 'RELATÓRIO DE PASSAGEM DE TURNO', 'LABORATÓRIO QUÍMICO')

        b = dict(border=1, border_color=BORDA)
        rot = self.f(font_size=10, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1, **b)
        val = self.f(font_size=11, bold=True, font_color=TEXTO, indent=1, **b)
        valc = self.f(font_size=11, bold=True, font_color=TEXTO, align='center', **b)
        txt = self.f(font_size=10, font_color=TEXTO, indent=1, text_wrap=True, valign='vcenter', **b)

        def valor(nm, vazio='—'):
            return '=IF(%s="","%s",%s)' % (c[nm], vazio, c[nm])

        def campo(r, c1, c2, rotulo, v1, v2, formula, fmt=val):
            if c1 == c2:
                ws.write(r, c1, rotulo, rot)
            else:
                ws.merge_range(r, c1, r, c2, rotulo, rot)
            if v1 == v2:
                ws.write_formula(r, v1, formula, fmt)
            else:
                ws.merge_range(r, v1, r, v2, '', fmt)
                ws.write_formula(r, v1, formula, fmt)

        # ---- identificacao
        ws.set_row(4, 8)
        r = 5
        ws.set_row(r, 22)
        campo(r, 1, 1, 'Data', 2, 3, '=IF(pData="","—",TEXT(DAY(pData),"00")&"/"&TEXT(MONTH(pData),"00")&"/"&'
                                     'YEAR(pData))', valc)
        campo(r, 4, 4, 'Turno', 5, 7, '="%s"' % sub_turno, valc)
        campo(r, 8, 9, 'Letra', 10, UC, valor('Letra'), valc)
        r += 1
        ws.set_row(r, 22)
        campo(r, 1, 1, 'Técnico', 2, 6, '=IF(%sTecnico="","—",%sTecnico)' % (t, t))
        campo(r, 7, 8, 'Produto US3', 9, 9, '=IF(%sProdUS3="","—",%sProdUS3)' % (t, t), valc)
        campo(r, 10, 11, 'Produto US4', 12, UC, '=IF(%sProdUS4="","—",%sProdUS4)' % (t, t), valc)
        r += 2

        # ---- ocorrencias (listas linha a linha, como no Preenchimento)
        # Chaves na coluna A (usadas pelo VBA para ocultar o que nao foi preenchido):
        #   S1 = faixa de secao, S2 = subtitulo, E = espaco da secao, T = linha de lista (texto em B),
        #   W = texto livre (texto em C), F/V = linha de campos preenchida/vazia (formula), Z = fim das secoes.
        # Secao (ou subtitulo) sem nenhuma linha preenchida fica oculta inteira.
        def espaco(r, altura=6):
            ws.set_row(r, altura)
            ws.write(r, 0, 'E', chave_fmt)

        def chave_campos(r, nomes):
            vazio = 'AND(%s)' % ','.join('%s=""' % c[n] for n in nomes)
            ws.write_formula(r, 0, '=IF(%s,"V","F")' % vazio, chave_fmt, 'F')

        self.secao_faixa(ws, r, 1, UC, 'OCORRÊNCIAS DO TURNO')
        ws.write(r, 0, 'S1', chave_fmt)
        r += 1
        sub = self.f(bold=True, font_size=10, font_color=AZUL_TITULO, bg_color=FUNDO_GRUPO, indent=1, **b)
        lin = self.f(font_size=10, font_color=TEXTO, indent=1, text_wrap=True, valign='top', left=1, right=1,
                     border_color=BORDA)
        for i, (titulo, nm) in enumerate((('Tarefas realizadas', 'Real'), ('Solicitações', 'Sol'),
                                          ('Equipamentos', 'Equip'), ('Tarefas a realizar (próximo turno)', 'AReal'))):
            ws.set_row(r, 18)
            ws.write(r, 0, 'S2', chave_fmt)
            ws.merge_range(r, 1, r, UC, '%d. %s' % (i + 1, titulo), sub)
            r += 1
            for k, ref in enumerate(c[nm]):
                # numero junto com o texto (sem coluna larga so para a numeracao)
                ws.set_row(r, 17)
                ws.write(r, 0, 'T', chave_fmt)
                ws.merge_range(r, 1, r, UC, '', lin)
                ws.write_formula(r, 1, '=IF(TRIM(%s%s)="","","%d.   "&TRIM(%s%s))' % (P, ref, k + 1, P, ref),
                                 lin)
                r += 1
            espaco(r, 4)
            r += 1

        def linha_t(r, formula):
            ws.set_row(r, 17)
            ws.write(r, 0, 'T', chave_fmt)
            ws.merge_range(r, 1, r, UC, '', lin)
            ws.write_formula(r, 1, formula, lin)

        # ---- dragas
        ws.set_row(r, 18)
        ws.write(r, 0, 'S2', chave_fmt)
        ws.merge_range(r, 1, r, UC, '5. Dragas', sub)
        r += 1
        for k, ref in enumerate(c['Dragas']):
            linha_t(r, '=IF(TRIM(%s%s)="","","%d.   "&TRIM(%s%s))' % (P, ref, k + 1, P, ref))
            r += 1
        espaco(r, 4)
        r += 1

        # ---- mineroduto: "Processando batch 267 com teores: Fe: ...; SiO2: ...; ..."
        ws.set_row(r, 18)
        ws.write(r, 0, 'S2', chave_fmt)
        ws.merge_range(r, 1, r, UC, '6. Mineroduto 03', sub)
        r += 1
        teores = [('BatchFe', 'Fe'), ('BatchSiO2', 'SiO2'), ('BatchAl2O3', 'Al2O3'), ('BatchP', 'P'),
                  ('BatchPPC', 'PPC')]
        partes = '&'.join('IF(%s="","","%s: "&%s&"; ")' % (c[k], rotulo, c[k]) for k, rotulo in teores)
        sem_teor = 'AND(%s)' % ','.join('%s=""' % c[k] for k, _ in teores)
        linha_t(r, '=IF(AND(%s="",%s),"","Processando batch "&%s&IF(%s,""," com teores: "&LEFT(%s,LEN(%s)-2)))'
                % (c['Batch'], sem_teor, c['Batch'], sem_teor, partes, partes))
        r += 1
        for ref in c['Mineroduto']:
            linha_t(r, '=IF(TRIM(%s%s)="","",TRIM(%s%s))' % (P, ref, P, ref))
            r += 1
        espaco(r, 4)
        r += 1
        # ---- filtragem
        ws.set_row(r, 18)
        ws.write(r, 0, 'S2', chave_fmt)
        ws.merge_range(r, 1, r, UC, '7. Filtragem', sub)
        r += 1
        for k, ref in enumerate(c['Filtragem']):
            linha_t(r, '=IF(TRIM(%s%s)="","","%d.   "&TRIM(%s%s))' % (P, ref, k + 1, P, ref))
            r += 1
        espaco(r, 4)
        r += 1

        # ---- comentarios por usina
        ws.set_row(r, 18)
        ws.write(r, 0, 'S2', chave_fmt)
        ws.merge_range(r, 1, r, UC, '8. Comentários por usina', sub)
        r += 1
        for us in ('US3', 'US4'):
            ws.set_row(r, 34)
            ws.write(r, 0, 'W', chave_fmt)
            ws.write(r, 1, 'Comentário ' + us, rot)
            ws.merge_range(r, 2, r, UC, '', txt)
            ws.write_formula(r, 2, valor('Com' + us, ''), txt)
            r += 1
        espaco(r, 15)
        r += 1

        # ---- controle do laboratorio
        self.secao_faixa(ws, r, 1, UC, 'CONTROLE DO LABORATÓRIO')
        ws.write(r, 0, 'S1', chave_fmt)
        r += 1
        ws.set_row(r, 30)
        ws.write(r, 0, 'W', chave_fmt)
        campo(r, 1, 1, 'Programa em uso', 2, UC, valor('Prog', ''), txt)
        r += 1
        ws.set_row(r, 20)
        chave_campos(r, ['Padroes', 'PadroesQuais'])
        campo(r, 1, 1, 'Foi necessário preparar padrões?', 2, 3, valor('Padroes'), valc)
        campo(r, 4, 5, 'Quais', 6, UC, valor('PadroesQuais'))
        r += 1
        ws.set_row(r, 20)
        chave_campos(r, ['Ar', 'ArStatus', 'Compressor', 'Nitrogenio'])
        campo(r, 1, 1, 'Sistema de ar utilizado', 2, 3, valor('Ar'), valc)
        campo(r, 4, 5, 'Status do equipamento', 6, 7, valor('ArStatus'), valc)
        campo(r, 8, 9, 'Compressor', 10, 10, valor('Compressor'), valc)
        campo(r, 11, 11, 'Nitrogênio', 12, UC, valor('Nitrogenio'), valc)
        r += 1
        espaco(r, 15)
        r += 1

        # ---- cadinhos e observacoes
        self.secao_faixa(ws, r, 1, UC, 'CADINHOS DE PLATINA')
        ws.write(r, 0, 'S1', chave_fmt)
        r += 1
        ws.set_row(r, 20)
        chave_campos(r, ['Cadinhos'])
        campo(r, 1, 1, 'Repassados para o turno', 2, UC, valor('Cadinhos'), txt)
        r += 1
        ws.set_row(r, 20)
        chave_campos(r, ['Reforma'])
        campo(r, 1, 1, 'Retirado para reforma', 2, UC, valor('Reforma'), txt)
        r += 1
        espaco(r, 15)
        r += 1
        self.secao_faixa(ws, r, 1, UC, 'OBSERVAÇÕES GERAIS')
        ws.write(r, 0, 'S1', chave_fmt)
        r += 1
        ws.set_row(r, 20)
        chave_campos(r, ['H2Carvao', 'H2Coque', 'Coque04'])
        campo(r, 1, 1, 'Hidrogênio no carvão', 2, 3, valor('H2Carvao'), valc)
        campo(r, 4, 5, 'Hidrogênio no coque', 6, 7, valor('H2Coque'), valc)
        campo(r, 8, 9, 'Coque Planta 04', 10, UC, valor('Coque04'))
        r += 1
        ws.set_row(r, 34)
        ws.write(r, 0, 'W', chave_fmt)
        campo(r, 1, 1, 'Observações', 2, UC, valor('Obs', ''), txt)
        r += 1
        espaco(r, 15)
        r += 1

        # ---- resultados quimicos
        self.secao_faixa(ws, r, 1, UC, 'INFORMATIVO DE QUALIDADE DO TURNO  |  RESULTADOS DO MES')
        ws.write(r, 0, 'Z', chave_fmt)
        r += 1
        ws.set_row(r, 17)
        ws.merge_range(r, 1, r, UC, 'Resultados não atualizados',
                       self.f(font_size=9, italic=True, font_color=TEXTO_SEC, indent=1))
        self.nome(t + 'Atualizado', ws, r, 1)
        r += 1
        horas = ['%02d:30' % ((hora_ini + 2 * i) % 24) for i in range(NSLOT)]
        r, _ = self.tabela(ws, r, NSLOT, horas, t + 'Horas', 'Média do turno', t + 'ProdUS3', t + 'ProdUS4',
                           UC + 4)
        ws.set_row(r, 17)
        ws.merge_range(r, 1, r, UC, 'Verde: dentro do limite   |   Vermelho: fora do limite   |   '
                       'Limites: SMIN-POP-GEA-001 rev. 12   |   Tags: Plano Amostral rev. 06',
                       self.f(font_size=8.5, font_color=TEXTO_SEC, italic=True, indent=1))
        fim = r
        self.nome(t + 'Area', ws, 1, 1, fim, UC)
        ws.print_area(1, 1, fim, UC)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.center_horizontally()

        T = nome_turno
        self.botao(ws, 1, CB, 'Atualizar dados do MES', 'AtualizarMES' + T, 210, 36, 'primario', x=4, y=8)
        self.botao(ws, 5, CB, 'Copiar imagem', 'CopiarImagem' + T, 210, 36, 'destaque', x=4, y=2)
        self.link(ws, 9, CB, "'Preenchimento'!B%d" % (self.bloco_row[t] + 1), 'Voltar ao Preenchimento', 10)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Resultados gerais
    def aba_resultados(self):
        """Resultados quimicos de qualquer periodo de ate 24 h (inicio e fim escolhidos pelo tecnico)."""
        ws = self.ws_ger
        CS = 4
        CM = CS + NSLOT_INF
        UC = CM + 2
        CB = UC + 2
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 26)
        ws.set_column(2, 2, 6)
        ws.set_column(3, 3, 6)
        ws.set_column(CS, CM - 1, 8.5)
        ws.set_column(CM, CM, 11)
        ws.set_column(CM + 1, UC, 9)
        ws.set_column(UC + 1, UC + 1, 3)
        ws.set_column(CB, CB, 28)
        self.cabecalho(ws, UC, 'RELATÓRIO DE PASSAGEM DE TURNO', 'LABORATÓRIO QUÍMICO', col_logo_fim=3)
        ws.set_row(4, 8)
        ws.set_row(5, 26)
        rot = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                     border=1, border_color=BORDA)
        sel = self.f(font_size=11, bold=True, font_color=AZUL_TITULO, num_format='dd/mm/yyyy hh:mm', align='center',
                     bg_color=FUNDO_INPUT, border=1, border_color=AMARELO, locked=False)
        ws.write(5, 1, 'Período (máx. 24 h)', rot)
        ws.merge_range(5, 2, 5, 3, 'Início', self.f(font_size=9, font_color=TEXTO_SEC, align='right'))
        ws.merge_range(5, CS, 5, CS + 2, '', sel)
        ws.write(5, CS + 3, 'Fim', self.f(font_size=9, font_color=TEXTO_SEC, align='center'))
        ws.merge_range(5, CS + 4, 5, CS + 6, '', sel)
        self.nome('gIni', ws, 5, CS)
        self.nome('gFim', ws, 5, CS + 4)
        for cc in (CS, CS + 4):
            ws.data_validation(5, cc, 5, cc, {'validate': 'date', 'criteria': '>', 'value': dtm.date(2020, 1, 1),
                                              'error_message': 'Digite data e hora: dd/mm/aaaa hh:mm'})
        ws.merge_range(5, CS + 7, 5, UC, 'Formato: dd/mm/aaaa hh:mm', self.f(font_size=8, font_color=TEXTO_SEC,
                                                                              indent=1))
        ws.set_row(6, 6)
        prod = self.f(bold=True, font_size=10, font_color=AZUL_TITULO, align='center', bg_color=FUNDO_INPUT,
                      border=1, border_color=BORDA, locked=False)
        for i, us in enumerate(('US3', 'US4')):
            rr = 7 + i
            ws.set_row(rr, 20)
            ws.write(rr, 1, 'Produto ' + us, rot)
            ws.merge_range(rr, 2, rr, 3, '', prod)
            ws.write_formula(rr, 2, '=IF(dProd{u}="","",dProd{u})'.format(u=us), prod)
            ws.data_validation(rr, 2, rr, 2, {'validate': 'list', 'source': '=lstProdutos'})
            self.nome('gProd' + us, ws, rr, 2)
        rodape = ('Verde: dentro do limite   |   Vermelho: fora do limite   |   '
                  'Limites: SMIN-POP-GEA-001 rev. 12   |   Tags: Plano Amostral rev. 06')
        ws.set_row(9, 15)
        ws.merge_range(9, 1, 9, UC, 'Resultados não atualizados', self.f(font_size=8, italic=True,
                                                                              font_color=TEXTO_SEC, indent=1))
        self.nome('gAtualizado', ws, 9, 1)
        r = 10
        r, _ = self.tabela(ws, r, NSLOT_INF, ['—'] * NSLOT_INF, 'gHoras', 'Média do período', 'gProdUS3',
                           'gProdUS4', CB + 2)
        ws.set_row(r, 16)
        ws.merge_range(r, 1, r, UC, rodape, self.f(font_size=7.5, font_color=TEXTO_SEC, italic=True, indent=1))
        fim = r
        self.nome('gArea', ws, 1, 1, fim, UC)
        ws.print_area(1, 1, fim, UC)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.center_horizontally()
        self.botao(ws, 1, CB, 'Atualizar dados do MES', 'AtualizarResultados', 200, 36, 'primario', x=4, y=8)
        self.botao(ws, 5, CB, 'Copiar imagem', 'CopiarImagemResultados', 200, 36, 'destaque', x=4, y=2)
        self.link(ws, 9, CB, "'Preenchimento'!A1", 'Voltar ao Preenchimento', 10)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ _Mapa
    def aba_mapa(self):
        ws = self.ws_mapa
        for i, t in enumerate(['Bloco', 'Consulta', 'Cálculo', 'Âncora', 'Saída', 'Parâmetros', 'Colunas']):
            ws.write(0, 10 + i, t)
        for r, (cod, titulo, calc, anc, out, ps, ncol) in enumerate(self.blocos, start=1):
            for i, v in enumerate([cod, titulo, calc, anc, out, ','.join(str(x) for x in ps)]):
                ws.write_string(r, 10 + i, v)
            ws.write_number(r, 16, ncol)

    # ------------------------------------------------------------ VBA
    def modulo_layout(self):
        c = CFG_COL
        linhas = ["Option Explicit", "",
                  "' Gerado por build_quimico.py - posicoes fixas das abas (nao editar a mao)",
                  "Public Const NPARAM As Long = %d" % NPARAM,
                  "Public Const NSLOT As Long = %d" % NSLOT,
                  "Public Const NSLOT_MAX As Long = %d" % NSLOT_INF,
                  "Public Const HORAS_SLOT As Double = 2",
                  "' 1a coluna das janelas de 2 h nas tabelas de resultados (E)",
                  "Public Const COL_SLOT1 As Long = 5",
                  "Public Const CFG_ROW1 As Long = %d" % CFG_ROW1]
        for k in ['GRUPO', 'PARAM', 'UNID', 'DEC', 'AGG', 'TAG3', 'TIPO3', 'TAG4', 'TIPO4', 'LIE', 'LSE',
                  'VMIN', 'VMAX', 'TIP3', 'TIP4']:
            linhas.append("Public Const CFG_COL_%s As Long = %d" % (k, c[k]))
        linhas += ["' Senha de protecao das abas (vazio = sem senha)",
                   'Public Const SENHA As String = ""', ""]
        return '\n'.join(linhas)


# ============================================================================
def checar_declaracoes(nome, codigo):
    """O Excel exige declaracoes de modulo (Const, Dim, Private x As...) antes da 1a rotina.
    O LibreOffice aceita fora de ordem, entao a verificacao e feita aqui."""
    em_rotina, viu_rotina = False, False
    for n, linha in enumerate(codigo.split('\n'), start=1):
        t = linha.strip()
        if re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property) ', t):
            em_rotina = viu_rotina = True
        elif re.match(r'^End (Sub|Function|Property)\b', t):
            em_rotina = False
        elif viu_rotina and not em_rotina and re.match(
                r'^(Public |Private |Global )?(Const |Dim |Declare |Type |Enum )|^(Public|Private|Global) \w+ As ', t):
            raise SystemExit('ERRO VBA em %s, linha %d: declaração depois de rotina (o Excel não compila): %s'
                             % (nome, n, t))
    return codigo


def ler_vba(nome):
    with open(os.path.join(VBA_DIR, nome), encoding='utf-8') as fh:
        return checar_declaracoes(nome, fh.read())


def montar_vba(layout_code):
    with open(os.path.join(VBA_DIR, 'ModLayout.bas'), 'w', encoding='utf-8') as fh:
        fh.write(layout_code)
    mods = [{'name': 'ThisWorkbook', 'kind': 'workbook', 'code': ler_vba('ThisWorkbook.cls')}]
    # abas de Resumo: ao abrir, ocultam as linhas vazias (evento simples, sem outras macros)
    evento = ('Option Explicit\n\nPrivate Sub Worksheet_Activate()\n'
              '    On Error Resume Next\n    CompactarAba Me\nEnd Sub\n')
    for cn in ['shPreenchimento', 'shResumoDia', 'shResumoNoite', 'shResultados', 'shLimites', 'shConfig', 'shDadosMES',
               'shMapa']:
        codigo = ''
        if cn in ('shResumoDia', 'shResumoNoite'):
            codigo = evento
        elif cn == 'shPreenchimento':
            # texto longo: a linha aumenta sozinha (celulas mescladas nao tem ajuste automatico no Excel)
            codigo = ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                      '    On Error Resume Next\n    AjustarAlturas Me, Target\nEnd Sub\n')
        mods.append({'name': cn, 'kind': 'sheet', 'code': codigo})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModImagem']:
        mods.append({'name': m, 'kind': 'module', 'code': ler_vba(m + '.bas')})
    if TESTE:
        mods.append({'name': 'ModTesteLO', 'kind': 'module', 'code': ler_vba('../build/ModTesteLO.bas')})
    return build_vba_project(mods)


def pos_processar(caminho):
    """Aplica macros, cantos arredondados e 'nao imprimir' nas caixas de texto usadas como botoes."""
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
    # cantos levemente arredondados (raio pequeno, como os botoes do kit visual Samarco)
    bloco = bloco.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
                          '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/>'
                          '</a:avLst></a:prstGeom>', 1)
    bloco = bloco.replace('<xdr:clientData/>', '<xdr:clientData fPrintsWithSheet="0"/>')
    bloco = bloco.replace('<a:p><a:r>', '<a:p><a:pPr algn="ctr"/><a:r>')
    # Botoes de duas linhas: 1a = titulo, demais = descricao menor e sem negrito
    paras = re.split(r'(?=<a:p>)', bloco)
    textos = [i for i, p in enumerate(paras) if p.startswith('<a:p>') and '<a:r>' in p]
    if len(textos) >= 2:
        for n, i in enumerate(textos):
            if n == 0:
                paras[i] = re.sub(r'sz="\d+"', 'sz="1150"', paras[i])
            else:
                paras[i] = re.sub(r'sz="\d+"', 'sz="850"', paras[i])
                paras[i] = paras[i].replace(' b="1"', ' b="0"')
        bloco = ''.join(paras)
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
