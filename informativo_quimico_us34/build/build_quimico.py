"""Gera a planilha Relatorio_Quimico_US3_US4.xlsm (Laboratório Químico - Usinas 3 e 4).

Mesma base do Informativo do Laboratório Físico (informativo_us34): turnos de 12 h, letras A-D,
resultados do MES (Aspen IP.21) de 2 em 2 horas e relatório do turno pronto para o e-mail.
Aqui o relatório é UMA página: ocorrências do turno + informativo de qualidade químico.

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
# Pontos de amostragem no MES (mesmos da planilha padrao e do Informativo do Fisico):
#   M650030010  = Pellet Feed (Mineroduto 03) - amostra unica para US3 e US4
#   M710050020  = Linha de Mistura / Pelota US3      M4710050020 = Linha de Mistura / Pelota US4
# Codigos de analise (sufixo -HHLQU = Laboratorio Quimico) CONFIRMADOS nas planilhas existentes:
#   0004 SiO2 · 0006 CaO · 0013 PPC · 0018 B2 · 0084 kg carvao/t · 0493 Carbono fixo
# As demais analises (FeT, Al2O3, MgO, P, Mn, TiO2, pH) ficam SEM TAG ate serem confirmadas no MES:
# a linha fica oculta no relatorio e a analise nao entra na consulta.
PF, LM3, LM4 = 'M650030010', 'M710050020', 'M4710050020'
MV, AM = 'IP_MESVALOR', 'IP_ANALOGMAP'
G_PF = 'Pellet Feed - Mineroduto 03'
G_LM = 'Linha de Mistura / Pelota'


def _pf(nome, un, dec, cod, tip):
    tag = '%s-%s-HHLQU' % (PF, cod) if cod else ''
    return (G_PF, nome, un, dec, 'Média', tag, MV, tag, MV, tip, tip, 'Q', True)


def _lm(nome, un, dec, cod, tip3, tip4):
    t3 = '%s-%s-HHLQU' % (LM3, cod) if cod else ''
    t4 = '%s-%s-HHLQU' % (LM4, cod) if cod else ''
    return (G_LM, nome, un, dec, 'Média', t3, MV, t4, MV, tip3, tip4, 'Q', False)


# (grupo, analise, unidade, decimais, agregacao, tag US3, tipo, tag US4, tipo, tipico US3, tipico US4,
#  consulta, amostra_unica)
PARAMS = [
    _pf('FeT', '%', 2, None, 67.50),
    _pf('SiO2', '%', 2, '0004', 1.29),
    _pf('Al2O3', '%', 2, None, 0.40),
    _pf('CaO', '%', 2, '0006', 0.09),
    _pf('MgO', '%', 2, None, 0.05),
    _pf('P', '%', 3, None, 0.040),
    _pf('Mn', '%', 3, None, 0.060),
    _pf('TiO2', '%', 3, None, 0.050),
    _pf('PPC', '%', 2, '0013', 3.66),
    _pf('pH - Mineroduto', '-', 2, None, 10.5),
    _lm('FeT', '%', 2, None, 65.50, 65.30),
    _lm('SiO2', '%', 2, '0004', 1.82, 1.90),
    _lm('Al2O3', '%', 2, None, 0.42, 0.44),
    _lm('CaO', '%', 2, '0006', 0.82, 0.88),
    _lm('MgO', '%', 2, None, 0.10, 0.11),
    _lm('B2 (CaO/SiO2)', '-', 2, '0018', 0.45, 0.46),
    _lm('Carvão', 'kg/t', 1, '0084', 12.7, 17.6),
    _lm('Carbono Fixo', '%', 2, '0493', 1.07, 1.20),
    _lm('P', '%', 3, None, 0.045, 0.046),
    _lm('Mn', '%', 3, None, 0.070, 0.071),
    _lm('TiO2', '%', 3, None, 0.055, 0.056),
]
NPARAM = len(PARAMS)
NSLOT = 6
USINAS = ('US3', 'US4')


def _idx(grupo, nome):
    return [i + 1 for i, pr in enumerate(PARAMS) if pr[0] == grupo and pr[1] == nome][0]


# indicadores do Painel: (parametro, titulo)
KPIS = [(_idx(G_PF, 'SiO2'), 'SiO2 - Pellet Feed'), (_idx(G_LM, 'SiO2'), 'SiO2 - Mistura/Pelota'),
        (_idx(G_LM, 'B2 (CaO/SiO2)'), 'B2 - Mistura/Pelota'), (_idx(G_LM, 'Carbono Fixo'), 'Carbono Fixo')]
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
SH_OC = "'Ocorrência'"


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
        self.wb.set_properties({'title': 'Relatório de Turno - Laboratório Químico - US3/US4',
                                'company': 'Samarco - Laboratório Químico'})
        self._fmts = {}
        self.linhas_inf = {}   # (p, k) -> linha (1-based) na aba Informativo

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
        """Faixa Azul Samarco com o logo (versao para fundo escuro) e filete amarelo,
        como no kit visual Samarco."""
        ws.set_row(r0, 6)
        ws.set_row(r0 + 1, 26)
        ws.set_row(r0 + 2, 20)
        ws.set_row(r0 + 3, 4)
        faixa = self.f(bg_color=AZUL)
        for r in (r0 + 1, r0 + 2):
            for c in range(col_ini, col_logo_fim + 1):
                ws.write_blank(r, c, None, faixa)
        ws.merge_range(r0 + 1, col_logo_fim + 1, r0 + 1, ultima_col, titulo,
                       self.f(bold=True, font_size=15, font_color=BRANCO, bg_color=AZUL, indent=1, valign='bottom'))
        ws.merge_range(r0 + 2, col_logo_fim + 1, r0 + 2, ultima_col, subtitulo,
                       self.f(font_size=9, font_color='#CFE0EA', bg_color=AZUL, indent=1, valign='top'))
        for c in range(col_ini, ultima_col + 1):
            ws.write_blank(r0 + 3, c, None, self.f(bg_color=AMARELO))
        ws.insert_image(r0 + 1, col_ini, LOGO_ESCURO, {'x_scale': 0.16, 'y_scale': 0.16, 'x_offset': 10,
                                                  'y_offset': 6, 'object_position': 3,
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
        self.ws_painel = wb.add_worksheet('Painel')
        self.ws_rel_d = wb.add_worksheet('Relatório Dia')
        self.ws_rel_n = wb.add_worksheet('Relatório Noite')
        self.ws_oc = wb.add_worksheet('Ocorrência')
        self.ws_inf = wb.add_worksheet('Informativo')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_mapa = wb.add_worksheet('_Mapa')
        for ws, cn in [(self.ws_painel, 'shPainel'), (self.ws_rel_d, 'shRelDia'), (self.ws_rel_n, 'shRelNoite'),
                       (self.ws_oc, 'shOcorrencia'), (self.ws_inf, 'shInformativo'),
                       (self.ws_cfg, 'shConfig'), (self.ws_mes, 'shDadosMES'), (self.ws_mapa, 'shMapa')]:
            ws.set_vba_name(cn)
            ws.hide_gridlines(2)
        self.ws_painel.set_tab_color(AZUL)
        self.ws_rel_d.set_tab_color(AMARELO)
        self.ws_rel_n.set_tab_color(AMARELO)
        self.ws_oc.set_tab_color(OURO)
        self.ws_inf.set_tab_color(AZUL_MEDIO)
        self.ws_cfg.set_tab_color(CINZA)

        self.aba_config()
        self.aba_dados_mes()
        self.aba_informativo()
        self.aba_ocorrencia()
        self.aba_relatorio(self.ws_rel_d, 'rlD', 'Turno Dia  ·  07h às 19h', 'Dia', 7)
        self.aba_relatorio(self.ws_rel_n, 'rlN', 'Turno Noite  ·  19h às 07h', 'Noite', 19)
        self.aba_painel()
        self.aba_mapa()
        self.ws_mes.hide()
        self.ws_cfg.hide()
        self.ws_mapa.very_hidden() if hasattr(self.ws_mapa, 'very_hidden') else self.ws_mapa.hide()
        self.ws_painel.activate()

    # ------------------------------------------------------------ Configurações
    def aba_config(self):
        ws = self.ws_cfg
        larg = {0: 2, 1: 16, 2: 34, 3: 9, 4: 6, 5: 13, 6: 25, 7: 15, 8: 25, 9: 15,
                10: 8, 11: 8, 12: 9, 13: 9, 14: 9, 15: 9, 16: 2, 17: 2}
        for c, w in larg.items():
            ws.set_column(c, c, w)
        ws.set_column(18, 24, 17)
        self.cabecalho(ws, 15, 'Configurações', 'Parâmetros gerais, tags do MES e limites de especificação')
        self.secao(ws, 4, 1, 15, 'Parâmetros gerais')
        gerais = [
            ('cfgFonte', 'Fonte dos dados', 'MES', 'MES = busca no Aspen/IP.21  ·  SIMULAÇÃO = dados fictícios (treinamento e teste)'),
            ('cfgServidor', 'Servidor do MES', 'UBU', 'Fonte de dados do suplemento Aspen'),
            ('cfgInicioDia', 'Início do turno Dia', dtm.time(7, 0), 'Turno Noite começa 12 h depois (Dia 07h-19h · Noite 19h-07h)'),
            ('cfgTimeout', 'Espera máxima pelo MES (s)', 60, 'Tempo de espera pela resposta das fórmulas do Aspen'),
            ('cfgFormatoData', 'Formato da data para o MES', 'Texto dd/mm/aaaa',
             'Texto dd/mm/aaaa (padrão, validado no MES) · Data do Excel · Texto mm/dd/aaaa'),
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
        self.botao(ws, 5, 11, 'Voltar ao Painel', 'IrPainel', 230, 30, 'primario', x=5)
        self.botao(ws, 7, 11, 'Ver dados brutos do MES', 'MostrarDadosMES', 230, 30, 'claro', x=5)

        ws.merge_range(14, 1, 14, 15,
                       'Análise SEM TAG = fica fora da consulta ao MES e oculta no relatório. Preencha a tag '
                       '(ex.: M710050020-00xx-HHLQU) quando confirmada no MES. Pellet Feed: amostra única, a tag US4 '
                       'repete a US3. LIE/LSE = especificação (vazio = sem avaliação).',
                       self.f(font_size=9, italic=True, font_color=AZUL_ACINZ, text_wrap=True, indent=1))
        ws.set_row(14, 28)
        self.secao(ws, 15, 1, 15, 'Tags do MES e limites - Análises químicas US3 e US4')
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
        ws.write(CFG_ROW1 - 1 + NPARAM + 1, 2, 'Tags em laranja: análise ainda sem tag no MES (confirmar com a TI / MES).',
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
        ws.merge_range(4, 18, 4, 27, 'LISTAS (usadas nas caixas de seleção)',
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
        self.botao(ws, 0, 4, 'Voltar às Configurações', 'OcultarDadosMES', 200, 28, 'primario', y=2)

    # ------------------------------------------------------------ Informativo
    def aba_informativo(self):
        """Informativo completo: resultados de 2 em 2 h de qualquer periodo (ate 24 h).
        O botao 1 do Painel carrega o turno; aqui o tecnico tambem pode escolher inicio e fim."""
        ws = self.ws_inf
        CS = INF_COL_H1 - 1            # 1a janela (0-based)
        CR = INF_COL_RES - 1           # resultado
        CMIN, CMAX, CLIE, CLSE, CST = CR + 1, CR + 2, CR + 3, CR + 4, CR + 5
        CB = CST + 2                   # botoes
        ws.set_column(0, 0, 3, None, {'hidden': True})
        ws.set_column(1, 1, 32)
        ws.set_column(2, 2, 7)
        ws.set_column(3, 3, 6)
        ws.set_column(CS, CS + NSLOT_INF - 1, 7.2)
        ws.set_column(CR, CR, 10)
        ws.set_column(CMIN, CMAX, 7.5)
        # LIE, LSE e Status ficam ocultos (nao vao para o e-mail); o Status continua marcando "Fora"
        ws.set_column(CLIE, CST, 7, None, {'hidden': True})
        ws.set_column(CST + 1, CST + 1, 3)
        ws.set_column(CB, CB, 27)
        self.cabecalho(ws, CST, 'Informativo de Qualidade - Usinas 3 e 4',
                       'Resultados do MES de 2 em 2 horas  ·  Laboratório Químico')
        lab = self.f(font_size=8, font_color=TEXTO_SEC, indent=1)
        val = self.f(font_size=11, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1, border=1,
                     border_color=BORDA)
        info = [((1, 1), 'Data', 'iData', 'dd/mm/yyyy'), ((2, 6), 'Turno / período', 'iTurno', None),
                ((7, 8), 'Turma', 'iTurma', None), ((9, 12), 'Responsável', 'iResp', None),
                ((13, 15), 'Atualizado em', 'iAtualizado', 'dd/mm/yyyy hh:mm'), ((CR, CST), 'Fonte', 'iFonte', None)]
        ws.set_row(5, 15)
        ws.set_row(6, 22)
        for (c1, c2), t, nm, nf in info:
            fv = self.f(font_size=11, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1, border=1,
                        border_color=BORDA, num_format=nf) if nf else val
            if c1 == c2:
                ws.write(5, c1, t, lab)
                ws.write_blank(6, c1, None, fv)
            else:
                ws.merge_range(5, c1, 5, c2, t, lab)
                ws.merge_range(6, c1, 6, c2, '', fv)
            self.nome(nm, ws, 6, c1)

        # escolha do periodo
        ws.set_row(7, 8)
        ws.set_row(8, 26)
        sel = self.f(font_size=11, bold=True, font_color=AZUL_TITULO, num_format='dd/mm/yyyy hh:mm', align='center',
                     bg_color=FUNDO_INPUT, border=1, border_color=AMARELO, locked=False)
        ws.write(8, 1, 'Período para consultar (até 24 h):', self.f(bold=True, font_size=10, font_color=AZUL_TITULO,
                                                                      indent=1))
        ws.merge_range(8, 2, 8, 3, 'Início', self.f(font_size=9, font_color=TEXTO_SEC, align='right'))
        ws.merge_range(8, CS, 8, CS + 2, '', sel)
        ws.write(8, CS + 3, 'Fim', self.f(font_size=9, font_color=TEXTO_SEC, align='center'))
        ws.merge_range(8, CS + 4, 8, CS + 6, '', sel)
        self.nome('iSelIni', ws, 8, CS)
        self.nome('iSelFim', ws, 8, CS + 4)
        for c in (CS, CS + 4):
            ws.data_validation(8, c, 8, c, {'validate': 'date', 'criteria': '>', 'value': dtm.date(2020, 1, 1),
                                            'error_message': 'Digite data e hora: dd/mm/aaaa hh:mm'})
        ws.merge_range(8, CS + 7, 8, CST, 'Ex.: 28/09/2026 07:00  →  29/09/2026 07:00. '
                       'Depois clique em Atualizar dados do MES.',
                       self.f(font_size=8, font_color=TEXTO_SEC, indent=1, text_wrap=True))
        ws.set_row(9, 16)
        ws.write(9, 1, 'Período consultado:', lab)
        ws.merge_range(9, 2, 9, CST, '', self.f(font_size=9, font_color=TEXTO_SEC))
        self.nome('iPeriodo', ws, 9, 2)

        hdr = self.f(bold=True, font_size=9, font_color=BRANCO, bg_color=AZUL, align='center', text_wrap=True,
                     border=1, border_color=AZUL)
        R0 = 11
        ws.set_row(R0, 30)
        cabs = ['Parâmetro', 'Unid.', 'Usina'] + ['—'] * NSLOT_INF + ['Resultado do período', 'Mín.', 'Máx.',
                                                                       'LIE', 'LSE', 'Status']
        for i, t in enumerate(cabs):
            ws.write(R0, 1 + i, t, hdr)
        self.nome('iHoras', ws, R0, CS, R0, CS + NSLOT_INF - 1)

        r = R0 + 1
        grupo_atual = None
        gi = 0
        chave_fmt = self.f(font_color=BRANCO, font_size=6)
        for p, pr in enumerate(PARAMS, start=1):
            g, nome, un, dec, agg = pr[:5]
            unica = pr[12]
            if g != grupo_atual:
                grupo_atual = g
                gi += 1
                ws.set_row(r, 17)
                ws.write(r, 0, 'G%d' % gi, chave_fmt)
                ws.merge_range(r, 1, r, CST, g, self.f(bold=True, font_size=9, font_color=AZUL_TITULO,
                                                       bg_color=FUNDO_GRUPO, indent=1, border=1, border_color=BORDA))
                r += 1
            nf = fmt_dec(dec)
            zebra = BRANCO if p % 2 else '#F8F9FA'
            borda = dict(border=1, border_color=BORDA)
            nlin = 1 if unica else 2
            fnome = self.f(font_size=9, bg_color=zebra, indent=1, text_wrap=True, **borda)
            fun = self.f(font_size=8, font_color=TEXTO_SEC, bg_color=zebra, align='center', **borda)
            if nlin == 2:
                ws.merge_range(r, 1, r + 1, 1, nome, fnome)
                ws.merge_range(r, 2, r + 1, 2, un, fun)
            else:
                ws.write(r, 1, nome, fnome)
                ws.write(r, 2, un, fun)
            for k in range(nlin):
                us = 'MD03' if unica else USINAS[k]
                rr = r + k
                ws.set_row(rr, 15)
                self.linhas_inf[(p, k + 1)] = rr + 1
                if unica:
                    self.linhas_inf[(p, 2)] = rr + 1     # amostra unica: US4 = mesma linha
                ws.write(rr, 0, 'P%02dU%d' % (p, 3 + k), chave_fmt)
                ws.write(rr, 3, us, self.f(bold=True, font_size=8, font_color=AZUL if k == 0 else AZUL_ACINZ,
                                           bg_color=zebra, align='center', **borda))
                vals = '%s:%s' % (rc(rr, CS), rc(rr, CS + NSLOT_INF - 1))
                for s_ in range(NSLOT_INF):
                    ws.write_blank(rr, CS + s_, None,
                                   self.f(font_size=9, num_format=nf, align='center', bg_color=zebra, locked=False,
                                          **borda))
                agg_ref = cfg_ref(p, 'AGG')
                res = ('=IFERROR(IF(COUNT({v})=0,"",IF({a}="Soma",SUM({v}),IF({a}="Último valor",'
                       'LOOKUP(2,1/({v}<>""),{v}),AVERAGE({v})))),"")').format(v=vals, a=agg_ref)
                ws.write_formula(rr, CR, res, self.f(num_format=nf, bold=True, align='center', font_size=10,
                                                     font_color=AZUL_TITULO, bg_color=FUNDO_AZUL_CLARO, **borda))
                ws.write_formula(rr, CMIN, '=IF(COUNT({v})=0,"",MIN({v}))'.format(v=vals),
                                 self.f(font_size=9, num_format=nf, align='center', font_color=AZUL_ACINZ,
                                        bg_color=zebra, **borda))
                ws.write_formula(rr, CMAX, '=IF(COUNT({v})=0,"",MAX({v}))'.format(v=vals),
                                 self.f(font_size=9, num_format=nf, align='center', font_color=AZUL_ACINZ,
                                        bg_color=zebra, **borda))
                for c, key in ((CLIE, 'LIE'), (CLSE, 'LSE')):
                    ref = cfg_ref(p, key)
                    ws.write_formula(rr, c, '=IF(%s="","",%s)' % (ref, ref),
                                     self.f(num_format=nf, align='center', font_color=AZUL_ACINZ, font_size=9,
                                            bg_color=zebra, **borda))
                K, N, O = rc(rr, CR), rc(rr, CLIE), rc(rr, CLSE)
                ws.write_formula(rr, CST, ('=IF(OR({k}="",AND({n}="",{o}="")),"",IF(OR(AND({n}<>"",{k}<{n}),'
                                           'AND({o}<>"",{k}>{o})),"Fora","OK"))').format(k=K, n=N, o=O),
                                 self.f(bold=True, align='center', font_size=9, bg_color=zebra, **borda))
            r += nlin
        ultima = r - 1
        # valores fora da especificacao em laranja (janelas e resultado)
        rng_vals = '%s:%s' % (rc(R0 + 1, CS), rc(ultima, CR))
        prim = rc(R0 + 1, CS, False, False)
        n_abs = '$%s%d' % (colname(CLIE), R0 + 2)
        o_abs = '$%s%d' % (colname(CLSE), R0 + 2)
        fora = 'AND(ISNUMBER({c}),OR(AND(ISNUMBER({n}),{c}<{n}),AND(ISNUMBER({o}),{c}>{o})))'.format(
            c=prim, n=n_abs, o=o_abs)
        ws.conditional_format(rng_vals, {'type': 'formula', 'criteria': '=' + fora,
                                         'format': self.wb.add_format({'font_color': LARANJA, 'bold': True,
                                                                       'bg_color': LARANJA_FUNDO})})

        r = ultima + 1
        ws.set_row(r, 22)
        ws.merge_range(r, 1, r, CST,
                       'Resultado do período: média das janelas de 2 h. Valores em laranja: fora da especificação '
                       '(LIE/LSE). Valores fora da faixa válida são descartados. Análises sem tag no MES ficam ocultas.',
                       self.f(font_size=8, font_color=TEXTO_SEC, italic=True, text_wrap=True, indent=1))
        self.nome('iAreaImagem', ws, 1, 1, r, CST)
        ws.print_area(1, 1, r, CST)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.center_horizontally()
        ws.freeze_panes(R0 + 1, 0)
        # botoes (fora da area da imagem)
        self.botao(ws, 1, CB, 'Voltar ao Painel', 'IrPainel', 195, 28, 'claro', x=4, y=4)
        self.botao(ws, 5, CB, 'Atualizar dados do MES\n(período escolhido)', 'AtualizarMESPeriodo', 195, 42,
                   'primario', x=4)
        self.botao(ws, 8, CB, 'Copiar como imagem', 'CopiarImagem', 195, 30, 'destaque', x=4, y=2)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Ocorrência (entrada de dados)
    def aba_ocorrencia(self):
        """Ocorrencia do turno, no mesmo formato usado hoje pelo Laboratorio Quimico:
        tarefas realizadas, solicitacoes, equipamentos, tarefas a realizar e controle do laboratorio.
        Um tecnico por letra (sem laboratoristas)."""
        ws = self.ws_oc
        UC = 10
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 5)
        ws.set_column(2, UC, 10.3)
        ws.set_column(UC + 1, UC + 1, 3)
        ws.set_column(UC + 2, UC + 2, 27)
        self.cabecalho(ws, UC, 'Ocorrência do Turno - Laboratório Químico',
                       'Preencha os campos amarelos  ·  o relatório do turno é montado a partir daqui',
                       col_logo_fim=3)
        ws.set_row(4, 6)
        ws.set_row(5, 22)
        ws.merge_range(5, 1, 5, UC, '', self.f(bold=True, font_size=10, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO,
                                               indent=1, border=1, border_color=BORDA))
        ws.write_formula(5, 1, '=IF(pData="","Selecione o turno no Painel",TEXT(DAY(pData),"00")&"/"&'
                               'TEXT(MONTH(pData),"00")&"/"&YEAR(pData)&"   ·   "&pTurno&"   ·   Letra "&pTurma&'
                               '"   ·   "&pResp)',
                         self.f(bold=True, font_size=10, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                                border=1, border_color=BORDA))
        self.campos_oc = []     # (nome, tipo, padrao) -> aba _Mapa (limpeza e leitura)
        num = self.f(font_size=8, font_color=TEXTO_SEC, align='center', bg_color=FUNDO_CLARO, border=1,
                     border_color=BORDA)
        inp = self.f_input(indent=1, font_size=10)
        rot = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                     border=1, border_color=BORDA, text_wrap=True)

        def lista(r, titulo, nome, n, dica):
            self.secao(ws, r, 1, UC - 4, titulo)
            ws.merge_range(r, UC - 3, r, UC, dica, self.f(font_size=8, italic=True, font_color=TEXTO_SEC,
                                                          align='right', valign='bottom', bottom=2, bottom_color=AZUL))
            for k in range(n):
                rr = r + 1 + k
                ws.set_row(rr, 19)
                ws.write(rr, 1, k + 1, num)
                ws.merge_range(rr, 2, rr, UC, '', inp)
            self.nome(nome, ws, r + 1, 2, r + n, 2)
            self.campos_oc.append((nome, 'lista', ''))
            return r + n + 2

        def campo(rr, c_rot1, c_rot2, texto, c_val1, c_val2, nome, padrao='', validar=None, fmt=None):
            fmt = fmt or inp
            if c_rot1 == c_rot2:
                ws.write(rr, c_rot1, texto, rot)
            else:
                ws.merge_range(rr, c_rot1, rr, c_rot2, texto, rot)
            if c_val1 == c_val2:
                ws.write(rr, c_val1, padrao, fmt)
            else:
                ws.merge_range(rr, c_val1, rr, c_val2, padrao, fmt)
            if validar:
                ws.data_validation(rr, c_val1, rr, c_val1, {'validate': 'list', 'source': '=' + validar})
            self.nome(nome, ws, rr, c_val1)
            self.campos_oc.append((nome, 'campo', padrao))

        r = 7
        r = lista(r, '1. Tarefas realizadas', 'ocReal', 12, 'uma tarefa por linha')
        r = lista(r, '2. Solicitações', 'ocSol', 4, 'nº da solicitação, amostra, pendência')
        r = lista(r, '3. Equipamentos', 'ocEquip', 4, 'falhas, vazamentos, manutenção')
        r = lista(r, '4. Tarefas a realizar (próximo turno)', 'ocAReal', 8, 'uma tarefa por linha')

        self.secao(ws, r, 1, UC, '5. Controle do laboratório')
        prog = ('OREGON (RX), Carbono (Leco CS-230), PCS (Calorímetro), mufla 1000 ºC, fotômetro, balanças, '
                'estufa e máquina de fusão')
        rr = r + 1
        ws.set_row(rr, 32)
        campo(rr, 1, 3, 'Programas / equipamentos em uso', 4, UC, 'ocProg', prog,
              fmt=self.f_input(indent=1, font_size=9, text_wrap=True))
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Foi necessário preparar padrões?', 4, 5, 'ocPadroes', 'Não', '=lstSN')
        campo(rr, 6, 6, 'Quais', 7, UC, 'ocPadroesQuais')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Sistema de ar utilizado', 4, 5, 'ocAr', 'Compressor', '=lstAr')
        campo(rr, 6, 6, 'Status', 7, 8, 'ocArStatus', '', '=lstLigado')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Condição do compressor', 4, 5, 'ocCompressor', '', '=lstCond')
        campo(rr, 6, 6, 'Nitrogênio', 7, 8, 'ocNitrogenio', '', '=lstCond')
        r = rr + 2

        self.secao(ws, r, 1, UC, '6. Pessoal (ausências, trocas e hora extra)')
        rr = r + 1
        for rotulo, nm in (('Houve ausência no turno?', 'ocAus'), ('Houve troca combinada?', 'ocTroca'),
                           ('Houve hora extra?', 'ocHE')):
            ws.set_row(rr, 20)
            campo(rr, 1, 3, rotulo, 4, 5, nm, 'Não', '=lstSN')
            campo(rr, 6, 6, 'Quem', 7, UC, nm + 'Quem')
            rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Letra que recebe o turno', 4, 5, 'ocRecebe', '', '=lstTurma')
        r = rr + 2

        self.secao(ws, r, 1, UC, '7. Cadinhos de platina')
        rr = r + 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Repassados para o turno', 4, UC, 'ocCadinhos')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Retirados para reforma', 4, UC, 'ocReforma')
        r = rr + 2

        self.secao(ws, r, 1, UC, '8. Observações gerais')
        rr = r + 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Hidrogênio no carvão', 4, 5, 'ocH2Carvao')
        campo(rr, 6, 7, 'Hidrogênio no coque', 8, UC, 'ocH2Coque')
        rr += 1
        ws.set_row(rr, 20)
        campo(rr, 1, 3, 'Coque Planta 04', 4, UC, 'ocCoque04')
        rr += 1
        ws.merge_range(rr, 1, rr + 2, 3, 'Observações', rot)
        ws.merge_range(rr, 4, rr + 2, UC, '', self.f_input(indent=1, text_wrap=True, valign='top'))
        self.nome('ocObs', ws, rr, 4)
        self.campos_oc.append(('ocObs', 'campo', ''))
        fim = rr + 2
        self.oc_fim = fim

        ws.print_area(1, 1, fim, UC)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        CB = UC + 2
        self.botao(ws, 1, CB, 'Voltar ao Painel', 'IrPainel', 195, 28, 'claro', x=4, y=4)
        self.botao(ws, 5, CB, 'Ver relatório do turno', 'IrRelatorio', 195, 34, 'primario', x=4)
        self.botao(ws, 8, CB, 'Limpar ocorrência', 'LimparOcorrencia', 195, 28, 'claro', x=4, y=6)
        self.botao(ws, 11, CB, 'Finalizar turno', 'FecharTurno', 195, 34, 'destaque', x=4)
        ws.merge_range(15, CB, 20, CB, 'Dica: uma tarefa por linha. Para textos longos, a linha quebra '
                       'sozinha no relatório.\nOs campos Sim/Não, sistema de ar e letras têm lista de seleção.',
                       self.f(font_size=8, font_color=TEXTO_SEC, text_wrap=True, valign='top', indent=1))
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Relatório Dia / Relatório Noite
    def aba_relatorio(self, ws, pref, sub_turno, nome_turno, hora_ini):
        """Relatorio do turno em UMA pagina (imagem para o e-mail): ocorrencias + informativo quimico.
        Os valores sao GRAVADOS pelo VBA (ModRelatorio.GerarRelatorio): uma "foto" do turno, para o
        outro turno poder atualizar o MES e a ocorrencia sem apagar este relatorio."""
        CS = 4                      # 1a janela (E)
        CR = CS + NSLOT             # media (K)
        UC = CR + 2                 # ultima coluna da imagem (M)
        CB = UC + 2                 # botoes (O)
        CC = CB + 2                 # controle oculto (Q, R, S)
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 23)
        ws.set_column(2, 2, 6)
        ws.set_column(3, 3, 6)
        ws.set_column(CS, CS + NSLOT - 1, 7)
        ws.set_column(CR, CR, 9)
        ws.set_column(CR + 1, UC, 7)
        ws.set_column(UC + 1, UC + 1, 3)
        ws.set_column(CB, CB, 27)
        ws.set_column(CC, CC + 2, 16, None, {'hidden': True})
        chave_fmt = self.f(font_color=BRANCO, font_size=6)
        vazio = 'Relatório ainda não gerado  ·  use o botão 3 do Painel (Relatório do turno)'

        for i, (t, nm) in enumerate([('Gerado em', 'Gerado'), ('Finalizado em', 'Final'),
                                     ('Resultados do MES em', 'MES')]):
            ws.write(0, CC + i, t)
            ws.write_blank(1, CC + i, None, self.f(num_format='dd/mm/yyyy hh:mm'))
            self.nome(pref + nm, ws, 1, CC + i)

        self.cabecalho(ws, UC, 'Relatório de Turno - Laboratório Químico', 'Usinas 3 e 4  ·  ' + sub_turno)
        ws.set_row(4, 6)
        ws.set_row(5, 22)
        ws.merge_range(5, 1, 5, UC, vazio, self.f(bold=True, font_size=10, font_color=AZUL_TITULO,
                                                  bg_color=FUNDO_CLARO, indent=1, border=1, border_color=BORDA))
        self.nome(pref + 'Info', ws, 5, 1)
        ws.set_row(6, 6)

        # ---- ocorrencias
        self.secao(ws, 7, 1, UC, 'Ocorrências do turno')
        rot = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                     border=1, border_color=BORDA, text_wrap=True, valign='top')
        val = self.f(font_size=9, text_wrap=True, indent=1, border=1, border_color=BORDA, valign='top')
        itens = [('Tarefas realizadas', 'Real'), ('Solicitações', 'Sol'), ('Equipamentos', 'Equip'),
                 ('Tarefas a realizar', 'AReal'), ('Controle do laboratório', 'Lab'), ('Pessoal', 'Pessoal'),
                 ('Cadinhos de platina', 'Cad'), ('Observações', 'Obs')]
        r = 8
        self.rel_textos = []
        for t, nm in itens:
            ws.set_row(r, 18)
            ws.write(r, 1, t, rot)
            ws.merge_range(r, 2, r, UC, '-', val)
            self.nome(pref + nm, ws, r, 2)
            self.rel_textos.append(nm)
            r += 1

        # ---- informativo de qualidade quimico
        r += 1
        ws.set_row(r - 1, 8)
        self.secao(ws, r, 1, UC, 'Informativo de qualidade do turno  ·  resultados do MES de 2 em 2 horas')
        r += 1
        hdr = self.f(bold=True, font_size=8, font_color=BRANCO, bg_color=AZUL, align='center', text_wrap=True,
                     border=1, border_color=BRANCO)
        ws.set_row(r, 26)
        horas = ['%02dh-%02dh' % ((hora_ini + 2 * i) % 24, (hora_ini + 2 * i + 2) % 24) for i in range(NSLOT)]
        cabs = ['Amostra / análise', 'Un.', 'Usina'] + horas + ['Média do turno', 'Mín.', 'Máx.']
        for i, t in enumerate(cabs):
            ws.write(r, 1 + i, t, hdr)
        self.nome(pref + 'Horas', ws, r, CS, r, CS + NSLOT - 1)
        r += 1
        self.rel_tab_ini = r + 1    # 1-based
        grupo = None
        gi = 0
        for p, pr in enumerate(PARAMS, start=1):
            g, nome, un, dec = pr[:4]
            unica = pr[12]
            if g != grupo:
                grupo = g
                gi += 1
                ws.set_row(r, 15)
                ws.write(r, 0, 'G%d' % gi, chave_fmt)
                ws.merge_range(r, 1, r, UC, g, self.f(bold=True, font_size=9, font_color=AZUL_TITULO,
                                                      bg_color=FUNDO_GRUPO, indent=1, border=1, border_color=BORDA))
                r += 1
            zebra = BRANCO if p % 2 else '#F8F9FA'
            b = dict(border=1, border_color=BORDA, bg_color=zebra)
            nf = fmt_dec(dec)
            nlin = 1 if unica else 2
            fnome = self.f(font_size=9, indent=1, **b)
            fun = self.f(font_size=8, font_color=TEXTO_SEC, align='center', **b)
            if nlin == 2:
                ws.merge_range(r, 1, r + 1, 1, nome, fnome)
                ws.merge_range(r, 2, r + 1, 2, un, fun)
            else:
                ws.write(r, 1, nome, fnome)
                ws.write(r, 2, un, fun)
            for k in range(nlin):
                rr = r + k
                ws.set_row(rr, 14)
                ws.write(rr, 0, 'P%02dU%d' % (p, 3 + k), chave_fmt)
                ws.write(rr, 3, 'MD03' if unica else USINAS[k],
                         self.f(bold=True, font_size=8, font_color=AZUL if k == 0 else AZUL_ACINZ, align='center', **b))
                for c in range(CS, UC + 1):
                    if c == CR:
                        fmt = self.f(font_size=9, bold=True, font_color=AZUL_TITULO, align='center',
                                     num_format=nf, **dict(b, bg_color=FUNDO_AZUL_CLARO))
                    elif c > CR:
                        fmt = self.f(font_size=8, font_color=AZUL_ACINZ, align='center', num_format=nf, **b)
                    else:
                        fmt = self.f(font_size=9, font_color=TEXTO, align='center', num_format=nf, **b)
                    ws.write(rr, c, '-', fmt)
            r += nlin
        ws.set_row(r, 16)
        ws.merge_range(r, 1, r, UC, 'Média do turno = média das janelas de 2 h. Valores em laranja: fora da '
                       'especificação. MD03 = Pellet Feed do Mineroduto 03 (amostra única para US3 e US4).',
                       self.f(font_size=7.5, font_color=TEXTO_SEC, italic=True, indent=1))
        fim = r
        self.rel_fim = fim + 1      # 1-based
        self.nome(pref + 'Area', ws, 1, 1, fim, UC)

        ws.print_area(1, 1, fim, UC)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.35, 0.35, 0.45, 0.45)
        ws.center_horizontally()

        T = nome_turno
        self.botao(ws, 1, CB, 'Voltar ao Painel', 'IrPainel', 195, 28, 'claro', x=4, y=4)
        self.botao(ws, 5, CB, 'Copiar relatório\n(imagem para o e-mail)', 'CopiarRelatorio' + T, 195, 44,
                   'destaque', x=4)
        self.botao(ws, 9, CB, 'Atualizar este relatório', 'AtualizarRelatorio' + T, 195, 28, 'claro', x=4, y=4)
        ws.merge_range(12, CB, 17, CB, 'Este relatório guarda uma "foto" do turno %s. Ele é refeito ao '
                       'atualizar o MES, pelo botão 3 do Painel e antes de copiar, enquanto o turno não for '
                       'finalizado.\nPara print: Windows + Shift + S.' % T,
                       self.f(font_size=8, font_color=TEXTO_SEC, text_wrap=True, valign='top', indent=1))
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Painel
    def aba_painel(self):
        ws = self.ws_painel
        ws.set_column(0, 0, 2)
        ws.set_column(1, 16, 9.3)
        ws.set_column(17, 17, 2)
        self.cabecalho(ws, 16, 'Relatório de Turno - Laboratório Químico',
                       'Usinas 3 e 4  ·  Ocorrências e resultados químicos do MES', col_logo_fim=3)
        lab = self.f(font_size=8, font_color=TEXTO_SEC, indent=1)
        self.secao(ws, 5, 1, 16, 'Turno selecionado')
        ws.merge_range(6, 1, 6, 3, 'Data do turno', lab)
        ws.merge_range(6, 4, 6, 7, 'Turno', lab)
        ws.merge_range(6, 8, 6, 9, 'Letra', lab)
        ws.merge_range(6, 10, 6, 16, 'Técnico', lab)
        ws.set_row(7, 28)
        big = dict(font_size=13, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_INPUT, border=1,
                   border_color=BORDA, locked=False)
        ws.merge_range(7, 1, 7, 3, '', self.f(num_format='dd/mm/yyyy', align='center', **big))
        ws.merge_range(7, 4, 7, 7, '', self.f(align='center', **big))
        ws.merge_range(7, 8, 7, 9, '', self.f(align='center', **big))
        ws.merge_range(7, 10, 7, 16, '', self.f(indent=1, **big))
        self.nome('pData', ws, 7, 1)
        self.nome('pTurno', ws, 7, 4)
        self.nome('pTurma', ws, 7, 8)
        self.nome('pResp', ws, 7, 10)
        ws.data_validation(7, 1, 7, 1, {'validate': 'date', 'criteria': '>',
                                        'value': dtm.date(2020, 1, 1),
                                        'error_message': 'Digite uma data válida (dd/mm/aaaa).'})
        ws.data_validation(7, 4, 7, 4, {'validate': 'list', 'source': '=lstTurno'})
        ws.data_validation(7, 8, 7, 8, {'validate': 'list', 'source': '=lstTurma'})
        ws.set_row(8, 26)
        self.botao(ws, 8, 1, 'Usar turno atual', 'BtnTurnoAtual', 150, 22, 'claro', y=3, tamanho=9)
        self.botao(ws, 8, 4, 'Editar turno Dia', 'SelecionarDia', 140, 22, 'claro', y=3, tamanho=9)
        self.botao(ws, 8, 6, 'Editar turno Noite', 'SelecionarNoite', 140, 22, 'claro', x=12, y=3, tamanho=9)
        ws.merge_range(8, 9, 8, 16, 'Dia: 07h às 19h   ·   Noite: 19h às 07h do dia seguinte   ·   '
                                    'Letras A, B, C e D',
                       self.f(font_size=8, font_color=TEXTO_SEC, indent=1))

        self.secao(ws, 10, 1, 16, 'Ações do turno')
        for r in range(11, 15):
            ws.set_row(r, 18)
        bw, bh = 262, 62
        acoes = [(1, '1 · Atualizar dados do MES\nResultados químicos US3 e US4', 'AtualizarMES', 'primario'),
                 (5, '2 · Ocorrência do turno\nTarefas, solicitações e equipamentos', 'IrOcorrencia', 'primario'),
                 (9, '3 · Relatório do turno\nUma página para o e-mail', 'IrRelatorio', 'primario'),
                 (13, '4 · Finalizar turno\nFixa o relatório e salva o arquivo', 'FecharTurno', 'destaque')]
        for col, txt, mac, est in acoes:
            self.botao(ws, 11, col, txt, mac, bw, bh, est, x=4, y=6)

        self.secao(ws, 16, 1, 16, 'Situação do turno')
        tl = self.f(font_size=8, font_color=TEXTO_SEC, bg_color=FUNDO_CLARO, indent=1,
                    top=1, left=1, right=1, border_color=BORDA)
        tv = dict(font_size=13, bold=True, font_color=AZUL_TITULO, bg_color=FUNDO_CLARO, indent=1,
                  bottom=1, left=1, right=1, border_color=BORDA)
        ws.set_row(18, 28)
        d = 'iData'
        tiles = [
            (1, 4, 'Informativo carregado (MES)',
             '=IF(%s="","— não carregado —",TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)&"  ·  "&IFERROR(LEFT(iTurno,FIND(" (",iTurno)-1),iTurno))'
             % (d, d, d, d), self.f(**tv)),
            (5, 8, 'Itens lançados na ocorrência', '=COUNTA(ocReal)+COUNTA(ocSol)+COUNTA(ocEquip)+COUNTA(ocAReal)',
             self.f(num_format='0', **tv)),
            (9, 12, 'Relatório Dia', self._formula_resumo('rlD'), self.f(**tv)),
            (13, 16, 'Relatório Noite', self._formula_resumo('rlN'), self.f(**tv)),
        ]
        for c1, c2, t, fm, fv in tiles:
            ws.merge_range(17, c1, 17, c2, t, tl)
            ws.merge_range(18, c1, 18, c2, '', fv)
            ws.write_formula(18, c1, fm, fv)
        for c in (9, 13):
            ws.conditional_format(18, c, 18, c, {'type': 'text', 'criteria': 'begins with', 'value': 'Finalizado',
                                                 'format': self.wb.add_format({'font_color': VERDE_TXT,
                                                                               'bg_color': VERDE_FUNDO})})
        ws.conditional_format(18, 5, 18, 5, {'type': 'cell', 'criteria': '>=', 'value': 3,
                                             'format': self.wb.add_format({'font_color': VERDE_TXT,
                                                                           'bg_color': VERDE_FUNDO})})

        self.secao(ws, 20, 1, 16, 'Indicadores do Informativo carregado')
        for i, (p, titulo) in enumerate(KPIS):
            c = 1 + i * 4
            un = PARAMS[p - 1][2]
            nf = fmt_dec(PARAMS[p - 1][3])
            ws.merge_range(21, c, 21, c + 3, '%s (%s)' % (titulo, un),
                           self.f(bold=True, font_size=9, font_color=AZUL_TITULO, bg_color=FUNDO_GRUPO, indent=1,
                                  border=1, border_color=BORDA))
            unica = PARAMS[p - 1][12]
            for k, us in enumerate(USINAS):
                r = 22 + k
                ws.set_row(r, 24)
                celula = dict(bg_color=BRANCO, border=1, border_color=BORDA)
                if unica and k == 1:
                    ws.write_blank(r, c, None, self.f(**celula))
                    ws.merge_range(r, c + 1, r, c + 2, '', self.f(**celula))
                    ws.write_blank(r, c + 3, None, self.f(**celula))
                    continue
                if unica:
                    us = 'MD03'
                chave = 'P%02dU%d' % (p, 3 + k)
                ws.write(r, c, us, self.f(bold=True, font_size=9, font_color=AZUL_ACINZ, align='center', **celula))
                ws.merge_range(r, c + 1, r, c + 2, '', self.f(num_format=nf, font_size=14, bold=True,
                                                              font_color=AZUL_TITULO, align='center', **celula))
                ws.write_formula(r, c + 1, '=IFERROR(INDEX(Informativo!$%s:$%s,MATCH("%s",Informativo!$A:$A,0)),"")'
                                 % (INF_L_RES, INF_L_RES, chave), self.f(num_format=nf, font_size=14, bold=True, font_color=AZUL_TITULO,
                                                 align='center', **celula))
                ws.write_formula(r, c + 3, '=IFERROR(INDEX(Informativo!$%s:$%s,MATCH("%s",Informativo!$A:$A,0)),"")'
                                 % (INF_L_ST, INF_L_ST, chave), self.f(bold=True, font_size=9, align='center', **celula))
            cf = '%s:%s' % (rc(22, c + 3), rc(23, c + 3))
            ws.conditional_format(cf, {'type': 'cell', 'criteria': '==', 'value': '"Fora"',
                                       'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': LARANJA})})
            ws.conditional_format(cf, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                       'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': VERDE_TXT})})

        self.secao(ws, 25, 1, 16, 'Mais opções')
        ws.set_row(26, 28)
        extras = [(1, 'Informativo completo (MES)', 'IrInformativo')]
        for col, txt, mac in extras:
            self.botao(ws, 26, col, txt, mac, 262, 28, 'claro', x=4, y=4)
        ws.merge_range(28, 1, 28, 16, '', self.f(font_size=8, font_color=TEXTO_SEC))
        ws.write_formula(28, 1, '="Fonte dos dados: "&cfgFonte&"   ·   Servidor MES: "&cfgServidor',
                         self.f(font_size=8, font_color=TEXTO_SEC))
        ws.protect('', {'format_columns': True, 'format_rows': True})
        ws.set_landscape()
        ws.fit_to_pages(1, 1)

    @staticmethod
    def _formula_resumo(pref):
        g, f = pref + 'Gerado', pref + 'Final'
        return ('=IF(ISNUMBER({f}),"Finalizado "&TEXT({f},"hh:mm"),IF(ISNUMBER({g}),"Gerado "&TEXT({g},"hh:mm"),'
                '"— pendente —"))').format(f=f, g=g)

    # ------------------------------------------------------------ _Mapa
    def aba_mapa(self):
        ws = self.ws_mapa
        for i, t in enumerate(['Campo da ocorrência', 'Tipo', 'Padrão']):
            ws.write(0, i, t)
        for r, linha in enumerate(self.campos_oc, start=1):
            for c, v in enumerate(linha):
                ws.write_string(r, c, v)
        for i, t in enumerate(['Param', 'Usina', 'Linha Informativo']):
            ws.write(0, 6 + i, t)
        r = 1
        for p in range(1, NPARAM + 1):
            for k in (1, 2):
                ws.write_number(r, 6, p)
                ws.write_number(r, 7, k)
                ws.write_number(r, 8, self.linhas_inf[(p, k)])
                r += 1
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
                  "Public Const HORAS_SLOT As Double = 2",
                  "Public Const CFG_ROW1 As Long = %d" % CFG_ROW1]
        for k in ['GRUPO', 'PARAM', 'UNID', 'DEC', 'AGG', 'TAG3', 'TIPO3', 'TAG4', 'TIPO4', 'LIE', 'LSE',
                  'VMIN', 'VMAX', 'TIP3', 'TIP4']:
            linhas.append("Public Const CFG_COL_%s As Long = %d" % (k, c[k]))
        linhas += ["Public Const INF_COL_H1 As Long = %d" % INF_COL_H1,
                   "Public Const INF_COL_RES As Long = %d" % INF_COL_RES,
                   "Public Const NSLOT_INF As Long = %d" % NSLOT_INF,
                   "Public Const REL_TAB_INI As Long = %d" % self.rel_tab_ini,
                   "Public Const REL_FIM As Long = %d" % self.rel_fim,
                   "' Senha de protecao das abas (vazio = sem senha)",
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
    for cn in ['shPainel', 'shRelDia', 'shRelNoite', 'shOcorrencia', 'shInformativo',
               'shConfig', 'shDadosMES', 'shMapa']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': ''})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModTurno', 'ModExportar', 'ModRelatorio']:
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
