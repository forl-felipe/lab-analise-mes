"""Gera a planilha Informativo_Qualidade_US3_US4.xlsm.

Uso:  python3 build_workbook.py [--teste]
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
SAIDA = os.path.join(RAIZ, 'Informativo_Qualidade_US3_US4%s.xlsm' % ('_TESTE' if TESTE else ''))

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
# tons de apoio (derivados da paleta)
FUNDO_INPUT = '#FFF7DC'      # ouro bem claro
FUNDO_CLARO = '#EEF3F6'      # cinza-azulado claro
FUNDO_GRUPO = '#DCE8F0'
FUNDO_AZUL_CLARO = '#E6F6FD'
VERDE_TXT, VERDE_FUNDO = '#1E7B4A', '#E2F3E8'
LARANJA_FUNDO = '#FDE6D8'
FONTE = 'Segoe UI'

# ---------------------------------------------------------------- parametros US3 / US4
# (grupo, parametro, unidade, decimais, agregacao, tag US3, tipo US3, tag US4, tipo US4, tipico US3, tipico US4)
MV, AM = 'IP_MESVALOR', 'IP_ANALOGMAP'
PARAMS = [
    ('Alimentação', 'Alimentação da Grelha', 't/h', 0, 'Média', '306TP001-FX001-R', AM, '406-FX-001', AM, 1000, 1100),
    ('Prensa de Rolos', 'Superfície Específica - Alimentação RP', 'cm²/g', 0, 'Média', 'M650060010-0017-HHLFU', MV, 'M4650060010-0017-HHLFU', MV, 2010, 2010),
    ('Prensa de Rolos', 'Superfície Específica - Saída RP', 'cm²/g', 0, 'Média', 'M650060030-0017-HHLFU', MV, 'M4650060030-0017-HHLFU', MV, 2160, 2120),
    ('Prensa de Rolos', '#325 - Saída RP', '%', 1, 'Média', 'M650060030-0016-HHLFU', MV, 'M4650060030-0016-HHLFU', MV, 90.2, 89.9),
    ('Prensa de Rolos', 'H2O - Alimentação RP', '%', 2, 'Média', 'M650060010-0001-HHLFU', MV, 'M4650060010-0001-HHLFU', MV, 11.2, 11.2),
    ('Prensa de Rolos', 'H2O - Saída RP', '%', 2, 'Média', 'M650060030-0001-HHLFU', MV, 'M4650060030-0001-HHLFU', MV, 11.0, 11.0),
    ('Pellet Feed', 'PPC - Pellet Feed', '%', 2, 'Média', 'M650030010-0013-HHLQU', MV, 'M650030010-0013-HHLQU', MV, 3.66, 3.66),
    ('Pellet Feed', 'SiO2 - Pellet Feed', '%', 2, 'Média', 'M650030010-0004-HHLQU', MV, 'M650030010-0004-HHLQU', MV, 1.29, 1.29),
    ('Pellet Feed', 'CaO - Pellet Feed', '%', 2, 'Média', 'M650030010-0006-HHLQU', MV, 'M650030010-0006-HHLQU', MV, 0.09, 0.09),
    ('Mistura', 'SiO2 - Mistura', '%', 2, 'Média', 'M710050020-0004-HHLQU', MV, 'M4710050020-0004-HHLQU', MV, 1.82, 1.90),
    ('Mistura', 'CaO - Mistura', '%', 2, 'Média', 'M710050020-0006-HHLQU', MV, 'M4710050020-0006-HHLQU', MV, 0.82, 0.88),
    ('Mistura', 'B2 - Mistura', '-', 2, 'Média', 'M710050020-0018-HHLQU', MV, 'M4710050020-0018-HHLQU', MV, 0.45, 0.46),
    ('Mistura', 'Dosagem de Carvão - Mistura', 'kg/t', 1, 'Média', 'M710050020-0084-HHLQU', MV, 'M4710050020-0084-HHLQU', MV, 12.7, 17.6),
    ('Mistura', 'Carbono Fixo - Mistura', '%', 2, 'Média', 'M710050020-0493-HHLQU', MV, 'M4710050020-0493-HHLQU', MV, 1.07, 1.20),
    ('Pelota Queimada', 'Faixa +16,0 -8,0 mm', '%', 1, 'Média', 'M710050020-0028-HHLFU', MV, 'M4710050020-0028-HHLFU', MV, 92.5, 90.9),
    ('Pelota Queimada', 'Relação Granulométrica', '-', 2, 'Média', 'M710050020-2021-HHLFU', MV, 'M4710050020-2021-HHLFU', MV, 0.73, 0.80),
    ('Pelota Queimada', 'Tamboramento', '%', 1, 'Média', 'M710050020-0029-HHLFU', MV, 'M4710050020-0029-HHLFU', MV, 94.0, 93.9),
    ('Pelota Queimada', 'Resistência à Compressão', 'kgf/pel', 0, 'Média', 'M710050020-0031-HHLFU', MV, 'M4710050020-0031-HHLFU', MV, 320, 334),
    ('Pelota Queimada', 'Compressão < 200 kgf/pel', '%', 0, 'Média', 'M710050020-0032-HHLFU', MV, 'M4710050020-0032-HHLFU', MV, 17, 16),
    ('Pelota Queimada', 'Finos -6,3 mm', '%', 1, 'Média', 'M710050020-0026-HHLFU', MV, 'M4710050020-0026-HHLFU', MV, 0.97, 1.17),
    ('Produção', 'Produção', 't', 0, 'Soma', 'M710050020-0150-HHCC', MV, 'M4710050020-0150-HHCC', MV, 700, 750),
    ('Produção', 'Ritmo', 't/dia', 0, 'Último valor', 'M710050031-0150-HHCC', MV, 'M4710050031-0150-HHCC', MV, 19000, 20600),
]
NPARAM = len(PARAMS)
NSLOT = 6
USINAS = ('US3', 'US4')
KPIS = [(18, 'Resistência à Compressão'), (17, 'Tamboramento'), (10, 'SiO2 - Mistura'), (21, 'Produção')]

# Configuracoes: colunas (1-based) da tabela de tags
CFG_ROW1 = 18
CFG_COL = dict(GRUPO=2, PARAM=3, UNID=4, DEC=5, AGG=6, TAG3=7, TIPO3=8, TAG4=9, TIPO4=10,
               LIE=11, LSE=12, VMIN=13, VMAX=14, TIP3=15, TIP4=16)
INF_COL_H1, INF_COL_RES = 5, 11
HIST_ROW_HDR, REG_ROW_HDR = 6, 6
HIST_FIXAS = ['Data', 'Turno', 'Turma', 'Responsável', 'Turma que recebe', 'Fechado em', 'Fonte dos dados']
HIST_COL_RES1 = len(HIST_FIXAS) + 1
HIST_COL_NOK = HIST_COL_RES1 + NPARAM * 2
HIST_FINAIS = ['Equip. NÃO OK', 'Ocorrência SSMA?', 'Segurança (SSMA)', 'Pendências p/ próximo turno',
               'Observações gerais', 'Arquivo PDF']

SH_CFG = "'Configurações'"
SH_PASS = "'Passagem de Turno'"
SH_HIST = "'Histórico'"


def cfg_ref(p, key, absolute=True):
    """Referencia a celula da tabela de tags (p 1-based)."""
    cell = rc(CFG_ROW1 + p - 2, CFG_COL[key] - 1, True, True)
    return '%s!%s' % (SH_CFG, cell)


def fmt_dec(d):
    return '0' if d == 0 else '0.' + '0' * d


# ============================================================================
class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_properties({'title': 'Informativo de Qualidade e Passagem de Turno - US3/US4',
                                'company': 'Samarco - Laboratório Físico'})
        self._fmts = {}
        self.mapa = []         # linhas da aba _Mapa (registro da passagem)
        self.linhas_inf = {}   # (p, k) -> linha (1-based) na aba Informativo

    # ----------------------------------------------------------- formatos
    def f(self, **kw):
        base = dict(font_name=FONTE, font_size=10, font_color=AZUL_PROFUNDO, valign='vcenter')
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self._fmts:
            self._fmts[chave] = self.wb.add_format(base)
        return self._fmts[chave]

    def f_input(self, **kw):
        base = dict(bg_color=FUNDO_INPUT, border=1, border_color=CINZA, locked=False)
        base.update(kw)
        return self.f(**base)

    # ----------------------------------------------------------- blocos visuais
    def cabecalho(self, ws, ultima_col, titulo, subtitulo, nome_logo, col_logo_fim=2, col_ini=1):
        ws.set_row(0, 8)
        ws.set_row(1, 30)
        ws.set_row(2, 22)
        ws.set_row(3, 5)
        ws.merge_range(1, col_ini, 2, col_logo_fim, 'SAMARCO',
                       self.f(bold=True, font_size=16, font_color=AZUL, align='center',
                              border=1, border_color=CINZA, bg_color=BRANCO))
        self.wb.define_name(nome_logo, '=%s!%s' % (self._qn(ws), rc(1, col_ini, True, True)))
        ws.merge_range(1, col_logo_fim + 1, 1, ultima_col, titulo,
                       self.f(bold=True, font_size=18, font_color=AZUL, indent=1))
        ws.merge_range(2, col_logo_fim + 1, 2, ultima_col, subtitulo,
                       self.f(font_size=10, font_color=AZUL_ACINZ, indent=1, italic=True))
        for c in range(col_ini, ultima_col + 1):
            ws.write_blank(3, c, None, self.f(bg_color=OURO))

    def secao(self, ws, row, c1, c2, texto, cor=AZUL):
        ws.set_row(row, 21)
        ws.merge_range(row, c1, row, c2, texto,
                       self.f(bold=True, font_size=10, font_color=BRANCO, bg_color=cor, indent=1))

    def botao(self, ws, row, col, texto, macro, largura=200, altura=40, estilo='primario',
              x=0, y=0, tamanho=None):
        estilos = {
            'primario': (AZUL, BRANCO, 12),
            'medio': (AZUL_MEDIO, BRANCO, 12),
            'titulo': (AZUL_TITULO, BRANCO, 12),
            'destaque': (OURO, AZUL_PROFUNDO, 12),
            'claro': (FUNDO_GRUPO, AZUL, 10),
            'pequeno': (AZUL, BRANCO, 9),
        }
        fundo, cor, tam = estilos[estilo]
        ws.insert_textbox(row, col, texto, {
            'width': largura, 'height': altura, 'x_offset': x, 'y_offset': y,
            'font': {'name': FONTE, 'size': tamanho or tam, 'bold': True, 'color': cor},
            'align': {'vertical': 'middle', 'horizontal': 'center'},
            'fill': {'color': fundo}, 'line': {'none': True},
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

    def mapa_add(self, tipo, secao, item, status, texto):
        self.mapa.append((tipo, secao, item, status, texto))

    # ============================================================ abas
    def construir(self):
        wb = self.wb
        self.ws_painel = wb.add_worksheet('Painel')
        self.ws_inf = wb.add_worksheet('Informativo')
        self.ws_pass = wb.add_worksheet('Passagem de Turno')
        self.ws_hist = wb.add_worksheet('Histórico')
        self.ws_reg = wb.add_worksheet('Registro Passagem')
        self.ws_tend = wb.add_worksheet('Tendências')
        self.ws_cfg = wb.add_worksheet('Configurações')
        self.ws_mes = wb.add_worksheet('Dados_MES')
        self.ws_mapa = wb.add_worksheet('_Mapa')
        for ws, cn in [(self.ws_painel, 'shPainel'), (self.ws_inf, 'shInformativo'),
                       (self.ws_pass, 'shPassagem'), (self.ws_hist, 'shHistorico'),
                       (self.ws_reg, 'shRegistro'), (self.ws_tend, 'shTendencias'),
                       (self.ws_cfg, 'shConfig'), (self.ws_mes, 'shDadosMES'), (self.ws_mapa, 'shMapa')]:
            ws.set_vba_name(cn)
            ws.hide_gridlines(2)
        self.ws_painel.set_tab_color(AZUL)
        self.ws_inf.set_tab_color(AZUL_MEDIO)
        self.ws_pass.set_tab_color(OURO)
        self.ws_hist.set_tab_color(AZUL_ACINZ)
        self.ws_reg.set_tab_color(AZUL_ACINZ)
        self.ws_tend.set_tab_color(AZUL_CLARO)
        self.ws_cfg.set_tab_color(CINZA)

        self.aba_config()
        self.aba_dados_mes()
        self.aba_informativo()
        self.aba_passagem()
        self.aba_historico()
        self.aba_registro()
        self.aba_tendencias()
        self.aba_painel()
        self.aba_mapa()
        self.ws_mes.hide()
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
        self.cabecalho(ws, 15, 'Configurações', 'Parâmetros gerais, tags do MES e limites de especificação',
                       'logoConfig')
        self.secao(ws, 4, 1, 15, 'PARÂMETROS GERAIS')
        gerais = [
            ('cfgFonte', 'Fonte dos dados', 'MES', 'MES = busca no Aspen/IP.21  ·  SIMULAÇÃO = dados fictícios (treinamento e teste)'),
            ('cfgServidor', 'Servidor do MES', 'UBU', 'Nome do servidor IP.21 usado nas fórmulas do suplemento Aspen'),
            ('cfgInicioDia', 'Início do turno Dia', dtm.time(7, 0), 'Turno Noite começa 12 h depois (Dia 07h-19h · Noite 19h-07h)'),
            ('cfgTimeout', 'Espera máxima pelo MES (s)', 60, 'Tempo de espera pela resposta das fórmulas do Aspen'),
            ('cfgPastaPDF', 'Pasta dos PDFs', '', 'Vazio = pasta "Informativos PDF" ao lado desta planilha'),
            ('cfgEmail', 'Criar e-mail ao fechar turno', 'Não', 'Sim = abre um e-mail no Outlook com resumo e PDF anexado'),
            ('cfgEmailPara', 'E-mail: Para', '', 'Endereços separados por ponto e vírgula'),
            ('cfgEmailCC', 'E-mail: Cópia', '', ''),
        ]
        for i, (nm, rot, val, desc) in enumerate(gerais):
            r = 5 + i
            ws.set_row(r, 20)
            ws.merge_range(r, 1, r, 2, rot, self.f(bold=True, indent=1, bg_color=FUNDO_CLARO,
                                                   border=1, border_color=BRANCO))
            if isinstance(val, dtm.time):
                ws.merge_range(r, 3, r, 5, '', self.f_input(num_format='hh:mm', align='center'))
                ws.write_datetime(r, 3, dtm.datetime.combine(dtm.date(1899, 12, 31), val),
                                  self.f_input(num_format='hh:mm', align='center'))
            else:
                ws.merge_range(r, 3, r, 5, val, self.f_input(align='center'))
            ws.merge_range(r, 6, r, 9, desc, self.f(font_size=9, font_color=AZUL_ACINZ, indent=1))
            self.nome(nm, ws, r, 3)
        ws.data_validation(5, 3, 5, 3, {'validate': 'list', 'source': '=lstFonte'})
        ws.data_validation(10, 3, 10, 3, {'validate': 'list', 'source': '=lstSN'})
        # botoes
        self.botao(ws, 5, 11, 'Voltar ao Painel', 'IrPainel', 230, 30, 'primario', x=5)
        self.botao(ws, 7, 11, 'Inserir logo nas abas', 'InserirLogo', 230, 30, 'destaque', x=5)
        self.botao(ws, 9, 11, 'Ver dados brutos do MES', 'MostrarDadosMES', 230, 30, 'claro', x=5)

        ws.merge_range(14, 1, 14, 15,
                       'Tags herdadas da planilha padrão (Usinas 3 e 4). Confira no MES antes do uso. '
                       'LIE/LSE = limites de especificação (vazio = sem avaliação). '
                       'Faixa válida: valores fora dela são descartados (o padrão antigo usava 0 a 50000).',
                       self.f(font_size=9, italic=True, font_color=AZUL_ACINZ, text_wrap=True, indent=1))
        ws.set_row(14, 28)
        self.secao(ws, 15, 1, 15, 'TAGS DO MES E LIMITES - USINAS 3 E 4')
        cab = ['Grupo', 'Parâmetro', 'Unidade', 'Dec.', 'Resultado do turno', 'Tag US3', 'Tipo US3',
               'Tag US4', 'Tipo US4', 'LIE', 'LSE', 'Válido mín.', 'Válido máx.', 'Típico US3 (simul.)',
               'Típico US4 (simul.)']
        ws.set_row(16, 32)
        for i, t in enumerate(cab):
            ws.write(16, 1 + i, t, self.f(bold=True, font_color=BRANCO, bg_color=AZUL_MEDIO, text_wrap=True,
                                          align='center', border=1, border_color=BRANCO))
        for p, pr in enumerate(PARAMS, start=1):
            r = CFG_ROW1 - 1 + p - 1
            g, nome, un, dec, agg, t3, ty3, t4, ty4, tip3, tip4 = pr
            zebra = FUNDO_CLARO if p % 2 else BRANCO
            fl = self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1)
            ws.write(r, 1, g, self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1, font_color=AZUL_ACINZ))
            ws.write(r, 2, nome, self.f(bg_color=zebra, border=1, border_color=BRANCO, indent=1, bold=True))
            ws.write(r, 3, un, self.f_input(align='center'))
            ws.write(r, 4, dec, self.f_input(align='center'))
            ws.write(r, 5, agg, self.f_input(align='center'))
            ws.write(r, 6, t3, self.f_input(font_name='Consolas', font_size=9))
            ws.write(r, 7, ty3, self.f_input(font_size=9))
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
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.print_area(1, 1, CFG_ROW1 + NPARAM, 15)

        # listas
        listas = [('lstTurno', 'Turnos', ['Dia (07h às 19h)', 'Noite (19h às 07h)']),
                  ('lstTurma', 'Turmas', ['A', 'B', 'C', 'D']),
                  ('lstStatus', 'Status', ['OK', 'NÃO OK', 'FORA DE OPERAÇÃO']),
                  ('lstAgreg', 'Resultado do turno', ['Média', 'Soma', 'Último valor']),
                  ('lstFonte', 'Fonte', ['MES', 'SIMULAÇÃO']),
                  ('lstSN', 'Sim/Não', ['Sim', 'Não']),
                  ('lstTipo', 'Tipo de tag', [MV, AM])]
        ws.merge_range(4, 18, 4, 24, 'LISTAS (usadas nas caixas de seleção)',
                       self.f(bold=True, font_color=BRANCO, bg_color=AZUL_ACINZ, indent=1))
        for i, (nm, tit, itens) in enumerate(listas):
            c = 18 + i
            ws.write(5 - 1 + 1, c, tit, self.f(bold=True, font_size=9, bg_color=FUNDO_GRUPO, text_wrap=True))
            for j, it in enumerate(itens):
                ws.write(6 + j, c, it, self.f(font_size=9, border=1, border_color=FUNDO_CLARO))
            self.nome(nm, ws, 6, c, 6 + len(itens) - 1, c)

    # ------------------------------------------------------------ Dados_MES
    def aba_dados_mes(self):
        ws = self.ws_mes
        ws.set_column(0, 0, 18)
        ws.set_column(1, 2 * NPARAM, 14)
        ws.write(0, 0, 'Dados brutos do MES (Aspen IP.21) - não editar', self.f(bold=True, font_size=13, font_color=AZUL))
        txt = self.f(num_format='@', bg_color=FUNDO_CLARO)
        rot = self.f(bold=True)
        ws.write(2, 0, 'Início', rot)
        ws.write_string(2, 1, '', txt)
        ws.write(3, 0, 'Fim', rot)
        ws.write_string(3, 1, '', txt)
        self.nome('mesInicio', ws, 2, 1)
        self.nome('mesFim', ws, 3, 1)
        tags = '&","&'.join(cfg_ref(p, k) for p in range(1, NPARAM + 1) for k in ('TAG3', 'TAG4'))
        tipos = '&","&'.join(cfg_ref(p, k) for p in range(1, NPARAM + 1) for k in ('TIPO3', 'TIPO4'))
        serv = 'REPT(cfgServidor&",",%d)&cfgServidor' % (2 * NPARAM - 1)
        ws.write(4, 0, 'Tags', rot)
        ws.write_formula(4, 1, '=' + tags, self.f(font_size=8))
        ws.write(5, 0, 'Servidores', rot)
        ws.write_formula(5, 1, '=' + serv, self.f(font_size=8))
        ws.write(6, 0, 'Tipos', rot)
        ws.write_formula(6, 1, '=' + tipos, self.f(font_size=8))
        ncol = 1 + 2 * NPARAM
        saida = 'Dados_MES!%s:%s' % (rc(10, 0), rc(10 + NSLOT - 1, ncol - 1))
        ancora = 'ADDRESS(ROW(Dados_MES!A9),COLUMN(Dados_MES!A9),1,,"Dados_MES")'
        f_get = ('=_xll.AspenTech.PME.ProcessData.Functions.GetCalculationValues("time,attribute",'
                 'Dados_MES!$B$5,Dados_MES!$B$6,Dados_MES!$B$7,Dados_MES!$B$3,Dados_MES!$B$4,'
                 '"2h",0,"",0,"1",0,1560,0,0,1,1,%s,ROWS(%s)&"#"&COLUMNS(%s),1)' % (ancora, saida, saida))
        ws.write_formula(8, 0, f_get, self.f(font_size=8))
        self.nome('mesConsulta', ws, 8, 0)
        ws.write(9, 0, 'Data/hora', self.f(bold=True, bg_color=AZUL, font_color=BRANCO))
        for p in range(1, NPARAM + 1):
            for k, us in enumerate(USINAS):
                ws.write(9, 1 + (p - 1) * 2 + k, 'P%02d %s · %s' % (p, us, PARAMS[p - 1][1]),
                         self.f(bold=True, bg_color=AZUL, font_color=BRANCO, font_size=8, text_wrap=True))
        ws.set_row(9, 45)
        f_show = ('{=_xll.AspenTech.PME.ProcessData.Functions.ShowCalculationValues(%s,Dados_MES!A9, 0)}' % ancora)
        ws.write_array_formula(10, 0, 10 + NSLOT - 1, ncol - 1, f_show, self.f(num_format='0.00'))
        ws.set_column(0, 0, 18, self.f(num_format='dd/mm/yyyy hh:mm'))
        self.nome('mesSaida', ws, 10, 0, 10 + NSLOT - 1, ncol - 1)
        self.botao(ws, 0, 4, 'Voltar às Configurações', 'OcultarDadosMES', 200, 28, 'primario', y=2)

    # ------------------------------------------------------------ Informativo
    def aba_informativo(self):
        ws = self.ws_inf
        ws.set_column(0, 0, 3, None, {'hidden': True})
        ws.set_column(1, 1, 34)
        ws.set_column(2, 2, 9)
        ws.set_column(3, 3, 7)
        ws.set_column(4, 9, 9.5)
        ws.set_column(10, 10, 12)
        ws.set_column(11, 15, 8)
        ws.set_column(16, 16, 3)
        ws.set_column(17, 17, 26)
        self.cabecalho(ws, 15, 'Informativo de Qualidade - Usinas 3 e 4',
                       'Resultados do MES por turno de 12 horas  ·  Laboratório Físico', 'logoInformativo')
        lab = self.f(font_size=8, bold=True, font_color=AZUL_ACINZ, indent=1)
        val = self.f(font_size=11, bold=True, font_color=AZUL, bg_color=FUNDO_CLARO, indent=1, border=1,
                     border_color=BRANCO)
        info = [((1, 1), 'Data do turno', 'iData', 'dd/mm/yyyy'), ((2, 3), 'Turno', 'iTurno', None),
                ((4, 5), 'Turma', 'iTurma', None), ((6, 9), 'Responsável', 'iResp', None),
                ((10, 12), 'Atualizado em', 'iAtualizado', 'dd/mm/yyyy hh:mm'), ((13, 15), 'Fonte', 'iFonte', None)]
        ws.set_row(5, 15)
        ws.set_row(6, 22)
        for (c1, c2), t, nm, nf in info:
            fv = self.f(font_size=11, bold=True, font_color=AZUL, bg_color=FUNDO_CLARO, indent=1, border=1,
                        border_color=BRANCO, num_format=nf) if nf else val
            if c1 == c2:
                ws.write(5, c1, t, lab)
                ws.write_blank(6, c1, None, fv)
            else:
                ws.merge_range(5, c1, 5, c2, t, lab)
                ws.merge_range(6, c1, 6, c2, '', fv)
            self.nome(nm, ws, 6, c1)
        ws.write(7, 1, 'Período consultado:', lab)
        ws.merge_range(7, 2, 7, 15, '', self.f(font_size=9, font_color=AZUL_ACINZ))
        self.nome('iPeriodo', ws, 7, 2)

        hdr = self.f(bold=True, font_color=BRANCO, bg_color=AZUL, align='center', text_wrap=True,
                     border=1, border_color=BRANCO)
        R0 = 9
        ws.set_row(R0, 32)
        cabs = ['Parâmetro', 'Unid.', 'Usina'] + ['—'] * NSLOT + ['Resultado do turno', 'Mín.', 'Máx.', 'LIE', 'LSE',
                                                                   'Status']
        for i, t in enumerate(cabs):
            ws.write(R0, 1 + i, t, hdr)
        self.nome('iHoras', ws, R0, INF_COL_H1 - 1, R0, INF_COL_H1 - 1 + NSLOT - 1)

        r = R0 + 1
        grupo_atual = None
        for p, pr in enumerate(PARAMS, start=1):
            g, nome, un, dec, agg = pr[:5]
            if g != grupo_atual:
                grupo_atual = g
                ws.set_row(r, 18)
                ws.merge_range(r, 1, r, 15, g.upper(), self.f(bold=True, font_size=9, font_color=AZUL,
                                                               bg_color=FUNDO_GRUPO, indent=1))
                r += 1
            nf = fmt_dec(dec)
            zebra = BRANCO if p % 2 else '#F7F9FB'
            borda = dict(border=1, border_color='#E3E8EC')
            ws.merge_range(r, 1, r + 1, 1, nome, self.f(bold=True, bg_color=zebra, indent=1, text_wrap=True, **borda))
            ws.merge_range(r, 2, r + 1, 2, un, self.f(font_color=AZUL_ACINZ, bg_color=zebra, align='center', **borda))
            for k, us in enumerate(USINAS):
                rr = r + k
                self.linhas_inf[(p, k + 1)] = rr + 1
                ws.write(rr, 0, 'P%02dU%d' % (p, 3 + k), self.f(font_color=BRANCO, font_size=6))
                ws.write(rr, 3, us, self.f(bold=True, font_size=9, font_color=BRANCO,
                                           bg_color=AZUL if k == 0 else AZUL_ACINZ, align='center', **borda))
                vals = '%s:%s' % (rc(rr, INF_COL_H1 - 1), rc(rr, INF_COL_H1 - 1 + NSLOT - 1))
                for s in range(NSLOT):
                    ws.write_blank(rr, INF_COL_H1 - 1 + s, None,
                                   self.f(num_format=nf, align='center', bg_color=zebra, locked=False, **borda))
                agg_ref = cfg_ref(p, 'AGG')
                res = ('=IFERROR(IF(COUNT({v})=0,"",IF({a}="Soma",SUM({v}),IF({a}="Último valor",'
                       'LOOKUP(2,1/({v}<>""),{v}),AVERAGE({v})))),"")').format(v=vals, a=agg_ref)
                ws.write_formula(rr, 10, res, self.f(num_format=nf, bold=True, align='center', font_size=11,
                                                     bg_color=FUNDO_AZUL_CLARO, **borda))
                ws.write_formula(rr, 11, '=IF(COUNT({v})=0,"",MIN({v}))'.format(v=vals),
                                 self.f(num_format=nf, align='center', font_color=AZUL_ACINZ, bg_color=zebra, **borda))
                ws.write_formula(rr, 12, '=IF(COUNT({v})=0,"",MAX({v}))'.format(v=vals),
                                 self.f(num_format=nf, align='center', font_color=AZUL_ACINZ, bg_color=zebra, **borda))
                for c, key in ((13, 'LIE'), (14, 'LSE')):
                    ref = cfg_ref(p, key)
                    ws.write_formula(rr, c, '=IF(%s="","",%s)' % (ref, ref),
                                     self.f(num_format=nf, align='center', font_color=AZUL_ACINZ, font_size=9,
                                            bg_color=zebra, **borda))
                K, N, O = rc(rr, 10), rc(rr, 13), rc(rr, 14)
                ws.write_formula(rr, 15, ('=IF(OR({k}="",AND({n}="",{o}="")),"",IF(OR(AND({n}<>"",{k}<{n}),'
                                          'AND({o}<>"",{k}>{o})),"Fora","OK"))').format(k=K, n=N, o=O),
                                 self.f(bold=True, align='center', font_size=9, bg_color=zebra, **borda))
            r += 2
        ultima = r - 1
        # formatacao condicional
        rng_vals = '%s:%s' % (rc(R0 + 1, 4), rc(ultima, 10))
        prim = rc(R0 + 1, 4, False, False)
        n_abs = '$N%d' % (R0 + 2)
        o_abs = '$O%d' % (R0 + 2)
        fora = 'AND(ISNUMBER({c}),OR(AND(ISNUMBER({n}),{c}<{n}),AND(ISNUMBER({o}),{c}>{o})))'.format(
            c=prim, n=n_abs, o=o_abs)
        ws.conditional_format(rng_vals, {'type': 'formula', 'criteria': '=' + fora,
                                         'format': self.wb.add_format({'font_color': LARANJA, 'bold': True,
                                                                       'bg_color': LARANJA_FUNDO})})
        kcol = '%s:%s' % (rc(R0 + 1, 10), rc(ultima, 10))
        ok = 'AND(ISNUMBER(K{r}),OR(ISNUMBER($N{r}),ISNUMBER($O{r})))'.format(r=R0 + 2)
        ws.conditional_format(kcol, {'type': 'formula', 'criteria': '=' + ok,
                                     'format': self.wb.add_format({'bg_color': VERDE_FUNDO, 'font_color': VERDE_TXT})})
        pcol = '%s:%s' % (rc(R0 + 1, 15), rc(ultima, 15))
        ws.conditional_format(pcol, {'type': 'cell', 'criteria': '==', 'value': '"Fora"',
                                     'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': LARANJA})})
        ws.conditional_format(pcol, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                     'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': VERDE_TXT})})

        r = ultima + 2
        self.secao(ws, r, 1, 15, 'DESTAQUES / OBSERVAÇÕES DO TURNO')
        ws.merge_range(r + 1, 1, r + 4, 15, '', self.f_input(valign='top', text_wrap=True))
        self.nome('iObs', ws, r + 1, 1)
        r += 6
        ws.merge_range(r, 1, r, 15,
                       'Resultado do turno: média dos resultados (Produção = soma; Ritmo = último valor). '
                       'Status compara com LIE/LSE da aba Configurações. Valores fora da faixa válida são descartados.',
                       self.f(font_size=8, italic=True, font_color=AZUL_ACINZ, text_wrap=True))
        ws.set_row(r, 24)
        self.nome('iAreaImagem', ws, 1, 1, r, 15)
        ws.print_area(1, 1, r, 15)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.center_horizontally()
        ws.set_footer('&L&8Informativo US3/US4&R&8Página &P de &N')
        ws.freeze_panes(R0 + 1, 0)
        # botoes (fora da area de impressao)
        self.botao(ws, 1, 17, 'Voltar ao Painel', 'IrPainel', 185, 34, 'primario', x=4)
        self.botao(ws, 4, 17, 'Atualizar dados do MES', 'AtualizarMES', 185, 34, 'medio', x=4)
        self.botao(ws, 7, 17, 'Copiar como imagem', 'CopiarImagem', 185, 34, 'titulo', x=4)
        self.botao(ws, 10, 17, 'Gerar PDF', 'ExportarPDFManual', 185, 34, 'destaque', x=4)
        ws.protect('', {'autofilter': True, 'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Passagem de Turno
    def aba_passagem(self):
        ws = self.ws_pass
        Q = self._qn(ws)
        ws.set_column(0, 0, 2)
        ws.set_column(1, 1, 26)
        ws.set_column(2, 12, 9.3)
        ws.set_column(13, 13, 3)
        ws.set_column(14, 14, 28)
        self.cabecalho(ws, 12, 'Passagem de Turno - Laboratório Físico',
                       'Usinas 3 e 4  ·  Dia 07h-19h  ·  Noite 19h-07h  ·  Turmas A, B, C e D', 'logoPassagem')
        lab = self.f(font_size=8, bold=True, font_color=AZUL_ACINZ, indent=1)
        info = self.f(font_size=11, bold=True, font_color=AZUL, bg_color=FUNDO_CLARO, indent=1)
        inp = self.f_input(font_size=11, bold=True, indent=1)
        ws.set_row(6, 24)
        ws.write(5, 1, 'Data', lab)
        ws.write_formula(6, 1, '=IF(pData="","",pData)', self.f(font_size=11, bold=True, font_color=AZUL,
                                                               bg_color=FUNDO_CLARO, indent=1, num_format='dd/mm/yyyy'))
        ws.merge_range(5, 2, 5, 3, 'Turno', lab)
        ws.merge_range(6, 2, 6, 3, '', info)
        ws.write_formula(6, 2, '=IF(pTurno="","",pTurno)', info)
        ws.write(5, 4, 'Turma', lab)
        ws.write_formula(6, 4, '=IF(pTurma="","",pTurma)', self.f(font_size=11, bold=True, font_color=AZUL,
                                                                 bg_color=FUNDO_CLARO, align='center'))
        ws.merge_range(5, 5, 5, 6, 'Responsável (entrega)', lab)
        ws.merge_range(6, 5, 6, 6, '', info)
        ws.write_formula(6, 5, '=IF(pResp="","",pResp)', info)
        ws.merge_range(5, 7, 5, 8, 'Turma que recebe', lab)
        ws.merge_range(6, 7, 6, 8, '', self.f_input(font_size=11, bold=True, align='center'))
        ws.data_validation(6, 7, 6, 7, {'validate': 'list', 'source': '=lstTurma'})
        self.nome('ptRecebe', ws, 6, 7)
        ws.merge_range(5, 9, 5, 12, 'Recebido por', lab)
        ws.merge_range(6, 9, 6, 12, '', inp)
        self.nome('ptRecebidoPor', ws, 6, 9)
        self.mapa_add('Cabeçalho', '', 'Turma que recebe', '', '#ptRecebe')
        self.mapa_add('Cabeçalho', '', 'Recebido por', '', '#ptRecebidoPor')

        rot = self.f(bold=True, font_size=9, bg_color=FUNDO_CLARO, indent=1, text_wrap=True, locked=False,
                     border=1, border_color=BRANCO)
        txt = self.f_input(valign='top', text_wrap=True)
        hdr = self.f(bold=True, font_size=9, font_color=AZUL, bg_color=FUNDO_GRUPO, align='center',
                     border=1, border_color=BRANCO)

        # 1. Equipe
        r = 8
        self.secao(ws, r, 1, 12, '1.  EQUIPE DO TURNO')
        esq = ['Técnico / Supervisor', 'Amostrador 01', 'Amostrador 02', 'LTF 01', 'LTF 02']
        dir_ = ['Laboratorista 01', 'Laboratorista 02', 'Laboratorista 03', 'Laboratorista 04', 'Apoio / T03']
        for i in range(5):
            rr = r + 1 + i
            ws.set_row(rr, 20)
            ws.write(rr, 1, esq[i], rot)
            ws.merge_range(rr, 2, rr, 6, '', self.f_input())
            ws.merge_range(rr, 7, rr, 8, dir_[i], rot)
            ws.merge_range(rr, 9, rr, 12, '', self.f_input())
            self.mapa_add('Equipe', '', '@' + rc(rr, 1), '', '@' + rc(rr, 2))
            self.mapa_add('Equipe', '', '@' + rc(rr, 7), '', '@' + rc(rr, 9))

        # 2. Seguranca
        r = 15
        self.secao(ws, r, 1, 12, '2.  SEGURANÇA (SSMA)')
        ws.set_row(r + 1, 20)
        ws.write(r + 1, 1, 'Houve ocorrência de segurança?', rot)
        ws.merge_range(r + 1, 2, r + 1, 3, '', self.f_input(align='center', bold=True))
        ws.data_validation(r + 1, 2, r + 1, 2, {'validate': 'list', 'source': '=lstSN'})
        self.nome('ptSegFlag', ws, r + 1, 2)
        ws.merge_range(r + 1, 4, r + 1, 12, 'Relate abaixo acidentes, incidentes, desvios ou condições inseguras.',
                       self.f(font_size=8, italic=True, font_color=AZUL_ACINZ, indent=1))
        ws.merge_range(r + 2, 1, r + 4, 1, 'Descrição', rot)
        ws.merge_range(r + 2, 2, r + 4, 12, '', txt)
        self.nome('ptSeguranca', ws, r + 2, 2)
        self.mapa_add('Segurança', '', 'Houve ocorrência?', '#ptSegFlag', '')
        self.mapa_add('Segurança', '', 'Descrição', '', '#ptSeguranca')

        # 3. Testes e pendencias
        r = 21
        self.secao(ws, r, 1, 12, '3.  TESTES E PENDÊNCIAS')
        blocos = [('Pendências recebidas do turno anterior', 'ptPendRecebidas', FUNDO_AZUL_CLARO),
                  ('Testes realizados no turno', 'ptTestes', None),
                  ('Pendências / testes a realizar (para o próximo turno)', 'ptPendencias', None)]
        rr = r + 1
        for t, nm, cor in blocos:
            ws.merge_range(rr, 1, rr + 2, 1, t, rot)
            ws.merge_range(rr, 2, rr + 2, 12, '', self.f_input(valign='top', text_wrap=True,
                                                               **({'bg_color': cor} if cor else {})))
            self.nome(nm, ws, rr, 2)
            self.mapa_add('Texto', '', t, '', '#' + nm)
            rr += 3

        # 4. Minerodutos
        r = 32
        self.secao(ws, r, 1, 12, '4.  MINERODUTOS / BATCH')
        ws.write(r + 1, 1, 'Mineroduto', hdr)
        for (c1, c2, t) in [(2, 3, 'Batch'), (4, 5, 'Início'), (6, 7, 'Término'), (8, 12, 'Observação')]:
            ws.merge_range(r + 1, c1, r + 1, c2, t, hdr)
        for i in range(3):
            rr = r + 2 + i
            ws.set_row(rr, 20)
            ws.write(rr, 1, 'Mineroduto %02d' % (i + 1), rot)
            for (c1, c2, t) in [(2, 3, 'Batch'), (4, 5, 'Início'), (6, 7, 'Término'), (8, 12, 'Observação')]:
                ws.merge_range(rr, c1, rr, c2, '', self.f_input(align='center' if c2 < 8 else 'left'))
                self.mapa_add('Batch', '@' + rc(rr, 1), t, '', '@' + rc(rr, c1))

        # 5. Equipamentos
        r = 38
        self.secao(ws, r, 1, 12, '5.  STATUS DOS EQUIPAMENTOS (US3 / US4)')
        ws.write(r + 1, 1, 'Área', hdr)
        ws.merge_range(r + 1, 2, r + 1, 4, 'Equipamento', hdr)
        ws.merge_range(r + 1, 5, r + 1, 6, 'Status', hdr)
        ws.merge_range(r + 1, 7, r + 1, 12, 'Comentário', hdr)
        equips = [('Preparação', 'Moinho 01'), ('Preparação', 'Moinho 02'),
                  ('Preparação', 'Planta 03'), ('Preparação', 'Planta 04'),
                  ('Amostragem PF / Mistura', 'Alimentação US 03 - 02TP04'),
                  ('Amostragem PF / Mistura', 'Alimentação US 04'),
                  ('Amostragem PF / Mistura', 'Descarga US 03 - 03TP03'),
                  ('Amostragem PF / Mistura', 'Descarga US 04'),
                  ('Amostragem PF / Mistura', 'Linha de Mistura US 03 - 04TP05'),
                  ('Amostragem PF / Mistura', 'Linha de Mistura US 04 - 04TP04'),
                  ('Amostragem PQ', 'Cortador PQ US 03 Primário - 07CR02'),
                  ('Amostragem PQ', 'Cortador PQ US 03 Secundário - 07CR05'),
                  ('Amostragem PQ', 'Cortador PQ US 04 Primário - 07CR001'),
                  ('Amostragem PQ', 'Cortador PQ US 04 Secundário - 07CR002')]
        e0 = r + 2
        for i, (area, eq) in enumerate(equips):
            rr = e0 + i
            ws.set_row(rr, 19)
            ws.write(rr, 1, area, self.f(font_size=8, font_color=AZUL_ACINZ, bg_color=FUNDO_CLARO, indent=1,
                                         locked=False, border=1, border_color=BRANCO))
            ws.merge_range(rr, 2, rr, 4, eq, self.f(font_size=9, bold=True, bg_color=FUNDO_CLARO, indent=1,
                                                    locked=False, border=1, border_color=BRANCO, shrink=True))
            ws.merge_range(rr, 5, rr, 6, '', self.f_input(align='center', bold=True, font_size=8))
            ws.merge_range(rr, 7, rr, 12, '', self.f_input(font_size=9))
            self.mapa_add('Equipamento', '@' + rc(rr, 1), '@' + rc(rr, 2), '@' + rc(rr, 5), '@' + rc(rr, 7))
        e1 = e0 + len(equips) - 1
        ws.data_validation(e0, 5, e1, 5, {'validate': 'list', 'source': '=lstStatus'})
        self.nome('ptStatus', ws, e0, 5, e1, 5)
        st = '%s:%s' % (rc(e0, 5), rc(e1, 5))
        ws.conditional_format(st, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                   'format': self.wb.add_format({'bg_color': VERDE_FUNDO, 'font_color': VERDE_TXT})})
        ws.conditional_format(st, {'type': 'cell', 'criteria': '==', 'value': '"NÃO OK"',
                                   'format': self.wb.add_format({'bg_color': LARANJA, 'font_color': BRANCO})})
        ws.conditional_format(st, {'type': 'cell', 'criteria': '==', 'value': '"FORA DE OPERAÇÃO"',
                                   'format': self.wb.add_format({'bg_color': CINZA, 'font_color': AZUL_PROFUNDO})})
        self.eq_rows = (e0, e1)

        # 6. Comentarios por usina
        r = e1 + 2
        self.secao(ws, r, 1, 12, '6.  COMENTÁRIOS POR USINA')
        ws.write(r + 1, 1, 'Assunto', hdr)
        ws.merge_range(r + 1, 2, r + 1, 6, 'Usina 3', self.f(bold=True, font_color=BRANCO, bg_color=AZUL, align='center'))
        ws.merge_range(r + 1, 7, r + 1, 12, 'Usina 4', self.f(bold=True, font_color=BRANCO, bg_color=AZUL_ACINZ,
                                                              align='center'))
        campos = ['Produto', 'Retorno do Pátio', 'Aglomerante', 'Combustível sólido', 'Qualidade das pelotas',
                  'Processos', 'Sistema de amostragem']
        c0 = r + 2
        for i, cp in enumerate(campos):
            rr = c0 + i
            ws.set_row(rr, 32)
            ws.write(rr, 1, cp, rot)
            ws.merge_range(rr, 2, rr, 6, '', self.f_input(valign='top', text_wrap=True, font_size=9))
            ws.merge_range(rr, 7, rr, 12, '', self.f_input(valign='top', text_wrap=True, font_size=9))
            self.mapa_add('Comentário', 'US3', '@' + rc(rr, 1), '', '@' + rc(rr, 2))
            self.mapa_add('Comentário', 'US4', '@' + rc(rr, 1), '', '@' + rc(rr, 7))
        self.com_rows = (c0, c0 + len(campos) - 1)

        # 7. Embarques
        r = c0 + len(campos) + 1
        self.secao(ws, r, 1, 12, '7.  EMBARQUES / CLIENTES')
        cols = [(1, 1, 'Cliente'), (2, 2, 'Produto'), (3, 3, 'Navio'), (4, 4, 'Início'), (5, 5, 'Término'),
                (6, 6, 'Qtd (t)'), (7, 7, 'Tamb. (%)'), (8, 8, 'Compr. (kgf)'), (9, 9, '-6,3 mm (%)'),
                (10, 12, 'Observação')]
        ws.set_row(r + 1, 26)
        for c1, c2, t in cols:
            if c1 == c2:
                ws.write(r + 1, c1, t, self.f(bold=True, font_size=8, font_color=AZUL, bg_color=FUNDO_GRUPO,
                                              align='center', text_wrap=True, border=1, border_color=BRANCO))
            else:
                ws.merge_range(r + 1, c1, r + 1, c2, t, hdr)
        for i in range(4):
            rr = r + 2 + i
            ws.set_row(rr, 19)
            for c1, c2, t in cols:
                if c1 == c2:
                    ws.write_blank(rr, c1, None, self.f_input(font_size=9, align='center' if c1 > 1 else 'left'))
                else:
                    ws.merge_range(rr, c1, rr, c2, '', self.f_input(font_size=9))
                self.mapa_add('Embarque', 'Embarque %d' % (i + 1), t, '', '@' + rc(rr, c1))

        # 8. Observacoes
        r = r + 7
        self.secao(ws, r, 1, 12, '8.  OBSERVAÇÕES GERAIS')
        ws.merge_range(r + 1, 1, r + 3, 12, '', txt)
        self.nome('ptObs', ws, r + 1, 1)
        self.mapa_add('Texto', '', 'Observações gerais', '', '#ptObs')
        self.mapa_add('Informativo', '', 'Destaques do turno', '', '#iObs')

        r = r + 5
        ws.set_row(r, 30)
        ws.merge_range(r, 1, r, 5, 'Entregue por: ____________________________',
                       self.f(font_size=9, font_color=AZUL_ACINZ, valign='bottom'))
        ws.merge_range(r, 7, r, 12, 'Recebido por: ____________________________',
                       self.f(font_size=9, font_color=AZUL_ACINZ, valign='bottom'))
        self.pass_ultima = r
        ws.print_area(1, 1, r, 12)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.center_horizontally()
        ws.set_footer('&L&8Passagem de Turno US3/US4&R&8Página &P de &N')

        self.botao(ws, 1, 14, 'Voltar ao Painel', 'IrPainel', 205, 34, 'primario', x=4)
        self.botao(ws, 4, 14, 'Carregar pendências\ndo turno anterior', 'CarregarPendencias', 205, 44, 'medio', x=4)
        self.botao(ws, 8, 14, 'Limpar formulário', 'LimparPassagem', 205, 34, 'claro', x=4)
        self.botao(ws, 11, 14, 'Fechar turno', 'FecharTurno', 205, 44, 'destaque', x=4)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Histórico
    def aba_historico(self):
        ws = self.ws_hist
        self.cabecalho(ws, 12, 'Histórico dos Turnos', 'Um registro por turno fechado  ·  resultados US3 e US4',
                       'logoHistorico', col_logo_fim=1, col_ini=0)
        cab = HIST_FIXAS + ['%s · %s' % (p[1], us) for p in PARAMS for us in USINAS] + HIST_FINAIS
        self.hist_cab = cab
        h = HIST_ROW_HDR - 1
        ws.set_row(h, 48)
        for i, t in enumerate(cab):
            cor = AZUL if i < len(HIST_FIXAS) else (AZUL_MEDIO if i < HIST_COL_NOK - 1 else AZUL_ACINZ)
            ws.write(h, i, t, self.f(bold=True, font_color=BRANCO, bg_color=cor, text_wrap=True, align='center',
                                     font_size=9, border=1, border_color=BRANCO))
        base = dict(font_size=9)
        ws.set_column(0, 0, 11, self.f(num_format='dd/mm/yyyy', align='center', **base))
        ws.set_column(1, 1, 17, self.f(**base))
        ws.set_column(2, 2, 7, self.f(align='center', **base))
        ws.set_column(3, 3, 20, self.f(**base))
        ws.set_column(4, 4, 9, self.f(align='center', **base))
        ws.set_column(5, 5, 15, self.f(num_format='dd/mm/yyyy hh:mm', **base))
        ws.set_column(6, 6, 16, self.f(**base))
        for p, pr in enumerate(PARAMS, start=1):
            for k in range(2):
                c = HIST_COL_RES1 - 1 + (p - 1) * 2 + k
                ws.set_column(c, c, 11, self.f(num_format=fmt_dec(pr[3]), align='center', **base))
        c = HIST_COL_NOK - 1
        ws.set_column(c, c, 9, self.f(align='center', **base))
        ws.set_column(c + 1, c + 1, 10, self.f(align='center', **base))
        ws.set_column(c + 2, c + 4, 40, self.f(text_wrap=True, valign='top', **base))
        ws.set_column(c + 5, c + 5, 40, self.f(font_size=8, font_color=AZUL_ACINZ))
        ws.freeze_panes(HIST_ROW_HDR, 3)
        ws.autofilter(h, 0, 20000, len(cab) - 1)
        self.botao(ws, 4, 4, 'Voltar ao Painel', 'IrPainel', 160, 26, 'pequeno', x=0, y=1)
        self.botao(ws, 4, 6, 'Tendências', 'IrTendencias', 120, 26, 'pequeno', x=0, y=1)
        ws.set_row(4, 30)

    # ------------------------------------------------------------ Registro
    def aba_registro(self):
        ws = self.ws_reg
        self.cabecalho(ws, 8, 'Registro da Passagem de Turno',
                       'Cada informação preenchida vira uma linha (ideal para filtros e Power BI)',
                       'logoRegistro', col_logo_fim=1, col_ini=0)
        cab = ['Data', 'Turno', 'Turma', 'Tipo', 'Seção / Usina', 'Item', 'Status', 'Texto', 'Registrado em']
        h = REG_ROW_HDR - 1
        ws.set_row(h, 26)
        for i, t in enumerate(cab):
            ws.write(h, i, t, self.f(bold=True, font_color=BRANCO, bg_color=AZUL, align='center', font_size=9,
                                     border=1, border_color=BRANCO))
        base = dict(font_size=9)
        larg = [11, 17, 7, 13, 22, 34, 16, 70, 15]
        for i, w in enumerate(larg):
            extra = {}
            if i == 0:
                extra = dict(num_format='dd/mm/yyyy', align='center')
            if i == 8:
                extra = dict(num_format='dd/mm/yyyy hh:mm')
            if i == 7:
                extra = dict(text_wrap=True, valign='top')
            ws.set_column(i, i, w, self.f(**base, **extra))
        ws.freeze_panes(REG_ROW_HDR, 0)
        ws.autofilter(h, 0, 50000, len(cab) - 1)
        ws.conditional_format(REG_ROW_HDR, 6, 50000, 6, {'type': 'cell', 'criteria': '==', 'value': '"NÃO OK"',
                                                          'format': self.wb.add_format({'font_color': LARANJA,
                                                                                        'bold': True})})
        self.botao(ws, 4, 3, 'Voltar ao Painel', 'IrPainel', 160, 26, 'pequeno', y=1)
        ws.set_row(4, 30)

    # ------------------------------------------------------------ Tendências
    def aba_tendencias(self):
        ws = self.ws_tend
        N = 30
        ws.set_column(0, 0, 3, None, {'hidden': True})
        ws.set_column(1, 1, 12)
        ws.set_column(2, 5, 10)
        ws.set_column(6, 6, 3)
        ws.set_column(7, 16, 9)
        ws.set_column(17, 17, 10, None, {'hidden': True})
        self.cabecalho(ws, 16, 'Tendências por Parâmetro', 'Últimos %d turnos fechados  ·  Usina 3 x Usina 4' % N,
                       'logoTendencias')
        ws.set_row(5, 26)
        ws.write(5, 1, 'Parâmetro:', self.f(bold=True, align='right'))
        ws.merge_range(5, 2, 5, 5, 'Resistência à Compressão', self.f_input(bold=True, font_size=11, indent=1))
        ws.data_validation(5, 2, 5, 2, {'validate': 'list', 'source': '=lstParam'})
        self.nome('tdParam', ws, 5, 2)
        # auxiliares (coluna R oculta)
        ws.write_formula(1, 17, "=COUNTA(%s!$A$%d:$A$20000)" % (SH_HIST, HIST_ROW_HDR + 1))
        ws.write_formula(2, 17, '=MATCH(tdParam&" · US3",%s!$%d:$%d,0)' % (SH_HIST, HIST_ROW_HDR, HIST_ROW_HDR))
        ws.write_formula(3, 17, '=MATCH(tdParam&" · US4",%s!$%d:$%d,0)' % (SH_HIST, HIST_ROW_HDR, HIST_ROW_HDR))
        ws.write_formula(4, 17, '=MATCH(tdParam,lstParam,0)')
        hdr = self.f(bold=True, font_color=BRANCO, bg_color=AZUL, align='center', font_size=9)
        R0 = 7
        for i, t in enumerate(['Turno', 'US3', 'US4', 'LIE', 'LSE']):
            ws.write(R0, 1 + i, t, hdr)
        corpo = self.f(font_size=9, align='center', num_format='0.00')
        for k in range(1, N + 1):
            r = R0 + k
            linha_hist = '%d+$R$2-%d+$A%d' % (HIST_ROW_HDR, N, r + 1)
            idx = '$R$2-%d+$A%d' % (N, r + 1)
            ws.write(r, 0, k)
            ws.write_formula(r, 1, ('=IF({i}<1,"",TEXT(DAY(INDEX({h}!$A:$A,{l})),"00")&"/"&TEXT(MONTH(INDEX({h}!$A:$A,{l})),"00")'
                                    '&" "&LEFT(INDEX({h}!$B:$B,{l}),1))').format(i=idx, h=SH_HIST, l=linha_hist),
                             self.f(font_size=9, align='center'))
            for c, colref in ((2, '$R$3'), (3, '$R$4')):
                v = 'INDEX({h}!$A:$ZZ,{l},{c})'.format(h=SH_HIST, l=linha_hist, c=colref)
                ws.write_formula(r, c, '=IF(OR({i}<1,ISNA({c})),NA(),IF({v}="",NA(),{v}))'.format(i=idx, c=colref, v=v),
                                 corpo)
            for c, key in ((4, 'LIE'), (5, 'LSE')):
                rng = '%s!$%s$%d:$%s$%d' % (SH_CFG, colname(CFG_COL[key] - 1), CFG_ROW1,
                                            colname(CFG_COL[key] - 1), CFG_ROW1 + NPARAM - 1)
                v = 'INDEX(%s,$R$5)' % rng
                ws.write_formula(r, c, '=IF(ISNA($R$5),NA(),IF(%s="",NA(),%s))' % (v, v), corpo)
        ultima = R0 + N
        ws.conditional_format(R0 + 1, 2, ultima, 5, {'type': 'formula', 'criteria': '=ISNA(C%d)' % (R0 + 2),
                                                     'format': self.wb.add_format({'font_color': '#D0D7DC'})})
        ch = self.wb.add_chart({'type': 'line'})
        cats = ['Tendências', R0 + 1, 1, ultima, 1]
        series = [(2, 'Usina 3', AZUL, 2.5, 'solid', True), (3, 'Usina 4', OURO, 2.5, 'solid', True),
                  (4, 'LIE', LARANJA, 1.25, 'dash', False), (5, 'LSE', LARANJA, 1.25, 'dash', False)]
        for c, nm, cor, w, dash, mk in series:
            s = {'name': nm, 'categories': cats, 'values': ['Tendências', R0 + 1, c, ultima, c],
                 'line': {'color': cor, 'width': w, 'dash_type': dash}}
            s['marker'] = {'type': 'circle', 'size': 5, 'fill': {'color': cor}, 'border': {'color': cor}} if mk \
                else {'type': 'none'}
            ch.add_series(s)
        ch.set_title({'name': ['Tendências', 5, 2], 'name_font': {'name': FONTE, 'size': 12, 'color': AZUL}})
        ch.set_legend({'position': 'bottom', 'font': {'name': FONTE, 'size': 9}})
        ch.set_x_axis({'num_font': {'name': FONTE, 'size': 8, 'rotation': -45}, 'line': {'color': CINZA}})
        ch.set_y_axis({'num_font': {'name': FONTE, 'size': 8}, 'major_gridlines': {'visible': True,
                       'line': {'color': '#E3E8EC'}}, 'line': {'none': True}})
        ch.set_chartarea({'border': {'none': True}})
        ch.show_na_as_empty_cell()
        ch.show_blanks_as('gap')
        ch.set_size({'width': 760, 'height': 420})
        ws.insert_chart(R0, 7, ch, {'x_offset': 6})

        r = ultima + 2
        ws.write(r, 1, 'Estatística', hdr)
        ws.write(r, 2, 'US3', hdr)
        ws.write(r, 3, 'US4', hdr)
        stats = [('Nº de turnos', 'COUNT({c})'), ('Média', 'AGGREGATE(1,6,{c})'), ('Mínimo', 'AGGREGATE(5,6,{c})'),
                 ('Máximo', 'AGGREGATE(4,6,{c})'), ('Desvio padrão', 'AGGREGATE(7,6,{c})')]
        for i, (t, fm) in enumerate(stats):
            ws.write(r + 1 + i, 1, t, self.f(bold=True, font_size=9, bg_color=FUNDO_CLARO, indent=1))
            for j, col in enumerate('CD'):
                rngc = '%s%d:%s%d' % (col, R0 + 2, col, ultima + 1)
                ws.write_formula(r + 1 + i, 2 + j, '=IFERROR(%s,"")' % fm.format(c=rngc),
                                 self.f(font_size=9, align='center', num_format='0' if i == 0 else '0.00',
                                        bg_color=FUNDO_CLARO))
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 1)
        ws.print_area(1, 1, r + len(stats), 16)
        self.botao(ws, 5, 7, 'Voltar ao Painel', 'IrPainel', 160, 28, 'pequeno', x=4)
        self.botao(ws, 5, 10, 'Ver Histórico', 'IrHistorico', 140, 28, 'pequeno', x=4)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    # ------------------------------------------------------------ Painel
    def aba_painel(self):
        ws = self.ws_painel
        ws.set_column(0, 0, 2)
        ws.set_column(1, 16, 9.3)
        ws.set_column(17, 17, 2)
        self.cabecalho(ws, 16, 'Informativo de Qualidade & Passagem de Turno',
                       'Laboratório Físico  ·  Usinas 3 e 4  ·  Resultados do MES', 'logoPainel', col_logo_fim=3)
        sub = self.f(bold=True, font_size=9, font_color=AZUL_ACINZ)
        lab = self.f(font_size=8, bold=True, font_color=AZUL_ACINZ, indent=1)
        ws.write(5, 1, 'TURNO SELECIONADO', sub)
        ws.merge_range(6, 1, 6, 3, 'Data do turno', lab)
        ws.merge_range(6, 4, 6, 7, 'Turno', lab)
        ws.merge_range(6, 8, 6, 9, 'Turma', lab)
        ws.merge_range(6, 10, 6, 16, 'Responsável', lab)
        ws.set_row(7, 30)
        big = dict(font_size=14, bold=True, font_color=AZUL, bg_color=FUNDO_INPUT, border=2, border_color=OURO,
                   locked=False)
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
        ws.merge_range(8, 4, 8, 16, 'Dia: 07h às 19h   ·   Noite: 19h às 07h do dia seguinte   ·   '
                                    'Turmas A, B, C e D',
                       self.f(font_size=8, italic=True, font_color=AZUL_ACINZ, indent=1))

        ws.write(10, 1, 'AÇÕES DO TURNO', sub)
        for r in range(11, 17):
            ws.set_row(r, 20)
        bw, bh = 262, 116
        acoes = [(1, '1\nATUALIZAR DADOS DO MES\nBusca os resultados das Usinas 3 e 4', 'AtualizarMES', 'primario'),
                 (5, '2\nPASSAGEM DE TURNO\nEquipe, segurança, equipamentos e pendências', 'IrPassagem', 'medio'),
                 (9, '3\nINFORMATIVO DO TURNO\nResultados, limites e status por usina', 'IrInformativo', 'titulo'),
                 (13, '4\nFECHAR TURNO\nGrava o histórico e gera o PDF', 'FecharTurno', 'destaque')]
        for col, txt, mac, est in acoes:
            self.botao(ws, 11, col, txt, mac, bw, bh, est, x=4, y=2)

        ws.write(18, 1, 'SITUAÇÃO DO TURNO', sub)
        tl = self.f(font_size=8, bold=True, font_color=AZUL_ACINZ, bg_color=FUNDO_CLARO, indent=1)
        tv = dict(font_size=13, bold=True, font_color=AZUL, bg_color=FUNDO_CLARO, indent=1)
        ws.set_row(20, 30)
        d = 'iData'
        tiles = [
            (1, 4, 'Informativo carregado',
             '=IF(%s="","— não carregado —",TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)&"  ·  "&IFERROR(LEFT(iTurno,FIND(" (",iTurno)-1),iTurno))'
             % (d, d, d, d), self.f(**tv)),
            (5, 8, 'Última atualização do MES', '=IF(iAtualizado="","—",iAtualizado)',
             self.f(num_format='dd/mm/yyyy hh:mm', **tv)),
            (9, 12, 'Passagem de turno preenchida', self._formula_preench(), self.f(num_format='0%', **tv)),
            (13, 16, 'Equipamentos NÃO OK', '=COUNTIF(ptStatus,"NÃO OK")', self.f(num_format='0', **tv)),
        ]
        for c1, c2, t, fm, fv in tiles:
            ws.merge_range(19, c1, 19, c2, t, tl)
            ws.merge_range(20, c1, 20, c2, '', fv)
            ws.write_formula(20, c1, fm, fv)
        ws.conditional_format(20, 13, 20, 13, {'type': 'cell', 'criteria': '>', 'value': 0,
                                               'format': self.wb.add_format({'font_color': BRANCO,
                                                                             'bg_color': LARANJA})})
        ws.conditional_format(20, 9, 20, 9, {'type': 'cell', 'criteria': '>=', 'value': 1,
                                             'format': self.wb.add_format({'font_color': VERDE_TXT,
                                                                           'bg_color': VERDE_FUNDO})})

        ws.write(22, 1, 'INDICADORES DO TURNO (resultado do turno)', sub)
        for i, (p, titulo) in enumerate(KPIS):
            c = 1 + i * 4
            un = PARAMS[p - 1][2]
            nf = fmt_dec(PARAMS[p - 1][3])
            ws.merge_range(23, c, 23, c + 3, '%s (%s)' % (titulo, un),
                           self.f(bold=True, font_size=9, font_color=BRANCO, bg_color=AZUL, indent=1))
            for k, us in enumerate(USINAS):
                r = 24 + k
                ws.set_row(r, 26)
                chave = 'P%02dU%d' % (p, 3 + k)
                ws.write(r, c, us, self.f(bold=True, font_size=9, font_color=BRANCO,
                                          bg_color=AZUL_MEDIO if k == 0 else AZUL_ACINZ, align='center'))
                ws.merge_range(r, c + 1, r, c + 2, '', self.f(num_format=nf, font_size=15, bold=True,
                                                              font_color=AZUL, bg_color=FUNDO_CLARO, align='center'))
                ws.write_formula(r, c + 1, '=IFERROR(INDEX(Informativo!$K:$K,MATCH("%s",Informativo!$A:$A,0)),"")'
                                 % chave, self.f(num_format=nf, font_size=15, bold=True, font_color=AZUL,
                                                 bg_color=FUNDO_CLARO, align='center'))
                ws.write_formula(r, c + 3, '=IFERROR(INDEX(Informativo!$P:$P,MATCH("%s",Informativo!$A:$A,0)),"")'
                                 % chave, self.f(bold=True, font_size=9, align='center', bg_color=FUNDO_CLARO))
            cf = '%s:%s' % (rc(24, c + 3), rc(25, c + 3))
            ws.conditional_format(cf, {'type': 'cell', 'criteria': '==', 'value': '"Fora"',
                                       'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': LARANJA})})
            ws.conditional_format(cf, {'type': 'cell', 'criteria': '==', 'value': '"OK"',
                                       'format': self.wb.add_format({'font_color': BRANCO, 'bg_color': VERDE_TXT})})

        ws.write(27, 1, 'MAIS OPÇÕES', sub)
        ws.set_row(28, 30)
        extras = [(1, 'Histórico dos turnos', 'IrHistorico'), (5, 'Registro da passagem', 'IrRegistro'),
                  (9, 'Tendências', 'IrTendencias'), (13, 'Configurações', 'IrConfig')]
        for col, txt, mac in extras:
            self.botao(ws, 28, col, txt, mac, 262, 32, 'claro', x=4, y=4)
        ws.merge_range(30, 1, 30, 16, '', self.f(font_size=8, italic=True, font_color=AZUL_ACINZ))
        ws.write_formula(30, 1, '="Fonte dos dados: "&cfgFonte&"   ·   Servidor MES: "&cfgServidor&'
                                '"   ·   Tags e limites na aba Configurações"',
                         self.f(font_size=8, italic=True, font_color=AZUL_ACINZ))
        ws.protect('', {'format_columns': True, 'format_rows': True})
        ws.set_landscape()
        ws.fit_to_pages(1, 1)

    def _formula_preench(self):
        e0, e1 = self.eq_rows
        c0, c1 = self.com_rows
        Q = SH_PASS
        partes = ['COUNTA(%s!$C$10:$C$14)' % Q, 'COUNTA(%s!$J$10:$J$14)' % Q, 'COUNTA(ptRecebe)',
                  'COUNTA(ptRecebidoPor)', 'COUNTA(ptSegFlag)', 'COUNTA(ptTestes)',
                  'COUNTA(ptStatus)', 'COUNTA(%s!$C$%d:$C$%d)' % (Q, c0 + 1, c1 + 1),
                  'COUNTA(%s!$H$%d:$H$%d)' % (Q, c0 + 1, c1 + 1)]
        total = 5 + 5 + 1 + 1 + 1 + 1 + (e1 - e0 + 1) + 2 * (c1 - c0 + 1)
        return '=(%s)/%d' % ('+'.join(partes), total)

    # ------------------------------------------------------------ _Mapa
    def aba_mapa(self):
        ws = self.ws_mapa
        for i, t in enumerate(['Tipo', 'Seção', 'Item', 'Status', 'Texto']):
            ws.write(0, i, t)
        for r, linha in enumerate(self.mapa, start=1):
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

    # ------------------------------------------------------------ VBA
    def modulo_layout(self):
        c = CFG_COL
        linhas = ["Option Explicit", "",
                  "' Gerado por build_workbook.py - posicoes fixas das abas (nao editar a mao)",
                  "Public Const NPARAM As Long = %d" % NPARAM,
                  "Public Const NSLOT As Long = %d" % NSLOT,
                  "Public Const HORAS_SLOT As Double = 2",
                  "Public Const CFG_ROW1 As Long = %d" % CFG_ROW1]
        for k in ['GRUPO', 'PARAM', 'UNID', 'DEC', 'AGG', 'TAG3', 'TIPO3', 'TAG4', 'TIPO4', 'LIE', 'LSE',
                  'VMIN', 'VMAX', 'TIP3', 'TIP4']:
            linhas.append("Public Const CFG_COL_%s As Long = %d" % (k, c[k]))
        linhas += ["Public Const INF_COL_H1 As Long = %d" % INF_COL_H1,
                   "Public Const INF_COL_RES As Long = %d" % INF_COL_RES,
                   "Public Const HIST_ROW_HDR As Long = %d" % HIST_ROW_HDR,
                   "Public Const HIST_COL_RES1 As Long = %d" % HIST_COL_RES1,
                   "Public Const HIST_COL_NOK As Long = %d" % HIST_COL_NOK,
                   "Public Const REG_ROW_HDR As Long = %d" % REG_ROW_HDR,
                   "' Senha de protecao das abas (vazio = sem senha)",
                   'Public Const SENHA As String = ""', ""]
        return '\n'.join(linhas)


# ============================================================================
def ler_vba(nome):
    with open(os.path.join(VBA_DIR, nome), encoding='utf-8') as fh:
        return fh.read()


def montar_vba(layout_code):
    with open(os.path.join(VBA_DIR, 'ModLayout.bas'), 'w', encoding='utf-8') as fh:
        fh.write(layout_code)
    mods = [{'name': 'ThisWorkbook', 'kind': 'workbook', 'code': ler_vba('ThisWorkbook.cls')}]
    for cn in ['shPainel', 'shInformativo', 'shPassagem', 'shHistorico', 'shRegistro', 'shTendencias',
               'shConfig', 'shDadosMES', 'shMapa']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': ''})
    for m in ['ModLayout', 'ModGeral', 'ModMES', 'ModTurno', 'ModExportar']:
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
    bloco = bloco.replace('prst="rect"', 'prst="roundRect"', 1)
    bloco = bloco.replace('<xdr:clientData/>', '<xdr:clientData fPrintsWithSheet="0"/>')
    bloco = bloco.replace('<a:p><a:r>', '<a:p><a:pPr algn="ctr"/><a:r>')
    # Botoes de varias linhas: 1a linha = numero (pequeno), 2a = titulo, demais = descricao pequena
    paras = re.split(r'(?=<a:p>)', bloco)
    textos = [i for i, p in enumerate(paras) if p.startswith('<a:p>') and '<a:r>' in p]
    if len(textos) >= 3:
        for n, i in enumerate(textos):
            if n == 0:
                paras[i] = re.sub(r'sz="\d+"', 'sz="1000"', paras[i])
                paras[i] = paras[i].replace('<a:solidFill><a:srgbClr val="FFFFFF"/>',
                                            '<a:solidFill><a:srgbClr val="FDB913"/>')
            elif n == 1:
                paras[i] = re.sub(r'sz="\d+"', 'sz="1300"', paras[i])
            else:
                paras[i] = re.sub(r'sz="\d+"', 'sz="900"', paras[i])
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
