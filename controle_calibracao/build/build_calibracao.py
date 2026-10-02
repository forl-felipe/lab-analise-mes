"""Gera Controle_Calibracao_LCP.xlsm (Laboratorio Fisico - calibracoes, afericoes e comparativos).

Abas visiveis:
  Lancamento - a unica aba de digitacao: cabecalho (data, turno, letra, responsavel), agenda do turno
               e os 9 ensaios, com calculo e resultado (Conforme / Nao conforme) na hora.
               Botao "Registrar lancamento" grava os resultados na base (BD_Afericoes).
  Painel     - indicadores do periodo: conformidade, aderencia a agenda, situacao de cada equipamento,
               graficos de tendencia e lista de nao conformidades.
Abas ocultas:
  BD_Afericoes (tabela tbl_Afericoes) e BD_Limites (tabela tbl_Limites): mesmas colunas usadas pelo Power BI.
  BD_Limites e tambem a aba de criterios (nominal, limites e tolerancias), editavel.
  Config (listas, agenda semanal, backup), _Staging (linhas prontas para a base), _Graficos (dados dos graficos).

Uso:  python3 build_calibracao.py [--teste]
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
SAIDA = os.path.join(RAIZ, 'Controle_Calibracao_LCP%s.xlsm' % ('_TESTE' if TESTE else ''))

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
TEXTO = '#1D2B36'
TEXTO_SEC = '#5B6B77'
BORDA = '#DDE3E8'
FUNDO_INPUT = '#FFF8E1'
FUNDO_CLARO = '#F3F5F7'
FUNDO_GRUPO = '#E8EDF1'
FUNDO_AZUL_CLARO = '#EEF8FD'
VERDE_TXT, VERDE_FUNDO = '#1E7B4A', '#E2F3E8'
VERM_TXT, VERM_FUNDO = '#B42318', '#FDE3E1'
LAR_TXT, LAR_FUNDO = '#B54708', '#FDEBD3'
LOGO_ESCURO = os.path.join(AQUI, 'assets', 'logo_fundo_escuro.png')
FONTE = 'Segoe UI'

LAN = 'Lançamento'
QL = "'Lançamento'!"
BD_MAX = 60000            # linhas da base consideradas nas formulas (agenda / ultimo registro)

# ---------------------------------------------------------------- ensaios
# (codigo, nome na base, titulo na tela, norma, frequencia)
ENSAIOS = [
    ('BLA', 'Blaine', 'Calibração Blaine', 'ISO 21283 · área da superfície específica', 'Início de todo turno'),
    ('TAM', 'Tambor de Abrasão', 'Tambor de abrasão · rotação', 'ISO 3271 · 25 ± 1 rpm', 'Seg, qua e sáb · noite'),
    ('ALP', 'Gran. Fina Alpine', 'Granulometria fina Alpine', 'ISO 4701 · passante', 'Seg, qua e sáb · noite'),
    ('UMI', 'Umidade', 'Umidade · analisadores x estufa', 'ISO 3087 · umidade de lote', 'Ter, qui e dom · noite'),
    ('COM', 'Compressão', 'Compressão · prensas', 'ISO 4700 · resistência à compressão', 'Segunda · noite'),
    ('GRA', 'Granulometria', 'Granulometria · LTF x LCE', 'ISO 4701 · comparativo de peneiramento', 'Segunda · noite'),
    ('T515', 'Tamb 5 kg x 15 kg', 'Tamboramento 5 kg x 15 kg', 'ISO 3271 · LTF x T03', 'Segunda · noite'),
    ('PEN', 'Verificação Peneiradores', 'Verificação dos peneiradores', 'Telas, inclinação, vibração, travamento, calibração',
     'Ter, qui e dom · noite'),
    ('FIS', 'Comparativo Fisher', 'Comparativo Fisher', 'Superfície específica 66PB01 x 66PB09', 'Eventual'),
]
COD = [e[0] for e in ENSAIOS]
NOME = {e[0]: e[1] for e in ENSAIOS}
# Agenda (planilha "Controle"): por dia da semana (seg..dom) x turno (D = 07x19, N = 19x07)
_D, _N = 'D', 'N'
AGENDA = {
    'BLA': {(d, t) for d in range(7) for t in (_D, _N)},
    'TAM': {(0, _N), (2, _N), (5, _N)},
    'ALP': {(0, _N), (2, _N), (5, _N)},
    'UMI': {(1, _N), (3, _N), (6, _N)},
    'COM': {(0, _N)},
    'GRA': {(0, _N)},
    'T515': {(0, _N)},
    'PEN': {(1, _N), (3, _N), (6, _N)},
    'FIS': set(),
}
DIAS = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']

# Peneiradores: (grupo, malha, tag)
PENEIRAS = [('Circulares', m, t) for m, t in [
    ('19', 'PN0423'), ('16', 'PN0482'), ('12,5', 'PN  614'), ('10', 'PN 303'), ('9', 'PN431'), ('8', 'PN432'),
    ('6,3', 'PN0430'), ('5', 'PN0425'), ('3,15', 'PN0,46'), ('0,5', 'PN0500'), ('0,5', 'PN0501')]] + \
    [('Peneirador 09', m, t) for m, t in [
        ('19', 'PN478'), ('16', 'PN516'), ('14', 'PN665'), ('12,5', 'PN479598'), ('9', 'PN510'), ('8', 'PN647'),
        ('6,3', 'PN0459')]] + \
    [('Peneirador 05', m, t) for m, t in [
        ('19', 'PN398'), ('16', 'PN664'), ('14', 'PN665'), ('12,5', 'PN666'), ('9', 'PN646'), ('8', 'PN647'),
        ('6,3', 'PN649')]] + \
    [('Peneirador 06', '6,3', 'PN648'), ('Peneirador 06', '10', 'N/CONT')]
ITENS_PEN = ['Condições das telas', 'Inclinação das telas', 'Vibração do peneirador', 'Travamento das peneiras',
             'Calibração']
STATUS_PEN = ['OK', 'NÃO OK', 'SEM TAG']

RESPONSAVEIS = ['Alexsandro', 'Anderson', 'Cassiano', 'Cleybson', 'Feliciano', 'Felipe', 'Humberto', 'Renato',
                'Wesley']
BACKUP_PADRAO = ('G:\\Grupos de Trabalho\\UBU Gerência de Processos e Automação\\Laboratório\\01 - Lab. Físico\\'
                 '02 - Gerenciamento da Rotina\\AFERIÇÃO - PADRÃO\\Backup')

DADOS = json.load(open(os.path.join(AQUI, 'dados_iniciais.json'), encoding='utf-8'))
LIMITES = DADOS['limites']          # cabecalho + 18 linhas (mesmo conteudo de tbl_Limites)
CHAVES = [l[0] for l in LIMITES[1:]]
BD_COLS = ['Ensaio', 'Tipo', 'Data', 'Equipamento', 'Parâmetro', 'Unidade', 'Valor', 'Referência', 'Diferença',
           'Lim. Inferior', 'Lim. Superior', 'Tolerância', 'Resultado', 'Responsável', 'Letra', 'Observação',
           'Origem']


def crit(chave, campo):
    """Nome definido com o criterio resolvido (vazio quando nao ha valor). campo: NOM, LI, LS, TOL."""
    return 'c_%s_%s' % (chave, campo)


def nf(dec):
    return '0' if dec == 0 else '0.' + '0' * dec


# ============================================================================
class Construtor:
    def __init__(self, caminho):
        self.wb = xlsxwriter.Workbook(caminho)
        self.wb.set_vba_name('ThisWorkbook')
        self.wb.set_properties({'title': 'Controle de Calibração e Aferição - Laboratório Físico',
                                'company': 'Samarco - Laboratório Físico'})
        self.wb.set_calc_mode('auto')
        self._fmts = {}
        self.stg = []           # linhas da _Staging (dicts)
        self.inputs = {}        # codigo do ensaio -> lista de enderecos de entrada (para limpar e contar)
        self.resultados = {}    # codigo -> lista de celulas de resultado (Conforme / Nao conforme)
        self.secoes = {}        # codigo -> linha do titulo da secao
        self.combos = []        # (ensaio, equipamento, parametro, unidade, criterio, dec) para o Painel
        self.series = {}        # graficos: codigo -> lista de (equipamento, parametro)

    # ----------------------------------------------------------- formatos
    def f(self, **kw):
        base = dict(font_name=FONTE, font_size=10, font_color=TEXTO, valign='vcenter')
        base.update(kw)
        chave = tuple(sorted(base.items()))
        if chave not in self._fmts:
            self._fmts[chave] = self.wb.add_format(base)
        return self._fmts[chave]

    def fin(self, **kw):
        base = dict(bg_color=FUNDO_INPUT, border=1, border_color='#E6C76A', locked=False, align='center',
                    font_size=11, bold=True, font_color=AZUL_PROFUNDO)
        base.update(kw)
        return self.f(**base)

    def fcalc(self, **kw):
        base = dict(border=1, border_color=BORDA, align='center', font_size=10.5, font_color=TEXTO)
        base.update(kw)
        return self.f(**base)

    def fref(self, **kw):
        base = dict(border=1, border_color=BORDA, align='center', font_size=9.5, font_color=AZUL_ACINZ,
                    bg_color=FUNDO_CLARO)
        base.update(kw)
        return self.f(**base)

    def flbl(self, **kw):
        base = dict(border=1, border_color=BORDA, font_size=10, bold=True, font_color=AZUL_TITULO,
                    bg_color=FUNDO_CLARO, indent=1)
        base.update(kw)
        return self.f(**base)

    def fhdr(self, **kw):
        base = dict(bold=True, font_size=9, font_color=BRANCO, bg_color=AZUL_MEDIO, align='center',
                    text_wrap=True, border=1, border_color=BRANCO)
        base.update(kw)
        return self.f(**base)

    def fres(self):
        return self.f(border=1, border_color=BORDA, align='center', bold=True, font_size=10.5,
                      font_color=TEXTO_SEC)

    def cf_status(self, ws, r1, c1, r2, c2):
        """Cores dos textos de situacao (Conforme, Nao conforme, Informativo, Pendente...)."""
        regras = [('Não conforme', VERM_TXT, VERM_FUNDO), ('Conforme', VERDE_TXT, VERDE_FUNDO),
                  ('Pendente', LAR_TXT, LAR_FUNDO), ('Atrasado', LAR_TXT, LAR_FUNDO),
                  ('Previsto', AZUL_TITULO, FUNDO_AZUL_CLARO), ('Preenchido', VERDE_TXT, VERDE_FUNDO),
                  ('Em dia', VERDE_TXT, VERDE_FUNDO), ('Informativo', AZUL_ACINZ, FUNDO_GRUPO)]
        for txt, cor, fundo in regras:
            fm = self.wb.add_format({'font_color': cor, 'bg_color': fundo, 'bold': True})
            ws.conditional_format(r1, c1, r2, c2, {'type': 'text', 'criteria': 'begins with', 'value': txt,
                                                   'format': fm})

    def cel(self, ws, r, c1, c2, valor, fmt, formula=False, resultado=None):
        if c2 > c1:
            ws.merge_range(r, c1, r, c2, '', fmt)
        if formula:
            ws.write_formula(r, c1, valor, fmt, resultado if resultado is not None else '')
        elif valor is None or valor == '':
            ws.write_blank(r, c1, None, fmt)
        else:
            ws.write(r, c1, valor, fmt)

    # ----------------------------------------------------------- blocos visuais
    def cabecalho(self, ws, ultima_col, titulo, subtitulo, col_logo_fim=3, col_ini=1, r0=0):
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

    def faixa(self, ws, r, c1, c2, texto, altura=24, cor=AZUL_TITULO, tamanho=11.5):
        ws.set_row(r, altura)
        ws.merge_range(r, c1, r, c2, texto, self.f(bold=True, font_size=tamanho, font_color=BRANCO, bg_color=cor,
                                                   indent=1))

    def botao(self, ws, row, col, texto, macro, largura=200, altura=40, estilo='primario', x=0, y=0,
              tamanho=None):
        estilos = {'primario': (AZUL, BRANCO, None, 11), 'destaque': (AMARELO, AZUL_PROFUNDO, None, 12),
                   'claro': (BRANCO, AZUL, AZUL, 10), 'alerta': (BRANCO, VERM_TXT, VERM_TXT, 10)}
        fundo, cor, borda, tam = estilos[estilo]
        ws.insert_textbox(row, col, texto, {
            'width': largura, 'height': altura, 'x_offset': x, 'y_offset': y,
            'font': {'name': FONTE, 'size': tamanho or tam, 'bold': True, 'color': cor},
            'align': {'vertical': 'middle', 'horizontal': 'center'},
            'fill': {'color': fundo},
            'line': {'color': borda, 'width': 1} if borda else {'none': True},
            'description': 'macro:' + macro, 'object_position': 3})

    def nome(self, nome, ws, r1, c1, r2=None, c2=None):
        ref = rc(r1, c1, True, True)
        if r2 is not None:
            ref += ':' + rc(r2, c2, True, True)
        self.wb.define_name(nome, "='%s'!%s" % (ws.name, ref))

    def link(self, ws, r, c, destino, texto, tamanho=9, fmt=None):
        ws.write_url(r, c, 'internal:' + destino, fmt or self.f(font_size=tamanho, bold=True,
                                                                font_color=AZUL_MEDIO, underline=1), string=texto)

    # ============================================================ estrutura
    def construir(self):
        wb = self.wb
        self.ws_lan = wb.add_worksheet(LAN)
        self.ws_pai = wb.add_worksheet('Painel')
        self.ws_bd = wb.add_worksheet('BD_Afericoes')
        self.ws_lim = wb.add_worksheet('BD_Limites')
        self.ws_cfg = wb.add_worksheet('Config')
        self.ws_stg = wb.add_worksheet('_Staging')
        self.ws_gra = wb.add_worksheet('_Graficos')
        for ws, cn in [(self.ws_lan, 'shLancamento'), (self.ws_pai, 'shPainel'), (self.ws_bd, 'shBD'),
                       (self.ws_lim, 'shLimites'), (self.ws_cfg, 'shConfig'), (self.ws_stg, 'shStaging'),
                       (self.ws_gra, 'shGraficos')]:
            ws.set_vba_name(cn)
            ws.hide_gridlines(2)
        self.ws_lan.set_tab_color(OURO)
        self.ws_pai.set_tab_color(AZUL)
        self.aba_limites()
        self.aba_config_listas()
        self.aba_lancamento()
        self.aba_config_resto()
        self.aba_staging()
        self.aba_bd()
        self.aba_graficos()
        self.aba_painel()
        for ws in (self.ws_bd, self.ws_lim, self.ws_cfg):
            ws.hide()
        self.ws_stg.very_hidden() if hasattr(self.ws_stg, 'very_hidden') else self.ws_stg.hide()
        self.ws_gra.very_hidden() if hasattr(self.ws_gra, 'very_hidden') else self.ws_gra.hide()
        self.ws_lan.activate()

    # ------------------------------------------------------------ BD_Limites (criterios)
    def aba_limites(self):
        ws = self.ws_lim
        larg = [10, 22, 26, 9, 10, 12, 12, 11, 70]
        for i, w in enumerate(larg):
            ws.set_column(i, i, w)
        n = len(LIMITES) - 1
        dados = []
        for l in LIMITES[1:]:
            dados.append([v if v is not None else None for v in l])
        num = self.f(num_format='0.###', align='center', bg_color=FUNDO_INPUT, border=1, border_color='#E6C76A')
        ws.add_table(0, 0, n, 8, {
            'name': 'tbl_Limites', 'style': 'Table Style Medium 2',
            'columns': [{'header': h} for h in LIMITES[0]], 'data': dados})
        for r in range(1, n + 1):
            for c in range(4, 8):
                v = LIMITES[r][c]
                if v is None:
                    ws.write_blank(r, c, None, num)
                else:
                    ws.write_number(r, c, v, num)
        ws.write(n + 2, 0, 'COMO FUNCIONA', self.f(bold=True, font_size=11, font_color=AZUL_TITULO))
        texto = ('Critérios de aceitação das calibrações, aferições e comparativos.\n'
                 ' Altere aqui o valor nominal, os limites ou a tolerância (células em amarelo). O Lançamento e o '
                 'Painel passam a usar o novo valor na hora.\n'
                 ' Resultado: com limites, Conforme quando Lim. Inferior <= valor <= Lim. Superior; com tolerância, '
                 'Conforme quando |valor - referência| <= tolerância.\n'
                 ' As linhas V_TA.. guardam o número de voltas de cada tambor (usado no cálculo do rpm).\n'
                 ' Não renomeie a coluna Chave nem apague linhas: o Lançamento procura o critério pela chave.\n'
                 ' Esta tabela (tbl_Limites) e a tabela tbl_Afericoes (aba BD_Afericoes) são lidas pelo Power BI.')
        ws.merge_range(n + 3, 0, n + 10, 8, texto, self.f(text_wrap=True, valign='top', font_size=10,
                                                          bg_color=FUNDO_CLARO, border=1, border_color=BORDA))
        ws.freeze_panes(1, 0)
        self.wb.define_name('crit_chave', "=BD_Limites!$A$2:$A$200")
        for campo, col in (('NOM', 'E'), ('LI', 'F'), ('LS', 'G'), ('TOL', 'H')):
            self.wb.define_name('crit_' + campo, "=BD_Limites!$%s$2:$%s$200" % (col, col))

    # ------------------------------------------------------------ Config: listas e criterios
    def aba_config_listas(self):
        ws = self.ws_cfg
        ws.set_column(0, 0, 2)
        ws.set_column(1, 30, 13)
        self.cabecalho(ws, 18, 'CONFIGURAÇÕES', 'LISTAS, AGENDA SEMANAL E BACKUP', col_logo_fim=3)
        hdr = self.fhdr()
        listas = [('lstResp', 'Responsáveis', RESPONSAVEIS + [''] * 12),
                  ('lstLetra', 'Letras', ['A', 'B', 'C', 'D']),
                  ('lstTurno', 'Turnos', ['07x19', '19x07']),
                  ('lstPen', 'Peneiradores', STATUS_PEN)]
        inp = self.f(bg_color=FUNDO_INPUT, border=1, border_color=BORDA, locked=False)
        for i, (nm, tit, itens) in enumerate(listas):
            c = 1 + i
            ws.write(5, c, tit, hdr)
            for j, it in enumerate(itens):
                ws.write(6 + j, c, it, inp)
            if nm == 'lstResp':
                self.wb.define_name(nm, "=OFFSET(Config!$B$7,0,0,MAX(1,COUNTA(Config!$B$7:$B$%d)),1)"
                                    % (6 + len(itens)))
            else:
                self.nome(nm, ws, 6, c, 6 + len(itens) - 1, c)
        # criterios resolvidos (nomes c_CHAVE_CAMPO)
        r0 = 30
        ws.merge_range(r0 - 1, 1, r0 - 1, 6, 'Critérios em uso (lidos da aba BD_Limites)', self.flbl())
        for j, h in enumerate(['Chave', 'Nominal', 'Lim. inf.', 'Lim. sup.', 'Tolerância']):
            ws.write(r0, 1 + j, h, hdr)
        for i, ch in enumerate(CHAVES):
            r = r0 + 1 + i
            ws.write(r, 1, ch, self.f(bold=True))
            for j, (campo, faixa) in enumerate((('NOM', 'crit_NOM'), ('LI', 'crit_LI'), ('LS', 'crit_LS'),
                                               ('TOL', 'crit_TOL'))):
                idx = 'INDEX(%s,MATCH("%s",crit_chave,0))' % (faixa, ch)
                lim = LIMITES[1 + i][4 + j]
                ws.write_formula(r, 2 + j, '=IFERROR(IF(ISNUMBER(%s),%s,""),"")' % (idx, idx),
                                 self.f(num_format='0.###', align='center'), lim if lim is not None else '')
                self.nome(crit(ch, campo), ws, r, 2 + j)

    # ------------------------------------------------------------ Lancamento
    def aba_lancamento(self):
        ws = self.ws_lan
        self.UC = UC = 15
        ws.set_column(0, 0, 1.5)
        ws.set_column(1, 1, 1.2)
        ws.set_column(2, UC, 10.2)
        ws.set_column(UC + 1, UC + 1, 2.5)
        ws.set_column(UC + 2, UC + 3, 27)
        self.cabecalho(ws, UC, 'CONTROLE DE CALIBRAÇÃO E AFERIÇÃO', 'LABORATÓRIO FÍSICO  |  LANÇAMENTO', col_logo_fim=5)
        lbl = self.flbl()
        # ---- identificacao (linha 5)
        ws.set_row(4, 8)
        r = 5
        ws.set_row(r, 26)
        self.cel(ws, r, 2, 2, 'Data', lbl)
        self.cel(ws, r, 3, 4, '', self.fin(num_format='dd/mm/yyyy', font_size=12))
        self.nome('pData', ws, r, 3)
        ws.data_validation(r, 3, r, 3, {'validate': 'date', 'criteria': '>', 'value': dtm.date(2020, 1, 1),
                                        'error_message': 'Digite a data (dd/mm/aaaa).'})
        self.cel(ws, r, 5, 5, 'Turno', lbl)
        self.cel(ws, r, 6, 6, '', self.fin(font_size=12))
        self.nome('pTurno', ws, r, 6)
        ws.data_validation(r, 6, r, 6, {'validate': 'list', 'source': '=lstTurno'})
        self.cel(ws, r, 7, 7, 'Letra', lbl)
        self.cel(ws, r, 8, 8, '', self.fin(font_size=12))
        self.nome('pLetra', ws, r, 8)
        ws.data_validation(r, 8, r, 8, {'validate': 'list', 'source': '=lstLetra'})
        self.cel(ws, r, 9, 10, 'Responsável', lbl)
        self.cel(ws, r, 11, UC, '', self.fin(font_size=12, align='left', indent=1))
        self.nome('pResp', ws, r, 11)
        ws.data_validation(r, 11, r, 11, {'validate': 'list', 'source': '=lstResp', 'error_type': 'information',
                                          'error_message': 'Nome fora da lista: confirme para usar mesmo assim.'})
        r += 1
        ws.set_row(r, 18)
        ws.merge_range(r, 2, r, UC, '', self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
        ws.write_formula(r, 2, '=IF(cfgUltimoLanc="","Nenhum lançamento registrado nesta planilha.",'
                               '"Último lançamento: "&cfgUltimoLanc)',
                         self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1), '')
        ws.freeze_panes(r + 1, 0)
        # botoes (area congelada, a direita)
        self.botao(ws, 1, UC + 2, 'Registrar lançamento', 'RegistrarLancamento', 190, 44, 'destaque', x=4, y=4)
        self.botao(ws, 1, UC + 3, 'Ver painel', 'IrPainel', 190, 44, 'primario', x=4, y=4)
        self.botao(ws, 4, UC + 2, 'Limpar tela', 'LimparLancamento', 190, 30, 'claro', x=4, y=6)
        self.botao(ws, 4, UC + 3, 'Desfazer último lançamento', 'DesfazerUltimo', 190, 30, 'alerta', x=4, y=6)

        # ---- agenda do turno
        r += 2
        self.faixa(ws, r, 2, UC, 'AGENDA DO TURNO')
        self.r_agenda = r
        r += 1
        ws.set_row(r, 30)
        hdr = self.fhdr()
        for c1, c2, t in [(2, 5, 'Ensaio'), (6, 7, 'Frequência'), (8, 9, 'Neste turno'), (10, 11, 'Nesta tela'),
                          (12, 13, 'Resultado na tela'), (14, UC, 'Último registro')]:
            self.cel(ws, r, c1, c2, t, hdr)
        self.r_agenda1 = r + 1
        r += 1 + len(ENSAIOS)
        ws.set_row(r, 16)
        ws.merge_range(r, 2, r, UC, 'Clique no nome do ensaio para ir ao formulário. Ensaios sem dados nesta tela '
                       'não são registrados.', self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
        r += 2

        # ---- secoes
        r = self.sec_blaine(ws, r)
        r = self.sec_tambor(ws, r)
        r = self.sec_alpine(ws, r)
        r = self.sec_umidade(ws, r)
        r = self.sec_compressao(ws, r)
        r = self.sec_granulometria(ws, r)
        r = self.sec_t515(ws, r)
        r = self.sec_peneiradores(ws, r)
        r = self.sec_fisher(ws, r)
        self.fim_lan = r
        # agenda (precisa das linhas das secoes)
        self.preencher_agenda(ws)
        ws.print_area(1, 1, r, UC)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.protect('', {'format_columns': True, 'format_rows': True, 'select_locked_cells': True,
                        'select_unlocked_cells': True})

    # -------- auxiliares das secoes
    def titulo_secao(self, ws, r, cod, num):
        e = ENSAIOS[COD.index(cod)]
        self.secoes[cod] = r
        ws.set_row(r, 28)
        ws.write_blank(r, 1, None, self.f(bg_color=AMARELO))
        ws.merge_range(r, 2, r, 10, '%d.  %s' % (num, e[2].upper()),
                       self.f(bold=True, font_size=12.5, font_color=BRANCO, bg_color=AZUL, indent=1))
        ws.merge_range(r, 11, r, self.UC, e[4], self.f(font_size=9, bold=True, font_color=AMARELO, bg_color=AZUL,
                                                       align='right', indent=1))
        ws.set_row(r + 1, 16)
        ws.merge_range(r + 1, 2, r + 1, self.UC, e[3], self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC,
                                                              indent=1, bg_color=FUNDO_CLARO))
        self.nome('sec_' + cod, ws, r, 2)
        self.inputs.setdefault(cod, [])
        self.resultados.setdefault(cod, [])
        return r + 2

    def cabec(self, ws, r, cols, altura=30):
        ws.set_row(r, altura)
        for c1, c2, t in cols:
            self.cel(ws, r, c1, c2, t, self.fhdr())
        return r + 1

    def entrada(self, ws, cod, r, c1, c2, fmt=None, valor=''):
        self.cel(ws, r, c1, c2, valor, fmt or self.fin())
        self.inputs[cod].append(rc(r, c1, True, True) if c2 == c1 else
                                '%s:%s' % (rc(r, c1, True, True), rc(r, c2, True, True)))
        return rc(r, c1, True, True)

    def resultado(self, ws, cod, r, c1, c2, formula):
        self.cel(ws, r, c1, c2, formula, self.fres(), formula=True)
        self.resultados[cod].append(rc(r, c1, True, True))
        self.cf_status(ws, r, c1, r, c2)
        return rc(r, c1, True, True)

    def obs(self, ws, cod, r):
        ws.set_row(r, 22)
        self.cel(ws, r, 2, 3, 'Observação', self.flbl())
        ref = self.entrada(ws, cod, r, 4, self.UC, self.fin(align='left', bold=False, font_size=10, indent=1,
                                                          text_wrap=True))
        return ref

    @staticmethod
    def f_faixa(v, li, ls):
        """Resultado por limites (vazio = sem limite)."""
        return ('=IF({v}="","",IF(AND(OR({li}="",{v}>={li}),OR({ls}="",{v}<={ls})),"Conforme","Não conforme"))'
                .format(v=v, li=li, ls=ls))

    @staticmethod
    def f_tol(dif, tol):
        return '=IF(OR({d}="",{t}=""),"",IF(ROUND(ABS({d}),6)<=ROUND({t},6),"Conforme","Não conforme"))'.format(
            d=dif, t=tol)

    def L(self, ref):
        """Referencia absoluta a uma celula da aba Lancamento (para a _Staging)."""
        return QL + ref

    def add_stg(self, cod, tipo, equip, param, unid, valor, refv, dif, li, ls, tol, res, obs, dec=2,
                criterio='', serie=True):
        """Uma linha candidata para a base. Todos os argumentos sao formulas/refs do ponto de vista da _Staging
        (sem '='), ou textos entre aspas."""
        self.stg.append(dict(cod=cod, ensaio='"%s"' % NOME[cod], tipo='"%s"' % tipo, equip=equip, param=param,
                             unid='"%s"' % unid, valor=valor, refv=refv, dif=dif, li=li, ls=ls, tol=tol, res=res,
                             obs=obs))
        eq = equip.strip('"') if equip.startswith('"') else None
        pa = param.strip('"') if param.startswith('"') else None
        if cod != 'PEN' and eq is not None and pa is not None:
            self.combos.append((NOME[cod], eq, pa, unid, criterio, dec))

    # -------- 1. Blaine
    def sec_blaine(self, ws, r):
        cod = 'BLA'
        r = self.titulo_secao(ws, r, cod, 1)
        r = self.cabec(ws, r, [(2, 3, 'Equipamento'), (4, 6, 'Parâmetro'), (7, 8, 'Resultado (cm²/g)'),
                               (9, 9, 'Nominal'), (10, 10, 'Lim. inf.'), (11, 11, 'Lim. sup.'),
                               (12, 12, 'Diferença'), (13, self.UC, 'Situação')])
        linhas = [('66PB08', 'Blaine Manual', 'BLAM'), ('Blaine Star', 'Calibração Star', 'BLAS'),
                  ('Blaine Automático', 'Blaine Automático', 'BLAA')]
        refs = []
        for equip, param, ch in linhas:
            ws.set_row(r, 24)
            if equip == '66PB08':
                eqref = self.entrada_fixa(ws, r, 2, 3, equip)
            else:
                self.cel(ws, r, 2, 3, equip, self.flbl())
                eqref = None
            self.cel(ws, r, 4, 6, param, self.f(border=1, border_color=BORDA, indent=1))
            v = self.entrada(ws, cod, r, 7, 8, self.fin(num_format='0'))
            nom, li, ls = crit(ch, 'NOM'), crit(ch, 'LI'), crit(ch, 'LS')
            self.cel(ws, r, 9, 9, '=%s' % nom, self.fref(num_format='0'), formula=True)
            self.cel(ws, r, 10, 10, '=%s' % li, self.fref(num_format='0'), formula=True)
            self.cel(ws, r, 11, 11, '=%s' % ls, self.fref(num_format='0'), formula=True)
            d = rc(r, 12, True, True)
            self.cel(ws, r, 12, 12, '=IF(OR(%s="",%s=""),"",%s-%s)' % (v, nom, v, nom), self.fcalc(num_format='0'),
                     formula=True)
            res = self.resultado(ws, cod, r, 13, self.UC, self.f_faixa(v, li, ls))
            refs.append((equip, param, ch, v, d, res, eqref))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for equip, param, ch, v, d, res, eqref in refs:
            eq = ('IF(TRIM(%s)="","66PB08",TRIM(%s))' % (self.L(eqref), self.L(eqref))) if eqref else '"%s"' % equip
            self.add_stg(cod, 'Calibração', eq if eqref is None else eq, '"%s"' % param, 'cm²/g', self.L(v),
                         crit(ch, 'NOM'), self.L(d), crit(ch, 'LI'), crit(ch, 'LS'), '""', self.L(res),
                         'pTurno&IF(TRIM(%s)="",""," - "&TRIM(%s))' % (self.L(o), self.L(o)), dec=0,
                         criterio='faixa:' + ch)
            if eqref:
                self.combos.append((NOME[cod], '66PB08', param, 'cm²/g', 'faixa:' + ch, 0))
        self.series[cod] = [(e, p) for e, p, _ in linhas]
        return r + 1

    def entrada_fixa(self, ws, r, c1, c2, valor):
        """Campo editavel com valor padrao (nao e apagado pelo Limpar)."""
        self.cel(ws, r, c1, c2, valor, self.fin(font_size=10, font_color=AZUL_TITULO, bg_color='#FFFDF5'))
        return rc(r, c1, True, True)

    # -------- 2. Tambor de abrasao (rpm)
    def sec_tambor(self, ws, r):
        cod = 'TAM'
        r = self.titulo_secao(ws, r, cod, 2)
        r = self.cabec(ws, r, [(2, 3, 'Tambor'), (4, 4, 'Voltas'), (5, 5, 'Minutos'), (6, 6, 'Segundos'),
                               (7, 7, 'Tempo (s)'), (8, 9, 'Rotação (rpm)'), (10, 10, 'Nominal'),
                               (11, 11, 'Lim. inf.'), (12, 12, 'Lim. sup.'), (13, self.UC, 'Situação')])
        tambores = [('66TA05', 'V_TA05'), ('66TA08', 'V_TA08'), ('66TA09', 'V_TA09'), ('66TA06', 'V_TA06'),
                    ('66TA07', 'V_TA07')]
        refs = []
        for tag, chv in tambores:
            ws.set_row(r, 24)
            self.cel(ws, r, 2, 3, tag, self.flbl())
            self.cel(ws, r, 4, 4, '=%s' % crit(chv, 'NOM'), self.fref(num_format='0'), formula=True)
            vol = rc(r, 4, True, True)
            mi = self.entrada(ws, cod, r, 5, 5, self.fin(num_format='0'))
            se = self.entrada(ws, cod, r, 6, 6, self.fin(num_format='0'))
            t = rc(r, 7, True, True)
            self.cel(ws, r, 7, 7, '=IF(AND(%s="",%s=""),"",N(%s)*60+N(%s))' % (mi, se, mi, se),
                     self.fcalc(num_format='0'), formula=True)
            rpm = rc(r, 8, True, True)
            self.cel(ws, r, 8, 9, '=IF(OR(%s="",N(%s)=0,%s=""),"",%s*60/%s)' % (t, t, vol, vol, t),
                     self.fcalc(num_format='0.00', bold=True, font_size=11.5), formula=True)
            for c, campo in ((10, 'NOM'), (11, 'LI'), (12, 'LS')):
                self.cel(ws, r, c, c, '=%s' % crit('TAMB', campo), self.fref(num_format='0.0'), formula=True)
            res = self.resultado(ws, cod, r, 13, self.UC, self.f_faixa(rpm, crit('TAMB', 'LI'), crit('TAMB', 'LS')))
            refs.append((tag, rpm, res))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for tag, rpm, res in refs:
            self.add_stg(cod, 'Calibração', '"%s"' % tag, '"Rotação"', 'rpm', self.L(rpm), crit('TAMB', 'NOM'),
                         'IF(OR(%s="",%s=""),"",%s-%s)' % (self.L(rpm), crit('TAMB', 'NOM'), self.L(rpm),
                                                           crit('TAMB', 'NOM')),
                         crit('TAMB', 'LI'), crit('TAMB', 'LS'), '""', self.L(res), 'TRIM(%s)' % self.L(o),
                         criterio='faixa:TAMB')
        self.series[cod] = [(t, 'Rotação') for t, _ in tambores]
        return r + 1

    # -------- 3. Alpine
    def sec_alpine(self, ws, r):
        cod = 'ALP'
        r = self.titulo_secao(ws, r, cod, 3)
        r = self.cabec(ws, r, [(2, 3, 'Equipamento'), (4, 6, 'Peneira'), (7, 8, 'Passante (%)'), (9, 9, 'Nominal'),
                               (10, 10, 'Lim. inf.'), (11, 11, 'Lim. sup.'), (12, 12, 'Diferença'),
                               (13, self.UC, 'Situação')])
        eqs = [('66AG09', 'Passante (peneira 66PN670)', '66PN670'), ('66AG10', 'Passante (peneira 66PN671)', '66PN671'),
               ('66AG11', 'Passante (peneira 66PN671)', '66PN671')]
        refs = []
        for tag, param, pen in eqs:
            ws.set_row(r, 24)
            self.cel(ws, r, 2, 3, tag, self.flbl())
            self.cel(ws, r, 4, 6, pen, self.f(border=1, border_color=BORDA, indent=1))
            v = self.entrada(ws, cod, r, 7, 8, self.fin(num_format='0.00'))
            for c, campo in ((9, 'NOM'), (10, 'LI'), (11, 'LS')):
                self.cel(ws, r, c, c, '=%s' % crit('ALP', campo), self.fref(num_format='0.0'), formula=True)
            d = rc(r, 12, True, True)
            self.cel(ws, r, 12, 12, '=IF(OR(%s="",%s=""),"",%s-%s)' % (v, crit('ALP', 'NOM'), v, crit('ALP', 'NOM')),
                     self.fcalc(num_format='0.00'), formula=True)
            res = self.resultado(ws, cod, r, 13, self.UC, self.f_faixa(v, crit('ALP', 'LI'), crit('ALP', 'LS')))
            refs.append((tag, param, v, d, res))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for tag, param, v, d, res in refs:
            self.add_stg(cod, 'Calibração', '"%s"' % tag, '"%s"' % param, '%', self.L(v), crit('ALP', 'NOM'),
                         self.L(d), crit('ALP', 'LI'), crit('ALP', 'LS'), '""', self.L(res), 'TRIM(%s)' % self.L(o),
                         criterio='faixa:ALP')
        self.series[cod] = [(t, p) for t, p, _ in eqs]
        return r + 1

    # -------- 4. Umidade
    def sec_umidade(self, ws, r):
        cod = 'UMI'
        r = self.titulo_secao(ws, r, cod, 4)
        lbl = self.flbl()
        ws.set_row(r, 24)
        self.cel(ws, r, 2, 4, 'Estufa 100 °C · umidade (%)', lbl)
        direto = self.entrada(ws, cod, r, 5, 6, self.fin(num_format='0.00'))
        self.cel(ws, r, 7, self.UC, 'ou informe as pesagens da estufa na linha abaixo (o cálculo é automático)',
                 self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
        r += 1
        ws.set_row(r, 24)
        self.cel(ws, r, 2, 4, 'Pesagens da estufa (g)', lbl)
        self.cel(ws, r, 5, 5, 'Tara', self.fref())
        tara = self.entrada(ws, cod, r, 6, 6, self.fin(num_format='0.00'))
        self.cel(ws, r, 7, 7, 'P. inicial', self.fref())
        pi = self.entrada(ws, cod, r, 8, 8, self.fin(num_format='0.00'))
        self.cel(ws, r, 9, 9, 'P. final', self.fref())
        pf = self.entrada(ws, cod, r, 10, 10, self.fin(num_format='0.00'))
        self.cel(ws, r, 11, 12, 'Estufa usada (%)', self.fref(bold=True))
        est = rc(r, 13, True, True)
        self.cel(ws, r, 13, self.UC,
                 '=IF(ISNUMBER(%s),%s,IF(AND(ISNUMBER(%s),ISNUMBER(%s),ISNUMBER(%s)),IF(%s>%s,(%s-%s)/(%s-%s)*100,""),""))'
                 % (direto, direto, tara, pi, pf, pi, tara, pi, pf, pi, tara),
                 self.fcalc(num_format='0.00', bold=True, font_size=11.5, bg_color=FUNDO_AZUL_CLARO), formula=True)
        r += 1
        r = self.cabec(ws, r, [(2, 4, 'Analisador'), (5, 6, 'Umidade (%)'), (7, 8, 'Estufa (%)'),
                               (9, 9, 'Diferença'), (10, 10, 'Tolerância'), (11, self.UC, 'Situação')])
        refs = []
        for tag in ('66AN10', '66AN11', '66AN12'):
            ws.set_row(r, 24)
            self.cel(ws, r, 2, 4, tag, lbl)
            v = self.entrada(ws, cod, r, 5, 6, self.fin(num_format='0.00'))
            self.cel(ws, r, 7, 8, '=IF(%s="","",%s)' % (est, est), self.fref(num_format='0.00'), formula=True)
            d = rc(r, 9, True, True)
            self.cel(ws, r, 9, 9, '=IF(OR(%s="",%s=""),"",%s-%s)' % (v, est, v, est), self.fcalc(num_format='0.00'),
                     formula=True)
            self.cel(ws, r, 10, 10, '=%s' % crit('UMI', 'TOL'), self.fref(num_format='0.00'), formula=True)
            res = self.resultado(ws, cod, r, 11, self.UC, self.f_tol(d, crit('UMI', 'TOL')))
            refs.append((tag, v, d, res))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for tag, v, d, res in refs:
            self.add_stg(cod, 'Comparativo', '"%s x Estufa"' % tag, '"Umidade"', '%', self.L(v),
                         'IF(%s="","",%s)' % (self.L(est), self.L(est)), self.L(d), '""', '""', crit('UMI', 'TOL'),
                         self.L(res), 'TRIM(%s)' % self.L(o), criterio='tol:UMI')
        self.series[cod] = [('%s x Estufa' % t, 'Umidade') for t in ('66AN10', '66AN11', '66AN12')]
        return r + 1

    # -------- 5. Compressao
    def sec_compressao(self, ws, r):
        cod = 'COM'
        r = self.titulo_secao(ws, r, cod, 5)
        prensas = ['66PS04', '66PS07', '66PS08', '66PS10', '66PS05']
        lbl = self.flbl()
        todos = []
        for fx, ch, nomefx in (('16', 'COM16', 'FX -16,0 +12,5 mm'), ('12', 'COM12', 'FX -12,5 +10,0 mm')):
            ws.set_row(r, 20)
            ws.merge_range(r, 2, r, 10, nomefx, self.f(bold=True, font_size=10.5, font_color=AZUL_TITULO,
                                                       bg_color=FUNDO_GRUPO, indent=1, border=1, border_color=BORDA))
            ws.merge_range(r, 11, r, self.UC, '', self.f(font_size=8.5, font_color=AZUL_ACINZ, bg_color=FUNDO_GRUPO,
                                                         align='right', indent=1, border=1, border_color=BORDA))
            ws.write_formula(r, 11, '="Nominal "&TEXT(%s,"0")&" ± "&TEXT(%s,"0")&" kgf/pel  |  velocidade "&'
                                    'TEXT(%s,"0")&" ± "&TEXT(%s-%s,"0")&" mm/min"'
                             % (crit(ch, 'NOM'), crit(ch, 'TOL'), crit('VEL', 'NOM'), crit('VEL', 'NOM'),
                                crit('VEL', 'LI')),
                             self.f(font_size=8.5, font_color=AZUL_ACINZ, bg_color=FUNDO_GRUPO, align='right',
                                    indent=1, border=1, border_color=BORDA), '')
            r += 1
            r = self.cabec(ws, r, [(2, 3, 'Prensa'), (4, 5, 'Resistência (kgf/pel)'), (6, 6, 'DP'),
                                   (7, 7, 'Diferença'), (8, 9, 'Situação'), (10, 11, 'Velocidade (mm/min)'),
                                   (12, self.UC, 'Situação velocidade')])
            for p in prensas:
                ws.set_row(r, 23)
                self.cel(ws, r, 2, 3, p, lbl)
                v = self.entrada(ws, cod, r, 4, 5, self.fin(num_format='0'))
                dp = self.entrada(ws, cod, r, 6, 6, self.fin(num_format='0', font_size=10, bold=False))
                d = rc(r, 7, True, True)
                self.cel(ws, r, 7, 7, '=IF(%s="","",%s-%s)' % (v, v, crit(ch, 'NOM')), self.fcalc(num_format='0'),
                         formula=True)
                res = self.resultado(ws, cod, r, 8, 9, self.f_tol(d, crit(ch, 'TOL')))
                if p == '66PS10':
                    self.cel(ws, r, 10, 11, 'não se aplica', self.fref(italic=True))
                    self.cel(ws, r, 12, self.UC, '', self.fref())
                    vel = rv = None
                else:
                    vel = self.entrada(ws, cod, r, 10, 11, self.fin(num_format='0.0'))
                    rv = self.resultado(ws, cod, r, 12, self.UC, self.f_faixa(vel, crit('VEL', 'LI'),
                                                                              crit('VEL', 'LS')))
                todos.append((fx, ch, nomefx, p, v, dp, d, res, vel, rv))
                r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for fx, ch, nomefx, p, v, dp, d, res, vel, rv in todos:
            obs = 'IF(%s="","","DP "&%s)&IF(AND(%s<>"",TRIM(%s)<>"")," - ","")&TRIM(%s)' % (
                self.L(dp), self.L(dp), self.L(dp), self.L(o), self.L(o))
            self.add_stg(cod, 'Comparativo', '"%s"' % p, '"Resistência %s"' % nomefx, 'kgf/pel', self.L(v),
                         crit(ch, 'NOM'), self.L(d), '""', '""', crit(ch, 'TOL'), self.L(res), obs, dec=0,
                         criterio='tol:' + ch)
            if vel:
                self.add_stg(cod, 'Verificação', '"%s"' % p, '"Velocidade %s"' % nomefx, 'mm/min', self.L(vel),
                             crit('VEL', 'NOM'), 'IF(%s="","",%s-%s)' % (self.L(vel), self.L(vel), crit('VEL', 'NOM')),
                             crit('VEL', 'LI'), crit('VEL', 'LS'), '""', self.L(rv), 'TRIM(%s)' % self.L(o), dec=1,
                             criterio='faixa:VEL')
        self.series[cod] = [(p, 'Resistência FX -16,0 +12,5 mm') for p in prensas]
        return r + 1

    # -------- 6. Granulometria LTF x LCE
    def sec_granulometria(self, ws, r):
        cod = 'GRA'
        r = self.titulo_secao(ws, r, cod, 6)
        r = self.cabec(ws, r, [(2, 4, 'Fração (mm)'), (5, 6, 'LTF 66PN09 (g)'), (7, 7, 'LTF (%)'),
                               (8, 9, 'LCE 66PN08 (g)'), (10, 10, 'LCE (%)'), (11, 11, 'Diferença'),
                               (12, self.UC, '')])
        fr = ['+19,0', '-19,0 +16,0', '-16,0 +14,0', '-14,0 +12,5', '-12,5 +9,0', '-9,0 +8,0', '-8,0 +6,3',
              '-6,3 +5,0', '-5,0 +0,5', '-0,5']
        r1 = r
        ltf, lce, pl, pc = [], [], [], []
        for i, f in enumerate(fr):
            ws.set_row(r, 20)
            self.cel(ws, r, 2, 4, f, self.flbl(font_size=9.5))
            ltf.append(self.entrada(ws, cod, r, 5, 6, self.fin(num_format='0.0', font_size=10)))
            lce.append(self.entrada(ws, cod, r, 8, 9, self.fin(num_format='0.0', font_size=10)))
            r += 1
        tot_l, tot_c = rc(r, 5, True, True), rc(r, 8, True, True)
        for i in range(len(fr)):
            rr = r1 + i
            a, b = rc(rr, 7, True, True), rc(rr, 10, True, True)
            self.cel(ws, rr, 7, 7, '=IF(OR(%s="",N(%s)=0),"",%s/%s*100)' % (ltf[i], tot_l, ltf[i], tot_l),
                     self.fcalc(num_format='0.00', font_size=9.5), formula=True)
            self.cel(ws, rr, 10, 10, '=IF(OR(%s="",N(%s)=0),"",%s/%s*100)' % (lce[i], tot_c, lce[i], tot_c),
                     self.fcalc(num_format='0.00', font_size=9.5), formula=True)
            self.cel(ws, rr, 11, 11, '=IF(OR(%s="",%s=""),"",%s-%s)' % (a, b, a, b),
                     self.fcalc(num_format='0.00', font_size=9.5, font_color=AZUL_ACINZ), formula=True)
            pl.append(a)
            pc.append(b)
        ws.set_row(r, 20)
        self.cel(ws, r, 2, 4, 'Total', self.flbl(font_size=9.5))
        self.cel(ws, r, 5, 6, '=IF(COUNT(%s:%s)=0,"",SUM(%s:%s))' % (ltf[0], ltf[-1], ltf[0], ltf[-1]),
                 self.fcalc(num_format='0.0', bold=True), formula=True)
        self.cel(ws, r, 8, 9, '=IF(COUNT(%s:%s)=0,"",SUM(%s:%s))' % (lce[0], lce[-1], lce[0], lce[-1]),
                 self.fcalc(num_format='0.0', bold=True), formula=True)
        self.cel(ws, r, 7, 7, '', self.fcalc())
        self.cel(ws, r, 10, 11, '', self.fcalc())
        r += 2
        r = self.cabec(ws, r, [(2, 4, 'Indicador'), (5, 6, 'LTF'), (7, 8, 'LCE'), (9, 9, 'Diferença'),
                               (10, 10, 'Tolerância'), (11, self.UC, 'Situação')])

        def soma(lst, idx):
            return '+'.join('N(%s)' % lst[i] for i in idx)

        ind = [('-6,3 mm', '%', [7, 8, 9], None, 'G63', 2), ('RG (Relação)', '-', None, 'rg', 'GRG', 2),
               ('-16,0 +12,5 mm', '%', [2, 3], None, None, 2), ('-16,0 +9,0 mm', '%', [2, 3, 4], None, None, 2),
               ('-16,0 +8,0 mm', '%', [2, 3, 4, 5], None, None, 2), ('Diâmetro médio', 'mm', None, 'dm', None, 2)]
        pesos = [20, 17.5, 15, 13.25, 10.75, 8.5, 7.15]
        refs = []
        for nome, un, idx, esp, ch, dec in ind:
            ws.set_row(r, 23)
            self.cel(ws, r, 2, 4, nome, self.flbl())
            vals = []
            for (c1, c2), p, tot in (((5, 6), pl, tot_l), ((7, 8), pc, tot_c)):
                if esp == 'rg':
                    fm = '=IF(OR(N(%s)=0,N(%s)=0),"",(N(%s)+N(%s))/N(%s))' % (tot, p[4], p[2], p[3], p[4])
                elif esp == 'dm':
                    fm = ('=IF(N(%s)=0,"",(%s+(N(%s)+N(%s)+N(%s))*3.15)/100)'
                          % (tot, '+'.join('N(%s)*%s' % (p[i], w) for i, w in enumerate(pesos)), p[7], p[8], p[9]))
                else:
                    fm = '=IF(N(%s)=0,"",%s)' % (tot, soma(p, idx))
                self.cel(ws, r, c1, c2, fm, self.fcalc(num_format=nf(dec), bold=True), formula=True)
                vals.append(rc(r, c1, True, True))
            d = rc(r, 9, True, True)
            self.cel(ws, r, 9, 9, '=IF(OR(%s="",%s=""),"",%s-%s)' % (vals[0], vals[1], vals[0], vals[1]),
                     self.fcalc(num_format='0.00'), formula=True)
            if ch:
                self.cel(ws, r, 10, 10, '=%s' % crit(ch, 'TOL'), self.fref(num_format='0.00'), formula=True)
                res = self.resultado(ws, cod, r, 11, self.UC, self.f_tol(d, crit(ch, 'TOL')))
            else:
                self.cel(ws, r, 10, 10, '—', self.fref())
                res = self.resultado(ws, cod, r, 11, self.UC, '=IF(%s="","","Informativo")' % vals[0])
            if nome != 'Diâmetro médio':
                refs.append((nome, un, vals, d, ch, res))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for nome, un, vals, d, ch, res in refs:
            self.add_stg(cod, 'Comparativo', '"66PN09 (LTF) x 66PN08 (LCE)"', '"%s"' % nome, un, self.L(vals[0]),
                         self.L(vals[1]), self.L(d), '""', '""', crit(ch, 'TOL') if ch else '""', self.L(res),
                         'TRIM(%s)' % self.L(o), criterio=('tol:' + ch) if ch else 'info')
        return r + 1

    # -------- 7. Tamboramento 5 kg x 15 kg
    def sec_t515(self, ws, r):
        cod = 'T515'
        r = self.titulo_secao(ws, r, cod, 7)
        r = self.cabec(ws, r, [(2, 4, 'Amostra'), (5, 6, '+6,3 mm (g)'), (7, 8, '-6,3 +0,5 mm (g)'),
                               (9, 9, '+6,3 (%)'), (10, 10, '-6,3 +0,5 (%)'), (11, 11, '-0,5 (%)'),
                               (12, self.UC, '')])
        lin = []
        for nome, div in (('LTF 5,0 kg · 66TA05', 50), ('LCE 15,0 kg · T03', 150)):
            ws.set_row(r, 24)
            self.cel(ws, r, 2, 4, nome, self.flbl())
            a = self.entrada(ws, cod, r, 5, 6, self.fin(num_format='0.0'))
            b = self.entrada(ws, cod, r, 7, 8, self.fin(num_format='0.0'))
            pa, pb, pc = rc(r, 9, True, True), rc(r, 10, True, True), rc(r, 11, True, True)
            self.cel(ws, r, 9, 9, '=IF(%s="","",%s/%d)' % (a, a, div), self.fcalc(num_format='0.00', bold=True),
                     formula=True)
            self.cel(ws, r, 10, 10, '=IF(%s="","",%s/%d)' % (b, b, div), self.fcalc(num_format='0.00'), formula=True)
            self.cel(ws, r, 11, 11, '=IF(OR(%s="",%s=""),"",100-(%s+%s))' % (pa, pb, pa, pb),
                     self.fcalc(num_format='0.00'), formula=True)
            self.cel(ws, r, 12, self.UC, 'massa de ensaio %s g' % ('5.000' if div == 50 else '15.000'),
                     self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
            lin.append((pa, pc))
            r += 1
        r = self.cabec(ws, r, [(2, 4, 'Comparativo'), (5, 6, 'LTF'), (7, 8, 'LCE'), (9, 9, 'Diferença'),
                               (10, 10, 'Tolerância'), (11, self.UC, 'Situação')], altura=22)
        refs = []
        for nome, i, ch in (('+6,3 mm', 0, 'T515'), ('-0,5 mm', 1, None)):
            ws.set_row(r, 24)
            self.cel(ws, r, 2, 4, nome, self.flbl())
            x, y = lin[0][i], lin[1][i]
            self.cel(ws, r, 5, 6, '=IF(%s="","",%s)' % (x, x), self.fcalc(num_format='0.00', bold=True), formula=True)
            self.cel(ws, r, 7, 8, '=IF(%s="","",%s)' % (y, y), self.fcalc(num_format='0.00', bold=True), formula=True)
            d = rc(r, 9, True, True)
            self.cel(ws, r, 9, 9, '=IF(OR(%s="",%s=""),"",%s-%s)' % (x, y, x, y), self.fcalc(num_format='0.00'),
                     formula=True)
            if ch:
                self.cel(ws, r, 10, 10, '=%s' % crit(ch, 'TOL'), self.fref(num_format='0.00'), formula=True)
                res = self.resultado(ws, cod, r, 11, self.UC, self.f_tol(d, crit(ch, 'TOL')))
            else:
                self.cel(ws, r, 10, 10, '—', self.fref())
                res = self.resultado(ws, cod, r, 11, self.UC, '=IF(OR(%s="",%s=""),"","Informativo")' % (x, y))
            refs.append((nome, x, y, d, ch, res))
            r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for nome, x, y, d, ch, res in refs:
            self.add_stg(cod, 'Comparativo', '"66TA05 (5 kg) x T03 (15 kg)"', '"%s"' % nome, '%',
                         'IF(%s="","",%s)' % (self.L(x), self.L(x)), 'IF(%s="","",%s)' % (self.L(y), self.L(y)),
                         self.L(d), '""', '""', crit(ch, 'TOL') if ch else '""', self.L(res), 'TRIM(%s)' % self.L(o),
                         criterio=('tol:' + ch) if ch else 'info')
        return r + 1

    # -------- 8. Peneiradores
    def sec_peneiradores(self, ws, r):
        cod = 'PEN'
        r = self.titulo_secao(ws, r, cod, 8)
        r = self.cabec(ws, r, [(2, 3, 'Peneirador'), (4, 4, 'Malha (mm)'), (5, 6, 'Tag'), (7, 7, 'Condições\ndas telas'),
                               (8, 8, 'Inclinação\ndas telas'), (9, 9, 'Vibração'), (10, 10, 'Travamento'),
                               (11, 11, 'Calibração'), (12, 12, 'Itens OK'), (13, self.UC, 'Situação')], altura=34)
        grupo = None
        r_ini = r
        refs = []
        for g, malha, tag in PENEIRAS:
            ws.set_row(r, 21)
            primeira = g != grupo
            grupo = g
            self.cel(ws, r, 2, 3, g if primeira else '', self.flbl(font_size=9.5) if primeira else
                     self.f(bg_color=FUNDO_CLARO, border=1, border_color=BORDA, top=0 if not primeira else 1))
            self.cel(ws, r, 4, 4, malha, self.fref(font_size=10, bold=True, font_color=AZUL_TITULO))
            tg = self.entrada_fixa(ws, r, 5, 6, tag)
            itens = []
            for k in range(5):
                c = 7 + k
                itens.append(self.entrada(ws, cod, r, c, c, self.fin(font_size=9.5)))
            rng = '%s:%s' % (itens[0], itens[-1])
            ok = rc(r, 12, True, True)
            self.cel(ws, r, 12, 12, '=IF(COUNTA(%s)=0,"",COUNTIF(%s,"OK")&" / "&COUNTA(%s))' % (rng, rng, rng),
                     self.fcalc(font_size=9.5), formula=True)
            res = self.resultado(ws, cod, r, 13, self.UC,
                                 '=IF(COUNTA(%s)=0,"",IF(COUNTIF(%s,"NÃO OK")>0,"Não conforme",IF(COUNTIF(%s,"SEM TAG")>0,'
                                 '"Conforme · sem TAG","Conforme")))' % (rng, rng, rng))
            refs.append((g, malha, tg, itens, rng, res))
            r += 1
        self.pen_area = (r_ini, 7, r - 1, 11)
        self.nome('penItens', ws, r_ini, 7, r - 1, 11)
        ws.data_validation(r_ini, 7, r - 1, 11, {'validate': 'list', 'source': '=lstPen'})
        ws.set_row(r, 26)
        ws.merge_range(r, 2, r, self.UC, '', self.f())
        self.botao(ws, r, 2, 'Marcar itens vazios como OK', 'MarcarPeneiradoresOK', 230, 26, 'claro', x=2, y=0)
        r += 1
        o = self.obs(ws, cod, r)
        r += 1
        for g, malha, tg, itens, rng, res in refs:
            R = self.L(rng.split(':')[0]) + ':' + rng.split(':')[1]
            nomes = ''.join('&IF(%s="NÃO OK","%s; ","")' % (self.L(it), ITENS_PEN[k]) for k, it in enumerate(itens))
            obs = ('IF(COUNTIF(%s,"SEM TAG")>0,"Sem TAG de calibração; ","")%s&TRIM(%s)' % (R, nomes, self.L(o)))
            self.add_stg(cod, 'Verificação', 'IF(TRIM(%s)="","%s",TRIM(%s))' % (self.L(tg), g, self.L(tg)),
                         '"Malha %s mm (%s)"' % (malha, g), 'itens OK',
                         'IF(COUNTA(%s)=0,"",COUNTIF(%s,"OK"))' % (R, R), 'IF(COUNTA(%s)=0,"",COUNTA(%s))' % (R, R),
                         '""', '""', '""', '""',
                         'IF(COUNTA(%s)=0,"",IF(COUNTIF(%s,"NÃO OK")>0,"Não conforme","Conforme"))' % (R, R), obs)
        return r + 1

    # -------- 9. Fisher
    def sec_fisher(self, ws, r):
        cod = 'FIS'
        r = self.titulo_secao(ws, r, cod, 9)
        r = self.cabec(ws, r, [(2, 4, 'Análise'), (5, 6, '66PB01 (cm²/g)'), (7, 8, '66PB09 (cm²/g)'),
                               (9, 9, 'Diferença'), (10, 10, 'Tolerância'), (11, self.UC, 'Situação')])
        ws.set_row(r, 24)
        self.cel(ws, r, 2, 4, 'Superfície específica', self.flbl())
        a = self.entrada(ws, cod, r, 5, 6, self.fin(num_format='0'))
        b = self.entrada(ws, cod, r, 7, 8, self.fin(num_format='0'))
        d = rc(r, 9, True, True)
        self.cel(ws, r, 9, 9, '=IF(OR(%s="",%s=""),"",%s-%s)' % (a, b, a, b), self.fcalc(num_format='0'), formula=True)
        self.cel(ws, r, 10, 10, '=%s' % crit('FIS', 'TOL'), self.fref(num_format='0'), formula=True)
        res = self.resultado(ws, cod, r, 11, self.UC, self.f_tol(d, crit('FIS', 'TOL')))
        r += 1
        o = self.obs(ws, cod, r)
        r += 1
        self.add_stg(cod, 'Comparativo', '"66PB01 x 66PB09"', '"Superfície específica"', 'cm²/g', self.L(a),
                     'IF(%s="","",%s)' % (self.L(b), self.L(b)), self.L(d), '""', '""', crit('FIS', 'TOL'),
                     self.L(res), 'TRIM(%s)' % self.L(o), dec=0, criterio='tol:FIS')
        return r + 1

    # -------- agenda do turno (topo do Lancamento)
    def preencher_agenda(self, ws):
        r = self.r_agenda1
        for i, (cod, nome, tit, norma, freq) in enumerate(ENSAIOS):
            rr = r + i
            ws.set_row(rr, 21)
            zebra = BRANCO if i % 2 == 0 else '#F8F9FA'
            b = dict(border=1, border_color=BORDA, bg_color=zebra)
            ws.merge_range(rr, 2, rr, 5, '', self.f(**b))
            self.link(ws, rr, 2, "'%s'!C%d" % (LAN, self.secoes[cod] + 1), '%d. %s' % (i + 1, tit),
                      fmt=self.f(font_size=10, bold=True, font_color=AZUL_MEDIO, underline=1, indent=1, **b))
            self.cel(ws, rr, 6, 7, freq, self.f(font_size=8.5, font_color=TEXTO_SEC, align='center', **b))
            prev = ('=IF(OR(pData="",pTurno=""),"—",IF(INDEX(agenda,%d,(WEEKDAY(pData,2)-1)*2+IF(pTurno="07x19",1,2))'
                    '="X","Previsto","—"))' % (i + 1))
            self.cel(ws, rr, 8, 9, prev, self.f(font_size=9.5, bold=True, align='center', font_color=TEXTO_SEC, **b),
                     formula=True)
            self.cel(ws, rr, 10, 11, '=IF(cnt_%s>0,"Preenchido",IF(%s="Previsto","Pendente","—"))'
                     % (cod, rc(rr, 8)), self.f(font_size=9.5, bold=True, align='center', font_color=TEXTO_SEC, **b),
                     formula=True)
            res = self.resultados[cod]
            nc = '+'.join('COUNTIF(%s,"Não conforme")' % x for x in res)
            cf = '+'.join('COUNTIF(%s,"Conforme*")' % x for x in res)
            self.cel(ws, rr, 12, 13, '=IF((%s)>0,"Não conforme ("&(%s)&")",IF((%s)>0,"Conforme","—"))' % (nc, nc, cf),
                     self.f(font_size=9.5, bold=True, align='center', font_color=TEXTO_SEC, **b), formula=True)
            ult = ('{=IFERROR(1/(1/MAX(IF(BD_Afericoes!$A$2:$A$%d="%s",BD_Afericoes!$C$2:$C$%d))),"")}'
                   % (BD_MAX, nome, BD_MAX))
            fu = self.f(font_size=9.5, align='center', num_format='dd/mm/yyyy', **b)
            ws.merge_range(rr, 14, rr, self.UC, '', fu)
            ws.write_array_formula(rr, 14, rr, 14, ult, fu, '')
        self.cf_status(ws, r, 8, r + len(ENSAIOS) - 1, 13)

    # ------------------------------------------------------------ Config: agenda, contagens, mapa de entradas
    def aba_config_resto(self):
        ws = self.ws_cfg
        hdr = self.fhdr()
        inp = self.f(bg_color=FUNDO_INPUT, border=1, border_color=BORDA, locked=False, align='center', bold=True)
        r0 = 5
        c0 = 7
        ws.merge_range(r0 - 1, c0, r0 - 1, c0 + 14, 'Agenda semanal (X = previsto). Base: aba "Controle" da planilha '
                       'anterior. D = turno 07x19, N = turno 19x07.', self.flbl())
        ws.write(r0, c0, 'Ensaio', hdr)
        for d in range(7):
            for t in range(2):
                ws.write(r0, c0 + 1 + d * 2 + t, '%s %s' % (DIAS[d], 'D' if t == 0 else 'N'), hdr)
        for i, (cod, nome, tit, _, _) in enumerate(ENSAIOS):
            ws.write(r0 + 1 + i, c0, tit, self.f(bold=True, font_size=9))
            for d in range(7):
                for t in range(2):
                    marca = 'X' if (d, _D if t == 0 else _N) in AGENDA[cod] else ''
                    ws.write(r0 + 1 + i, c0 + 1 + d * 2 + t, marca, inp)
        ws.set_column(c0, c0, 30)
        self.nome('agenda', ws, r0 + 1, c0 + 1, r0 + len(ENSAIOS), c0 + 14)
        self.nome('agendaNomes', ws, r0 + 1, c0, r0 + len(ENSAIOS), c0)
        # parametros gerais
        r = 17
        ws.merge_range(r, c0, r, c0 + 6, 'Parâmetros gerais', self.flbl())
        gerais = [('cfgBackup', 'Pasta do backup automático (ao salvar)', BACKUP_PADRAO),
                  ('cfgUltimoID', 'Último número de lançamento', 0),
                  ('cfgUltimoLanc', 'Último lançamento', ''),
                  ('cfgPontosGrafico', 'Pontos nos gráficos do Painel', 30)]
        for i, (nm, rot, val) in enumerate(gerais):
            rr = r + 1 + i
            ws.write(rr, c0, rot, self.f(bold=True, font_size=9))
            ws.merge_range(rr, c0 + 1, rr, c0 + 10, '', inp)
            ws.write(rr, c0 + 1, val, self.f(bg_color=FUNDO_INPUT, border=1, border_color=BORDA, locked=False,
                                             align='left'))
            self.nome(nm, ws, rr, c0 + 1)
        # contagem de entradas por ensaio (Lancamento -> "Nesta tela")
        r = 24
        ws.merge_range(r, c0, r, c0 + 6, 'Campos preenchidos no Lançamento (uso interno)', self.flbl())
        for i, cod in enumerate(COD):
            rr = r + 1 + i
            ws.write(rr, c0, cod, self.f(bold=True))
            partes = ','.join(QL + a for a in self.inputs[cod] if not self.eh_obs(cod, a))
            ws.write_formula(rr, c0 + 1, '=COUNTA(%s)' % partes, self.f(), 0)
            self.nome('cnt_' + cod, ws, rr, c0 + 1)
            # enderecos a limpar (texto lido pelo VBA)
            ws.write_string(rr, c0 + 2, ';'.join(self.inputs[cod]))
        self.nome('mapaInputs', ws, r + 1, c0, r + len(COD), c0 + 2)

    def eh_obs(self, cod, a):
        # observacao: ultima entrada da secao, larga (coluna E ate P)
        return a.startswith('$E$') and ':$P$' in a

    # ------------------------------------------------------------ _Staging
    def aba_staging(self):
        ws = self.ws_stg
        cab = BD_COLS + ['Registrar', 'Ensaio (cód.)']
        for i, h in enumerate(cab):
            ws.write(0, i, h)
        ws.set_column(0, 18, 16)
        for i, s in enumerate(self.stg):
            r = 1 + i
            campos = [s['ensaio'], s['tipo'], 'IF(pData="","",pData)', s['equip'], s['param'], s['unid'], s['valor'],
                      s['refv'], s['dif'], s['li'], s['ls'], s['tol'], s['res'], 'TRIM(pResp)', 'pLetra', s['obs'],
                      '""']
            for c, fml in enumerate(campos):
                if c in (6, 7, 8, 9, 10, 11) and not fml.startswith('IF') and fml != '""':
                    fml = 'IF(%s="","",%s)' % (fml, fml)
                ws.write_formula(r, c, '=' + fml, None, '')
            ws.write_formula(r, 17, '=IF(AND(G%d<>"",M%d<>""),1,0)' % (r + 1, r + 1), None, 0)
            ws.write_string(r, 18, s['cod'])
        self.n_stg = len(self.stg)

    # ------------------------------------------------------------ BD_Afericoes
    def aba_bd(self):
        ws = self.ws_bd
        larg = [22, 12, 11, 26, 30, 9, 10, 11, 10, 11, 11, 10, 14, 14, 7, 40, 46]
        for i, w in enumerate(larg):
            ws.set_column(i, i, w)
        fdata = self.wb.add_format({'num_format': 'dd/mm/yyyy'})
        dados = []
        for v in DADOS['afericoes']:
            v = list(v)
            v[2] = dtm.datetime.strptime(v[2], '%Y-%m-%d')
            dados.append(v)
        ws.add_table(0, 0, len(dados), 16, {
            'name': 'tbl_Afericoes', 'style': 'Table Style Medium 2',
            'columns': [{'header': h, 'format': fdata} if h == 'Data' else {'header': h} for h in BD_COLS],
            'data': dados})
        ws.set_column(2, 2, 11, fdata)
        ws.freeze_panes(1, 0)

    # ------------------------------------------------------------ _Graficos (dados preenchidos pelo VBA)
    GRAFICOS = [
        # (codigo, titulo, campo (G valor | I diferenca), unidade, criterio para linhas de referencia, dec)
        ('BLA', 'Blaine · resultado (cm²/g)', 'G', 'cm²/g', 'faixa:BLAM', 0),
        ('TAM', 'Tambores · rotação (rpm)', 'G', 'rpm', 'faixa:TAMB', 2),
        ('ALP', 'Alpine · passante (%)', 'G', '%', 'faixa:ALP', 2),
        ('UMI', 'Umidade · analisador - estufa (p.p.)', 'I', 'p.p.', 'tol:UMI', 2),
        ('COM', 'Compressão FX -16,0 +12,5 · diferença ao nominal (kgf/pel)', 'I', 'kgf/pel', 'tol:COM16', 0),
    ]
    GRA_LIN = 40          # linhas reservadas por bloco de grafico
    GRA_NPT = 30

    def aba_graficos(self):
        ws = self.ws_gra
        ws.write(0, 0, 'Dados dos gráficos do Painel (preenchidos pela macro AtualizarPainel)')
        self.gra_blocos = []
        for k, (cod, tit, campo, un, criterio, dec) in enumerate(self.GRAFICOS):
            r0 = 2 + k * self.GRA_LIN
            series = self.series[cod]
            ws.write(r0, 0, cod)
            ws.write(r0, 1, campo)
            ws.write(r0, 2, criterio)
            ws.write(r0 + 1, 0, 'Ocasião')
            for j, (eq, pa) in enumerate(series):
                ws.write(r0 + 1, 1 + j, eq)
                ws.write(r0, 3 + j, pa)
            nser = len(series)
            ws.write(r0 + 1, 1 + nser, 'Lim. inferior')
            ws.write(r0 + 1, 2 + nser, 'Lim. superior')
            self.gra_blocos.append((cod, r0, nser))

    # ------------------------------------------------------------ Painel
    def aba_painel(self):
        ws = self.ws_pai
        UC = 16
        self.P = P = {}
        ws.set_column(0, 0, 1.5)
        ws.set_column(1, UC, 10.5)
        ws.set_column(UC + 1, UC + 1, 2.5)
        ws.set_column(UC + 2, UC + 2, 27)
        self.cabecalho(ws, UC, 'CONTROLE DE CALIBRAÇÃO E AFERIÇÃO', 'LABORATÓRIO FÍSICO  |  PAINEL DE RESULTADOS')
        lbl = self.flbl()
        ws.set_row(4, 8)
        r = 5
        ws.set_row(r, 26)
        self.cel(ws, r, 1, 2, 'Período', lbl)
        self.cel(ws, r, 3, 4, '', self.fin(num_format='dd/mm/yyyy', font_size=12))
        self.nome('painelIni', ws, r, 3)
        self.cel(ws, r, 5, 5, 'até', self.f(align='center', font_color=TEXTO_SEC))
        self.cel(ws, r, 6, 7, '', self.fin(num_format='dd/mm/yyyy', font_size=12))
        self.nome('painelFim', ws, r, 6)
        for c in (3, 6):
            ws.data_validation(r, c, r, c, {'validate': 'date', 'criteria': '>', 'value': dtm.date(2020, 1, 1)})
        ws.merge_range(r, 8, r, UC, '', self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
        ws.write(r, 8, 'Painel não atualizado', self.f(font_size=8.5, italic=True, font_color=TEXTO_SEC, indent=1))
        self.nome('painelInfo', ws, r, 8)
        ws.freeze_panes(r + 1, 0)
        self.botao(ws, 1, UC + 2, 'Atualizar painel', 'AtualizarPainel', 190, 34, 'destaque', x=4, y=4)
        self.botao(ws, 3, UC + 2, 'Mês atual', 'PainelMesAtual', 92, 26, 'claro', x=4, y=6)
        self.botao(ws, 3, UC + 2, 'Últimos 30 dias', 'PainelUltimos30', 94, 26, 'claro', x=100, y=6)
        self.botao(ws, 5, UC + 2, 'Ir ao lançamento', 'IrLancamento', 190, 28, 'primario', x=4, y=0)
        self.botao(ws, 7, UC + 2, 'Importar planilha antiga', 'ImportarArquivoAntigo', 190, 28, 'claro', x=4, y=4)

        # ---- KPIs
        r += 2
        kpis = [('kpiReg', 'Resultados registrados'), ('kpiConf', 'Conformidade'), ('kpiNC', 'Não conformidades'),
                ('kpiAder', 'Aderência à agenda')]
        ws.set_row(r, 18)
        ws.set_row(r + 1, 40)
        ws.set_row(r + 2, 18)
        for i, (nm, t) in enumerate(kpis):
            c1 = 1 + i * 4
            c2 = c1 + 3
            cor = [AZUL, VERDE_TXT, VERM_TXT, AZUL_MEDIO][i]
            ws.merge_range(r, c1, r, c2, t, self.f(bold=True, font_size=9, font_color=TEXTO_SEC, bg_color=FUNDO_CLARO,
                                                   align='center', top=1, left=1, right=1, border_color=BORDA))
            ws.merge_range(r + 1, c1, r + 1, c2, '—', self.f(bold=True, font_size=24, font_color=cor,
                                                             bg_color=FUNDO_CLARO, align='center', left=1, right=1,
                                                             border_color=BORDA))
            ws.merge_range(r + 2, c1, r + 2, c2, '', self.f(font_size=8.5, font_color=TEXTO_SEC, bg_color=FUNDO_CLARO,
                                                            align='center', bottom=1, left=1, right=1,
                                                            border_color=BORDA))
            self.nome(nm, ws, r + 1, c1)
            self.nome(nm + 'Sub', ws, r + 2, c1)
        r += 4

        # ---- por ensaio
        self.faixa(ws, r, 1, UC, 'RESULTADOS POR ENSAIO')
        r += 1
        cols = [(1, 4, 'Ensaio'), (5, 6, 'Previstos na agenda'), (7, 7, 'Realizados'), (8, 8, 'Aderência'),
                (9, 9, 'Resultados'), (10, 10, 'Conformes'), (11, 11, 'Não conformes'), (12, 12, 'Conformidade'),
                (13, 14, 'Último registro'), (15, UC, 'Situação')]
        ws.set_row(r, 30)
        for c1, c2, t in cols:
            self.cel(ws, r, c1, c2, t, self.fhdr())
        r += 1
        P['ens1'] = r
        for i, (cod, nome, tit, _, _) in enumerate(ENSAIOS):
            rr = r + i
            ws.set_row(rr, 21)
            zebra = BRANCO if i % 2 == 0 else '#F8F9FA'
            b = dict(border=1, border_color=BORDA, bg_color=zebra)
            self.cel(ws, rr, 1, 4, tit, self.f(bold=True, indent=1, **b))
            for c1, c2, fm in ((5, 6, '0'), (7, 7, '0'), (8, 8, '0%'), (9, 9, '0'), (10, 10, '0'), (11, 11, '0'),
                               (12, 12, '0.0%'), (13, 14, 'dd/mm/yyyy'), (15, UC, '@')):
                self.cel(ws, rr, c1, c2, '', self.f(align='center', num_format=fm, **b))
        self.cf_status(ws, r, 15, r + len(ENSAIOS) - 1, UC)
        ws.conditional_format(r, 11, r + len(ENSAIOS) - 1, 11, {
            'type': 'cell', 'criteria': '>', 'value': 0,
            'format': self.wb.add_format({'font_color': VERM_TXT, 'bold': True})})
        ws.conditional_format(r, 8, r + len(ENSAIOS) - 1, 8, {
            'type': 'data_bar', 'bar_color': '#8FC7E8', 'bar_solid': True, 'min_type': 'num', 'min_value': 0,
            'max_type': 'num', 'max_value': 1})
        r += len(ENSAIOS) + 1

        # ---- graficos
        self.faixa(ws, r, 1, UC, 'TENDÊNCIAS (ÚLTIMOS LANÇAMENTOS DO PERÍODO)')
        r += 1
        P['graf1'] = r
        self.graficos_painel(ws, r)
        r += 3 * 17 + 1

        # ---- situacao por equipamento
        self.faixa(ws, r, 1, UC, 'SITUAÇÃO ATUAL POR EQUIPAMENTO (ÚLTIMO RESULTADO NO PERÍODO)')
        r += 1
        cols = [(1, 2, 'Ensaio'), (3, 4, 'Equipamento'), (5, 7, 'Parâmetro'), (8, 8, 'Resultado'),
                (9, 9, 'Referência'), (10, 10, 'Diferença'), (11, 12, 'Critério'), (13, 13, 'Data'),
                (14, 14, 'Responsável'), (15, UC, 'Situação')]
        ws.set_row(r, 26)
        for c1, c2, t in cols:
            self.cel(ws, r, c1, c2, t, self.fhdr())
        r += 1
        P['eq1'] = r
        for i, (ens, eq, pa, un, criterio, dec) in enumerate(self.combos):
            rr = r + i
            ws.set_row(rr, 18)
            zebra = BRANCO if i % 2 == 0 else '#F8F9FA'
            b = dict(border=1, border_color=BORDA, bg_color=zebra, font_size=9)
            self.cel(ws, rr, 1, 2, ens, self.f(font_color=TEXTO_SEC, indent=1, **b))
            self.cel(ws, rr, 3, 4, eq, self.f(bold=True, indent=1, **b))
            self.cel(ws, rr, 5, 7, pa, self.f(indent=1, **b))
            for c1, c2, fm in ((8, 8, nf(dec)), (9, 9, nf(dec)), (10, 10, nf(dec))):
                self.cel(ws, rr, c1, c2, '', self.f(align='center', num_format=fm, bold=(c1 == 8), **b))
            self.cel(ws, rr, 11, 12, '', self.f(align='center', font_color=AZUL_ACINZ, **b))
            self.cel(ws, rr, 13, 13, '', self.f(align='center', num_format='dd/mm/yy', **b))
            self.cel(ws, rr, 14, 14, '', self.f(align='center', **b))
            self.cel(ws, rr, 15, UC, '', self.f(align='center', bold=True, **b))
        P['neq'] = len(self.combos)
        self.cf_status(ws, r, 15, r + len(self.combos) - 1, UC)
        r += len(self.combos) + 1

        # ---- peneiradores
        self.faixa(ws, r, 1, UC, 'PENEIRADORES (ÚLTIMA VERIFICAÇÃO NO PERÍODO)')
        r += 1
        cols = [(1, 4, 'Peneirador'), (5, 6, 'Data'), (7, 8, 'Telas verificadas'), (9, 10, 'Telas com NÃO OK'),
                (11, 12, 'Telas sem TAG'), (13, 14, 'Responsável'), (15, UC, 'Situação')]
        ws.set_row(r, 26)
        for c1, c2, t in cols:
            self.cel(ws, r, c1, c2, t, self.fhdr())
        r += 1
        P['pen1'] = r
        grupos = []
        for g, _, _ in PENEIRAS:
            if g not in grupos:
                grupos.append(g)
        for i, g in enumerate(grupos):
            rr = r + i
            ws.set_row(rr, 20)
            b = dict(border=1, border_color=BORDA, bg_color=BRANCO if i % 2 == 0 else '#F8F9FA')
            self.cel(ws, rr, 1, 4, g, self.f(bold=True, indent=1, **b))
            for c1, c2, fm in ((5, 6, 'dd/mm/yyyy'), (7, 8, '0'), (9, 10, '0'), (11, 12, '0'), (13, 14, '@'),
                               (15, UC, '@')):
                self.cel(ws, rr, c1, c2, '', self.f(align='center', num_format=fm, **b))
        self.grupos_pen = grupos
        self.cf_status(ws, r, 15, r + len(grupos) - 1, UC)
        r += len(grupos) + 1

        # ---- nao conformidades
        self.faixa(ws, r, 1, UC, 'NÃO CONFORMIDADES DO PERÍODO (MAIS RECENTES)', cor=VERM_TXT)
        r += 1
        cols = [(1, 1, 'Data'), (2, 2, 'Turno'), (3, 4, 'Ensaio'), (5, 6, 'Equipamento'), (7, 9, 'Parâmetro'),
                (10, 10, 'Resultado'), (11, 11, 'Referência'), (12, 12, 'Critério'), (13, 13, 'Responsável'),
                (14, UC, 'Observação')]
        ws.set_row(r, 26)
        for c1, c2, t in cols:
            self.cel(ws, r, c1, c2, t, self.fhdr(bg_color=VERM_TXT))
        r += 1
        P['nc1'] = r
        P['nnc'] = 25
        for i in range(P['nnc']):
            rr = r + i
            ws.set_row(rr, 18)
            b = dict(border=1, border_color=BORDA, bg_color=BRANCO if i % 2 == 0 else '#FDF6F5', font_size=9)
            for c1, c2, fm in ((1, 1, 'dd/mm/yy'), (2, 2, '@'), (3, 4, '@'), (5, 6, '@'), (7, 9, '@'),
                               (10, 10, '0.00'), (11, 11, '0.00'), (12, 12, '@'), (13, 13, '@'), (14, UC, '@')):
                self.cel(ws, rr, c1, c2, '', self.f(align='center' if c1 < 14 else 'left', num_format=fm, **b))
        r += P['nnc']
        P['fim'] = r
        ws.print_area(1, 1, r, UC)
        ws.set_portrait()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        ws.protect('', {'format_columns': True, 'format_rows': True})

    def graficos_painel(self, ws, r):
        cores = ['#00335A', '#F37021', '#59C6F2', '#1E7B4A', '#9C27B0', '#B5BEC4']
        # 1) conformidade por ensaio (barras empilhadas, dados da tabela por ensaio)
        e1, e2 = self.P['ens1'], self.P['ens1'] + len(ENSAIOS) - 1
        ch = self.wb.add_chart({'type': 'bar', 'subtype': 'stacked'})
        ch.add_series({'name': 'Conformes', 'categories': ['Painel', e1, 1, e2, 1],
                       'values': ['Painel', e1, 10, e2, 10], 'fill': {'color': '#3BA272'}, 'gap': 60})
        ch.add_series({'name': 'Não conformes', 'categories': ['Painel', e1, 1, e2, 1],
                       'values': ['Painel', e1, 11, e2, 11], 'fill': {'color': '#D9534F'}})
        self.estilo_grafico(ch, 'Resultados por ensaio no período', None)
        ch.set_y_axis({'reverse': True, 'num_font': {'size': 8, 'name': FONTE}})
        ch.set_x_axis({'num_font': {'size': 8, 'name': FONTE}, 'major_gridlines': {'visible': True,
                                                                                    'line': {'color': '#E5E9EC'}}})
        ws.insert_chart(r, 1, ch, {'x_offset': 4, 'y_offset': 6, 'object_position': 1,
                                   'width': 560, 'height': 300})
        # 2..6) tendencias
        pos = [(r, 9), (r + 17, 1), (r + 17, 9), (r + 34, 1), (r + 34, 9)]
        for k, (cod, tit, campo, un, criterio, dec) in enumerate(self.GRAFICOS):
            _, r0, nser = self.gra_blocos[k]
            c = self.wb.add_chart({'type': 'line'})
            a, z = r0 + 2, r0 + 1 + self.GRA_NPT
            for j in range(nser):
                c.add_series({'name': ['_Graficos', r0 + 1, 1 + j], 'categories': ['_Graficos', a, 0, z, 0],
                              'values': ['_Graficos', a, 1 + j, z, 1 + j],
                              'line': {'color': cores[j % len(cores)], 'width': 1.75},
                              'marker': {'type': 'circle', 'size': 4,
                                         'fill': {'color': cores[j % len(cores)]},
                                         'border': {'color': cores[j % len(cores)]}}})
            for j, nomeLim in enumerate(('Lim. inferior', 'Lim. superior')):
                c.add_series({'name': ['_Graficos', r0 + 1, 1 + nser + j], 'categories': ['_Graficos', a, 0, z, 0],
                              'values': ['_Graficos', a, 1 + nser + j, z, 1 + nser + j],
                              'line': {'color': '#D9534F', 'width': 1.25, 'dash_type': 'dash'},
                              'marker': {'type': 'none'}})
            c.show_blanks_as('gap')
            self.estilo_grafico(c, tit, un)
            rr, cc = pos[k]
            ws.insert_chart(rr, cc, c, {'x_offset': 4, 'y_offset': 6, 'object_position': 1,
                                        'width': 560, 'height': 300})

    def estilo_grafico(self, ch, titulo, un):
        ch.set_title({'name': titulo, 'name_font': {'size': 10.5, 'bold': True, 'color': AZUL_TITULO, 'name': FONTE}})
        ch.set_legend({'position': 'bottom', 'font': {'size': 8, 'name': FONTE}})
        ch.set_chartarea({'border': {'color': BORDA}, 'fill': {'color': BRANCO}})
        ch.set_plotarea({'fill': {'color': BRANCO}})
        if un:
            ch.set_x_axis({'num_font': {'size': 7.5, 'name': FONTE, 'rotation': -45}})
            ch.set_y_axis({'num_font': {'size': 8, 'name': FONTE},
                           'major_gridlines': {'visible': True, 'line': {'color': '#E5E9EC'}}})

    # ------------------------------------------------------------ VBA: posicoes
    def modulo_layout(self):
        P = self.P
        linhas = ["Option Explicit", "",
                  "' Gerado por build_calibracao.py - posicoes fixas (nao editar a mao)",
                  "Public Const NSTG As Long = %d" % self.n_stg,
                  "Public Const N_ENSAIOS As Long = %d" % len(ENSAIOS),
                  "Public Const PAI_ENS1 As Long = %d" % (P['ens1'] + 1),
                  "Public Const PAI_EQ1 As Long = %d" % (P['eq1'] + 1),
                  "Public Const PAI_NEQ As Long = %d" % P['neq'],
                  "Public Const PAI_PEN1 As Long = %d" % (P['pen1'] + 1),
                  "Public Const PAI_NPEN As Long = %d" % len(self.grupos_pen),
                  "Public Const PAI_NC1 As Long = %d" % (P['nc1'] + 1),
                  "Public Const PAI_NNC As Long = %d" % P['nnc'],
                  "Public Const GRA_NBLOCOS As Long = %d" % len(self.GRAFICOS),
                  "Public Const GRA_LIN As Long = %d" % self.GRA_LIN,
                  "Public Const GRA_NPT As Long = %d" % self.GRA_NPT,
                  "' Coluna da 1a celula de cada tabela do Painel",
                  "Public Const PAI_COL1 As Long = 2",
                  'Public Const SENHA As String = ""', ""]
        # nomes dos ensaios (ordem da agenda e da tabela por ensaio)
        linhas.append("Public Function NomeEnsaio(ByVal i As Long) As String")
        linhas.append("    Select Case i")
        for i, e in enumerate(ENSAIOS):
            linhas.append('        Case %d: NomeEnsaio = "%s"' % (i + 1, e[1]))
        linhas.append("    End Select")
        linhas.append("End Function")
        linhas.append("")
        linhas.append("' Criterio de cada linha da tabela 'Situacao por equipamento' (faixa:CHAVE, tol:CHAVE ou info)")
        linhas.append("Public Function CriterioEquip(ByVal i As Long) As String")
        linhas.append("    Select Case i")
        for i, cb in enumerate(self.combos):
            linhas.append('        Case %d: CriterioEquip = "%s"' % (i + 1, cb[4]))
        linhas.append("    End Select")
        linhas.append("End Function")
        return '\n'.join(linhas)


# ============================================================================
def checar_declaracoes(nome, codigo):
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
        'shPainel': ('Option Explicit\n\nPrivate Sub Worksheet_Activate()\n    On Error Resume Next\n'
                     '    AtualizarPainel True\nEnd Sub\n'),
        'shLancamento': ('Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n'
                         '    On Error Resume Next\n    AoAlterarLancamento Target\nEnd Sub\n'),
    }
    for cn in ['shLancamento', 'shPainel', 'shBD', 'shLimites', 'shConfig', 'shStaging', 'shGraficos']:
        mods.append({'name': cn, 'kind': 'sheet', 'code': eventos.get(cn, '')})
    for m in ['ModLayout', 'ModGeral', 'ModRegistro', 'ModPainel']:
        mods.append({'name': m, 'kind': 'module', 'code': ler_vba(m + '.bas')})
    if TESTE:
        mods.append({'name': 'ModTesteLO', 'kind': 'module', 'code': ler_vba('../build/ModTesteLO.bas')})
    return build_vba_project(mods)


def pos_processar(caminho):
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
                          '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/>'
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
    print('Gerado:', SAIDA, '| linhas candidatas:', c.n_stg, '| equipamentos no painel:', len(c.combos))


if __name__ == '__main__':
    main()
