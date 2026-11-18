"""Le o historico de uma planilha do comparativo ja em uso (tbl_Batches, tbl_Leituras e o registro _Acomp) e grava
um .json no formato de historico.json, para gerar a planilha nova ja com ele:

    python3 extrair_historico.py "Comparativo_Interlab_Germano_Ubu_2.xlsm" hist.json
    python3 build_interlab.py --historico hist.json      # -> ../Comparativo_Interlab_Germano_Ubu_COM_HISTORICO.xlsm

As colunas sao casadas pelo nome do cabecalho: coluna que nao existe mais na planilha nova fica de fora; coluna nova
fica vazia.
"""
import datetime as dtm
import json
import os
import sys

import openpyxl


def valor(v):
    if isinstance(v, dtm.datetime):
        return {'dt': v.isoformat()}
    if isinstance(v, dtm.date):
        return {'dt': dtm.datetime(v.year, v.month, v.day).isoformat()}
    if isinstance(v, str) and v.startswith('='):
        return None
    return v


def tabela(wb, nome):
    for ws in wb.worksheets:
        if nome in ws.tables:
            ref = ws.tables[nome].ref
            linhas = list(ws[ref])
            cab = [c.value for c in linhas[0]]
            dados = [[valor(c.value) for c in lin] for lin in linhas[1:]]
            dados = [d for d in dados if any(x is not None for x in d)]
            return cab, dados
    return [], []


def main(origem, destino):
    wb = openpyxl.load_workbook(origem)
    cab_b, dados_b = tabela(wb, 'tbl_Batches')
    cab_l, dados_l = tabela(wb, 'tbl_Leituras')
    acomp_cab, acomp = [], []
    if '_Acomp' in wb.sheetnames:
        ws = wb['_Acomp']
        linhas = [[valor(c.value) for c in lin] for lin in ws.iter_rows()]
        if linhas:
            acomp_cab = linhas[0]
            acomp = [lin for lin in linhas[1:] if lin and lin[0] is not None]
    saida = {
        'origem': '%s (extraído em %s)' % (os.path.basename(origem), dtm.date.today().strftime('%d/%m/%Y')),
        'colunas_batches': cab_b,
        'batches': [dict(zip(cab_b, d)) for d in dados_b],
        'colunas_leituras': cab_l,
        'leituras': dados_l,
        'colunas_acomp': acomp_cab,
        'acomp': acomp,
    }
    with open(destino, 'w', encoding='utf-8') as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=1)
    print('%d batches, %d leituras, %d no registro de acompanhamento -> %s'
          % (len(dados_b), len(dados_l), len(acomp), destino))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
