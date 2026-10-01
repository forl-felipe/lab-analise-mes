"""Extrai o logo Samarco (SVG) do kit visual e gera os PNGs usados na planilha.

Uso: python3 gerar_logos.py <caminho do samarco-kit-visual.html>
Requer: pip install cairosvg
"""
import os
import re
import sys

import cairosvg

AQUI = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(AQUI, 'assets')

# linha do kit com cada versao do logo e cor de fundo do PNG
VERSOES = {'logo_fundo_escuro': (64, '#00335A'), 'logo_fundo_claro': (67, '#FFFFFF')}


def main(kit):
    os.makedirs(ASSETS, exist_ok=True)
    linhas = open(kit, encoding='utf-8').read().split('\n')
    for nome, (linha, fundo) in VERSOES.items():
        svg = re.search(r'<svg.*?</svg>', linhas[linha - 1]).group(0)
        with open(os.path.join(ASSETS, nome + '.svg'), 'w', encoding='utf-8') as fh:
            fh.write(svg)
        png = os.path.join(ASSETS, nome + '.png')
        cairosvg.svg2png(bytestring=svg.encode('utf-8'), write_to=png, output_width=940,
                         output_height=206, background_color=fundo)
        print(nome, os.path.getsize(png), 'bytes')


if __name__ == '__main__':
    main(sys.argv[1])
