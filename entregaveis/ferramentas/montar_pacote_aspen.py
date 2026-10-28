"""Monta o pacote de execução da tela no Aspen a partir das imagens e posições geradas por gerar_imagens.js.

Uso (na pasta ferramentas):
    node gerar_imagens.js
    python montar_pacote_aspen.py

Saídas em ../03_aspen/:
    posicoes_objetos.xlsx                          posições em unidades do Aspen (tela de 50 x 35 unidades)
    posicoes_objetos.csv                           mesmas informações em texto
    roteiro_aspen_graphics.md / .html / .pdf       passo a passo completo, com uma ficha por objeto

Requer: openpyxl, markdown e Playwright (para o PDF).
"""
import csv
import json
import os
import subprocess

import markdown
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

AQUI = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(AQUI, "..", "03_aspen")

TEXTO = "230, 234, 240"
ROTULO = "139, 150, 165"
LARANJA = "242, 153, 74"
PAINEL = "31, 38, 48"
CABEC = "17, 22, 29"
HORA = "`VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B)"
QUAL = "`VAL` no arquivo: **trocar pelo atributo de qualidade** (Parte B)"
DISC = "`IP_DISCRETEMAP` no arquivo: **conferir** (Parte B)"
REGRA_QUAL = f"Texto igual a `Good`: RGB {TEXTO}. Qualquer outro texto: RGB {LARANJA}"
DH = "dd/mm/aaaa hh:mm:ss"


def campo(nome, tag, mapa, atrib, pt, negrito, alin, fundo, formato, regra="Nenhuma", obs=""):
    return dict(ferramenta="Campo de dados", nome=nome, tag=tag, mapa=mapa, atrib=atrib, pt=pt, negrito=negrito,
                alin=alin, cor=TEXTO, fundo=fundo, formato=formato, regra=regra, obs=obs)


def tendencia(nome, obs, tag=""):
    return dict(ferramenta="Tendência (plot)", nome=nome, obs=obs, tag=tag)


SPEC = {
    1: dict(ferramenta="Imagem", nome="Logotipo oficial", obs="Arquivo oficial fornecido pela Comunicação. Manter a proporção dentro do retângulo. Ver Parte D."),
    2: campo("Última atualização (cabeçalho)", "307-GI-001", "IP_ANALOGMAP", HORA, 14, True, "Direita", CABEC, DH,
             obs="Fica no cabeçalho, que tem fundo mais escuro que os painéis."),
    3: campo("Carga média: valor", "307-GI-001", "IP_ANALOGMAP", "VAL", 57, True, "Direita", PAINEL, "1 casa decimal (0,0)",
             obs="Objeto de calibração da fonte (ver início da Parte E). A unidade daN/pelota já está no fundo."),
    4: tendencia("Carga média: minitendência", "Já vem pronta no arquivo. Ver Parte C.", "307-GI-001"),
    5: campo("Carga média: atualizado em", "307-GI-001", "IP_ANALOGMAP", HORA, 11, False, "Esquerda", PAINEL, DH),
    6: campo("Carga média: qualidade", "307-GI-001", "IP_ANALOGMAP", QUAL, 11, True, "Esquerda", PAINEL, "Texto", REGRA_QUAL),
    7: campo("Desvio padrão: valor", "307-GI-002", "IP_ANALOGMAP", "VAL", 57, True, "Direita", PAINEL, "1 casa decimal (0,0)",
             obs="Copie o objeto 3 e troque a tag e a posição."),
    8: tendencia("Desvio padrão: minitendência", "Já vem pronta no arquivo. Ver Parte C.", "307-GI-002"),
    9: campo("Desvio padrão: atualizado em", "307-GI-002", "IP_ANALOGMAP", HORA, 11, False, "Esquerda", PAINEL, DH,
             obs="Copie o objeto 5 e troque a tag e a posição."),
    10: campo("Desvio padrão: qualidade", "307-GI-002", "IP_ANALOGMAP", QUAL, 11, True, "Esquerda", PAINEL, "Texto", REGRA_QUAL,
              obs="Copie o objeto 6 e troque a tag e a posição."),
    11: campo("Compressão MES: último resultado", "M4710050020-0031-HHLFU", "IP_MESVALOR", "VAL", 57, True, "Direita", PAINEL,
              "Sem casas decimais (0)", obs="Mesmo vínculo do campo Compressão da tela PQ US04. Mesma fonte do objeto 3: copie o 3 e troque tag, mapa e formato."),
    12: campo("Compressão MES: horário do resultado", "M4710050020-0031-HHLFU", "IP_MESVALOR", HORA, 11, False, "Esquerda", PAINEL,
              "dd/mm/aaaa hh:mm", obs="Copie o objeto 5 e troque tag, mapa e formato."),
    13: campo("Modo do cortador 07CR001", "U04-07CR001", "IP_PIDMAP03", "VAL", 14, True, "Direita", PAINEL, "Texto",
              f"Texto `Automático`: RGB {TEXTO}. Texto `Manual`: RGB {LARANJA}",
              obs="Mesmo vínculo do indicador Manual/Automático da tela PQ US04. Se o campo mostrar número, use a regra de cor pelo número equivalente."),
    14: campo("Modo do cortador 07CR002", "U04-07CR002", "IP_PIDMAP03", "VAL", 14, True, "Direita", PAINEL, "Texto",
              f"Texto `Automático`: RGB {TEXTO}. Texto `Manual`: RGB {LARANJA}", obs="Copie o objeto 13 e troque a tag."),
    15: tendencia("Tendência principal da carga média", "Já vem pronta no arquivo, com as penas e a escala. Ver Parte C.", "307-GI-001 e M4710050020-0031-HHLFU (ver Parte F)"),
    16: campo("Faixa menor que 200 (Atenção)", "307-QI-003", DISC, "VAL", 18, True, "Direita", PAINEL, "Texto (ou número inteiro)",
              f"Valor `1` ou `Ativo`: RGB 242, 201, 76. Valor `0` ou `Inativo`: RGB {ROTULO}"),
    17: campo("Faixa menor que 150 (Alerta)", "307-QI-002", DISC, "VAL", 18, True, "Direita", PAINEL, "Texto (ou número inteiro)",
              f"Valor `1` ou `Ativo`: RGB 242, 153, 74. Valor `0` ou `Inativo`: RGB {ROTULO}", obs="Copie o objeto 16."),
    18: campo("Faixa menor que 100 (Crítico)", "307-QI-001", DISC, "VAL", 18, True, "Direita", PAINEL, "Texto (ou número inteiro)",
              f"Valor `1` ou `Ativo`: RGB 235, 87, 87. Valor `0` ou `Inativo`: RGB {ROTULO}", obs="Copie o objeto 16."),
}


def un(v):
    """Unidades do Aspen com 2 casas e vírgula decimal."""
    return f"{v:.2f}".replace(".", ",")


def pt(v):
    """Tamanhos de fonte foram definidos para 1920 px de largura; a tela de 50 unidades aparece com cerca de 1400 px."""
    return max(8, round(v * 0.84))


def ficha(n, s, pos):
    linhas = [f"### Objeto {n}: {s['nome']}", "",
              "| Propriedade | Valor |", "|---|---|",
              f"| Ferramenta | {s['ferramenta']} |",
              f"| Posição em unidades (X ; Y) | {un(pos['x'])} ; {un(pos['y'])} |",
              f"| Tamanho em unidades (largura ; altura) | {un(pos['w'])} ; {un(pos['h'])} |"]
    if s["ferramenta"] == "Campo de dados":
        linhas += [f"| Tag | `{s['tag']}` |",
                   f"| Mapa | {s['mapa'] if '`' in s['mapa'] else '`' + s['mapa'] + '`'} |",
                   f"| Atributo | {s['atrib'] if '`' in s['atrib'] else '`' + s['atrib'] + '`'} |",
                   f"| Fonte | Arial{' negrito' if s['negrito'] else ''}, já definida no arquivo |",
                   f"| Alinhamento | {s['alin']} |",
                   f"| Cor do texto | RGB {s['cor']} |",
                   f"| Cor de fundo | RGB {s['fundo']} (ou transparente), sem borda |",
                   f"| Formato | {s['formato']} |",
                   f"| Regra de cor | {s['regra']} |"]
    if s.get("obs"):
        linhas.append(f"| Observação | {s['obs']} |")
    return "\n".join(linhas) + "\n"


def planilha(posicoes, caminho):
    wb = Workbook()
    ws = wb.active
    ws.title = "Posicoes"
    f = "Arial"
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    cab = PatternFill("solid", fgColor="1F2630")
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    ws["A1"] = "Largura da tela no editor, em unidades. O padrão do Aspen é 50; só altere se a tela tiver outro tamanho."
    ws["A1"].font = Font(name=f, bold=True)
    ws["F1"] = 50
    ws["F1"].fill = amarelo
    ws["F1"].border = borda
    ws["F1"].font = Font(name=f, bold=True, color="0000FF")
    ws["A2"] = "Fator de conversão (calculado)"
    ws["A2"].font = Font(name=f)
    ws["F2"] = "=F1/50"
    ws["F2"].number_format = "0.0000"
    ws["F2"].font = Font(name=f)
    titulos = ["Nº", "Objeto", "Ferramenta", "Tag", "Mapa", "Atributo", "X (tela 50)", "Y (tela 50)", "Largura (tela 50)", "Altura (tela 50)",
               "X no editor", "Y no editor", "Largura no editor", "Altura no editor"]
    larg = [5, 38, 18, 26, 22, 30, 8, 8, 10, 10, 10, 10, 12, 12]
    for c, (t, w) in enumerate(zip(titulos, larg), 1):
        cel = ws.cell(row=4, column=c, value=t)
        cel.font = Font(name=f, bold=True, color="FFFFFF")
        cel.fill = cab
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cel.column_letter].width = w
    linhas = [dict(n="Tela", nome="Tela inteira (referência)", ferramenta="-", tag="", mapa="", atrib="", x=0, y=0, w=50, h=35)]
    for p in posicoes:
        s = SPEC[p["n"]]
        linhas.append(dict(n=p["n"], nome=s["nome"], ferramenta=s["ferramenta"], tag=s.get("tag", ""),
                           mapa=s.get("mapa", "").replace("`", ""), atrib=s.get("atrib", "").replace("*", ""),
                           x=p["x"], y=p["y"], w=p["w"], h=p["h"]))
    for r, l in enumerate(linhas, 5):
        vals = [l["n"], l["nome"], l["ferramenta"], l["tag"], l["mapa"], l["atrib"], l["x"], l["y"], l["w"], l["h"]]
        for c, v in enumerate(vals, 1):
            cel = ws.cell(row=r, column=c, value=v)
            cel.font = Font(name=f)
            if c >= 7:
                cel.number_format = "0.00"
        for c, src in zip(range(11, 15), "GHIJ"):
            cel = ws.cell(row=r, column=c, value=f"=ROUND({src}{r}*$F$2,2)")
            cel.number_format = "0.00"
            cel.font = Font(name=f, bold=True)
            cel.fill = PatternFill("solid", fgColor="F2F2F2")
    ws.freeze_panes = "A5"
    wb.save(caminho)
    return linhas


def main():
    with open(os.path.join(SAIDA, "posicoes.json"), encoding="utf-8") as fh:
        posicoes = json.load(fh)
    assert {p["n"] for p in posicoes} == set(SPEC), "Objetos do protótipo e da especificação não conferem"


    linhas = planilha(posicoes, os.path.join(SAIDA, "posicoes_objetos.xlsx"))
    with open(os.path.join(SAIDA, "posicoes_objetos.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["n", "objeto", "ferramenta", "tag", "mapa", "atributo", "x_unid", "y_unid", "largura_unid", "altura_unid"])
        for l in linhas:
            w.writerow([l["n"], l["nome"], l["ferramenta"], l["tag"], l["mapa"], l["atrib"], un(l["x"]), un(l["y"]), un(l["w"]), un(l["h"])])

    with open(os.path.join(AQUI, "roteiro_modelo.md"), encoding="utf-8") as fh:
        modelo = fh.read()
    fichas = "\n".join(ficha(p["n"], SPEC[p["n"]], p) for p in posicoes)
    n_campos = sum(1 for s in SPEC.values() if s["ferramenta"] == "Campo de dados")
    n_tend = sum(1 for s in SPEC.values() if s["ferramenta"].startswith("Tendência"))
    with open(os.path.join(SAIDA, "objetos_tela.json"), encoding="utf-8") as fh:
        tela = json.load(fh)
    texto = (modelo.replace("{{FICHAS}}", fichas).replace("{{N_CAMPOS}}", str(n_campos)).replace("{{N_TEND}}", str(n_tend))
             .replace("{{N_FORMAS}}", str(len(tela["formas"]) + len(tela["simbolos"]))).replace("{{N_TEXTOS}}", str(len(tela["textos"]))))
    md_path = os.path.join(SAIDA, "roteiro_aspen_graphics.md")
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write(texto)

    # O conversor exige linha em branco antes de listas; o .md fica como está.
    saida_md, anterior = [], ""
    for linha in texto.splitlines():
        eh_lista = linha.lstrip().startswith(("- ", "1. ", "2. ", "3. ", "4. ", "5. ", "6. "))
        if eh_lista and anterior.strip() and not anterior.lstrip().startswith(("- ", "1. ", "2. ", "3. ", "4. ", "5. ", "6. ", ">")):
            saida_md.append("")
        saida_md.append(linha)
        anterior = linha
    corpo = markdown.markdown("\n".join(saida_md), extensions=["tables"])
    html = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>Passo a passo Aspen</title>
<style>
body {{ font-family: Arial, sans-serif; font-size: 14px; color: #1b1f24; max-width: 1000px; margin: 24px auto; padding: 0 16px; line-height: 1.5; }}
h1 {{ font-size: 24px; border-bottom: 3px solid #1f2630; padding-bottom: 6px; }}
h2 {{ font-size: 19px; margin-top: 32px; background: #1f2630; color: #fff; padding: 6px 10px; border-radius: 4px; page-break-after: avoid; }}
h3 {{ font-size: 16px; margin-top: 22px; color: #0b5c80; page-break-after: avoid; }}
table {{ border-collapse: collapse; width: 100%; margin: 8px 0 14px; page-break-inside: avoid; }}
th, td {{ border: 1px solid #c5ccd6; padding: 5px 8px; text-align: left; vertical-align: top; }}
th {{ background: #e9edf2; }}
td:first-child {{ white-space: nowrap; }}
code {{ background: #eef1f5; padding: 1px 4px; border-radius: 3px; font-size: 13px; }}
img {{ max-width: 100%; border: 1px solid #c5ccd6; }}
blockquote {{ border-left: 4px solid #f2994a; margin: 10px 0; padding: 4px 12px; background: #fff7ef; }}
hr {{ border: 0; border-top: 1px solid #c5ccd6; }}
</style></head><body>{corpo}</body></html>"""
    html_path = os.path.join(SAIDA, "roteiro_aspen_graphics.html")
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html)

    pdf_path = os.path.join(SAIDA, "roteiro_aspen_graphics.pdf")
    js = ("const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();"
          f"await p.goto('file://{html_path}');await p.pdf({{path:'{pdf_path}',format:'A4',margin:{{top:'15mm',bottom:'15mm',left:'12mm',right:'12mm'}},"
          "printBackground:true,displayHeaderFooter:true,headerTemplate:'<span></span>',"
          "footerTemplate:'<div style=\"font-size:8px;width:100%;text-align:center\">Página <span class=pageNumber></span> de <span class=totalPages></span></div>'});"
          "await b.close()})()")
    subprocess.run(["node", "-e", js], check=True)
    print(f"Pacote montado em {os.path.abspath(SAIDA)}: {len(SPEC)} objetos")


if __name__ == "__main__":
    main()
