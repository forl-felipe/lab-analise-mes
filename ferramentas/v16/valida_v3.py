# -*- coding: utf-8 -*-
import os, re, json, glob
MOD = '/home/user/lab-analise-mes/Painel Samarco/Painel.SemanticModel'
REP = '/tmp/claude-0/-home-user/0dd78099-f388-5411-9f86-70215a660d00'
REP = '/home/user/lab-analise-mes/Painel Samarco/Painel.Report'
PG  = REP + '/definition/pages'

cols, meas, tabs = set(), set(), set()
for f in glob.glob(MOD + '/definition/tables/*.tmdl'):
    txt = open(f, encoding='utf-8').read().replace('\r\n', '\n')
    m0 = re.match(r"table ('([^']+)'|(\S+))", txt)
    tname = m0.group(2) or m0.group(3)
    tabs.add(tname)
    for m in re.finditer(r"^\tcolumn ('([^']+)'|(\S+))", txt, re.M):
        cols.add((tname, m.group(2) or m.group(3)))
    for m in re.finditer(r"^\tmeasure ('([^']+)'|(\S+))", txt, re.M):
        meas.add((tname, m.group(2) or m.group(3)))
print('modelo: %d tabelas · %d colunas · %d medidas' % (len(tabs), len(cols), len(meas)))
dup = {}
for t, c in meas: dup.setdefault(c, []).append(t)
d2 = {k: v for k, v in dup.items() if len(v) > 1}
print('medidas com nome duplicado em tabelas diferentes:', d2 or 'nenhuma')

ref_c, ref_m = set(), set()
def walk(o):
    if isinstance(o, dict):
        if isinstance(o.get('Measure'), dict):
            e = o['Measure'].get('Expression', {}).get('SourceRef', {}).get('Entity')
            if e: ref_m.add((e, o['Measure'].get('Property')))
        if isinstance(o.get('Column'), dict):
            e = o['Column'].get('Expression', {}).get('SourceRef', {}).get('Entity')
            if e: ref_c.add((e, o['Column'].get('Property')))
        for v in o.values(): walk(v)
    elif isinstance(o, list):
        for v in o: walk(v)

# 1. JSON válido
ruim = []
todos = glob.glob(PG + '/*/visuals/*/visual.json') + glob.glob(PG + '/*/page.json') + [PG + '/pages.json']
for f in todos:
    try: walk(json.load(open(f, encoding='utf-8')))
    except Exception as ex: ruim.append('%s -> %s' % (f.split('/pages/')[-1], ex))
print('JSON: %d arquivos · %s' % (len(todos), 'INVÁLIDOS: ' + str(ruim) if ruim else 'todos válidos'))

falhas = []
for e, p in sorted(ref_c):
    if (e, p) not in cols: falhas.append('COLUNA %s[%s]' % (e, p))
for e, p in sorted(ref_m):
    if (e, p) not in meas: falhas.append('MEDIDA %s[%s]' % (e, p))
print('referências: %d colunas · %d medidas · %s'
      % (len(ref_c), len(ref_m), ('QUEBRADAS: ' + '; '.join(falhas)) if falhas else 'todas resolvem'))

# 2. recursos registrados
res = set(os.listdir(REP + '/StaticResources/RegisteredResources'))
usados = set()
for f in glob.glob(PG + '/*/visuals/*/visual.json'):
    for m in re.finditer(r'"ItemName":\s*"([^"]+)"', open(f, encoding='utf-8').read()):
        usados.add(m.group(1))
print('recursos: %d registrados · faltando -> %s · não usados -> %s'
      % (len(res), sorted(usados - res) or 'nenhum',
         sorted(r for r in res - usados if r.endswith('.png')) or 'nenhum'))

# 3. destinos de navegação
paginas = {p for p in os.listdir(PG) if os.path.isdir(PG + '/' + p)}
navs = set()
for f in glob.glob(PG + '/*/visuals/*/visual.json'):
    for m in re.finditer(r'"navigationSection".*?"Value":\s*"\'([^\']+)\'"',
                         open(f, encoding='utf-8').read(), re.S):
        navs.add(m.group(1))
print('navegação: %d destinos · inexistentes -> %s' % (len(navs), sorted(navs - paginas) or 'nenhum'))

# 4. layout: limites e sobreposição
W, H = 1920, 1080
FUNDO = ('textbox',)   # barras de fundo podem estar sob tudo
prob = []
for pid in sorted(paginas):
    _pj = json.load(open(PG + '/' + pid + '/page.json', encoding='utf-8'))
    nome = _pj['displayName']; W, H = _pj.get('width', 1920), _pj.get('height', 1080)
    vs = []
    for v in sorted(os.listdir(PG + '/' + pid + '/visuals')):
        j = json.load(open(PG + '/' + pid + '/visuals/' + v + '/visual.json', encoding='utf-8'))
        p = j['position']; t = j.get('visual', {}).get('visualType')
        vs.append((v, t, p['x'], p['y'], p['width'], p['height'], p.get('z', 0)))
        if p['x'] < 0 or p['y'] < 0 or p['x'] + p['width'] > W or p['y'] + p['height'] > H:
            prob.append('%s: %s(%s) fora da tela x=%s y=%s w=%s h=%s'
                        % (nome, v[:6], t, p['x'], p['y'], p['width'], p['height']))
    for i in range(len(vs)):
        for k in range(i + 1, len(vs)):
            a, b = vs[i], vs[k]
            # ignora pares intencionais: botão+ícone, card+ícone, barra de fundo
            if {a[1], b[1]} <= {'image', 'actionButton'}: continue
            if {a[1], b[1]} == {'image', 'cardVisual'}: continue
            if a[1] == 'textbox' and a[4] >= 224 and a[5] >= 88 : continue
            if b[1] == 'textbox' and b[4] >= 224 and b[5] >= 88 : continue
            ox = min(a[2]+a[4], b[2]+b[4]) - max(a[2], b[2])
            oy = min(a[3]+a[5], b[3]+b[5]) - max(a[3], b[3])
            if ox > 0 and oy > 0:
                prob.append('%s: %s(%s) x %s(%s) sobrepõem %dx%d px'
                            % (nome, a[0][:6], a[1], b[0][:6], b[1], ox, oy))
print('layout: %s' % ('\n   ' + '\n   '.join(prob) if prob else 'sem sobreposição nem estouro de tela'))

# 5. resourcePackages x arquivos em disco
import json as _j
rj = _j.load(open(REP + '/definition/report.json', encoding='utf-8'))
decl = {i['name'] for pk in rj['resourcePackages'] if pk['name'] == 'RegisteredResources' for i in pk['items']}
disco = set(os.listdir(REP + '/StaticResources/RegisteredResources'))
print('resourcePackages: declarado sem arquivo -> %s · arquivo sem declaração -> %s · usado sem declaração -> %s'
      % (sorted(decl - disco) or 'nenhum', sorted(disco - decl) or 'nenhum', sorted(usados - decl) or 'nenhum'))

# 6. os $schema dos arquivos de controle (o .pbip foi escrito a mao por mim)
import re as _re
PROJ = os.path.dirname(REP)
CONTROLE = [
 (PROJ + '/Painel.pbip',
  r'^https://developer\.microsoft\.com/json-schemas/fabric/pbip/pbipProperties/1\.\d+\.\d+/schema\.json$'),
 (REP + '/definition.pbir',
  r'^https://developer\.microsoft\.com/json-schemas/fabric/item/report/definitionProperties/\d+\.\d+(\.\d+)?/schema\.json$'),
 (REP + '/definition/report.json',
  r'^https://developer\.microsoft\.com/json-schemas/fabric/item/report/definition/report/'),
 (REP + '/definition/pages/pages.json',
  r'^https://developer\.microsoft\.com/json-schemas/fabric/item/report/definition/pagesMetadata/'),
]
ruins = []
for f, pat in CONTROLE:
    s = _j.load(open(f, encoding='utf-8')).get('$schema', '')
    if not _re.match(pat, s):
        ruins.append('%s -> %s' % (os.path.basename(f), s))
print('$schema dos arquivos de controle: %s' % ('INVALIDO: ' + '; '.join(ruins) if ruins else 'todos no padrao'))
