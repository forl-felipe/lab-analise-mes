# -*- coding: utf-8 -*-
"""Confere o M de cada partição: bloco inteiro e delimitadores balanceados."""
import io, glob, os, sys
DIR = "Painel Samarco/Painel.SemanticModel/definition/tables"
IND = "\t\t\t\t"

def bloco_m(caminho):
    linhas = io.open(caminho, encoding="utf-8", newline="").read().split("\r\n")
    # so particoes "= m"; as calculadas sao DAX e nao tem 'in'
    try:
        p = next(k for k, l in enumerate(linhas) if l.strip().startswith("partition "))
    except StopIteration:
        return None, []
    if not linhas[p].rstrip().endswith("= m"): return None, []
    try: i = next(k for k, l in enumerate(linhas) if k > p and l.strip() == "source =")
    except StopIteration: return None, []
    corpo, vazias = [], []
    for k in range(i + 1, len(linhas)):
        l = linhas[k]
        if l == "":                       # fim do bloco no TMDL
            break
        if not l.startswith(IND):
            vazias.append(("linha %d nao recuada dentro do M" % (k + 1), l[:40]))
            break
        corpo.append(l[len(IND):])
    return "\n".join(corpo), vazias

def balanco(m):
    pares = {")": "(", "]": "[", "}": "{"}
    pilha, i, n = [], 0, len(m)
    while i < n:
        c = m[i]
        if c == '"':                       # string (aspas duplas dobram para escapar)
            i += 1
            while i < n:
                if m[i] == '"':
                    if i + 1 < n and m[i+1] == '"': i += 2; continue
                    break
                i += 1
        elif c == "/" and i + 1 < n and m[i+1] == "/":
            while i < n and m[i] != "\n": i += 1
        elif c == "/" and i + 1 < n and m[i+1] == "*":
            i = m.find("*/", i) + 1
        elif c in "([{": pilha.append((c, i))
        elif c in pares:
            if not pilha or pilha[-1][0] != pares[c]:
                return "fecha %r sem abrir, posicao %d" % (c, i)
            pilha.pop()
        i += 1
    if pilha:
        ab, pos = pilha[-1]
        linha = m[:pos].count("\n") + 1
        return "abre %r na linha %d do M e nunca fecha" % (ab, linha)
    return None

problemas, conferidas = [], 0
for f in sorted(glob.glob(DIR + "/*.tmdl")):
    nome = os.path.basename(f)
    m, avisos = bloco_m(f)
    if m is None: continue
    conferidas += 1
    for a in avisos: problemas.append("%s: %s %r" % (nome, a[0], a[1]))
    if not m.rstrip().split("\n")[-1].strip():
        problemas.append("%s: bloco M termina em branco (truncado?)" % nome)
    linhas = [l for l in m.split("\n") if l.strip()]
    corpo = m.split("\n")
    if not any(l == "in" for l in corpo):
        problemas.append("%s: bloco M sem o 'in' de topo (truncado por linha vazia?)" % nome)
    else:
        k = max(i for i, l in enumerate(corpo) if l == "in")
        if not [l for l in corpo[k+1:] if l.strip()]:
            problemas.append("%s: nada depois do 'in' de topo" % nome)
    b = balanco(m)
    if b: problemas.append("%s: %s" % (nome, b))

print("M das particoes: %s" % ("\n   " + "\n   ".join(problemas) if problemas else "%d tabelas, todas fechadas e balanceadas" % conferidas))
falhou = bool(problemas)

# ── virgula entre os passos do let ───────────────────────────────────────────
import re as _re
PASSO = _re.compile(r"^    [A-Za-z_][A-Za-z0-9_]* =($| )")   # binding no nivel do let
faltando = []
for f in sorted(glob.glob(DIR + "/*.tmdl")):
    nome = os.path.basename(f)
    m, _ = bloco_m(f)
    if m is None: continue
    linhas = m.split("\n")
    for k, l in enumerate(linhas):
        if not PASSO.match(l): continue
        # ultima linha com conteudo antes deste passo, ignorando comentario
        j = k - 1
        while j >= 0 and (not linhas[j].strip() or linhas[j].strip().startswith("//")):
            j -= 1
        if j < 0: continue
        ant = linhas[j].rstrip()
        if ant.strip() in ("let",) or ant.endswith((",", "(", "[", "{")): continue
        faltando.append("%s: linha %d do M comeca um passo mas a anterior nao termina em virgula -> %r"
                        % (nome, k + 1, ant.strip()[-60:]))
print("virgulas entre passos do let: %s"
      % ("\n   " + "\n   ".join(faltando) if faltando else "todas no lugar"))
sys.exit(1 if (falhou or faltando) else 0)
