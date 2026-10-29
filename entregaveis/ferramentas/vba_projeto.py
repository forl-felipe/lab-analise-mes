"""Monta um projeto VBA (vbaProject.bin) a partir do código-fonte e o injeta num .xlsx, gerando um .xlsm.

Segue a especificação MS-OVBA: fontes comprimidas com o algoritmo do VBA, fluxo "dir" com os módulos,
fluxo PROJECT com as chaves de proteção (CMG, DPB e GC sem senha) e _VBA_PROJECT na versão 0xFFFF,
que faz o Excel recompilar o código ao abrir (não há p-code pré-compilado).

Uso no código:
    from vba_projeto import gerar_xlsm
    gerar_xlsm("base.xlsx", "saida.xlsm", modulos, documentos)
      modulos    = {"ModPIMS": "código..."}                     (módulos padrão, .bas)
      documentos = {"ThisWorkbook": "código...", "shPainel": ""} (pasta de trabalho e abas, pelo codeName)
"""
import os
import random
import re
import struct
import uuid
import zipfile

CP = "cp1252"


# ---------------------------------------------------------------- compressão MS-OVBA 2.4.1
def comprimir(dados: bytes) -> bytes:
    saida = bytearray(b"\x01")
    pos = 0
    while pos < len(dados):
        bloco = dados[pos:pos + 4096]
        pos += len(bloco)
        corpo = bytearray()
        i = 0
        while i < len(bloco):
            flags, tokens = 0, bytearray()
            for bit in range(8):
                if i >= len(bloco):
                    break
                # tamanho do deslocamento depende da posição dentro do bloco
                dif = i
                nbits = max((dif - 1).bit_length(), 4) if dif > 0 else 4
                max_len = (0xFFFF >> nbits) + 3
                melhor_len, melhor_off = 0, 0
                inicio = max(0, i - (1 << nbits))
                for c in range(i - 1, inicio - 1, -1):
                    n = 0
                    while n < max_len and i + n < len(bloco) and bloco[c + n] == bloco[i + n]:
                        n += 1
                    if n > melhor_len:
                        melhor_len, melhor_off = n, i - c
                        if n == max_len:
                            break
                if melhor_len >= 3:
                    token = ((melhor_off - 1) << (16 - nbits)) | (melhor_len - 3)
                    tokens += struct.pack("<H", token)
                    flags |= 1 << bit
                    i += melhor_len
                else:
                    tokens.append(bloco[i])
                    i += 1
            corpo.append(flags)
            corpo += tokens
        if len(corpo) > 4096 and len(bloco) == 4096:
            saida += struct.pack("<H", 0x3FFF) + bloco    # bloco sem compressão (só para blocos cheios)
            continue
        assert len(corpo) <= 4096, "bloco final sem compressão possível"
        cabec = 0xB000 | (len(corpo) + 2 - 3)   # assinatura 0b011, bloco comprimido
        saida += struct.pack("<H", cabec) + corpo
    return bytes(saida)


def descomprimir(dados: bytes) -> bytes:
    assert dados[0] == 1
    saida = bytearray()
    p = 1
    while p < len(dados):
        cabec = struct.unpack_from("<H", dados, p)[0]
        tam = (cabec & 0x0FFF) + 3
        comprimido = cabec & 0x8000
        fim = p + tam
        p += 2
        base = len(saida)
        if not comprimido:
            saida += dados[p:p + 4096]
            p += 4096
            continue
        while p < fim:
            flags = dados[p]
            p += 1
            for bit in range(8):
                if p >= fim:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", dados, p)[0]
                    p += 2
                    dif = len(saida) - base
                    nbits = max((dif - 1).bit_length(), 4)
                    comp = (token & (0xFFFF >> nbits)) + 3
                    off = (token >> (16 - nbits)) + 1
                    for _ in range(comp):
                        saida.append(saida[-off])
                else:
                    saida.append(dados[p])
                    p += 1
    return bytes(saida)


# ---------------------------------------------------------------- criptografia de dados MS-OVBA 2.4.3
def cifrar(id_projeto: str, dados: bytes, semente=None) -> str:
    semente = random.randint(0, 255) if semente is None else semente
    versao = 2
    chave = sum(id_projeto.encode(CP)) & 0xFF
    out = bytearray([semente, semente ^ versao, semente ^ chave])
    nao_cif, cif1, cif2 = chave, semente ^ chave, semente ^ versao
    for _ in range((semente & 6) // 2):
        temp = 0
        b = temp ^ ((cif2 + nao_cif) & 0xFF)
        out.append(b)
        cif2, cif1, nao_cif = cif1, b, temp
    for byte in struct.pack("<I", len(dados)) + dados:
        b = byte ^ ((cif2 + nao_cif) & 0xFF)
        out.append(b)
        cif2, cif1, nao_cif = cif1, b, byte
    return out.hex().upper()


def decifrar(id_projeto: str, texto: str) -> bytes:
    d = bytes.fromhex(texto)
    semente = d[0]
    chave = d[2] ^ semente
    nao_cif, cif1, cif2 = chave, d[2], d[1] ^ semente  # versão
    cif1 = d[2]
    cif2 = d[1]
    p = 3
    for _ in range((semente & 6) // 2):
        b = d[p]
        nao = b ^ ((cif2 + nao_cif) & 0xFF)
        cif2, cif1, nao_cif = cif1, b, nao
        p += 1
    saida = bytearray()
    for b in d[p:]:
        nao = b ^ ((cif2 + nao_cif) & 0xFF)
        saida.append(nao)
        cif2, cif1, nao_cif = cif1, b, nao
    n = struct.unpack_from("<I", saida, 0)[0]
    return bytes(saida[4:4 + n])


# ---------------------------------------------------------------- fluxos do projeto
def _rec(id_, dados: bytes) -> bytes:
    return struct.pack("<HI", id_, len(dados)) + dados


def fluxo_dir(modulos):
    """modulos: lista de (nome, documento: bool)."""
    nome = b"VBAProject"
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 3))              # SysKind Win64 (como o Relatório Físico)
    d += _rec(0x004A, struct.pack("<I", 6))              # versão de compatibilidade
    d += _rec(0x0002, struct.pack("<I", 0x0409))         # LCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))         # LCID invoke
    d += _rec(0x0003, struct.pack("<H", 1252))           # code page
    d += _rec(0x0004, nome)
    d += _rec(0x0005, b"") + _rec(0x0040, b"")           # doc string
    d += _rec(0x0006, b"") + _rec(0x003D, b"")           # help file
    d += _rec(0x0007, struct.pack("<I", 0))              # help context
    d += _rec(0x0008, struct.pack("<I", 0))              # lib flags
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)           # versão do projeto
    d += _rec(0x000C, b"") + _rec(0x003C, b"")           # constantes
    # referência OLE Automation (stdole); VBA e Excel são implícitas no projeto do Excel
    libid = b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
    d += _rec(0x0016, b"stdole") + _rec(0x003E, "stdole".encode("utf-16-le"))
    d += _rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))
    d += _rec(0x000F, struct.pack("<H", len(modulos)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))
    for m, documento in modulos:
        mb, mu = m.encode(CP), m.encode("utf-16-le")
        d += _rec(0x0019, mb) + _rec(0x0047, mu)
        d += _rec(0x001A, mb) + _rec(0x0032, mu)
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))          # offset do fonte no fluxo
        d += _rec(0x001E, struct.pack("<I", 0))
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))
        d += struct.pack("<HI", 0x0022 if documento else 0x0021, 0)
        d += struct.pack("<HI", 0x002B, 0)
    d += struct.pack("<HI", 0x0010, 0)
    return bytes(d)


def fluxo_project(id_projeto, modulos):
    linhas = [f'ID="{id_projeto}"']
    for m, documento in modulos:
        linhas.append(f"Document={m}/&H00000000" if documento else f"Module={m}")
    linhas += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
               f'CMG="{cifrar(id_projeto, struct.pack("<I", 0))}"',
               f'DPB="{cifrar(id_projeto, bytes([0]))}"',
               f'GC="{cifrar(id_projeto, bytes([0xFF]))}"',
               "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "",
               "[Workspace]"]
    linhas += [f"{m}=0, 0, 0, 0, C" for m, _ in modulos]
    return ("\r\n".join(linhas) + "\r\n").encode(CP)


def fluxo_projectwm(modulos):
    d = bytearray()
    for m, _ in modulos:
        d += m.encode(CP) + b"\0" + m.encode("utf-16-le") + b"\0\0"
    return bytes(d + b"\0\0")


BASE_DOC = {"ThisWorkbook": "0{00020819-0000-0000-C000-000000000046}"}
BASE_ABA = "0{00020820-0000-0000-C000-000000000046}"


def fonte_documento(nome, codigo):
    base = BASE_DOC.get(nome, BASE_ABA)
    attrs = [f'Attribute VB_Name = "{nome}"', f'Attribute VB_Base = "{base}"', "Attribute VB_GlobalNameSpace = False",
             "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = True",
             "Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = True"]
    return "\r\n".join(attrs) + "\r\n" + codigo


def fonte_modulo(nome, codigo):
    codigo = re.sub(r'^Attribute VB_Name = ".*?"\r?\n', "", codigo)
    return f'Attribute VB_Name = "{nome}"\r\n' + codigo


# ---------------------------------------------------------------- arquivo composto (CFB v3) com pastas e fluxos
def _cfb(entradas):
    """entradas: lista de (caminho, bytes) com caminhos tipo 'PROJECT' ou 'VBA/dir'."""
    SETOR, MINI = 512, 64
    FREE, END, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD

    # árvore de diretório
    nos = [{"nome": "Root Entry", "tipo": 5, "filhos": [], "dados": b""}]
    pastas = {"": 0}
    for caminho, dados in entradas:
        partes = caminho.split("/")
        pai = ""
        for k, parte in enumerate(partes):
            chave = "/".join(partes[:k + 1])
            if k < len(partes) - 1:
                if chave not in pastas:
                    nos.append({"nome": parte, "tipo": 1, "filhos": [], "dados": b""})
                    pastas[chave] = len(nos) - 1
                    nos[pastas[pai]]["filhos"].append(pastas[chave])
            else:
                nos.append({"nome": parte, "tipo": 2, "filhos": [], "dados": dados})
                nos[pastas[pai]]["filhos"].append(len(nos) - 1)
            pai = chave if k < len(partes) - 1 else pai

    def chave_ordem(i):
        n = nos[i]["nome"]
        return (len(n), n.upper())

    esq = [FREE] * len(nos)
    dir_ = [FREE] * len(nos)
    filho = [FREE] * len(nos)

    def arvore(lista):
        if not lista:
            return FREE
        lista = sorted(lista, key=chave_ordem)
        m = len(lista) // 2
        raiz = lista[m]
        esq[raiz] = arvore(lista[:m])
        dir_[raiz] = arvore(lista[m + 1:])
        return raiz

    for i, no in enumerate(nos):
        if no["tipo"] in (1, 5):
            filho[i] = arvore(no["filhos"])

    # mini fluxo (fluxos < 4096 bytes) e setores normais
    mini = bytearray()
    minifat = []
    grandes = []
    inicio = [END] * len(nos)
    for i, no in enumerate(nos):
        if no["tipo"] != 2:
            continue
        d = no["dados"]
        if len(d) == 0:
            inicio[i] = END
        elif len(d) < 4096:
            n = -(-len(d) // MINI)
            inicio[i] = len(mini) // MINI
            minifat += [inicio[i] + k + 1 for k in range(n - 1)] + [END]
            mini += d.ljust(n * MINI, b"\0")
        else:
            grandes.append(i)

    n_dir = -(-len(nos) // 4)
    n_minifat = -(-len(minifat) // 128) if minifat else 0
    n_mini = -(-len(mini) // SETOR) if mini else 0
    n_grandes = [-(-len(nos[i]["dados"]) // SETOR) for i in grandes]
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + sum(n_grandes)
        if n_fat * 128 >= total:
            break
        n_fat += 1
    fat = [FATSECT] * n_fat
    s = n_fat

    def cadeia(n):
        nonlocal s
        ini = s
        fat.extend([s + k + 1 for k in range(n - 1)] + [END])
        s += n
        return ini

    dir_ini = cadeia(n_dir)
    minifat_ini = cadeia(n_minifat) if n_minifat else END
    mini_ini = cadeia(n_mini) if n_mini else END
    for i, n in zip(grandes, n_grandes):
        inicio[i] = cadeia(n)
    fat += [FREE] * (n_fat * 128 - len(fat))

    def entrada(i):
        no = nos[i]
        nome = (no["nome"] + "\0").encode("utf-16-le")
        if no["tipo"] == 5:
            ini, tam = mini_ini, len(mini)
        elif no["tipo"] == 2:
            ini, tam = inicio[i], len(no["dados"])
        else:
            ini, tam = 0, 0
        return (nome.ljust(64, b"\0") + struct.pack("<HBB", len(nome), no["tipo"], 1)
                + struct.pack("<3I", esq[i], dir_[i], filho[i]) + bytes(16) + bytes(4) + bytes(16)
                + struct.pack("<IIi", ini, tam, 0))

    vazia = b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<3I", FREE, FREE, FREE) + bytes(48)
    diretorio = b"".join(entrada(i) for i in range(len(nos)))
    diretorio += vazia * (n_dir * 4 - len(nos))
    difat = list(range(n_fat)) + [FREE] * (109 - n_fat)
    cab = (bytes.fromhex("D0CF11E0A1B11AE1") + bytes(16) + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
           + struct.pack("<IIIIIIIII", 0, n_fat, dir_ini, 0, 4096, minifat_ini, n_minifat, END, 0)
           + struct.pack("<109I", *difat))
    corpo = struct.pack(f"<{len(fat)}I", *fat) + diretorio
    if n_minifat:
        mf = minifat + [FREE] * (n_minifat * 128 - len(minifat))
        corpo += struct.pack(f"<{len(mf)}I", *mf)
    if n_mini:
        corpo += bytes(mini).ljust(n_mini * SETOR, b"\0")
    for i, n in zip(grandes, n_grandes):
        corpo += nos[i]["dados"].ljust(n * SETOR, b"\0")
    return cab + corpo


def vba_project_bin(modulos: dict, documentos: dict) -> bytes:
    id_projeto = "{" + str(uuid.uuid4()).upper() + "}"
    ordem = [(n, True) for n in documentos] + [(n, False) for n in modulos]
    entradas = [("PROJECT", fluxo_project(id_projeto, ordem)), ("PROJECTwm", fluxo_projectwm(ordem)),
                ("VBA/_VBA_PROJECT", bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
                ("VBA/dir", comprimir(fluxo_dir(ordem)))]
    for n, codigo in documentos.items():
        entradas.append((f"VBA/{n}", comprimir(fonte_documento(n, codigo).encode(CP))))
    for n, codigo in modulos.items():
        entradas.append((f"VBA/{n}", comprimir(fonte_modulo(n, codigo).encode(CP))))
    return _cfb(entradas)


def gerar_xlsm(xlsx: str, xlsm: str, modulos: dict, documentos: dict):
    """Copia o .xlsx para .xlsm acrescentando o projeto VBA e os tipos de conteúdo de pasta com macros."""
    binario = vba_project_bin(modulos, documentos)
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                t = dados.decode("utf-8")
                t = t.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                              "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                if 'Extension="bin"' not in t:
                    t = t.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
                dados = t.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                t = dados.decode("utf-8")
                t = t.replace("</Relationships>",
                              '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                              'Target="vbaProject.bin"/></Relationships>')
                dados = t.encode("utf-8")
            zout.writestr(item, dados)
        zout.writestr("xl/vbaProject.bin", binario)
    return binario


if __name__ == "__main__":
    # autoteste da compressão e da criptografia
    for amostra in (b"", b"abc", b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa", os.urandom(8192),
                    ("Attribute VB_Name = \"X\"\r\nOption Explicit\r\n" * 300).encode()):
        assert descomprimir(comprimir(amostra)) == amostra, amostra[:20]
    pid = "{12345678-1234-1234-1234-123456789ABC}"
    for d in (bytes([0]), bytes([0xFF]), struct.pack("<I", 0)):
        for sem in range(256):
            assert decifrar(pid, cifrar(pid, d, sem)) == d
    print("compressão e criptografia: ok")
