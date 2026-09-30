"""Gera um vbaProject.bin (MS-OVBA) a partir de arquivos-fonte VBA.

O projeto gerado e "version-independent" (_VBA_PROJECT sem cache de p-code),
entao o Excel recompila o codigo-fonte na primeira abertura.
Referencias: [MS-CFB] (Compound File Binary) e [MS-OVBA].
"""
import math
import random
import struct
import uuid

# --------------------------------------------------------------------------
# Compressao MS-OVBA 2.4.1
# --------------------------------------------------------------------------

def _copy_token_help(difference):
    bit_count = max(4, math.ceil(math.log2(difference))) if difference > 1 else 4
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, maximum_length


def _compress_chunk(chunk):
    body = bytearray()
    dpos = 0
    n = len(chunk)
    index = {}
    while dpos < n:
        flag_pos = len(body)
        body.append(0)
        flag = 0
        for bit in range(8):
            if dpos >= n:
                break
            best_len, best_off = 0, 0
            if dpos > 0:
                bit_count, max_len = _copy_token_help(dpos)
                max_off = 1 << bit_count
                key = bytes(chunk[dpos:dpos + 3])
                for cand in reversed(index.get(key, [])):
                    off = dpos - cand
                    if off > max_off:
                        break
                    ln = 0
                    limit = min(max_len, n - dpos)
                    while ln < limit and chunk[cand + ln] == chunk[dpos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, off
                        if ln == limit:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_help(dpos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                body += struct.pack('<H', token)
                flag |= 1 << bit
                step = best_len
            else:
                body.append(chunk[dpos])
                step = 1
            for i in range(dpos, dpos + step):
                if i + 3 <= n:
                    index.setdefault(bytes(chunk[i:i + 3]), []).append(i)
            dpos += step
        body[flag_pos] = flag
    if len(body) > 4094 and n == 4096:
        header = (4096 + 2 - 3) | (0b011 << 12)  # bloco sem compressao
        return struct.pack('<H', header & 0x7FFF) + bytes(chunk)
    size = len(body) + 2
    header = ((size - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15)
    return struct.pack('<H', header) + bytes(body)


def compress(data: bytes) -> bytes:
    out = bytearray([1])
    for pos in range(0, len(data), 4096):
        out += _compress_chunk(data[pos:pos + 4096])
    return bytes(out)


# --------------------------------------------------------------------------
# Criptografia de dados do stream PROJECT (MS-OVBA 2.4.3.2)
# --------------------------------------------------------------------------

def _encrypt(project_id: str, data: bytes, rnd: random.Random) -> str:
    seed = rnd.randrange(256)
    version = 2
    proj_key = sum(project_id.encode('ascii')) & 0xFF
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    unenc1 = proj_key
    enc1 = proj_key_enc
    enc2 = version_enc
    out = bytearray([seed, version_enc, proj_key_enc])
    ignored_length = (seed & 6) // 2
    for _ in range(ignored_length):
        temp = rnd.randrange(256)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack('<I', len(data)) + data:
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()


def _decrypt(hexdata: str) -> bytes:
    """Usado apenas para autoverificacao."""
    d = bytes.fromhex(hexdata)
    seed, version_enc, proj_key_enc = d[0], d[1], d[2]
    proj_key = seed ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    pos = 3
    for _ in range((seed & 6) // 2):
        b = d[pos]; pos += 1
        temp = b ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, b, temp
    raw = bytearray()
    for b in d[pos:]:
        byte = b ^ ((enc2 + unenc1) & 0xFF)
        raw.append(byte)
        enc2, enc1, unenc1 = enc1, b, byte
    length = struct.unpack('<I', raw[:4])[0]
    return bytes(raw[4:4 + length])


# --------------------------------------------------------------------------
# Stream "dir" (MS-OVBA 2.3.4.2)
# --------------------------------------------------------------------------

CODEPAGE = 1252


def _rec(rid, payload: bytes) -> bytes:
    return struct.pack('<HI', rid, len(payload)) + payload


def _dir_stream(project_name, modules):
    enc = lambda s: s.encode('cp1252')
    uni = lambda s: s.encode('utf-16-le')
    d = bytearray()
    d += _rec(0x0001, struct.pack('<I', 1))            # SYSKIND Win32
    d += _rec(0x0002, struct.pack('<I', 0x0409))       # LCID
    d += _rec(0x0014, struct.pack('<I', 0x0409))       # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', CODEPAGE))     # CODEPAGE
    d += _rec(0x0004, enc(project_name))               # NAME
    d += _rec(0x0005, b'') + _rec(0x0040, b'')         # DOCSTRING
    d += _rec(0x0006, b'') + _rec(0x003D, b'')         # HELPFILEPATH
    d += _rec(0x0007, struct.pack('<I', 0))            # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))            # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)         # VERSION
    d += _rec(0x000C, b'') + _rec(0x003C, b'')         # CONSTANTS
    # Referencia: OLE Automation (stdole)
    name = 'stdole'
    d += _rec(0x0016, enc(name)) + _rec(0x003E, uni(name))
    libid = enc('*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#'
                'C:\\Windows\\System32\\stdole2.tlb#OLE Automation')
    d += struct.pack('<HII', 0x000D, 4 + len(libid) + 6, len(libid)) + libid + struct.pack('<IH', 0, 0)
    # Modulos
    d += struct.pack('<HIH', 0x000F, 2, len(modules))
    d += struct.pack('<HIH', 0x0013, 2, 0xFFFF)
    for m in modules:
        d += _rec(0x0019, enc(m['name']))
        d += _rec(0x0047, uni(m['name']))
        d += _rec(0x001A, enc(m['name'])) + _rec(0x0032, uni(m['name']))
        d += _rec(0x001C, b'') + _rec(0x0048, b'')
        d += _rec(0x0031, struct.pack('<I', 0))        # MODULEOFFSET
        d += _rec(0x001E, struct.pack('<I', 0))        # HELPCONTEXT
        d += struct.pack('<HIH', 0x002C, 2, 0xFFFF)    # COOKIE
        d += struct.pack('<HI', 0x0022 if m['document'] else 0x0021, 0)
        d += struct.pack('<HI', 0x002B, 0)             # terminador do modulo
    d += struct.pack('<HI', 0x0010, 0)                 # terminador
    return bytes(d)


# --------------------------------------------------------------------------
# Compound File Binary (MS-CFB), versao 3
# --------------------------------------------------------------------------

SECTOR = 512
MINI = 64
CUTOFF = 4096
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


class _Entry:
    def __init__(self, name, kind, data=b''):
        self.name, self.kind, self.data = name, kind, data  # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.id = None


def _cfb_key(e):
    return (len(e.name), e.name.upper())


def _build_tree(entries):
    entries = sorted(entries, key=_cfb_key)
    if not entries:
        return NOSTREAM
    mid = len(entries) // 2
    node = entries[mid]
    node.left = _build_tree(entries[:mid])
    node.right = _build_tree(entries[mid + 1:])
    return node.id


def _write_cfb(root: _Entry) -> bytes:
    # numera entradas (pre-ordem)
    order = []

    def walk(e):
        e.id = len(order)
        order.append(e)
        for c in e.children:
            walk(c)
    walk(root)
    for e in order:
        if e.kind in (1, 5):
            e.child = _build_tree(e.children)

    streams = [e for e in order if e.kind == 2]
    small = [e for e in streams if len(e.data) < CUTOFF]
    big = [e for e in streams if len(e.data) >= CUTOFF]

    # mini stream
    ministream = bytearray()
    minifat = []
    for e in small:
        if not e.data:
            e.start = ENDOFCHAIN
            continue
        nsec = -(-len(e.data) // MINI)
        e.start = len(ministream) // MINI
        for i in range(nsec):
            minifat.append(e.start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        ministream += e.data + b'\0' * (nsec * MINI - len(e.data))
    root.data = bytes(ministream)

    def nsectors(n):
        return -(-n // SECTOR)

    dir_bytes_len = len(order) * 128
    n_dir = nsectors(dir_bytes_len)
    n_minifat = nsectors(len(minifat) * 4)
    n_ministream = nsectors(len(ministream))
    n_big = sum(nsectors(len(e.data)) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_ministream + n_big
        if n_fat * (SECTOR // 4) >= total:
            break
        n_fat += 1
    if n_fat > 109:
        raise ValueError('arquivo grande demais para este gerador')

    fat = []
    sec = 0

    def alloc(count):
        nonlocal sec
        start = sec
        for i in range(count):
            fat.append(start + i + 1 if i < count - 1 else ENDOFCHAIN)
        sec += count
        return start if count else ENDOFCHAIN

    fat_start = sec
    for _ in range(n_fat):
        fat.append(FATSECT)
    sec += n_fat
    dir_start = alloc(n_dir)
    minifat_start = alloc(n_minifat) if n_minifat else ENDOFCHAIN
    root.start = alloc(n_ministream) if n_ministream else ENDOFCHAIN
    for e in big:
        e.start = alloc(nsectors(len(e.data)))
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))

    header = bytearray()
    header += bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 16
    header += struct.pack('<HHHHH', 0x003E, 0x0003, 0xFFFE, 9, 6) + b'\0' * 6
    header += struct.pack('<IIIIIIIII', 0, n_fat, dir_start, 0, CUTOFF,
                          minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += struct.pack('<109I', *difat)
    assert len(header) == 512

    def dir_entry(e):
        name = e.name.encode('utf-16-le') + b'\0\0'
        size = len(e.data) if e.kind in (2, 5) else 0
        start = e.start if e.kind in (2, 5) else 0
        if e.kind == 2 and size == 0:
            start = ENDOFCHAIN
        return (name.ljust(64, b'\0') + struct.pack('<HBB', len(name), e.kind, 1)
                + struct.pack('<III', e.left, e.right, e.child) + b'\0' * 16
                + struct.pack('<I', 0) + b'\0' * 16 + struct.pack('<IQ', start, size))

    dirdata = b''.join(dir_entry(e) for e in order)
    empty = (b'\0' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
             + b'\0' * 36 + struct.pack('<IQ', 0, 0))
    while len(dirdata) % SECTOR:
        dirdata += empty

    def pad(b):
        return b + b'\0' * (-len(b) % SECTOR)

    body = bytearray()
    body += struct.pack('<%dI' % len(fat), *fat)
    body += dirdata
    if n_minifat:
        mf = minifat + [FREESECT] * (n_minifat * (SECTOR // 4) - len(minifat))
        body += struct.pack('<%dI' % len(mf), *mf)
    body += pad(bytes(ministream))
    for e in big:
        body += pad(e.data)
    return bytes(header) + bytes(body)


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------

DOC_ATTRS = {
    'workbook': '0{00020819-0000-0000-C000-000000000046}',
    'sheet': '0{00020820-0000-0000-C000-000000000046}',
}


def _module_source(name, code, doc_kind=None):
    lines = ['Attribute VB_Name = "%s"' % name]
    if doc_kind:
        lines += ['Attribute VB_Base = "%s"' % DOC_ATTRS[doc_kind],
                  'Attribute VB_GlobalNameSpace = False',
                  'Attribute VB_Creatable = False',
                  'Attribute VB_PredeclaredId = True',
                  'Attribute VB_Exposed = True',
                  'Attribute VB_TemplateDerived = False',
                  'Attribute VB_Customizable = True']
    text = '\r\n'.join(lines) + '\r\n' + code.replace('\r\n', '\n').replace('\n', '\r\n')
    if not text.endswith('\r\n'):
        text += '\r\n'
    return text.encode('cp1252')


def build_vba_project(modules, project_name='VBAProject', seed=1234):
    """modules: lista de dicts {name, code, kind} com kind em
    'workbook' | 'sheet' | 'module'. Retorna os bytes do vbaProject.bin."""
    rnd = random.Random(seed)
    project_id = '{' + str(uuid.UUID(int=rnd.getrandbits(128))).upper() + '}'
    mods = []
    for m in modules:
        doc_kind = m['kind'] if m['kind'] in DOC_ATTRS else None
        mods.append({'name': m['name'], 'document': bool(doc_kind),
                     'source': _module_source(m['name'], m['code'], doc_kind)})

    dir_data = compress(_dir_stream(project_name, mods))
    vba_project = bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])

    lines = ['ID="%s"' % project_id]
    for m in mods:
        lines.append(('Document=%s/&H00000000' % m['name']) if m['document'] else ('Module=%s' % m['name']))
    lines += ['Name="%s"' % project_name, 'HelpContextID="0"', 'VersionCompatible32="393222000"']
    cmg = _encrypt(project_id, struct.pack('<I', 0), rnd)
    dpb = _encrypt(project_id, b'\0', rnd)
    gc = _encrypt(project_id, b'\xFF', rnd)
    assert _decrypt(cmg) == struct.pack('<I', 0) and _decrypt(dpb) == b'\0' and _decrypt(gc) == b'\xFF'
    lines += ['CMG="%s"' % cmg, 'DPB="%s"' % dpb, 'GC="%s"' % gc, '',
              '[Host Extender Info]',
              '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
              '[Workspace]']
    for m in mods:
        lines.append('%s=0, 0, 0, 0, C' % m['name'])
    project_stream = ('\r\n'.join(lines) + '\r\n').encode('cp1252')

    wm = bytearray()
    for m in mods:
        wm += m['name'].encode('cp1252') + b'\0' + m['name'].encode('utf-16-le') + b'\0\0'
    wm += b'\0\0'

    root = _Entry('Root Entry', 5)
    vba = _Entry('VBA', 1)
    vba.children = [_Entry('_VBA_PROJECT', 2, vba_project), _Entry('dir', 2, dir_data)]
    for m in mods:
        vba.children.append(_Entry(m['name'], 2, compress(m['source'])))
    root.children = [vba, _Entry('PROJECT', 2, project_stream), _Entry('PROJECTwm', 2, bytes(wm))]
    return _write_cfb(root)
