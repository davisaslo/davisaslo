"""Construye un vbaProject.bin (MS-OVBA / MS-CFB) a partir de código fuente VBA.

Reutiliza la información de proyecto y referencias (stdole, Office) del archivo
original del usuario y genera módulos sólo-fuente (sin p-code), de modo que Excel
compila el código al abrir el libro.
"""
import struct
import uuid

CP = 'cp1252'


# ---------------------------------------------------------------- MS-OVBA compression
def _compress_chunk(data):
    out = bytearray()
    p = 0
    n = len(data)
    while p < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if p >= n:
                break
            best_len = 0
            best_off = 0
            if p > 0:
                diff = p
                bitcount = max((diff - 1).bit_length(), 4)
                max_len = (0xFFFF >> bitcount) + 3
                lim = min(max_len, n - p)
                start = max(0, p - (1 << bitcount)) if bitcount < 16 else 0
                # búsqueda codiciosa del match más largo
                for c in range(p - 1, start - 1, -1):
                    if data[c] != data[p]:
                        continue
                    ln = 0
                    while ln < lim and data[c + ln] == data[p + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, p - c
                        if ln == lim:
                            break
            if best_len >= 3:
                bitcount = max((p - 1).bit_length(), 4)
                token = ((best_off - 1) << (16 - bitcount)) | (best_len - 3)
                out += struct.pack('<H', token)
                flags |= 1 << bit
                p += best_len
            else:
                out.append(data[p])
                p += 1
        out[flag_pos] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    res = bytearray(b'\x01')
    for i in range(0, len(data), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) + 2 > 4098 and len(chunk) == 4096:
            hdr = 0x3000 | (4098 - 3)  # sin compresión
            res += struct.pack('<H', hdr) + chunk
        else:
            size = len(comp) + 2
            hdr = 0x8000 | 0x3000 | (size - 3)
            res += struct.pack('<H', hdr) + comp
    return bytes(res)


# ---------------------------------------------------------------- MS-CFB writer
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


class Node:
    def __init__(self, name, data=None, children=None):
        self.name = name
        self.data = data
        self.children = children  # None -> stream
        self.sid = None
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0


def _key(name):
    return (len(name), name.upper())


def _build_tree(children):
    s = sorted(children, key=lambda c: _key(c.name))

    def rec(lst):
        if not lst:
            return NOSTREAM
        mid = len(lst) // 2
        node = lst[mid]
        node.left = rec(lst[:mid])
        node.right = rec(lst[mid + 1:])
        return node.sid
    return rec(s)


def write_cfb(root_children):
    root = Node('Root Entry', children=root_children)
    nodes = []

    def walk(n):
        n.sid = len(nodes)
        nodes.append(n)
        if n.children:
            for c in n.children:
                walk(c)
    walk(root)
    for n in nodes:
        if n.children is not None:
            n.child = _build_tree(n.children)

    # ministream
    mini = bytearray()
    minifat = []
    big = []
    for n in nodes:
        if n.children is None:
            n.size = len(n.data)
            if n.size < 4096:
                if n.size == 0:
                    n.start = ENDOFCHAIN
                    continue
                first = len(mini) // 64
                cnt = (n.size + 63) // 64
                n.start = first
                for k in range(cnt):
                    minifat.append(first + k + 1 if k < cnt - 1 else ENDOFCHAIN)
                mini += n.data + b'\x00' * (cnt * 64 - n.size)
            else:
                big.append(n)

    dir_bytes_len = len(nodes) * 128
    n_dir = (dir_bytes_len + 511) // 512
    n_minifat = (len(minifat) * 4 + 511) // 512 if minifat else 0
    n_mini = (len(mini) + 511) // 512
    n_big = sum((b.size + 511) // 512 for b in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        if (total + 127) // 128 <= n_fat:
            break
        n_fat += 1
    assert n_fat <= 109
    fat = [FREESECT] * (n_fat * 128)
    sec = 0
    for _ in range(n_fat):
        fat[sec] = FATSECT
        sec += 1

    def chain(count):
        nonlocal sec
        first = sec
        for k in range(count):
            fat[sec] = sec + 1 if k < count - 1 else ENDOFCHAIN
            sec += 1
        return first if count else ENDOFCHAIN
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    mini_start = chain(n_mini)
    for b in big:
        b.start = chain((b.size + 511) // 512)
    root.start = mini_start if n_mini else ENDOFCHAIN
    root.size = len(mini)

    # directory entries
    d = bytearray()
    for n in nodes:
        nm = n.name.encode('utf-16-le') + b'\x00\x00'
        assert len(nm) <= 64
        typ = 5 if n is root else (1 if n.children is not None else 2)
        d += nm.ljust(64, b'\x00')
        d += struct.pack('<HBB', len(nm), typ, 1)
        d += struct.pack('<III', n.left, n.right, n.child)
        d += b'\x00' * 16 + b'\x00' * 4 + b'\x00' * 16
        d += struct.pack('<IQ', n.start if (typ != 1) else 0, n.size if typ != 1 else 0)
    while len(d) % 512:
        d += (b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
              + b'\x00' * 48)

    hdr = struct.pack('<8s16sHHHHH6sIIIIIIIII', b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1', b'\x00' * 16,
                      0x3E, 3, 0xFFFE, 9, 6, b'\x00' * 6, 0, n_fat, dir_start, 0, 4096,
                      minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    hdr += struct.pack('<109I', *difat)
    assert len(hdr) == 512

    body = bytearray()
    body += struct.pack('<%dI' % len(fat), *fat)
    body += d
    if n_minifat:
        mf = struct.pack('<%dI' % len(minifat), *minifat)
        body += mf + struct.pack('<I', FREESECT) * ((n_minifat * 512 - len(mf)) // 4)
    if n_mini:
        body += mini + b'\x00' * (n_mini * 512 - len(mini))
    for b in big:
        body += b.data + b'\x00' * (((b.size + 511) // 512) * 512 - b.size)
    return bytes(hdr + body)


# ---------------------------------------------------------------- dir stream
def rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload


def build_dir(info_and_refs: bytes, modules):
    out = bytearray(info_and_refs)
    out += struct.pack('<HIH', 0x000F, 2, len(modules))
    out += struct.pack('<HIH', 0x0013, 2, 0xFFFF)
    for m in modules:
        nm = m['name'].encode(CP)
        nmu = m['name'].encode('utf-16-le')
        out += rec(0x0019, nm)
        out += rec(0x0047, nmu)
        out += rec(0x001A, nm)
        out += rec(0x0032, nmu)
        out += rec(0x001C, b'') + rec(0x0048, b'')
        out += struct.pack('<HII', 0x0031, 4, 0)       # offset del código fuente
        out += struct.pack('<HII', 0x001E, 4, 0)       # help context
        out += struct.pack('<HIH', 0x002C, 2, 0xFFFF)  # cookie
        out += rec(0x0022 if m['document'] else 0x0021, b'')
        out += struct.pack('<HI', 0x002B, 0)
    out += struct.pack('<HI', 0x0010, 0)
    return bytes(out)


def build_vba_project(template_bin, modules, out_path):
    """modules: lista de dict(name, code, document(bool), base(str|None))"""
    import olefile
    from oletools.olevba import decompress_stream
    o = olefile.OleFileIO(template_bin)
    olddir = bytes(decompress_stream(bytearray(o.openstream('VBA/dir').read())))
    oldproj = o.openstream('PROJECT').read().decode(CP)
    o.close()
    cut = olddir.index(struct.pack('<HIH', 0x000F, 2, 0)[:6])
    info = olddir[:cut]

    def keep(key):
        for line in oldproj.splitlines():
            if line.startswith(key + '='):
                return line
        raise KeyError(key)

    proj = [keep('ID')]
    for m in modules:
        proj.append(('Document=%s/&H00000000' % m['name']) if m['document'] else ('Module=%s' % m['name']))
    proj += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
             keep('CMG'), keep('DPB'), keep('GC'), '',
             '[Host Extender Info]',
             '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
             '[Workspace]']
    for m in modules:
        proj.append('%s=0, 0, 0, 0, C' % m['name'])
    project_stream = ('\r\n'.join(proj) + '\r\n').encode(CP)

    wm = bytearray()
    for m in modules:
        wm += m['name'].encode(CP) + b'\x00' + m['name'].encode('utf-16-le') + b'\x00\x00'
    wm += b'\x00\x00'

    vba_children = [Node('_VBA_PROJECT', b'\xcc\x61\xff\xff\x00\x00\x00'),
                    Node('dir', compress(build_dir(info, modules)))]
    for m in modules:
        if m['document']:
            attrs = ['Attribute VB_Name = "%s"' % m['name'],
                     'Attribute VB_Base = "%s"' % m['base'],
                     'Attribute VB_GlobalNameSpace = False',
                     'Attribute VB_Creatable = False',
                     'Attribute VB_PredeclaredId = True',
                     'Attribute VB_Exposed = True',
                     'Attribute VB_TemplateDerived = False',
                     'Attribute VB_Customizable = True']
        else:
            attrs = ['Attribute VB_Name = "%s"' % m['name']]
        code = m['code'].replace('\r\n', '\n').strip('\n').split('\n')
        src = '\r\n'.join(attrs + code) + '\r\n'
        vba_children.append(Node(m['name'], compress(src.encode(CP))))

    data = write_cfb([Node('VBA', children=vba_children),
                      Node('PROJECT', project_stream),
                      Node('PROJECTwm', bytes(wm))])
    with open(out_path, 'wb') as f:
        f.write(data)
    return data
