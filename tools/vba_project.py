"""vbaProject.bin schrijven vanuit VBA-broncode (MS-OVBA + MS-CFB).

Het project wordt zonder gecompileerde p-code geschreven (zoals de specificatie voorschrijft voor
schrijvers: _VBA_PROJECT versie 0xFFFF, geen PerformanceCache). Excel compileert de broncode bij het
openen. Referenties, projectnaam en beveiligingsvelden (CMG/DPB/GC) worden overgenomen uit een
bestaand vbaProject.bin.
"""
from __future__ import annotations

import struct
import olefile

# ----------------------------------------------------------------------------------------------
#  MS-OVBA 2.4.1 compressie
# ----------------------------------------------------------------------------------------------


def _copy_token_help(diff: int):
    bit_count = 4
    while (1 << bit_count) < diff:
        bit_count += 1
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, maximum_length


def _compress_chunk(data: bytes, start: int, end: int) -> bytes:
    out = bytearray()
    p = start
    index: dict[bytes, list[int]] = {}

    def add(pos: int):
        if pos + 3 <= end:
            index.setdefault(data[pos:pos + 3], []).append(pos)

    while p < end:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if p >= end:
                break
            diff = p - start
            best_len, best_off = 0, 0
            if diff > 0 and p + 3 <= end:
                bit_count, max_len = _copy_token_help(diff)
                for c in reversed(index.get(data[p:p + 3], [])):
                    ln = 3
                    lim = min(max_len, end - p)
                    while ln < lim and data[c + ln] == data[p + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, p - c
                        if ln == lim:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_help(diff)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack('<H', token)
                flags |= 1 << bit
                for q in range(p, p + best_len):
                    add(q)
                p += best_len
            else:
                out.append(data[p])
                add(p)
                p += 1
        out[flag_pos] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b'\x01')
    pos = 0
    if not data:
        return bytes(out)
    while pos < len(data):
        end = min(pos + 4096, len(data))
        comp = _compress_chunk(data, pos, end)
        if len(comp) > 4096 and end - pos == 4096:
            header = (4096 + 2 - 3) | (0b011 << 12)          # ongecomprimeerd blok
            out += struct.pack('<H', header) + data[pos:end]
        else:
            if len(comp) > 4096:
                raise ValueError('blok niet comprimeerbaar')
            header = (len(comp) + 2 - 3) | (0b011 << 12) | (1 << 15)
            out += struct.pack('<H', header) + comp
        pos = end
    return bytes(out)


def decompress(data: bytes) -> bytes:
    from oletools.olevba import decompress_stream
    return decompress_stream(bytearray(data))


# ----------------------------------------------------------------------------------------------
#  dir-stream
# ----------------------------------------------------------------------------------------------


def _parse_dir(raw: bytes):
    recs = []
    i = 0
    while i < len(raw):
        rid, size = struct.unpack_from('<HI', raw, i)
        i += 6
        if rid == 0x0009:          # PROJECTVERSION: Reserved=4, daarna 4 + 2 bytes
            size = 6
        recs.append((rid, raw[i:i + size]))
        i += size
        if rid == 0x0010:
            break
    return recs


def _rec(rid: int, data: bytes) -> bytes:
    return struct.pack('<HI', rid, len(data)) + data


def _build_dir(orig_recs, modules, codepage='cp1252') -> bytes:
    out = bytearray()
    for rid, data in orig_recs:
        if rid == 0x000F:
            break
        if rid == 0x0009:
            out += struct.pack('<HI', rid, 4) + data
        else:
            out += _rec(rid, data)
    out += _rec(0x000F, struct.pack('<H', len(modules)))
    out += _rec(0x0013, struct.pack('<H', 0xFFFF))
    for m in modules:
        name = m['name']
        nb, nu = name.encode(codepage), name.encode('utf-16-le')
        out += _rec(0x0019, nb)
        out += _rec(0x0047, nu)
        out += _rec(0x001A, nb)
        out += _rec(0x0032, nu)
        out += _rec(0x001C, b'')
        out += _rec(0x0048, b'')
        out += _rec(0x0031, struct.pack('<I', 0))
        out += _rec(0x001E, struct.pack('<I', 0))
        out += _rec(0x002C, struct.pack('<H', 0xFFFF))
        out += _rec(0x0022 if m['document'] else 0x0021, b'')
        out += _rec(0x002B, b'')
    out += _rec(0x0010, b'')
    return bytes(out)


# ----------------------------------------------------------------------------------------------
#  MS-CFB schrijven (versie 3, sectoren van 512 bytes)
# ----------------------------------------------------------------------------------------------

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
SECT, MINI, CUTOFF = 512, 64, 4096


class _Node:
    def __init__(self, name, kind, data=b''):
        self.name, self.kind, self.data = name, kind, data      # kind: 1 storage, 2 stream, 5 root
        self.children: list[_Node] = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size, self.sid = ENDOFCHAIN, 0, 0


def _sort_key(n: _Node):
    return (len(n.name), n.name.upper())


def _tree(nodes):
    if not nodes:
        return NOSTREAM
    mid = len(nodes) // 2
    n = nodes[mid]
    n.left = _tree(nodes[:mid])
    n.right = _tree(nodes[mid + 1:])
    return n.sid


def write_cfb(streams: dict[str, bytes]) -> bytes:
    """streams: {'PROJECT': b'..', 'VBA/dir': b'..'}"""
    root = _Node('Root Entry', 5)
    storages = {'': root}
    for path in sorted(streams):
        parts = path.split('/')
        parent = root
        for k in range(len(parts) - 1):
            key = '/'.join(parts[:k + 1])
            if key not in storages:
                st = _Node(parts[k], 1)
                parent.children.append(st)
                storages[key] = st
            parent = storages[key]
        parent.children.append(_Node(parts[-1], 2, streams[path]))

    order: list[_Node] = []

    def walk(n):
        n.sid = len(order)
        order.append(n)
        for c in sorted(n.children, key=_sort_key):
            walk(c)
    walk(root)
    for n in order:
        if n.kind in (1, 5):
            n.child = _tree(sorted(n.children, key=_sort_key))

    # mini stream
    mini = bytearray()
    minifat: list[int] = []
    big = []
    for n in order:
        if n.kind != 2:
            continue
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
        elif n.size < CUTOFF:
            first = len(minifat)
            cnt = (n.size + MINI - 1) // MINI
            n.start = first
            for k in range(cnt):
                minifat.append(first + k + 1 if k < cnt - 1 else ENDOFCHAIN)
            mini += n.data + b'\0' * (cnt * MINI - n.size)
        else:
            big.append(n)

    def nsect(nbytes):
        return (nbytes + SECT - 1) // SECT

    n_dir = nsect(len(order) * 128)
    n_minifat = nsect(len(minifat) * 4)
    n_mini = nsect(len(mini))
    n_big = sum(nsect(len(n.data)) for n in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        need = (total + 127) // 128
        if need <= n_fat:
            break
        n_fat = need
    if n_fat > 109:
        raise ValueError('te groot voor een CFB zonder DIFAT')

    fat = [FREESECT] * (n_fat * 128)
    cur = 0

    def chain(count):
        nonlocal cur
        first = cur
        for k in range(count):
            fat[cur] = cur + 1 if k < count - 1 else ENDOFCHAIN
            cur += 1
        return first if count else ENDOFCHAIN

    for _ in range(n_fat):
        fat[cur] = FATSECT
        cur += 1
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    mini_start = chain(n_mini)
    for n in big:
        n.start = chain(nsect(len(n.data)))
    root.start = mini_start if n_mini else ENDOFCHAIN
    root.size = len(mini)

    # directory
    dbytes = bytearray()
    for n in order:
        nm = n.name.encode('utf-16-le')
        e = nm + b'\0' * (64 - len(nm))
        e += struct.pack('<HBB', len(nm) + 2, n.kind, 1)
        e += struct.pack('<III', n.left, n.right, n.child)
        e += b'\0' * 16 + struct.pack('<I', 0) + b'\0' * 16
        e += struct.pack('<IQ', n.start, n.size)
        dbytes += e
    while len(dbytes) % SECT:
        dbytes += b'\0' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\0' * 36 + struct.pack('<IQ', 0, 0)

    header = bytearray(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\0' * 16)
    header += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\0' * 6
    header += struct.pack('<IIIIIIIII', 0, n_fat, dir_start, 0, CUTOFF,
                          minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += struct.pack('<109I', *difat)
    assert len(header) == 512

    body = bytearray()
    body += struct.pack('<%dI' % len(fat), *fat)
    body += dbytes
    mf = minifat + [FREESECT] * (n_minifat * 128 - len(minifat))
    body += struct.pack('<%dI' % len(mf), *mf)
    body += mini + b'\0' * (n_mini * SECT - len(mini))
    for n in big:
        body += n.data + b'\0' * (nsect(len(n.data)) * SECT - len(n.data))
    return bytes(header + body)


# ----------------------------------------------------------------------------------------------
#  project
# ----------------------------------------------------------------------------------------------


def build_vba_project(template_bin: str, modules: list[dict], codepage='cp1252') -> bytes:
    """modules: [{'name': 'ThisWorkbook', 'document': True, 'code': bytes (met Attribute-regels, CRLF)}]"""
    ole = olefile.OleFileIO(template_bin)
    orig_dir = _parse_dir(decompress(ole.openstream('VBA/dir').read()))
    project_txt = ole.openstream('PROJECT').read().decode(codepage)
    ole.close()

    keep = []
    for line in project_txt.split('\r\n'):
        if line.startswith(('ID=', 'Name=', 'HelpContextID=', 'VersionCompatible32=', 'CMG=', 'DPB=', 'GC=')):
            keep.append(line)
    pid = [l for l in keep if l.startswith('ID=')]
    rest = [l for l in keep if not l.startswith('ID=')]
    lines = pid[:]
    for m in modules:
        if m['document']:
            lines.append('Document=%s/&H00000000' % m['name'])
        else:
            lines.append('Module=%s' % m['name'])
    lines += rest
    lines += ['', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
              '', '[Workspace]']
    lines += ['%s=0, 0, 0, 0, C' % m['name'] for m in modules]
    project = ('\r\n'.join(lines) + '\r\n').encode(codepage)

    wm = bytearray()
    for m in modules:
        wm += m['name'].encode(codepage) + b'\0' + m['name'].encode('utf-16-le') + b'\0\0'
    wm += b'\0\0'

    streams = {
        'PROJECT': project,
        'PROJECTwm': bytes(wm),
        'VBA/_VBA_PROJECT': b'\xCC\x61\xFF\xFF\x00\x00\x00',
        'VBA/dir': compress(_build_dir(orig_dir, modules, codepage)),
    }
    for m in modules:
        streams['VBA/' + m['name']] = compress(m['code'])
    return write_cfb(streams)


def document_module_source(name: str, workbook=False) -> bytes:
    base = '0{00020819-0000-0000-C000-000000000046}' if workbook else '0{00020820-0000-0000-C000-000000000046}'
    return ('Attribute VB_Name = "%s"\r\nAttribute VB_Base = "%s"\r\nAttribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n' % (name, base)).encode('cp1252')
