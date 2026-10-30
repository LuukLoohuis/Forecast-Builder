"""Schrijft een vbaProject.bin (VBA-project voor .xlsm) uit platte VBA-broncode.

Volgt [MS-OVBA] (compressie, dir-stream, PROJECT-stream, data-encryptie van CMG/DPB/GC) en [MS-CFB] versie 3.
Er wordt geen p-code geschreven: _VBA_PROJECT bevat Version 0xFFFF, waardoor Office de broncode bij het openen
compileert (dit schrijft de specificatie voor bij het aanmaken van een project).

Gebruik:
    build_vba_project(std_modules={"modRDFO": "Sub X()\\r\\nEnd Sub"},
                      document_modules=["ThisWorkbook", "Blad1", ...]) -> bytes
"""
import math
import random
import struct
import uuid

CP = "cp1252"
NOSTREAM = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
SECTOR = 512
MINI = 64
CUTOFF = 4096


# --------------------------------------------------------------------------- MS-OVBA 2.4.1 compressie
def _compress_chunk(chunk):
    n = len(chunk)
    out = bytearray()
    heads = {}
    i = 0

    def remember(p):
        if p + 3 <= n:
            heads.setdefault(bytes(chunk[p:p + 3]), []).append(p)

    while i < n:
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if i >= n:
                break
            best_len, best_off = 0, 0
            if i > 0 and i + 3 <= n:
                bit_count = max(math.ceil(math.log2(i)), 4) if i > 1 else 4
                max_len = (0xFFFF >> bit_count) + 3
                max_off = 1 << bit_count
                for j in reversed(heads.get(bytes(chunk[i:i + 3]), [])):
                    if i - j > max_off:
                        break
                    ln = 0
                    while i + ln < n and ln < max_len and chunk[j + ln] == chunk[i + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, i - j
                        if ln == max_len:
                            break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= 1 << bit
                for p in range(i, i + best_len):
                    remember(p)
                i += best_len
            else:
                out.append(chunk[i])
                remember(i)
                i += 1
        out[flag_pos] = flag
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for pos in range(0, len(data), 4096):
        chunk = data[pos:pos + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) <= 4096 or len(chunk) < 4096:
            assert len(comp) + 2 <= 4098, "chunk niet comprimeerbaar"
            out += struct.pack("<H", ((len(comp) + 2 - 3) & 0x0FFF) | 0x3000 | 0x8000) + comp
        else:
            out += struct.pack("<H", 0x0FFF | 0x3000) + chunk
    return bytes(out)


# --------------------------------------------------------------------------- MS-OVBA 2.4.3 data-encryptie
def encrypt(data, project_id):
    seed = random.randint(0, 255)
    version = 2
    proj_key = sum(project_id.encode(CP)) % 256
    out = [seed, seed ^ version, seed ^ proj_key]
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version
    ignored_len = (seed & 6) // 2
    for _ in range(ignored_len):
        temp = random.randint(0, 255)
        b = temp ^ ((enc2 + unenc1) % 256)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + bytes(data):
        b = byte ^ ((enc2 + unenc1) % 256)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return "".join(f"{x:02X}" for x in out)


def decrypt(hexstr):
    """Alleen voor de eigen test: draait encrypt() terug."""
    d = bytes.fromhex(hexstr)
    seed, ver_enc, pk_enc = d[0], d[1], d[2]
    assert seed ^ ver_enc == 2
    proj_key = seed ^ pk_enc
    unenc1, enc1, enc2 = proj_key, pk_enc, ver_enc
    pos = 3
    for _ in range((seed & 6) // 2):
        b = d[pos]
        unenc1, enc2, enc1 = b ^ ((enc2 + unenc1) % 256), enc1, b
        pos += 1
    raw = []
    for b in d[pos:]:
        byte = b ^ ((enc2 + unenc1) % 256)
        raw.append(byte)
        enc2, enc1, unenc1 = enc1, b, byte
    ln = struct.unpack("<I", bytes(raw[:4]))[0]
    return bytes(raw[4:4 + ln]), proj_key


# --------------------------------------------------------------------------- dir-stream
def _rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload


def _dir_stream(modules):
    """modules: lijst (naam, is_document)."""
    name = b"VBAProject"
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))            # SYSKIND win32
    d += _rec(0x0002, struct.pack("<I", 0x0409))       # LCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))       # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", 1252))         # CODEPAGE
    d += _rec(0x0004, name)                            # NAME
    d += _rec(0x0005, b"") + struct.pack("<HI", 0x0040, 0)            # DOCSTRING (+unicode)
    d += _rec(0x0006, b"") + struct.pack("<HI", 0x003D, 0)            # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)         # VERSION (Reserved=4, major, minor)
    d += _rec(0x000C, b"") + struct.pack("<HI", 0x003C, 0)            # CONSTANTS
    # geen REFERENCE-records: VBA- en Excel-bibliotheek zijn impliciet; PowerPoint via late binding
    d += _rec(0x000F, struct.pack("<H", len(modules)))  # MODULES count
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))         # PROJECTCOOKIE
    for mname, is_doc in modules:
        b = mname.encode(CP)
        u = mname.encode("utf-16-le")
        d += _rec(0x0019, b)                             # MODULENAME
        d += _rec(0x0047, u)                             # MODULENAMEUNICODE
        d += _rec(0x001A, b) + struct.pack("<HI", 0x0032, len(u)) + u    # STREAMNAME
        d += _rec(0x001C, b"") + struct.pack("<HI", 0x0048, 0)           # DOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))          # OFFSET (geen p-code)
        d += _rec(0x001E, struct.pack("<I", 0))          # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))     # COOKIE
        d += struct.pack("<HI", 0x0022 if is_doc else 0x0021, 0)          # TYPE
        d += struct.pack("<HI", 0x002B, 0)               # module-terminator
    d += struct.pack("<HI", 0x0010, 0)                   # dir-terminator
    return bytes(d)


def _module_source(name, code, doc_base=None):
    lines = [f'Attribute VB_Name = "{name}"']
    if doc_base:
        lines += [f'Attribute VB_Base = "{doc_base}"', "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
                  "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = True", "Attribute VB_TemplateDerived = False",
                  "Attribute VB_Customizable = True"]
    text = "\r\n".join(lines) + "\r\n" + code.replace("\r\n", "\n").replace("\n", "\r\n")
    if not text.endswith("\r\n"):
        text += "\r\n"
    return text.encode(CP)


# --------------------------------------------------------------------------- MS-CFB writer (versie 3)
class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size = ENDOFCHAIN, 0
        self.sid = None


def _cfb_key(e):
    return (len(e.name), e.name.upper())


def _tree(nodes):
    if not nodes:
        return NOSTREAM
    nodes = sorted(nodes, key=_cfb_key)
    mid = len(nodes) // 2
    root = nodes[mid]
    root.left = _tree(nodes[:mid])
    root.right = _tree(nodes[mid + 1:])
    return root.sid


def write_cfb(root_children):
    """root_children: lijst _Entry (storages met .children, streams met .data)."""
    root = _Entry("Root Entry", 5)
    root.children = root_children
    entries = []

    def walk(e):
        e.sid = len(entries)
        entries.append(e)
        for c in e.children:
            walk(c)
    walk(root)
    for e in entries:
        if e.kind in (1, 5):
            e.child = _tree(e.children)

    # mini stream en grote streams
    mini = bytearray()
    minifat = []
    bigs = []
    for e in entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            first = len(mini) // MINI
            nsec = math.ceil(e.size / MINI)
            for k in range(nsec):
                minifat.append(first + k + 1 if k < nsec - 1 else ENDOFCHAIN)
            mini += e.data + b"\x00" * (nsec * MINI - e.size)
            e.start = first
        else:
            bigs.append(e)

    sectors = []      # lijst van (bytes, chain-id) per sector
    fat_chain = []

    def place(blob):
        n = math.ceil(len(blob) / SECTOR)
        first = len(sectors)
        blob = blob + b"\x00" * (n * SECTOR - len(blob))
        for k in range(n):
            sectors.append(blob[k * SECTOR:(k + 1) * SECTOR])
            fat_chain.append(first + k + 1 if k < n - 1 else ENDOFCHAIN)
        return first, n

    for e in bigs:
        e.start, _ = place(e.data)
    if mini:
        root.start, _ = place(bytes(mini))
        root.size = len(mini)
    if minifat:
        mf = b"".join(struct.pack("<I", x) for x in minifat)
        minifat_start, minifat_n = place(mf + struct.pack("<I", FREESECT) * ((-len(minifat)) % 128))
    else:
        minifat_start, minifat_n = ENDOFCHAIN, 0

    dirblob = bytearray()
    for e in entries:
        nm = (e.name.encode("utf-16-le") + b"\x00\x00") if e.kind else b""
        dirblob += nm.ljust(64, b"\x00")
        dirblob += struct.pack("<HBB", len(nm), e.kind, 1)
        dirblob += struct.pack("<III", e.left, e.right, e.child)
        dirblob += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        dirblob += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size)
    while len(dirblob) % SECTOR:
        dirblob += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 48
    dir_start, _ = place(bytes(dirblob))

    # FAT
    nonfat = len(sectors)
    nfat = 1
    while nonfat + nfat > nfat * 128:
        nfat += 1
    assert nfat <= 109
    fat_start = len(sectors)
    fat = fat_chain + [FATSECT] * nfat
    fat += [FREESECT] * (nfat * 128 - len(fat))
    fatblob = b"".join(struct.pack("<I", x) for x in fat)
    for k in range(nfat):
        sectors.append(fatblob[k * SECTOR:(k + 1) * SECTOR])

    hdr = bytearray()
    hdr += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIII", 0, nfat, dir_start, 0)
    hdr += struct.pack("<IIIII", CUTOFF, minifat_start, minifat_n, ENDOFCHAIN, 0)
    difat = [fat_start + k for k in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr += b"".join(struct.pack("<I", x) for x in difat)
    assert len(hdr) == 512
    return bytes(hdr) + b"".join(sectors)


# --------------------------------------------------------------------------- VBA-project
WB_BASE = "0{00020819-0000-0000-C000-000000000046}"
WS_BASE = "0{00020820-0000-0000-C000-000000000046}"


def build_vba_project(std_modules, document_modules, workbook_code="", sheet_code=None):
    """std_modules: {naam: code}; document_modules: ['ThisWorkbook', 'Blad1', ...] (codeNames)."""
    sheet_code = sheet_code or {}
    pid = "{" + str(uuid.uuid4()).upper() + "}"
    mods = []
    for dm in document_modules:
        base = WB_BASE if dm == "ThisWorkbook" else WS_BASE
        code = workbook_code if dm == "ThisWorkbook" else sheet_code.get(dm, "")
        mods.append((dm, True, _module_source(dm, code, base)))
    for nm, code in std_modules.items():
        mods.append((nm, False, _module_source(nm, code)))

    project = [f'ID="{pid}"']
    project += [f"Document={m}/&H00000000" for m, d, _ in mods if d]
    project += [f"Module={m}" for m, d, _ in mods if not d]
    project += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
                f'CMG="{encrypt(struct.pack("<I", 0), pid)}"',
                f'DPB="{encrypt(b"\x00", pid)}"',
                f'GC="{encrypt(b"\xff", pid)}"',
                "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
                "", "[Workspace]"]
    project += [f"{m}=0, 0, 0, 0, C" for m, _d, _s in mods]
    project_stream = ("\r\n".join(project) + "\r\n").encode(CP)
    wm = bytearray()
    for m, _d, _s in mods:
        wm += m.encode(CP) + b"\x00" + m.encode("utf-16-le") + b"\x00\x00"
    wm += b"\x00\x00"

    vba = _Entry("VBA", 1)
    vba.children = [_Entry("_VBA_PROJECT", 2, struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000)),
                    _Entry("dir", 2, compress(_dir_stream([(m, d) for m, d, _ in mods])))]
    vba.children += [_Entry(m, 2, compress(src)) for m, _d, src in mods]
    return write_cfb([vba, _Entry("PROJECT", 2, project_stream), _Entry("PROJECTwm", 2, bytes(wm))])
