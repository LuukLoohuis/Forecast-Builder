"""Bouwt een nieuw xl/vbaProject.bin uit een bestaand vbaProject.bin plus nieuwe VBA-broncode.

Doel: een .xlsm afleveren waarin de bijgewerkte macro al zit, zonder dat de gebruiker een .bas hoeft
te importeren. Er wordt geen p-code of performance cache geschreven ("VBA purging"): alle
MODULEOFFSET-waarden worden 0, _VBA_PROJECT wordt de minimale 7-byte variant en de __SRP_*-streams
vervallen. Excel hercompileert dan uit de bron.

Gebaseerd op MS-OVBA (compressie 2.4.1, dir-stream 2.3.4.2) en MS-CFB (compound file versie 3).

Gebruik:
    python3 -m forecast_builder.vbabuild --template <bin> --module Naam=<pad .bas> --out <bin>
    build_vbaproject(template_bin: bytes, modules: dict[str, str]) -> bytes
"""
from __future__ import annotations

import argparse
import io
import struct
from dataclasses import dataclass, field

# ----------------------------------------------------------------------------------------------
#  MS-OVBA 2.4.1 compressie
# ----------------------------------------------------------------------------------------------

CHUNK = 4096


def _copy_token_bits(difference: int) -> tuple[int, int, int]:
    """Geeft (bit_count, length_mask, maximum_length) voor een copy token (MS-OVBA 2.4.1.3.19.1)."""
    bit_count = max((difference - 1).bit_length(), 4)      # == max(ceil(log2(difference)), 4)
    length_mask = 0xFFFF >> bit_count
    return bit_count, length_mask, length_mask + 3


def decompress(data: bytes) -> bytes:
    """Decomprimeert een CompressedContainer (MS-OVBA 2.4.1.3.1)."""
    if not data:
        raise ValueError("lege compressed container")
    if data[0] != 0x01:
        raise ValueError("compressed container begint niet met signature 0x01")
    out = bytearray()
    pos = 1
    end = len(data)
    while pos < end:
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("ongeldige chunk signature op positie %d" % pos)
        compressed = header & 0x8000
        chunk_start = pos
        chunk_end = min(end, chunk_start + size)
        pos += 2
        dec_start = len(out)
        if not compressed:
            out += data[pos:pos + CHUNK]
            pos = chunk_end
            continue
        while pos < chunk_end and len(out) - dec_start < CHUNK:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end or len(out) - dec_start >= CHUNK:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bit_count, length_mask, _ = _copy_token_bits(len(out) - dec_start)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    src = len(out) - offset
                    if src < dec_start:
                        raise ValueError("copy token wijst vóór het begin van de chunk")
                    for _ in range(length):          # byte voor byte, overlap is toegestaan
                        out.append(out[src])
                        src += 1
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


def _compress_chunk_tokens(data: bytes, start: int, end: int) -> bytes:
    """Codeert data[start:end] als token sequences (zonder chunk header); greedy matching."""
    out = bytearray()
    pos = start
    # hash op 3 bytes -> posities binnen deze chunk (meest recente achteraan)
    table: dict[bytes, list[int]] = {}
    while pos < end:
        flags = 0
        tokens = bytearray()
        for bit in range(8):
            if pos >= end:
                break
            difference = pos - start
            best_len = 0
            best_off = 0
            if difference > 0 and pos + 3 <= end:
                bit_count, _, max_len = _copy_token_bits(difference)
                max_off = min(difference, 1 << bit_count)
                limit = min(max_len, end - pos)
                key = data[pos:pos + 3]
                cands = table.get(key)
                if cands:
                    tried = 0
                    for cand in reversed(cands):
                        if pos - cand > max_off:
                            break
                        tried += 1
                        if tried > 64:
                            break
                        n = 3
                        while n < limit and data[cand + n] == data[pos + n]:
                            n += 1
                        if n > best_len:
                            best_len = n
                            best_off = pos - cand
                            if n == limit:
                                break
            if best_len >= 3:
                bit_count, length_mask, _ = _copy_token_bits(difference)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                tokens += struct.pack("<H", token)
                flags |= 1 << bit
                step = best_len
            else:
                tokens.append(data[pos])
                step = 1
            for p in range(pos, pos + step):
                if p + 3 <= end:
                    table.setdefault(data[p:p + 3], []).append(p)
            pos += step
        out.append(flags)
        out += tokens
    return bytes(out)


def _compress_chunk(data: bytes, start: int, end: int, out: bytearray) -> None:
    body = _compress_chunk_tokens(data, start, end)
    if len(body) <= CHUNK:
        out += struct.pack("<H", 0x8000 | 0x3000 | (len(body) + 2 - 3))
        out += body
    elif end - start == CHUNK:
        # niet comprimeerbaar: ruwe chunk van precies 4096 bytes (MS-OVBA 2.4.1.3.7 / 2.4.1.3.3)
        out += struct.pack("<H", 0x3000 | 0x0FFF)
        out += data[start:end]
    else:
        # niet-comprimeerbare staart korter dan 4096: splitsen in twee kleinere compressed chunks
        # (een compressed chunk van <= 3640 bytes past altijd: 3640 + 455 flagbytes < 4096)
        mid = start + (end - start) // 2
        _compress_chunk(data, start, mid, out)
        _compress_chunk(data, mid, end, out)


def compress(data: bytes) -> bytes:
    """Comprimeert bytes naar een CompressedContainer (MS-OVBA 2.4.1.3.6)."""
    data = bytes(data)
    out = bytearray(b"\x01")
    for start in range(0, len(data), CHUNK):
        _compress_chunk(data, start, min(len(data), start + CHUNK), out)
    return bytes(out)


# ----------------------------------------------------------------------------------------------
#  MS-OVBA 2.3.4.2 dir-stream
# ----------------------------------------------------------------------------------------------

REC_PROJECTCODEPAGE = 0x0003
REC_PROJECTVERSION = 0x0009
REC_PROJECTMODULES = 0x000F
REC_TERMINATOR = 0x0010
REC_MODULENAME = 0x0019
REC_MODULESTREAMNAME = 0x001A
REC_MODULETYPE_PROC = 0x0021
REC_MODULETYPE_DOC = 0x0022
REC_MODULE_END = 0x002B
REC_MODULEOFFSET = 0x0031
REC_MODULESTREAMNAME_U = 0x0032


@dataclass
class DirRecord:
    rid: int
    data: bytes

    def tobytes(self) -> bytes:
        if self.rid == REC_PROJECTVERSION:
            # Id, Reserved (uint32 = 4), VersionMajor (uint32), VersionMinor (uint16): vaste lengte
            return struct.pack("<HI", self.rid, 4) + self.data
        return struct.pack("<HI", self.rid, len(self.data)) + self.data


@dataclass
class DirModule:
    name: str
    stream_name: str
    text_offset: int
    is_document: bool
    records: list[DirRecord] = field(default_factory=list)


def parse_dir(decompressed: bytes) -> list[DirRecord]:
    """Splitst een gedecomprimeerde dir-stream in records (id, data)."""
    records = []
    pos = 0
    n = len(decompressed)
    while pos < n:
        rid, size = struct.unpack_from("<HI", decompressed, pos)
        pos += 6
        if rid == REC_PROJECTVERSION:
            size = 6
        records.append(DirRecord(rid, bytes(decompressed[pos:pos + size])))
        pos += size
    return records


def serialize_dir(records: list[DirRecord]) -> bytes:
    return b"".join(r.tobytes() for r in records)


def dir_modules(records: list[DirRecord], codepage: str) -> list[DirModule]:
    """Groepeert de MODULE-records per module."""
    modules: list[DirModule] = []
    cur: DirModule | None = None
    for r in records:
        if r.rid == REC_MODULENAME:
            cur = DirModule(r.data.decode(codepage), "", 0, False)
            modules.append(cur)
        if cur is None:
            continue
        cur.records.append(r)
        if r.rid == REC_MODULESTREAMNAME:
            cur.stream_name = r.data.decode(codepage)
        elif r.rid == REC_MODULEOFFSET:
            cur.text_offset = struct.unpack("<I", r.data)[0]
        elif r.rid == REC_MODULETYPE_DOC:
            cur.is_document = True
        elif r.rid == REC_MODULE_END:
            cur = None
    return modules


def dir_codepage(records: list[DirRecord]) -> str:
    for r in records:
        if r.rid == REC_PROJECTCODEPAGE:
            return "cp%d" % struct.unpack("<H", r.data)[0]
    return "cp1252"


# ----------------------------------------------------------------------------------------------
#  MS-CFB schrijver (versie 3)
# ----------------------------------------------------------------------------------------------

SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF
TYPE_STORAGE = 1
TYPE_STREAM = 2
TYPE_ROOT = 5
BLACK = 1


@dataclass
class CfbNode:
    name: str
    is_storage: bool
    data: bytes = b""
    children: list["CfbNode"] = field(default_factory=list)
    # ingevuld tijdens het schrijven
    sid: int = -1
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def _name_key(name: str):
    # MS-CFB 2.6.4: eerst op lengte (UTF-16 code units), dan hoofdletterongevoelig per code unit
    units = name.encode("utf-16-le")
    return (len(units), name.upper().encode("utf-16-le"))


def _build_bst(nodes: list[CfbNode]) -> int:
    """Bouwt een gebalanceerde BST uit gesorteerde knopen; geeft de sid van de wortel (of NOSTREAM)."""
    if not nodes:
        return NOSTREAM
    mid = len(nodes) // 2
    root = nodes[mid]
    root.left = _build_bst(nodes[:mid])
    root.right = _build_bst(nodes[mid + 1:])
    return root.sid


def _chain(fat: list[int], first: int, count: int) -> None:
    for i in range(count):
        fat[first + i] = first + i + 1 if i + 1 < count else ENDOFCHAIN


def _dir_entry(node: CfbNode, etype: int) -> bytes:
    name = node.name.encode("utf-16-le") + b"\x00\x00"
    if len(name) > 64:
        raise ValueError("naam te lang: %r" % node.name)
    return (
        name.ljust(64, b"\x00")
        + struct.pack("<HBB", len(name), etype, BLACK)
        + struct.pack("<III", node.left, node.right, node.child)
        + b"\x00" * 16                         # CLSID
        + struct.pack("<I", 0)                 # state bits
        + b"\x00" * 16                         # tijden
        + struct.pack("<I", node.start)
        + struct.pack("<Q", node.size)
    )


def write_cfb(root: CfbNode) -> bytes:
    """Schrijft een compound file (versie 3, sector 512, mini-sector 64) voor de boom onder root."""
    # 1. directory-entries nummeren (root = 0, daarna per storage de kinderen in boomvolgorde)
    entries: list[CfbNode] = []

    def assign(node: CfbNode) -> None:
        node.sid = len(entries)
        entries.append(node)
        node.children.sort(key=lambda c: _name_key(c.name))
        for c in node.children:
            assign(c)
        node.child = _build_bst(node.children)

    assign(root)

    # 2. stream-data verdelen over mini stream en gewone sectoren
    streams = [n for n in entries if not n.is_storage]
    mini_container = bytearray()
    regular: list[CfbNode] = []
    for n in streams:
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
        elif n.size < MINI_CUTOFF:
            n.start = len(mini_container) // MINI_SECTOR
            mini_container += n.data
            mini_container += b"\x00" * (-len(n.data) % MINI_SECTOR)
        else:
            regular.append(n)
    n_mini = len(mini_container) // MINI_SECTOR
    minifat = [FREESECT] * n_mini
    for n in streams:
        if 0 < n.size < MINI_CUTOFF:
            _chain(minifat, n.start, -(-n.size // MINI_SECTOR))
    minifat_bytes = b"".join(struct.pack("<I", v) for v in minifat)
    minifat_bytes += b"\xff" * (-len(minifat_bytes) % SECTOR)
    n_minifat_sect = len(minifat_bytes) // SECTOR

    # directory
    n_dir_sect = max(1, -(-len(entries) // 4))

    def sectors_of(nbytes: int) -> int:
        return -(-nbytes // SECTOR)

    n_minicont_sect = sectors_of(len(mini_container))
    regular_sects = [sectors_of(n.size) for n in regular]
    n_data = n_dir_sect + n_minifat_sect + n_minicont_sect + sum(regular_sects)
    n_fat = 1
    while n_fat * 128 < n_data + n_fat:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("bestand te groot voor een DIFAT zonder extra sectoren")
    total = n_fat + n_data
    fat = [FREESECT] * (n_fat * 128)

    # 3. sectoren toewijzen: FAT, directory, mini-FAT, mini stream, gewone streams
    sect = 0
    fat_sects = list(range(sect, sect + n_fat))
    for s in fat_sects:
        fat[s] = FATSECT
    sect += n_fat
    dir_start = sect
    _chain(fat, dir_start, n_dir_sect)
    sect += n_dir_sect
    minifat_start = ENDOFCHAIN
    if n_minifat_sect:
        minifat_start = sect
        _chain(fat, minifat_start, n_minifat_sect)
        sect += n_minifat_sect
    root.start = ENDOFCHAIN
    root.size = len(mini_container)
    if n_minicont_sect:
        root.start = sect
        _chain(fat, root.start, n_minicont_sect)
        sect += n_minicont_sect
    for n, cnt in zip(regular, regular_sects):
        n.start = sect
        _chain(fat, n.start, cnt)
        sect += cnt
    assert sect == total

    # 4. bytes uitschrijven
    out = io.BytesIO()
    difat = fat_sects + [FREESECT] * (109 - len(fat_sects))
    header = (
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        + b"\x00" * 16
        + struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
        + b"\x00" * 6
        + struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, MINI_CUTOFF,
                      minifat_start, n_minifat_sect, ENDOFCHAIN, 0)
        + b"".join(struct.pack("<I", v) for v in difat)
    )
    assert len(header) == SECTOR
    out.write(header)
    out.write(b"".join(struct.pack("<I", v) for v in fat))

    def padded(b: bytes) -> bytes:
        return b + b"\x00" * (-len(b) % SECTOR)

    dir_bytes = bytearray()
    for n in entries:
        etype = TYPE_ROOT if n is root else (TYPE_STORAGE if n.is_storage else TYPE_STREAM)
        if n.is_storage and n is not root:
            n.start, n.size = 0, 0
        dir_bytes += _dir_entry(n, etype)
    unused = b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 48
    assert len(unused) == 128
    while len(dir_bytes) % SECTOR:
        dir_bytes += unused
    out.write(bytes(dir_bytes))
    out.write(minifat_bytes)
    out.write(padded(bytes(mini_container)))
    for n in regular:
        out.write(padded(n.data))
    result = out.getvalue()
    assert len(result) == SECTOR * (1 + total)
    return result


# ----------------------------------------------------------------------------------------------
#  vbaProject.bin opbouwen
# ----------------------------------------------------------------------------------------------

def _normalize_source(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text and not text.endswith("\n"):
        text += "\n"
    return text.replace("\n", "\r\n")


def build_vbaproject(template_bin: bytes, modules: dict[str, str]) -> bytes:
    """Nieuw vbaProject.bin: template met de broncode van `modules` ({modulenaam: brontekst}) vervangen.

    De bron wordt opgeslagen in de code page van het project (PROJECTCODEPAGE), met CRLF-regeleinden.
    """
    import olefile

    ole = olefile.OleFileIO(template_bin)
    try:
        def read(path: str) -> bytes:
            with ole.openstream(path) as f:
                return f.read()

        project = read("PROJECT")
        projectwm = read("PROJECTwm")
        records = parse_dir(decompress(read("VBA/dir")))
        codepage = dir_codepage(records)
        mods = dir_modules(records, codepage)
        known = {m.name for m in mods}
        onbekend = set(modules) - known
        if onbekend:
            raise KeyError("module(s) niet in het template: %s" % ", ".join(sorted(onbekend)))

        vba = CfbNode("VBA", True)
        vba.children.append(CfbNode("_VBA_PROJECT", False, bytes.fromhex("cc61ffff000000")))
        for m in mods:
            raw = read("VBA/" + m.stream_name)
            if m.name in modules:
                source = _normalize_source(modules[m.name]).encode(codepage)
            else:
                source = decompress(raw[m.text_offset:])
            vba.children.append(CfbNode(m.stream_name, False, compress(source)))
            for r in m.records:
                if r.rid == REC_MODULEOFFSET:
                    r.data = struct.pack("<I", 0)
        vba.children.append(CfbNode("dir", False, compress(serialize_dir(records))))

        root = CfbNode("Root Entry", True)
        root.children.append(CfbNode("PROJECT", False, project))
        root.children.append(CfbNode("PROJECTwm", False, projectwm))
        root.children.append(vba)
        # overige streams in de root (bijv. een digitale handtekening) nemen we niet over
        return write_cfb(root)
    finally:
        ole.close()


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Maakt een nieuw vbaProject.bin met vervangen modulebron.")
    p.add_argument("--template", required=True, help="bestaand vbaProject.bin")
    p.add_argument("--module", action="append", default=[], metavar="NAAM=PAD",
                   help="modulenaam=pad naar .bas/.cls (meerdere keren toegestaan)")
    p.add_argument("--out", required=True, help="uitvoer vbaProject.bin")
    a = p.parse_args(argv)
    modules = {}
    for spec in a.module:
        if "=" not in spec:
            p.error("--module verwacht NAAM=PAD, kreeg %r" % spec)
        naam, pad = spec.split("=", 1)
        with open(pad, "rb") as f:
            raw = f.read()
        if raw.startswith(b"\xef\xbb\xbf"):
            raw = raw[3:]
        try:
            modules[naam] = raw.decode("utf-8")
        except UnicodeDecodeError:
            modules[naam] = raw.decode("cp1252")
    with open(a.template, "rb") as f:
        template = f.read()
    result = build_vbaproject(template, modules)
    with open(a.out, "wb") as f:
        f.write(result)
    print("%s: %d bytes, modules vervangen: %s" % (a.out, len(result), ", ".join(modules) or "-"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
