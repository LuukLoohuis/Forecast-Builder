"""Bouwt vbaProject.bin opnieuw: broncode per module (optioneel vervangen), zonder gecompileerde p-code.
_VBA_PROJECT krijgt versie 0xFFFF, zodat Excel bij openen opnieuw compileert uit de broncode.
usage: vba_rebuild.py in.bin out.bin [Module=pad.bas ...]"""
import sys, struct, os
sys.path.insert(0, os.path.dirname(__file__))
src = open(os.path.join(os.path.dirname(__file__), 'readvba.py')).read().split("data=open(sys.argv[1]")[0]
exec(src)   # cfb(), decompress()

def compress(data):
    out = bytearray(b'\x01'); pos = 0
    while pos < len(data):
        chunk = data[pos:pos+4096]; pos += 4096
        body = bytearray(); i = 0; n = len(chunk)
        index = {}
        while i < n:
            flagpos = len(body); body.append(0); flags = 0
            for b in range(8):
                if i >= n: break
                best_len = 0; best_off = 0
                if i > 0:
                    bc = max((i - 1).bit_length(), 4)
                    maxlen = (0xFFFF >> bc) + 3; maxoff = 1 << (16 - bc)
                    key = bytes(chunk[i:i+3])
                    for c in reversed(index.get(key, [])):
                        off = i - c
                        if off > maxoff: break
                        l = 0
                        while l < maxlen and i + l < n and chunk[c + l] == chunk[i + l]: l += 1
                        if l > best_len: best_len, best_off = l, off
                        if l >= maxlen: break
                if best_len >= 3:
                    bc = max((i - 1).bit_length(), 4)
                    tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                    body += struct.pack('<H', tok); flags |= (1 << b)
                    for k in range(best_len):
                        index.setdefault(bytes(chunk[i+k:i+k+3]), []).append(i + k)
                    i += best_len
                else:
                    body.append(chunk[i]); index.setdefault(bytes(chunk[i:i+3]), []).append(i); i += 1
            body[flagpos] = flags
        if len(body) + 2 > 4098 and n == 4096:
            out += struct.pack('<H', 0x3000 | 4095) + chunk          # ongecomprimeerd
        else:
            assert len(body) + 2 <= 4098
            out += struct.pack('<H', 0x8000 | 0x3000 | (len(body) + 2 - 3)) + body
    return bytes(out)

def name_key(nm): return (len(nm), nm.upper())

def write_cfb(tree):
    """tree: dict naam -> bytes (stream) of dict (storage). Geeft bytes van een CFB v3-bestand."""
    SS, MSS, CUT = 512, 64, 4096
    entries = []   # [name, type, children_idx(list), data]
    def add(name, node, typ):
        idx = len(entries); entries.append([name, typ, [], node if isinstance(node, bytes) else None, None, None, None])
        if isinstance(node, dict):
            for k in sorted(node, key=name_key): entries[idx][2].append(add(k, node[k], 1 if isinstance(node[k], dict) else 2))
        return idx
    add('Root Entry', tree, 5)
    # gebalanceerde boom per storage
    for e in entries:
        kids = e[2]
        def build(lo, hi):
            if lo > hi: return 0xFFFFFFFF
            mid = (lo + hi) // 2; k = kids[mid]
            entries[k][4] = build(lo, mid - 1); entries[k][5] = build(mid + 1, hi); return k
        e[6] = build(0, len(kids) - 1) if kids else 0xFFFFFFFF
    # mini stream
    mini = bytearray(); minifat = []
    big = []   # (entry idx, data)
    for i, e in enumerate(entries):
        d = e[3]
        if d is None: continue
        if len(d) < CUT:
            start = len(mini) // MSS; cnt = (len(d) + MSS - 1) // MSS
            if cnt == 0: e.append((0xFFFFFFFE, 0)); continue
            for k in range(cnt): minifat.append(start + k + 1 if k < cnt - 1 else 0xFFFFFFFE)
            mini += d + b'\0' * (cnt * MSS - len(d)); e.append((start, len(d)))
        else:
            big.append(i)
    nsec = lambda b: (len(b) + SS - 1) // SS
    ndir = (len(entries) * 128 + SS - 1) // SS
    mf_bytes = b''.join(struct.pack('<I', x) for x in minifat); nmf = nsec(mf_bytes)
    nmini = nsec(mini)
    nbig = sum(nsec(entries[i][3]) for i in big)
    data_secs = ndir + nmf + nmini + nbig
    nfat = 1
    while (data_secs + nfat) > nfat * (SS // 4): nfat += 1
    assert nfat <= 109
    fat = [0xFFFFFFFF] * (nfat * (SS // 4))
    cur = 0
    def chain(n):
        nonlocal cur
        s = cur
        for k in range(n): fat[cur + k] = cur + k + 1 if k < n - 1 else 0xFFFFFFFE
        cur += n; return s if n else 0xFFFFFFFE
    dir_start = chain(ndir); mf_start = chain(nmf); mini_start = chain(nmini)
    big_start = {}
    for i in big: big_start[i] = chain(nsec(entries[i][3]))
    fat_secs = list(range(cur, cur + nfat))
    for s in fat_secs: fat[s] = 0xFFFFFFFD
    body = bytearray()
    # directory
    dirb = bytearray()
    for i, e in enumerate(entries):
        nm = e[0].encode('utf-16le') + b'\0\0'
        rec = bytearray(128); rec[0:len(nm)] = nm
        struct.pack_into('<H', rec, 64, len(nm)); rec[66] = e[1]; rec[67] = 1
        left = e[4] if e[4] is not None else 0xFFFFFFFF; right = e[5] if e[5] is not None else 0xFFFFFFFF
        child = e[6] if e[1] in (1, 5) else 0xFFFFFFFF
        struct.pack_into('<III', rec, 68, left, right, child)
        if e[1] == 5: start, size = (mini_start if nmini else 0xFFFFFFFE), len(mini)
        elif e[1] == 1: start, size = 0, 0
        elif i in big_start: start, size = big_start[i], len(e[3])
        else: start, size = e[7]
        struct.pack_into('<IQ', rec, 116, start, size)
        dirb += rec
    pad = lambda b: bytes(b) + b'\0' * (nsec(b) * SS - len(b))
    body += pad(dirb) + (pad(mf_bytes) if nmf else b'') + (pad(mini) if nmini else b'')
    for i in big: body += pad(entries[i][3])
    body += b''.join(struct.pack('<I', x) for x in fat)
    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex('D0CF11E0A1B11AE1')
    struct.pack_into('<HHHHHH', hdr, 24, 0x3E, 3, 0xFFFE, 9, 6, 0)
    struct.pack_into('<IIIIIIIII', hdr, 40, 0, nfat, dir_start, 0, CUT, mf_start if nmf else 0xFFFFFFFE, nmf, 0xFFFFFFFE, 0)
    difat = fat_secs + [0xFFFFFFFF] * (109 - nfat)
    struct.pack_into('<109I', hdr, 76, *difat)
    return bytes(hdr) + bytes(body)

def main(inp, outp, repl):
    data = open(inp, 'rb').read(); paths, read = cfb(data)
    dirraw = read(paths['/VBA/dir']); d = bytearray(decompress(dirraw))
    i = 0; mods = []; cur = None
    while i < len(d):
        rid, sz = struct.unpack_from('<HI', d, i); hdr_i = i; i += 6
        if rid == 9: sz = 6
        if rid == 0x19: cur = {'name': d[i:i+sz].decode('latin1')}
        elif rid == 0x1A: cur['stream'] = d[i:i+sz].decode('latin1')
        elif rid == 0x31: cur['off'] = struct.unpack_from('<I', d, i)[0]; cur['offpos'] = i
        elif rid == 0x2B: mods.append(cur)
        i += sz
    tree = {}
    for p, e in paths.items():
        parts = p.strip('/').split('/')
        if e['typ'] != 2: continue
        if parts[0] == 'VBA' and (parts[-1].startswith('__SRP_') or parts[-1] == '_VBA_PROJECT' or parts[-1] == 'dir'): continue
        node = tree
        for q in parts[:-1]: node = node.setdefault(q, {})
        node[parts[-1]] = read(e)
    vba = tree.setdefault('VBA', {})
    for m in mods:
        raw = read(paths['/VBA/' + m['stream']]); source = decompress(raw[m['off']:])
        if m['name'] in repl:
            new = open(repl[m['name']], 'rb').read()
            if b'\r\n' not in new: new = new.replace(b'\n', b'\r\n')
            source = new
        vba[m['stream']] = compress(source)
        assert decompress(vba[m['stream']]) == source
        struct.pack_into('<I', d, m['offpos'], 0)
    vba['dir'] = compress(bytes(d)); assert decompress(vba['dir']) == bytes(d)
    vba['_VBA_PROJECT'] = bytes.fromhex('CC61FFFF000000')
    out = write_cfb(tree); open(outp, 'wb').write(out)
    # controle: opnieuw lezen
    paths2, read2 = cfb(out)
    for m in mods:
        s2 = decompress(read2(paths2['/VBA/' + m['stream']]))
        assert s2.startswith(b'Attribute VB_Name'), m['name']
    print('ok', len(mods), 'modules,', len(out), 'bytes')

if __name__ == '__main__':
    repl = dict(a.split('=', 1) for a in sys.argv[3:])
    main(sys.argv[1], sys.argv[2], repl)
