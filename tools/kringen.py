"""Kringverwijzingen zoeken zoals Excel ze ziet: elke verwijzing telt, ook in een ALS-tak die niet gekozen wordt.
Verwijzingen binnen KOLOM/RIJ tellen niet (Excel ziet die niet als kring).

    python tools/kringen.py werkbestand.xlsm
"""
import sys, re, bisect, collections
import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string as ci, get_column_letter as L
sys.setrecursionlimit(100000)
wb = openpyxl.load_workbook(sys.argv[1], keep_vba=False)
names = {}
for n, dn in wb.defined_names.items():
    names[n.lower()] = dn.attr_text
for ws in wb.worksheets:
    for n, dn in getattr(ws, 'defined_names', {}).items():
        names.setdefault(n.lower(), dn.attr_text)
formules = {}           # (blad, r, c) -> formule
percol = collections.defaultdict(list)    # (blad, c) -> rijen met formule
maxr = {}
for ws in wb.worksheets:
    maxr[ws.title] = ws.max_row
    for row in ws.iter_rows():
        for cel in row:
            v = cel.value
            if isinstance(v, str) and v.startswith('='):
                formules[(ws.title, cel.row, cel.column)] = v
            elif type(v).__name__ == 'ArrayFormula':
                formules[(ws.title, cel.row, cel.column)] = v.text
for (b, r, c) in formules:
    percol[(b, c)].append(r)
for k in percol:
    percol[k].sort()
REF = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?\$?([A-Z]{1,3})?\$?(\d+)?(?::\$?([A-Z]{1,3})?\$?(\d+)?)?$")
def refs(blad, f):
    try:
        tok = Tokenizer(f)
    except Exception:
        return []
    uit = []
    diepte = []
    for t in tok.items:
        if t.type == 'FUNC' and t.subtype == 'OPEN':
            diepte.append(t.value.upper().rstrip('('))
            continue
        if t.type == 'FUNC' and t.subtype == 'CLOSE':
            if diepte: diepte.pop()
            continue
        if t.type != 'OPERAND' or t.subtype != 'RANGE':
            continue
        if diepte and diepte[-1] in ('COLUMN', 'ROW', 'COLUMNS', 'ROWS'):
            continue
        v = t.value
        if v.lower() in names:
            v = names[v.lower()]
            if v.upper().startswith('OFFSET('):
                v = v[7:].split(',')[0]
        m = REF.match(v.replace('[0]!', ''))
        if not m:
            continue
        b = (m.group(1) or m.group(2) or blad).replace("''", "'")
        if b not in maxr:
            continue
        c1, r1, c2, r2 = m.group(3), m.group(4), m.group(5), m.group(6)
        if not (c1 or r1):
            continue
        if m.group(0).count(':') == 0:
            c2, r2 = c1, r1
        c1n = ci(c1) if c1 else 1
        c2n = ci(c2) if c2 else 16384
        r1n = int(r1) if r1 else 1
        r2n = int(r2) if r2 else maxr[b]
        uit.append((b, min(c1n, c2n), max(c1n, c2n), min(r1n, r2n), max(r1n, r2n)))
    return uit
# grof: per kolom
cols_graph = collections.defaultdict(set)
cell_refs = {}
for (b, r, c), f in formules.items():
    rs = refs(b, f[1:] if not f.startswith('=') else f)
    cell_refs[(b, r, c)] = rs
    for (b2, c1, c2, r1, r2) in rs:
        for cc in range(c1, min(c2, c1 + 400) + 1):
            if (b2, cc) in percol:
                cols_graph[(b, c)].add((b2, cc))
# SCC op kolomniveau (Tarjan iteratief)
idx, low, onst, st, sccs, teller = {}, {}, set(), [], [], [0]
def tarjan(v0):
    stack = [(v0, iter(cols_graph.get(v0, ())))]
    idx[v0] = low[v0] = teller[0]; teller[0] += 1; st.append(v0); onst.add(v0)
    while stack:
        v, it = stack[-1]
        for w in it:
            if w not in idx:
                idx[w] = low[w] = teller[0]; teller[0] += 1; st.append(w); onst.add(w)
                stack.append((w, iter(cols_graph.get(w, ()))))
                break
            elif w in onst:
                low[v] = min(low[v], idx[w])
        else:
            stack.pop()
            if stack:
                low[stack[-1][0]] = min(low[stack[-1][0]], low[v])
            if low[v] == idx[v]:
                comp = []
                while True:
                    w = st.pop(); onst.discard(w); comp.append(w)
                    if w == v: break
                sccs.append(comp)
for v in list(percol):
    if v not in idx:
        tarjan(v)
kandidaat = set()
for comp in sccs:
    if len(comp) > 1 or comp[0] in cols_graph.get(comp[0], ()):
        kandidaat.update(comp)
print('kolommen in een mogelijke kring:', len(kandidaat))
# fijn: celniveau binnen de kandidaatkolommen
cel_graph = {}
for (b, r, c), rs in cell_refs.items():
    if (b, c) not in kandidaat:
        continue
    doel = []
    for (b2, c1, c2, r1, r2) in rs:
        for cc in range(c1, min(c2, c1 + 400) + 1):
            if (b2, cc) not in kandidaat:
                continue
            rows = percol[(b2, cc)]
            i = bisect.bisect_left(rows, r1); j = bisect.bisect_right(rows, r2)
            doel.extend((b2, rr, cc) for rr in rows[i:j])
    cel_graph[(b, r, c)] = doel
idx2, low2, on2, st2, gevonden, t2 = {}, {}, set(), [], [], [0]
def tarjan2(v0):
    stack = [(v0, iter(cel_graph.get(v0, ())))]
    idx2[v0] = low2[v0] = t2[0]; t2[0] += 1; st2.append(v0); on2.add(v0)
    while stack:
        v, it = stack[-1]
        for w in it:
            if w not in cel_graph:
                continue
            if w not in idx2:
                idx2[w] = low2[w] = t2[0]; t2[0] += 1; st2.append(w); on2.add(w)
                stack.append((w, iter(cel_graph.get(w, ()))))
                break
            elif w in on2:
                low2[v] = min(low2[v], idx2[w])
        else:
            stack.pop()
            if stack:
                low2[stack[-1][0]] = min(low2[stack[-1][0]], low2[v])
            if low2[v] == idx2[v]:
                comp = []
                while True:
                    w = st2.pop(); on2.discard(w); comp.append(w)
                    if w == v: break
                if len(comp) > 1 or v in cel_graph.get(v, ()):
                    gevonden.append(comp)
for v in cel_graph:
    if v not in idx2:
        tarjan2(v)
print('kringen op celniveau:', len(gevonden))
for comp in gevonden[:15]:
    cs = sorted(comp)[:6]
    print('  %d cellen, o.a.:' % len(comp), ', '.join("%s!%s%d" % (b, L(c), r) for b, r, c in cs))
    b, r, c = cs[0]
    print('     ', formules[(b, r, c)][:200])
