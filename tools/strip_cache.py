"""Verwijdert opgeslagen rekenresultaten (Excel rekent bij openen alles opnieuw). usage: strip_cache.py in out"""
import sys, os, re
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import N
bk = Book(sys.argv[1])
for name in list(bk.paths):
    s = bk.sheet(name)
    for row in s.rows.values():
        for el in row.findall(N+'c'):
            if el.find(N+'f') is None: continue
            v = el.find(N+'v')
            if v is not None: el.remove(v)
            if el.get('t') in ('str', 'e', 'b', 'n'): del el.attrib['t']
for fn in list(bk.data):
    if fn.startswith('xl/charts/chart') and fn.endswith('.xml'):
        d = bk.text(fn); d = re.sub(r'<c:pt idx="\d+">.*?</c:pt>', '', d, flags=re.S); d = re.sub(r'<c:ptCount val="\d+"/>', '<c:ptCount val="0"/>', d); bk.settext(fn, d)
bk.save(sys.argv[2])
