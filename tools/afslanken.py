"""Maakt het bestand kleiner door doorgetrokken formules als 'gedeelde formule' op te slaan.

Gebruik:  python3 tools/afslanken.py IN.xlsm UIT.xlsm
Laatste stap: de andere scripts in tools/ werken op het bestand vóór deze stap (bron/RDFO_template_bron.xlsm).

Een kolom met honderden keren dezelfde, doorgetrokken formule staat daarna één keer in het bestand.
Er verandert niets aan de berekening: elke formule wordt vooraf vergeleken met wat Excel uit de
gedeelde formule afleidt, en alleen bij een exacte overeenkomst gedeeld.
"""
import re
import sys
import zipfile
from xml.sax.saxutils import unescape

from openpyxl.formula.translate import Translator

SRC, DST = sys.argv[1], sys.argv[2]
CEL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)><f>([^<]*)</f>')
REL_RIJ = re.compile(r'(?<![A-Za-z0-9_$])\$?[A-Z]{1,3}[0-9]+')        # een verwijzing met een relatieve rij


def afgeleid(basis, kolom, rij0, rij):
    """De formule die Excel in (kolom, rij) afleidt uit de formule in (kolom, rij0)."""
    if not REL_RIJ.search(basis):
        return basis
    return Translator('=' + basis, '%s%d' % (kolom, rij0)).translate_formula('%s%d' % (kolom, rij))[1:]


def blad(x):
    per_kolom = {}
    for m in CEL.finditer(x):
        per_kolom.setdefault(m.group(1), []).append((int(m.group(2)), m))
    vervang = []                 # (begin, eind, nieuwe tekst) voor het <f>-element
    si = 0
    for kolom, cellen in per_kolom.items():
        i = 0
        while i < len(cellen):
            rij0, m0 = cellen[i]
            basis = unescape(m0.group(4))
            j = i + 1
            while (j < len(cellen) and cellen[j][0] == rij0 + (j - i)
                   and unescape(cellen[j][1].group(4)) == afgeleid(basis, kolom, rij0, cellen[j][0])):
                j += 1
            if j - i >= 3 and '[' not in basis:
                vervang.append((m0.start(4) - 3, m0.end(4) + 4,
                                '<f t="shared" ref="%s%d:%s%d" si="%d">%s</f>' % (kolom, rij0, kolom, cellen[j - 1][0], si, m0.group(4))))
                for _, m in cellen[i + 1:j]:
                    vervang.append((m.start(4) - 3, m.end(4) + 4, '<f t="shared" si="%d"/>' % si))
                si += 1
            i = j
    vervang.sort()
    uit, pos = [], 0
    for a, b, t in vervang:
        assert x[a:a + 3] == '<f>' and x[b - 4:b] == '</f>'
        uit.append(x[pos:a])
        uit.append(t)
        pos = b
    uit.append(x[pos:])
    return ''.join(uit), len(vervang), si


def main():
    zi = zipfile.ZipFile(SRC)
    zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED, compresslevel=9)
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
            assert b't="shared"' not in d, 'bestand is al afgeslankt'
            x, n, groepen = blad(d.decode('utf8'))
            print('%-28s %7d formules in %5d groepen   %9d -> %9d bytes' % (it.filename, n, groepen, len(d), len(x.encode('utf8'))))
            d = x.encode('utf8')
        zo.writestr(it, d)
    zo.close()


if __name__ == '__main__':
    main()
