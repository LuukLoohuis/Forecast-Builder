"""Controle van het gebouwde werkbestand met LibreOffice (headless).

1. Een testkopie maken waarin XLOOKUP is vervangen door INDEX/MATCH (LibreOffice 24.2 kent XLOOKUP nog niet;
   bij een exacte zoekopdracht zonder 'niet gevonden'-waarde is de uitkomst gelijk).
2. De testkopie volledig laten doorrekenen en als .xlsx bewaren.
3. Vergelijken: 1. Cashflow, CF - hulp VENA, 2. Ontw PEW en 3. Restant bdgt in het werkbestand tegen de waarden die
   Excel in het oorspronkelijke FO heeft opgeslagen; daarna de controles op tab Start.

    python tools/controleer.py [werkbestand.xlsm]
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def split_args(s, start):
    """argumenten van de functie waarvan '(' op positie start staat; geeft (args, eindpositie)"""
    depth, i, args, cur, q = 0, start, [], '', None
    while i < len(s):
        ch = s[i]
        if q:
            cur += ch
            if ch == q:
                q = None
        elif ch in '"\'':
            q = ch
            cur += ch
        elif ch == '(':
            depth += 1
            if depth > 1:
                cur += ch
        elif ch == ')':
            depth -= 1
            if depth == 0:
                args.append(cur)
                return args, i
            cur += ch
        elif ch == ',' and depth == 1:
            args.append(cur)
            cur = ''
        else:
            cur += ch
        i += 1
    raise ValueError(s)


def zonder_xlookup(f):
    while '_xlfn.XLOOKUP(' in f:
        p = f.index('_xlfn.XLOOKUP(')
        args, end = split_args(f, p + len('_xlfn.XLOOKUP'))
        if len(args) != 3:
            raise ValueError(f)
        a, b, c = [zonder_xlookup(x) for x in args]
        f = f[:p] + 'INDEX(%s,MATCH(%s,%s,0))' % (c, a, b) + f[end + 1:]
    return f


def testkopie(src, dst):
    """XLOOKUP -> INDEX/MATCH; in FO - actuals de jaar- en periodekolommen (tekst uit SAP) als getal, zoals Excel ze in
    SUMIFS-criteria toch al als getal ziet (LibreOffice doet dat niet)"""
    import html
    from xlsxpkg import Package
    pkg = Package(src)
    act = pkg.sheet_part('FO - actuals')
    sst = [''.join(t.itertext()) for t in pkg.x('xl/sharedStrings.xml')]
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            b = zi.read(n)
            if n.startswith('xl/worksheets/sheet') and b'XLOOKUP' in b:
                s = b.decode('utf-8')

                def rep(m):
                    return m.group(1) + html.escape(zonder_xlookup(html.unescape(m.group(2))), quote=False) + '</f>'
                s = re.sub(r'(<f[^>]*>)([^<]*)</f>', rep, s)
                b = s.encode('utf-8')
            if n == act:
                s = b.decode('utf-8')

                def num(m):
                    txt = sst[int(m.group(3))]
                    if txt.strip().isdigit():
                        return '<c r="%s"%s><v>%s</v></c>' % (m.group(1), m.group(2), txt.strip())
                    return m.group(0)
                s = re.sub(r'<c r="(P\d+)"([^>]*?) t="s"><v>(\d+)</v></c>', num, s)
                s = re.sub(r'<c r="(P\d+)"([^>]*?) t="inlineStr"><is><t>\s*(\d+)\s*</t></is></c>',
                           lambda m: '<c r="%s"%s><v>%s</v></c>' % (m.group(1), m.group(2), m.group(3)), s)
                b = s.encode('utf-8')
            zo.writestr(n, b)

MACRO = r'''
import uno, sys
from com.sun.star.beans import PropertyValue
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for i in range(60):
    try:
        ctx = resolver.resolve("uno:pipe,name=fobouw;urp;StarOffice.ComponentContext"); break
    except Exception:
        import time; time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[1]), "_blank", 0, (pv("Hidden", True),))
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(sys.argv[2]), (pv("FilterName", "Calc MS Excel 2007 XML"),))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
'''


def doorrekenen(src, dst):
    tmp = tempfile.mkdtemp()
    script = os.path.join(tmp, 'reken.py')
    open(script, 'w').write(MACRO)
    prof = 'file://' + os.path.join(tmp, 'profiel')
    lo = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', '-env:UserInstallation=' + prof,
                           '--accept=pipe,name=fobouw;urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        subprocess.run([sys.executable, script, src, dst], check=True, timeout=900)
    finally:
        time.sleep(1)
        lo.kill()


def vergelijk(nieuw, fo_path):
    import openpyxl
    a = openpyxl.load_workbook(fo_path, data_only=True)
    b = openpyxl.load_workbook(nieuw, data_only=True)
    fouten = 0
    for blad, bereik in (('1. Cashflow', 'B4:J26'), ('CF - hulp VENA', 'C3:BK4'), ('2. Ontw PEW', 'E7:I27'),
                         ('3. Restant bdgt', 'E7:K20'), ('Aansluiting opbrengsten', 'B15:J29')):
        n = 0
        for ra, rb in zip(a[blad][bereik], b[blad][bereik]):
            for ca, cb in zip(ra, rb):
                va, vb = ca.value, cb.value
                if isinstance(va, (int, float)) or isinstance(vb, (int, float)):
                    x = va if isinstance(va, (int, float)) else 0
                    y = vb if isinstance(vb, (int, float)) else 0
                    if abs(x - y) > 0.01 and not (va in (None, '') and vb in (None, '', 0)):
                        if n < 5:
                            print('    VERSCHIL %s!%s: FO %r, werkbestand %r' % (blad, ca.coordinate, va, vb))
                        n += 1
                elif (va or '') != (vb or ''):
                    if n < 5:
                        print('    VERSCHIL %s!%s: FO %r, werkbestand %r' % (blad, ca.coordinate, va, vb))
                    n += 1
        print('  %-26s %-8s %s' % (blad, bereik, 'gelijk' if n == 0 else '%d verschillen' % n))
        fouten += n
    return fouten, b


def main():
    xlsm = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'output', 'FO_Parkzoom_werkbestand.xlsm')
    fo = os.path.join(ROOT, 'input', 'FO_Parkzoom_fictieve_demoversie.xlsx')
    tmp = tempfile.mkdtemp()
    test = os.path.join(tmp, 'test.xlsx')
    uit = os.path.join(tmp, 'doorgerekend.xlsx')
    print('Testkopie zonder XLOOKUP')
    testkopie(xlsm, test)
    print('Doorrekenen met LibreOffice')
    doorrekenen(test, uit)
    print('Vergelijken met het oorspronkelijke FO (waarden van Excel)')
    fouten, wb = vergelijk(uit, fo)
    ws = wb['Start']
    print('Controles op tab Start')
    for r in range(15, 19):
        print('  %-16s %-12s %s' % (ws.cell(r, 2).value, ws.cell(r, 4).value, ws.cell(r, 8).value))
    for r in range(23, 28):
        print('  %-50s %s %s' % (ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value or ''))
    for r in range(31, 44):
        print('  %-62s %-28s %s' % (ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value))
    print('  Start E8/E9/C11: %s / %s / %s' % (ws['E8'].value, ws['E9'].value, ws['C11'].value))
    fouten_cellen = []
    for blad in ('Start', 'Prognose kosten', 'Prognose opbrengsten', 'Controle oud FO', '1. Cashflow', 'CF - hulp VENA'):
        for row in wb[blad].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#') and c.value[1:4].isupper():
                    fouten_cellen.append('%s!%s=%s' % (blad, c.coordinate, c.value))
    print('Foutwaarden (#...) in de nieuwe tabbladen: %d %s' % (len(fouten_cellen), fouten_cellen[:10]))
    out = os.path.join(ROOT, 'output', 'doorgerekend_libreoffice.tmp.xlsx')
    os.replace(uit, out)
    print('Doorgerekende kopie: ' + out)
    return 1 if fouten or fouten_cellen else 0


if __name__ == '__main__':
    sys.exit(main())
