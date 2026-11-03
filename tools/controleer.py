"""Controle van het gebouwde werkbestand met LibreOffice (headless).

1. Een testkopie maken waarin XLOOKUP is vervangen door INDEX/MATCH (LibreOffice 24.2 kent XLOOKUP nog niet;
   bij een exacte zoekopdracht zonder 'niet gevonden'-waarde is de uitkomst gelijk).
2. De testkopie volledig laten doorrekenen en als .xlsx bewaren.
3. Vergelijken: 1. Cashflow, CF - hulp VENA, 2. Ontw PEW en 3. Restant bdgt in het werkbestand tegen de waarden die
   Excel in het oorspronkelijke FO heeft opgeslagen; daarna de controles op tab Start.

    python tools/controleer.py [werkbestand.xlsm] [--fo FO.xlsx]

De opzet van het FO (welke rijen en kolommen) staat in <werkbestand>.opzet.json, die build_werkbestand.py schrijft.
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


def vergelijkingen(cfg, fo_wb):
    """welke bereiken van het oude FO terug moeten komen in het werkbestand, volgens de opzet van het FO"""
    from openpyxl.utils import get_column_letter as L
    oz = cfg['opzet']
    cf, ob, k, pw = oz['cashflow'], oz['opbrengsten'], oz['kosten'], oz['pew']
    r1, rn = cf['periodes'][0][2], cf['periodes'][-1][2]
    il, it = cfg['invoer_laatste'], cfg['invoer_totaal']
    v = fo_wb['CF - hulp VENA']
    out = [('1. Cashflow', 'B%d:J%d' % (r1, rn), 'Invoer', 'B8:J%d' % il),
           ('1. Cashflow', 'D%d:F%d' % (cf['totaal'], cf['totaal']), 'Invoer', 'D%d:F%d' % (it, it)),
           ('1. Cashflow', 'D%d:F%d' % (cf['ok'], cf['ok']), 'Invoer', 'D%d:F%d' % (it + 3, it + 3)),
           ('CF - hulp VENA', 'C3:%s4' % L(v.max_column), 'CF - hulp VENA', 'C3:%s4' % L(v.max_column)),
           ('2. Ontw PEW', 'E%d:%s%d' % (pw['eerste'], pw['kolommen'][-1], pw['resultaat_pct']), '2. Ontw PEW',
            'E%d:%s%d' % (pw['eerste'], pw['kolommen'][-1], pw['resultaat_pct']))]
    a = fo_wb['Aansluiting opbrengsten']
    out.append(('Aansluiting opbrengsten', 'B1:J%d' % a.max_row, 'Aansluiting opbrengsten', 'B1:J%d' % a.max_row))
    o = '%s%d:%s%d' % (ob['eerste'], ob['types'][0]['omzet'][0], ob['laatste'], ob['totaal'])
    out.append(('CF - opbrengsten', o, 'CF - opbrengsten', o))
    kk = '%s%d:%s%d' % (k['eerste'], k['posten'][0], k['laatste'], k['periode'])
    out.append(('CF - kosten', kk, 'CF - kosten', kk))
    bekend = {'2. Ontw PEW'}
    for n, _ in oz['bladen']:
        if 'PEW' in n and n not in bekend:
            ws = fo_wb[n]
            rng = 'E1:%s%d' % (L(ws.max_column), ws.max_row)
            out.append((n, rng, n, rng))
    return out


def vergelijk(nieuw, fo_path, cfg, toon=5):
    import openpyxl
    a = openpyxl.load_workbook(fo_path, data_only=True)
    b = openpyxl.load_workbook(nieuw, data_only=True)
    fouten = 0
    for blad_a, ber_a, blad_b, ber_b in vergelijkingen(cfg, a):
        n = 0
        for ra, rb in zip(a[blad_a][ber_a], b[blad_b][ber_b]):
            for ca, cb in zip(ra, rb):
                va, vb = ca.value, cb.value
                if blad_a == '1. Cashflow' and ca.column == 3 and not (isinstance(va, str) and va.startswith('Q')):
                    va = None           # jaarrij: in sommige FO's staat het jaar ook in kolom Q
                if isinstance(va, (int, float)) or isinstance(vb, (int, float)):
                    x = va if isinstance(va, (int, float)) else 0
                    y = vb if isinstance(vb, (int, float)) else 0
                    if abs(x - y) > 0.01 and not (va in (None, '') and vb in (None, '', 0)):
                        if n < toon:
                            print('    VERSCHIL %s!%s: FO %r, werkbestand %r' % (blad_b, cb.coordinate, va, vb))
                        n += 1
                elif (va or '') != (vb or ''):
                    if n < toon:
                        print('    VERSCHIL %s!%s: FO %r, werkbestand %r' % (blad_b, cb.coordinate, va, vb))
                    n += 1
        print('  %-24s %-9s -> %-24s %-9s %s' % (blad_a, ber_a, blad_b, ber_b, 'gelijk' if n == 0 else '%d verschillen' % n))
        fouten += n
    return fouten, b


def scenario(wb, cfg):
    """wat het scenariomodel nu (met formules) ziet"""
    w = wb['Woningtypes']
    for k in range(cfg['daeb_type']):
        c = 3 + 3 * k
        naam = w.cell(7, c).value
        if not naam:
            continue
        print('  Woningtype %d: %s · %s woningen · koopsom %s · start bouw Q%s %s · grond %.1f%%' % (
            k + 1, naam, w.cell(8, c).value, w.cell(9, c).value, w.cell(11, c).value, w.cell(10, c).value,
            100 * (w.cell(15, c + 1).value or 0)))
        for titel, r1, n in (('termijnen', 16, 10), ("extra's", 33, 6), ('bouwkosten', 82, 10)):
            regels = ['%s %s kw%s' % (w.cell(r, c).value, round(w.cell(r, c + 1).value, 4) if isinstance(w.cell(r, c + 1).value, float)
                                      else w.cell(r, c + 1).value, w.cell(r, c + 2).value)
                      for r in range(r1, r1 + n) if w.cell(r, c).value]
            print('     %-10s %s' % (titel, ' | '.join(regels)))
    i = wb['Invoer']
    print('  Invoer verkocht/transport (type 1, 2) per periode:')
    print('     ' + ' '.join('%s%s:%s/%s,%s/%s' % (i.cell(r, 2).value, i.cell(r, 3).value or '', i.cell(r, 12).value,
                                                    i.cell(r, 13).value, i.cell(r, 14).value, i.cell(r, 15).value)
                             for r in range(8, cfg['invoer_laatste'] + 1)))
    k = wb['Scenario kosten']
    print('  Scenario kosten: ' + ' | '.join('%s (%s) gerealiseerd %s totaal %s' % (k.cell(r, 2).value, k.cell(r, 3).value,
                                                                                   round(k.cell(r, 5).value or 0), round(k.cell(r, 4).value or 0))
                                             for r in range(8, 11)))
    print('  Scenario kosten totaal %s · Invoer kosten %s' % (k['D28'].value, k['D29'].value))
    m = wb['Model']
    d = wb['Dashboard']
    print('  Model: periodes %s · actuals t/m %s · stand %s · max. financiering %s (%s) · eindsaldo %s · break-even %s'
          % (m['BY8'].value, m['BY11'].value, m['BY12'].value, m['BY18'].value, m['BY19'].value, m['BY20'].value, m['BY21'].value))
    print('  Dashboard: model dekt opbrengsten %s · start bouw project %s' % (d['W13'].value, d['W9'].value))


def main():
    import argparse
    import json
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsm', nargs='?', default=os.path.join(ROOT, 'output', 'FO_Parkzoom_werkbestand.xlsm'))
    ap.add_argument('--fo', help='het FO waaruit het werkbestand is gebouwd (standaard: zoals in de .opzet.json)')
    ap.add_argument('--toon', type=int, default=5, help='aantal verschillen per bereik om te tonen')
    a = ap.parse_args()
    xlsm = a.xlsm
    cfg = json.load(open(os.path.splitext(xlsm)[0] + '.opzet.json'))
    fo = a.fo or os.path.join(ROOT, 'input', cfg['fo'])
    tmp = tempfile.mkdtemp()
    test = os.path.join(tmp, 'test.xlsx')
    uit = os.path.join(tmp, 'doorgerekend.xlsx')
    print('Testkopie zonder XLOOKUP')
    testkopie(xlsm, test)
    print('Doorrekenen met LibreOffice')
    doorrekenen(test, uit)
    print('Vergelijken met het oorspronkelijke FO (waarden van Excel)')
    fouten, wb = vergelijk(uit, fo, cfg, a.toon)
    ws = wb['Start']
    print('Controles op tab Start')
    for r in range(15, 19):
        print('  %-16s %-12s %s' % (ws.cell(r, 2).value, ws.cell(r, 4).value, ws.cell(r, 8).value))
    for r in range(23, 28):
        print('  %-50s %s %s' % (ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value or ''))
    for r in range(31, 45):
        print('  %-62s %-28s %s' % (ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value))
    print('  Start E8/E9/C11: %s / %s / %s' % (ws['E8'].value, ws['E9'].value, ws['C11'].value))
    vt = wb['Verkoop & transport']
    for row in vt.iter_rows(min_row=9, max_row=vt.max_row):
        if row[1].value and any(isinstance(c.value, str) and c.value[:2] in ('OK', 'FO', 'LE') for c in row[20:]):
            st = [c.value for c in row[20:] if isinstance(c.value, str) and c.value[:2] in ('OK', 'FO', 'LE')]
            print('  Verkoop & transport rij %d %-50s %s' % (row[0].row, row[1].value, st))
    sp = wb['SAP-correcties']
    print('  SAP-correcties: ' + ' | '.join('%s/%s %s' % (sp.cell(r, 2).value, sp.cell(r, 3).value or '', round(sp.cell(r, 6).value or 0))
                                            for r in range(7, 13)) + ' | niet toegewezen %s · totaal %s' % (sp['F13'].value, sp['F14'].value))
    fouten_cellen = []
    scenario(wb, cfg)
    for blad in ('Start', 'Invoer', 'Woningtypes', 'Scenario kosten', 'Dashboard', 'Model', '_koppeling', '2. Ontw PEW',
                 'Controle oud FO', 'CF - opbrengsten', 'CF - kosten', 'CF - hulp VENA', 'PowerPoint',
                 'Verkoop & transport', 'SAP-correcties'):
        for row in wb[blad].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#') and c.value[1:4].isupper() and not (
                        c.value == '#N/A' and blad in ('Model', 'PowerPoint')):   # NA() = bewust gat in grafiekreeksen
                    fouten_cellen.append('%s!%s=%s' % (blad, c.coordinate, c.value))
    print('Foutwaarden (#...): %d %s' % (len(fouten_cellen), fouten_cellen[:12]))
    out = os.path.join(os.path.dirname(os.path.abspath(xlsm)), 'doorgerekend_libreoffice.tmp.xlsx')
    os.replace(uit, out)
    print('Doorgerekende kopie: ' + out)
    return 1 if fouten or fouten_cellen else 0


if __name__ == '__main__':
    sys.exit(main())
