"""Alles testen met data, in één commando (exit-code 0 = alles goed).

    python tools/test_alles.py                 bouwen + alle tests op het demo-FO
    python tools/test_alles.py --geen-bouw     de bestaande bestanden in output/ testen
    python tools/test_alles.py --fo FO.xlsx    daarnaast een eigen (echt) oud FO omzetten en vergelijken (blijft lokaal)

Tests
  1  controleer.py op het demo: geen foutwaarden, de cashflow gelijk aan het oude FO (op de 4 bekende kosten-
     verschillen van samen €9 na), alle controles op Start OK
  2  geen kringverwijzingen in het sjabloon (kringen.py)
  3  de macro Oud FO inlezen (VBA, in LibreOffice) op het lege sjabloon: cashflow gelijk aan het oude FO en per periode
     gelijk aan de Python-omzetting (pariteit)
  4  de macro Nieuw kwartaal starten (NieuwKwartaalUitvoeren): vorige prognose = de cashflow van daarvoor,
     FO - vorig = het oude FO - huidig, FO - huidig en FO - actuals leeg
  5  scenario's met data op het demo: kans grondtermijn 50% (netto 0), parkeren naar rato van de bouwtermijnen,
     Voorstel met een zelf ingevuld kwartaal, verkoop vóór de exportdatum telt niet, overlappende sleutels per type
  6  PowerPoint-sjabloon: elke vorm die de macro vult staat erin

LibreOffice (soffice) en de Python-module uno zijn nodig. Excel zelf wordt niet getest: daar is geen Excel voor.
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
import controleer  # noqa: E402
import standaard_fo as S  # noqa: E402

OUT = os.path.join(ROOT, 'output')
SJABLOON = os.path.join(OUT, 'FO_werkbestand_sjabloon.xlsm')
DEMO_WB = os.path.join(OUT, 'FO_Parkzoom_werkbestand.xlsm')
PPTX = os.path.join(OUT, 'Kwartaal_Template_cashflow_v22.pptx')
uitslag = []


def ok(naam, goed, detail=''):
    uitslag.append((naam, bool(goed), detail))
    print('  %s  %s%s' % ('OK  ' if goed else 'FOUT', naam, ('  · ' + detail) if detail else ''))


def run(cmd, timeout=1800):
    p = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, timeout=timeout)
    return p.returncode, p.stdout + p.stderr


# ---- LibreOffice ---------------------------------------------------------------------------------------
class LO:
    def __init__(self):
        self.tmp = tempfile.mkdtemp(prefix='lo_test_')
        self.pipe = 'fotest%d' % os.getpid()
        self.proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                                      '-env:UserInstallation=file://%s/prof' % self.tmp,
                                      '--accept=pipe,name=%s;urp;' % self.pipe],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        import uno
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(120):
            try:
                ctx = res.resolve('uno:pipe,name=%s;urp;StarOffice.ComponentContext' % self.pipe)
                break
            except Exception:
                time.sleep(1)
        else:
            raise RuntimeError('LibreOffice start niet')
        self.desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

    @staticmethod
    def pv(n, v):
        from com.sun.star.beans import PropertyValue
        p = PropertyValue()
        p.Name, p.Value = n, v
        return p

    def open(self, pad, macros=False):
        import uno
        props = [self.pv('Hidden', True)] + ([self.pv('MacroExecutionMode', 4)] if macros else [])
        return self.desk.loadComponentFromURL(uno.systemPathToFileUrl(pad), '_blank', 0, tuple(props))

    def macro(self, doc, naam, args=()):
        sp = doc.getScriptProvider()
        s = sp.getScript('vnd.sun.star.script:VBAProject.FOWerkbestand.%s?language=Basic&location=document' % naam)
        return s.invoke(tuple(args), (), ())[0]

    def stop(self):
        try:
            self.desk.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(20)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.tmp, ignore_errors=True)


def jaren_als_getal(doc):
    """LibreOffice ziet een jaartal als tekst in SUMIFS niet als getal (Excel wel): zoals in controleer.testkopie"""
    a = doc.Sheets.getByName('FO - actuals')
    for r in range(1, 7000):
        c = a.getCellByPosition(15, r)
        t = c.getString().strip()
        if t.isdigit() and c.getType().value == 'TEXT':
            c.setValue(float(t))


def vba_lint(tekst):
    """fouten die LibreOffice Basic niet ziet maar Excel wel: toewijzing aan een andere functie in een functie; Sub/Function
    zonder bijpassende End"""
    regels = tekst.replace('\r\n', '\n').split('\n')
    fout = []
    func = set(m.group(1).lower() for m in re.finditer(r'^(?:Public |Private |Friend )?(?:Static )?Function (\w+)', tekst, re.M | re.I))
    huidig, soort = None, None
    for nr, r in enumerate(regels, 1):
        code = re.sub(r"'.*$", '', re.sub(r'"[^"]*"', '""', r)).strip()
        m = re.match(r'^(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set)) (\w+)', code, re.I)
        if m:
            if huidig:
                fout.append('%d: %s %s begint voordat %s klaar is' % (nr, m.group(1), m.group(2), huidig))
            soort, huidig = m.group(1).split()[0].lower(), m.group(2)
            continue
        m = re.match(r'^End (Sub|Function|Property)\b', code, re.I)
        if m:
            if not huidig or m.group(1).lower() != soort:
                fout.append('%d: End %s zonder bijpassend begin' % (nr, m.group(1)))
            huidig = soort = None
            continue
        if huidig:
            for deel in code.split(':'):
                deel = re.sub(r'^.*\bThen\b\s*', '', deel.strip(), flags=re.I) if re.search(r'\bThen\b\s*\S', deel, re.I) else deel.strip()
                deel = re.sub(r'^Else\s+', '', deel, flags=re.I)
                a = re.match(r'^(\w+)\s*=(?!=)', deel)
                if a and a.group(1).lower() in func and a.group(1).lower() != huidig.lower():
                    fout.append('%d: in %s een toewijzing aan functie %s' % (nr, huidig, a.group(1)))
    if huidig:
        fout.append('einde bestand: %s niet afgesloten' % huidig)
    return fout


def waarde(sh, ref):
    return sh.getCellRangeByName(ref).getValue()


def tekst(sh, ref):
    return sh.getCellRangeByName(ref).getString()


def invoer_rijen(doc):
    inv = doc.Sheets.getByName('Invoer')
    return {(inv.getCellByPosition(1, r).getString(), inv.getCellByPosition(2, r).getString()):
            (inv.getCellByPosition(3, r).getValue(), inv.getCellByPosition(5, r).getValue(), inv.getCellByPosition(8, r).getValue())
            for r in range(7, 52) if inv.getCellByPosition(1, r).getString()}


# ---- tests ----------------------------------------------------------------------------------------------
def t_controleer(werkbestand, fo, naam, max_kosten_verschillen, streng=True):
    code, uit = run([sys.executable, 'tools/controleer.py', werkbestand] + (['--fo', fo] if fo else []))
    fout = re.search(r'Foutwaarden \(#\.\.\.\): (\d+)', uit)
    ok('%s: geen foutwaarden' % naam, fout and fout.group(1) == '0')
    m = re.search(r'Invoer, (\d+) periodes: (gelijk|(\d+) verschillen)', uit)
    omzet = re.findall(r'VERSCHIL \S+ omzet', uit)
    n = 0 if not m or m.group(2) == 'gelijk' else int(m.group(3))
    ok('%s: cashflow gelijk aan het oude FO' % naam, m and not omzet and n <= max_kosten_verschillen,
       m.group(0) if m else 'geen vergelijking')
    start = uit.split('Controles op tab Start')[-1]
    fouten = [re.sub(r'\s{2,}', ' ', l.strip()) for l in start.splitlines() if re.search(r'\s(FOUT)\s*$', l)]
    if streng:
        ok('%s: geen FOUT in de controles op Start' % naam, not fouten, '; '.join(fouten)[:200])
    else:                       # een eigen FO: de controles op Start zijn bevindingen over dat FO, geen testfout
        print('  INFO  %s: controles op Start: %s' % (naam, '; '.join(fouten) if fouten else 'alles OK'))


def t_kringen():
    code, uit = run([sys.executable, 'tools/kringen.py', SJABLOON])
    m = re.search(r'kringen op celniveau: (\d+)', uit)
    ok('sjabloon: geen kringverwijzingen', m and m.group(1) == '0', m.group(0) if m else uit[-200:])


def t_vba(lo, fo, naam, python_doc=None):
    kopie = os.path.join(lo.tmp, 'vba_%s.xlsm' % re.sub(r'\W', '', naam))
    controleer.testkopie(SJABLOON, kopie)
    doc = lo.open(kopie, macros=True)
    try:
        lo.macro(doc, 'OudFOOmzettenTest', (fo,))
    except Exception as e:
        ok('%s: Oud FO inlezen (VBA) draait' % naam, False, str(e)[:200])
        return doc
    jaren_als_getal(doc)
    doc.calculateAll()
    uit = os.path.join(lo.tmp, 'vba_%s.xlsx' % re.sub(r'\W', '', naam))
    import uno
    doc.storeToURL(uno.systemPathToFileUrl(uit), (lo.pv('FilterName', 'Calc MS Excel 2007 XML'),))
    import openpyxl
    wb = openpyxl.load_workbook(uit, data_only=True)
    import io
    import contextlib
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        n = controleer.vergelijk_cashflow(wb, fo, 0)
    omzet = 'omzet' in buf.getvalue()
    ok('%s: Oud FO inlezen (VBA) → cashflow gelijk aan het oude FO' % naam, not omzet and n <= (4 if python_doc else 0),
       buf.getvalue().strip().splitlines()[-1] if buf.getvalue().strip() else '')
    if python_doc is not None:
        a, b = invoer_rijen(doc), invoer_rijen(python_doc)
        verschil = [k for k in set(a) | set(b) if k not in a or k not in b or any(abs(x - y) > 1 for x, y in zip(a[k], b[k]))]
        ok('%s: VBA-omzetting = Python-omzetting (kosten, omzet, cashflow per periode)' % naam, not verschil,
           '%d periodes vergeleken' % len(a) + (', verschil in %s' % sorted(verschil)[:4] if verschil else ''))
    return doc


def t_nieuw_kwartaal(lo, doc):
    inv_voor = invoer_rijen(doc)
    h = doc.Sheets.getByName('FO - huidig')
    inhoud = lambda ws: [(lambda c: round(c.getValue(), 6) if c.getType().value == 'VALUE' else c.getString().strip())(
        ws.getCellByPosition(c, r)) for r in range(16, 60) for c in range(0, 21)]
    huidig_voor = inhoud(h)
    try:
        n = lo.macro(doc, 'NieuwKwartaalUitvoeren')
    except Exception as e:
        ok('Nieuw kwartaal starten draait', False, str(e)[:200])
        return
    doc.calculateAll()
    v = doc.Sheets.getByName('CF vorig en huidig')
    bewaard = {(v.getCellByPosition(1, r).getString(), v.getCellByPosition(2, r).getString()): v.getCellByPosition(6, r).getValue()
               for r in range(7, 67) if v.getCellByPosition(1, r).getString()}
    goed = all(abs(bewaard.get(k, 1e99) - x[2]) < 0.01 for k, x in inv_voor.items())
    ok('Nieuw kwartaal: vorige prognose = de cashflow van daarvoor', goed and n, '%s periodes bewaard' % n)
    fv = doc.Sheets.getByName('FO - vorig')
    vorig_na = inhoud(fv)
    afwijk = [i for i, (x, y) in enumerate(zip(vorig_na, huidig_voor)) if x != y and str(x) != str(y)]
    ok('Nieuw kwartaal: FO - vorig = het oude FO - huidig', not afwijk,
       '' if not afwijk else '%d cellen anders, bv. %s: %r ↔ %r' % (
           len(afwijk), '%s%d' % (S.L(afwijk[0] % 21 + 1), afwijk[0] // 21 + 17), huidig_voor[afwijk[0]], vorig_na[afwijk[0]]))
    leeg = doc.Sheets.getByName('FO - huidig').getCellByPosition(8, 20).getString() == '' and \
        doc.Sheets.getByName('FO - actuals').getCellByPosition(0, 1).getString() == ''
    ok('Nieuw kwartaal: FO - huidig en FO - actuals leeg voor de nieuwe exports', leeg)


def t_scenarios(doc):
    p, o, k = (doc.Sheets.getByName(n) for n in ('Aansluiting opbrengsten', 'CF - opbrengsten', 'CF - kosten'))

    def bereken():
        doc.calculateAll()

    # kans grondtermijn
    m_, d_ = doc.Sheets.getByName('Model'), doc.Sheets.getByName('Dashboard')

    def scenariolijn():                 # cumulatieve opbrengsten van de scenariolijn (Model AL) per periode
        import itertools
        return list(itertools.accumulate(waarde(m_, 'AL%d' % r) for r in range(8, 68)))
    bereken()
    zonder = scenariolijn()
    p.getCellRangeByName('Y13').setValue(0.5)
    bereken()
    extra, netto = waarde(o, 'D%d' % (S.KT + 2)), waarde(o, 'D%d' % (S.KT + 4))
    ok('kans grondtermijn 50%: extra grondtermijn eerder, netto 0 over de looptijd', extra > 0 and abs(netto) < 1,
       'extra %.0f, netto %.0f' % (extra, netto))
    met = scenariolijn()
    hoogste = max(b - a for a, b in zip(zonder, met))
    ok('kans grondtermijn in de scenariolijn (Dashboard Y27 = ja): eerder ontvangen, aan het eind gelijk, naam + kans grond',
       hoogste > 1 and abs(met[-1] - zonder[-1]) < 1 and tekst(m_, 'BY71').endswith('+ kans grond'),
       'hoogste stand extra %.0f, eind verschil %.0f, naam %r' % (hoogste, met[-1] - zonder[-1], tekst(m_, 'BY71')))
    d_.getCellRangeByName('Y27').setString('nee')
    bereken()
    uit = scenariolijn()
    ok('kans grondtermijn uit de scenariolijn (Dashboard Y27 = nee): scenariolijn als zonder kans',
       max(abs(a - b) for a, b in zip(zonder, uit)) < 1 and not tekst(m_, 'BY71').endswith('kans grond'))
    d_.getCellRangeByName('Y27').setString('ja')
    p.getCellRangeByName('Y13').setString('')
    # parkeren naar rato van de bouwtermijnen
    cel = 'D%d' % (S.p_aans(1) + 2 + S.P_REGELS.index('Opbrengsten parkeren'))
    p.getCellRangeByName(cel).setValue(1000000)
    bereken()
    som = waarde(o, '%s%d' % (S.CHK, S.OVR2))
    ok('parkeren €1 mln naar rato van de bouwtermijnen: volledig verdeeld', abs(som - 1000000) < 1, 'verdeeld %.0f' % som)
    p.getCellRangeByName(cel).setString('')
    # Voorstel met een zelf ingevuld kwartaal
    r = S.K_POST1
    oud = {c: k.getCellByPosition(c, r - 1).getFormula() for c in range(S.C_ACT - 1, S.C_QN)}
    methode = tekst(k, 'C%d' % r)
    k.getCellRangeByName('C%d' % r).setString('Voorstel')
    for c in range(S.C_Q1 - 1, S.C_QN):
        k.getCellByPosition(c, r - 1).setValue(0)
    k.getCellByPosition(S.C_Q1 - 1 + 3, r - 1).setValue(1500000)           # Q4 van het eerste jaar
    bereken()
    pew, verdeeld = waarde(k, 'D%d' % r), waarde(k, 'E%d' % r)
    ok('Voorstel (nullen + één zelf ingevuld kwartaal): verdeeld = PEW', abs(pew - verdeeld) < 10 and not tekst(k, 'A%d' % r),
       'PEW %.0f, verdeeld %.0f' % (pew, verdeeld))
    k.getCellRangeByName('C%d' % r).setString(methode)
    for c, f in oud.items():
        k.getCellByPosition(c, r - 1).setFormula(f)
    # verkoop vóór het kwartaal van de prijslijstexport telt niet mee
    c1 = p.getCellRangeByName('C1').getFormula()
    p.getCellRangeByName('C1').setFormula('=DATE(2026;11;4)')
    bereken()
    pct, status = waarde(o, '%s%d' % (S.CHK, S.tr(2) + 1)), tekst(o, '%s%d' % (S.STAT, S.tr(2) + 1))
    ok('verkoop vóór de exportdatum telt niet nog eens mee', pct < 0.999 and status.startswith('LET OP'),
       'verkocht %.1f%% · %s' % (pct * 100, status))
    p.getCellRangeByName('C1').setFormula(c1)
    # overlappende sleutels: elk bouwnummer bij één type
    c13 = tekst(p, 'C13')
    p.getCellRangeByName('C13').setString('*')
    bereken()
    a1, a2, tot = waarde(p, 'D13'), waarde(p, 'D14'), waarde(p, 'D%d' % (S.P_TYPE1 + S.NT))
    ok('overlappende sleutels (* en APP): geen dubbeltelling', (a1, a2, tot) == (53, 24, 77), '%d + %d = %d' % (a1, a2, tot))
    p.getCellRangeByName('C13').setString(c13)
    bereken()


def t_palet_werkbestand():
    """het werkbestand heeft hetzelfde palet als het PowerPoint-sjabloon (geen oude kleuren in grafieken en opmaak)"""
    import zipfile
    import pptx_v22
    oud = {}
    with zipfile.ZipFile(SJABLOON) as z:
        for n in z.namelist():
            if n.endswith('.xml') and n.startswith(('xl/charts/', 'xl/drawings/', 'xl/styles', 'xl/worksheets/sheet')):
                s = z.read(n).decode('utf-8', 'ignore').upper()
                for k in pptx_v22.PALET:
                    if re.search(r'(VAL|RGB)="(FF)?%s"' % k, s):
                        oud.setdefault(k, set()).add(n.split('/')[1])
    ok('werkbestand: zelfde palet als de slides (geen oude kleuren)', not oud,
       ', '.join('%s in %s' % (k, '/'.join(sorted(v))) for k, v in sorted(oud.items())))
    # Dashboard: vorige prognose geel, scenario rood, typekleuren (verkoop / transport en de legenda in Model)
    with zipfile.ZipFile(SJABLOON) as z:
        ch = [z.read(n).decode('utf-8') for n in z.namelist() if re.match(r'xl/charts/chart\d+\.xml$', n)]
    typen = [c for c in ch if 'g_vk_pct' in c and all(t in c for t in pptx_v22.TYPEN[:4])]
    vorig = [c for c in ch if re.search(r'Vorige prognose</c:v>.*?srgbClr val="%s"' % pptx_v22.VORIG_KL, c, re.S)]
    opd = [c for c in ch if 'g_vk_som_' in c and c.count('datalabelsRange>') == 4
           and ('PowerPoint!$IG$' in c or 'PowerPoint!$IH$' in c)]
    ok('Dashboard: opsomming boven de staaf en de percentagerij uit cellen (verkoop en transport)', len(opd) == 2,
       '%d grafieken' % len(opd))
    ok('Dashboard: vorige prognose geel, scenario rood, eigen kleuren per woningtype',
       typen and vorig and any('srgbClr val="%s"' % pptx_v22.SCEN_KL in c for c in vorig),
       '%d typegrafieken, %d met vorige prognose' % (len(typen), len(vorig)))


def t_zonder_prijslijst():
    """leeg sjabloon, vroeg in het project: alleen de SKO, het aantal, een termijnschema en een planning (geen
    prijslijst): de aansluiting sluit (PEW = SKO, geen tegenboeking) en de omzet komt in de cashflow"""
    import tempfile
    import openpyxl
    from xlsxpkg import Package, Sheet, SST
    tmp = tempfile.mkdtemp(prefix='zonderpl_')
    k, k2, d = (os.path.join(tmp, x) for x in ('k.xlsm', 'k2.xlsm', 'd.xlsx'))
    controleer.testkopie(SJABLOON, k)
    pkg = Package(k)
    sst = SST(pkg.x('xl/sharedStrings.xml'))

    def zet(blad, cellen):
        part = pkg.sheet_part(blad)
        sh = Sheet(pkg.x(part), sst)
        for ref, v in cellen:
            sh.set_value(ref, v)
        pkg.put(part, sh.t)
    a, q = S.p_aans(1), S.QCOLS
    zet('Aansluiting opbrengsten', [('B%d' % S.P_TYPE1, 'Rijwoning'), ('N%d' % S.P_TYPE1, 10), ('B%d' % (a + 2), 600000),
                                    ('B%d' % (a + 3), 2900000), ('L%d' % (a + 2), 'T1'), ('M%d' % (a + 2), 0.5),
                                    ('L%d' % (a + 3), 'T2'), ('M%d' % (a + 3), 0.5)])
    zet('CF - opbrengsten', [('%s%d' % (q[4], S.plan(1, 1)), 10), ('%s%d' % (q[6], S.plan(1, 2)), 10),
                             ('%s%d' % (q[2], S.vk_rij(1)), 10), ('%s%d' % (q[3], S.tt_rij(1)), 10)])
    sst.finish()
    pkg.put('xl/sharedStrings.xml', sst.t)
    pkg.save(k2)
    controleer.doorrekenen(k2, d)
    wb = openpyxl.load_workbook(d, data_only=True)
    p, o = wb['Aansluiting opbrengsten'], wb['CF - opbrengsten']
    regels = [[p['%s%d' % (c, r)].value for c in 'BDGHJ'] for r in (a + 2, a + 3)]
    tot = next((o['D%d' % r].value for r in range(1, 80) if str(o['B%d' % r].value or '').startswith('Totaal opbrengsten type 1')), None)
    ok('zonder prijslijst: de SKO loopt door naar de aansluiting (PEW = SKO, Corr 0) en de cashflow',
       all(abs(b - d_) < 1 and abs(g) < 1 and abs(h - b) < 1 and j == 'OK' for b, d_, g, h, j in regels)
       and tot is not None and abs(tot - 3500000) < 1 and 'geen prijslijst' in str(p['E%d' % a].value),
       'opbrengsten type 1 %s (SKO 3500000), kop: %s' % (tot, p['E%d' % a].value))


def t_gedeelde_formules():
    """het sjabloon bewaart doorgetrokken formules als gedeelde formule (kleiner, sneller openen); de opbouw klopt"""
    import zipfile
    from xlsxpkg import M, parse, split_ref, col2num
    n_kind = n_bron = groot = 0
    fout = []
    with zipfile.ZipFile(SJABLOON) as z:
        for n in z.namelist():
            if not re.match(r'xl/worksheets/sheet\d+\.xml$', n):
                continue
            b = z.read(n)
            groot += len(b)
            bron, kind = {}, []
            for f in parse(b).iter(M + 'f'):
                if f.get('t') != 'shared':
                    continue
                cel = f.getparent().get('r')
                if f.get('ref'):
                    if f.get('si') in bron or not f.text:
                        fout.append('%s %s: bron dubbel of zonder tekst' % (n, cel))
                    bron[f.get('si')] = f.get('ref')
                else:
                    kind.append((f.get('si'), cel))
            for si, cel in kind:
                if si not in bron:
                    fout.append('%s %s: geen bron' % (n, cel))
                    continue
                (k1, r1), (k2, r2) = (split_ref(x) for x in bron[si].split(':'))
                k, r = split_ref(cel)
                if not (col2num(k1) <= col2num(k) <= col2num(k2) and r1 <= r <= r2):
                    fout.append('%s %s: buiten het bereik %s' % (n, cel, bron[si]))
            n_kind += len(kind)
            n_bron += len(bron)
    ok('sjabloon: doorgetrokken formules gedeeld opgeslagen (kleiner bestand, sneller openen), opbouw klopt',
       n_kind > 50000 and not fout and groot < 14e6 and os.path.getsize(SJABLOON) < 1.7e6,
       '%d cellen in %d reeksen, werkbladen %.1f MB, bestand %.2f MB%s' % (n_kind, n_bron, groot / 1e6, os.path.getsize(SJABLOON) / 1e6,
                                                                        '; ' + '; '.join(fout[:3]) if fout else ''))


def t_dianummers():
    """de teksten in het werkbestand noemen geen verouderde dianummers (sinds v22 is dia 4 nieuw)"""
    import openpyxl                      # de cellen zelf (in sharedStrings blijven vervangen teksten ongebruikt staan)
    wb = openpyxl.load_workbook(SJABLOON, read_only=True)
    sst = '\n'.join(str(c) for ws in wb.worksheets for rij in ws.iter_rows(values_only=True) for c in rij
                    if isinstance(c, str))
    oud = [t for t in ('dia 1 en 10', 'dia 4 en Update PEW', 'kopie van dia 8', 'kopie van dia 9', 'Dashboard en dia 5', 'DIA 10') if t in sst]
    ok('werkbestand: geen verouderde dianummers in de teksten', not oud, ', '.join(oud))


def t_bereik_en_totaal():
    """grafiekbereik van het Dashboard werkt door in de blokken voor PowerPoint; totaalrij verkoop/transport; afgesloten
    kwartaal = huidig kwartaal - 1"""
    import tempfile
    import openpyxl
    from openpyxl.utils import column_index_from_string as ci
    from xlsxpkg import Package, Sheet, SST
    tmp = tempfile.mkdtemp(prefix='bereik_')
    k, k2, d = (os.path.join(tmp, x) for x in ('k.xlsm', 'k2.xlsm', 'd.xlsx'))
    controleer.testkopie(DEMO_WB, k)
    pkg = Package(k)
    part, sst = pkg.sheet_part('Dashboard'), SST(pkg.x('xl/sharedStrings.xml'))
    sh = Sheet(pkg.x(part), sst)
    for ref, v in (('W89', 2026), ('X89', 3), ('W90', 2027), ('X90', 4)):
        sh.set_value(ref, v)
    pkg.put(part, sh.t)
    sst.finish()
    pkg.put('xl/sharedStrings.xml', sst.t)
    pkg.save(k2)
    controleer.doorrekenen(k2, d)
    wb = openpyxl.load_workbook(d, data_only=True)
    m, pp, st = wb['Model'], wb['PowerPoint'], wb['Start']
    hh = [pp.cell(r, ci('HH')).value for r in range(8, 68)]
    gevuld = [x for x in hh if x not in (None, '')]
    ok('grafiekbereik Q3 2026 t/m Q4 2027: het cashflowblok voor PowerPoint heeft die 6 kwartalen',
       m['BY338'].value == 6 and len(gevuld) == 6 and gevuld[0] == 'Q3' and gevuld[-1] == 'Q4', '%s rijen: %s' % (len(gevuld), gevuld))
    tot = [pp.cell(8, ci('ER') + 1 + i).value for i in range(10)]
    som = sum(x for x in tot if isinstance(x, (int, float)))
    ok('verkoop en transport: totaalrij = alle woningen per type', pp['ER8'].value == 'Totaal' and som == m['BY39'].value > 0
       and m['BY339'].value == m['BY132'].value + 1, 'totaal %s' % som)
    verw = ' + '.join('%g' % x for x in tot if isinstance(x, (int, float)) and x > 0)
    rij = [(pp['FR%d' % r].value, pp['IG%d' % r].value, m['BO%d' % (r - 1)].value) for r in range(9, 69)
           if isinstance(pp['FS%d' % r].value, (int, float))]
    ok('verkoop en transport: percentagerij (vaste hoogte, tekst = cumulatief %)', rij and all(
        t == '%d%%' % round(p * 100) for p, t, _ in rij) and all(abs(pp['FS%d' % r].value - 1.12) < 1e-9 for r in range(9, 69)
                                                                   if isinstance(pp['FS%d' % r].value, (int, float))),
       '%d kwartalen, bijv. %s' % (len(rij), [t for _, t, _ in rij][:4]))
    ok('verkoop en transport: opsomming boven de staaf (Dashboard), bijv. totaalrij', pp['IC8'].value == verw
       and pp['IE8'].value == som, '%r (verwacht %r), hoogte %s' % (pp['IC8'].value, verw, pp['IE8'].value))
    # tijdassen: bouwtermijnen t/m het jaar van de laatste termijn; verkoop/transport t/m Q4 van het laatste transport
    laatste = max([x for x in (m.cell(r, ci('GK')).value for r in range(8, 141)) if isinstance(x, (int, float)) and x < 99999] or [0])
    ok('bouwtermijnen: de tijdas stopt aan het eind van het jaar van de laatste termijn', laatste > 0
       and m['BY99'].value == ((laatste - 1) // 4 + 1) * 4 + 1, 'laatste termijn %s, as tot %s' % (laatste, m['BY99'].value))
    vp = [pp.cell(r, ci('ER')).value for r in range(9, 69)]
    vp = [x for x in vp if x not in (None, '')]
    ok('verkoop en transport: de as stopt in Q4 van het jaar van het laatste transport', vp and str(vp[-1]).startswith('Q4')
       and m['BY340'].value <= m['BY132'].value, '%d kwartalen, laatste %s' % (len(vp), vp[-1] if vp else ''))
    # tabblad Aansluiting opbrengsten (was Prijslijst); CF-overzicht onderaan beide CF-tabbladen
    ok('tabblad Prijslijst heet Aansluiting opbrengsten (Prijslijst dump blijft)', 'Aansluiting opbrengsten' in wb.sheetnames
       and 'Prijslijst' not in wb.sheetnames and 'Prijslijst dump' in wb.sheetnames)
    inv = wb['Invoer']
    tot = next(r for r in range(8, 120) if inv.cell(r, 3).value == 'Totaal')
    goed = []
    for blad in ('CF - opbrengsten', 'CF - kosten'):
        ws = wb[blad]
        R = next(r for r in range(2, ws.max_row + 1) if str(ws.cell(r, 2).value or '').startswith('CASHFLOW · ZO WORDT'))
        o_, k_, c_ = (ws.cell(R + i, 4).value for i in (5, 6, 8))
        e_, f_ = ws.cell(R + 5, 5).value, ws.cell(R + 5, 6).value
        goed.append(abs(o_ - inv.cell(tot, 6).value) < 1 and abs(k_ - inv.cell(tot, 4).value) < 1
                    and abs(c_ - m['BY20'].value) < 1 and abs(e_ + f_ - o_) < 1 and ws.cell(R + 10, 4).value == 'OK')
    ok('CF-overzicht onderaan beide CF-tabbladen: totalen = Invoer, eindsaldo = Model, actuals + prognose = totaal, PEW OK',
       all(goed), str(goed))
    # RESULTAAT: delta = huidig - vorig (een stijging is positief), % in procentpunten; dia 2: t/m het jaar van de laatste cijfers
    p2 = wb['2. Ontw PEW']
    ok('RESULTAAT: delta huidig − vorig en % in procentpunten', abs(p2['H29'].value - (p2['G29'].value - p2['F29'].value)) < 0.01
       and abs(p2['H30'].value - (p2['G30'].value - p2['F30'].value) * 100) < 0.001, '%s / %s' % (p2['H29'].value, p2['H30'].value))
    kc = [pp.cell(r, ci('GZ')).value for r in range(8, 68)]
    ok('grafiek dia 2: t/m Q4 van het jaar van de laatste opbrengsten of kosten', 0 < m['BY341'].value <= m['BY8'].value
       and sum(1 for x in kc if isinstance(x, (int, float))) == m['BY341'].value, 'BY341 = %s' % m['BY341'].value)
    # verkoop en transport: één tabel per type; t/m nu uit de Salesforce-status
    import standaard_fo as S
    ob = wb['CF - opbrengsten']
    dump = wb['Prijslijst dump']
    sts = [str(dump.cell(r, 7).value or '').strip().lower() for r in range(2, dump.max_row + 1)]
    sf_t = sum(1 for x in sts if x.startswith('getransporteer'))          # Getransporteerd of Getransporteert
    sf_v = sf_t + sts.count('verkocht')                                   # In optie / Te koop tellen niet als verkocht
    tv, tt = ob['D%d' % S.VK_TOT].value, ob['D%d' % S.TT_TOT].value
    ok('verkoop en transport: één tabel per type, t/m nu uit Salesforce (verkocht en getransporteerd)',
       str(ob['B%d' % S.VK].value).startswith('VERKOOP EN TRANSPORT') and str(ob['B%d' % S.vk_rij(1)].value).endswith('verkocht')
       and 'getransporteerd' in str(ob['B%d' % S.tt_rij(1)].value) and tv == sf_v and tt == sf_t
       and ob['G%d' % S.TT_TOT].value == tt, 'verkocht %s (SF %s), getransporteerd %s (SF %s)' % (tv, sf_v, tt, sf_t))
    # 2B. Ontw PEW: A+B vorig/huidig = 2. Ontw PEW
    pw, pb = wb['2. Ontw PEW'], wb['2B. Ontw PEW']
    ok('2B. Ontw PEW: vorig, huidig en FD gelijk aan 2. Ontw PEW', all(abs((pb['%s20' % a].value or 0) - (pw['%s20' % b].value or 0)) < 0.01
       for a, b in (('E', 'E'), ('H', 'F'), ('I', 'G'))) and abs(pb['I20'].value) > 0, '%s' % pb['I20'].value)
    import zipfile
    with zipfile.ZipFile(SJABLOON) as z:
        wbx = z.read('xl/workbook.xml').decode('utf-8')
    namen = re.findall(r'<definedName name="(g_(?:vk_\w+|v\d+|t\d+))"[^>]*>([^<]*)<', wbx)
    ok('Dashboard: de grafieken verkoop en transport lezen het blok met de totaalrij', len(namen) >= 30
       and all(t.startswith('OFFSET(PowerPoint!$') and 'BY$339' in t for _, t in namen), '%d namen' % len(namen))
    ok('Start: afgesloten t/m = het kwartaal voor het huidige kwartaal', st['F8'].value and st['F12'].value == st['F8'].value - 1,
       '%s / %s' % (st['E8'].value, st['E12'].value))


def t_export(lo, export=None):
    """de ruwe SAP-export (een rij lager, een kolom naar links, bedragen als Nederlandse tekst, nullen leeg) wordt bij het
    inlezen omgezet naar de vaste indeling; een export die al goed staat blijft ongewijzigd"""
    kopie = os.path.join(lo.tmp, 'export.xlsm')
    controleer.testkopie(DEMO_WB, kopie)
    doc = lo.open(kopie, macros=True)
    try:
        sh = doc.Sheets.getByName('FO - huidig')
        pew = doc.Sheets.getByName('2. Ontw PEW')
        doc.calculateAll()
        n = 0
        for r in range(400):
            if any(sh.getCellByPosition(c, r).getString() for c in range(21)):
                n = r + 1
        bereik = sh.getCellRangeByPosition(0, 0, 20, n + 1)
        orig = [list(rij) for rij in bereik.getDataArray()]
        pew0 = [pew.getCellByPosition(6, r).getValue() for r in (19, 26)]
        # al goed: ongewijzigd
        bereik.setDataArray(tuple(tuple(r) for r in orig))
        m0 = lo.macro(doc, 'ExportNormaliserenTest', ('FO - huidig',))
        na0 = [list(rij) for rij in bereik.getDataArray()]
        ok('export in de vaste indeling: blijft ongewijzigd', not m0 and na0 == orig, m0 or '')

        def nl(v):
            t = ('%.2f' % abs(v)) if abs(v - round(v)) > 0.004 else '%d' % round(abs(v))
            heel, _, dec = t.partition('.')
            heel = '.'.join([heel[max(0, i - 3):i] for i in range(len(heel), 0, -3)][::-1])
            return ('-' if v < 0 else '') + heel + (',' + dec if dec else '')
        ruw = [[''] * 21 for _ in range(n + 2)]
        for r in range(n):
            for c in range(21):
                v = orig[r][c]
                if v == '' or (r > 14 and c == 0):
                    continue
                if r > 14 and 10 <= c <= 18 and isinstance(v, float):
                    v = nl(v) if abs(v) > 0.004 else ''
                elif isinstance(v, float) and r < 14:
                    v = '%d' % v
                c2 = 0 if (r <= 14 and c == 0) else c - 1
                ruw[r + 1][c2] = v
        bereik.setDataArray(tuple(tuple(r) for r in ruw))
        m1 = lo.macro(doc, 'ExportNormaliserenTest', ('FO - huidig',))
        doc.calculateAll()
        na = [list(rij) for rij in bereik.getDataArray()]
        fout = []
        for r in range(n):
            for c in range(21):
                a, b = orig[r][c], na[r][c]
                if r <= 14 and c == 1:
                    continue
                gelijk = (abs((a or 0) - (b or 0)) < 0.006 if isinstance(a, float) or isinstance(b, float) and a == '' else a == b)
                if not gelijk:
                    fout.append('%s%d: %r -> %r' % (chr(65 + c), r + 1, a, b))
        pew1 = [pew.getCellByPosition(6, r).getValue() for r in (19, 26)]
        ok('ruwe SAP-export: omgezet naar de vaste indeling (kopregel rij 15, periode E4, bedragen als getal)',
           m1 and not fout and sh.getCellByPosition(0, 14).getString().startswith('Business'), (m1 or '') + ' ' + '; '.join(fout[:4]))
        ok('ruwe SAP-export: PEW op 2. Ontw PEW gelijk aan voor het omzetten', all(abs(x - y) < 0.01 for x, y in zip(pew0, pew1))
           and abs(pew0[0]) > 0, '%s / %s' % (pew0, pew1))
        if export:                                   # een eigen export (blijft lokaal): totaal van de export = PEW
            import openpyxl
            ws = openpyxl.load_workbook(export, data_only=True).worksheets[0]
            data = [[('' if c is None else c) for c in rij[:21]] + [''] * (21 - len(rij[:21])) for rij in ws.iter_rows(values_only=True)]
            leeg = sh.getCellRangeByPosition(0, 0, 20, max(n, len(data)) + 5)
            leeg.clearContents(1023)
            sh.getCellRangeByPosition(0, 0, 20, len(data) - 1).setDataArray(tuple(tuple(r) for r in data))
            m2 = lo.macro(doc, 'ExportNormaliserenTest', ('FO - huidig',))
            doc.calculateAll()
            ok('eigen export: omgezet', bool(m2) and sh.getCellByPosition(0, 14).getString().startswith('Business'), m2 or '')
            print('        periode %s · regel 17: %s' % (sh.getCellByPosition(4, 3).getString(),
                  [sh.getCellByPosition(c, 16).getString() for c in (1, 10, 12, 14)]))
            print('        2. Ontw PEW huidig: kosten %.0f · opbrengsten %.0f · resultaat %.0f' % tuple(
                pew.getCellByPosition(6, r).getValue() for r in (19, 26, 28)))
    finally:
        doc.close(True)


def t_pptx():
    from pptx import Presentation
    namen = set()

    def alle(shapes):
        for sh in shapes:
            namen.add(sh.name)
            if sh.shape_type == 6:                  # groep
                alle(sh.shapes)
    pres = Presentation(PPTX)
    for slide in pres.slides:
        alle(slide.shapes)
    for indeling in pres.slide_layouts:            # de macro zoekt ook in de dia-indelingen (voettekst)
        alle(indeling.shapes)
    vba = open(os.path.join(ROOT, 'vba', 'CashflowNaarPowerPoint.bas'), encoding='cp1252').read()
    fow = open(os.path.join(ROOT, 'vba', 'FOWerkbestand.bas'), encoding='cp1252').read()
    npp = vba[vba.index('Public Sub NaarPowerPoint()'):]
    npp = npp[:npp.index('\nEnd Sub')]
    vgv = vba[vba.index('Private Sub VulGrafiekVorm('):]
    vgv = vgv[:vgv.index('\nEnd Sub')]
    import glob
    lint = {os.path.basename(p_): vba_lint(open(p_, encoding='cp1252').read())
            for p_ in sorted(glob.glob(os.path.join(ROOT, 'vba', '*.bas')) + glob.glob(os.path.join(ROOT, 'vba', '*.cls')))}
    ok('VBA: geen toewijzing aan een andere functie, elke Sub/Function netjes afgesloten (controle die LibreOffice mist)',
       not any(lint.values()), '; '.join('%s %s' % (k, v[:3]) for k, v in lint.items() if v) or '%d modules' % len(lint))
    zv = vba[vba.index('Private Function ZoekVorm('):]
    zv = zv[:zv.index('\nEnd Function')]
    ok('Naar PowerPoint: vormen rechtstreeks op naam, reeksbereik in één keer met terugval, tijd per grafiek',
       's.Shapes(naam)' in zv and 'shp.Name = naam' not in zv and 'ZetReeksFormule(reeks, fReeks, kol, n) Then ZetBereik' in vgv
       and 'NoteerGrafiekTijd vorm, tGraf' in vgv and 'Per grafiek (s)' in vba)
    ok('Naar PowerPoint sneller: tijden in de eindmelding, PowerPoint geminimaliseerd, vormlijst, eerst (uit) weg',
       'TijdOverzicht()' in npp and 'WindowState = 2' in npp and 'BouwVormIndex pres' in npp
       and 'ScreenUpdating = False' in npp and vgv.index('VerwijderUitgezetteReeksen ch') < vgv.index('ZetBereik reeks')
       and 'ClearContents' not in vgv and 'Application.CalculateFull\n    CashflowNaarPowerPoint' not in fow
       and 'CashflowNaarPowerPoint.RekenTijd' in fow)
    nodig = set()
    for f in ('Grafieken', 'Tabellen'):
        body = vba[vba.index('Private Function %s()' % f):]
        body = body[:body.index('End Function')]
        nodig |= set(re.findall(r'Array\("([A-Z0-9_]+)"', body))
    nodig.discard('VT_TABEL')                       # sinds v20 niet meer in het sjabloon (de macro slaat hem over)
    import openpyxl
    pp = openpyxl.load_workbook(SJABLOON, read_only=True)['PowerPoint']
    for row in pp.iter_rows(min_row=8, max_row=320, min_col=3, max_col=6, values_only=True):
        if row[0] and isinstance(row[3], str) and re.fullmatch(r'[A-Z0-9_]+', row[3].strip()):
            nodig.add(row[3].strip())
    mist = sorted(nodig - namen)
    ok('PowerPoint-sjabloon: alle vormen die de macro vult zijn er', not mist, '%d vormen' % len(nodig) +
       (', mist: %s' % ', '.join(mist[:8]) if mist else ''))
    # v22: tabellen, palet en lettergrootte
    ok('PowerPoint-sjabloon: SC_TABEL is een tabel; de dia Update PEW heeft het vak PEW_BEELD (geen tabel, geen top 3)',
       any(sh.name == 'SC_TABEL' and sh.has_table for sl in pres.slides for sh in sl.shapes) and 'PEW_BEELD' in namen
       and 'PEW_TOP3' not in namen and 'PEW_TABEL' not in namen)
    rood = [sh.name for sl in pres.slides for sh in sl.shapes if sh.has_chart and sh.name in ('VP1_GRAFIEK', 'VP2_GRAFIEK')
            and sh.chart._chartSpace.xpath('.//c:ser[contains(string(c:tx), "umulatief %")]//a:srgbClr[@val="A16207"]')]
    ok('verkoop en transport: percentages bij de cumulatieve lijn donkergeel op beide grafieken', len(rood) == 2, ', '.join(rood))
    import pptx_v22 as PV
    kl = {}
    for sl in pres.slides:
        for sh in sl.shapes:
            if sh.has_chart:
                for ser in sh.chart._chartSpace.iter(PV.C + 'ser'):
                    sp = ser.find(PV.C + 'spPr')
                    c = sp.find('.//' + PV.A + 'srgbClr') if sp is not None else None
                    kl.setdefault((sh.name, PV._naam(ser).lower()), c.get('val') if c is not None else None)
    vorig = {k: v for k, v in kl.items() if k[1] == 'vorige prognose'}
    scen = {k: v for k, v in kl.items() if k[1] == 'scenario'}
    vp = [kl.get(('VP1_GRAFIEK', n)) for n in ('rijwoning', 'tweekapper', 'vrijstaand', 'appartement')]
    op = []
    for sl in pres.slides:
        for sh in sl.shapes:
            if sh.has_chart and sh.name in ('VP1_GRAFIEK', 'VP2_GRAFIEK'):
                cs_ = sh.chart._chartSpace
                ser = {PV._naam(x_): x_ for x_ in cs_.iter(PV.C + 'ser')}
                o, lb = ser.get('Opsomming'), ser.get('Cumulatief % (label)')
                mk = [x_ for x_ in cs_.iter(PV.C + 'ser') if PV._marker_reeks(x_)]
                if (o is not None and o.find('.//' + PV.C + 'dLblPos').get('val') == 't' and lb is not None
                        and lb.find('.//' + PV.C + 'dLblPos').get('val') == 'ctr' and mk
                        and mk[0].find(PV.C + 'marker/' + PV.C + 'symbol').get('val') == 'none'):
                    op.append(sh.name)
    ok('verkoop en transport (dia): opsomming boven de staaf, percentages in een rij bovenin, geen bolletje op de markering',
       len(op) == 2, ', '.join(op))
    ok('kleuren: vorige prognose geel, scenario rood, eigen kleuren per woningtype (dia)',
       vorig and set(vorig.values()) == {PV.VORIG_KL} and scen and set(scen.values()) == {PV.SCEN_KL} and vp == PV.TYPEN[:4],
       'vorige %s, scenario %s, types %s' % (sorted(set(vorig.values())), sorted(set(scen.values())), vp))
    import pptx_v22
    oud_palet, klein = set(), []
    for i, sl in enumerate(pres.slides, 1):
        for sh in sl.shapes:
            delen = [sh._element] + ([sh.chart._chartSpace] if sh.has_chart else [])
            for d in delen:
                oud_palet |= {c.get('val').upper() for c in d.iter(pptx_v22.A + 'srgbClr')} & set(pptx_v22.PALET)
                if sh.name.startswith('Slide Number') or sh.name == 'BT_GRAFIEK':
                    continue
                for tag in ('rPr', 'defRPr', 'endParaRPr'):
                    klein += ['dia %d %s %s' % (i, sh.name, r.get('sz')) for r in d.iter(pptx_v22.A + tag)
                              if r.get('sz') and int(r.get('sz')) < pptx_v22.MIN_SZ]
    ok('PowerPoint-sjabloon: alleen kleuren uit het werkboekpalet', not oud_palet, ', '.join(sorted(oud_palet)))
    ok('PowerPoint-sjabloon: geen tekst onder 8 pt', not klein, ', '.join(klein[:5]))
    vervallen = sorted(n for n in namen if n.startswith(('VR_', 'VO_')))
    ok('PowerPoint-sjabloon: 9 dia\'s, zonder vorige-huidig-realisatie en van verkoop naar omzet', len(pres.slides) == 9
       and not vervallen, ', '.join(vervallen[:5]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--geen-bouw', action='store_true')
    ap.add_argument('--fo', help='een eigen oud FO om ook te testen (blijft lokaal)')
    ap.add_argument('--export', help='een eigen SAP-export voor FO - huidig om te testen (blijft lokaal)')
    ap.add_argument('--alleen-export', action='store_true', help='alleen de exporttest (snel)')
    a = ap.parse_args()
    if not a.geen_bouw:
        print('== Bouwen')
        code, uit = run([sys.executable, 'tools/bouw.py'])
        ok('bouwen', code == 0, uit.strip().splitlines()[-1] if uit.strip() else '')
        if code:
            print(uit[-2000:])
            return 1
    if a.alleen_export:
        lo = LO()
        try:
            t_export(lo, a.export)
        finally:
            lo.stop()
        return 1 if [n for n, g, _ in uitslag if not g] else 0
    print('== 1-2 Controles en kringen')
    t_controleer(DEMO_WB, None, 'demo', 4)
    t_kringen()
    print('== 6 PowerPoint-sjabloon')
    t_pptx()
    t_palet_werkbestand()
    t_gedeelde_formules()
    t_zonder_prijslijst()
    t_dianummers()
    t_bereik_en_totaal()
    print('== 3-5 Met LibreOffice: macro\'s en scenario\'s')
    lo = LO()
    try:
        t_export(lo, a.export)
        kopie = os.path.join(lo.tmp, 'demo.xlsm')
        controleer.testkopie(DEMO_WB, kopie)
        demo = lo.open(kopie)
        jaren_als_getal(demo)
        demo.calculateAll()
        vba_doc = t_vba(lo, S.DEMO, 'demo', python_doc=demo)
        t_nieuw_kwartaal(lo, vba_doc)
        try:                                       # de PowerPoint-module moet compileren (de knop zelf vraagt PowerPoint)
            sp = vba_doc.getScriptProvider()
            sp.getScript('vnd.sun.star.script:VBAProject.CashflowNaarPowerPoint.KnoppenControleren'
                         '?language=Basic&location=document').invoke((), (), ())
            ok('PowerPoint-module compileert en de knoppen worden gezet', True)
        except Exception as e:
            ok('PowerPoint-module compileert en de knoppen worden gezet', False, str(e)[:200])
        vba_doc.close(True)
        t_scenarios(demo)
        demo.close(True)
        if a.fo:
            eigen = os.path.join(lo.tmp, 'eigen.xlsm')
            code, uit = run([sys.executable, 'tools/bouw.py', '--oud', a.fo, '--project', 'Test', '--out', eigen])
            ok('eigen FO: omzetten (Python)', code == 0)
            if code == 0:
                t_controleer(eigen, a.fo, 'eigen FO', 0, streng=False)
            d = t_vba(lo, os.path.abspath(a.fo), 'eigen FO')
            d.close(True)
    finally:
        lo.stop()
    fout = [n for n, g, _ in uitslag if not g]
    print('\n%d tests, %d fout' % (len(uitslag), len(fout)))
    for n in fout:
        print('  FOUT: ' + n)
    return 1 if fout else 0


if __name__ == '__main__':
    sys.exit(main())
