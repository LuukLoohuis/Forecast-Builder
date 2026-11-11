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
    p, o, k = (doc.Sheets.getByName(n) for n in ('Prijslijst', 'CF - opbrengsten', 'CF - kosten'))

    def bereken():
        doc.calculateAll()

    # kans grondtermijn
    p.getCellRangeByName('Y13').setValue(0.5)
    bereken()
    extra, netto = waarde(o, 'D%d' % (S.KT + 2)), waarde(o, 'D%d' % (S.KT + 4))
    ok('kans grondtermijn 50%: extra grondtermijn eerder, netto 0 over de looptijd', extra > 0 and abs(netto) < 1,
       'extra %.0f, netto %.0f' % (extra, netto))
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


def t_dianummers():
    """de teksten in het werkbestand noemen geen verouderde dianummers (sinds v22 is dia 4 nieuw)"""
    import openpyxl                      # de cellen zelf (in sharedStrings blijven vervangen teksten ongebruikt staan)
    wb = openpyxl.load_workbook(SJABLOON, read_only=True)
    sst = '\n'.join(str(c) for ws in wb.worksheets for rij in ws.iter_rows(values_only=True) for c in rij
                    if isinstance(c, str))
    oud = [t for t in ('dia 1 en 10', 'kopie van dia 8', 'kopie van dia 9', 'Dashboard en dia 5', 'titel dia 6') if t in sst]
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
    ok('Start: afgesloten t/m = het kwartaal voor het huidige kwartaal', st['F8'].value and st['F12'].value == st['F8'].value - 1,
       '%s / %s' % (st['E8'].value, st['E12'].value))


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
    tab = {sh.name: sh.has_table for sl in pres.slides for sh in sl.shapes if sh.name in ('VR_TABEL', 'PEW_TABEL', 'SC_TABEL')}
    ok('PowerPoint-sjabloon: VR_TABEL, PEW_TABEL en SC_TABEL zijn tabellen', len(tab) == 3 and all(tab.values()), str(tab))
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
    ok('PowerPoint-sjabloon: 11 dia\'s, dia 4 = vorige ↔ huidig ↔ realisatie', len(pres.slides) == 11 and
       any(sh.name == 'VR_TABEL' for sh in pres.slides[3].shapes))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--geen-bouw', action='store_true')
    ap.add_argument('--fo', help='een eigen oud FO om ook te testen (blijft lokaal)')
    a = ap.parse_args()
    if not a.geen_bouw:
        print('== Bouwen')
        code, uit = run([sys.executable, 'tools/bouw.py'])
        ok('bouwen', code == 0, uit.strip().splitlines()[-1] if uit.strip() else '')
        if code:
            print(uit[-2000:])
            return 1
    print('== 1-2 Controles en kringen')
    t_controleer(DEMO_WB, None, 'demo', 4)
    t_kringen()
    print('== 6 PowerPoint-sjabloon')
    t_pptx()
    t_palet_werkbestand()
    t_dianummers()
    t_bereik_en_totaal()
    print('== 3-5 Met LibreOffice: macro\'s en scenario\'s')
    lo = LO()
    try:
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
