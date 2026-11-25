"""FO overzetten (output/FO_overzetten.xlsm): een klein los bestand met een macro die werkbestanden van een eerdere versie
van het FO-werkbestand in het nieuwste sjabloon zet (per bestand een kopie van het sjabloon, daarin FO overnemen).

    python tools/overzetten.py [uit.xlsm]
"""
from __future__ import annotations

import os
import sys
import tempfile

from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
import vba_project  # noqa: E402
from xlsxpkg import Package, rels_name  # noqa: E402

UIT = os.path.join(ROOT, 'output', 'FO_overzetten.xlsm')
SCENARIO = os.path.join(ROOT, 'input', 'Cashflow_scenario_leeg_3.xlsm')
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
KNOPPEN = [('B8:C9', 'btnSjabloon', '1. Nieuwste sjabloon kiezen...', 'modOverzetten.SjabloonKiezen', '5B6169'),
           ('D8:E9', 'btnOverzetten', '2. Werkbestanden kiezen en overzetten...', 'modOverzetten.Overzetten', '111F39')]


def blad(pad):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Overzetten'
    ws.sheet_properties.codeName = 'shOverzetten'
    wb.code_name = 'ThisWorkbook'
    navy, geel = PatternFill('solid', fgColor='FF003366'), PatternFill('solid', fgColor='FFFFF2CC')
    F = 'Heijmans-Light'
    ws['B2'] = 'FO overzetten'
    ws['B2'].font = Font(name='Heijmans-BoldCon', size=16, color='FF003366')
    uitleg = ['Zet FO-werkbestanden van een eerdere versie van het sjabloon over naar het nieuwste sjabloon: de exports en '
              'alle invoer gaan mee, de formules en verbeteringen van het nieuwe sjabloon blijven.',
              '1. Kies het nieuwste sjabloon (FO_werkbestand_sjabloon.xlsm).   2. Kies een of meer werkbestanden. '
              'Per werkbestand komt er een nieuw bestand naast: <naam>_nieuw.xlsm. Het oude bestand verandert niet.']
    for i, t in enumerate(uitleg):
        ws['B%d' % (3 + i)] = t
        ws['B%d' % (3 + i)].font = Font(name=F, size=10, italic=i == 1, color='FF4B5563')
    ws['B6'] = 'Nieuwste sjabloon'
    ws['B6'].font = Font(name=F, size=10, bold=True)
    ws['C6'].fill = geel
    ws['C6'].font = Font(name=F, size=10)
    ws.merge_cells('C6:E6')
    ws['B10'] = 'Knoppen werken niet? Alt+F8 > Overzetten > Uitvoeren (macro\'s moeten aan staan).'
    ws['B10'].font = Font(name=F, size=9, italic=True, color='FF6B7280')
    for c, t in zip('BCDE', ['Oud werkbestand', 'Nieuw werkbestand', 'Status', 'Melding (cellen per tabblad)']):
        ws['%s11' % c] = t
        ws['%s11' % c].font = Font(name=F, size=10, bold=True, color='FFFFFFFF')
        ws['%s11' % c].fill = navy
    for c, w in zip('ABCDE', [2, 60, 60, 10, 90]):
        ws.column_dimensions[c].width = w
    for r in range(12, 200):
        ws['E%d' % r].alignment = Alignment(wrap_text=False)
    ws.sheet_view.showGridLines = False
    wb.save(pad)


def maak(uit=UIT):
    import build_werkbestand as B
    tmp = tempfile.mkdtemp(prefix='overzetten_')
    x = os.path.join(tmp, 'o.xlsx')
    blad(x)
    pkg = Package(x)
    # werkmap als macro-werkmap met een vbaProject
    ct = pkg.x('[Content_Types].xml')
    for o in ct.findall(CT + 'Override'):
        if o.get('PartName') == '/xl/workbook.xml':
            o.set('ContentType', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    etree.SubElement(ct, CT + 'Default', Extension='bin', ContentType='application/vnd.ms-office.vbaProject')
    rn = rels_name('xl/workbook.xml')
    rels = pkg.x(rn)
    ids = {r.get('Id') for r in rels}
    rid = next('rId%d' % i for i in range(1, 99) if 'rId%d' % i not in ids)
    etree.SubElement(rels, '{%s}Relationship' % PR, Id=rid, Target='vbaProject.bin',
                     Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject')
    pkg.put(rn, rels)
    vdir = os.path.join(ROOT, 'vba')
    lees = lambda n: open(os.path.join(vdir, n), 'rb').read()
    modules = [{'name': 'ThisWorkbook', 'document': True, 'code': vba_project.document_module_source('ThisWorkbook', True)},
               {'name': 'shOverzetten', 'document': True, 'code': vba_project.document_module_source('shOverzetten')},
               {'name': 'modOverzetten', 'document': False, 'code': lees('Overzetten.bas')}]
    sjab = os.path.join(tmp, 'vba.bin')
    open(sjab, 'wb').write(Package(SCENARIO).parts['xl/vbaProject.bin'])
    pkg.parts['xl/vbaProject.bin'] = vba_project.build_vba_project(sjab, modules)
    # knoppen als vaste vormen (zoals in het sjabloon)

    class _C:
        def __init__(self, p, ct_):
            self.dst, self.ct = p, ct_

        def next_name(self, pattern):
            i = 1
            while pattern % i in self.dst.parts:
                i += 1
            return pattern % i

        def add_override(self, part, ctype):
            etree.SubElement(self.ct, CT + 'Override', PartName='/' + part, ContentType=ctype)
    oud = (B.START_KNOPPEN, B.DASH_KNOPPEN)
    try:
        B.START_KNOPPEN, B.DASH_KNOPPEN = KNOPPEN, []
        part = pkg.sheet_part('Overzetten')
        n = B.vaste_knoppen(pkg, {'Start': part, 'Dashboard': part}, _C(pkg, ct))
    finally:
        B.START_KNOPPEN, B.DASH_KNOPPEN = oud
    pkg.put('[Content_Types].xml', ct)
    pkg.save(uit)
    print('FO overzetten: %s (%d knoppen)' % (uit, n))
    return uit


if __name__ == '__main__':
    maak(sys.argv[1] if len(sys.argv) > 1 else UIT)
