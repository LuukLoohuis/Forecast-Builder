"""Bouwt het FO-werkbestand: FO-onderlegger + cashflow-scenariomodel + nieuwe tabbladen in één .xlsm.

    python tools/build_werkbestand.py            (standaard: input/ -> output/)

Bronnen (input/):
  FO_Parkzoom_fictieve_demoversie.xlsx     FO-onderlegger (tabbladen, formules, opmaak)
  Cashflow_scenario_leeg_3.xlsm            scenariomodel met grafieken, knoppen en macro's
  Kwartaal_Template_cashflow_v20.pptx      PowerPoint-sjabloon
VBA (vba/): de broncode van alle modules; wordt als nieuw vbaProject.bin in het bestand gezet.
"""
from __future__ import annotations

import argparse
import copy
import os
import re
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import nieuwe_bladen  # noqa: E402
import vba_project  # noqa: E402
from xlsxpkg import (CT, CT_CHART, CT_CHARTCOLORS, CT_CHARTSTYLE, CT_DRAWING, CT_SHEET, M, PR, R, SST, T_DRAWING,  # noqa: E402
                     T_SHEET, Package, Sheet, Styles, col2num, num2col, parse, rels_name, relative, replace_in_formulas,
                     resolve, split_ref, tobytes)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ------------------------------------------------------------------------------------------------
#  indeling van het werkbestand: (naam, bron, naam in de bron, codenaam)
# ------------------------------------------------------------------------------------------------
SC_RENAME = {'CF kosten': 'Scenario kosten'}
LAYOUT = [
    ('Start', 'new', 'Start', 'shStart'),
    ('FO - vorig', 'fo', 'FO - vorig', 'shFOVorig'),
    ('FO - huidig', 'fo', 'FO - huidig', 'shFOHuidig'),
    ('FO - actuals', 'fo', 'FO - actuals', 'shFOActuals'),
    ('Prijslijst dump', 'fo', 'Prijslijst dump', 'shPrijslijstDump'),
    ('Prijslijst', 'fo', 'Prijslijst', 'shPrijslijst'),
    ('Aansluiting opbrengsten', 'fo', 'Aansluiting opbrengsten', 'shAansluiting'),
    ('Verkooptempo', 'fo', 'Verkooptempo', 'shVerkooptempo'),
    ('CF - opbrengsten', 'fo', 'CF - opbrengsten', 'shCFOpbrengsten'),
    ('CF - kosten', 'fo', 'CF - kosten', 'shCFKostenFO'),
    ('Prognose opbrengsten', 'new', 'Prognose opbrengsten', 'shPrognoseOpbr'),
    ('Prognose kosten', 'new', 'Prognose kosten', 'shPrognoseKosten'),
    ('1. Cashflow', 'fo', '1. Cashflow', 'shCashflow'),
    ('2. Ontw PEW', 'fo', '2. Ontw PEW', 'shOntwPEW'),
    ('3. Restant bdgt', 'fo', '3. Restant bdgt', 'shRestantBdgt'),
    ('CF - hulp VENA', 'fo', 'CF - hulp VENA', 'shVENA'),
    ('Dashboard', 'sc', 'Dashboard', 'shDashboard'),
    ("Scenario's", 'sc', "Scenario's", 'shScenarios'),
    ('Invoer', 'sc', 'Invoer', 'shInvoer'),
    ('Woningtypes', 'sc', 'Woningtypes', 'shWoningtypes'),
    ('Scenario kosten', 'sc', 'CF kosten', 'shCFKosten'),
    ('Model', 'sc', 'Model', 'shModel'),
    ('PowerPoint', 'sc', 'PowerPoint', 'shPowerPoint'),
    ('Notities FO gesprek', 'fo', 'Notities FO gesprek', 'shNotities'),
    ('Vorige prognose', 'new', 'Vorige prognose', 'shVorigePrognose'),
    ('Controle oud FO', 'new', 'Controle oud FO', 'shControleOudFO'),
    ('Behind the scens', 'fo', 'Behind the scens', 'shBehindScens'),
    ('Toelichting onderlegger', 'fo', 'Toelichting onderlegger', 'shToelichting'),
    ('1. Ontw businessc FD', 'fo', '1. Ontw businessc FD', 'shOntwBusinesscase'),
    ('5. K&R businessc', 'fo', '5. K&R businessc', 'shKRBusinesscase'),
]
REWIRE_SHEETS = ['1. Cashflow', 'CF - hulp VENA']
REWIRE = [("'CF - kosten'!", "'Prognose kosten'!"), ("'CF - opbrengsten'!", "'Prognose opbrengsten'!")]
VBA_STD = ['CashflowNaarPowerPoint', 'FOWerkbestand']
DEMO = {'project': 'Parkzoom', 'oplevering': (2028, 3)}


def note(msg):
    print('  ' + msg)


# ------------------------------------------------------------------------------------------------
#  werkblad overnemen uit een ander pakket
# ------------------------------------------------------------------------------------------------

class Copier:
    def __init__(self, dst: Package, styles: Styles, sst: SST):
        self.dst, self.styles, self.sst = dst, styles, sst
        self.ct = dst.x('[Content_Types].xml')

    def next_name(self, pattern):
        i = 1
        while pattern % i in self.dst.parts:
            i += 1
        return pattern % i

    def add_override(self, part, ctype):
        for o in self.ct.findall(CT + 'Override'):
            if o.get('PartName') == '/' + part:
                o.set('ContentType', ctype)
                return
        etree.SubElement(self.ct, CT + 'Override', PartName='/' + part, ContentType=ctype)

    def remove_override(self, part):
        for o in self.ct.findall(CT + 'Override'):
            if o.get('PartName') == '/' + part:
                self.ct.remove(o)

    def copy_sheet(self, src: Package, src_part: str, xf_map, dxf_map, src_si, codename):
        t = src.x(src_part)
        si_map = {}

        def map_si(i):
            if i not in si_map:
                si_map[i] = self.sst.add_si(src_si[i])
            return si_map[i]

        default_s = str(xf_map[0])
        for c in t.iter(M + 'c'):
            s = c.get('s')
            c.set('s', str(xf_map[int(s)]) if s is not None else default_s)
            if c.get('t') == 's':
                v = c.find(M + 'v')
                v.text = str(map_si(int(v.text)))
        for row in t.iter(M + 'row'):
            if row.get('s') is not None:
                row.set('s', str(xf_map[int(row.get('s'))]))
        for col in t.iter(M + 'col'):
            if col.get('style') is not None:
                col.set('style', str(xf_map[int(col.get('style'))]))
        for rule in t.iter(M + 'cfRule'):
            if rule.get('dxfId') is not None:
                rule.set('dxfId', str(dxf_map[int(rule.get('dxfId'))]))
        # paginainstellingen: printerinstellingen niet meenemen
        for ps in t.iter(M + 'pageSetup'):
            if R + 'id' in ps.attrib:
                del ps.attrib[R + 'id']
        for tag in ('legacyDrawing', 'legacyDrawingHF', 'tableParts'):
            for el in t.findall(M + tag):
                t.remove(el)
        for sv in t.iter(M + 'sheetView'):
            if 'tabSelected' in sv.attrib:
                del sv.attrib['tabSelected']
        pr = t.find(M + 'sheetPr')
        if pr is None:
            pr = etree.Element(M + 'sheetPr')
            t.insert(0, pr)
        pr.set('codeName', codename)

        new_part = self.next_name('xl/worksheets/sheet%d.xml')
        rels_out = []
        for el in t.findall(M + 'drawing'):
            rid = el.get(R + 'id')
            for i, ty, target, _ in src.rels(src_part):
                if i == rid and ty == T_DRAWING:
                    d_part = self.copy_drawing(src, resolve(src_part, target))
                    rels_out.append((rid, T_DRAWING, relative(new_part, d_part)))
        self.dst.put(new_part, t)
        if rels_out:
            rr = etree.Element(PR + 'Relationships', nsmap={None: PR[1:-1]})
            for rid, ty, target in rels_out:
                etree.SubElement(rr, PR + 'Relationship', Id=rid, Type=ty, Target=target)
            self.dst.put(rels_name(new_part), rr)
        self.add_override(new_part, CT_SHEET)
        return new_part

    def copy_drawing(self, src: Package, d_part: str):
        new_d = self.next_name('xl/drawings/drawing%d.xml')
        self.dst.parts[new_d] = src.parts[d_part]
        self.add_override(new_d, CT_DRAWING)
        rels = src.rels(d_part)
        if rels:
            rr = etree.Element(PR + 'Relationships', nsmap={None: PR[1:-1]})
            for rid, ty, target, mode in rels:
                p = resolve(d_part, target)
                if p.startswith('xl/charts/chart'):
                    new_c = self.next_name('xl/charts/chart%d.xml')
                    self.dst.parts[new_c] = src.parts[p]
                    self.add_override(new_c, CT_CHART)
                    crels = src.rels(p)
                    if crels:
                        cr = etree.Element(PR + 'Relationships', nsmap={None: PR[1:-1]})
                        for cid, cty, ctarget, _ in crels:
                            cp = resolve(p, ctarget)
                            base = os.path.basename(cp)
                            kind = 'style' if base.startswith('style') else 'colors'
                            new_s = self.next_name('xl/charts/' + kind + '%d.xml')
                            self.dst.parts[new_s] = src.parts[cp]
                            self.add_override(new_s, CT_CHARTSTYLE if kind == 'style' else CT_CHARTCOLORS)
                            etree.SubElement(cr, PR + 'Relationship', Id=cid, Type=cty, Target=relative(new_c, new_s))
                        self.dst.put(rels_name(new_c), cr)
                    etree.SubElement(rr, PR + 'Relationship', Id=rid, Type=ty, Target=relative(new_d, new_c))
                else:
                    raise ValueError('onbekend onderdeel in tekening: ' + p)
            self.dst.put(rels_name(new_d), rr)
        return new_d


# ------------------------------------------------------------------------------------------------
#  bouwen
# ------------------------------------------------------------------------------------------------

def oud_cashflow(fo_path):
    import openpyxl
    wb = openpyxl.load_workbook(fo_path, data_only=True)
    ws = wb['1. Cashflow']
    out = []
    for r in range(4, 56):
        if ws.cell(r, 2).value is None and ws.cell(r, 3).value is None:
            out.append({})
            continue
        if not isinstance(ws.cell(r, 2).value, (int, float)):
            out.append({})
            continue
        out.append({'D': ws.cell(r, 4).value, 'F': ws.cell(r, 6).value, 'H': ws.cell(r, 8).value})
    while out and not out[-1]:
        out.pop()
    return out


def merges_row57(fo: Package):
    part = fo.sheet_part('CF - kosten')
    t = fo.x(part)
    out = []
    for mc in t.iter(M + 'mergeCell'):
        ref = mc.get('ref')
        a, b = ref.split(':')
        if split_ref(a)[1] == 57 and split_ref(b)[1] == 57 and col2num(split_ref(a)[0]) >= 8:
            out.append(ref)
    return sorted(out, key=lambda r: col2num(split_ref(r.split(':')[0])[0]))


def extend_helpers(sh: Sheet, cols, first, last, row_to):
    """hulpkolommen (formules) doortrekken tot row_to, naar het voorbeeld van rij last"""
    for col in cols:
        c = sh.cell('%s%d' % (col, last), create=False)
        f = c.find(M + 'f') if c is not None else None
        if f is None:
            raise ValueError('geen formule in %s%d' % (col, last))
        text = f.text if f.text else None
        if text is None:          # gedeelde formule: de bron zoeken
            si = f.get('si')
            for ff in sh.t.iter(M + 'f'):
                if ff.get('si') == si and ff.get('ref'):
                    src_ref = ff.getparent().get('r')
                    text = shift_formula(ff.text, split_ref(src_ref)[1], last)
                    break
        style = c.get('s')
        for r in range(last + 1, row_to + 1):
            sh.set_formula('%s%d' % (col, r), shift_formula(text, last, r), style)


def shift_formula(text, row_from, row_to):
    """relatieve rijverwijzingen (zonder $) verschuiven, ook in kolombereiken als E:E niet"""
    d = row_to - row_from

    def rep(m):
        if m.group(1) is not None:      # tekst tussen aanhalingstekens of een bladnaam
            return m.group(0)
        col_abs, col, row_abs, row = m.group(2), m.group(3), m.group(4), m.group(5)
        if row_abs:
            return m.group(0)
        return '%s%s%s' % (col_abs, col, int(row) + d)
    return re.sub(r"(\"[^\"]*\"|'[^']*')|(?<![A-Za-z_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])", rep, text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--fo', default=os.path.join(ROOT, 'input', 'FO_Parkzoom_fictieve_demoversie.xlsx'))
    ap.add_argument('--scenario', default=os.path.join(ROOT, 'input', 'Cashflow_scenario_leeg_3.xlsm'))
    ap.add_argument('--pptx', default=os.path.join(ROOT, 'input', 'Kwartaal_Template_cashflow_v20.pptx'))
    ap.add_argument('--out', default=os.path.join(ROOT, 'output', 'FO_Parkzoom_werkbestand.xlsm'))
    ap.add_argument('--pptx-out', default=os.path.join(ROOT, 'output', 'Kwartaal_Template_cashflow_v21.pptx'))
    a = ap.parse_args()
    os.makedirs(os.path.dirname(a.out), exist_ok=True)

    print('Bronnen lezen')
    sc = Package(a.scenario)
    fo = Package(a.fo)
    tmp_new = a.out + '.nieuw.tmp'
    info = nieuwe_bladen.bouw(tmp_new, DEMO, merges_row57(fo), oud_cashflow(a.fo), os.path.basename(a.fo))
    new = Package(tmp_new)
    os.remove(tmp_new)

    dst = sc          # het scenariobestand is de basis: grafieken, knoppen en macro's blijven staan
    print('Thema en stijlen samenvoegen')
    dst.parts['xl/theme/theme1.xml'] = fo.parts['xl/theme/theme1.xml']   # Heijmans-thema (het FO leunt erop)
    styles = Styles(dst.x('xl/styles.xml'))
    maps = {'fo': styles.merge(fo.x('xl/styles.xml')), 'new': styles.merge(new.x('xl/styles.xml'))}
    sst = SST(dst.x('xl/sharedStrings.xml'))
    src_si = {k: (p.x('xl/sharedStrings.xml').findall(M + 'si') if 'xl/sharedStrings.xml' in p.parts else [])
              for k, p in (('fo', fo), ('new', new))}
    copier = Copier(dst, styles, sst)

    print('Werkbladen overnemen')
    sc_sheets = {n: (p, st) for n, p, st in sc.sheets()}
    sc_order = [n for n, _, _ in sc.sheets()]
    srcs = {'fo': fo, 'new': new}
    src_states = {k: {n: st for n, _, st in p.sheets()} for k, p in srcs.items()}
    final = []        # (naam, part, state)
    for name, origin, src_name, code in LAYOUT:
        if origin == 'sc':
            part, state = sc_sheets[src_name]
            t = dst.x(part)
            pr = t.find(M + 'sheetPr')
            assert pr is not None and pr.get('codeName') == code, (name, code)
            for sv in t.iter(M + 'sheetView'):
                sv.attrib.pop('tabSelected', None)
            dst.put(part, t)
        else:
            p = srcs[origin]
            xf_map, dxf_map = maps[origin]
            part = copier.copy_sheet(p, p.sheet_part(src_name), xf_map, dxf_map, src_si[origin], code)
            state = src_states[origin][src_name]
        final.append((name, part, state, code))
    note('%d tabbladen' % len(final))

    # ---- verwijzingen: CF kosten -> Scenario kosten; 1. Cashflow en VENA naar de prognosebladen -----------
    print('Formules aanpassen')
    for old, newname in SC_RENAME.items():
        n = 0
        for part in [p for _, p, _, _ in final] + [k for k in dst.parts if k.startswith('xl/charts/chart')]:
            t = dst.x(part)
            k = replace_in_formulas(t, "'%s'!" % old, "'%s'!" % newname)
            if k:
                dst.put(part, t)
                n += k
        note("'%s' -> '%s': %d formules" % (old, newname, n))
    part_of = {n: p for n, p, _, _ in final}
    for sname in REWIRE_SHEETS:
        t = dst.x(part_of[sname])
        n = sum(replace_in_formulas(t, o, nw) for o, nw in REWIRE)
        dst.put(part_of[sname], t)
        note('%s: %d formules naar Prognose kosten/opbrengsten' % (sname, n))

    # ---- cellen aanvullen --------------------------------------------------------------------------------
    print('Cellen aanvullen')

    def sheet(name):
        return Sheet(dst.x(part_of[name]), sst)

    def save(name, sh):
        dst.put(part_of[name], sh.t)

    sh = sheet('Dashboard')
    sh.set_formula('W10', 'IF(Start!$C$6="","",Start!$C$6)')
    save('Dashboard', sh)

    sh = sheet('PowerPoint')
    style = {c: sh.style_of('%s117' % c) for c in 'BCDEF'}
    d10 = info['dia10_row']
    rows = [
        (1, 'Cijfers per (tab Start)', '=Start!$C$11', 'COVER_DATUM'),
        (10, 'Cijfers per (tab Start)', '=Start!$C$11', 'PEW_DATUM'),
        (10, 'Meest recente FD', '=Start!$C$%d&""' % d10, 'PEW_FD_WAARDE'),
        (10, '  toelichting', '=Start!$D$%d&""' % d10, 'PEW_FD_SUB'),
        (10, 'Huidige PEW-basis', '=Start!$C$%d&""' % (d10 + 1), 'PEW_BASIS_WAARDE'),
        (10, '  toelichting', '=Start!$D$%d&""' % (d10 + 1), 'PEW_BASIS_SUB'),
        (10, 'Mogelijk resultaat', '=Start!$C$%d&""' % (d10 + 2), 'PEW_RES_WAARDE'),
        (10, '  toelichting', '=Start!$D$%d&""' % (d10 + 2), 'PEW_RES_SUB'),
    ]
    for i, (dia, onderdeel, f, vorm) in enumerate(rows):
        r = 118 + i
        assert sh.cell('C%d' % r, create=False) is None or not len(sh.cell('C%d' % r))
        sh.set_value('B%d' % r, dia, style['B'])
        sh.set_value('C%d' % r, onderdeel, style['C'])
        sh.set_formula('D%d' % r, f, style['D'])
        sh.set_value('F%d' % r, vorm, style['F'])
    save('PowerPoint', sh)

    for name, cols, first in (('FO - vorig', 'XY', 17), ('FO - huidig', 'XY', 17)):
        sh = sheet(name)
        last = max(int(split_ref(c.get('r'))[1]) for c in sh.t.iter(M + 'c')
                   if c.get('r').startswith('X') and c.find(M + 'f') is not None)
        extend_helpers(sh, cols, first, last, 1500)
        save(name, sh)
        note('%s: hulpkolommen X:Y t/m rij 1500 (was %d)' % (name, last))
    sh = sheet('FO - actuals')
    last = max(int(split_ref(c.get('r'))[1]) for c in sh.t.iter(M + 'c')
               if c.get('r').startswith('Q') and c.find(M + 'f') is not None)
    extend_helpers(sh, 'QR', 2, last, 6000)
    save('FO - actuals', sh)
    note('FO - actuals: hulpkolommen Q:R t/m rij 6000 (was %d)' % last)

    # ---- werkmap -------------------------------------------------------------------------------------------
    print('Werkmap samenstellen')
    wb = dst.x('xl/workbook.xml')
    sheets_el = wb.find(M + 'sheets')
    for s in list(sheets_el):
        sheets_el.remove(s)
    rels = etree.Element(PR + 'Relationships', nsmap={None: PR[1:-1]})
    for i, (name, part, state, code) in enumerate(final, 1):
        attrs = {'name': name, 'sheetId': str(i), R + 'id': 'rId%d' % i}
        if state:
            attrs['state'] = state
        etree.SubElement(sheets_el, M + 'sheet', attrs)
        etree.SubElement(rels, PR + 'Relationship', Id='rId%d' % i, Type=T_SHEET, Target=relative('xl/workbook.xml', part))
    n = len(final)
    extra = [('theme', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml'),
             ('styles', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml'),
             ('sst', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', 'sharedStrings.xml'),
             ('vba', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin')]
    for k, (_, ty, target) in enumerate(extra, n + 1):
        etree.SubElement(rels, PR + 'Relationship', Id='rId%d' % k, Type=ty, Target=target)
    dst.put('xl/_rels/workbook.xml.rels', rels)

    names = [s[0] for s in final]
    dn = wb.find(M + 'definedNames')
    new_dn = etree.Element(M + 'definedNames')

    def add_names(src_wb, src_order, rename):
        el = src_wb.find(M + 'definedNames')
        if el is None:
            return
        for d in el:
            d = copy.deepcopy(d)
            if d.get('localSheetId') is not None:
                sname = src_order[int(d.get('localSheetId'))]
                sname = rename.get(sname, sname)
                d.set('localSheetId', str(names.index(sname)))
            for old, nw in rename.items():
                d.text = d.text.replace("'%s'!" % old, "'%s'!" % nw)
            new_dn.append(d)
    add_names(wb, sc_order, SC_RENAME)
    add_names(fo.x('xl/workbook.xml'), [n for n, _, _ in fo.sheets()], {})
    for nm, ref in (('p_huidig', 'Start!$F$8'), ('p_oplevering', 'Start!$F$9')):
        d = etree.SubElement(new_dn, M + 'definedName', name=nm)
        d.text = ref
    # Excel wil de namen gesorteerd (hoofdletterongevoelig), _xlnm-namen per blad
    items = sorted(new_dn, key=lambda d: (d.get('name').lower(), int(d.get('localSheetId') or -1)))
    for d in list(new_dn):
        new_dn.remove(d)
    for d in items:
        new_dn.append(d)
    if dn is not None:
        dn.addprevious(new_dn)
        wb.remove(dn)
    else:
        wb.find(M + 'sheets').addnext(new_dn)
    calc = wb.find(M + 'calcPr')
    calc.set('fullCalcOnLoad', '1')
    bv = wb.find(M + 'bookViews')
    for v in bv:
        v.set('activeTab', '0')
        v.attrib.pop('firstSheet', None)
    dst.put('xl/workbook.xml', wb)

    # Start als actief blad
    t = dst.x(part_of['Start'])
    t.find(M + 'sheetViews')[0].set('tabSelected', '1')
    dst.put(part_of['Start'], t)

    # rekenketen weg (Excel bouwt hem opnieuw op), inhoudstypen bijwerken
    dst.parts.pop('xl/calcChain.xml', None)
    copier.remove_override('xl/calcChain.xml')
    dst.put('[Content_Types].xml', copier.ct)
    sst.finish()
    dst.put('xl/sharedStrings.xml', sst.t)
    dst.put('xl/styles.xml', styles.t)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application>'
           '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><AppVersion>16.0300</AppVersion></Properties>')
    dst.parts['docProps/app.xml'] = app.encode('utf-8')

    # ---- VBA ------------------------------------------------------------------------------------------------
    print('VBA-project schrijven')
    vdir = os.path.join(ROOT, 'vba')
    modules = [{'name': 'ThisWorkbook', 'document': True, 'code': open(os.path.join(vdir, 'ThisWorkbook.cls'), 'rb').read()}]
    for name, part, state, code in final:
        f = os.path.join(vdir, code + '.cls')
        src = open(f, 'rb').read() if os.path.exists(f) else vba_project.document_module_source(code)
        modules.append({'name': code, 'document': True, 'code': src})
    for m in VBA_STD:
        modules.append({'name': m, 'document': False, 'code': open(os.path.join(vdir, m + '.bas'), 'rb').read()})
    for m in modules:
        assert b'\r\n' in m['code'] and b'\n' not in m['code'].replace(b'\r\n', b''), 'regeleinden ' + m['name']
        m['code'].decode('cp1252')
    tmpl = os.path.join(os.path.dirname(a.out), 'vbaProject.orig.tmp')
    open(tmpl, 'wb').write(Package(a.scenario).parts['xl/vbaProject.bin'])
    dst.parts['xl/vbaProject.bin'] = vba_project.build_vba_project(tmpl, modules)
    os.remove(tmpl)
    note('%d modules' % len(modules))

    dst.save(a.out)
    print('Klaar: ' + a.out)

    # ---- PowerPoint-sjabloon v21: alleen vormnamen erbij -------------------------------------------------------
    maak_sjabloon(a.pptx, a.pptx_out)
    print('Klaar: ' + a.pptx_out)


def maak_sjabloon(src, out):
    """v21 = v20 met namen voor de vormen die de macro nu ook vult (lay-out ongewijzigd)"""
    from pptx import Presentation
    p = Presentation(src)
    namen = []
    for shp in p.slides[0].shapes:
        if shp.has_text_frame and shp.text_frame.text.strip().startswith('Cijfers per'):
            namen.append((shp, 'COVER_DATUM'))
    tekst = {shp.name: shp for shp in p.slides[9].shapes}
    plan = {'Text 8': 'PEW_FD_WAARDE', 'Text 9': 'PEW_FD_SUB', 'Text 13': 'PEW_BASIS_WAARDE', 'Text 14': 'PEW_BASIS_SUB',
            'Text 18': 'PEW_RES_WAARDE', 'Text 19': 'PEW_RES_SUB', 'Text 20': 'PEW_DATUM'}
    for old, new in plan.items():
        namen.append((tekst[old], new))
    assert len(namen) == 8, [n for _, n in namen]
    for shp, nm in namen:
        shp.name = nm
    p.save(out)

if __name__ == '__main__':
    main()
