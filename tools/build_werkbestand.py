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

import fo_aanpassingen  # noqa: E402
import opzet as opzet_fo  # noqa: E402
import scenario_koppeling  # noqa: E402
import nieuwe_bladen  # noqa: E402
import vba_project  # noqa: E402
from xlsxpkg import (CT, CT_CHART, CT_CHARTCOLORS, CT_CHARTSTYLE, CT_DRAWING, CT_SHEET, M, PR, R, SST, T_DRAWING,  # noqa: E402
                     T_SHEET, Package, Sheet, Styles, col2num, num2col, parse, rels_name, relative, replace_in_formulas,
                     resolve, rijen_invoegen, split_ref, tobytes, unshare)

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
    ('CF - opbrengsten', 'fo', 'CF - opbrengsten', 'shCFOpbrengsten'),
    ('CF - kosten', 'fo', 'CF - kosten', 'shCFKostenFO'),
    ('Dashboard', 'sc', 'Dashboard', 'shDashboard'),
    ("Scenario's", 'sc', "Scenario's", 'shScenarios'),
    ('2. Ontw PEW', 'new', '2. Ontw PEW', 'shOntwPEW'),
    ('CF - hulp VENA', 'fo', 'CF - hulp VENA', 'shVENA'),
    ('PowerPoint', 'sc', 'PowerPoint', 'shPowerPoint'),
    ('Invoer', 'sc', 'Invoer', 'shInvoer'),
    ('Woningtypes', 'sc', 'Woningtypes', 'shWoningtypes'),
    ('Scenario kosten', 'sc', 'CF kosten', 'shCFKosten'),
    ('Model', 'sc', 'Model', 'shModel'),
    ('Vorige prognose', 'new', 'Vorige prognose', 'shVorigePrognose'),
    ('Controle oud FO', 'new', 'Controle oud FO', 'shControleOudFO'),
    ('Behind the scens', 'fo', 'Behind the scens', 'shBehindScens'),
    ('_sjabloon', 'new', '_sjabloon', 'shSjabloon'),
    ('_koppeling', 'new', '_koppeling', 'shKoppeling'),
]
VERVALLEN = {'1. Cashflow', '3. Restant bdgt', '2. Ontw PEW', 'Notities FO gesprek', 'Verkooptempo',
             'Toelichting onderlegger', '1. Ontw businessc FD', '5. K&R businessc'}    # 2. Ontw PEW komt in een nieuw jasje
VERBORGEN = {'Invoer', 'Woningtypes', 'Scenario kosten', 'Model',   # rekenbladen van het scenariomodel: alles is formule
             'PowerPoint', 'Vorige prognose', 'Controle oud FO', 'Behind the scens'}
VBA_STD = ['CashflowNaarPowerPoint', 'FOWerkbestand']
DEMO = {'project': 'Parkzoom', 'oplevering': (2028, 3), 'startbouw': (2027, 2),
        'dia10': [('Meest recente FD', 'SKO', 'Inclusief uitruilregeling'),
                  ('Huidige PEW-basis', 'FD DO (Aanbieding)', '10 april 2026'),
                  ('Mogelijk resultaat', '4,8%', 'Bij rentelasten richting nul')]}


def indeling(oz):
    """LAYOUT voor dit FO: de vaste tabbladen die het FO heeft, plus de extra FO-tabbladen (zoals '2B. Ontw PEW') na
    '2. Ontw PEW', met een gemaakte codenaam"""
    fo = [n for n, _ in oz['bladen']]
    bekend = {src for _, o, src, _ in LAYOUT if o == 'fo'}
    out = []
    for item in LAYOUT:
        if item[1] == 'fo' and item[2] not in fo:
            note('tabblad %s staat niet in het FO: overgeslagen' % item[2])
            continue
        out.append(item)
        if item[0] == '2. Ontw PEW':
            for n in fo:
                if n not in bekend and n not in VERVALLEN:
                    code = 'shFO' + re.sub(r'[^A-Za-z0-9]', '', n)[:24]
                    out.append((n, 'fo', n, code))
                    note('extra FO-tabblad overgenomen: %s' % n)
    return out


def config(oz, layout):
    """de opzet voor de macro's (FOWerkbestand: Cfg), op _sjabloon kolom E:F"""
    from xlsxpkg import col2num as cn
    ob, k, cf = oz['opbrengsten'], oz['kosten'], oz['cashflow']
    hulp = {b: '|'.join('%s:%d' % (c, r[0]) for c, r in kol.items()) for b, kol in oz['hulpkolommen'].items()}
    controles = ['CF - kosten|B%d|onderdelen' % k['onderdelen'], 'CF - kosten|B%d|totaal kosten' % k['totaal'],
                 'CF - kosten|B%d|periode in cashflow' % k['periode'],
                 'CF - opbrengsten|B%d|termijnen' % ob['euro_kop'], 'CF - opbrengsten|B%d|totaal opbrengsten' % ob['totaal'],
                 'CF - opbrengsten|B%d|periode in cashflow' % ob['periode']]
    return [
        ('fo_bladen', '|'.join(src for _, o, src, _ in layout if o == 'fo') + '|2. Ontw PEW'),
        ('transportrijen', '|'.join(str(t['tr']) for t in ob['types'])),
        ('opbrengsten_kol1', cn(ob['eerste'])), ('opbrengsten_kol2', cn(ob['laatste'])),
        ('opbrengsten_totaal', ob['totaal']), ('opbrengsten_correctie', ob['correctie']),
        ('correctie_bereik', 'B%d:%s%d' % (ob['correctie'], ob['laatste'], ob['correctie'])),
        ('daeb_totaal', ob.get('daeb_totaal') or 0),
        ('daeb_bereik', 'B%d:%s%d' % (ob['daeb_sjabloon'], ob['laatste'], ob['daeb_eind'] + 1) if ob['daeb_sjabloon'] else ''),
        ('voorstel_bereik', 'B%d:%s%d' % (k['voorstel_start'], k['status'], k['voorstel_eind'])),
        ('controles', ';'.join(controles)),
        ('cashflow_rij1', cf['periodes'][0][2]),

        ('voorstel_offset', k['voorstel_offset']),
        ('vt_types', '|'.join('%d:%d:%d:%d' % (t['verkoop'], t['koop'], t['tr'] + 1, t['hb'] - 1) for t in ob['types'])),
        ('vt_invoeg', '|'.join('%d:%d' % (r, n) for r, n in oz['ingevoegd'])),
    ] + [('hulp|' + b, v) for b, v in hulp.items()]


def vervallen_als_waarde(pkg, part, blad, wb):
    """formules die naar een vervallen tabblad verwijzen (Verkooptempo, Notities, 1. Cashflow, 3. Restant bdgt): de waarde"""
    weg = VERVALLEN - {'2. Ontw PEW'}
    t = pkg.x(part)
    n = 0
    for c in t.iter(M + 'c'):
        f = c.find(M + 'f')
        if f is None or not f.text or not any(("'%s'!" % w) in f.text or (w + '!') in f.text for w in weg):
            continue
        v = wb[blad][c.get('r')].value
        c.remove(f)
        for x in c.findall(M + 'v') + c.findall(M + 'is'):
            c.remove(x)
        c.attrib.pop('t', None)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            etree.SubElement(c, M + 'v').text = repr(v) if isinstance(v, float) else str(v)
        elif v not in (None, ''):
            c.set('t', 'inlineStr')
            etree.SubElement(etree.SubElement(c, M + 'is'), M + 't').text = str(v)
        n += 1
    if n:
        pkg.put(part, t)
    return n


def daeb_uit_verkooptempo(sh, oz, wb, st, sjabloon):
    """FO met een eigen DAEB-blok: aantal en verkoop van de corporatie stonden op Verkooptempo (de rij met 'DAEB' in de
    naam, anders de eerste rij na de koopwoningtypes). Verkooptempo vervalt: als regels onder de periode-rij zetten"""
    ob = oz['opbrengsten']
    if not ob['daeb'] or 'Verkooptempo' not in wb.sheetnames:
        return
    o = wb['CF - opbrengsten']
    if any(str(o.cell(r, 2).value or '').lower().startswith('aantal daeb') for r in range(1, o.max_row + 1)):
        return
    v = wb['Verkooptempo']
    kop = next((r for r in range(1, 30) if str(v.cell(r, 1).value or '').strip().lower() == 'verkooptempo'), None)
    if not kop:
        return
    namen = [(r, str(v.cell(r, 1).value or '').strip()) for r in range(kop + 1, kop + 20)]
    rij = next((r for r, nm in namen if 'daeb' in nm.lower()), None)
    if rij is None:
        r = kop + 1 + len(ob['types'])
        nm = str(v.cell(r, 1).value or '').strip()
        rij = r if nm and not nm.lower().startswith('totaal') else None
    if rij is None:
        return
    acol = next((c for c in range(1, 40) if str(v.cell(kop, c).value or '').strip().lower() == 'aantal'), 3)
    aantal = v.cell(rij, acol).value
    # kwartaalsleutels op Verkooptempo en op CF - opbrengsten
    def sleutels(ws, jr, qr, c1, c2):
        out, jaar = {}, None
        for c in range(c1, c2 + 1):
            if isinstance(ws.cell(jr, c).value, (int, float)):
                jaar = int(ws.cell(jr, c).value)
            q = str(ws.cell(qr, c).value or '').strip()
            if jaar and re.match(r'Q[1-4]$', q):
                out[jaar * 4 + int(q[1])] = c
        return out
    kv = sleutels(v, kop - 1, kop, 2, 40)
    ko = sleutels(o, ob['jaarrij'], ob['euro_kop'], col2num(ob['eerste']), col2num(ob['laatste']))
    r1 = ob['periode'] + 2
    for c in ['B', 'D'] + [num2col(x) for x in ko.values()]:
        assert not sh.get('%s%d' % (c, r1)) and not sh.get('%s%d' % (c, r1 + 1)), 'rij %d/%d niet leeg' % (r1, r1 + 1)
    z = fo_aanpassingen.Zetter(sh, 'CF - opbrengsten', sjabloon)
    z('B%d' % r1, 'Aantal DAEB-woningen', st['tekst'])
    z('D%d' % r1, aantal, st['invoer_getal'])
    z('B%d' % (r1 + 1), 'Verkoop en levering DAEB (aantal woningen per kwartaal; transport = verkoop)', st['tekst'])
    for key, c in ko.items():
        x = v.cell(rij, kv[key]).value if key in kv else None
        z('%s%d' % (num2col(c), r1 + 1), x if isinstance(x, (int, float)) and x else None, st['invoer_getal'])
    note('CF - opbrengsten: DAEB-aantal (%s) en -verkoop van Verkooptempo (%s) naar rij %d-%d'
         % (aantal, v.cell(rij, 1).value, r1, r1 + 1))


def vba_constanten():
    """vaste plekken van het standaard-FO voor de macro Oud FO inlezen (OmzettenOudFO)"""
    import standaard_fo as S
    from xlsxpkg import col2num as cn
    w = {'NT': S.NT, 'NTERM': S.NTERM, 'NQ': S.NQ, 'C_ACT': S.C_ACT, 'C_Q1': S.C_Q1, 'BLOK': S.BLOK, 'VK_RIJ': S.VK,
         'OZ1': S.OZ1, 'OBLOK': S.OBLOK, 'OV1_RIJ': S.OV1, 'OV2_RIJ': S.OV2, 'PER_RIJ': S.PER, 'DAEB_KOP_RIJ': S.DAEB_KOP,
         'DAEB_AANTAL_RIJ': S.DAEB_AANTAL, 'DAEB_VERKOOP_RIJ': S.DAEB_VERKOOP, 'KBLOK': S.KBLOK, 'K_POST1': S.K_POST1,
         'NPOST': S.NPOST, 'C_TOEL': cn(S.TOEL), 'P_TYPE1': S.P_TYPE1, 'P_AANS1': S.P_AANS1, 'P_BLOK': S.P_BLOK,
         'P_BN1': S.P_BN1, 'P_BNN': S.P_BNN}
    regels = ["' vaste plekken (gegenereerd uit tools/standaard_fo.py)"]
    regels += ['Private Const %s As Long = %d' % (k, v) for k, v in w.items()]
    regels.append('Private Const P_REGELS As String = "%s"' % '|'.join(S.P_REGELS))
    return '\r\n'.join(regels)


def project_uit_bestandsnaam(pad):
    """'20261001_FO_Parkzoom_2_Q4.xlsx' -> 'Parkzoom 2'"""
    n = os.path.splitext(os.path.basename(pad))[0]
    n = re.sub(r'^\d{6,8}[_ ]*', '', n)
    n = re.sub(r'^FO[_ ]+', '', n, flags=re.I)
    n = re.sub(r'[_ ]+Q[1-4]\b.*$', '', n)
    n = re.sub(r'[_ ]+(fictieve|demo).*$', '', n, flags=re.I)
    return n.replace('_', ' ').strip() or 'Project'


EXTERN = re.compile(r"'?\[\d+\]([^'!\[\]]+)'?!")


def lokale_verwijzingen(t):
    """verwijzingen naar een extern bestand ('[2]FO - huidig'!) naar het tabblad in dit bestand"""
    n = 0
    for f in t.iter(M + 'f'):
        if f.text and EXTERN.search(f.text):
            f.text = EXTERN.sub(lambda m: "'%s'!" % m.group(1), f.text)
            n += 1
    return n


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

def verwijder_vormen(pkg: Package, sheet_part: str, namen) -> int:
    """vormen (knoppen) met deze namen uit de tekening van een werkblad halen"""
    n = 0
    for rid, ty, target, _ in pkg.rels(sheet_part):
        if ty != T_DRAWING:
            continue
        d = resolve(sheet_part, target)
        t = pkg.x(d)
        for anchor in list(t):
            cnv = [e for e in anchor.iter() if etree.QName(e).localname == 'cNvPr']
            if cnv and cnv[0].get('name') in namen:
                t.remove(anchor)
                n += 1
        pkg.put(d, t)
    return n


def oud_cashflow(fo_path, oz):
    """periodes van '1. Cashflow' uit het FO (waarden van Excel): jaar, Q, kosten, omzet, CF, CF vorig kwartaal"""
    import openpyxl
    wb = openpyxl.load_workbook(fo_path, data_only=True)
    ws = wb['1. Cashflow']
    return [{'B': jaar, 'C': q, 'D': ws.cell(r, 4).value, 'F': ws.cell(r, 6).value, 'H': ws.cell(r, 8).value,
             'J': ws.cell(r, 10).value} for jaar, q, r in oz['cashflow']['periodes']]


def pew_inhoud(fo_path):
    """alle gevulde cellen van '2. Ontw PEW' (formules en waarden), voor het nieuwe jasje"""
    import openpyxl
    ws = openpyxl.load_workbook(fo_path)['2. Ontw PEW']
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}


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
    ap.add_argument('--project', help='projectnaam (standaard uit de bestandsnaam van het FO)')
    ap.add_argument('--oud', help='het oude FO waaruit het standaard-FO is omgezet (voor de controle)')
    ap.add_argument('--demo', action='store_true', help='teksten van het demo-FO (dia 10)')
    ap.add_argument('--startbouw', help='jaar-kwartaal, bv. 2027-2 (standaard de markering Start bouw op CF - opbrengsten)')
    ap.add_argument('--oplevering', help='jaar-kwartaal, bv. 2028-3 (standaard de markering Oplevering op CF - opbrengsten)')
    a = ap.parse_args()
    os.makedirs(os.path.dirname(a.out), exist_ok=True)

    print('Opzet van het FO lezen')
    fo_bron = a.fo
    oz0 = opzet_fo.lees(fo_bron)
    # per woningtype 3 rijen onder de transportregel (verkocht %, getransporteerd %, grondprijs); van onder naar boven
    ingevoegd = [] if oz0['standaard'] else sorted(((t['tr'] + 1, 3) for t in oz0['opbrengsten']['types']), reverse=True)
    fo_pkg = Package(fo_bron)
    for at, n in ingevoegd:
        rijen_invoegen(fo_pkg, 'CF - opbrengsten', at, n)
    fo_tmp = a.out + '.fo.tmp.xlsx'
    fo_pkg.save(fo_tmp)
    a.fo = fo_tmp
    oz = opzet_fo.lees(fo_tmp)
    oz['ingevoegd'] = ingevoegd
    if ingevoegd:
        note('CF - opbrengsten: per woningtype 3 rijen ingevoegd (rij %s van het FO)' % ', '.join(str(r) for r, _ in ingevoegd[::-1]))
    ob, ko = oz['opbrengsten'], oz['kosten']
    note('CF - opbrengsten: %d woningtypes, kolommen %s..%s, totaal rij %d; CF - kosten: posten rij %d-%d, totaal rij %d'
         % (len(ob['types']), ob['eerste'], ob['laatste'], ob['totaal'], ko['posten'][0], ko['posten'][-1], ko['totaal']))
    if a.demo or os.path.abspath(fo_bron) == os.path.abspath(ap.get_default('fo')):
        project = dict(DEMO)
    else:
        opl = ob['mijlpalen'].get('oplevering')
        project = {'project': project_uit_bestandsnaam(fo_bron), 'oplevering': opl or (None, None)}
    if a.project:
        project['project'] = a.project
    if a.oplevering:
        j, q = a.oplevering.split('-')
        project['oplevering'] = (int(j), int(q))
    if a.startbouw:
        j, q = a.startbouw.split('-')
        project['startbouw'] = (int(j), int(q))
    elif 'startbouw' not in project:
        project['startbouw'] = ob['mijlpalen'].get('startbouw') or (None, None)
    note('project %s, oplevering %s' % (project['project'], project['oplevering']))
    layout = indeling(oz)
    scenario_koppeling.configureer(oz)

    print('Bronnen lezen')
    sc = Package(a.scenario)
    fo = Package(a.fo)
    tmp_new = a.out + '.nieuw.tmp'
    inv_tot = scenario_koppeling.INV_RN + 2
    info = nieuwe_bladen.bouw(tmp_new, project, oud_cashflow(a.fo, oz), os.path.basename(fo_bron), pew_inhoud(a.fo), oz,
                              inv_tot, scenario_koppeling.DAEB_TYPE)
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

    import openpyxl
    fo_waarden = openpyxl.load_workbook(a.fo, data_only=True)
    print('Werkbladen overnemen')
    sc_sheets = {n: (p, st) for n, p, st in sc.sheets()}
    sc_order = [n for n, _, _ in sc.sheets()]
    srcs = {'fo': fo, 'new': new}
    src_states = {k: {n: st for n, _, st in p.sheets()} for k, p in srcs.items()}
    final = []        # (naam, part, state)
    for name, origin, src_name, code in layout:
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
            if origin == 'fo':
                t = dst.x(part)
                g = unshare(t)
                n = lokale_verwijzingen(t)
                if g or n:
                    dst.put(part, t)
                if g:
                    note('%s: %d gedeelde formules uitgeschreven' % (name, g))
                if n:
                    note('%s: %d formules verwezen naar een ander bestand; nu naar het tabblad in dit bestand' % (name, n))
                v = vervallen_als_waarde(dst, part, src_name, fo_waarden)
                if v:
                    note('%s: %d formules met een verwijzing naar een vervallen tabblad (%s) zijn waarde geworden'
                         % (name, v, ', '.join(sorted(VERVALLEN - {'2. Ontw PEW'}))))
        if name in VERBORGEN:
            state = 'hidden'
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

    # ---- CF - opbrengsten en CF - kosten: transport, correctie, methode, voorstel --------------------------
    print('CF - opbrengsten en CF - kosten aanpassen')
    stijl_t = new.x(new.sheet_part('_stijlen'))
    st = {}
    for naam, ref in info['stijlen'].items():
        c = Sheet(stijl_t, sst).cell(ref, create=False)
        st[naam] = maps['new'][0][int(c.get('s', '0'))]
    sjabloon = []
    for blad, functie in (('CF - opbrengsten', lambda sh: fo_aanpassingen.opbrengsten(sh, oz, st, sjabloon, note)),
                          ('CF - kosten', lambda sh: fo_aanpassingen.kosten(sh, oz, st, sjabloon, styles, note))):
        sh = Sheet(dst.x(part_of[blad]), sst)
        functie(sh)
        if blad == 'CF - opbrengsten':
            daeb_uit_verkooptempo(sh, oz, fo_waarden, st, sjabloon)
        dst.put(part_of[blad], sh.t)
    ab = oz['aansluiting_blad']
    sh = Sheet(dst.x(part_of[ab]), sst)
    fo_aanpassingen.aansluiting(sh, oz, st, sjabloon, note)
    dst.put(part_of[ab], sh.t)
    for ref, tekst in info['pew_extra'].items():
        sjabloon.append(('2. Ontw PEW', ref, tekst))
    # sjabloonlijst op het verborgen blad _sjabloon
    sh = Sheet(dst.x(part_of['_sjabloon']), sst)
    for i, (blad, ref, inhoud) in enumerate(sjabloon, 2):
        sh.set_value('A%d' % i, blad)
        sh.set_value('B%d' % i, ref)
        sh.set_value('C%d' % i, inhoud if inhoud else None)
    for i, (sleutel, waarde) in enumerate(config(oz, layout) + [('top3_rij', info.get('top3_row', 0))], 1):
        sh.set_value('E%d' % i, sleutel)
        sh.set_value('F%d' % i, waarde)
    dst.put(part_of['_sjabloon'], sh.t)
    note('_sjabloon: %d cellen' % len(sjabloon))

    # ---- scenariotabbladen met formules aan het FO koppelen (geen 'Ophalen uit FO' meer nodig) ------------
    print('Scenario koppelen: Invoer, Woningtypes, Scenario kosten, Dashboard')
    sh_o = Sheet(dst.x(part_of['CF - opbrengsten']), sst)
    sh_k = Sheet(dst.x(part_of['CF - kosten']), sst)
    types = scenario_koppeling.structuur(sh_o, sh_k)
    sh = Sheet(dst.x(part_of['_koppeling']), sst)
    scenario_koppeling.sleutelrijen(sh)
    slots = scenario_koppeling.vul_hulp(sh, types, note)
    scenario_koppeling.daeb_hulp(sh, types, note)
    dst.put(part_of['_koppeling'], sh.t)
    sh = Sheet(dst.x(part_of['Woningtypes']), sst)
    scenario_koppeling.woningtypes(sh, types, slots, st)
    sh.set_value('B1', 'Woningtypes — rekent automatisch mee met CF - opbrengsten, Aansluiting opbrengsten en CF - kosten')
    sh.set_value('B2', 'Blok 1, 2, 3 en 7 zijn formules (verborgen hulpblad _koppeling): niet overschrijven. Blok 4 en 6 '
                       '(vorige prognose) vul je met de knoppen hieronder; blok 5 alleen bij DAEB.')
    sh.set_value('B3', 'Een ander woningtype of een andere termijnverdeling pas je aan op CF - opbrengsten (bouwplanning, '
                       'transport) en Aansluiting opbrengsten (koopsommen, % per termijn).')
    dst.put(part_of['Woningtypes'], sh.t)
    sh_w = Sheet(dst.x(part_of['Woningtypes']), sst)
    sh = Sheet(dst.x(part_of['Invoer']), sst)
    scenario_koppeling.invoer(sh, types, st)
    scenario_koppeling.daeb_woningtypes(sh_w, sh, st)
    dst.put(part_of['Woningtypes'], sh_w.t)
    sh.set_value('B1', 'Invoer — cashflow per periode en verkocht/transport per woningtype, automatisch uit het FO')
    sh.set_value('B2', 'Formules: verwerkte kwartalen uit FO - actuals, prognose uit CF - kosten (rij %d) en CF - opbrengsten '
                       '(rij %d); verkocht = verkooptempo, transport = transportregel op CF - opbrengsten. Niet overschrijven.'
                 % (ko['totaal'], ob['totaal']))
    sh.set_value('B3', 'Kolom J = vorige prognose (tab Vorige prognose). Rij %d-%d: totalen en aansluiting op het PEW.'
                 % (inv_tot, inv_tot + 3))
    sh.set_value('B4', None)
    sh.set_value('B5', '1 · CASHFLOW · automatisch uit FO - actuals en de CF-tabbladen')
    dst.put(part_of['Invoer'], sh.t)
    # knoppen die de formules zouden overschrijven of verschuiven: uit de tekeningen
    for blad, namen in (('Invoer', {'btnUitFO'}), ('Woningtypes', {'btnTypeInvoegen', 'btnTypeVerwijderen'})):
        n = verwijder_vormen(dst, part_of[blad], namen)
        note('%s: %d knop(pen) weggehaald' % (blad, n))
    sh = Sheet(dst.x(part_of['Scenario kosten']), sst)
    scenario_koppeling.scenario_kosten(sh, st)
    dst.put(part_of['Scenario kosten'], sh.t)
    sh = Sheet(dst.x(part_of['Dashboard']), sst)
    sh.set_formula('W5', 'Start!$C$12')
    sh.set_formula('X5', 'Start!$D$12')
    sh.set_value('V68', 'FO-KOPPELING (formules, rekent automatisch mee)')
    sh.set_value('W69', 'dit werkbestand')
    sh.set_value('Z69', 'tab Invoer, Woningtypes en Scenario kosten zijn formules op de FO-tabbladen')
    sh.set_formula('W70', 'Start!$E$12')
    sh.set_formula('W71', 'SUM(Invoer!$F$8:$F$%d)' % scenario_koppeling.INV_RN)
    sh.set_formula('W72', 'SUM(Invoer!$D$8:$D$%d)' % scenario_koppeling.INV_RN)
    sh.set_formula('W73', 'Start!$C$10')
    dst.put(part_of['Dashboard'], sh.t)

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

    for name, tot in (('FO - vorig', 1500), ('FO - huidig', 1500), ('FO - actuals', 6000)):
        sh = sheet(name)
        for col, (first, last) in oz['hulpkolommen'][name].items():
            extend_helpers(sh, col, first, last, tot)
            note('%s: hulpkolom %s rij %d t/m %d (was t/m %d)' % (name, col, first, tot, last))
        save(name, sh)

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
                if sname not in names:
                    continue
                d.set('localSheetId', str(names.index(sname)))
            for old, nw in rename.items():
                d.text = d.text.replace("'%s'!" % old, "'%s'!" % nw)
            if re.search(r'\[\d+\]|#REF!', d.text or '') or any("'%s'!" % v in (d.text or '') or '%s!' % v in (d.text or '')
                                                               for v in VERVALLEN - {'2. Ontw PEW'}):
                note('naam %s overgeslagen (%s)' % (d.get('name'), d.text))
                continue
            new_dn.append(d)
    add_names(wb, sc_order, SC_RENAME)
    add_names(fo.x('xl/workbook.xml'), [n for n, _, _ in fo.sheets()], {})
    for nm, ref in (('p_huidig', 'Start!$F$8'), ('p_oplevering', 'Start!$F$9'), ('p_actuals', 'Start!$F$12'),
                    ('p_startbouw', 'Start!$F$7')):
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
        code = open(os.path.join(vdir, m + '.bas'), 'rb').read()
        if b"'@@CONSTANTEN@@" in code:
            code = code.replace(b"'@@CONSTANTEN@@", vba_constanten().encode('cp1252'))
        modules.append({'name': m, 'document': False, 'code': code})
    for m in modules:
        assert b'\r\n' in m['code'] and b'\n' not in m['code'].replace(b'\r\n', b''), 'regeleinden ' + m['name']
        m['code'].decode('cp1252')
    tmpl = os.path.join(os.path.dirname(a.out), 'vbaProject.orig.tmp')
    open(tmpl, 'wb').write(Package(a.scenario).parts['xl/vbaProject.bin'])
    dst.parts['xl/vbaProject.bin'] = vba_project.build_vba_project(tmpl, modules)
    os.remove(tmpl)
    note('%d modules' % len(modules))

    dst.save(a.out)
    os.remove(fo_tmp)
    import json
    with open(os.path.splitext(a.out)[0] + '.opzet.json', 'w') as fh:
        json.dump({'fo': os.path.basename(fo_bron), 'oud_fo': os.path.abspath(a.oud) if a.oud else None, 'opzet': oz, 'invoer_laatste': scenario_koppeling.INV_RN,
                   'invoer_totaal': inv_tot, 'daeb_type': scenario_koppeling.DAEB_TYPE}, fh, indent=1, default=str)
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
            'Text 18': 'PEW_RES_WAARDE', 'Text 19': 'PEW_RES_SUB', 'Text 20': 'PEW_DATUM', 'Shape 3': 'PEW_TABEL'}
    for old, new in plan.items():
        namen.append((tekst[old], new))
    assert len(namen) == 9, [n for _, n in namen]
    for shp, nm in namen:
        shp.name = nm
    p.save(out)

if __name__ == '__main__':
    main()
