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
                     T_SHEET, Package, Sheet, Styles, rels_name, relative, replace_in_formulas,
                     resolve, split_ref, unshare)

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
    ('CF vorig en huidig', 'new', 'CF vorig en huidig', 'shVorigePrognose'),
    ('2. Ontw PEW', 'new', '2. Ontw PEW', 'shOntwPEW'),
    ('2B. Ontw PEW', 'new', '2B. Ontw PEW', 'shOntwPEWB'),
    ('CF - hulp VENA', 'fo', 'CF - hulp VENA', 'shVENA'),
    ('PowerPoint', 'sc', 'PowerPoint', 'shPowerPoint'),
    ('Invoer', 'sc', 'Invoer', 'shInvoer'),
    ('Woningtypes', 'sc', 'Woningtypes', 'shWoningtypes'),
    ('Scenario kosten', 'sc', 'CF kosten', 'shCFKosten'),
    ('Model', 'sc', 'Model', 'shModel'),
    ('Controle oud FO', 'new', 'Controle oud FO', 'shControleOudFO'),
    ('Behind the scens', 'fo', 'Behind the scens', 'shBehindScens'),
    ('_sjabloon', 'new', '_sjabloon', 'shSjabloon'),
    ('_koppeling', 'new', '_koppeling', 'shKoppeling'),
]
VERVALLEN = {'1. Cashflow', '3. Restant bdgt', '2. Ontw PEW', '2B. Ontw PEW', 'Notities FO gesprek', 'Verkooptempo',
             'Toelichting onderlegger', '1. Ontw businessc FD', '5. K&R businessc'}    # 2. Ontw PEW komt in een nieuw jasje
VERBORGEN = {'Invoer', 'Woningtypes', 'Scenario kosten', 'Model',   # rekenbladen van het scenariomodel: alles is formule
             'PowerPoint', 'Controle oud FO', 'Behind the scens',
             '2B. Ontw PEW'}                                         # zichtbaar zodra het project DAEB-woningen heeft (macro)
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
    cf = oz['cashflow']
    hulp = {b: '|'.join('%s:%d' % (c, r[0]) for c, r in kol.items()) for b, kol in oz['hulpkolommen'].items()}
    # alleen wat de macro's nog lezen (Cfg in vba/FOWerkbestand.bas): de eerste cashflowrij van het oude FO en de
    # hulpkolommen van de exports; top3_rij komt erbij in main()
    return [
        ('cashflow_rij1', cf['periodes'][0][2]),
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


def vba_constanten():
    """vaste plekken van het standaard-FO voor de macro Oud FO inlezen (OmzettenOudFO)"""
    import standaard_fo as S
    from xlsxpkg import col2num as cn
    w = {'NT': S.NT, 'NTERM': S.NTERM, 'NQ': S.NQ, 'C_ACT': S.C_ACT, 'C_Q1': S.C_Q1, 'BLOK': S.BLOK, 'VK_RIJ': S.VK,
         'OZ1': S.OZ1, 'TT_RIJ': S.TT, 'KANS_RIJ': S.KANS, 'OBLOK': S.OBLOK, 'OV1_RIJ': S.OV1, 'OV2_RIJ': S.OV2, 'PER_RIJ': S.PER, 'DAEB_KOP_RIJ': S.DAEB_KOP,
         'DAEB_AANTAL_RIJ': S.DAEB_AANTAL, 'DAEB_VERKOOP_RIJ': S.DAEB_VERKOOP, 'KBLOK': S.KBLOK, 'K_POST1': S.K_POST1,
         'NPOST': S.NPOST, 'C_TOEL': cn(S.TOEL), 'P_TYPE1': S.P_TYPE1, 'P_AANS1': S.P_AANS1, 'P_BLOK': S.P_BLOK,
         'P_BN1': S.P_BN1, 'P_BNN': S.P_BNN}
    regels = ["' vaste plekken (gegenereerd uit tools/standaard_fo.py)"]
    regels += ['Private Const %s As Long = %d' % (k, v) for k, v in w.items()]
    regels.append('Private Const P_REGELS As String = "%s"' % '|'.join(S.P_REGELS))
    return '\r\n'.join(regels)


def palet_werkbestand(dst):
    """het werkbestand in hetzelfde palet als het PowerPoint-sjabloon v22 (pptx_v22.PALET): grafieken, tekeningen,
    stijlen en de voorwaardelijke opmaak in de werkbladen. Gele tekst wordt donker amber (geel op wit is onleesbaar);
    gele vullingen en reeksen worden het Heijmans-geel. Rekent niets anders."""
    import pptx_v22
    pal = pptx_v22.PALET
    geel = {'EDA100', 'F2B800'}
    srgb = re.compile(r'(<a:srgbClr val=")([0-9A-Fa-f]{6})(")')
    rgb = re.compile(r'(rgb=")(FF)?([0-9A-Fa-f]{6})(")')
    font = re.compile(r'<(?:\w+:)?font\b[^>]*?(?<!/)>.*?</(?:\w+:)?font>', re.S)
    n = 0

    def kleur(m, tekst=False):
        nonlocal n
        v = m.group(len(m.groups()) - 1).upper()
        if v not in pal:
            return m.group(0)
        n += 1
        nieuw = pptx_v22.TEKST_GEEL if tekst and v in geel else pal[v]
        if len(m.groups()) == 4:
            return m.group(1) + (m.group(2) or '') + nieuw + m.group(4)
        return m.group(1) + nieuw + m.group(3)
    for naam in list(dst.parts):
        if not naam.endswith('.xml'):
            continue
        if naam.startswith(('xl/charts/', 'xl/drawings/')) and b'srgbClr' in dst.parts[naam]:
            s = dst.parts[naam].decode('utf-8')
            dst.parts[naam] = srgb.sub(kleur, s).encode('utf-8')
        elif naam == 'xl/styles.xml' or naam.startswith('xl/worksheets/sheet'):
            s = dst.parts[naam].decode('utf-8')
            s = font.sub(lambda f: rgb.sub(lambda m: kleur(m, True), f.group(0)), s)
            dst.parts[naam] = rgb.sub(kleur, s).encode('utf-8')
    return n


def harmoniseer_stijlen(t):
    """één vormgeving in het hele bestand: koppen donkerblauw met witte tekst (ook de gele koppen uit het oude FO en het
    scenariomodel), invoer overal hetzelfde geel met een dunne goudrand, totalen lichtblauw"""
    fills, fonts, borders, xfs = (t.find(M + x) for x in ('fills', 'fonts', 'borders', 'cellXfs'))
    kop, invoer = set(), set()
    for i, f in enumerate(fills):
        pf = f.find(M + 'patternFill')
        fg = pf.find(M + 'fgColor') if pf is not None else None
        if fg is None or pf.get('patternType') != 'solid':
            continue
        th, tint, rgb = fg.get('theme'), float(fg.get('tint') or 0), (fg.get('rgb') or '').upper()
        nieuw = None
        if (th == '4' and abs(tint) < 0.01) or rgb in ('FFFEB900', 'FFFFC000', 'FFFFB900'):
            nieuw = 'FF003366'; kop.add(i)
        elif th == '7' and abs(tint - 0.6) < 0.01 or rgb in ('FFFFF2B2', 'FFFFF2CC'):
            nieuw = 'FFFFF2CC'; invoer.add(i)
        elif th == '7' and abs(tint - 0.4) < 0.01:
            nieuw = 'FFDCE8F5'
        if nieuw:
            for tag in ('fgColor', 'bgColor'):
                e = pf.find(M + tag)
                if e is None:
                    e = etree.SubElement(pf, M + tag)
                e.attrib.clear()
                e.set('rgb', nieuw)
    VOLG = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family',
            'charset', 'scheme']
    wit = {}
    goud = {}
    for xf in xfs:
        fid = int(xf.get('fillId', 0))
        if fid in kop:
            fo = int(xf.get('fontId', 0))
            if fo not in wit:
                f = copy.deepcopy(fonts[fo])
                for e in list(f):
                    if e.tag in (M + 'color', M + 'b'):
                        f.remove(e)
                f.append(etree.Element(M + 'b'))
                f.append(etree.Element(M + 'color', rgb='FFFFFFFF'))
                kids = sorted(f, key=lambda e: VOLG.index(etree.QName(e).localname) if etree.QName(e).localname in VOLG else 99)
                for e in list(f):
                    f.remove(e)
                for e in kids:
                    f.append(e)
                fonts.append(f)
                wit[fo] = len(fonts) - 1
            xf.set('fontId', str(wit[fo]))
            xf.set('applyFont', '1')
        elif fid in invoer:
            bo = int(xf.get('borderId', 0))
            if bo not in goud:
                b = etree.Element(M + 'border')
                for zijde in ('left', 'right', 'top', 'bottom'):
                    z = etree.SubElement(b, M + zijde, style='thin')
                    etree.SubElement(z, M + 'color', rgb='FFE9CF7A')
                etree.SubElement(b, M + 'diagonal')
                borders.append(b)
                goud[bo] = len(borders) - 1
            xf.set('borderId', str(goud[bo]))
            xf.set('applyBorder', '1')
    for el in (fonts, borders):
        el.set('count', str(len(el)))
    return len(kop), len(invoer)


def foutkolom(sh, kol, st, styles):
    """kolom A: FOUT of LET OP uit de statuskolom van dezelfde rij (rood/oranje), C1: aantal fouten en let op"""
    rijen = sorted({split_ref(c.get('r'))[1] for c in sh.t.iter(M + 'c')
                    if split_ref(c.get('r'))[0] == kol and c.find(M + 'f') is not None} - {1})     # A1 = '← Start'
    for r in rijen:
        s_ = '%s%d&""' % (kol, r)
        sh.set('A%d' % r, '=IF(LEFT(%s,4)="FOUT","FOUT",IF(LEFT(%s,6)="LET OP","LET OP",""))' % (s_, s_), st['vet'])
    sh.set('B1', '=IF(COUNTIF(A:A,"FOUT")+COUNTIF(A:A,"LET OP")=0,"Geen fouten",COUNTIF(A:A,"FOUT")&" fout · "&'
                 'COUNTIF(A:A,"LET OP")&" let op · zie kolom A")', st['vet'])
    rood = styles.add_dxf('<font><b val="1"/><color rgb="FFB91C1C"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFECACA"/></patternFill></fill>')
    oranje = styles.add_dxf('<font><b val="1"/><color rgb="FF92400E"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFEF3C7"/></patternFill></fill>')
    sh.add_cf('A1:A1000', 'A1="FOUT"', rood)
    sh.add_cf('A1:A1000', 'A1="LET OP"', oranje)
    sh.add_cf('B1', 'LEFT(B1,5)<>"Geen "', rood)
    return len(rijen)


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


def pew_b_inhoud(fo_path):
    """A (DAEB), B (niet-DAEB) en toelichting per post voor '2B. Ontw PEW', door de omzetter op _omzetting (D:G) gezet"""
    import openpyxl
    wb = openpyxl.load_workbook(fo_path, read_only=True)
    if '_omzetting' not in wb.sheetnames:
        return None
    uit = {}
    for rij in wb['_omzetting'].iter_rows(min_col=4, max_col=7, values_only=True):
        if rij[0]:
            uit[str(rij[0]).strip().lower()] = (rij[1], rij[2], rij[3])
    return uit or None


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
    ap.add_argument('--pptx-out', default=os.path.join(ROOT, 'output', 'Kwartaal_Template_cashflow_v22.pptx'))
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
    # het standaard-FO heeft de rijen voor verkocht %, getransporteerd % en grondprijs al (geen rijen invoegen)
    assert oz0['standaard'], 'het werkbestand wordt gebouwd op het standaard-FO (tools/bouw.py)'
    fo_pkg = Package(fo_bron)
    fo_tmp = a.out + '.fo.tmp.xlsx'
    fo_pkg.save(fo_tmp)
    a.fo = fo_tmp
    oz = opzet_fo.lees(fo_tmp)
    oz['ingevoegd'] = []
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
                              inv_tot, scenario_koppeling.DAEB_TYPE, pew_b_inhoud(a.fo))
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
        dst.put(part_of[blad], sh.t)
    # foutkolom vooraan (kolom A) en een telling bovenaan, op CF - opbrengsten en CF - kosten
    for blad, kol in (('CF - opbrengsten', oz['opbrengsten']['status']), ('CF - kosten', oz['kosten']['status'])):
        sh = Sheet(dst.x(part_of[blad]), sst)
        n = foutkolom(sh, kol, st, styles)
        dst.put(part_of[blad], sh.t)
        note('%s: foutkolom A voor %d rijen met een status' % (blad, n))
    ab = oz['aansluiting_blad']
    sh = Sheet(dst.x(part_of[ab]), sst)
    fo_aanpassingen.aansluiting(sh, oz, st, sjabloon, note)
    dst.put(part_of[ab], sh.t)
    for ref, tekst in info['pew_extra'].items():
        sjabloon.append(('2. Ontw PEW', ref, tekst))
    # sjabloonlijst op het verborgen blad _sjabloon
    # (alleen de opzet in E:F: de macro's lezen geen sjabloonlijst meer, het sjabloon is zelf de bron)
    sh = Sheet(dst.x(part_of['_sjabloon']), sst)
    for i, (sleutel, waarde) in enumerate(config(oz, layout) + [('top3_rij', info.get('top3_row', 0))], 1):
        sh.set_value('E%d' % i, sleutel)
        sh.set_value('F%d' % i, waarde)
    dst.put(part_of['_sjabloon'], sh.t)
    note('_sjabloon: opzet voor de macro\'s')

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
    sh.set_value('B3', 'Kolom J = vorige prognose (tab CF vorig en huidig). Rij %d-%d: totalen en aansluiting op het PEW.'
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
    # aansluiting van het scenariomodel op de CF-tabbladen (voor de controle in W54); de overige regels van het oude
    # blok FO-KOPPELING (bestand, opgehaald op) zijn overbodig in het werkbestand
    sh.set_value('V68', 'AANSLUITING MODEL ↔ CF-TABBLADEN')
    for ref in ('V69', 'W69', 'Z69', 'V70', 'W70', 'V73', 'W73', 'Z73'):
        sh.set_value(ref, None)
    sh.set_formula('W71', 'SUM(Invoer!$F$8:$F$%d)' % scenario_koppeling.INV_RN)
    sh.set_formula('W72', 'SUM(Invoer!$D$8:$D$%d)' % scenario_koppeling.INV_RN)
    # teksten die nog naar verborgen tabbladen of de vervallen knop 'Ophalen uit FO' verwezen
    sh.set_value('V2', 'Geel = invoer. De cijfers komen uit CF - opbrengsten en CF - kosten; hier stel je de scenario\'s, '
                       'de rente en de grafiekperiode in.')
    sh.set_value('Z9', 'vroegste start bouw van de woningtypes (CF - opbrengsten); vergelijk met Start rij 7')
    sh.set_value('Z10', 'op de titeldia en in de voettekst van alle dia\'s (Start, cel C6)')
    sh.set_value('V38', 'Zonder macro\'s: de blokken voor de dia\'s staan op het verborgen tabblad PowerPoint.')
    sh.set_formula('W54', 'IF(AND(ABS(Model!$BY$15-N(Dashboard!$W$71))<1,ABS(SUM(Model!$H$8:$H$67)-N(Dashboard!$W$72))<1),"OK",'
                          '"LET OP: het model wijkt af van de CF-tabbladen (opbrengsten "&FIXED(Model!$BY$15/1000000,1)&" mln tegen "'
                          '&FIXED(N(Dashboard!$W$71)/1000000,1)&" mln)")')
    if oz.get('standaard'):
        import standaard_fo as S_
        rk = S_.KT
        for ref, v in (('V82', 'KANS GRONDTERMIJN (Prijslijst kolom Y; CF - opbrengsten rij %d)' % S_.KANS),
                       ('W83', 'Max. financiering · huidig'), ('W84', 'Max. financiering · met kans'),
                       ('W85', 'Extra grondtermijn eerder ontvangen (hoogste stand)'),
                       ('W86', 'Extra opbrengst grondtermijn (totaal prognose)')):
            sh.set_value(ref, v)
        sh.set_formula('Z83', "'CF - opbrengsten'!$D$%d" % (rk + 6), sh.style_of('W71'))
        sh.set_formula('Z84', "'CF - opbrengsten'!$D$%d" % (rk + 7), sh.style_of('W71'))
        sh.set_formula('Z85', "'CF - opbrengsten'!$D$%d" % (rk + 5), sh.style_of('W71'))
        sh.set_formula('Z86', "'CF - opbrengsten'!$D$%d" % (rk + 2), sh.style_of('W71'))
    # periode van de cashflowgrafiek: vanaf / t/m (leeg = automatisch: eerste t/m laatste kwartaal met opbrengsten of
    # kosten, plus één kwartaal); de grafiekreeksen (namen g_...) lopen via Model!BY336:BY338 mee
    st_k, st_l, st_i, st_n, st_w = (sh.style_of(x) for x in ('V82', 'V77', 'W18', 'Z18', 'W9'))
    sh.set_value('V87', 'GRAFIEK · PERIODE (cashflow per kwartaal en de andere grafieken over de hele looptijd)', st_k)
    sh.set_value('W88', 'jaar', st_l)
    sh.set_value('X88', 'kwartaal', st_l)
    sh.set_value('V89', 'Grafiek vanaf', st_l)
    sh.set_value('V90', 'Grafiek t/m', st_l)
    for r in (89, 90):
        sh.set_value('W%d' % r, None, st_i)
        sh.set_value('X%d' % r, None, st_i)
    sh.set_value('Z89', 'leeg = automatisch: het eerste kwartaal met opbrengsten of kosten', st_n)
    sh.set_value('Z90', 'leeg = automatisch: één kwartaal na het laatste kwartaal met opbrengsten of kosten', st_n)
    sh.set_value('V91', 'Getoond', st_l)
    sh.set_formula('W91', 'INDEX(Model!$E$8:$E$67,Model!$BY$336)&" t/m "&INDEX(Model!$E$8:$E$67,Model!$BY$337)', st_w)
    dst.put(part_of['Dashboard'], sh.t)
    sh = Sheet(dst.x(part_of['Model']), sst)
    st_lab = sh.style_of('BX132')
    st_val = sh.style_of('BY132')
    rng = lambda c: '$%s$8:$%s$67' % (c, c)
    # een periode telt mee als er opbrengsten of kosten van minstens €500 zijn (Model rekent in mln), vanaf Q1 van het
    # eerste jaar op CF - opbrengsten (D3): de grafiek loopt mee met de CF-tabbladen
    nz = ('((ISNUMBER({a})*(({a}>=0.0005)+({a}<=-0.0005))+ISNUMBER({u})*(({u}>=0.0005)+({u}<=-0.0005))>0)'
          '*(ISNUMBER({f})*({f}>=\'CF - opbrengsten\'!$D$3*4+1)))').format(a=rng('AT'), u=rng('AU'), f=rng('F'))
    fq = rng('F')
    for r, lab, formule in (
            (336, 'Grafiek: positie van de eerste getoonde periode (Dashboard W89:X89, leeg = eerste met opbrengsten of kosten)',
             'IF(N(Dashboard!$W$89)=0,IFERROR(MATCH(1,INDEX(%s*1,0),0),1),IFERROR(MATCH(1,INDEX(ISNUMBER(%s)*(%s>=Dashboard!$W$89*4+MAX(1,N(Dashboard!$X$89))),0),0),1))'
             % (nz, fq, fq)),
            (337, 'Grafiek: positie van de laatste getoonde periode (Dashboard W90:X90, leeg = één na de laatste met opbrengsten of kosten)',
             'MAX($BY$336,MIN(MAX(1,$BY$8),IF(N(Dashboard!$W$90)=0,IFERROR(LOOKUP(2,1/%s,ROW(%s)-7)+1,$BY$8),'
             'IFERROR(LOOKUP(2,1/(ISNUMBER(%s)*(%s<=Dashboard!$W$90*4+IF(N(Dashboard!$X$90)=0,4,Dashboard!$X$90))),ROW(%s)-7),$BY$8))))'
             % (nz, fq, fq, fq, fq)),
            (338, 'Grafiek: aantal getoonde periodes', 'MAX(1,$BY$337-$BY$336+1)')):
        sh.set_value('BX%d' % r, lab, st_lab)
        sh.set_formula('BY%d' % r, formule, st_val)
    dst.put(part_of['Model'], sh.t)
    # Dashboard: cashflow vorige en huidige prognose per periode (de periode van de grafiek)
    sh = Sheet(dst.x(part_of['Dashboard']), sst)
    st_k, st_l, st_v, st_h = sh.style_of('V82'), sh.style_of('W83'), sh.style_of('W71'), sh.style_of('W56')
    # controle 'rij met cijfers maar zonder jaar': alleen de periodes (niet de totaalregels eronder)
    sh.set_formula('W43', 'IF(SUMPRODUCT((Invoer!$B$7:$B$%d="")*(Invoer!$C$7:$ZZ$%d<>""))=0,"OK",'
                          '"LET OP: rij met cijfers maar zonder jaar telt niet mee")' % (scenario_koppeling.INV_RN, scenario_koppeling.INV_RN))
    # samenvatting: vorige ↔ huidige prognose en prognose ↔ realisatie (details per periode op tab CF vorig en huidig)
    T = "'CF vorig en huidig'!"
    rng = lambda c: '%s$%s$8:$%s$%d' % (T, c, c, scenario_koppeling.INV_RN)
    idx = 'MATCH(p_actuals,Invoer!$A$8:$A$%d,0)' % scenario_koppeling.INV_RN
    sh.set_value('V93', 'VORIGE ↔ HUIDIGE PROGNOSE ↔ REALISATIE (€)', st_k)
    for c, h in zip(['V', 'W', 'X', 'Y'], ['', 'Vorige prognose', 'Huidig / gerealiseerd', 'Verschil']):
        sh.set_value('%s94' % c, h or None, st_h)
    regels = [
        ('Eindsaldo cashflow', 'IFERROR(LOOKUP(2,1/ISNUMBER(%s),%s)*1000,"")' % (rng('N'), rng('N')),
         'IFERROR(LOOKUP(2,1/ISNUMBER(%s),%s)*1000,"")' % (rng('M'), rng('M'))),
        ('Laagste stand (max. financiering)', 'IF(COUNT(%s)=0,"",MIN(%s)*1000)' % (rng('N'), rng('N')),
         'IF(COUNT(%s)=0,"",MIN(%s)*1000)' % (rng('M'), rng('M'))),
        (None, None, None),
        ('Kosten in dat kwartaal · vorige prognose ↔ gerealiseerd', 'IFERROR(INDEX(%s,%s),"")' % (rng('V'), idx),
         'IFERROR(INDEX(%s,%s),"")' % (rng('U'), idx)),
        ('Omzet in dat kwartaal · vorige prognose ↔ gerealiseerd', 'IFERROR(INDEX(%s,%s),"")' % (rng('Y'), idx),
         'IFERROR(INDEX(%s,%s),"")' % (rng('X'), idx)),
    ]
    for i, (lab, fw, fx) in enumerate(regels):
        r = 95 + i
        if lab is None:
            sh.set_formula('V%d' % r, '"Laatste afgesloten kwartaal (actuals): "&IF(N(p_actuals)=0,"nog geen actuals",'
                                      '"Q"&(p_actuals-4*INT((p_actuals-1)/4))&" "&INT((p_actuals-1)/4))', st_l)
            for c in 'WXY':
                sh.set_value('%s%d' % (c, r), None, st_v)
            continue
        sh.set_value('V%d' % r, lab, st_l)
        sh.set_formula('W%d' % r, fw, st_v)
        sh.set_formula('X%d' % r, fx, st_v)
        sh.set_formula('Y%d' % r, 'IF(OR(W%d="",X%d=""),"",X%d-W%d)' % (r, r, r, r), st_v)
    r = 95 + len(regels)
    sh.set_value('V%d' % r, 'Afgesloten periodes: cashflow = realisatie (SAP)', st_l)
    sh.set_formula('X%d' % r, 'IF(COUNTIF(%s,"LET OP*")=0,"OK",COUNTIF(%s,"LET OP*")&" periode(s) wijken af")' % (rng('AA'), rng('AA')),
                   sh.style_of('W41'))
    sh.set_formula('V%d' % (r + 1), 'HYPERLINK("#\'CF vorig en huidig\'!A1","→ details per periode: CF vorig en huidig")', st_l)
    # oude tabel (periode per rij) weg
    for rr in range(r + 2, 155):
        for c in ['V', 'W', 'X', 'Y', 'Z', 'AA']:
            sh.set_value('%s%d' % (c, rr), None)
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
        (9, 'Cijfers per (tab Start)', '=Start!$C$11', 'PEW_DATUM'),
    ]
    for i, (dia, onderdeel, f, vorm) in enumerate(rows):
        r = 118 + i
        assert sh.cell('C%d' % r, create=False) is None or not len(sh.cell('C%d' % r))
        sh.set_value('B%d' % r, dia, style['B'])
        sh.set_value('C%d' % r, onderdeel, style['C'])
        sh.set_formula('D%d' % r, f, style['D'])
        sh.set_value('F%d' % r, vorm, style['F'])
    # vervallen vormen: het label van het nu-blok op dia 2 bestaat sinds sjabloon v17 niet meer (anders meldt de knop
    # bij elke presentatie 'niet gevonden')
    for r in range(8, 400):
        if sh.get('F%d' % r) == 'KPI4_LABEL':
            sh.set_value('F%d' % r, None)
    # instellingen (sjabloon, naam nieuwe presentatie) van rij 130-131 naar 160-161: de tekstregels lopen nu door tot
    # voorbij rij 130 (de macro leest ze tot de eerste lege regel)
    for r_oud, r_nieuw in ((130, 160), (131, 161)):
        for c in 'CDE':
            sh.set('%s%d' % (c, r_nieuw), sh.get('%s%d' % (c, r_oud)), sh.style_of('%s%d' % (c, r_oud)))
            sh.set('%s%d' % (c, r_oud), None)
    sh.set('E160', 'leeg = Kwartaal_Template_cashflow_v22.pptx (of v21) in dezelfde map als dit bestand')
    powerpoint_v22(sh, style, 118 + len(rows), d10)
    tekst_vervang(sh, 'GE5', 'een kopie van dia 9', 'een kopie van de dia per woningtype')
    save('PowerPoint', sh)
    # Model: typekleuren van de legenda's in het palet van sjabloon v22 (zelfde mapping als de grafieken)
    import pptx_v22
    sh = sheet('Model')
    for r in (321, 323):
        f = sh.get('BY%d' % r)
        assert f and 'EDA100' in f and '2A78D6' in f, f
        sh.set_formula('BY%d' % r, re.sub(r'"([0-9A-F]{6}),"', lambda m: '"%s,"' % pptx_v22.PALET.get(m.group(1), m.group(1)), f))
    save('Model', sh)
    grafiekbereik_en_totaal(sheet, save)
    for blad, ref, oud, nieuw in (('Dashboard', 'E156', 'kopie van dia 8 in het sjabloon', 'kopie van de dia per woningtype'),
                                  ('Woningtypes', 'B42', 'Dashboard en dia 5', 'Dashboard en de dia Bouwtermijnen')):
        sh = sheet(blad)
        tekst_vervang(sh, ref, oud, nieuw)
        save(blad, sh)

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
    # grafieken over de hele looptijd: periode vanaf/t-m van het Dashboard (Model!BY336:BY338)
    for d in new_dn:
        if d.get('name', '').startswith('g_') and d.text and ',0,0,MAX(1,Model!$BY$8),1)' in d.text:
            d.text = d.text.replace(',0,0,MAX(1,Model!$BY$8),1)', ',Model!$BY$336-1,0,MAX(1,Model!$BY$338),1)')
    # Dashboard-grafieken verkoop en transport: de reeksen lezen het blok op tab PowerPoint (ER7:FW68), dat begint met de
    # totaalrij (alle woningen per type) en hetzelfde venster volgt; Model!BY339 = totaalrij + getoonde periodes
    from openpyxl.utils import get_column_letter as L_, column_index_from_string as ci_
    vp = {'D': 'ER', 'BG': 'FM', 'BH': 'FN', 'BT': 'FO', 'BU': 'FP', 'BS': 'FQ', 'BK': 'FR', 'BO': 'FS', 'BL': 'FT', 'BP': 'FU',
          'MD': 'FV', 'ME': 'FW'}
    for i in range(20):
        vp[L_(ci_('EI') + i)] = L_(ci_('ES') + i)
    n_vp = 0
    for d in new_dn:
        m = re.fullmatch(r'OFFSET\(Model!\$([A-Z]+)\$8,Model!\$BY\$131-1,0,MAX\(1,Model!\$BY\$132\),1\)', d.text or '')
        if m and re.match(r'g_(vk_|v\d|t\d)', d.get('name', '')):
            d.text = 'OFFSET(PowerPoint!$%s$8,0,0,MAX(1,Model!$BY$339),1)' % vp[m.group(1)]
            n_vp += 1
    assert n_vp >= 30, n_vp
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
    # tabbladkleuren per soort: start (donkerblauw), exports (grijs), invoer en rekenwerk (geel), uitkomsten (groen)
    TABKLEUR = {'Start': 'FF003366',
                'FO - vorig': 'FF9CA3AF', 'FO - huidig': 'FF9CA3AF', 'FO - actuals': 'FF9CA3AF', 'Prijslijst dump': 'FF9CA3AF',
                'Prijslijst': 'FFFEB900', 'CF - opbrengsten': 'FFFEB900', 'CF - kosten': 'FFFEB900', '2. Ontw PEW': 'FFFEB900',
                '2B. Ontw PEW': 'FFFEB900',
                'Dashboard': 'FF15803D', "Scenario's": 'FF15803D', 'CF vorig en huidig': 'FF15803D', 'CF - hulp VENA': 'FF15803D'}
    for naam, kleur in TABKLEUR.items():
        if naam not in part_of:
            continue
        t = dst.x(part_of[naam])
        sp = t.find(M + 'sheetPr')
        if sp is None:
            sp = etree.Element(M + 'sheetPr')
            t.insert(0, sp)
        for oud in sp.findall(M + 'tabColor'):
            sp.remove(oud)
        tc = etree.Element(M + 'tabColor', rgb=kleur)
        sp.insert(0, tc)
        dst.put(part_of[naam], t)
    dst.parts.pop('xl/calcChain.xml', None)
    copier.remove_override('xl/calcChain.xml')
    dst.put('[Content_Types].xml', copier.ct)
    sst.finish()
    dst.put('xl/sharedStrings.xml', sst.t)
    n_kop, n_inv = harmoniseer_stijlen(styles.t)
    note('Vormgeving: %d kopvullingen donkerblauw, %d invoervullingen geel met rand' % (n_kop, n_inv))
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

    note('Palet (zoals sjabloon v22): %d kleuren omgezet' % palet_werkbestand(dst))
    dst.save(a.out)
    os.remove(fo_tmp)
    import json
    with open(os.path.splitext(a.out)[0] + '.opzet.json', 'w') as fh:
        json.dump({'fo': os.path.basename(fo_bron), 'oud_fo': os.path.abspath(a.oud) if a.oud else None, 'opzet': oz, 'invoer_laatste': scenario_koppeling.INV_RN,
                   'invoer_totaal': inv_tot, 'daeb_type': scenario_koppeling.DAEB_TYPE}, fh, indent=1, default=str)
    print('Klaar: ' + a.out)

    # ---- PowerPoint-sjabloon v22: v20 + vormnamen (v21) + strakker (pptx_v22) -----------------------------------
    maak_sjabloon(a.pptx, a.pptx_out)
    print('Klaar: ' + a.pptx_out)


def grafiekbereik_en_totaal(sheet, save):
    """1) 'Grafiek vanaf / t/m' van het Dashboard (Model!BY336:BY338) geldt ook voor de dia's: de blokken over de hele
       looptijd (scenario's P, verkoop naar omzet FY, cashflow HH) beginnen bij de eerste getoonde periode, en de vensters
       van verkoop/transport, verkoopscenario's (BY131/BY132) en de grafiek per type (BY127/BY128) blijven binnen het bereik.
       De kleine grafiek op dia 2 houdt de hele looptijd (de macro schikt hem op positie).
    2) verkoop en transport: eerste rij 'Totaal' met alle woningen per type, de kwartalen eronder (blok ER7 een rij lager);
       titels met het totaal en de legenda met het aantal per type"""
    from openpyxl.utils import get_column_letter as L, column_index_from_string as ci
    sh = sheet('Model')
    for start, aantal in ((131, 132), (127, 128)):
        f = sh.cell('BY%d' % start).find(M + 'f')
        assert f is not None and f.text and '$BY$336' not in f.text, start
        f.text = 'MAX(%s,$BY$336)' % f.text                       # (matrixformule: de tekst aanpassen, het type blijft)
        assert sh.get('BY%d' % aantal) == '=MAX(1,$BY$8-$BY$%d+1)' % start, sh.get('BY%d' % aantal)
        sh.set_formula('BY%d' % aantal, 'MAX(1,MIN($BY$8,$BY$337)-$BY$%d+1)' % start, sh.style_of('BY%d' % aantal))
    assert not sh.get('BX339') and not sh.get('BX340')
    # verkoop en transport: de as stopt aan het eind van het jaar van de laatste verkoop / het laatste transport (BY340);
    # de verkoopscenario's houden het ruimere venster BY132
    klaar = 'MIN($BY$8,MAX(COUNTIF($BK$8:$BK$67,"<"&0.99999),COUNTIF($BL$8:$BL$67,"<"&0.99999))+1)'
    idx = 'INDEX($F$8:$F$67,%s)' % klaar
    sh.set_value('BX340', 'Verkoopgrafiek: aantal getoonde periodes (t/m Q4 van het jaar van de laatste verkoop / het laatste transport)',
                 sh.style_of('BX338'))
    sh.set_formula('BY340', 'IF(OR($BY$8=0,$BY$39=0),$BY$132,MAX(1,MIN($BY$132,IFERROR(%s+4*(INT((%s-1)/4)+1)-%s,$BY$8)-$BY$131+1)))'
                   % (klaar, idx, idx), sh.style_of('BY338'))
    sh.set_value('BX339', 'Verkoopgrafiek in PowerPoint: aantal rijen (totaalrij + periodes)', sh.style_of('BX338'))
    sh.set_formula('BY339', '$BY$340+1', sh.style_of('BY338'))
    # bouwtermijnen: de tijdas stopt aan het eind van het jaar van de laatste termijn
    oud = sh.get('BY101')
    assert oud == '=IF($BY$8=0,0,INT((INDEX($F$8:$F$67,$BY$8)-1)/4))', oud
    laatste = 'SUMPRODUCT(MAX(($GA$8:$GA$140=1)*($GK$8:$GK$140<99999)*$GK$8:$GK$140))'
    sh.set_formula('BY101', 'IF($BY$8=0,0,IF(IFERROR(%s,0)=0,INT((INDEX($F$8:$F$67,$BY$8)-1)/4),MIN(INT((INDEX($F$8:$F$67,$BY$8)-1)/4),'
                            'INT((%s-1)/4))))' % (laatste, laatste), sh.style_of('BY101'))
    sh.set_formula('BY324', '"Verkoop per kwartaal · "&$BY$39&" woningen"', sh.style_of('BY324'))
    sh.set_formula('BY325', '"Transport per kwartaal (notarieel) · "&$BY$39&" woningen"', sh.style_of('BY325'))
    f = sh.get('BY322')
    for i in range(10):
        oud = '&%s$7&"   "' % L(ci('EI') + i)
        assert f.count(oud) == 1, oud
        f = f.replace(oud, '&%s$7&" ("&%s$6&")   "' % (L(ci('EI') + i), L(ci('CA') + i)))
    sh.set_formula('BY322', f, sh.style_of('BY322'))
    save('Model', sh)
    sh = sheet('PowerPoint')
    for a, b in (('P', 'Y'), ('FY', 'GC'), ('HH', 'HR')):
        for c in range(ci(a), ci(b) + 1):
            for r in range(8, 68):
                n, ref = r - 7, '%s%d' % (L(c), r)
                f = sh.get(ref)
                kop = '=IF(%d>Model!$BY$8,' % n
                assert f and f.startswith(kop) and f.count('$67,%d)' % n) == 1, (ref, f)
                f = '=IF(%d>Model!$BY$338,' % n + f[len(kop):].replace('$67,%d)' % n, '$67,%d+Model!$BY$336-1)' % n)
                sh.set_formula(ref, f, sh.style_of(ref))
    # totaalrij: de periodes een rij omlaag (van onder naar boven), daarna rij 8 = totaal per type
    k1, kn = ci('ER'), ci('FW')
    for c in range(k1, kn + 1):
        assert not sh.get('%s68' % L(c))
        for r in range(68, 8, -1):
            f = sh.get('%s%d' % (L(c), r - 1))
            assert f.count('>Model!$BY$132,') == 1, f
            sh.set('%s%d' % (L(c), r), f.replace('>Model!$BY$132,', '>Model!$BY$340,'), sh.style_of('%s%d' % (L(c), r - 1)))
    sh.set_value('ER8', 'Totaal', sh.style_of('ER8'))
    for c in range(k1 + 1, kn + 1):
        i = c - k1 - 1
        if i < 20:              # verkocht per type (10) en getransporteerd per type (10): alle woningen van het type
            f = 'IF(N(Model!$%s$6)=0,NA(),Model!$%s$6)' % ((L(ci('CA') + i % 10),) * 2)
        else:
            f = 'NA()'
        sh.set_formula('%s8' % L(c), f, sh.style_of('%s8' % L(c)))
    save('PowerPoint', sh)


def tekst_vervang(sh, ref, oud, nieuw):
    t = sh.get(ref)
    assert isinstance(t, str) and oud in t and not t.startswith('='), (ref, t)
    sh.set_value(ref, t.replace(oud, nieuw), sh.style_of(ref))


def powerpoint_v22(sh, style, r0, d10):
    """tab PowerPoint voor sjabloon v22 (9 dia's): kansblok op dia 3, PEW_SUBTITEL op de dia Update PEW (het overzicht
    zelf plakt de macro als afbeelding), geen scenariolijnen 1-3 in de cashflowgrafiek, accentkleuren; de regels van de
    vervallen dia 'Van verkoop naar omzet' (VO_*) gaan weg"""
    from openpyxl.utils import get_column_letter as L
    import pptx_v22
    D = 'Dashboard!'
    W = lambda c, r: '%s$%s$%d' % (D, c, r)
    P = lambda c, r: "'2. Ontw PEW'!$%s$%d" % (c, r)
    pew_d = '(%s-%s)' % (P('G', 29), P('F', 29))
    fd, basis = 'Start!$C$%d&""' % d10, 'Start!$C$%d&""' % (d10 + 1)
    rows = [
        (3, 'Blokje kans grondtermijn · label', 'KANS GRONDTERMIJN', '', 'CF_KANS_LABEL'),
        (9, 'Ondertitel', '="Resultaat "&IF(%s<0,"−","")&"€"&FIXED(ABS(%s)/1000,1)&" mln ("&FIXED(%s*100,1)&"%%) · "&'
                          'IF(ABS(%s)<50,"gelijk aan vorig kwartaal",IF(%s>0,"+","−")&"€"&FIXED(ABS(%s)/1000,1)&'
                          '" mln t.o.v. vorig kwartaal")&IF(AND(Model!$BY$135>0,COUNT(\'2B. Ontw PEW\'!$F$7:$G$26)>0),'
                          '" · niet-DAEB "&FIXED(\'2B. Ontw PEW\'!$G$30*100,1)&"%% · DAEB "&FIXED(\'2B. Ontw PEW\'!$F$30*100,1)&"%%","")'
                          '&IF(%s="",""," · FD: "&%s)&IF(%s="",""," · basis: "&%s)'
         % (P('G', 29), P('G', 29), P('G', 30), pew_d, pew_d, pew_d, fd, fd, basis, basis), '', 'PEW_SUBTITEL'),
    ]
    for i, (dia, onderdeel, f, richting, vorm) in enumerate(rows):
        r = r0 + i
        assert sh.cell('C%d' % r, create=False) is None or not len(sh.cell('C%d' % r))
        sh.set_value('B%d' % r, dia, style['B'])
        sh.set_value('C%d' % r, onderdeel, style['C'])
        sh.set('D%d' % r, f, style['D'])
        sh.set_value('F%d' % r, vorm, style['F'])
    r_eind = r0 + len(rows)
    for r in range(8, r0):
        vorm = sh.get('F%d' % r)
        if vorm == 'CF_BE_WAARDE':          # dia 3: het vierde blok is de kans grondtermijn (break-even staat al op dia 2)
            sh.set_value('C%d' % r, 'Blokje kans grondtermijn')
            sh.set('D%d' % r, '=IF(AND(N(%s)=0,N(%s)=0),"geen kans ingesteld",IF(%s<0,"−","+")&"€"&FIXED(ABS(%s)/1000000,1)'
                              '&"M eerder")' % (W('Z', 85), W('Z', 86), W('Z', 85), W('Z', 85)))
            sh.set_value('F%d' % r, 'CF_KANS_WAARDE')
        elif vorm in ('VP_STAT1_ACCENT', 'VP_STAT3_ACCENT'):
            sh.set_value('D%d' % r, '003366')            # tekst (sh.set zou er een getal van maken)
        elif vorm == 'VP_STAT2_ACCENT':
            sh.set('D%d' % r, '=IF(Model!$BY$326=1,"FEB900","003366")')
    # de regels van de vervallen dia weg en de rest aansluiten (de macro leest tot de eerste lege regel); dianummers:
    # het scenariomodel telt nog met 'Van verkoop naar omzet' als dia 8 (per woningtype 9 -> 8)
    regels = []
    for r in range(8, r_eind):
        w = {c: (sh.get('%s%d' % (c, r)), sh.style_of('%s%d' % (c, r))) for c in 'BCDEF'}
        if (w['F'][0] or '').startswith('VO_'):
            continue
        if r < 118 and str(w['B'][0] or '').strip() == '9':
            w['B'] = ('8', w['B'][1])
        regels.append(w)
    for i in range(r_eind - 8):
        r = 8 + i
        for c in 'BCDEF':
            v, st = regels[i][c] if i < len(regels) else (None, None)
            ref = '%s%d' % (c, r)
            if v is None or v == '':
                sh.set_value(ref, None)
            elif isinstance(v, str) and v.startswith('='):
                sh.set_formula(ref, v, st)
            elif c == 'B' and str(v).strip().isdigit():
                sh.set_value(ref, int(v), st)
            else:
                sh.set_value(ref, str(v), st)
    # cashflowgrafiek (dia 3): geen scenariolijnen 1-3 (die staan op de scenariodia's); de gele lijn blijft
    for c, n in (('HP', 1), ('HQ', 2), ('HR', 3)):
        sh.set_value('%s7' % c, 'Scenario %d (uit)' % n)


def maak_sjabloon(src, out):
    """v22 = v20 met namen voor de vormen die de macro vult (v21), daarna strakker gemaakt (tools/pptx_v22.py)"""
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
    import pptx_v22
    pptx_v22.maak(p)
    p.save(out)

if __name__ == '__main__':
    main()
