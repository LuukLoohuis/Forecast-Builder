"""Werkbladen tussen .xlsx/.xlsm-pakketten samenvoegen op XML-niveau.

openpyxl gooit grafieken, knoppen en uitbreidingen weg; daarom werkt dit direct op de XML:
stijlen en gedeelde teksten worden samengevoegd en omgenummerd, tekeningen en grafieken gaan mee.
"""
from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from collections import OrderedDict

from lxml import etree

NS_M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
M = '{%s}' % NS_M
R = '{%s}' % NS_R
PR = '{%s}' % NS_PR
CT = '{%s}' % NS_CT

T_SHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
T_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
T_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
CT_SHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
CT_CHARTSTYLE = 'application/vnd.ms-office.chartstyle+xml'
CT_CHARTCOLORS = 'application/vnd.ms-office.chartcolorstyle+xml'

PARSER = etree.XMLParser(remove_blank_text=False, huge_tree=True)


def parse(b: bytes):
    return etree.fromstring(b, PARSER)


def tobytes(el) -> bytes:
    return etree.tostring(el, xml_declaration=True, encoding='UTF-8', standalone=True)


def rels_name(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, '_rels', f + '.rels')


def resolve(base_part: str, target: str) -> str:
    if target.startswith('/'):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def relative(base_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(base_part))


# ----------------------------------------------------------------------------------------------
#  celverwijzingen
# ----------------------------------------------------------------------------------------------

def col2num(c: str) -> int:
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref: str):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    return m.group(1), int(m.group(2))


# ----------------------------------------------------------------------------------------------
#  pakket
# ----------------------------------------------------------------------------------------------

class Package:
    def __init__(self, path: str):
        self.parts: OrderedDict[str, bytes] = OrderedDict()
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                self.parts[n] = z.read(n)

    def save(self, path: str):
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', self.parts['[Content_Types].xml'])
            for n, b in self.parts.items():
                if n != '[Content_Types].xml':
                    z.writestr(n, b)

    def x(self, name):
        return parse(self.parts[name])

    def put(self, name, tree_or_bytes):
        self.parts[name] = tree_or_bytes if isinstance(tree_or_bytes, bytes) else tobytes(tree_or_bytes)

    def rels(self, part):
        n = rels_name(part)
        if n not in self.parts:
            return []
        return [(r.get('Id'), r.get('Type'), r.get('Target'), r.get('TargetMode'))
                for r in self.x(n).findall(PR + 'Relationship')]

    def sheets(self):
        """[(naam, part, state)] in volgorde van de werkmap"""
        wb = self.x('xl/workbook.xml')
        rel = {i: resolve('xl/workbook.xml', t) for i, ty, t, _ in self.rels('xl/workbook.xml')}
        out = []
        for s in wb.find(M + 'sheets'):
            out.append((s.get('name'), rel[s.get(R + 'id')], s.get('state')))
        return out

    def sheet_part(self, name):
        for n, p, _ in self.sheets():
            if n == name:
                return p
        raise KeyError(name)


# ----------------------------------------------------------------------------------------------
#  stijlen
# ----------------------------------------------------------------------------------------------

def _key(el) -> bytes:
    return etree.tostring(el, method='c14n')


class Styles:
    LISTS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs']

    def __init__(self, tree):
        self.t = tree
        for name in ('fonts', 'fills', 'borders', 'cellXfs', 'dxfs'):
            if self.t.find(M + name) is None:
                el = etree.Element(M + name, count='0')
                self._insert(el)
        self.index = {}
        for name in ('fonts', 'fills', 'borders', 'cellXfs', 'dxfs'):
            self.index[name] = {}
            for i, el in enumerate(self.t.find(M + name)):
                self.index[name].setdefault(_key(el), i)
        nf = self.t.find(M + 'numFmts')
        self.numfmts = {}
        if nf is not None:
            for el in nf:
                self.numfmts[el.get('formatCode')] = int(el.get('numFmtId'))

    def _insert(self, el):
        order = self.LISTS
        pos = order.index(etree.QName(el).localname)
        for i, ch in enumerate(self.t):
            ln = etree.QName(ch).localname
            if ln in order and order.index(ln) > pos or ln in ('tableStyles', 'colors', 'extLst'):
                self.t.insert(i, el)
                return
        self.t.append(el)

    def _add(self, name, el) -> int:
        k = _key(el)
        if k in self.index[name]:
            return self.index[name][k]
        lst = self.t.find(M + name)
        lst.append(el)
        i = len(lst) - 1
        self.index[name][k] = i
        lst.set('count', str(len(lst)))
        return i

    def add_dxf(self, xml: str) -> int:
        """voorwaardelijke opmaak: <font>/<fill>-xml -> dxfId"""
        el = etree.fromstring('<dxf xmlns="%s">%s</dxf>' % (NS_M, xml))
        return self._add('dxfs', el)

    def numfmt(self, code: str) -> int:
        if code in self.numfmts:
            return self.numfmts[code]
        nf = self.t.find(M + 'numFmts')
        if nf is None:
            nf = etree.Element(M + 'numFmts', count='0')
            self._insert(nf)
        nid = max([163] + list(self.numfmts.values())) + 1
        e = etree.SubElement(nf, M + 'numFmt', numFmtId=str(nid), formatCode=code)
        nf.set('count', str(len(nf)))
        self.numfmts[code] = nid
        return nid

    def merge(self, src_tree):
        """voegt de stijlen van een ander pakket toe; geeft (xf_map, dxf_map)"""
        src_nf = {}
        nf = src_tree.find(M + 'numFmts')
        if nf is not None:
            for el in nf:
                src_nf[int(el.get('numFmtId'))] = el.get('formatCode')

        def map_nf(i):
            i = int(i)
            return str(self.numfmt(src_nf[i]) if i in src_nf else i)

        maps = {}
        for name in ('fonts', 'fills', 'borders'):
            maps[name] = [self._add(name, copy.deepcopy(el)) for el in src_tree.find(M + name)]
        xf_map = []
        for el in src_tree.find(M + 'cellXfs'):
            e = copy.deepcopy(el)
            e.set('numFmtId', map_nf(e.get('numFmtId', '0')))
            e.set('fontId', str(maps['fonts'][int(e.get('fontId', '0'))]))
            e.set('fillId', str(maps['fills'][int(e.get('fillId', '0'))]))
            e.set('borderId', str(maps['borders'][int(e.get('borderId', '0'))]))
            e.set('xfId', '0')
            xf_map.append(self._add('cellXfs', e))
        dxf_map = []
        dx = src_tree.find(M + 'dxfs')
        if dx is not None:
            for el in dx:
                dxf_map.append(self._add('dxfs', copy.deepcopy(el)))
        return xf_map, dxf_map


# ----------------------------------------------------------------------------------------------
#  gedeelde teksten
# ----------------------------------------------------------------------------------------------

class SST:
    def __init__(self, tree):
        self.t = tree
        self.index = {}
        for i, si in enumerate(self.t.findall(M + 'si')):
            self.index.setdefault(_key(si), i)

    def add_si(self, si) -> int:
        k = _key(si)
        if k in self.index:
            return self.index[k]
        self.t.append(copy.deepcopy(si))
        i = len(self.t) - 1
        self.index[k] = i
        return i

    def add(self, text: str) -> int:
        si = etree.Element(M + 'si')
        t = etree.SubElement(si, M + 't')
        t.text = text
        if text != text.strip():
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return self.add_si(si)

    def finish(self):
        n = len(self.t.findall(M + 'si'))
        self.t.set('count', str(n))
        self.t.set('uniqueCount', str(n))


# ----------------------------------------------------------------------------------------------
#  werkblad bewerken
# ----------------------------------------------------------------------------------------------

SHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
               'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
               'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins',
               'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors',
               'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls',
               'webPublishItems', 'tableParts', 'extLst']


class Sheet:
    """een werkblad-XML met hulpjes om cellen te zetten"""

    def __init__(self, tree, sst: SST):
        self.t = tree
        self.sst = sst
        self.data = tree.find(M + 'sheetData')

    def row(self, r: int, create=True):
        prev = None
        for el in self.data:
            n = int(el.get('r'))
            if n == r:
                return el
            if n > r:
                break
            prev = el
        if not create:
            return None
        new = etree.Element(M + 'row', r=str(r))
        if prev is None:
            self.data.insert(0, new)
        else:
            prev.addnext(new)
        return new

    def cell(self, ref: str, create=True):
        col, r = split_ref(ref)
        row = self.row(r, create)
        if row is None:
            return None
        cn = col2num(col)
        prev = None
        for c in row:
            cc = col2num(split_ref(c.get('r'))[0])
            if cc == cn:
                return c
            if cc > cn:
                break
            prev = c
        if not create:
            return None
        new = etree.Element(M + 'c', r=ref)
        if prev is None:
            row.insert(0, new)
        else:
            prev.addnext(new)
        return new

    def style_of(self, ref):
        c = self.cell(ref, create=False)
        return None if c is None else c.get('s')

    def _clear(self, c):
        for ch in list(c):
            c.remove(ch)
        if 't' in c.attrib:
            del c.attrib['t']

    def set_formula(self, ref, formula: str, style=None):
        c = self.cell(ref)
        f = c.find(M + 'f')
        if f is not None and f.get('t') == 'shared' and f.get('ref'):
            raise ValueError('%s is de bron van een gedeelde formule' % ref)
        self._clear(c)
        f = etree.SubElement(c, M + 'f')
        f.text = formula[1:] if formula.startswith('=') else formula
        if style is not None:
            c.set('s', str(style))

    def set_value(self, ref, value, style=None):
        c = self.cell(ref)
        f = c.find(M + 'f')
        if f is not None and f.get('t') == 'shared' and f.get('ref'):
            raise ValueError('%s is de bron van een gedeelde formule' % ref)
        self._clear(c)
        if value is None:
            pass
        elif isinstance(value, str):
            c.set('t', 's')
            etree.SubElement(c, M + 'v').text = str(self.sst.add(value))
        else:
            etree.SubElement(c, M + 'v').text = repr(float(value)) if isinstance(value, float) else str(value)
        if style is not None:
            c.set('s', str(style))

    def get(self, ref):
        """inhoud van een cel als tekst zoals je hem typt: '=formule', een getal of een tekst; None = leeg"""
        c = self.cell(ref, create=False)
        if c is None:
            return None
        f = c.find(M + 'f')
        if f is not None and f.text:
            return '=' + f.text
        t = c.get('t')
        if t == 'inlineStr':
            return ''.join(c.find(M + 'is').itertext())
        v = c.find(M + 'v')
        if v is None:
            return None
        if t == 's':
            return ''.join(self.sst.t.findall(M + 'si')[int(v.text)].itertext())
        return v.text

    def set(self, ref, inhoud, style=None):
        """zet '=formule', een getal (als tekst of getal) of een tekst; None/'' = leegmaken (opmaak blijft)"""
        if inhoud is None or inhoud == '':
            self.set_value(ref, None, style)
        elif isinstance(inhoud, str) and inhoud.startswith('='):
            self.set_formula(ref, inhoud, style)
        else:
            try:
                x = float(inhoud)
                self.set_value(ref, int(x) if x == int(x) and 'e' not in str(inhoud).lower() else x, style)
            except (TypeError, ValueError):
                self.set_value(ref, str(inhoud), style)

    def add_element(self, el):
        """element op de juiste plek in de werkblad-XML (de volgorde ligt vast in het schema)"""
        name = etree.QName(el).localname
        pos = SHEET_ORDER.index(name)
        for i, ch in enumerate(self.t):
            ln = etree.QName(ch).localname
            if ln in SHEET_ORDER and SHEET_ORDER.index(ln) > pos:
                self.t.insert(i, el)
                return el
        self.t.append(el)
        return el

    def add_list_validation(self, sqref, items):
        dvs = self.t.find(M + 'dataValidations')
        if dvs is None:
            dvs = self.add_element(etree.Element(M + 'dataValidations', count='0'))
        dv = etree.SubElement(dvs, M + 'dataValidation', type='list', allowBlank='1', showErrorMessage='1', sqref=sqref)
        etree.SubElement(dv, M + 'formula1').text = '"%s"' % ','.join(items)
        dvs.set('count', str(len(dvs)))

    def add_cf(self, sqref, formula, dxf_id):
        prio = max([int(r.get('priority', '0')) for r in self.t.iter(M + 'cfRule')] + [0]) + 1
        cf = etree.Element(M + 'conditionalFormatting', sqref=sqref)
        rule = etree.SubElement(cf, M + 'cfRule', type='expression', dxfId=str(dxf_id), priority=str(prio))
        etree.SubElement(rule, M + 'formula').text = formula
        existing = self.t.findall(M + 'conditionalFormatting')
        if existing:
            existing[-1].addnext(cf)
        else:
            self.add_element(cf)

    def formulas(self):
        for f in self.t.iter(M + 'f'):
            yield f


def replace_in_formulas(tree, old: str, new: str) -> int:
    n = 0
    for f in tree.iter(M + 'f', '{http://schemas.openxmlformats.org/drawingml/2006/chart}f',
                       '{http://schemas.microsoft.com/office/excel/2006/main}f'):
        if f.text and old in f.text:
            f.text = f.text.replace(old, new)
            n += 1
    return n


def unshare(tree) -> int:
    """gedeelde formules (t="shared") uitschrijven tot gewone formules per cel, zodat elke cel zijn eigen formuletekst
    heeft (de aanpassingen lezen en zetten per cel)"""
    from openpyxl.formula.translate import Translator
    bron = {}
    for f in tree.iter(M + 'f'):
        if f.get('t') == 'shared' and f.get('ref') is not None and f.text:
            bron[f.get('si')] = (f.getparent().get('r'), '=' + f.text)
    n = 0
    for f in tree.iter(M + 'f'):
        if f.get('t') != 'shared':
            continue
        si = f.get('si')
        r = f.getparent().get('r')
        if si in bron:
            oorsprong, tekst = bron[si]
            if not f.text:
                f.text = Translator(tekst, origin=oorsprong).translate_formula(r)[1:]
            for a in ('t', 'si', 'ref'):
                f.attrib.pop(a, None)
            n += 1
    return n



def formuleteksten(tree) -> dict:
    """{celnaam: formuletekst} van alle cellen met een eigen formuletekst (voor de controle van share / unshare)"""
    return {f.getparent().get('r'): f.text for f in tree.iter(M + 'f') if f.text}


def share(tree) -> int:
    """het omgekeerde van unshare: buren met dezelfde formule (verschoven) als gedeelde formule opslaan, zoals Excel een
    doorgetrokken formule bewaart: de eerste cel houdt de tekst (t="shared" ref si), de andere alleen de verwijzing (si).
    Scheelt veel XML (kleiner bestand, sneller openen). Eerst reeksen langs een kolom, dan voor wat overblijft langs
    een rij. Een cel doet alleen mee als de formule van de eerste cel, verschoven naar die cel (dezelfde vertaler als
    unshare), letterlijk gelijk is aan zijn eigen tekst. Matrixformules, formules met kenmerken en bestaande gedeelde
    formules blijven zoals ze zijn. Geeft het aantal cellen terug dat geen eigen formuletekst meer heeft."""
    from openpyxl.formula.translate import Translator
    cellen, si = {}, -1
    for f in tree.iter(M + 'f'):
        if f.get('t') == 'shared' and f.get('si') is not None:
            si = max(si, int(f.get('si')))
        elif not f.attrib and f.text:
            k, r = split_ref(f.getparent().get('r'))
            cellen[(col2num(k), r)] = f
    naam = lambda k, r: '%s%d' % (num2col(k), r)
    gedaan, n = set(), 0
    for dk, dr, volgorde in ((0, 1, sorted(cellen)), (1, 0, sorted(cellen, key=lambda x: (x[1], x[0])))):
        for k, r in volgorde:
            if (k, r) in gedaan or (k + dk, r + dr) not in cellen or (k + dk, r + dr) in gedaan:
                continue
            vertaler = Translator('=' + cellen[(k, r)].text, origin=naam(k, r))
            leden, kk, rr = [], k + dk, r + dr
            while (kk, rr) in cellen and (kk, rr) not in gedaan:
                try:
                    zelfde = vertaler.translate_formula(naam(kk, rr))[1:] == cellen[(kk, rr)].text
                except Exception:
                    zelfde = False
                if not zelfde:
                    break
                leden.append((kk, rr))
                kk, rr = kk + dk, rr + dr
            if not leden:
                continue
            si += 1
            bron = cellen[(k, r)]
            bron.set('t', 'shared')
            bron.set('ref', '%s:%s' % (naam(k, r), naam(*leden[-1])))
            bron.set('si', str(si))
            gedaan.add((k, r))
            for lid in leden:
                f = cellen[lid]
                f.text = None
                f.set('t', 'shared')
                f.set('si', str(si))
                gedaan.add(lid)
            n += len(leden)
    return n


# ------------------------------------------------------------------------------------------------
#  rijen invoegen (zoals Excel: alle verwijzingen in het hele bestand schuiven mee)
# ------------------------------------------------------------------------------------------------

_REF = re.compile(
    r"(\"(?:[^\"]|\"\")*\")"                                              # 1 tekst tussen aanhalingstekens
    r"|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Z]{1,3}(?::\$?[A-Z]{1,3})?(?![\w$]))"   # 2 kolombereik met blad: overslaan
    r"|(?<![\w$.!'\]])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?"             # 3 bladnaam
    r"(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![\w(!])")        # 4, 5 cel of bereik


def _schuif_cel(ref, at, n):
    m = re.match(r'(\$?[A-Z]{1,3})(\$?)(\d+)$', ref)
    r = int(m.group(3))
    return '%s%s%d' % (m.group(1), m.group(2), r + n if r >= at else r)


def schuif_formule(text, blad, eigen, at, n):
    """verwijzingen naar rij >= at op blad `blad` n rijen omlaag; eigen = de formule staat op dat blad zelf"""
    if not text:
        return text

    def rep(m):
        if m.group(1) or m.group(2):
            return m.group(0)
        pre = m.group(3)
        if pre:
            naam = pre[:-1]
            if naam.startswith("'"):
                naam = naam[1:-1].replace("''", "'")
            if naam != blad:
                return m.group(0)
        elif not eigen:
            return m.group(0)
        a = _schuif_cel(m.group(4), at, n)
        b = ':' + _schuif_cel(m.group(5), at, n) if m.group(5) else ''
        return (pre or '') + a + b
    return _REF.sub(rep, text)


def _schuif_bereik(sq, at, n):
    return ' '.join(':'.join(_schuif_cel(x, at, n) if re.match(r'\$?[A-Z]{1,3}\$?\d+$', x) else x for x in part.split(':'))
                    for part in sq.split())


def rijen_invoegen(pkg, blad, at, n):
    """n lege rijen invoegen vóór rij `at` op blad `blad` (gedeelde formules eerst uitschrijven)"""
    doel = pkg.sheet_part(blad)
    for naam, part, _ in pkg.sheets():
        t = pkg.x(part)
        unshare(t)
        eigen = part == doel
        for el in t.iter():
            if not isinstance(el.tag, str):
                continue
            ln = etree.QName(el).localname
            if ln in ('f', 'formula', 'formula1', 'formula2') and el.text:
                el.text = schuif_formule(el.text, blad, eigen, at, n)
            if not eigen:
                continue
            if ln == 'sqref' and el.text:
                el.text = _schuif_bereik(el.text, at, n)
            for attr in ('sqref', 'ref'):
                if attr in el.attrib and ln not in ('c',):
                    el.set(attr, _schuif_bereik(el.get(attr), at, n))
            if ln == 'row':
                r = int(el.get('r'))
                if r >= at:
                    el.set('r', str(r + n))
            elif ln == 'c' and el.get('r'):
                el.set('r', _schuif_cel(el.get('r'), at, n))
            elif ln == 'brk' and el.getparent() is not None and etree.QName(el.getparent()).localname == 'rowBreaks':
                if int(el.get('id')) >= at:
                    el.set('id', str(int(el.get('id')) + n))
        pkg.put(part, t)
        if eigen:
            for rid, ty, target, _ in pkg.rels(part):
                if ty == T_DRAWING:
                    d = resolve(part, target)
                    dt = pkg.x(d)
                    for el in dt.iter():
                        if isinstance(el.tag, str) and etree.QName(el).localname == 'row' and el.text and \
                                etree.QName(el.getparent()).localname in ('from', 'to') and int(el.text) >= at - 1:
                            el.text = str(int(el.text) + n)
                    pkg.put(d, dt)
    wb = pkg.x('xl/workbook.xml')
    for d in wb.iter(M + 'definedName'):
        d.text = schuif_formule(d.text, blad, False, at, n)
    pkg.put('xl/workbook.xml', wb)
    for p in [k for k in pkg.parts if k.startswith('xl/charts/chart')]:
        t = pkg.x(p)
        for f in t.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}f'):
            f.text = schuif_formule(f.text, blad, False, at, n)
        pkg.put(p, t)
    pkg.parts.pop('xl/calcChain.xml', None)
