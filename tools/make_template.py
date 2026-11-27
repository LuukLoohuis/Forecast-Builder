"""Lege PowerPoint-template in RDFO/Heijmans-stijl (zelfde opmaak als 'Maak PowerPoint'):
model met navy lijn, gele accentbalk bij de titel, voetregel, paginanummer en logo; indelingen Titeldia,
Titel en inhoud, Twee kolommen, Alleen titel en Leeg. usage: make_template.py logo_navy.png uit.pptx"""
import sys, copy
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.oxml.ns import qn
from lxml import etree

NAVY, GEEL, GRIJS, LIJN = '111F39', 'FEB900', '6B7385', 'D9DEE8'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
logo, out = sys.argv[1:3]
prs = Presentation(); prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
E = lambda pt: int(pt * 12700)
_id = [100]

def nid():
    _id[0] += 1; return _id[0]

def rect(x, y, w, h, kleur, naam):
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{nid()}" name="{naam}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{kleur}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="nl-NL"/></a:p></p:txBody></p:sp>')

def lijn(x1, y, x2, kleur, naam):
    return etree.fromstring(
        f'<p:cxnSp xmlns:p="{P}" xmlns:a="{A}"><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="{naam}"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{E(x1)}" y="{E(y)}"/><a:ext cx="{E(x2 - x1)}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{kleur}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>')

def plaatje(part, rechts, boven, hoogte, naam):
    _, rid = part.get_or_add_image_part(logo)
    w = hoogte * 236 / 83
    return etree.fromstring(
        f'<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvPicPr><p:cNvPr id="{nid()}" name="{naam}" descr="Heijmans"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr userDrawn="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{E(rechts - w)}" y="{E(boven)}"/><a:ext cx="{E(w)}" cy="{E(hoogte)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')

def zet(ph_el, x, y, w, h):
    sppr = ph_el.find(qn('p:spPr'))
    for old in sppr.findall(qn('a:xfrm')): sppr.remove(old)
    xf = etree.SubElement(sppr, qn('a:xfrm')); sppr.insert(0, xf)
    etree.SubElement(xf, qn('a:off'), x=str(E(x)), y=str(E(y))); etree.SubElement(xf, qn('a:ext'), cx=str(E(w)), cy=str(E(h)))

def tekststijl(ph_el, sz, kleur, vet=False, algn='l', prompt=None, anchor=None):
    tb = ph_el.find(qn('p:txBody'))
    if tb is None:
        tb = etree.SubElement(ph_el, qn('p:txBody')); etree.SubElement(tb, qn('a:bodyPr')); etree.SubElement(tb, qn('a:lstStyle')); etree.SubElement(tb, qn('a:p'))
    bp = tb.find(qn('a:bodyPr'))
    for k in ('lIns', 'tIns', 'rIns', 'bIns'): bp.set(k, '0')
    if anchor: bp.set('anchor', anchor)
    ls = tb.find(qn('a:lstStyle'))
    for c in list(ls): ls.remove(c)
    lv = etree.SubElement(ls, qn('a:lvl1pPr'), algn=algn)
    etree.SubElement(lv, qn('a:buNone'))
    rp = etree.SubElement(lv, qn('a:defRPr'), sz=str(int(sz * 100)), b='1' if vet else '0')
    sf = etree.SubElement(rp, qn('a:solidFill')); etree.SubElement(sf, qn('a:srgbClr'), val=kleur)
    etree.SubElement(rp, qn('a:latin'), typeface='Arial')
    if prompt is not None:
        ph_el.find('.//' + qn('p:ph')).set('hasCustomPrompt', '1')
        for p_ in tb.findall(qn('a:p')): tb.remove(p_)
        p_ = etree.SubElement(tb, qn('a:p')); r_ = etree.SubElement(p_, qn('a:r'))
        etree.SubElement(r_, qn('a:rPr'), lang='nl-NL'); etree.SubElement(r_, qn('a:t')).text = prompt

def ph_of(tree, typ=None, idx=None):
    for sp in tree.iter(qn('p:sp')):
        ph = sp.find('.//' + qn('p:ph'))
        if ph is None: continue
        if typ is not None and ph.get('type', 'body') == typ: return sp
        if idx is not None and ph.get('idx') == str(idx): return sp
    return None

# ---------- thema: Heijmans-kleuren en Arial
master = prs.slide_master
thema = master.part.part_related_by('http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme')
th = etree.fromstring(thema.blob)
kleuren = dict(dk1=NAVY, lt1='FFFFFF', dk2=NAVY, lt2='F3F5F9', accent1=GEEL, accent2=NAVY, accent3='A4B3BB', accent4='E30613',
               accent5='3B7A1E', accent6='5B9BD5', hlink='1F4E9A', folHlink=GRIJS)
cs = th.find('.//' + qn('a:clrScheme')); cs.set('name', 'Heijmans RDFO')
for nm, v in kleuren.items():
    el = cs.find(qn('a:' + nm))
    for c in list(el): el.remove(c)
    etree.SubElement(el, qn('a:srgbClr'), val=v)
fs = th.find('.//' + qn('a:fontScheme')); fs.set('name', 'Arial')
for k in ('a:majorFont', 'a:minorFont'): fs.find(qn(k)).find(qn('a:latin')).set('typeface', 'Arial')
thema._blob = etree.tostring(th, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------- model (master)
mt = master._element; tree = mt.find('.//' + qn('p:spTree'))
titel = ph_of(tree, 'title'); zet(titel, 52, 28, 860, 34)
body = ph_of(tree, 'body'); zet(body, 52, 110, 872, 380)
for typ in ('dt',):
    el = ph_of(tree, typ)
    if el is not None: tree.remove(el)
voet = ph_of(tree, 'ftr'); zet(voet, 36, 516, 700, 16); tekststijl(voet, 9, GRIJS)
nr = ph_of(tree, 'sldNum'); tree.remove(nr)          # paginanummer als vast veld (altijd zichtbaar)
paginanr = etree.fromstring(
    f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{nid()}" name="Paginanummer"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
    f'<p:spPr><a:xfrm><a:off x="{E(800)}" y="{E(516)}"/><a:ext cx="{E(40)}" cy="{E(16)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
    '<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/><a:p><a:pPr algn="r"/>'
    f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="nl-NL" sz="900" b="1"><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>')
ts = mt.find(qn('p:txStyles'))
tl = ts.find(qn('p:titleStyle')).find(qn('a:lvl1pPr')); tl.set('algn', 'l')
dr = tl.find(qn('a:defRPr')); dr.set('sz', '2400'); dr.set('b', '1')
for lvl, sz in (('a:lvl1pPr', 1400), ('a:lvl2pPr', 1200), ('a:lvl3pPr', 1100), ('a:lvl4pPr', 1000), ('a:lvl5pPr', 1000)):
    el = ts.find(qn('p:bodyStyle')).find(qn(lvl))
    if el is not None: el.find(qn('a:defRPr')).set('sz', str(sz))
for el in (rect(0, 0, 960, 6, NAVY, 'Navy lijn'),
           lijn(36, 510, 924, LIJN, 'Voetlijn'), plaatje(master.part, 924, 513, 18, 'Logo'), paginanr):
    tree.insert(2, el)

# ---------- indelingen
houden = {'Title Slide': 'Titeldia', 'Title and Content': 'Titel en inhoud', 'Two Content': 'Twee kolommen',
          'Title Only': 'Alleen titel', 'Blank': 'Leeg'}
for lay in list(prs.slide_layouts):
    if lay.name not in houden: prs.slide_layouts.remove(lay)
for lay in prs.slide_layouts:
    nm = houden[lay.name]; lt = lay._element; cs_ = lt.find(qn('p:cSld')); cs_.set('name', nm); t = cs_.find(qn('p:spTree'))
    for typ in ('dt', 'sldNum'):
        el = ph_of(t, typ)
        if el is not None: t.remove(el)
    if nm == 'Titeldia':
        lt.set('showMasterSp', '0')
        bg = etree.fromstring(f'<p:bg xmlns:p="{P}" xmlns:a="{A}"><p:bgPr><a:solidFill><a:srgbClr val="{GEEL}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>')
        cs_.insert(0, bg)
        for el in (rect(0, 0, 960, 10, NAVY, 'Navy lijn'), rect(74, 322, 120, 8, NAVY, 'Navy balk'), plaatje(lay.part, 900, 456, 44, 'Logo')):
            t.insert(2, el)
        ct = ph_of(t, 'ctrTitle'); zet(ct, 72, 150, 820, 150); tekststijl(ct, 44, NAVY, True, 'l', 'RDFO-prognoserapportage', 'b')
        st = ph_of(t, 'subTitle'); zet(st, 72, 346, 820, 34); tekststijl(st, 20, NAVY, False, 'l', 'Q3 2026 · regio')
        for typ in ('ftr',):
            el = ph_of(t, typ)
            if el is not None: t.remove(el)
        # kleine regel boven de titel (regio)
        reg = copy.deepcopy(st); reg.find('.//' + qn('p:cNvPr')).set('id', str(nid())); reg.find('.//' + qn('p:cNvPr')).set('name', 'Regio')
        ph = reg.find('.//' + qn('p:ph')); ph.attrib.pop('type', None); ph.set('type', 'body'); ph.set('idx', '13'); ph.set('sz', 'quarter')
        zet(reg, 72, 64, 600, 30); tekststijl(reg, 16, NAVY, True, 'l', 'Heijmans Vastgoed · regio')
        t.append(reg)
        continue
    tp = ph_of(t, 'title')
    if tp is not None: tekststijl(tp, 24, NAVY, True, 'l', 'Titel van de dia')
    if nm in ('Titel en inhoud', 'Alleen titel', 'Twee kolommen'):
        t.insert(2, rect(36, 30, 6, 30, GEEL, 'Gele balk'))
        # kernboodschap onder de titel
        kern = etree.fromstring(
            f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{nid()}" name="Kernboodschap"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" sz="quarter" idx="14"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>')
        zet(kern, 52, 64, 860, 22); tekststijl(kern, 12, GRIJS, False, 'l', 'Kernboodschap in één zin')
        t.append(kern)
    if nm == 'Titel en inhoud':
        b = ph_of(t, idx=1)
        if b is not None: zet(b, 52, 110, 872, 380)
    if nm == 'Twee kolommen':
        b1, b2 = ph_of(t, idx=1), ph_of(t, idx=2)
        for b_, x_ in ((b1, 52), (b2, 496)):
            if b_ is None: continue
            zet(b_, x_, 110, 428, 380)
            ls = b_.find(qn('p:txBody')).find(qn('a:lstStyle'))
            for c in list(ls): ls.remove(c)          # zelfde tekstgrootte als het model

# ---------- voorbeelddia's (leeg, met invulteksten)
lay = {l.name: l for l in prs.slide_layouts}
prs.slides.add_slide(lay['Titeldia'])
s = prs.slides.add_slide(lay['Titel en inhoud'])
for ph in s.placeholders:
    if ph.placeholder_format.type is not None and ph.placeholder_format.idx in (11, 12): pass
prs.core_properties.title = 'RDFO-template Heijmans'; prs.core_properties.author = ''
prs.save(out); print('opgeslagen:', out, 'indelingen:', [l.name for l in prs.slide_layouts])
