"""Sjabloon v22 = v21, strakker (zelfde opbouw, 10 dia's + 1):

- palet: de kleuren van het werkboek (navy 003366, geel FEB900, groen 15803D, rood B91C1C) in dia's, grafieken en indelingen;
  rustige, onderscheidbare scenariokleuren;
- leesbaarheid: geen tekst onder 8 pt op de dia's en in de grafieken (behalve het dianummer en de tijdas van de
  bouwtermijnengrafiek, die de macro zelf schaalt);
- dia 3: het vierde cijferblok wordt de kans grondtermijn (CF_KANS_LABEL / CF_KANS_WAARDE);
- de dia 'Van verkoop naar omzet' vervalt (9 dia's);
- dia Update PEW: een leeg vak PEW_BEELD over de hele breedte; de macro plakt er het overzicht van 2. Ontw PEW (bij DAEB
  2B. Ontw PEW) als afbeelding in; FD en basis in PEW_SUBTITEL (geen top 3 kosten meer op de dia);
- verkoop en transport: de cumulatieve %-lijn, labels en rechter as rood (B91C1C), as tot 115%.

Gebruik: maak(presentation) na het hernoemen naar v21 (build_werkbestand.maak_sjabloon)."""
import copy
import re

from lxml import etree
from pptx.util import Emu, Pt

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
A = '{%s}' % NS['a']

PALET = {'002060': '003366', 'F2B800': 'FEB900', 'EDA100': 'FEB900', '12864E': '15803D', '2F9E63': '15803D',
         '5A8F3D': '15803D', 'C0392F': 'B91C1C', 'F08A00': '003366',
         '7B2FBE': '2563EB', 'E0218A': 'B45309', '00897B': '0F766E',
         # woningtypes 1-10 (type 4 is het geel hierboven): rustige reeks die bij het palet past
         '2A78D6': '2563EB', 'EB6834': 'B45309', '1BAF7A': '0F766E', 'E87BA4': '64748B', '008300': '15803D',
         '4A3AA7': '1E3A8A', '8A9A1B': 'D97706', '1F8A8A': '0891B2', '8C6D3F': '92400E',
         '00B0F0': '003366', 'C2185B': '003366'}        # cumulatieve lijn in de verkoop- en transportgrafiek
TEKST_GEEL = 'B45309'                       # gele tekst (EDA100/F2B800) in Excel: donker amber, geel op wit is onleesbaar
MIN_SZ = 800                                    # 8 pt
NAVY, GRIJS, LIJN = '003366', '4B5563', 'E4E5E6'


def cm(x):
    return Emu(int(round(x * 360000)))


# ---------------------------------------------------------------------------------------------------------------
ROOD = 'A16207'                                    # donkergeel: de percentages bij de cumulatieve lijn (naam historisch)
VORIG_KL, SCEN_KL = 'FEB900', 'DC2626'              # cashflowgrafiek: vorige prognose geel, scenario rood
OUD_TYPEN = ['2563EB', 'B45309', '0F766E', 'FEB900', '64748B', '15803D', '1E3A8A', 'D97706', '0891B2', '92400E']
TYPEN = ['2563EB', '0D9488', '7C3AED', '64748B', '0EA5E9', '16A34A', 'DB2777', '4F46E5', '65A30D', '57534E']
C = '{%s}' % NS['c']


def pct_rood(cs):
    """grafiek verkoop / transport: de cumulatieve %-lijn navy, de percentages (labels en rechter as) donkerrood, zodat ze
    afsteken tegen de staven; rechter as tot 115% zodat het label '100%' niet tegen de titel valt.
    Geeft het aantal rood gemaakte reeksen terug."""
    n = 0
    for ser in cs.iter(C + 'ser'):
        tx = ser.find(C + 'tx')
        naam = ''.join(tx.itertext()) if tx is not None else ''
        if 'cumulatief %' in naam.lower():
            dl = ser.find(C + 'dLbls')
            in_lbl = set(dl.iter(A + 'srgbClr')) if dl is not None else set()
            for c in ser.iter(A + 'srgbClr'):
                c.set('val', ROOD if c in in_lbl else NAVY)        # lijn en markers navy, labels donkergeel
            n += 1
    for ax in cs.iter(C + 'valAx'):
        pos = ax.find(C + 'axPos')
        if pos is None or pos.get('val') != 'r':
            continue
        sc = ax.find(C + 'scaling')
        mx = sc.find(C + 'max')
        if mx is not None and mx.get('val') == '1':
            mx.set('val', '1.15')
            mu = ax.find(C + 'majorUnit')
            if mu is None:
                mu = etree.Element(C + 'majorUnit')
                na = ax.find(C + 'minorUnit')
                if na is None:
                    na = ax.find(C + 'dispUnits')
                if na is None:
                    na = ax.find(C + 'extLst')
                if na is not None:
                    na.addprevious(mu)
                else:
                    ax.append(mu)
            mu.set('val', '0.25')
        tp = ax.find(C + 'txPr')
        if tp is not None:
            for c in tp.iter(A + 'srgbClr'):
                c.set('val', ROOD)
    return n


def _naam(ser):
    tx = ser.find(C + 'tx')
    if tx is None:
        return ''
    v = tx.find('.//' + C + 'v')                    # de getoonde naam (bij een celverwijzing de cache, niet de formule)
    return (v.text if v is not None and v.text else ''.join(tx.itertext())).strip()


def _lijn(ser, kleur, w, dash):
    sp = ser.find(C + 'spPr')
    ln = sp.find(A + 'ln') if sp is not None else None
    if ln is None:
        return False
    ln.set('w', str(w))
    for c in ln.iter(A + 'srgbClr'):
        c.set('val', kleur)
    d = ln.find(A + 'prstDash')
    if d is None:
        d = etree.SubElement(ln, A + 'prstDash')
        nxt = [e for e in ln if e.tag in (A + 'round', A + 'bevel', A + 'miter', A + 'headEnd', A + 'tailEnd')]
        if nxt:
            nxt[0].addprevious(d)
    d.set('val', dash)
    return True


def lijnkleuren(cs):
    """cashflowgrafiek: meer contrast tussen de lijnen. Huidig blijft navy; vorige prognose geel (gestreept, 2,25 pt),
    scenario rood (lange streep, 2 pt); het label 'Scenario €..' rood. Geeft het aantal aangepaste reeksen terug."""
    n = 0
    for ser in cs.iter(C + 'ser'):
        nm = _naam(ser).lower()
        if nm == 'vorige prognose':
            n += _lijn(ser, VORIG_KL, 28575, 'dash')
        elif nm == 'scenario':
            n += _lijn(ser, SCEN_KL, 25400, 'lgDash')
        elif nm.startswith('scenario ') and '€' in nm:
            for c in ser.iter(A + 'srgbClr'):
                if c.get('val') == 'FEB900':
                    c.set('val', SCEN_KL)
            n += 1
    return n


def typekleuren(cs):
    """grafieken per woningtype: de typen in een eigen reeks kleuren (geen geel: dat is vorige prognose en de percentages;
    geen rood: kosten en scenario; geen navy: de cumulatieve lijn). Herkent de typereeksen aan de oude kleuren op volgorde
    (10 reeksen achter elkaar). Geeft het aantal aangepaste reeksen terug."""
    sers = list(cs.iter(C + 'ser'))
    eerst = []
    for ser in sers:
        sp = ser.find(C + 'spPr')
        c = sp.find('.//' + A + 'srgbClr') if sp is not None else None
        eerst.append(c.get('val') if c is not None else None)
    n = 0
    for j in range(len(sers) - 9):
        if eerst[j:j + 10] == OUD_TYPEN:
            for i in range(10):
                for c in sers[j + i].iter(A + 'srgbClr'):
                    if c.get('val') == OUD_TYPEN[i]:
                        c.set('val', TYPEN[i])
                    elif c.get('val') == '1F2937' and i == 3:          # tekst op het oude geel: nu wit
                        c.set('val', 'FFFFFF')
                n += 1
            eerst[j:j + 10] = TYPEN
    return n


def typekleur_tekst(f):
    """legendakleuren als tekst ('2563EB,' ... per type in een formule): op volgorde naar TYPEN"""
    import re
    i = [0]

    def r(m):
        k = i[0]
        i[0] += 1
        return '"%s,"' % (TYPEN[k] if k < 10 and m.group(1) == OUD_TYPEN[k] else m.group(1))
    return re.sub(r'"([0-9A-F]{6}),"', r, f)


def daeb_lijn(cs):
    """verkoop / transport: de lijn van de DAEB-ruit naar beneden (foutbalk) goed zichtbaar: donkergoud, 1,5 pt, gestreept
    (was dun lichtgeel, nauwelijks te zien). Geeft het aantal aangepaste lijnen terug."""
    n = 0
    for eb in cs.iter(C + 'errBars'):
        for ln in eb.iter(A + 'ln'):
            ln.set('w', '19050')
            for c in ln.iter(A + 'srgbClr'):
                c.set('val', '7A5300')
            d = ln.find(A + 'prstDash')
            if d is not None:
                d.set('val', 'dash')
            n += 1
    return n


def palet(el):
    for c in el.iter(A + 'srgbClr'):
        v = (c.get('val') or '').upper()
        if v in PALET:
            c.set('val', PALET[v])


def min_letter(el, minimum=MIN_SZ):
    for tag in ('rPr', 'defRPr', 'endParaRPr'):
        for r in el.iter(A + tag):
            sz = r.get('sz')
            if sz and int(sz) < minimum:
                r.set('sz', str(minimum))


def vorm(slide, naam):
    return next(s for s in slide.shapes if s.name == naam)


def nieuw_id(slide):
    return max([int(e.get('id')) for e in slide.shapes._spTree.iter('{%s}cNvPr' % NS['p'])] + [1]) + 1


def zet_tekst(cel_txbody, tekst):
    """eerste run van de eerste alinea houdt zijn opmaak; de rest weg"""
    p = cel_txbody.find(A + 'p')
    runs = p.findall(A + 'r')
    for r in runs[1:]:
        p.remove(r)
    if runs:
        runs[0].find(A + 't').text = tekst
    for extra in cel_txbody.findall(A + 'p')[1:]:
        cel_txbody.remove(extra)


def run_kleur(tc, kleur=None, sz=None, b=None, algn=None):
    for r in tc.iter(A + 'rPr'):
        if sz:
            r.set('sz', str(sz))
        if b is not None:
            r.set('b', '1' if b else '0')
        if kleur:
            for c in r.iter(A + 'srgbClr'):
                c.set('val', kleur)
    if algn:
        p = tc.find('.//' + A + 'p')
        ppr = p.find(A + 'pPr')
        if ppr is None:
            ppr = etree.SubElement(p, A + 'pPr')
            p.remove(ppr)
            p.insert(0, ppr)
        ppr.set('algn', algn)


def maak_tabel(bron, slide, naam, x, y, kolommen, rijen, rijhoogte=0.72, kop_sz=800, sz=900, vet=None, streep=None,
               kopjes=(), tekstkol=(), marge=None):
    """tabel in de opmaak van SC_TABEL (dia 5): kopregel navy vet, eerste kolom grijs, bedragen rechts"""
    gf = copy.deepcopy(bron._element)
    nv = gf.find('.//{%s}cNvPr' % NS['p'])
    nv.set('id', str(nieuw_id(slide)))
    nv.set('name', naam)
    tbl = gf.find('.//' + A + 'tbl')
    grid = tbl.find(A + 'tblGrid')
    cols = grid.findall(A + 'gridCol')
    for g in cols[1:]:
        grid.remove(g)
    g0 = cols[0]
    for e in list(g0):
        g0.remove(e)                      # geen colId-extensies (dubbele ids)
    for i, w in enumerate(kolommen):
        g = g0 if i == 0 else copy.deepcopy(g0)
        g.set('w', str(int(cm(w))))
        if i:
            grid.append(g)
    trs = tbl.findall(A + 'tr')
    kop_tr, body_tr = trs[0], trs[1]
    for tr in trs:
        tbl.remove(tr)
    h = int(cm(rijhoogte))
    for ri, rij in enumerate(rijen):
        tr = copy.deepcopy(kop_tr if ri == 0 else body_tr)
        tr.set('h', str(h))
        tcs = tr.findall(A + 'tc')
        while len(tcs) < len(kolommen):
            tr.append(copy.deepcopy(tcs[-1]))
            tcs = tr.findall(A + 'tc')
        for tc in tcs[len(kolommen):]:
            tr.remove(tc)
        tcs = tr.findall(A + 'tc')
        for ci, tekst in enumerate(rij):
            tc = tcs[ci]
            bron_tc = (kop_tr if ri == 0 else body_tr).findall(A + 'tc')[min(1, ci) if ri else 1]
            body = tc.find(A + 'txBody')
            tc.replace(body, copy.deepcopy(bron_tc.find(A + 'txBody')))
            zet_tekst(tc.find(A + 'txBody'), tekst or '\u00a0')      # lege cel: harde spatie (houdt de rijhoogte)
            if ri == 0:
                run_kleur(tc, GRIJS if ci == 0 else NAVY, kop_sz, True, 'l' if ci == 0 or ci in tekstkol else 'r')
            else:
                dik = ri in vet if vet is not None else ci > 0       # vet: alleen de totaalregels, anders alle bedragen
                if ri in kopjes:                                     # kopje (KOSTEN / OPBRENGSTEN): klein, vet, navy
                    run_kleur(tc, NAVY, kop_sz, True, 'l')
                elif ci in tekstkol:                                 # tekstkolom (toelichting): links, grijs, niet vet
                    run_kleur(tc, GRIJS, sz, False, 'l')
                else:
                    run_kleur(tc, NAVY if (vet is not None and ri in vet) or ci else GRIJS, sz, dik, 'l' if ci == 0 else 'r')
                if ri == streep or (isinstance(streep, (set, tuple)) and ri in streep):   # dikkere lijn boven deze regel
                    ln = tc.find(A + 'tcPr').find(A + 'lnT')
                    ln.set('w', '15875')
                    for c in ln.iter(A + 'srgbClr'):
                        c.set('val', NAVY)
        for par in tr.iter(A + 'p'):                                 # lege cellen: zelfde lettergrootte (anders rekt de rij op)
            e = par.find(A + 'endParaRPr')
            if e is None:
                e = etree.SubElement(par, A + 'endParaRPr')
            e.set('sz', str(kop_sz if ri == 0 else sz))
        if marge is not None:                                        # celmarge boven/onder (EMU), voor compacte tabellen
            for pr in tr.iter(A + 'tcPr'):
                pr.set('marT', str(marge))
                pr.set('marB', str(marge))
        tbl.append(tr)
    xfrm = gf.find('{%s}xfrm' % NS['p'])
    xfrm.find(A + 'off').set('x', str(int(cm(x))))
    xfrm.find(A + 'off').set('y', str(int(cm(y))))
    xfrm.find(A + 'ext').set('cx', str(int(sum(cm(w) for w in kolommen))))
    xfrm.find(A + 'ext').set('cy', str(h * len(rijen)))
    slide.shapes._spTree.append(gf)
    return gf


def tekstvak(slide, naam, x, y, w, h, tekst, sz=8, kleur=GRIJS, vet=False):
    tb = slide.shapes.add_textbox(cm(x), cm(y), cm(w), cm(h))
    tb.name = naam
    tf = tb.text_frame
    tf.word_wrap = True
    for m in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'):
        setattr(tf, m, 0)
    tf.text = tekst
    from pptx.dml.color import RGBColor
    for i, par in enumerate(tf.paragraphs):
        for r in par.runs:
            r.font.size = Pt(sz)
            r.font.bold = vet or i == 0
            r.font.color.rgb = RGBColor.from_string(kleur)
    return tb


# ---------------------------------------------------------------------------------------------------------------
def dia3_kans(p):
    s = p.slides[2]
    vorm(s, 'CF_BE_LABEL').name = 'CF_KANS_LABEL'
    vorm(s, 'CF_BE_WAARDE').name = 'CF_KANS_WAARDE'
    zet_tekst(vorm(s, 'CF_KANS_LABEL').text_frame._txBody, 'KANS GRONDTERMIJN')
    zet_tekst(vorm(s, 'CF_KANS_WAARDE').text_frame._txBody, '+€1,2M eerder')


def dia2_labels(p):
    """kaart verkocht / getransporteerd: labels breed genoeg voor 8 pt, de balken iets korter"""
    s = p.slides[1]
    zet_tekst(vorm(s, 'KPI7_LABEL').text_frame._txBody, 'VERKOCHT / TRANSPORT')
    for x in 'VT':
        lab, spoor, balk = (vorm(s, 'KPI7_%s_%s' % (n, x)) for n in ('LABEL', 'SPOOR', 'BALK'))
        aandeel = balk.width / spoor.width if spoor.width else 0
        lab.width = cm(2.5)
        spoor.left, spoor.width = cm(10.1), cm(2.0)
        balk.left, balk.width = spoor.left, int(spoor.width * aandeel)


def kop_als_sc(p, s, kicker, titel, sub):
    """kopje, titel en ondertitel in dezelfde opmaak en positie als dia 5 (Scenario's)"""
    bron = p.slides[4]
    for oud, nieuw_naam, tekst in ((kicker, 'KICKER', None), (titel, 'TITEL', None), (sub, 'SUBTITEL', None)):
        o = vorm(s, oud)
        txt = o.text_frame.text
        e = copy.deepcopy(vorm(bron, 'SC_' + nieuw_naam)._element)
        o._element.addprevious(e)
        s.shapes._spTree.remove(o._element)
        shp = vorm(s, 'SC_' + nieuw_naam)
        shp.name = oud
        shp._element.find('.//{%s}cNvPr' % NS['p']).set('id', str(nieuw_id(s)))
        zet_tekst(shp.text_frame._txBody, txt or ' ')


def dia_pew(p, sc_tabel):
    """dia 'Update PEW' (laatste): kaart met een echte tabel en de top 3 kosten eronder"""
    s = p.slides[len(p.slides) - 1]
    vorm(s, 'Text 0').name = 'PEW_KICKER'
    vorm(s, 'Text 1').name = 'PEW_TITEL'
    vorm(s, 'Text 2').name = 'PEW_SUBTITEL'
    vorm(s, 'PEW_SUBTITEL').text_frame.text = 'Resultaat €5,7 mln (16,8%) · gelijk aan vorig kwartaal · FD: SKO · basis: FD DO'
    kop_als_sc(p, s, 'PEW_KICKER', 'PEW_TITEL', 'PEW_SUBTITEL')
    kaart = vorm(s, 'PEW_TABEL')
    kaart.name = 'PEW_KAART'
    kaart.left, kaart.width, kaart.top, kaart.height = cm(1.1), cm(23.1), cm(3.0), cm(9.75)
    # de drie kaarten rechts vervallen: FD en PEW-basis staan in de ondertitel
    for naam in ('Shape 5', 'Shape 6', 'Text 7', 'PEW_FD_WAARDE', 'PEW_FD_SUB', 'Shape 10', 'Shape 11', 'Text 12',
                 'PEW_BASIS_WAARDE', 'PEW_BASIS_SUB', 'Shape 15', 'Shape 16', 'Text 17', 'PEW_RES_WAARDE', 'PEW_RES_SUB'):
        s.shapes._spTree.remove(vorm(s, naam)._element)
    vorm(s, 'PEW_DATUM').top = cm(12.85)
    # leeg vak voor de afbeelding van 2. Ontw PEW (of 2B bij DAEB); de macro plakt hem er passend in
    from pptx.enum.shapes import MSO_SHAPE
    vak = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, cm(1.5), cm(3.2), cm(22.3), cm(9.4))
    vak.name = 'PEW_BEELD'
    vak.fill.background()
    vak.line.fill.background()
    return s


def dia_weg(p, vormnaam):
    """de dia met deze vorm uit het sjabloon (dia en relatie)"""
    lst = p.slides._sldIdLst
    for i, sl in enumerate(p.slides):
        if any(shp.name == vormnaam for shp in sl.shapes):
            p.part.drop_rel(lst[i].rId)
            del lst[i]
            return
    raise ValueError('geen dia met ' + vormnaam)


def palet_en_letters(p):
    for i, s in enumerate(p.slides):
        palet(s._element)
        for shp in s.shapes:
            if shp.name.startswith('Slide Number'):
                continue
            min_letter(shp._element)
            if shp.has_chart:
                cs = shp.chart._chartSpace
                palet(cs)
                if shp.name != 'BT_GRAFIEK':           # de tijdas van de bouwtermijnen schaalt de macro zelf
                    min_letter(cs)
                if shp.name in ('VP1_GRAFIEK', 'VP2_GRAFIEK'):
                    assert pct_rood(cs) == 2, shp.name
                    assert daeb_lijn(cs) >= 1, shp.name
                lijnkleuren(cs)
                typekleuren(cs)
    for m in p.slide_masters:
        palet(m._element)
        for lay in m.slide_layouts:
            palet(lay._element)


def afwerking(p):
    """voorbeeldteksten zonder projectnaam, het dianummer leesbaar (8 pt grijs), korte ondertitel op de dia per type"""
    zet_tekst(vorm(p.slides[0], 'COVER_PROJECT').text_frame._txBody, 'Project')
    pt = next(s for s in p.slides if any(sh.name == 'PT_SUBTITEL' for sh in s.shapes))
    zet_tekst(vorm(pt, 'PT_SUBTITEL').text_frame._txBody,
              't/m Q4 ’26 €0,0 mln → daarna €15,4 mln · grijs = gerealiseerd (volgens planning)')
    for lay in p.slide_layouts:
        if lay.name == 'ORION':
            lay.name = 'FO'
        for shp in lay.shapes:
            if shp.name == 'VOET_TEKST':
                zet_tekst(shp.text_frame._txBody, 'PROJECT  |  Q3 2026')
    for boom in [s.shapes for s in p.slides] + [lay.shapes for lay in p.slide_layouts if lay.name == 'FO']:
        for shp in boom:
            if not shp.name.startswith('Slide Number'):
                continue
            el = shp._element
            for r in list(el.iter(A + 'defRPr')) + list(el.iter(A + 'rPr')):
                r.set('sz', str(MIN_SZ))
                for f in r.findall(A + 'solidFill'):
                    r.remove(f)
                fill = etree.SubElement(r, A + 'solidFill')
                r.remove(fill)
                r.insert(0, fill)
                etree.SubElement(fill, A + 'srgbClr').set('val', '7A7A7A')
            rechts = shp.left + shp.width                   # breed genoeg voor twee cijfers, rechts uitgelijnd
            shp.width, shp.height = cm(0.9), cm(0.4)
            shp.left = rechts - shp.width


def maak(p):
    sc_tabel = vorm(p.slides[4], 'SC_TABEL')
    dia3_kans(p)
    dia2_labels(p)
    dia_pew(p, sc_tabel)
    palet_en_letters(p)
    afwerking(p)
    dia_weg(p, 'VO_GRAFIEK')                 # 'Van verkoop naar omzet' vervalt
    return p
