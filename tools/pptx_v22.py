"""Sjabloon v22 = v21, strakker (zelfde opbouw, 10 dia's + 1):

- palet: de kleuren van het werkboek (navy 003366, geel FEB900, groen 15803D, rood B91C1C) in dia's, grafieken en indelingen;
  rustige, onderscheidbare scenariokleuren;
- leesbaarheid: geen tekst onder 8 pt op de dia's en in de grafieken (behalve het dianummer en de tijdas van de
  bouwtermijnengrafiek, die de macro zelf schaalt);
- dia 3: het vierde cijferblok wordt de kans grondtermijn (CF_KANS_LABEL / CF_KANS_WAARDE);
- nieuwe dia 4 'Vorige ↔ huidig ↔ realisatie' (VR_*) met tabel VR_TABEL;
- dia Update PEW: PEW_TABEL is een echte tabel (geen geplakte afbeelding meer), met PEW_TOP3 en PEW_SUBTITEL.

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


def maak_tabel(bron, slide, naam, x, y, kolommen, rijen, rijhoogte=0.72, kop_sz=800, sz=900, vet=None, streep=None):
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
            zet_tekst(tc.find(A + 'txBody'), tekst)
            if ri == 0:
                run_kleur(tc, GRIJS if ci == 0 else NAVY, kop_sz, True, 'l' if ci == 0 else 'r')
            else:
                dik = ri in vet if vet is not None else ci > 0       # vet: alleen de totaalregels, anders alle bedragen
                run_kleur(tc, NAVY if (vet is not None and ri in vet) or ci else GRIJS, sz, dik, 'l' if ci == 0 else 'r')
                if streep == ri:                                     # dikkere lijn boven deze regel (Resultaat)
                    ln = tc.find(A + 'tcPr').find(A + 'lnT')
                    ln.set('w', '15875')
                    for c in ln.iter(A + 'srgbClr'):
                        c.set('val', NAVY)
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
    vorm(s, 'PEW_SUBTITEL').text_frame.text = 'Resultaat €3,5 mln (4,4%) · gelijk aan vorig kwartaal'
    kop_als_sc(p, s, 'PEW_KICKER', 'PEW_TITEL', 'PEW_SUBTITEL')
    kaart = vorm(s, 'PEW_TABEL')
    kaart.name = 'PEW_KAART'
    kaart.left, kaart.width = cm(1.1), cm(16.6)
    for i, (old, new) in enumerate((('Shape 6', 'PEW_K1_ACCENT'), ('Shape 11', 'PEW_K2_ACCENT'), ('Shape 16', 'PEW_K3_ACCENT'))):
        vorm(s, old).name = new
    rijen = [['x € 1.000', 'Vorig', 'Huidig', 'Δ'],
             ['Grondkosten', '8.390', '8.390', '0'],
             ['Aanneemsom', '53.341', '53.341', '0'],
             ['Overige kosten', '14.123', '14.123', '0'],
             ['Totaal kosten', '75.854', '75.854', '0'],
             ['Totaal opbrengsten', '79.336', '79.336', '0'],
             ['Resultaat', '3.482', '3.482', '0'],
             ['Resultaat %', '4,4%', '4,4%', '0,0%']]
    maak_tabel(sc_tabel, s, 'PEW_TABEL', 1.6, 3.6, [6.2, 3.1, 3.1, 3.1], rijen, rijhoogte=0.8, vet={4, 5, 6}, streep=6)
    tekstvak(s, 'PEW_TOP3', 1.6, 10.3, 15.5, 2.0,
             'Top 3 kosten Q3 2026 (actuals)\n1. Aanneemsom: € 1.234.567\n2. Honoraria: € 234.567\n3. Grondkosten: € 34.567',
             sz=9, kleur='111F39')
    return s


def dia_vr(p, pew, sc_tabel):
    """nieuwe dia 'Vorige ↔ huidig ↔ realisatie' (kopie van de opbouw van Update PEW), direct na de cashflow (dia 3)"""
    nieuw = p.slides.add_slide(pew.slide_layout)
    boom = nieuw.shapes._spTree
    for e in list(boom)[2:]:
        boom.remove(e)
    for e in list(pew.shapes._spTree)[2:]:
        boom.append(copy.deepcopy(e))
    hernoem = {'PEW_KICKER': 'VR_KICKER', 'PEW_TITEL': 'VR_TITEL', 'PEW_SUBTITEL': 'VR_SUBTITEL', 'PEW_KAART': 'VR_KAART',
               'PEW_TABEL': 'VR_TABEL', 'PEW_TOP3': 'VR_TOELICHTING', 'PEW_DATUM': 'VR_DATUM',
               'Text 7': 'VR_K1_LABEL', 'PEW_FD_WAARDE': 'VR_K1_WAARDE', 'PEW_FD_SUB': 'VR_K1_SUB',
               'Text 12': 'VR_K2_LABEL', 'PEW_BASIS_WAARDE': 'VR_K2_WAARDE', 'PEW_BASIS_SUB': 'VR_K2_SUB',
               'Text 17': 'VR_K3_LABEL', 'PEW_RES_WAARDE': 'VR_K3_WAARDE', 'PEW_RES_SUB': 'VR_K3_SUB',
               'PEW_K1_ACCENT': 'VR_K1_ACCENT', 'PEW_K2_ACCENT': 'VR_K2_ACCENT', 'PEW_K3_ACCENT': 'VR_K3_ACCENT'}
    for shp in nieuw.shapes:
        if shp.name in hernoem:
            shp.name = hernoem[shp.name]
    teksten = {'VR_KICKER': 'VORIGE PROGNOSE ↔ HUIDIG ↔ REALISATIE',
               'VR_TITEL': 'Eindsaldo −€0,1 mln t.o.v. vorige prognose · Q3 ’26 sluit aan op SAP',
               'VR_SUBTITEL': 'Vorige prognose = de cashflow van vorig kwartaal · gerealiseerd = SAP-actuals van het laatst afgesloten kwartaal',
               'VR_K1_LABEL': 'EINDSALDO', 'VR_K1_WAARDE': '€3,4M', 'VR_K1_SUB': '−€0,1M t.o.v. vorige prognose',
               'VR_K2_LABEL': 'MAX. FINANCIERING', 'VR_K2_WAARDE': '−€10,9M', 'VR_K2_SUB': '€1,3M minder diep dan vorige prognose',
               'VR_K3_LABEL': 'REALISATIE Q3 ’26', 'VR_K3_WAARDE': 'sluit aan', 'VR_K3_SUB': 'afgesloten periodes = SAP-actuals',
               'VR_TOELICHTING': 'Details per kwartaal: tab CF vorig en huidig in het werkbestand.'}
    for shp in nieuw.shapes:
        if shp.name in teksten and shp.has_text_frame:
            zet_tekst(shp.text_frame._txBody, teksten[shp.name])
    # tabel: andere inhoud en kolommen (zelfde opmaak)
    oud = vorm(nieuw, 'VR_TABEL')
    boom.remove(oud._element)
    rijen = [['', 'Vorige prognose', 'Huidig / gerealiseerd', 'Verschil'],
             ['Eindsaldo cashflow', '€3,5M', '€3,4M', '−€0,1M'],
             ['Laagste stand (max. financiering)', '−€12,2M', '−€10,9M', '€1,3M'],
             ['Kosten Q3 ’26', '€0,1M', '€0,1M', '€0,0M'],
             ['Omzet Q3 ’26', '€0,0M', '€0,0M', '€0,0M'],
             ['Afgesloten periodes = SAP', '', 'OK', '']]
    maak_tabel(sc_tabel, nieuw, 'VR_TABEL', 1.6, 3.6, [6.7, 2.9, 3.3, 2.6], rijen, rijhoogte=1.0)
    toel = vorm(nieuw, 'VR_TOELICHTING')
    toel.top, toel.height = cm(10.1), cm(0.8)
    kaart = vorm(nieuw, 'VR_KAART')                  # de kaart eindigt gelijk met de drie kaarten rechts
    kaart.height = vorm(nieuw, 'Shape 15').top + vorm(nieuw, 'Shape 15').height - kaart.top
    for r in toel._element.iter(A + 'rPr'):
        r.set('b', '0')
    # positie: direct na de cashflow (dia 3)
    lst = p.slides._sldIdLst
    el = lst[-1]
    lst.remove(el)
    lst.insert(3, el)
    return nieuw


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
    pew = dia_pew(p, sc_tabel)
    dia_vr(p, pew, sc_tabel)
    palet_en_letters(p)
    afwerking(p)
    return p
