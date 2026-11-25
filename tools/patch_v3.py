"""Ronde 3: soort/deelgebiedfilter doorvoeren, Invoer met status, AK-factor als instelling, links naar berekeningen,
mijlpalen op Projectdetail, mutatie + toelichting op Projecten, logo's voor de PowerPoint, tabvolgorde en -kleuren.
Draai na patch_versies.py en patch_eenvoud.py. usage: patch_v3.py in.xlsm out.xlsm logo_navy.png logo_wit.png"""
import sys, os, re
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book, skeleton, q
from patch_ak import split, col2n, n2col
from patch_eenvoud import (N, clear, hide_rows, add_cf, add_dv_list, add_hyperlink, hyperlinks_keep, expand_shared,
                           get_or_make, merges_drop, add_dxf)
from lxml import etree

def ftexts(s, rows=None):
    for r, row in s.rows.items():
        if rows and r not in rows: continue
        for el in row.findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.text: yield el, f

E_PATTERNS = [
    (re.compile(r'INDEX\(Projectlijst!\$E\$8:\$E\$107,([^)]*)\)="Nee"'), r'INDEX(Projectlijst!$AN$8:$AN$107,\1)<>"Ja"'),
    (re.compile(r'Projectlijst!\$E(\$?\d+)="Nee"'), r'Projectlijst!$AN\1<>"Ja"'),
    (re.compile(r'Projectlijst!\$E(\$?\d+)<>"Nee"'), r'Projectlijst!$AN\1="Ja"'),
]
def e_to_an(bk):
    for nm in ['VERW_Project', 'AK-dekking', "Kansen en risico's", 'Invoer', 'Dias']:
        s = bk.sheet(nm); expand_shared(s, s.rows.keys())
        for el, f in ftexts(s):
            t = f.text
            for pat, rep in E_PATTERNS: t = pat.sub(rep, t)
            f.text = t

def filters(bk):
    # keuzelijst deelgebied in Lijsten (Q4:Q20) + Sel_Deelgebied op Dashboard N5
    L = bk.sheet('Lijsten')
    L.set('Q3', text='L_DeelgebiedKeuze'); L.set('Q4', text='Alle'); L.set('Q5', text='(geen)')
    for k in range(1, 16): L.set(f'Q{5+k}', is_str=True, formula=f'IF(INDEX(Gebieden,{k})="","",INDEX(Gebieden,{k})&"")')
    bk.set_name('L_DeelgebiedKeuze', 'Lijsten!$Q$4:$Q$20')
    d = bk.sheet('Dashboard')
    mc = d.root.find(N+'mergeCells')
    for m in mc:
        if m.get('ref') == 'K5:N5': m.set('ref', 'K5:M5')
    d.set('N4', text='Deelgebied', style=d.style('K4')); d.set('N5', text='Alle', style=d.style('K5'))
    add_dv_list(d, 'N5', 'L_DeelgebiedKeuze')
    bk.set_name('Sel_Deelgebied', 'Dashboard!$N$5')
    d.set('E4', text='Kwartaal (nu)', style=d.style('E4'))
    cv = bk.sheet('CALC_Versies')
    d.set('E6', is_str=True, style=d.style('E4'), formula='IF(Sel_Prev="","",IF(Sel_Versie="","","vergeleken met "&Sel_Prev))')
    # woningen van de huidige versie komen uit de upload (er is maar een Woningetalage)
    for r in range(5, 10):
        for col in ('E', 'Q', 'R', 'W', 'X'):
            f = cv.cell(f'{col}{r}').find(N+'f')
            if f is not None and f.text and 'P_UploadVersie' not in f.text:
                f.text = f'IF(AND($A{r}=P_UploadVersie,P_UploadVJ=P_VJ),{col}$10,{f.text})'
    cv = bk.sheet('CALC_Versies')
    cv.set('B21', is_str=True, style=cv.style('B21'), formula='IF(AND(Sel_Type="Alle",Sel_Project="Alle"),"geen filter",IF(Sel_Project<>"Alle",Sel_Project,Sel_Type))&IF(OR(Sel_Deelgebied="Alle",Sel_Deelgebied=""),""," · deelgebied "&Sel_Deelgebied)')
    # Reserves
    rs = bk.sheet('Reserves'); expand_shared(rs, rs.rows.keys())
    for el, f in ftexts(rs):
        t = f.text
        t = re.sub(r'(?<![A-Z])SUM\((Projectlijst!\$[A-Z]+\$8:\$[A-Z]+\$107)\)', r'SUMIFS(\1,PL_Telt,"Ja")', t)
        t = re.sub(r'SUMPRODUCT\((Projectlijst!\$[LMN]\$8:\$[LMN]\$107),', r'SUMPRODUCT((PL_Telt="Ja")*\1,', t)
        t = re.sub(r'IF\(AND\(Projectlijst!\$A(\d+)<>"",', r'IF(AND(Projectlijst!$AN\1="Ja",Projectlijst!$A\1<>"",', t)
        f.text = t
    o = bk.sheet('Outlook')
    for el, f in ftexts(o):
        f.text = re.sub(r'(?<![A-Z])SUM\((RS_Effect\w*)\)', r'SUMIFS(\1,PL_Telt,"Ja")', f.text)
    # Kansen en risico's: telt alleen binnen filter (projectnaam -> Projectlijst)
    k = bk.sheet("Kansen en risico's")
    for r in range(8, 48):
        k.set(f'H{r}', style=k.style(f'H{r}'), formula=(f'IF(OR(F{r}="",C{r}=""),"",IF(E{r}="Risico",-ABS(F{r}),ABS(F{r}))*IF(OR(Sel_Deelgebied="Alle",Sel_Deelgebied=""),1,'
                                                         f'IFERROR(--(INDEX(PL_Telt,MATCH($A{r},Projectlijst!$B$8:$B$107,0))="Ja"),IF(Sel_Deelgebied="(geen)",1,0))))'))

def invoer(bk):
    s = bk.sheet('Invoer'); st = dict(lbl=s.style('C9'), inp=s.style('D9'), txt=s.style('E9'), hdr=s.style('C7'), sec=s.style('A5'))
    s.set('B2', style=s.style('B2'), text='Instellingen van dit bestand. Gele cel = invullen. Kolom F zegt wat nog ontbreekt. Projectgegevens staan op de Projectlijst.')
    # AK-factor in blok 1 (rij 25); oude status vastleggen naar D590
    s.set('C25', text='AK-dekking: kosten delen door', style=st['lbl']); s.set('D25', formula=None, style=st['inp'])
    el = s.cell('D25')
    for ch in list(el): el.remove(ch)
    if 't' in el.attrib: del el.attrib['t']
    v = etree.SubElement(el, N+'v'); v.text = '1.06'
    s.set('E25', text='AK-dekking = kosten per kwartaal / deze factor x AK %. Standaard 1,06.', style=st['txt'])
    s.set('D590', text='Vastgesteld'); bk.set_name('P_StatusVastleg', 'Invoer!$D$590')
    bk.set_name('P_AKFactor', 'Invoer!$D$25')
    s.set('C589', formula='$D$25', style=s.style('C589')); s.set('A589', text='AK-factor: zie blok 1 (Instellingen)', style=s.style('A589'))
    hide_rows(s, 11, 12); hide_rows(s, 17, 24); hide_rows(s, 26, 27); hide_rows(s, 338, 562); hide_rows(s, 589, 590)
    s.set('F7', text='Status', style=st['hdr'])
    checks = {8: 'IF(OR(D8="",D8="Regio (invullen)"),"⚠ nog invullen","✔")', 9: 'IF(ISNUMBER(D9),"✔","⚠ nog invullen")',
              10: 'IF(ISNUMBER(D10),IF(YEAR(D10)<>D9,"⚠ peildatum hoort niet bij het rapportagejaar","✔"),"⚠ nog invullen")',
              13: 'IF(ISNUMBER(D13),"✔","⚠ nog invullen")', 14: 'IF(OR(D14="Budget",D14="Q1",D14="Q2",D14="Q3",D14="Q4"),IF(OR(D14="Budget",D14="Q"&ROUNDUP(MONTH(D10)/3,0)),"✔","⚠ versie past niet bij de peildatum"),"⚠ nog invullen")',
              15: 'IF(ISNUMBER(D15),"✔","⚠ nog invullen")', 16: 'IF(OR(D16="EUR",D16="EUR x 1.000",D16="EUR mln"),"✔","⚠ nog invullen")',
              25: 'IF(AND(ISNUMBER(D25),D25>0),"✔","⚠ nog invullen")'}
    for r, f in checks.items(): s.set(f'F{r}', is_str=True, formula=f, style=st['txt'])
    # blokken hernummeren
    s.set('A133', text='2. Aanvullende RDFO-posten', style=s.style('A133'))
    s.set('A563', text='3. Deelgebieden', style=s.style('A563'))
    s.set('A564', text='Maak hier de deelgebieden aan. Op de Projectlijst kies je per project een deelgebied (leeg = geen). Het Dashboard filtert erop.', style=s.style('A564'))
    s.set('A565', text='Deelgebied', style=s.style('A565'))
    s.set('C563', is_str=True, formula='IF(COUNTA(Gebieden)=0,"⚠ nog geen deelgebieden (optioneel)","✔ "&COUNTA(Gebieden)&" deelgebied(en)")', style=st['txt'])
    s.set('A583', text='4. Funnel en target per jaar', style=s.style('A583'))
    s.set('A591', text='5. DAEB-omzet en verdeling over jaren', style=s.style('A591'))
    nav = [('C3', '1. Instellingen', 'A5'), ('E3', '2. Aanvullende posten', 'A133'), ('G3', '3. Deelgebieden', 'A563'),
           ('I3', '4. Funnel en target', 'A583'), ('K3', '5. DAEB', 'A591')]
    hyperlinks_keep(s, {'N5', 'N133', 'N563', 'N583', 'N591'})
    for c in ('M3', 'C4', 'E4', 'G4', 'I4', 'K4'): clear(s, c)
    for ref, t, loc in nav:
        s.set(ref, text=t, style=s.style('C3')); add_hyperlink(s, ref, f"'Invoer'!{loc}", t)
    s.set('C4', is_str=True, style=st['txt'], formula=
          'IF(COUNTIF(F8:F25,"⚠*")=0,"✔ Instellingen compleet","⚠ "&COUNTIF(F8:F25,"⚠*")&" instelling(en) nog invullen of controleren (kolom F)")'
          '&"   ·   Projectlijst: "&IF(COUNTIFS(Projectlijst!$A$8:$A$107,"?*",Projectlijst!$C$8:$C$107,"")=0,"✔ compleet","⚠ "&COUNTIFS(Projectlijst!$A$8:$A$107,"?*",Projectlijst!$C$8:$C$107,"")&" project(en) zonder soort")')
    add_hyperlink(s, 'C4', "'Projectlijst'!A1", 'Projectlijst')

def akdekking(bk):
    s = bk.sheet('AK-dekking')
    s.set('E20', text='Bron AK %\n(klik = regel)', style=s.style('E20'))
    for r in range(21, 121):
        s.set(f'E{r}', is_str=True, style=s.style(f'E{r}'), formula=(
            f'IF(C{r}<>"Ja","",IF(INDEX(Projectlijst!$I$8:$I$107,$AW{r})<>"","eigen",IF(INDEX(Invoer!$V$31:$V$130,$AW{r})<>"",'
            f'HYPERLINK("#\'Bron_AK\'!A"&IFERROR(MATCH($A{r},INDEX(Bron_AK!$A$1:$Z$3000,0,AK_KolWBS),0),1),"SAP →"),"ontbreekt")))'))

def projecten(bk, dxf_red, dxf_green):
    s = bk.sheet('Projecten'); inp = bk.sheet('Invoer').style('D14')
    s.set('Q9', is_str=True, style=s.style('Q9'), formula='"Mutatie"&CHAR(10)&Sel_Versie&" t.o.v. "&IF(Sel_Prev="","vorige",Sel_Prev)')
    s.set('S9', is_str=True, style=s.style('S9'), formula='"Toelichting mutatie"&CHAR(10)&Sel_Versie&" t.o.v. "&IF(Sel_Prev="","vorige",Sel_Prev)&" (typ hier)"')
    for r in range(10, 110):
        el = s.cell(f'S{r}')
        for ch in list(el): el.remove(ch)
        if 't' in el.attrib: del el.attrib['t']
        el.set('s', inp)
    for sq in ('Q10:Q109', 'P10:P109'):
        add_cf(s, sq, f'AND(ISNUMBER({sq.split(":")[0]}),{sq.split(":")[0]}<0)', dxf_red, prio=1)
        add_cf(s, sq, f'AND(ISNUMBER({sq.split(":")[0]}),{sq.split(":")[0]}>0)', dxf_green, prio=2)
    cols = s.root.find(N+'cols')
    for c in cols:
        if int(c.get('min')) <= 19 <= int(c.get('max')) and c.get('min') == c.get('max'): c.set('width', '44')

def mijlpalen(bk, st):
    m = bk.add_sheet('Bron_Mijlpalen', skeleton(tab='FFFEB900'))
    m.set('A1', text='Plak hier de mijlpalen-export (hele blad, vanaf A1): rij 1 = mijlpalen, rij 2 = verwacht/werkelijk, kolom A = "WBS - naam".')
    p = bk.sheet('Projectdetail')
    T, H, LBL, DATE = p.style('B15'), p.style('B16'), p.style('B17'), bk.sheet('Historie').style('C6')
    r0 = 62
    p.set(f'B{r0}', text='Mijlpalen (tabblad Bron_Mijlpalen)', style=T)
    p.set(f'H{r0}', formula='IF(OR(Det_Idx=0,Det_WBS=""),0,IFERROR(MATCH(Det_WBS&" *",Bron_Mijlpalen!$A:$A,0),IFERROR(MATCH(Det_WBS&"*",Bron_Mijlpalen!$A:$A,0),0)))')
    p.set(f'I{r0}', is_str=True, formula=f'IF(H{r0}=0,IF(Det_Idx=0,"","geen mijlpalen gevonden voor dit project"),"")')
    for j, h in enumerate(['Mijlpaal', 'Verwacht', 'Werkelijk', 'Status']): p.set(f'{"BCDE"[j]}{r0+1}', text=h, style=H)
    for k in range(1, 17):
        r = r0 + 1 + k; cn = 2*k; cw = 2*k + 1
        name = f'IFERROR(INDEX(Bron_Mijlpalen!$1:$1,{cn})&"","")'
        p.set(f'B{r}', is_str=True, style=LBL, formula=f'IF($H${r0}=0,"",{name})')
        for col, cc in (('C', cn), ('D', cw)):
            p.set(f'{col}{r}', style=DATE, formula=f'IF(OR($H${r0}=0,$B{r}=""),"",IF(INDEX(Bron_Mijlpalen!$A$1:$AZ$1000,$H${r0},{cc})="","",INDEX(Bron_Mijlpalen!$A$1:$AZ$1000,$H${r0},{cc})))')
        p.set(f'E{r}', is_str=True, style=LBL, formula=(f'IF($B{r}="","",IF(ISNUMBER(D{r}),IF(AND(ISNUMBER(C{r}),D{r}>C{r}),"gerealiseerd (later dan verwacht)","gerealiseerd"),'
                                                         f'IF(ISNUMBER(C{r}),IF(C{r}<TODAY(),"verlopen (nog niet gerealiseerd)","gepland"),"–")))'))
    dim = p.root.find(N+'dimension'); dim.set('ref', f'A1:V{r0+17}')

BANNER = {'RDFO': [('CALC_Versies', 'CALC_Versies'), ('Versies', 'Versies')], 'Projecten': [('VERW_Project', 'VERW_Project')],
          'Projectdetail': [('VERW_Project', 'VERW_Project'), ('Bron_Mijlpalen', 'Mijlpalen')], 'Woningetalage': [('STG_Woningetalage', 'STG_Woningetalage')],
          'Outlook': [('VERW_Project', 'VERW_Project')], 'Bandbreedte': [('SCEN', 'SCEN')], 'AK-dekking': [('STG_AKR', 'STG_AKR'), ('Bron_AK', 'Bron_AK')],
          "Kansen en risico's": [('Projectlijst', 'Projectlijst')], 'Reserves': [('Projectlijst', 'Projectlijst')], 'Dashboard': [('CALC_Versies', 'CALC_Versies')],
          'Versies': [('VERSIE_REKEN', 'VERSIE_REKEN')]}
def banners(bk):
    for sh, links in BANNER.items():
        s = bk.sheet(sh); row = 3
        used = {split(el.get('r'))[0] for el in s.rows.get(row, etree.Element('x')).findall(N+'c') if len(el)}
        free = [c for c in ['B','C','D','E','F','G','H','I','J','K','L','M','N'] if c not in used]
        # zoek 1 vrije cel voor de filtertekst en daarna vrije cellen voor de links
        if not free: print('geen vrije cel op', sh); continue
        sty = bk.sheet('Invoer').style('E9')
        if sh not in ('Versies',):
            s.set(f'{free[0]}{row}', is_str=True, style=sty, formula='"Filter deelgebied: "&IF(OR(Sel_Deelgebied="",Sel_Deelgebied="Alle"),"alle",Sel_Deelgebied)&"  (kies op Dashboard)"')
            free = [c for c in free if col2n(c) >= col2n(free[0]) + 3]
        for (target, label), c in zip(links, free[::2]):
            s.set(f'{c}{row}', text=f'Berekening: {label} →', style=sty)
            add_hyperlink(s, f'{c}{row}', f"'{target}'!A1", label)
    s = bk.sheet('Versies'); add_hyperlink(s, 'B15', "'VERSIE_REKEN'!A1", 'VERSIE_REKEN')

def logos(bk, navy, wit):
    sh = 'PPT'; p = bk.paths[sh]
    rels_fn = p.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if rels_fn in bk.data and b'drawing' in bk.data[rels_fn]: raise SystemExit('PPT heeft al een drawing')
    n = max([int(re.search(r'drawing(\d+)\.xml', k).group(1)) for k in bk.data if re.match(r'xl/drawings/drawing\d+\.xml$', k)] + [0]) + 1
    bk.data['xl/media/heijmans_logo_navy.png'] = open(navy, 'rb').read(); bk.data['xl/media/heijmans_logo_wit.png'] = open(wit, 'rb').read()
    def pic(i, name, rid, row):
        return (f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>23</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row+3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i}" name="{name}" descr="{name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2247900" cy="790575"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>')
    d = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
         + pic(2, 'LogoNavy', 'rId1', 2) + pic(3, 'LogoWit', 'rId2', 7) + '</xdr:wsDr>')
    bk.data[f'xl/drawings/drawing{n}.xml'] = d.encode()
    bk.data[f'xl/drawings/_rels/drawing{n}.xml.rels'] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/heijmans_logo_navy.png"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/heijmans_logo_wit.png"/></Relationships>').encode()
    if rels_fn in bk.data:
        t = bk.text(rels_fn).replace('</Relationships>', f'<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{n}.xml"/></Relationships>')
    else:
        t = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             f'<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{n}.xml"/></Relationships>')
    bk.settext(rels_fn, t)
    s = bk.sheet(sh); dr = get_or_make(s, 'drawing'); dr.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rIdLogo')
    s.root.nsmap  # zorg dat r-namespace bekend is (lxml voegt hem toe)
    ct = bk.text('[Content_Types].xml')
    if 'Extension="png"' not in ct: ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>')
    ct = ct.replace('</Types>', f'<Override PartName="/xl/drawings/drawing{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    bk.settext('[Content_Types].xml', ct)

ORDER = ['Start', 'Dashboard', 'RDFO', 'Projecten', 'Projectdetail', 'Woningetalage', 'Outlook', 'Bandbreedte', 'AK-dekking',
         "Kansen en risico's", 'Reserves', 'Projectlijst', 'Invoer', 'Versies', 'Handleiding', 'Controle',
         'Bron_ZPMW03', 'Budget ZPMW03', 'Q1 ZPMW03', 'Q2 ZPMW03', 'Q3 ZPMW03', 'Q4 ZPMW03', 'Bron_Woningetalage', 'Bron_AK', 'Bron_AKR', 'Bron_Mijlpalen',
         'VERSIE_REKEN', 'CALC_Versies', 'VERW_Project', 'SCEN', 'STG_ZPMW03', 'STG_Woningetalage', 'STG_AKR',
         'Historie', 'Concept', 'Lijsten', 'PPT', 'Dias']
HIDDEN = {'Historie', 'Concept', 'Lijsten', 'PPT', 'Dias'}
GEEL = {'Bron_ZPMW03', 'Budget ZPMW03', 'Q1 ZPMW03', 'Q2 ZPMW03', 'Q3 ZPMW03', 'Q4 ZPMW03', 'Bron_Woningetalage', 'Bron_AK', 'Bron_AKR', 'Bron_Mijlpalen'}
GRIJS = {'VERSIE_REKEN', 'CALC_Versies', 'VERW_Project', 'SCEN', 'STG_ZPMW03', 'STG_Woningetalage', 'STG_AKR'}
def tab_color(s, rgb):
    pr = s.root.find(N+'sheetPr')
    if pr is None:
        pr = etree.Element(N+'sheetPr'); s.root.insert(0, pr)
    tc = pr.find(N+'tabColor')
    if tc is None:
        tc = etree.Element(N+'tabColor'); pr.insert(0, tc)
    for a in list(tc.attrib): del tc.attrib[a]
    tc.set('rgb', rgb)
def reorder(bk):
    for nm in GEEL: tab_color(bk.sheet(nm), 'FFFEB900')
    for nm in GRIJS: tab_color(bk.sheet(nm), 'FFA4B3BB')
    for nm in bk.paths: bk.sheet(nm)   # alle bladen laden (sheetPr kan wijzigen)
    wb = bk.text('xl/workbook.xml')
    sheets = re.findall(r'<sheet [^>]*/>', wb)
    byname = {}
    for el in sheets:
        nm = re.search(r'name="([^"]+)"', el).group(1).replace('&apos;', "'")
        byname[nm] = el
    old = [re.search(r'name="([^"]+)"', el).group(1).replace('&apos;', "'") for el in sheets]
    assert set(ORDER) == set(old), (set(ORDER) ^ set(old))
    new_els = []
    for nm in ORDER:
        el = re.sub(r' state="\w+"', '', byname[nm])
        if nm in HIDDEN: el = el.replace(' r:id=', ' state="hidden" r:id=')
        new_els.append(el)
    wb = re.sub(r'<sheets>.*</sheets>', '<sheets>' + ''.join(new_els) + '</sheets>', wb, flags=re.S)
    idx = {old[i]: ORDER.index(old[i]) for i in range(len(old))}
    wb = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{idx[old[int(m.group(1))]]}"', wb)
    wb = re.sub(r'firstSheet="\d+"', 'firstSheet="0"', wb); wb = re.sub(r'activeTab="\d+"', 'activeTab="0"', wb)
    bk.settext('xl/workbook.xml', wb)
    # alleen Start geselecteerd
    for nm in bk.paths:
        sv = bk.sheet(nm).root.find(N+'sheetViews')
        if sv is None: continue
        for v in sv.findall(N+'sheetView'):
            if 'tabSelected' in v.attrib: del v.attrib['tabSelected']
    sv = bk.sheet('Start').root.find(N+'sheetViews').find(N+'sheetView'); sv.set('tabSelected', '1')

def handleiding_factor(bk):
    pass

if __name__ == '__main__':
    bk = Book(sys.argv[1])
    red = add_dxf(bk, 'FFF8D0D0'); green = add_dxf(bk, 'FFD5EFD5')
    e_to_an(bk); filters(bk); invoer(bk); akdekking(bk); projecten(bk, red, green)
    mijlpalen(bk, None); banners(bk); logos(bk, sys.argv[3], sys.argv[4]); reorder(bk)
    bk.save(sys.argv[2])
