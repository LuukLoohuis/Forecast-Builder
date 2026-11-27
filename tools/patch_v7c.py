"""Ronde 7, deel 3: reserves koppelen op projectnaam (kolom N 'Omschrijving' van ZPMW03) en gekleurde
voortgangstegels op Start."""
import sys, os, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from patch_ak import split, col2n, n2col, N
from patch_eenvoud import add_cf, add_hyperlink, clear, get_or_make
from lxml import etree


# ------------------------------------------------------------------ stijlhulpjes
class Stijlen:
    def __init__(self, bk):
        self.bk = bk; self.x = etree.fromstring(bk.data['xl/styles.xml'])
    def _lijst(self, tag):
        return self.x.find(N + tag)
    def _add(self, tag, el):
        l = self._lijst(tag); l.append(el); l.set('count', str(len(l))); return len(l) - 1
    def font(self, sz, kleur, vet=False):
        f = etree.Element(N+'font')
        if vet: etree.SubElement(f, N+'b')
        etree.SubElement(f, N+'sz', val=str(sz)); etree.SubElement(f, N+'color', rgb='FF' + kleur)
        etree.SubElement(f, N+'name', val='Arial'); etree.SubElement(f, N+'family', val='2')
        return self._add('fonts', f)
    def fill(self, kleur):
        f = etree.Element(N+'fill'); p = etree.SubElement(f, N+'patternFill', patternType='solid')
        etree.SubElement(p, N+'fgColor', rgb='FF' + kleur); etree.SubElement(p, N+'bgColor', indexed='64')
        return self._add('fills', f)
    def rand(self, kleur, dikte='medium'):
        b = etree.Element(N+'border')
        for k in ('left', 'right', 'top', 'bottom'):
            e = etree.SubElement(b, N+k, style=dikte); etree.SubElement(e, N+'color', rgb='FF' + kleur)
        etree.SubElement(b, N+'diagonal')
        return self._add('borders', b)
    def numfmt(self, code):
        l = self._lijst('numFmts')
        nid = max(int(e.get('numFmtId')) for e in l) + 1
        etree.SubElement(l, N+'numFmt', numFmtId=str(nid), formatCode=code); l.set('count', str(len(l))); return nid
    def xf(self, font=0, fill=0, rand=0, numfmt=0, **align):
        x = etree.Element(N+'xf', numFmtId=str(numfmt), fontId=str(font), fillId=str(fill), borderId=str(rand), xfId='0',
                          applyFont='1', applyFill='1', applyBorder='1', applyAlignment='1')
        if numfmt: x.set('applyNumberFormat', '1')
        if align: etree.SubElement(x, N+'alignment', **{k: str(v) for k, v in align.items()})
        return str(self._add('cellXfs', x))
    def dxf(self, fill, kleur, vet=False):
        d = etree.Element(N+'dxf'); f = etree.SubElement(d, N+'font')
        if vet: etree.SubElement(f, N+'b')
        etree.SubElement(f, N+'color', rgb='FF' + kleur)
        fl = etree.SubElement(d, N+'fill'); p = etree.SubElement(fl, N+'patternFill'); etree.SubElement(p, N+'bgColor', rgb='FF' + fill)
        return self._add('dxfs', d)
    def bewaar(self):
        self.bk.data['xl/styles.xml'] = etree.tostring(self.x, xml_declaration=True, encoding='UTF-8', standalone=True)


# ------------------------------------------------------------------ reserves op projectnaam
def reserves_naam(bk, sty):
    b = bk.sheet('Berekening'); H, NUM, TXT, LNK = sty['hdr'], sty['num'], sty['txt'], sty['lnk']
    rows = [(10, 'Regels (WBS of projectnaam)',
             'IF(T6=0,0,COUNTIF(INDEX(Z_Res,0,T6),"?.*")+IF(T11=0,0,COUNTIFS(INDEX(Z_Res,0,T6),"Total",INDEX(Z_Res,0,T11),"<>Total")))'),
            (11, 'Projectnaam (Proj Omschrijving)', 'IF(T5=0,0,IFERROR(MATCH("*Omschrijving*",INDEX(Z_Res,T5,0),0),IFERROR(MATCH("*Omschrijving*",INDEX(Z_Res,T5+1,0),0),0)))'),
            (12, 'ZPMW03-upload: kolom WBS-element', 'IFERROR(MATCH("WBS-element",INDEX(B_ZPMW03,P_ZHdr,0),0),0)'),
            (13, 'ZPMW03-upload: kolom Omschrijving', 'IFERROR(MATCH("Omschrijving",INDEX(B_ZPMW03,P_ZHdr,0),0),0)')]
    for r, lab, f in rows:
        b.set(f'S{r}', text=lab, style=TXT); b.set(f'T{r}', formula=f, style=NUM)
        if r >= 11:
            b.set(f'U{r}', is_str=True, style=TXT, formula=f'IF(T{r}=0,"–",SUBSTITUTE(ADDRESS(1,T{r},4),"1",""))')
    # controle: wat staat er in de dump en wat is gekoppeld aan een project
    b.set('S15', text='Controle reserves', style=H); b.set('T15', text='In dump', style=H); b.set('U15', text='Gekoppeld', style=H); b.set('V15', text='Niet gekoppeld', style=H)
    for r, lab, kol, pl in ((16, 'Commerciële reserve', '$T$7', 'AX'), (17, 'Onvoorzien', '$T$8', 'AY'), (18, 'Nazorg', '$T$9', 'AZ')):
        c = f'INDEX(Z_Res,0,{kol})'; w = 'INDEX(Z_Res,0,$T$6)'; nm = 'INDEX(Z_Res,0,$T$11)'
        b.set(f'S{r}', text=lab, style=TXT)
        b.set(f'T{r}', style=NUM, formula=(f'IF(OR({kol}=0,$T$6=0),0,IF(COUNTIF({w},"?.*")>0,SUMIFS({c},{w},"?.*"),'
                                            f'IF($T$11=0,0,SUMIFS({c},{w},"Total",{nm},"<>Total"))))'))
        b.set(f'U{r}', style=NUM, formula=f'SUM(Projectlijst!${pl}$8:${pl}$107)')
        b.set(f'V{r}', style=NUM, formula=f'T{r}-U{r}')
    b.set('S19', text='Niet gekoppeld = naam of WBS uit de dump staat niet op de Projectlijst / in ZPMW03 kolom N.', style=sty['sub'])
    # Projectlijst: naam voor de koppeling (kolom N van de ZPMW03-upload; anders de naam op de Projectlijst)
    p = bk.sheet('Projectlijst'); H7, NUMr = p.style('AX7'), p.style('AX8')
    p.set('BA6', text='KOPPELING RESERVES', style=p.style('AX6'))
    p.set('BA7', text='Projectnaam voor koppeling\n(ZPMW03 kolom N)', style=H7)
    for r in range(8, 108):
        p.set(f'BA{r}', is_str=True, style=p.style(f'B{r}'),
              formula=(f'IF($A{r}="","",IFERROR(INDEX(INDEX(B_ZPMW03,0,Berekening!$T$13),'
                       f'MATCH($A{r},INDEX(B_ZPMW03,0,Berekening!$T$12),0))&"",$B{r}))'))
        for dmp, kol in (('AX', '$T$7'), ('AY', '$T$8'), ('AZ', '$T$9')):
            c = f'INDEX(Z_Res,0,Berekening!{kol})'; w = 'INDEX(Z_Res,0,Berekening!$T$6)'; nm = 'INDEX(Z_Res,0,Berekening!$T$11)'
            opwbs = f'(SUMIFS({c},{w},$A{r})+SUMIFS({c},{w},$A{r}&".*"))'
            opnaam = f'IF(OR(Berekening!$T$11=0,$BA{r}=""),0,SUMIFS({c},{nm},$BA{r},{w},"<>?.*"))'
            p.set(f'{dmp}{r}', style=NUMr,
                  formula=f'IF(OR($A{r}="",Berekening!{kol}=0,Berekening!$T$6=0),"",IF({opwbs}<>0,{opwbs},{opnaam}))')
    cols = p.root.find(N+'cols')
    if not any(c.get('min') == '53' for c in cols):
        e = etree.SubElement(cols, N+'col'); e.set('min', '53'); e.set('max', '53'); e.set('width', '34'); e.set('customWidth', '1')
    p.root.find(N+'dimension').set('ref', 'A1:BA107')
    rs = bk.sheet('Reserves')
    rs.set('B10', is_str=True, style=rs.style('B2'),
           formula=('IF(SUM(Berekening!$V$16:$V$18)>0.5,"⚠ Niet gekoppeld aan een project: "&TEXT(SUM(Berekening!$V$16:$V$18),"#,##0")'
                    '&" (zie Berekening, kolom S–V)","✔ Alle reserves uit de dump zijn gekoppeld aan een project.")'))
    add_hyperlink(rs, 'B10', "'Berekening'!S15", 'Berekening')


# ------------------------------------------------------------------ Start: gekleurde tegels
def start_tegels(bk):
    st = bk.sheet('Start'); S = Stijlen(bk)
    navy, grijs = '14284B', '6B7385'
    f_tegel = S.font(9, navy); f_kop = S.font(11, navy, True)
    tegel = S.xf(font=f_tegel, fill=S.fill('F3F5F9'), rand=S.rand('FFFFFF', 'thick'), horizontal='left', wrapText=1, vertical='center', indent=1)
    balk = S.xf(font=f_kop, numfmt=S.numfmt('0" van 10 geplakt"'), vertical='center')
    rood = S.dxf('FBE0DD', '9B1C1C', True); geel = S.dxf('FFE699', navy, True); groen = S.dxf('E2EFDA', '1E5E2B', False)
    S.bewaar()
    # oude lijst weg
    for r in range(7, 21):
        for c in ('G', 'H'): clear(st, f'{c}{r}')
    hl = st.root.find(N+'hyperlinks')
    if hl is not None:
        for h in list(hl):
            c, r = split(h.get('ref'))
            if c in ('G', 'H') and 7 <= r <= 20: hl.remove(h)
    vers = []
    for k, v in enumerate(['Budget', 'Q1', 'Q2', 'Q3', 'Q4']):
        n = f'N(Versies!$I${6 + k})'
        vers.append((f"'{v} ZPMW03'", f'{v} ZPMW03', n, f'{n}&" projecten"', f'Invoer!$D$14="{v}"'))
    nw = 'MAX(0,COUNTIF(INDEX(Bron_Woningetalage!$A$1:$ZZ$501,0,MAX(1,STG_Woningetalage!$B$3)),"?*")-1)'
    rest = [("'Bron_Woningetalage'", 'Woningetalage', nw, f'{nw}&" regels"', None),
            ("'Bron_AK'", 'AK-percentages', 'COUNTA(Bron_AK!$A$2:$A$3000)', 'COUNTA(Bron_AK!$A$2:$A$3000)&" regels"', None),
            ("'Bron_AKR'", 'AK-realisatie', 'N(AKR_Rijen)', 'N(AKR_Rijen)&" boekingen"', None),
            ("'Bron_Mijlpalen'", 'Mijlpalen', 'MAX(0,COUNTA(Bron_Mijlpalen!$A$3:$A$1000))', 'MAX(0,COUNTA(Bron_Mijlpalen!$A$3:$A$1000))&" projecten"', None),
            ("'Bron_Reserves'", 'Reserves', 'N(Berekening!$T$10)', 'N(Berekening!$T$10)&IF(N(Berekening!$T$10)=1," regel"," regels")', None)]
    alle = vers + rest
    for i, (blad, label, n, tel, huidig) in enumerate(alle):
        col = 'G' if i < 5 else 'H'; r = 15 + (i % 5); ref = f'{col}{r}'
        ster = f'&IF({huidig},"   ★ huidig","")' if huidig else ''
        st.set(ref, is_str=True, style=tegel,
               formula=f'IF({n}>0,"✔  ","–  ")&"{label}"{ster}&CHAR(10)&IF({n}>0,{tel},"nog niet geplakt")')
        add_hyperlink(st, ref, f'{blad}!A1', label)
    tot = '+'.join(f'IF({n}>0,1,0)' for _, _, n, _, _ in alle)
    st.set('G14', text='Dumps', style=st.style('B14'))
    st.set('H14', formula=tot, style=balk)
    st.set('G20', text='Groen = geplakt · rood = nog plakken · geel = huidig kwartaal. Klik op een tegel om erheen te gaan.', style=st.style('B12'))
    add_cf(st, 'G15:H19', 'LEFT(G15,1)="–"', rood, prio=1)
    add_cf(st, 'G15:H19', 'ISNUMBER(SEARCH("★",G15))', geel, prio=2)
    add_cf(st, 'G15:H19', 'LEFT(G15,1)="✔"', groen, prio=3)
    # voortgangsbalk
    cf = etree.Element(N+'conditionalFormatting'); cf.set('sqref', 'H14')
    ru = etree.SubElement(cf, N+'cfRule', type='dataBar', priority='4')
    db = etree.SubElement(ru, N+'dataBar')
    etree.SubElement(db, N+'cfvo', type='num', val='0'); etree.SubElement(db, N+'cfvo', type='num', val='10')
    etree.SubElement(db, N+'color', rgb='FF70AD47')
    st.root.findall(N+'conditionalFormatting')[-1].addnext(cf)
    cols = st.root.find(N+'cols')
    for c in list(cols):
        if c.get('min') == '7': c.set('width', '27'); c.set('max', '7')
    if not any(c.get('min') == '8' for c in cols):
        e = etree.SubElement(cols, N+'col'); e.set('min', '8'); e.set('max', '8'); e.set('width', '27'); e.set('customWidth', '1')
    st.set('C19', style=st.style('C19'), text="Knop 'Maak PowerPoint'. Rechts zie je welke dumps er staan (groen) en welke nog ontbreken (rood).")
    st.root.find(N+'dimension').set('ref', 'A1:H21')
    st.set('H1', style=st.style('G1'))     # navy titelbalk tot en met de tegels
    wb = bk.text('xl/workbook.xml')
    wb = wb.replace('<definedName name="_xlnm.Print_Area" localSheetId="0">Start!$A$1:$G$35</definedName>',
                    '<definedName name="_xlnm.Print_Area" localSheetId="0">Start!$A$1:$H$35</definedName>')
    bk.settext('xl/workbook.xml', wb)
