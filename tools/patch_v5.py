"""Ronde 5: remake. Budget = alleen OP.BOB jaar; OPR i.p.v. BOR; '–' i.p.v. n.b.; RDFO zonder upload-kolom en met
vrijval reserves per kwartaal; Projecten compacter met filter; Projectlijst: Strapo, woningen B2B/B2C, vrijvalkwartalen;
mijlpalen-heatmap; Woningetalage-filter; Outlook t/m RJ+5; bandbreedte optie A; nieuwe dia-tabellen; knoppen.
usage: patch_v5.py in.xlsm out.xlsm"""
import sys, os, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import split, col2n, n2col, N
from patch_eenvoud import add_cf, add_dv_list, add_hyperlink, expand_shared, get_or_make, add_dxf, hide_rows, clear
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI
from lxml import etree

def formulas(s):
    for row in s.rows.values():
        for el in row.findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.text: yield el, f

def hide_cols(s, letters):
    cols = s.root.find(N+'cols')
    for L in letters:
        n = col2n(L); done = False
        for c in list(cols):
            mn, mx = int(c.get('min')), int(c.get('max'))
            if mn <= n <= mx:
                parts = [(mn, n-1, False), (n, n, True), (n+1, mx, False)]
                for a, b, hid in parts:
                    if a <= b:
                        e = copy.deepcopy(c); e.set('min', str(a)); e.set('max', str(b))
                        if hid: e.set('hidden', '1')
                        c.addprevious(e)
                cols.remove(c); done = True; break
        if not done:
            e = etree.SubElement(cols, N+'col'); e.set('min', str(n)); e.set('max', str(n)); e.set('width', '9'); e.set('hidden', '1'); e.set('customWidth', '1')
    items = sorted(cols, key=lambda c: int(c.get('min')))
    for c in list(cols): cols.remove(c)
    for c in items: cols.append(c)
def unhide_col(s, L):
    n = col2n(L)
    for c in s.root.find(N+'cols'):
        if int(c.get('min')) <= n <= int(c.get('max')) and c.get('hidden'):
            mn, mx = int(c.get('min')), int(c.get('max'))
            for a, b, hid in ((mn, n-1, True), (n, n, False), (n+1, mx, True)):
                if a <= b:
                    e = copy.deepcopy(c); e.set('min', str(a)); e.set('max', str(b))
                    if not hid and 'hidden' in e.attrib: del e.attrib['hidden']
                    c.addprevious(e)
            c.getparent().remove(c); break

REPORT = ['Start', 'Dashboard', 'RDFO', 'Projecten', 'Projectdetail', 'Woningetalage', 'Outlook', 'Bandbreedte', 'AK-dekking',
          "Kansen en risico's", 'Reserves', 'Projectlijst', 'PPT', 'Dias', 'Versies', 'Handleiding', 'Invoer', 'Controle']
KEYS = {'BOR FY', 'BOR actual YTD', 'BOR'}
def labels(bk):
    # 1. n.b. -> – (overal, consequent)
    for nm in bk.paths:
        s = bk.sheet(nm)
        for el, f in formulas(s):
            if '"n.b."' in f.text: f.text = f.text.replace('"n.b."', '"–"')
    # 2. BOR -> OPR in teksten (niet in sleutels)
    # gedeelde teksten NIET globaal wijzigen (ook kolomkoppen van geplakte exports staan daar); alleen labelcellen
    # op de rapportbladen worden losse tekst met OPR
    sst = [re.sub(r'<[^>]+>', '', m) for m in re.findall(r'<si>(.*?)</si>', bk.text('xl/sharedStrings.xml'), re.S)]
    import html
    sst = [html.unescape(x) for x in sst]
    for nm in REPORT:
        s = bk.sheet(nm)
        for row in s.rows.values():
            for el in row.findall(N+'c'):
                if el.get('t') != 's' or el.find(N+'f') is not None: continue
                v = el.find(N+'v')
                if v is None: continue
                txt = sst[int(v.text)]
                if 'BOR' not in txt or txt in KEYS: continue
                new = re.sub(r'(?<![A-Za-z_.])BOR(?![A-Za-z_])', 'OPR', txt)
                if new == txt: continue
                for ch in list(el): el.remove(ch)
                el.set('t', 'inlineStr'); is_ = etree.SubElement(el, N+'is'); t = etree.SubElement(is_, N+'t'); t.text = new
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    lit = re.compile(r'"([^"]*)"')
    for nm in REPORT:
        s = bk.sheet(nm)
        for el, f in formulas(s):
            t = f.text
            def rl(m):
                before = t[max(0, m.start()-12):m.start()]
                if 'Kengetal' in before or m.group(1) in KEYS: return m.group(0)
                return '"' + re.sub(r'(?<![A-Za-z_.])BOR(?![A-Za-z_])', 'OPR', m.group(1)) + '"'
            f.text = lit.sub(rl, t)
        for row in s.rows.values():
            for el in row.findall(N+'c'):
                if el.get('t') == 'inlineStr':
                    for tt in el.iter(N+'t'):
                        if tt.text and tt.text not in KEYS: tt.text = re.sub(r'(?<![A-Za-z_.])BOR(?![A-Za-z_])', 'OPR', tt.text)

def budget(bk):
    v = bk.sheet('Versies')
    v.set('C6', text='n.v.t.', style=v.style('C6')); v.set('D6', formula='DATE(P_VJ,1,1)', style=v.style('D6'))
    v.set('E6', formula='P_VJ', style=v.style('E6')); v.set('F6', formula='P_VJ', style=v.style('F6'))
    v.set('B3', style=v.style('B3'), text='Budget = de begroting voor het rapportagejaar: alleen OP.BOB / OP.BOR / Kosten van dat jaar (kolom AR e.v.), zonder realisatie LJ. Q1–Q4 = Bedrijfsopbrengsten LJ + OP.BOB jaar.')
    R = bk.sheet('VERSIE_REKEN')
    for w in range(6, 156):
        R.set(f'BM{w}', formula=f'IF($AA{w}=0,"",-N(AF{w}))'); R.set(f'BN{w}', formula=f'IF($AA{w}=0,"",-N(AG{w}))')
        R.set(f'BO{w}', formula='""'); R.set(f'BP{w}', formula='""')
        R.set(f'BQ{w}', formula=f'IF(OR($AA{w}=0,$BL{w}=""),"",N(AH{w})/P_AKFactor*$BL{w})')
        R.set(f'BR{w}', formula=f'IF($AA{w}=0,"",-N(AI{w}))'); R.set(f'BS{w}', formula=f'IF($AA{w}=0,"",-N(AJ{w}))')
        R.set(f'BT{w}', formula=f'IF(OR($AA{w}=0,$BL{w}=""),"",N(AK{w})/P_AKFactor*$BL{w})')
    for el, f in formulas(R):
        r = split(el.get('r'))
        if r[0] == 'J' and r[1] >= 6 and not f.text.startswith('IF(O'):
            f.text = f'IF(O{r[1]}="","",{f.text})'
        if r[0] == 'M' and r[1] >= 6 and '$CD$' in f.text:
            f.text = re.sub(r'IF\(\$CD\$(\d+)<>"",\$CD\$\1,"Opstal"\)', r'IF(OR($CD$\1="",$CD$\1="Strapo"),"Opstal",$CD$\1)', f.text)

def rdfo(bk):
    s = bk.sheet('RDFO'); hide_cols(s, ['H'])
    st = dict(h=s.style('C33'), l=s.style('B34'), n=s.style('C34'), t=s.style('B37'), tn=s.style('C37'), sec=s.style('B33'))
    for r in range(39, 47):
        if r in s.rows:
            for el in list(s.rows[r].findall(N+'c')): s.rows[r].remove(el)
    s.set('B39', is_str=True, formula='"Vrijval reserves "&P_RJ&" ("&P_Eenheid&")  ·  klik + om open te klappen"', style=st['sec'])
    for q in range(4): s.set(f'{"CDEF"[q]}39', is_str=True, formula=f'P_RJ&"-Q{q+1}"', style=st['h'])
    s.set('G39', text='Jaar', style=st['h'])
    for r, lab, col, kcol in ((40, 'Commerciële reserve', 'L', 'AR'), (41, 'Operationele reserve', 'M', 'AS')):
        s.set(f'B{r}', text=lab, style=st['l'])
        for q in range(4):
            c = 'CDEF'[q]
            s.set(f'{c}{r}', style=st['n'], formula=f'SUMIFS(Projectlijst!${col}$8:${col}$107,Projectlijst!${kcol}$8:${kcol}$107,{c}$39,PL_Telt,"Ja")/P_Factor')
        s.set(f'G{r}', style=st['n'], formula=f'SUM(C{r}:F{r})')
        row = s.row(r); row.set('outlineLevel', '1'); row.set('hidden', '1')
    s.set('B42', text='Vrijval reserves (plus op het resultaat)', style=st['t'])
    for c in 'CDEFG': s.set(f'{c}42', style=st['tn'], formula=f'SUM({c}40:{c}41)')
    s.row(42).set('collapsed', '1')
    s.set('B43', text='OPR-resultaat incl. vrijval (geselecteerde versie)', style=st['l'])
    s.set('G43', style=st['n'], formula='IF(AND(ISNUMBER(J22),Sel_Doeljaar=P_RJ),J22+G42,"–")')
    pr = s.root.find(N+'sheetPr')
    if pr is None: pr = etree.Element(N+'sheetPr'); s.root.insert(0, pr)
    op = pr.find(N+'outlinePr')
    if op is None:
        op = etree.Element(N+'outlinePr'); tc = pr.find(N+'tabColor')
        if tc is not None: tc.addnext(op)
        else: pr.insert(0, op)
    op.set('summaryBelow', '1')
    fmt = s.root.find(N+'sheetFormatPr')
    if fmt is not None: fmt.set('outlineLevelRow', '1')
    d = bk.sheet('Dias'); d.set('H21', formula='0', style=d.style('H21'))      # upload-kolom niet in de PowerPoint

def projectlijst(bk, dxf_or):
    p = bk.sheet('Projectlijst')
    for dv in p.root.iter(N+'dataValidation'):
        if dv.get('sqref') == 'C8:C107': dv.find(N+'formula1').text = '"Opstal,Deelneming,Strapo,Buiten scope"'
    hide_cols(p, ['X'])
    H6, H7, IN, NUM = p.style('AG6'), p.style('AG7'), p.style('AG8') or bk.sheet('Invoer').style('D9'), p.style('O8')
    inp = bk.sheet('Invoer').style('D9'); inpt = bk.sheet('Invoer').style('D14')
    p.set('AO6', text='WONINGEN', style=H6); p.set('AR6', text='VRIJVAL RESERVES', style=H6)
    for c, t in (('AO', 'Woningen totaal\n(Woningetalage)'), ('AP', 'Waarvan B2B\n(invullen)'), ('AQ', 'B2C\n(= totaal − B2B)'),
                 ('AR', 'Commerciële reserve\nvalt vrij in'), ('AS', 'Operationele reserve\nvalt vrij in')):
        p.set(f'{c}7', text=t, style=H7)
    for r in range(8, 108):
        p.set(f'AO{r}', style=NUM, formula=f'IF(OR($A{r}="",COUNTIF(STG_Woningetalage!$B$6:$B$505,$A{r})=0),"",SUMIFS(STG_Woningetalage!$Q$6:$Q$505,STG_Woningetalage!$B$6:$B$505,$A{r}))')
        p.set(f'AP{r}', style=inp)
        p.set(f'AQ{r}', style=NUM, formula=f'IF(AO{r}="","",AO{r}-N(AP{r}))')
        p.set(f'AR{r}', style=inpt); p.set(f'AS{r}', style=inpt)
    add_dv_list(p, 'AR8:AS107', 'L_Kwartalen')
    p.set('AT6', text='MIJLPALEN', style=H6); p.set('AT7', text='WBS in mijlpalen-dump\n(leeg = zelfde WBS)', style=H7)
    for r in range(8, 108): p.set(f'AT{r}', style=inpt)
    add_cf(p, 'AQ8:AQ107', 'AND(ISNUMBER(AQ8),AQ8<0)', dxf_or, prio=3)
    L = bk.sheet('Lijsten'); L.set('S3', text='L_Kwartalen'); k = 4
    for y in range(3):
        for q in range(1, 5): L.set(f'S{k}', is_str=True, formula=f'(P_RJ+{y})&"-Q{q}"'); k += 1
    bk.set_name('L_Kwartalen', 'Lijsten!$S$4:$S$15')
    inv = bk.sheet('Invoer'); expand_shared(inv, range(31, 131))
    for i in range(100):
        r = 31 + i; q = 8 + i
        inv.set(f'F{r}', is_str=True, style=inv.style(f'F{r}'), formula=f'IF(Projectlijst!$A${q}="","",IF(OR(Projectlijst!$C${q}="",Projectlijst!$C${q}="Strapo"),"Opstal",Projectlijst!$C${q}))')
        inv.set(f'J{r}', is_str=True, style=inv.style(f'J{r}'), formula=f'IF(Projectlijst!$C${q}="Strapo","Ja",IF(Projectlijst!$X${q}="","",Projectlijst!$X${q}))')
    p.set('B2', style=p.style('B2'), text='Gevuld vanuit ZPMW03 (nieuwe projecten komen er bij het openen van dit tabblad vanzelf bij, oranje). Kies de soort: Opstal, Deelneming, Strapo of Buiten scope.')

def projecten(bk, dxf_yel, dxf_grey):
    s = bk.sheet('Projecten'); hide_cols(s, ['D', 'J', 'K'])
    s.set('X9', text='Toon', style=s.style('W9'))
    for r in range(10, 110):
        s.set(f'X{r}', is_str=True, style=s.style(f'W{r}'), formula=f'IF($A{r}="","",IF(COUNTIF(E{r}:I{r},">0.5")+COUNTIF(E{r}:I{r},"<-0.5")+COUNTIF(L{r}:N{r},">0.5")+COUNTIF(L{r}:N{r},"<-0.5")=0,"leeg","met cijfers"))')
    add_cf(s, 'E10:I109', 'E$8=Sel_Versie', dxf_yel, prio=5)
    add_cf(s, 'A10:W109', 'AND($A10<>"",$X10="leeg")', dxf_grey, prio=6)
    af = s.root.find(N+'autoFilter')
    if af is not None: af.set('ref', 'A9:X109')
    wb = bk.text('xl/workbook.xml')
    wb = re.sub(r'(<definedName name="_xlnm\._FilterDatabase" localSheetId="\d+" hidden="1">)Projecten!\$A\$9:\$W\$109', r'\1Projecten!$A$9:$X$109', wb)
    bk.settext('xl/workbook.xml', wb)
    s.set('B7', style=s.style('B7'), text="Geel = kwartaal gekozen op het Dashboard. Mutatie: groen = hoger, rood = lager. Grijs = geen cijfers dit jaar (knop rechtsboven of filter 'Toon' verbergt ze). Niet sorteren.")

def projectdetail(bk, red, orange):
    p = bk.sheet('Projectdetail')
    for r in range(62, 110):
        if r in p.rows:
            for el in list(p.rows[r].findall(N+'c')): p.rows[r].remove(el)
    H, L, DATE, TXT = p.style('B16'), p.style('B17'), bk.sheet('Historie').style('C6'), bk.sheet('Invoer').style('E9')
    key = 'IF(IFERROR(INDEX(Projectlijst!$AT$8:$AT$107,MATCH(Det_WBS,Projectlijst!$A$8:$A$107,0))&"","")<>"",INDEX(Projectlijst!$AT$8:$AT$107,MATCH(Det_WBS,Projectlijst!$A$8:$A$107,0))&"",Det_WBS)'
    p.set('B62', text='Mijlpalen', style=p.style('B15'))
    p.set('L62', formula=f'IF(Det_Idx=0,0,IFERROR(MATCH({key}&" *",Bron_Mijlpalen!$A:$A,0),IFERROR(MATCH({key}&"*",Bron_Mijlpalen!$A:$A,0),0)))')
    p.set('C62', is_str=True, style=TXT, formula='IF(Det_Idx=0,"",IF($L$62=0,"Geen mijlpalen gevonden: vul op de Projectlijst (kolom AT) de WBS uit de mijlpalen-dump in.","Rood = verwachte datum verstreken · oranje = binnen 2 maanden · geen kleur = gerealiseerd"))')
    for r, t in ((63, 'Mijlpaal'), (64, 'Verwacht'), (65, 'Werkelijk'), (66, 'Status')): p.set(f'B{r}', text=t, style=H if r == 63 else L)
    for k in range(8):
        c = CL(3 + k); cv, cw = 2*(k+1), 2*(k+1)+1
        p.set(f'{c}63', is_str=True, style=H, formula=f'IF($L$62=0,"",IFERROR(INDEX(Bron_Mijlpalen!$1:$1,{cv})&"",""))')
        for r, cc in ((64, cv), (65, cw)):
            p.set(f'{c}{r}', style=DATE, formula=f'IF(OR($L$62=0,{c}$63=""),"",IF(INDEX(Bron_Mijlpalen!$A$1:$AZ$1000,$L$62,{cc})="","",INDEX(Bron_Mijlpalen!$A$1:$AZ$1000,$L$62,{cc})))')
        p.set(f'{c}66', is_str=True, style=L, formula=(f'IF({c}$63="","",IF(ISNUMBER({c}65),"✔ gerealiseerd",IF(ISNUMBER({c}64),'
                                                        f'IF({c}64<TODAY(),"verlopen",IF({c}64<=EDATE(TODAY(),2),"< 2 maanden","gepland")),"–")))'))
    r63 = p.row(63); r63.set('ht', '42'); r63.set('customHeight', '1')
    add_cf(p, 'C64:J64', 'AND(ISNUMBER(C64),NOT(ISNUMBER(C65)),C64<TODAY())', red, prio=1)
    add_cf(p, 'C64:J64', 'AND(ISNUMBER(C64),NOT(ISNUMBER(C65)),C64>=TODAY(),C64<=EDATE(TODAY(),2))', orange, prio=2)
    p.root.find(N+'dimension').set('ref', 'A1:V66')

def woningetalage(bk):
    w = bk.sheet('Woningetalage'); inp = bk.sheet('Invoer').style('D14')
    w.set('B3', text='Deelgebied (filter):', style=w.style('B3'))
    w.set('C3', text='Alle', style=inp); add_dv_list(w, 'C3', 'L_DeelgebiedKeuze')

def outlook(bk):
    o = bk.sheet('Outlook'); unhide_col(o, 'I')
    o.set('L5', is_str=True, formula='"Totaal "&(P_RJ+1)&"–"&(P_RJ+5)', style=o.style('L5'))
    for el, f in formulas(o):
        c, r = split(el.get('r'))
        if c == 'L' and re.fullmatch(r'SUM\(E%d:H%d\)' % (r, r), f.text): f.text = f'SUM(E{r}:I{r})'
    d = bk.sheet('Dias')
    for el, f in formulas(d):
        if 'MIN(OUT_N,6)' in f.text: f.text = f.text.replace('MIN(OUT_N,6)', 'MIN(OUT_N,7)')

CHART_A = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>
{ser}<c:gapWidth val="60"/><c:overlap val="100"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>
<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="111F39"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr><c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>
<c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="t"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9DEE8"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"><a:solidFill><a:srgbClr val="6B7385"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr><c:crossAx val="5001"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx></c:plotArea>
<c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr></c:chartSpace>'''
def ser_xml(i, name_ref, val_ref, fill, labels, fmt='+#,##0.0'):
    sp = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' if fill is None else f'<c:spPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></c:spPr>'
    dl = ('<c:dLbls><c:numFmt formatCode="{fmt}" sourceLinked="0"/><c:spPr><a:noFill/></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr>'
          '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>').replace('{fmt}', fmt) if labels else ''
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:f>{name_ref}</c:f></c:strRef></c:tx>{sp}<c:invertIfNegative val="0"/>{dl}'
            f'<c:cat><c:strRef><c:f>Bandbreedte!$Y$15:$Y$17</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>{val_ref}</c:f></c:numRef></c:val></c:ser>')
def bandbreedte(bk):
    b = bk.sheet('Bandbreedte')
    lab, num, txt = b.style('A11'), b.style('B11'), b.style('G6')
    b.set('A11', text='Upside t.o.v. basis (optimistisch − basis)', style=lab)
    b.set('A13', text='Downside t.o.v. basis (laagste − basis)', style=lab)
    for c in 'BCD':
        b.set(f'{c}11', formula=f'IF(AND(ISNUMBER({c}6),ISNUMBER({c}7)),{c}6-{c}7,"–")', style=num)
        b.set(f'{c}13', formula=f'IF(ISNUMBER({c}7),MIN({c}8:{c}10)-{c}7,"–")', style=num)
    b.set('G11', text='extra omzet als de verkopen meevallen', style=txt)
    b.set('G13', text='minder omzet in het slechtste scenario', style=txt)
    b.set('J4', text='Bandbreedte per jaar (EUR mln): rood = downside onder basis, groen = upside boven basis', style=b.style('J4'))
    b.set('Y14', text='jaar'); b.set('Z14', text='onder'); b.set('AA14', text='Downside'); b.set('AB14', text='Upside')
    for k, c in enumerate('BCD'):
        r = 15 + k
        b.set(f'Y{r}', formula=f'{c}5', is_str=True)
        b.set(f'Z{r}', formula=f'IF(ISNUMBER({c}7),MIN({c}8:{c}10,{c}7),0)')
        b.set(f'AA{r}', formula=f'IF(ISNUMBER({c}7),{c}7-Z{r},0)')
        b.set(f'AB{r}', formula=f'IF(ISNUMBER({c}7),MAX(0,{c}6-{c}7),0)')
    sers = (ser_xml(0, 'Bandbreedte!$Z$14', 'Bandbreedte!$Z$15:$Z$17', None, False) +
            ser_xml(1, 'Bandbreedte!$AA$14', 'Bandbreedte!$AA$15:$AA$17', 'E30613', True, '-#,##0.0') +
            ser_xml(2, 'Bandbreedte!$AB$14', 'Bandbreedte!$AB$15:$AB$17', '3B7A1E', True))
    xml = CHART_A.replace('{ser}', sers).encode()
    # grafiek op Bandbreedte en Dias (nr 9) vervangen
    def chart_of(sheet, nr):
        p = bk.paths[sheet]; rels = bk.text(p.replace('worksheets/', 'worksheets/_rels/') + '.rels')
        drw = 'xl/drawings/' + re.search(r'Target="\.\./drawings/([^"]+)"', rels).group(1)
        dr = bk.text(drw); drels = bk.text(drw.replace('drawings/', 'drawings/_rels/') + '.rels')
        rid = re.findall(r'<xdr:graphicFrame.*?r:id="([^"]+)"', dr, re.S)[nr - 1]
        tg = re.search(r'Id="%s"[^>]*Target="\.\./charts/([^"]+)"' % rid, drels) or re.search(r'Target="\.\./charts/([^"]+)"[^>]*Id="%s"' % rid, drels)
        return 'xl/charts/' + tg.group(1)
    for sheet, nr in (('Bandbreedte', 1), ('Dias', 9)):
        fn = chart_of(sheet, nr); bk.data[fn] = xml
    ct = bk.text('[Content_Types].xml')
    bk.settext('[Content_Types].xml', ct)
    p = bk.sheet('PPT'); p.set('H37', text='')
    d = bk.sheet('Dias')
    for rr, src_r in ((187, 13), (188, 12)):          # downside direct onder upside, budget als laatste regel
        d.set(f'B{rr}', formula=f'Bandbreedte!A{src_r}', is_str=True, style=d.style('B186'))
        for c, src in (('C', 'B'), ('D', 'C'), ('E', 'D')): d.set(f'{c}{rr}', formula=f'Bandbreedte!{src}{src_r}', style=d.style(f'{c}186'))
    bk.set_name('T_BB', 'Dias!$B$180:$E$188')

def dias_tabellen(bk, red, orange):
    d = bk.sheet('Dias'); H, T, NUM, DATE = d.style('B236'), d.style('B237'), d.style('D237'), bk.sheet('Historie').style('C6')
    # --- T_WONP: woningetalage per project (activiteit, start verkoop, kwartalen)
    r0 = 670
    w = [2.4, 1.6, 0.9] + [0.7] * 8 + [0.8]
    d.set(f'A{r0-1}', text='gewicht')
    for j, g in enumerate(w): d.set(f'{CL(2+j)}{r0-1}', formula=str(g))
    d.set(f'A{r0}', text='T_WONP')
    heads = ['Project', 'Activiteit', 'Start verkoop'] + [f'=P_RJ&" Q{q}"' for q in range(1, 5)] + [f'=(P_RJ+1)&" Q{q}"' for q in range(1, 5)] + ['Totaal']
    for j, h in enumerate(heads):
        if h.startswith('='): d.set(f'{CL(2+j)}{r0}', formula=h[1:], is_str=True, style=H)
        else: d.set(f'{CL(2+j)}{r0}', text=h, style=H)
    act = 'MATCH("Activiteit",INDEX(B_ZPMW03,P_ZHdr,0),0)'
    for k in range(15):
        r = r0 + 1 + k; src = 71 + k
        d.set(f'B{r}', is_str=True, style=T, formula=f'IF(Woningetalage!$B{src}="","",Woningetalage!$B{src})')
        d.set(f'C{r}', is_str=True, style=T, formula=f'IF(Woningetalage!$L{src}="","",IFERROR(INDEX(B_ZPMW03,MATCH(Woningetalage!$L{src},INDEX(B_ZPMW03,0,STG_ZPMW03!$B$3),0),{act})&"",""))')
        d.set(f'D{r}', style=DATE, formula=f'IF(Woningetalage!$L{src}="","",IFERROR(INDEX(STG_Woningetalage!$N$6:$N$505,MATCH(Woningetalage!$L{src},STG_Woningetalage!$B$6:$B$505,0)),""))')
        for j in range(8): d.set(f'{CL(5+j)}{r}', style=NUM, formula=f'IF(Woningetalage!$B{src}="","",Woningetalage!{CL(3+j)}{src})')
        d.set(f'M{r}', style=NUM, formula=f'IF(Woningetalage!$B{src}="","",Woningetalage!K{src})')
    bk.set_name('T_WONP', f'Dias!$B${r0}:$M${r0+15}')
    # --- T_OIJ: orderintake per jaar
    r1 = 690
    d.set(f'A{r1-1}', text='gewicht')
    for j, g in enumerate([1, 1, 1.4, 1.4, 1]): d.set(f'{CL(2+j)}{r1-1}', formula=str(g))
    d.set(f'A{r1}', text='T_OIJ')
    for j, h in enumerate(['Jaar', '# projecten', '="PEW-omzet ("&P_Eenheid&")"', '="PEW-OPR ("&P_Eenheid&")"', 'Marge %']):
        if h.startswith('='): d.set(f'{CL(2+j)}{r1}', formula=h[1:], is_str=True, style=H)
        else: d.set(f'{CL(2+j)}{r1}', text=h, style=H)
    pct = bk.sheet('Projecten').style('O10')
    for k in range(6):
        r = r1 + 1 + k; y = f'(P_RJ-1+{k})'
        d.set(f'B{r}', formula=y, style=T)
        d.set(f'C{r}', formula=f'COUNTIFS(V_OI_Jaar,{y})', style=NUM)
        d.set(f'D{r}', formula=f'SUMIFS(V_PEWOmz,V_OI_Jaar,{y})/P_Factor', style=NUM)
        d.set(f'E{r}', formula=f'SUMIFS(V_PEWBor,V_OI_Jaar,{y})/P_Factor', style=NUM)
        d.set(f'F{r}', formula=f'IF(D{r}=0,"–",E{r}/D{r})', style=pct)
    bk.set_name('T_OIJ', f'Dias!$B${r1}:$F${r1+6}')
    # --- T_MIJL: mijlpalen-heatmap (projecten x mijlpalen, verwachte datum gekleurd)
    r2 = 705; NM = 8; NPJ = 20
    d.set(f'A{r2-1}', text='gewicht'); d.set(f'B{r2-1}', formula='2.6')
    for k in range(NM): d.set(f'{CL(3+k)}{r2-1}', formula=f'IF(C{r2}="",0,1)'.replace('C', CL(3+k), 1))
    d.set(f'A{r2}', text='T_MIJL'); d.set(f'B{r2}', text='Project', style=H)
    for k in range(NM): d.set(f'{CL(3+k)}{r2}', is_str=True, style=H, formula=f'IFERROR(INDEX(Bron_Mijlpalen!$1:$1,{2*(k+1)})&"","")')
    # hulp: per bronregel (3..62) WBS en volgnummer in kolommen R..T
    for i in range(3, 63):
        rr = r2 + 30 + i
        d.set(f'R{rr}', is_str=True, formula=f'TRIM(Bron_Mijlpalen!$A{i}&"")')
        d.set(f'S{rr}', is_str=True, formula=f'IF(R{rr}="","",LEFT(R{rr},FIND(" - ",R{rr}&" - ")-1))')
        d.set(f'T{rr}', formula=(f'IF(OR(S{rr}="",ISERROR(FIND(".",S{rr}))),"",IF(IF(COUNTIF(Projectlijst!$A$8:$A$107,S{rr})+COUNTIF(Projectlijst!$AT$8:$AT$107,S{rr})>0,COUNTIFS(Projectlijst!$A$8:$A$107,S{rr},PL_Telt,"Ja")+COUNTIFS(Projectlijst!$AT$8:$AT$107,S{rr},PL_Telt,"Ja")>0,OR(Sel_Deelgebied="Alle",Sel_Deelgebied="",Sel_Deelgebied="(geen)")),MAX($T${r2+32}:T{rr-1})+1,""))'))
    d.set(f'T{r2+32}', text='nr')
    for k in range(NPJ):
        r = r2 + 1 + k
        m = f'MATCH({k+1},$T${r2+33}:$T${r2+92},0)'
        d.set(f'Q{r}', formula=f'IFERROR({m}+2,0)')
        d.set(f'B{r}', is_str=True, style=T, formula=f'IF($Q{r}=0,"",LEFT(MID(INDEX(Bron_Mijlpalen!$A:$A,$Q{r}),FIND(" - ",INDEX(Bron_Mijlpalen!$A:$A,$Q{r})&" - ")+3,60),34))')
        for j in range(NM):
            cv, cw = 2*(j+1), 2*(j+1)+1
            vcol, wcol = CL(22 + j), CL(22 + NM + j)          # V.. = verwacht, AD.. = werkelijk
            d.set(f'{vcol}{r}', formula=f'IF($Q{r}=0,"",IF(INDEX(Bron_Mijlpalen!$A$1:$AZ$200,$Q{r},{cv})="","",INDEX(Bron_Mijlpalen!$A$1:$AZ$200,$Q{r},{cv})))')
            d.set(f'{wcol}{r}', formula=f'IF($Q{r}=0,"",IF(INDEX(Bron_Mijlpalen!$A$1:$AZ$200,$Q{r},{cw})="","",INDEX(Bron_Mijlpalen!$A$1:$AZ$200,$Q{r},{cw})))')
            dt = lambda x: f'RIGHT("0"&MONTH({x}),2)&"-"&RIGHT(YEAR({x}),2)'
            d.set(f'{CL(3+j)}{r}', is_str=True, style=T, formula=f'IF(ISNUMBER({wcol}{r}),"✔ "&{dt(wcol+str(r))},IF(ISNUMBER({vcol}{r}),{dt(vcol+str(r))},""))')
    add_cf(d, f'C{r2+1}:J{r2+NPJ}', f'AND(ISNUMBER(V{r2+1}),NOT(ISNUMBER(AD{r2+1})),V{r2+1}<TODAY())', red, prio=1)
    add_cf(d, f'C{r2+1}:J{r2+NPJ}', f'AND(ISNUMBER(V{r2+1}),NOT(ISNUMBER(AD{r2+1})),V{r2+1}>=TODAY(),V{r2+1}<=EDATE(TODAY(),2))', orange, prio=2)
    bk.set_name('T_MIJL', f'Dias!$B${r2}:$J${r2+NPJ}')
    d.root.find(N+'dimension').set('ref', f'A1:AK{r2+95}')

def ppt_spec(bk):
    p = bk.sheet('PPT')
    rows = [
     (2, 'Kerncijfers', '=PPT_Msg1', 'kpi', 'K_DASH', 36, 96, 888, 74, '', 1),
     (2, '', '', 'grafiek', 1, 36, 200, 436, 296, '="Omzet "&P_RJ&": van budget naar "&Sel_Versie&" (EUR mln)"', 1),
     (2, '', '', 'grafiek', 2, 488, 200, 436, 296, 'Omzet per jaar (EUR mln)', 1),
     (3, 'Omzet en resultaat', '="Bedragen in "&P_Eenheid', 'tabel', 'T_RDFO', 36, 96, 888, 0, '', 1),
     (4, 'Projecten en stoplicht', '=PPT_Msg3', 'tabel', 'T_PROJ1', 36, 96, 888, 0, '', 1),
     (5, 'Projecten en stoplicht (vervolg)', '', 'tabel', 'T_PROJ2', 36, 96, 888, 0, '', 1),
     (6, 'Projecten en stoplicht (vervolg)', '', 'tabel', 'T_PROJ3', 36, 96, 888, 0, '', 1),
     (7, 'Prognoseverloop en realisatie', '', 'grafiek', 3, 36, 120, 436, 376, '="Grootste mutaties "&Sel_Prev&" → "&Sel_Versie', 1),
     (7, '', '', 'grafiek', 4, 488, 120, 436, 376, 'Realisatie en nog te gaan, top 10 (EUR mln)', 1),
     (8, 'Outlook', '="Outlook "&(P_RJ+1)&"–"&(P_RJ+5)', 'tabel', 'T_OUT', 36, 96, 888, 0, '', 1),
     (9, 'Outlook in beeld', '', 'grafiek', 5, 36, 120, 290, 376, 'Omzet per jaar (EUR mln)', 1),
     (9, '', '', 'grafiek', 6, 335, 120, 290, 376, 'Woningverkopen per jaar', 1),
     (9, '', '', 'grafiek', 7, 634, 120, 290, 376, 'Orderintake per jaar (EUR mln)', 1),
     (10, 'Bandbreedte', '=PPT_Msg9', 'grafiek', 9, 36, 120, 470, 360, 'Rood = downside onder basis · groen = upside boven basis (EUR mln)', 1),
     (10, '', '', 'tabel', 'T_BB', 530, 120, 394, 0, 'Scenario’s en paden (EUR mln)', 1),
     (11, "Kansen en risico's per project", '=PPT_Msg10', 'grafiek', 10, 36, 110, 888, 370, '', '=IF(COUNTIF(KR_Soort,"?*")>0,1,0)'),
     (12, "Kansen en risico's: overzicht", '', 'tabel', 'T_KR', 36, 96, 888, 0, '', '=IF(COUNTIF(KR_Soort,"?*")>0,1,0)'),
     (13, 'Reserves', '=PPT_Msg12', 'tabel', 'T_RESS', 36, 110, 520, 0, 'Per soort', '=IF(N(Reserves!$C$9)<>0,1,0)'),
     (13, '', '', 'tabel', 'T_RES', 36, 238, 888, 0, 'Per project', '=IF(N(Reserves!$C$9)<>0,1,0)'),
     (14, 'Per gebied', '="Bedragen in "&P_Eenheid', 'tabel', 'T_GEB', 36, 96, 888, 0, '', 1),
     (15, 'Orderintake', '=PPT_Msg6', 'tabel', 'T_OIJ', 36, 120, 400, 0, 'Per jaar', 1),
     (15, '', '', 'tabel', 'T_OI', 460, 120, 464, 0, 'Recentste orderintake', 1),
     (16, 'Woningverkopen', '=PPT_Msg7', 'kpi', 'K_WON', 36, 96, 888, 74, '', 1),
     (16, '', '', 'grafiek', 8, 36, 200, 436, 296, 'Verkopen per kwartaal (aantal)', 1),
     (16, '', '', 'grafiek', 11, 488, 200, 436, 296, 'Jaarverwachting verkopen per versie (aantal)', 1),
     (17, 'Woningetalage per project', '="Verkopen per kwartaal "&P_RJ&"–"&(P_RJ+1)&" · top 15"', 'tabel', 'T_WONP', 36, 100, 888, 0, '', 1),
     (18, 'Mijlpalen', 'Rood = verwachte datum verstreken · oranje = binnen 2 maanden · ✔ = gerealiseerd', 'tabel', 'T_MIJL', 36, 96, 888, 0, '', 1),
     (19, 'Aandachtspunten', '', 'opsomming', 'PPT_Bullets', 36, 100, 888, 390, '', 1),
    ]
    r0 = 40
    for r in range(r0, r0 + 40):
        if r in p.rows:
            for el in list(p.rows[r].findall(N+'c')):
                if col2n(split(el.get('r'))[0]) >= 8: p.rows[r].remove(el)
    for i, row in enumerate(rows):
        r = r0 + i
        for j, v in enumerate(row):
            ref = f'{CL(8+j)}{r}'
            if isinstance(v, (int, float)): p.set(ref, formula=str(v))
            elif isinstance(v, str) and v.startswith('='): p.set(ref, formula=v[1:], is_str=(j in (1, 2, 9)))
            elif v != '': p.set(ref, text=v)
    bk.set_name('PPT_Spec', f'PPT!$H${r0}:$R${r0+len(rows)-1}')
    p.set('H39', text='Dia'); p.set('I39', text='Titel'); p.set('J39', text='Kernboodschap')

BTN = None
def knoppen(bk):
    """Start: alleen 'Maak PowerPoint'. Projectlijst en Controle: geen knoppen. Projecten: filterknop."""
    def drawing_of(sheet):
        p = bk.paths[sheet]; rfn = p.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if rfn not in bk.data: return None, rfn
        m = re.search(r'Target="\.\./drawings/([^"]+)"', bk.text(rfn)); return ('xl/drawings/' + m.group(1)) if m else None, rfn
    start, _ = drawing_of('Start'); d = bk.text(start); tmpl = None
    for a in re.findall(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', d, re.S):
        if 'RDFO_MaakPowerPoint' in a: tmpl = a
        elif 'macro="[0]!RDFO_' in a: d = d.replace(a, '')
    # PowerPoint-knop naar kolom B
    d = d.replace(tmpl, re.sub(r'(<xdr:(?:from|to)><xdr:col>)(\d+)', lambda m: m.group(1) + str(int(m.group(2)) - 1), tmpl))
    bk.settext(start, d)
    for sheet in ('Projectlijst', 'Controle'):
        fn, _ = drawing_of(sheet)
        if fn:
            t = bk.text(fn)
            for a in re.findall(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', t, re.S):
                if 'macro="[0]!RDFO_' in a: t = t.replace(a, '')
            bk.settext(fn, t)
    # Projecten: nieuwe tekening met filterknop
    fn, rfn = drawing_of('Projecten')
    d2 = bk.text(start); d2 = re.sub(r'<a:t>\s*\d+\s+Maak PowerPoint</a:t>', '<a:t>Maak PowerPoint</a:t>', d2); d2 = d2.replace('descr="3  Maak PowerPoint"', 'descr="Maak PowerPoint"'); bk.settext(start, d2)
    knop = tmpl.replace('RDFO_MaakPowerPoint', 'RDFO_FilterProjecten')
    knop = re.sub(r'descr="[^"]*"', 'descr="Toon alleen projecten met cijfers / alles"', knop)
    knop = re.sub(r'(<a:t>)[^<]*(</a:t>)', r'\1Alleen projecten met cijfers / alles\2', knop, count=1)
    knop = re.sub(r'<xdr:from><xdr:col>\d+</xdr:col><xdr:colOff>\d+</xdr:colOff><xdr:row>\d+</xdr:row>', '<xdr:from><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>', knop)
    knop = re.sub(r'<xdr:to><xdr:col>\d+</xdr:col><xdr:colOff>\d+</xdr:colOff><xdr:row>\d+</xdr:row>', '<xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row>', knop)
    head = re.search(r'^.*?<xdr:wsDr[^>]*>', d, re.S).group(0)
    n = max(int(re.search(r'drawing(\d+)\.xml', k).group(1)) for k in bk.data if re.match(r'xl/drawings/drawing\d+\.xml$', k)) + 1
    newfn = f'xl/drawings/drawing{n}.xml'
    bk.data[newfn] = (head + knop + '</xdr:wsDr>').encode()
    if rfn in bk.data:
        bk.settext(rfn, bk.text(rfn).replace('</Relationships>', f'<Relationship Id="rIdKnop" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{n}.xml"/></Relationships>'))
    else:
        bk.settext(rfn, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        f'<Relationship Id="rIdKnop" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{n}.xml"/></Relationships>')
    s = bk.sheet('Projecten'); dr = get_or_make(s, 'drawing'); dr.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rIdKnop')
    ct = bk.text('[Content_Types].xml').replace('</Types>', f'<Override PartName="/xl/drawings/drawing{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    bk.settext('[Content_Types].xml', ct)
    # Start: teksten
    st = bk.sheet('Start')
    for c in ('C12', 'D12'): clear(st, c)
    st.set('B12', text='Maakt de presentatie (Heijmans-stijl).', style=st.style('B12'))

def controle(bk):
    c = bk.sheet('Controle')
    X = 'MAX(0,COUNTIF(INDEX(Bron_Woningetalage!$A$1:$ZZ$501,0,MAX(1,STG_Woningetalage!$B$3)),"?*")-1)'
    c.set('F10', is_str=True, style=c.style('F10'), formula=f'"verwacht "&{X}')
    c.set('G10', is_str=True, style=c.style('G10'), formula=f'IF(E10={X},"OK","FOUT")')

def voortgang(bk):
    st = bk.sheet('Start'); H, L, TXT = st.style('B14'), st.style('C15'), st.style('B12')
    items = [(f"'{v} ZPMW03'", f'{v} ZPMW03', f'N(Versies!$I${6+i})', f'IF(N(Versies!$I${6+i})>0," ("&Versies!$I${6+i}&" projecten)","")&IF(Invoer!$D$14="{v}","  ★ huidig","")') for i, v in enumerate(['Budget', 'Q1', 'Q2', 'Q3', 'Q4'])]
    nw = 'MAX(0,COUNTIF(INDEX(Bron_Woningetalage!$A$1:$ZZ$501,0,MAX(1,STG_Woningetalage!$B$3)),"?*")-1)'
    items += [("'Bron_Woningetalage'", 'Woningetalage', nw, f'IF({nw}>0," ("&{nw}&" regels)","")'),
              ("'Bron_AK'", 'AK-percentages', 'COUNTA(Bron_AK!$A$2:$A$3000)', 'IF(COUNTA(Bron_AK!$A$2:$A$3000)>0," ("&COUNTA(Bron_AK!$A$2:$A$3000)&" regels)","")'),
              ("'Bron_AKR'", 'AK-realisatie (grootboek)', 'N(AKR_Rijen)', 'IF(N(AKR_Rijen)>0," ("&AKR_Rijen&" boekingen)","")'),
              ("'Bron_Mijlpalen'", 'Mijlpalen', 'MAX(0,COUNTA(Bron_Mijlpalen!$A$3:$A$1000))', 'IF(COUNTA(Bron_Mijlpalen!$A$3:$A$1000)>0," ("&COUNTA(Bron_Mijlpalen!$A$3:$A$1000)&" projecten)","")')]
    tel = '+'.join(f'IF({n}>0,1,0)' for _, _, n, _ in items)
    st.set('G7', is_str=True, style=H, formula=f'"Geüpload: "&({tel})&" van {len(items)}"')
    for i, (sheet, label, n, extra) in enumerate(items):
        r = 8 + i
        st.set(f'G{r}', is_str=True, style=TXT, formula=f'IF({n}>0,"✔  ","–  ")&"{label}"&{extra}')
        add_hyperlink(st, f'G{r}', f"{sheet}!A1", label)
    cols = st.root.find(N+'cols')
    for c in cols:
        if c.get('min') == '7' and c.get('max') == '7': c.set('width', '44')

def main(src, dst):
    bk = Book(src)
    red = add_dxf(bk, 'FFF4B6B2'); orange = add_dxf(bk, 'FFFCD9A0'); yel = add_dxf(bk, 'FFFFF2CC'); grey = None
    st = bk.text('xl/styles.xml'); m = re.search(r'<dxfs count="(\d+)"', st); grey = int(m.group(1))
    st = st.replace(m.group(0), f'<dxfs count="{grey+1}"', 1).replace('</dxfs>', '<dxf><font><color rgb="FFA4B3BB"/></font></dxf></dxfs>', 1); bk.settext('xl/styles.xml', st)
    labels(bk); budget(bk); rdfo(bk); projectlijst(bk, orange); projecten(bk, yel, grey); projectdetail(bk, red, orange)
    woningetalage(bk); outlook(bk); bandbreedte(bk); controle(bk); voortgang(bk); dias_tabellen(bk, red, orange); ppt_spec(bk); knoppen(bk)
    bk.save(dst)

if __name__ == '__main__': main(sys.argv[1], sys.argv[2])
