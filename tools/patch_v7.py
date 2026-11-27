"""Ronde 7: één eenvoudig rekenblad 'Berekening' (omzet en OPR per project, per versie, per jaar, rechtstreeks uit de
ZPMW03-tabbladen), handmatige correcties per versie op Projecten, reserves-dump, bandbreedte-grafiek staand, dumpknoppen.
usage: patch_v7.py in.xlsm out.xlsm [herberekend.xlsx]"""
import sys, os, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import split, col2n, n2col, N
from patch_eenvoud import add_cf, add_dv_list, add_hyperlink, clear, get_or_make, add_dxf, expand_shared
from lxml import etree

VERS = ['Budget', 'Q1', 'Q2', 'Q3', 'Q4']
NY = 6                       # jaren P_RJ .. P_RJ+5
BW = 17                      # kolommen per versieblok
G0 = 7                       # eerste blokkolom (G)
R_HDR, R_LAB, R_SRC, R_C1, R_C2 = 22, 23, 24, 25, 26
R_OP, R_DN, R_NOL, R_TOT, R_DUMP, R_NIET, R_FOP, R_FDN = 27, 28, 29, 30, 31, 32, 33, 34
R_LIST, R0, R1 = 36, 37, 136   # projectregels 37..136 = Projectlijst 8..107
# Projecten: correctiekolommen per versie (Q1..Q4): omzet, OPR; daarna toelichting
CORR = {'Q1': ('AC', 'AD'), 'Q2': ('AE', 'AF'), 'Q3': ('AG', 'AH'), 'Q4': ('AI', 'AJ')}
CORR_TXT = 'AK'

def L(n): return n2col(n)
def bcol(k, j): return L(G0 + BW * k + j)          # kolomletter in blok k, positie j
def zn(v): return f'Z_{v}'
VI = lambda k: 5 + k                                 # rij in de versietabel (J..Q)
DET = f'Berekening!$G${R0}:${L(G0 + BW * 5 - 1)}${R1}'
DETA = f'Berekening!$A${R0}:$A${R1}'
def DROW(r): return f'Berekening!$G${r}:${L(G0 + BW * 5 - 1)}${r}'

def berekening(bk, sty):
    if 'Berekening' in bk.paths: return bk.sheet('Berekening')
    last = G0 + BW * 5 - 1
    cols = ('<cols><col min="1" max="1" width="17" customWidth="1"/><col min="2" max="2" width="34" customWidth="1"/>'
            '<col min="3" max="3" width="11" customWidth="1"/><col min="4" max="6" width="8" customWidth="1"/>'
            f'<col min="7" max="{last}" width="11" customWidth="1"/></cols>')
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><tabColor rgb="FFA4B3BB"/></sheetPr><dimension ref="A1"/>'
           '<sheetViews><sheetView showGridLines="0" zoomScale="90" workbookViewId="0"><pane xSplit="6" topLeftCell="G1" activePane="topRight" state="frozen"/></sheetView></sheetViews>'
           f'<sheetFormatPr defaultRowHeight="14.4"/>{cols}<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode()
    s = bk.add_sheet('Berekening', xml)
    for k, v in enumerate(VERS):
        bk.set_name(zn(v), f"'{v} ZPMW03'!$A$1:$ZZ$600")
    T, SUB, H, NUM, TXT, LNK, TOT = sty['title'], sty['sub'], sty['hdr'], sty['num'], sty['txt'], sty['lnk'], sty['tot']
    s.set('A1', text='Berekening omzet en OPR per project — rechtstreeks uit de ZPMW03-tabbladen (bedragen in EUR)', style=T)
    s.set('A2', text='Omzet in het jaar van de dump = Bedrijfsopbrengsten LJ (kolom X) + OP.BOB van dat jaar (bijv. AR). Latere jaren = alleen OP.BOB van dat jaar. '
                     'OPR net zo: Bruto (omzet) result. LJ (Z) + OP.BOR. SAP geeft opbrengsten negatief, daarom × −1.', style=SUB)
    s.set('A3', text='Rij 24 laat per kolom zien uit welke kolom(men) van het tabblad het komt. Correcties komen van het tabblad Projecten (kolom AC–AJ). '
                     'Onderaan per project; daarboven de totalen. Projecten die niet op de Projectlijst staan tellen als opstal (rij 29).', style=SUB)
    # ---- versietabel J4:Q9
    for c, t in zip('JKLMNOPQ', ['Versie', 'Tabblad', 'Koprij', 'Kolom WBS (nr)', 'Kolom WBS', 'Dumpjaar', 'Regels', 'Status']):
        s.set(f'{c}4', text=t, style=H)
    for k, v in enumerate(VERS):
        r = VI(k); tab = f"'{v} ZPMW03'"
        s.set(f'J{r}', text=v, style=TXT)
        s.set(f'K{r}', text=f'{v} ZPMW03 →', style=LNK); add_hyperlink(s, f'K{r}', f'{tab}!A1', f'{v} ZPMW03')
        kop = '0'
        for rr in (5, 4, 3, 2, 1): kop = f'IF(COUNTIF({tab}!${rr}:${rr},"WBS-element")>0,{rr},{kop})'
        s.set(f'L{r}', formula=kop, style=NUM)
        s.set(f'M{r}', formula=f'IF(L{r}=0,0,IFERROR(MATCH("WBS-element",INDEX({zn(v)},L{r},0),0),0))', style=NUM)
        s.set(f'N{r}', is_str=True, formula=f'IF(M{r}=0,"–",SUBSTITUTE(ADDRESS(1,M{r},4),"1",""))', style=TXT)
        s.set(f'O{r}', formula=f'IF(L{r}=0,0,IFERROR(VALUE(MID(INDEX(INDEX({zn(v)},L{r},0),MATCH("OP.BOB ????-Q1",INDEX({zn(v)},L{r},0),0)),8,4)),0))', style=TXT)
        s.set(f'P{r}', formula=f'IF(M{r}=0,0,MAX(0,COUNTIF(INDEX({zn(v)},0,M{r}),"?*")-1))', style=NUM)
        s.set(f'Q{r}', is_str=True, formula=f'IF(P{r}>0,"✔ geplakt","– leeg")', style=TXT)
    s.set('H4', text='Jaar-index', style=H); s.set('H5', formula='MAX(0,MIN(5,Sel_Doeljaar-P_RJ))', style=NUM)
    # ---- detail: kopregels per blok
    labels = []
    for y in range(NY):
        labels += [(f'"Omzet"&CHAR(10)&(P_RJ+{y})', 'Bedrijfsopbrengsten LJ', f'"OP.BOB "&(P_RJ+{y})', y),
                   (f'"OPR"&CHAR(10)&(P_RJ+{y})', 'Bruto (omzet) result. LJ', f'"OP.BOR "&(P_RJ+{y})', y)]
    for k, v in enumerate(VERS):
        vr = VI(k); Z = zn(v); kop, wbs, dj = f'$L${vr}', f'$M${vr}', f'$O${vr}'
        c0 = bcol(k, 0)
        s.set(f'{c0}{R_HDR}', text=f'{v} ZPMW03 →', style=LNK); add_hyperlink(s, f'{c0}{R_HDR}', f"'{v} ZPMW03'!A1", f'{v} ZPMW03')
        s.set(f'{bcol(k, 2)}{R_HDR}', is_str=True, style=SUB,
              formula=f'IF({wbs}=0,"tabblad leeg",{dj}&"-dump · WBS in kolom "&$N${vr}&" · "&$P${vr}&" regels")')
        cnt = bcol(k, 16)
        for j in range(BW):
            c = bcol(k, j)
            if j < 12:
                lab, lj, op, y = labels[j]
                s.set(f'{c}{R_LAB}', is_str=True, formula=lab, style=H)
                s.set(f'{c}{R_C1}', style=NUM, formula=(f'IF({wbs}=0,0,IF((P_RJ+{y})={dj},IFERROR(MATCH("{lj}",INDEX({Z},{kop},0),0),0),'
                                                       f'IFERROR(MATCH({op},INDEX({Z},{kop},0),0),0)))'))
                s.set(f'{c}{R_C2}', style=NUM, formula=f'IF(OR({wbs}=0,(P_RJ+{y})<>{dj}),0,IFERROR(MATCH({op},INDEX({Z},{kop},0),0),0))')
            elif j in (12, 13):
                s.set(f'{c}{R_LAB}', text=('Omzet' if j == 12 else 'OPR') + '\nt/m peildatum', style=H)
                nm = 'Bedrijfsopbrengsten LJ' if j == 12 else 'Bruto (omzet) result. LJ'
                s.set(f'{c}{R_C1}', style=NUM, formula=f'IF({wbs}=0,0,IFERROR(MATCH("{nm}",INDEX({Z},{kop},0),0),0))')
                s.set(f'{c}{R_C2}', formula='0', style=NUM)
            elif j in (14, 15):
                s.set(f'{c}{R_LAB}', is_str=True, style=H, formula=f'"Correctie "&{"" if True else ""}"{"omzet" if j == 14 else "OPR"}"&CHAR(10)&P_RJ')
            else:
                s.set(f'{c}{R_LAB}', text='Regels\nin dump', style=H)
            # bronregel
            if j < 14:
                s.set(f'{c}{R_SRC}', is_str=True, style=SUB,
                      formula=f'IF({c}${R_C1}=0,"–",SUBSTITUTE(ADDRESS(1,{c}${R_C1},4),"1","")&IF({c}${R_C2}>0," + "&SUBSTITUTE(ADDRESS(1,{c}${R_C2},4),"1",""),""))')
            elif j in (14, 15):
                s.set(f'{c}{R_SRC}', text=('Projecten ' + CORR[v][j - 14]) if v in CORR else 'geen correcties', style=SUB)
            else:
                s.set(f'{c}{R_SRC}', is_str=True, style=SUB, formula=f'"kolom "&$N${vr}')
            # projectregels
            for r in range(R0, R1 + 1):
                if j < 14:
                    f = (f'IF(OR($A{r}="",{c}${R_C1}=0,${cnt}{r}=0),"",-(SUMIFS(INDEX({Z},0,{c}${R_C1}),INDEX({Z},0,{wbs}),$A{r})'
                         f'+IF({c}${R_C2}>0,SUMIFS(INDEX({Z},0,{c}${R_C2}),INDEX({Z},0,{wbs}),$A{r}),0)))')
                    s.set(f'{c}{r}', formula=f, style=NUM)
                elif j in (14, 15):
                    if v in CORR:
                        pc = CORR[v][j - 14]
                        s.set(f'{c}{r}', style=NUM, formula=f'IF($A{r}="","",IFERROR(N(INDEX(Projecten!${pc}$10:${pc}$109,MATCH($A{r},Projecten!$A$10:$A$109,0)))*P_Factor,0))')
                else:
                    s.set(f'{c}{r}', style=NUM, formula=f'IF(OR($A{r}="",{wbs}=0),0,COUNTIF(INDEX({Z},0,{wbs}),$A{r}))')
            # totaalregels
            rng = f'{c}{R0}:{c}{R1}'
            if j < 16 and not (j in (14, 15) and v not in CORR):
                s.set(f'{c}{R_OP}', style=TOT, formula=f'SUMIFS({rng},$C${R0}:$C${R1},"Opstal",$D${R0}:$D${R1},"Ja")')
                s.set(f'{c}{R_DN}', style=TOT, formula=f'SUMIFS({rng},$C${R0}:$C${R1},"Deelneming",$D${R0}:$D${R1},"Ja")')
                s.set(f'{c}{R_FOP}', style=NUM, formula=f'SUMIFS({rng},$C${R0}:$C${R1},"Opstal",$F${R0}:$F${R1},"Ja")'
                      + (f'+IF(AND(Sel_WBS="*",OR(Sel_TypeCrit="*",Sel_TypeCrit="Opstal")),{c}{R_NOL},0)' if j < 14 else ''))
                s.set(f'{c}{R_FDN}', style=NUM, formula=f'SUMIFS({rng},$C${R0}:$C${R1},"Deelneming",$F${R0}:$F${R1},"Ja")')
                if j < 14:
                    s.set(f'{c}{R_DUMP}', style=NUM, formula=(f'IF({c}${R_C1}=0,0,-(SUMIFS(INDEX({Z},0,{c}${R_C1}),INDEX({Z},0,{wbs}),"?*")'
                                                              f'+IF({c}${R_C2}>0,SUMIFS(INDEX({Z},0,{c}${R_C2}),INDEX({Z},0,{wbs}),"?*"),0)))'))
                    s.set(f'{c}{R_NOL}', style=NUM, formula=f'IF(OR(Sel_Deelgebied="Alle",Sel_Deelgebied="",Sel_Deelgebied="(geen)"),{c}{R_DUMP}-SUM({rng}),0)')
                    s.set(f'{c}{R_NIET}', style=NUM, formula=f'{c}{R_DUMP}-{c}{R_TOT}')
                else:
                    s.set(f'{c}{R_NOL}', formula='0', style=NUM)
                s.set(f'{c}{R_TOT}', style=TOT, formula=f'{c}{R_OP}+{c}{R_DN}+{c}{R_NOL}')
    for r, t in ((R_OP, 'Opstal (telt mee)'), (R_DN, 'Deelnemingen (telt mee)'), (R_NOL, 'Niet op de Projectlijst (telt als opstal)'),
                 (R_TOT, 'Totaal meegeteld'), (R_DUMP, 'Controle: hele dump (alle regels met WBS)'), (R_NIET, 'Niet meegeteld (buiten scope, Nee of deelgebied)'),
                 (R_FOP, 'Opstal, met filter Dashboard (type/project)'), (R_FDN, 'Deelnemingen, met filter Dashboard')):
        s.set(f'B{r}', text=t, style=TOT if r in (R_OP, R_DN, R_TOT) else TXT)
    s.set(f'B{R_LAB}', text='Kengetal en jaar', style=H); s.set(f'B{R_SRC}', text='Kolom(men) in het tabblad', style=SUB)
    for r in (R_C1, R_C2): s.row(r).set('hidden', '1')
    for c, t in zip('ABCDEF', ['WBS-element', 'Project', 'Soort', 'Telt mee', 'Strapo', 'Filter\nDashboard']):
        s.set(f'{c}{R_LIST}', text=t, style=H)
    for i in range(R1 - R0 + 1):
        r = R0 + i; q = 8 + i
        s.set(f'A{r}', is_str=True, style=TXT, formula=f'IF(Projectlijst!$A{q}="","",Projectlijst!$A{q})')
        s.set(f'B{r}', is_str=True, style=TXT, formula=f'IF($A{r}="","",Projectlijst!$B{q})')
        s.set(f'C{r}', is_str=True, style=TXT, formula=f'IF($A{r}="","",IF(Projectlijst!$C{q}="Deelneming","Deelneming",IF(Projectlijst!$C{q}="Buiten scope","Buiten scope","Opstal")))')
        s.set(f'D{r}', is_str=True, style=TXT, formula=f'IF($A{r}="","",Projectlijst!$AN{q})')
        s.set(f'E{r}', is_str=True, style=TXT, formula=f'IF($A{r}="","",IF(OR(Projectlijst!$C{q}="Strapo",Projectlijst!$X{q}="Ja"),"Ja","Nee"))')
        s.set(f'F{r}', is_str=True, style=TXT, formula=f'IF(AND($D{r}="Ja",OR(Sel_TypeCrit="*",$C{r}=Sel_TypeCrit),OR(Sel_WBS="*",$A{r}=Sel_WBS)),"Ja","Nee")')
    for c in range(G0, G0 + BW * 5):
        s.set(f'{L(c)}{R_LIST}', is_str=True, style=H, formula=f'{L(c)}{R_LAB}')
    # ---- samenvatting B4:G20 (doeljaar)
    s.set('B4', is_str=True, formula='"Totalen "&Sel_Doeljaar&" (EUR)"', style=H)
    for k, v in enumerate(VERS): s.set(f'{L(3 + k)}4', text=v, style=H)
    rows = [('Omzet opstal (dump)', 'Z', R_OP, 1, True), ('Omzet deelnemingen (dump)', 'Z', R_DN, 1, False),
            ('OPR opstal (dump)', 'Z', R_OP, 2, True), ('OPR deelnemingen (dump)', 'Z', R_DN, 2, False),
            ('Correctie omzet opstal (Projecten)', 'C', R_OP, 15, False), ('Correctie omzet deelnemingen', 'C', R_DN, 15, False),
            ('Correctie OPR opstal', 'C', R_OP, 16, False), ('Correctie OPR deelnemingen', 'C', R_DN, 16, False),
            ('Omzet t/m peildatum opstal', 'Y', R_OP, 13, True), ('Omzet t/m peildatum deelnemingen', 'Y', R_DN, 13, False),
            ('OPR t/m peildatum opstal', 'Y', R_OP, 14, True), ('OPR t/m peildatum deelnemingen', 'Y', R_DN, 14, False),
            ('Omzet incl. correcties, filter Dashboard', 'F', None, 1, None), ('OPR incl. correcties, filter Dashboard', 'F', None, 2, None),
            ('Omzet t/m peildatum, filter Dashboard', 'FY', None, 13, None), ('OPR t/m peildatum, filter Dashboard', 'FY', None, 14, None)]
    for i, (lab, kind, src, pos, nol) in enumerate(rows):
        r = 5 + i; s.set(f'B{r}', text=lab, style=TXT)
        for k, v in enumerate(VERS):
            c = L(3 + k); vr = VI(k); base = BW * k
            jr = f'{base}+2*$H$5+{pos}' if kind in ('Z', 'F') else f'{base + pos}'
            ix = lambda row, p: f'INDEX({DROW(row)},{p})'.replace('Berekening!', '')
            if kind == 'Z':
                f = f'IF($P${vr}=0,"–",{ix(src, jr)}' + (f'+{ix(R_NOL, jr)}' if nol else '') + ')'
            elif kind == 'C':
                f = '0' if v not in CORR else f'IF(Sel_Doeljaar<>P_RJ,0,{ix(src, jr)})'
            elif kind == 'Y':
                f = f'IF($P${vr}=0,"–",IF($O${vr}<>Sel_Doeljaar,"n.v.t.",{ix(src, jr)}' + (f'+{ix(R_NOL, jr)}' if nol else '') + '))'
            elif kind == 'F':
                cr = base + 14 + pos
                corr = f'+IF(Sel_Doeljaar=P_RJ,{ix(R_FOP, cr)}+{ix(R_FDN, cr)},0)' if v in CORR else ''
                f = f'IF($P${vr}=0,"–",{ix(R_FOP, jr)}+{ix(R_FDN, jr)}{corr})'
            else:
                f = f'IF($P${vr}=0,"–",IF($O${vr}<>Sel_Doeljaar,"n.v.t.",{ix(R_FOP, jr)}+{ix(R_FDN, jr)}))'
            s.set(f'{c}{r}', formula=f, style=NUM)
    s.root.find(N+'dimension').set('ref', f'A1:{L(G0 + BW * 5 - 1)}{R1}')
    return s

def move_sheet(bk, name, after):
    wb = bk.text('xl/workbook.xml')
    sheets = re.findall(r'<sheet [^>]*/>', wb)
    names = [re.search(r'name="([^"]+)"', e).group(1).replace('&apos;', "'") for e in sheets]
    i = names.index(name); el = sheets.pop(i); nm = names.pop(i)
    j = names.index(after) + 1; sheets.insert(j, el); names.insert(j, nm)
    old = [re.search(r'name="([^"]+)"', e).group(1).replace('&apos;', "'") for e in re.findall(r'<sheet [^>]*/>', wb)]
    wb = re.sub(r'<sheets>.*</sheets>', '<sheets>' + ''.join(sheets) + '</sheets>', wb, flags=re.S)
    wb = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{names.index(old[int(m.group(1))])}"', wb)
    bk.settext('xl/workbook.xml', wb)

def styles(bk):
    c, r = bk.sheet('CALC_Versies'), bk.sheet('RDFO')
    return dict(title=c.style('A1'), sub=c.style('A2'), hdr=c.style('A4'), num=c.style('G5'), txt=c.style('A5'),
                lnk=r.style('E3'), tot=bk.sheet('Projecten').style('E6'), inpn=bk.sheet('Projectlijst').style('L8'),
                inpt=bk.sheet('Invoer').style('D14'))

# ------------------------------------------------------------------ rapporten aan Berekening koppelen
SUMC = lambda k: L(3 + k)                         # kolom van versie k in de samenvatting (C..G)
KSEL = 'MATCH(Sel_Versie,Berekening!$J$5:$J$9,0)-1'
KPRV = 'MATCH(Sel_Prev,Berekening!$J$5:$J$9,0)-1'

def proj_expr(wbs, k, kind):
    """bedrag van één project (k EUR) voor versie k (0..4, mag een formule zijn), doeljaar; incl. correctie."""
    ri = f'MATCH({wbs},{DETA},0)'
    base = f'{BW}*({k})'
    cnt = f'INDEX({DET},{ri},{base}+17)'
    if kind == 'ytd':
        val = f'INDEX({DET},{ri},{base}+13)'
        return f'IFERROR(IF(INDEX(Berekening!$O$5:$O$9,({k})+1)<>Sel_Doeljaar,"n.v.t.",IF({cnt}=0,"–",N({val})/P_Factor)),"–")'
    p, pc = (1, 15) if kind == 'omz' else (2, 16)
    val = f'INDEX({DET},{ri},{base}+2*Berekening!$H$5+{p})'
    cor = f'INDEX({DET},{ri},{base}+{pc})'
    return f'IFERROR(IF(AND({cnt}=0,N({cor})=0),"–",(N({val})+IF(Berekening!$H$5=0,N({cor}),0))/P_Factor),"–")'

def calc_versies(bk):
    c = bk.sheet('CALC_Versies')
    m = {'G': (5, 9), 'H': (6, 10), 'I': (7, 11), 'J': (8, 12), 'K': (13,), 'L': (14,), 'M': (15,), 'N': (16,),
         'S': (17,), 'T': (18,), 'U': (19,), 'V': (20,)}
    for k in range(5):
        r = 5 + k; X = SUMC(k)
        c.set(f'C{r}', formula=f'Berekening!$P${5 + k}', style=c.style(f'C{r}'))
        for col, rows in m.items():
            a = f'Berekening!${X}${rows[0]}'
            f = f'IF(ISNUMBER({a}),{a}+Berekening!${X}${rows[1]},{a})' if len(rows) == 2 else a
            c.set(f'{col}{r}', formula=f, style=c.style(f'{col}{r}'), is_str=False)

def rdfo(bk, sty):
    s = bk.sheet('RDFO')
    for row, (a, b) in ((7, (5, 9)), (8, (6, 10)), (15, (7, 11)), (16, (8, 12))):
        for k in range(5):
            X = SUMC(k); A = f'Berekening!${X}${a}'
            s.set(f'{L(3 + k)}{row}', style=s.style(f'{L(3 + k)}{row}'),
                  formula=f'IF(ISNUMBER({A}),({A}+Berekening!${X}${b})/P_Factor,{A})')
        s.set(f'R{row}', text='Berekening →', style=sty['lnk']); add_hyperlink(s, f'R{row}', f"'Berekening'!B{a}", 'Berekening')
    s.set('S7', text='Berekening: Bedrijfsopbrengsten LJ (X) + OP.BOB van het jaar; budget = alleen OP.BOB. Incl. correcties van Projecten.', style=s.style('S7'))
    s.set('S8', text='Idem, soort Deelneming (Projectlijst).', style=s.style('S8'))
    s.set('S15', text='Berekening: Bruto (omzet) result. LJ (Z) + OP.BOR van het jaar. Incl. correcties van Projecten.', style=s.style('S15'))
    s.set('S16', text='Idem, soort Deelneming.', style=s.style('S16'))
    s.set('E3', text='Berekening →', style=s.style('E3'))
    for h in s.root.iter(N+'hyperlink'):
        if h.get('ref') == 'E3': h.set('location', "'Berekening'!A1"); h.set('display', 'Berekening')
    H, LAB, NUM = s.style('B33'), s.style('B7'), s.style('C7')
    s.set('B46', text='Handmatige correcties (tabblad Projecten) — zitten al in de cijfers hierboven', style=H)
    for c in 'CDEFG': s.set(f'{c}46', style=H)
    for row, (a, b), lab in ((47, (9, 10), 'Correctie omzet'), (48, (11, 12), 'Correctie OPR')):
        s.set(f'B{row}', text=lab, style=LAB)
        for k in range(5):
            X = SUMC(k); s.set(f'{L(3 + k)}{row}', style=NUM, formula=f'(N(Berekening!${X}${a})+N(Berekening!${X}${b}))/P_Factor')
        s.set(f'R{row}', text='Projecten →', style=sty['lnk']); add_hyperlink(s, f'R{row}', "'Projecten'!AC9", 'Projecten')

def projecten(bk, sty, dxf_sel):
    s = bk.sheet('Projecten')
    for r in range(10, 110):
        for k, c in enumerate('EFGHI'):
            s.set(f'{c}{r}', style=s.style(f'{c}{r}'), formula=f'IF($A{r}="","",{proj_expr(f"$A{r}", k, "omz")})')
        s.set(f'L{r}', style=s.style(f'L{r}'), formula=f'IF($A{r}="","",{proj_expr(f"$A{r}", KSEL, "ytd")})')
        s.set(f'N{r}', style=s.style(f'N{r}'), formula=f'IF($A{r}="","",{proj_expr(f"$A{r}", KSEL, "opr")})')
        s.set(f'AB{r}', style=s.style(f'AB{r}'), formula=f'IF(OR($A{r}="",Sel_Prev=""),"–",{proj_expr(f"$A{r}", KPRV, "omz")})')
        s.set(f'X{r}', is_str=True, style=s.style(f'X{r}'),
              formula=f'IF($A{r}="","",IF(COUNTIF(E{r}:I{r},">0.5")+COUNTIF(E{r}:I{r},"<-0.5")=0,"leeg","met cijfers"))')
        for c in 'ACDEFGHIJ'[0:0]: pass
        for v, (co, cp) in CORR.items():
            for c in (co, cp):
                el = s.cell(f'{c}{r}'); keep = [ch for ch in el if ch.tag == N+'v']
                if el.find(N+'f') is not None or el.find(N+'is') is not None: s.set(f'{c}{r}', style=sty['inpn'])
                else: el.set('s', sty['inpn'])
        el = s.cell(f'{CORR_TXT}{r}')
        if el.find(N+'f') is not None: s.set(f'{CORR_TXT}{r}', style=sty['inpt'])
        else: el.set('s', sty['inpt'])
    H, TOT = s.style('E9'), s.style('E6')
    s.set('AC7', text='HANDMATIGE CORRECTIE per versie (EUR x 1.000, + of −) — telt mee in RDFO, Dashboard en Outlook', style=sty['hdr'])
    for v, (co, cp) in CORR.items():
        for c, kg in ((co, 'omzet'), (cp, 'OPR')):
            s.set(f'{c}8', text=v)
            s.set(f'{c}9', is_str=True, style=H, formula=f'"{v} "&P_VJ&CHAR(10)&"correctie {kg}"')
            s.set(f'{c}6', style=TOT, formula=f'SUBTOTAL(109,{c}10:{c}109)')
    s.set(f'{CORR_TXT}9', text='Toelichting correctie (typ hier)', style=H)
    add_cf(s, 'AC10:AJ109', 'AC$8=Sel_Versie', dxf_sel, prio=4)
    cols = s.root.find(N+'cols')
    def addcol(mn, mx, w, before=None):
        e = etree.Element(N+'col'); e.set('min', str(mn)); e.set('max', str(mx)); e.set('width', str(w)); e.set('customWidth', '1')
        kids = list(cols); pos = len(kids)
        for i, c in enumerate(kids):
            if int(c.get('min')) > mn: pos = i; break
        cols.insert(pos, e)
    if not any(c.get('min') == '29' for c in cols):
        addcol(25, 25, 2); addcol(29, 36, 10); addcol(37, 37, 42)
    s.root.find(N+'dimension').set('ref', 'A1:AK109')
    s.set('B7', style=s.style('B7'), text="Geel = kwartaal gekozen op het Dashboard. Mutatie: groen = hoger, rood = lager. Projecten zonder cijfers zijn verborgen "
          "(knop rechtsboven: alles tonen). Zelf corrigeren: kolom AC–AK (per versie). Niet sorteren.")

def projectdetail(bk, sty):
    p = bk.sheet('Projectdetail')
    for k in range(5):
        r = 17 + k
        p.set(f'C{r}', style=p.style(f'C{r}'), formula=proj_expr('Det_WBS', k, 'omz'))
        p.set(f'D{r}', style=p.style(f'D{r}'), formula=proj_expr('Det_WBS', k, 'opr'))
        p.set(f'F{r}', style=p.style(f'F{r}'), formula=proj_expr('Det_WBS', k, 'ytd'))
    p.set('E3', text='Berekening →', style=p.style('E3'))
    for h in p.root.iter(N+'hyperlink'):
        if h.get('ref') == 'E3': h.set('location', "'Berekening'!A1"); h.set('display', 'Berekening')

def outlook(bk):
    o = bk.sheet('Outlook'); KU = f'{BW}*(P_UploadIdx-1)'
    R = lambda row: DROW(row)
    for y in range(NY):
        c = L(4 + y)          # D..I
        corr_o = f'+INDEX({R(R_FOP)},{KU}+15)' if y == 0 else ''
        corr_d = f'+INDEX({R(R_FDN)},{KU}+15)' if y == 0 else ''
        o.set(f'{c}6', style=o.style(f'{c}6'), formula=f'(INDEX({R(R_FOP)},{KU}+{2*y+1}){corr_o})/P_Factor')
        o.set(f'{c}8', style=o.style(f'{c}8'), formula=f'(INDEX({R(R_FDN)},{KU}+{2*y+1}){corr_d})/P_Factor')
        o.set(f'{c}7', style=o.style(f'{c}7'), formula=(f'SUMIFS(INDEX({DET},0,{KU}+{2*y+1}),Berekening!$C${R0}:$C${R1},"Opstal",'
                                                        f'Berekening!$E${R0}:$E${R1},"Ja",Berekening!$F${R0}:$F${R1},"Ja")/P_Factor'))
        corr = f'+INDEX({R(R_FOP)},{KU}+16)+INDEX({R(R_FDN)},{KU}+16)' if y == 0 else ''
        o.set(f'{c}16', style=o.style(f'{c}16'), formula=f'(INDEX({R(R_FOP)},{KU}+{2*y+2})+INDEX({R(R_FDN)},{KU}+{2*y+2}){corr})/P_Factor')
        o.set(f'{c}21', style=o.style(f'{c}21'), formula=f'IF(Berekening!$P$5=0,"–",(INDEX({R(R_FOP)},{2*y+1})+INDEX({R(R_FDN)},{2*y+1}))/P_Factor)')
    o.set('D10', style=o.style('D10'), formula=f'(INDEX({R(R_FOP)},{KU}+13)+INDEX({R(R_FDN)},{KU}+13))/P_Factor')
    o.set('E3', text='Berekening →', style=o.style('E3'))
    for h in o.root.iter(N+'hyperlink'):
        if h.get('ref') == 'E3': h.set('location', "'Berekening'!A1"); h.set('display', 'Berekening')

def bandbreedte_budget(bk):
    b = bk.sheet('Bandbreedte')
    for y, c in enumerate('BCD'):
        b.set(f'{c}12', style=b.style(f'{c}12'), formula=f'IF(Berekening!$P$5=0,"–",INDEX({DROW(R_TOT)},{2*y+1})/1000000)')

def main(src, dst, herberekend=None, reserves=None):
    from patch_v7b import bron_reserves, projectlijst_reserves, bandbreedte_grafiek, start, verberg_rekenbladen, projecten_zichtbaar, outlook_grafiek, outlook_jaren_vijf
    bk = Book(src); sty = styles(bk)
    berekening(bk, sty); move_sheet(bk, 'Berekening', 'Bron_Mijlpalen')
    sel = add_dxf(bk, 'FFFFD966')
    calc_versies(bk); rdfo(bk, sty); projecten(bk, sty, sel); projectdetail(bk, sty); outlook(bk); bandbreedte_budget(bk)
    oranje = add_dxf(bk, 'FFFCD9A0')
    bron_reserves(bk, sty, reserves); move_sheet(bk, 'Bron_Reserves', 'Bron_Mijlpalen')
    projectlijst_reserves(bk, sty, oranje); bandbreedte_grafiek(bk); outlook_grafiek(bk); outlook_jaren_vijf(bk); start(bk); verberg_rekenbladen(bk)
    projecten_zichtbaar(bk, herberekend)
    from patch_v7c import reserves_naam, start_tegels
    reserves_naam(bk, sty); start_tegels(bk)
    pp = bk.sheet('PPT')
    sst = [''.join(si.itertext()) for si in etree.fromstring(bk.data['xl/sharedStrings.xml'])]
    for r in range(40, 80):
        el = pp.cell(f'Q{r}'); v = el.find(N+'v')
        txt = sst[int(v.text)] if (el.get('t') == 's' and v is not None) else ''.join(el.itertext())
        if txt.startswith('Rood = downside'):
            pp.set(f'Q{r}', text='Per jaar: ondergrens (rood), basis (navy) en bovengrens (groen), EUR mln', style=pp.style(f'Q{r}'))
    c = bk.sheet('Controle')
    for r in range(60, 100):
        f = c.cell(f'E{r}').find(N+'f')
        if f is not None and f.text == 'IFERROR(Projecten!$K$6*P_Factor,"–")':
            f.text = 'IFERROR(SUM(Projecten!$K$10:$K$109)*P_Factor,"–")'
            c.set(f'H{r}', text='Verschil = project dat niet op de Projectlijst staat (Berekening, rij 29).', style=c.style(f'H{r}'))
    bk.save(dst)

if __name__ == '__main__': main(*sys.argv[1:])
