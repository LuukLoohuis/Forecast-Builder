"""Ronde 7, deel 2: reserves-dump, Projectlijst-reserves, bandbreedte-grafiek staand, Start-knoppen, rekenbladen verbergen."""
import sys, os, re
sys.path.insert(0, os.path.dirname(__file__))
from patch_ak import split, col2n, n2col, N
from patch_eenvoud import add_cf, add_hyperlink
from lxml import etree
L = n2col


def bron_reserves(bk, sty, bron=None):
    if 'Bron_Reserves' not in bk.paths:
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><tabColor rgb="FFFEB900"/></sheetPr><dimension ref="A1"/>'
               '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="14.4"/>'
               '<cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="5" width="18" customWidth="1"/></cols><sheetData/>'
               '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode()
        bk.add_sheet('Bron_Reserves', xml)
    s = bk.sheet('Bron_Reserves')
    if bron:
        import openpyxl
        ws = openpyxl.load_workbook(bron, data_only=True).worksheets[0]
        for row in ws.iter_rows():
            for c in row:
                if c.value is None: continue
                ref = f'{L(c.column)}{c.row}'
                if isinstance(c.value, (int, float)):
                    el = s.cell(ref); el.attrib.pop('t', None)
                    for ch in list(el): el.remove(ch)
                    etree.SubElement(el, N+'v').text = repr(float(c.value)) if isinstance(c.value, float) else str(c.value)
                else:
                    s.set(ref, text=str(c.value))
        s.root.find(N+'dimension').set('ref', f'A1:{L(ws.max_column)}{ws.max_row}')
    else:
        s.set('A1', text='Plak hier de reserves-export (hele blad, vanaf A1): per WBS de kolommen Onvoorzien, Commerciële reserve en Nazorg.')
    bk.set_name('Z_Res', 'Bron_Reserves!$A$1:$Z$3000')
    b = bk.sheet('Berekening'); H, NUM, TXT, LNK = sty['hdr'], sty['num'], sty['txt'], sty['lnk']
    b.set('S4', text='Reserves-dump →', style=LNK); add_hyperlink(b, 'S4', "'Bron_Reserves'!A1", 'Bron_Reserves')
    b.set('T4', text='Kolom (nr)', style=H); b.set('U4', text='Kolom', style=H)
    kop = '0'
    for rr in (5, 4, 3, 2, 1):
        kop = f'IF(COUNTIF(Bron_Reserves!${rr}:${rr},"*Onvoorzien*")>0,{rr},{kop})'
    rows = [(5, 'Koprij', kop),
            (6, 'WBS', 'IF(T5=0,0,IFERROR(MATCH("*WBS*",INDEX(Z_Res,T5,0),0),IFERROR(MATCH("*WBS*",INDEX(Z_Res,T5+1,0),0),1)))'),
            (7, 'Commerciële reserve', 'IF(T5=0,0,IFERROR(MATCH("*Commerci*",INDEX(Z_Res,T5,0),0),0))'),
            (8, 'Onvoorzien', 'IF(T5=0,0,IFERROR(MATCH("*Onvoorzien*",INDEX(Z_Res,T5,0),0),0))'),
            (9, 'Nazorg', 'IF(T5=0,0,IFERROR(MATCH("*Nazorg*",INDEX(Z_Res,T5,0),0),0))'),
            (10, 'Regels met WBS', 'IF(T6=0,0,COUNTIF(INDEX(Z_Res,0,T6),"?.*"))')]
    for r, lab, f in rows:
        b.set(f'S{r}', text=lab, style=TXT); b.set(f'T{r}', formula=f, style=NUM)
        if 6 <= r <= 9:
            b.set(f'U{r}', is_str=True, style=TXT, formula=f'IF(T{r}=0,"–",SUBSTITUTE(ADDRESS(1,T{r},4),"1",""))')


def projectlijst_reserves(bk, sty, dxf_or):
    p = bk.sheet('Projectlijst'); H6, H7, NUMr = p.style('AG6'), p.style('AG7'), p.style('S8')
    H7r = p.style('L7')
    p.set('L7', text='Commerciële\nreserve', style=H7r); p.set('M7', text='Onvoorzien', style=H7r); p.set('N7', text='Nazorg', style=H7r)
    p.set('Q7', text='Onvoorzien\nvalt vrij in', style=H7r)
    p.set('AU6', text='RESERVES: EIGEN BEDRAG (leeg = dump)', style=H6); p.set('AX6', text='RESERVES UIT DUMP (Bron_Reserves)', style=H6)
    for c, t in zip(['AU', 'AV', 'AW', 'AX', 'AY', 'AZ'],
                    ['Commerciële\neigen', 'Onvoorzien\neigen', 'Nazorg\neigen', 'Commerciële\nuit dump', 'Onvoorzien\nuit dump', 'Nazorg\nuit dump']):
        p.set(f'{c}7', text=t, style=H7)
    for r in range(8, 108):
        for amt, eig, dmp, kol in (('L', 'AU', 'AX', '$T$7'), ('M', 'AV', 'AY', '$T$8'), ('N', 'AW', 'AZ', '$T$9')):
            el = p.cell(f'{amt}{r}'); v = el.find(N+'v')
            old = v.text if (v is not None and el.find(N+'f') is None) else None
            p.set(f'{amt}{r}', style=NUMr, formula=f'IF($A{r}="","",IF({eig}{r}<>"",{eig}{r},{dmp}{r}))')
            e2 = p.cell(f'{eig}{r}')
            if old not in (None, '') and e2.find(N+'v') is None:      # eerder handmatig ingevuld bedrag blijft als eigen bedrag
                p.set(f'{eig}{r}', style=sty['inpn']); etree.SubElement(p.cell(f'{eig}{r}'), N+'v').text = old
            else:
                e2.set('s', sty['inpn'])
            p.set(f'{dmp}{r}', style=NUMr,
                  formula=(f'IF(OR($A{r}="",Berekening!{kol}=0),"",SUMIFS(INDEX(Z_Res,0,Berekening!{kol}),INDEX(Z_Res,0,Berekening!$T$6),$A{r})'
                           f'+SUMIFS(INDEX(Z_Res,0,Berekening!{kol}),INDEX(Z_Res,0,Berekening!$T$6),$A{r}&".*"))'))
    add_cf(p, 'L8:N107', 'AND($A8<>"",AU8<>"")', dxf_or, prio=2)
    bk.set_name('RS_Invoer', 'Projectlijst!$P$8:$R$107'); bk.set_name('RS_Eigen', 'Projectlijst!$AU$8:$AW$107')
    # kwartaal vrij typen met controle op jjjj-QX
    for dv in p.root.iter(N+'dataValidation'):
        if dv.get('sqref') == 'P8:R107':
            dv.set('type', 'custom'); dv.set('allowBlank', '1'); dv.set('showErrorMessage', '1'); dv.set('showInputMessage', '1')
            dv.set('errorTitle', 'Kwartaal'); dv.set('error', 'Vul het kwartaal in als jjjj-QX, bijvoorbeeld 2026-Q4 (of Later).')
            dv.set('promptTitle', 'Valt vrij in'); dv.set('prompt', 'Typ jjjj-QX, bijv. 2026-Q4, of Later.')
            dv.find(N+'formula1').text = ('OR(P8="",UPPER(P8)="LATER",AND(LEN(P8)=7,ISNUMBER(-LEFT(P8,4)),UPPER(MID(P8,5,2))="-Q",'
                                          'ISNUMBER(FIND(RIGHT(P8,1),"1234"))))')
    p.set('AK7', text='Kans % eigen\n(leeg = ZPMW03)', style=p.style('AK7'))
    cols = p.root.find(N+'cols')
    if not any(c.get('min') == '47' for c in cols):
        e = etree.SubElement(cols, N+'col'); e.set('min', '47'); e.set('max', '52'); e.set('width', '11'); e.set('customWidth', '1')
    p.set('B2', style=p.style('B2'), text='Gevuld vanuit ZPMW03 (nieuwe projecten komen er bij het openen vanzelf bij, oranje). Kies de soort: Opstal, Deelneming, '
          'Strapo of Buiten scope. Reserves komen uit Bron_Reserves; een eigen bedrag in AU–AW gaat voor (oranje).')
    p.root.find(N+'dimension').set('ref', 'A1:AZ107')
    # Onvoorzien i.p.v. Operationele reserve (alleen eigen labels; geen exportkoppen)
    sst = bk.text('xl/sharedStrings.xml')
    for a, b2 in (('Operationele\nreserve', 'Onvoorzien'), ('Operationele reserve', 'Onvoorzien'), ('operationele reserve', 'onvoorzien')):
        sst = sst.replace(a, b2)
    bk.settext('xl/sharedStrings.xml', sst)
    for nm in ('Reserves', 'RDFO', 'Projectlijst', 'Dias'):
        sh = bk.sheet(nm)
        for row in sh.rows.values():
            for el in row.findall(N+'c'):
                for t in el.iter(N+'t'):
                    if t.text and 'perationele' in t.text:
                        t.text = (t.text.replace('Operationele\nreserve', 'Onvoorzien').replace('Operationele reserve', 'Onvoorzien')
                                  .replace('Operationele\n', 'Onvoorzien\n'))
    rs = bk.sheet('Reserves')
    rs.set('B2', style=rs.style('B2'), text='Bedragen uit Bron_Reserves (eigen bedrag op de Projectlijst AU–AW gaat voor); vrijvalkwartaal per soort op de Projectlijst (P–R).')


def bandbreedte_grafiek(bk):
    import chart_bb
    b = bk.sheet('Bandbreedte')
    for c, t in zip(['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'],
                    ['jaar', 'onder', 'Downside (ondergrens tot basis)', 'Upside (basis tot bovengrens)', 'Ondergrens', 'Basis', 'Bovengrens', 'Basis (prognose)']):
        b.set(f'{c}14', text=t)
    for k, c in enumerate('BCD'):
        r = 15 + k
        b.set(f'Y{r}', formula=f'{c}5', is_str=True)
        b.set(f'AC{r}', formula=f'IF(ISNUMBER({c}7),MIN({c}6:{c}10),NA())')
        b.set(f'AD{r}', formula=f'IF(ISNUMBER({c}7),{c}7,NA())')
        b.set(f'AE{r}', formula=f'IF(ISNUMBER({c}7),MAX({c}6:{c}10),NA())')
        b.set(f'AF{r}', formula='IFERROR(MAX($AE$15:$AE$17)*0.012,0)')
        b.set(f'Z{r}', formula=f'IFERROR(AC{r},0)')
        b.set(f'AA{r}', formula=f'IFERROR(MAX(0,AD{r}-AC{r}-AF{r}/2),0)')
        b.set(f'AB{r}', formula=f'IFERROR(MAX(0,AE{r}-AD{r}-AF{r}/2),0)')
    b.set('J4', text='Bandbreedte omzet per jaar (EUR mln): ondergrens, basis en bovengrens', style=b.style('J4'))
    xml = chart_bb.chart_xml().encode()

    def chart_of(sheet, nr):
        p = bk.paths[sheet]; rels = bk.text(p.replace('worksheets/', 'worksheets/_rels/') + '.rels')
        drw = 'xl/drawings/' + re.search(r'Target="\.\./drawings/([^"]+)"', rels).group(1)
        dr = bk.text(drw); drels = bk.text(drw.replace('drawings/', 'drawings/_rels/') + '.rels')
        rid = re.findall(r'<xdr:graphicFrame.*?r:id="([^"]+)"', dr, re.S)[nr - 1]
        tg = (re.search(r'Id="%s"[^>]*Target="\.\./charts/([^"]+)"' % rid, drels)
              or re.search(r'Target="\.\./charts/([^"]+)"[^>]*Id="%s"' % rid, drels))
        return 'xl/charts/' + tg.group(1), drw
    for sheet, nr in (('Bandbreedte', 1), ('Dias', 9)):
        fn, drw = chart_of(sheet, nr); bk.data[fn] = xml
        if sheet == 'Bandbreedte':
            d = bk.text(drw).replace('<xdr:ext cx="5580000" cy="2376000"/>', '<xdr:ext cx="5580000" cy="3060000"/>'); bk.settext(drw, d)


def start(bk):
    st = bk.sheet('Start')
    rfn = bk.paths['Start'].replace('worksheets/', 'worksheets/_rels/') + '.rels'
    fn = 'xl/drawings/' + re.search(r'Target="\.\./drawings/([^"]+)"', bk.text(rfn)).group(1)
    d = bk.text(fn)
    if 'RDFO_ToonDumps' not in d:
        tmpl = next(a for a in re.findall(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', d, re.S) if 'RDFO_ToonRekenbladen' in a)
        extra = ''
        for i, (mac, txt) in enumerate((('RDFO_ToonDumps', 'Dumps tonen / verbergen'), ('RDFO_NaarDump', 'Naar ZPMW03-dump'))):
            k = tmpl.replace('RDFO_ToonRekenbladen', mac)
            k = re.sub(r'(<xdr:(?:from|to)><xdr:col>)(\d+)', lambda m, i=i: m.group(1) + str(int(m.group(2)) + 1 + i), k)
            k = re.sub(r'descr="[^"]*"', f'descr="{txt}"', k); k = re.sub(r'name="Knop \d+"', f'name="Knop {1102 + i}"', k)
            k = re.sub(r'<xdr:cNvPr id="\d+"', f'<xdr:cNvPr id="{1102 + i}"', k)
            k = re.sub(r'(<a:t>)[^<]*(</a:t>)', lambda m, txt=txt: m.group(1) + txt + m.group(2), k, count=1)
            extra += k
        d = d.replace(tmpl, tmpl + extra); bk.settext(fn, d)
    T = st.style('B12')
    st.set('D12', text='Verbergt of toont de gele plaktabbladen.', style=T)
    st.set('E12', is_str=True, style=T, formula='"Gaat naar het tabblad "&P_UploadVersie&" ZPMW03 (Invoer › Versie van de upload)."')
    st.set('C15', style=st.style('C15'), text="Gele tabbladen: ZPMW03 in het tabblad van het kwartaal (Budget, Q1–Q4), Woningetalage, AK-percentages, AK-realisatie, "
           "Mijlpalen en Reserves. Knop 'Naar ZPMW03-dump' brengt je erheen.")
    st.set('C18', style=st.style('C18'), text='Dashboard: kies kwartaal en deelgebied. Omzet en OPR staan op het grijze tabblad Berekening; elk rapport linkt ernaar.')
    st.set('C19', style=st.style('C19'), text="Knop 'Maak PowerPoint'. Correcties per project: tabblad Projecten (kolom AC–AK).")
    n = 'N(Berekening!$T$10)'
    st.set('G17', is_str=True, style=st.style('G16'), formula=f'IF({n}>0,"✔  ","–  ")&"Reserves"&IF({n}>0," ("&{n}&IF({n}=1," regel)"," regels)"),"")')
    add_hyperlink(st, 'G17', "'Bron_Reserves'!A1", 'Reserves')
    g7 = st.cell('G7').find(N+'f')
    if g7 is not None and 'Berekening!$T$10' not in g7.text:
        g7.text = g7.text.replace(')&" van 9"', f'+IF({n}>0,1,0))&" van 10"')


def verberg_rekenbladen(bk):
    wb = bk.text('xl/workbook.xml')
    for nm in ('VERSIE_REKEN', 'CALC_Versies', 'VERW_Project', 'SCEN', 'STG_ZPMW03', 'STG_Woningetalage', 'STG_AKR'):
        wb = re.sub(r'(<sheet name="%s" sheetId="\d+")( state="\w+")?' % re.escape(nm), r'\1 state="hidden"', wb)
    bk.settext('xl/workbook.xml', wb)


def projecten_zichtbaar(bk, herberekend):
    s = bk.sheet('Projecten'); toon = {}
    if herberekend:
        import openpyxl
        v = openpyxl.load_workbook(herberekend, data_only=True, read_only=True)
        for i, (x,) in enumerate(v['Projecten'].iter_rows(min_row=10, max_row=109, min_col=24, max_col=24, values_only=True)):
            toon[10 + i] = (x == 'met cijfers')
    for r in range(10, 110):
        if toon.get(r, False): s.row(r).attrib.pop('hidden', None)
        else: s.row(r).set('hidden', '1')
