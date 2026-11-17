"""Voegt tabbladen 'Oude versies' en 'Bron_Versie' toe: oude RDFO-tabbladen (Budget / Prognose Qx)
plakken en met formules omzetten naar Historie-regels (A:R) om in één keer te plakken.
usage: python3 -I patch_oude_versies.py in.xlsm out.xlsm"""
import sys, os, zipfile, re
sys.path.insert(0, os.path.dirname(__file__))
from patch_ak import Sheet, sheet_paths, N
from lxml import etree

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SKEL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{{pr}}<dimension ref="A1"/>'
        '<sheetViews><sheetView workbookViewId="0"{frz}</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="14.4"/>{cols}<sheetData/>{dv}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')

def skeleton(cols='', dv='', tab=None, freeze=None):
    pr = f'<sheetPr><tabColor rgb="{tab}"/></sheetPr>' if tab else ''
    frz = (f'><pane ySplit="{freeze-1}" topLeftCell="A{freeze}" activePane="bottomLeft" state="frozen"/>'
           if freeze else '>')
    return SKEL.replace('{cols}', cols).replace('{dv}', dv).replace('{pr}', pr).replace('{frz}', frz).encode()

ROWS_SRC = 300            # regels in het plakblad
OUT0 = 22                 # eerste outputregel
NOUT = 900                # 300 projecten x 3 kengetallen
HIST_HDR = ['ImportID','Bronbestand','Importdatum','Status','Versiejaar','Versie','Revisie','Peildatum','Doeljaar',
            'WBS','Projectdefinitie','Projectnaam','Type','IFRS','Kengetal','Waarde','Eenheid','Toelichting']

def build_oude(s, st):
    T, TXT, LBL, INP, INPD, INPN, HDR, CELL, NUM, DATE, HELP = (st[k] for k in
        ['title','text','label','inp_txt','inp_date','inp_num','hdr','cell','num','date','help'])
    s.set('B1', text='Oude versies overzetten (Budget / Q1 / Q2 / Q3 / Q4)', style=T)
    steps = [
        "1. Kopieer in je oude RDFO het tabblad 'Budget' of 'Prognose Qx' (alles, Ctrl+A) en plak het als waarden in A1 van tabblad Bron_Versie.",
        "2. Vul hieronder de versie, het versiejaar en zo nodig het doeljaar in. Controleer de aansluiting rechts (verschil = 0).",
        "3. Ga naar Historie, rij 'Plak in Historie op rij' hieronder: selecteer A22:R" + str(OUT0+NOUT-1) + " hier, kopieer, en plak als WAARDEN in kolom A van die rij (lege regels onderaan zijn niet erg).",
        "4. Staat deze versie al in de Historie (bijv. opnieuw uit ZPMW03 berekend)? Dan krijgen deze regels revisie +1 en vervangen ze die cijfers. Officiële AK-totaal van de regio: zie regel 18 (naar Invoer › 3).",
    ]
    for i, t in enumerate(steps): s.set(f'B{2+i}', text=t, style=TXT)
    lab = [('B7','Versie','C7',None,INP,"Budget / Q1 / Q2 / Q3 / Q4"),
           ('B8','Versiejaar','C8',None,INPN,"Jaar van de prognoseronde, bijv. 2026"),
           ('B9','Doeljaar (leeg = jaar uit kop)','C9',None,INPN,None),
           ('B10','Peildatum (optioneel)','C10',None,INPD,None),
           ('B11','Oud bestand / tabblad','C11',None,INP,"Komt in de kolom Bronbestand"),
           ('B12','Bedragen in plakblad','C12','EUR x 1.000',INP,"EUR / EUR x 1.000 / EUR mln"),
           ('B13','AK-dekking regio officieel (EUR)','C13',None,INPN,"Optioneel; maakt regel 18 voor Invoer › 3")]
    for lr, lt, ir, iv, ist, note in lab:
        s.set(lr, text=lt, style=LBL)
        if iv is not None: s.set(ir, text=iv, style=ist)
        else: s.set(ir, style=ist)
        if note: s.set('D'+ir[1:], text=note, style=HELP)
    s.set('D9', is_str=True, style=HELP, formula='IF($AE$10="","","Gevonden in kop: "&$AE$10)')
    info = [('F7','Kolom BOB totaal','G7','IF($AG$7=0,"niet gevonden",$AG$7)'),
            ('F8','Kolom BOR totaal','G8','IF($AG$8=0,"niet gevonden",$AG$8)'),
            ('F9','Kolom AK totaal','G9','IF($AG$9=0,"niet gevonden",$AG$9)'),
            ('F10','Projecten gevonden','G10','N(MAX($AA$22:$AA$321))'),
            ('F11','AK per project meenemen','G11','IF(AND($AG$9>0,SUMPRODUCT(ABS(N(+INDEX(Bron_Versie!$A$1:$AZ$300,0,MAX(1,$AG$9)))*(ISNUMBER(FIND(".",Bron_Versie!$A$1:$A$300)))))>0),"Ja","Nee (AK = 0)")'),
            ('F12','Regels om te plakken','G12',f'COUNTIF($T${OUT0}:$T${OUT0+NOUT-1},">0")'),
            ('F13','Plak in Historie op rij','G13','Historie!$D$3')]
    for lr, lt, vr, f in info:
        s.set(lr, text=lt, style=LBL); s.set(vr, formula=f, style=NUM, is_str=False)
    # controle
    s.set('I7', text='Controle', style=HDR); s.set('J7', text="Plakblad 'Totalen'", style=HDR)
    s.set('K7', text='Naar Historie', style=HDR); s.set('L7', text='Verschil', style=HDR)
    for r, (lbl, ag) in zip((8, 9, 10), (('Omzet FY','$AG$7'),('BOR FY','$AG$8'),('AK-dekking FY','$AG$9'))):
        s.set(f'I{r}', text=lbl, style=LBL)
        s.set(f'J{r}', style=NUM, formula=f'IF({ag}=0,"n.b.",IFERROR(N(INDEX(Bron_Versie!$A$1:$AZ$300,MATCH("Totalen",Bron_Versie!$C$1:$C$300,0),{ag}))*$AE$12,"n.b."))')
        s.set(f'K{r}', style=NUM, formula=f'SUMIFS($P${OUT0}:$P${OUT0+NOUT-1},$O${OUT0}:$O${OUT0+NOUT-1},I{r})')
        s.set(f'L{r}', style=NUM, formula=f'IF(ISNUMBER(J{r}),K{r}-J{r},"")')
    s.set('I12', text='Al in Historie', style=LBL)
    s.set('J12', is_str=True, style=TXT, formula='IF(OR($C$7="",$C$8=""),"Vul versie en versiejaar in.",IF(COUNTIFS(L_Versie,$C$7,L_Versiejaar,$C$8,L_Actief,1)=0,"Nieuw: deze versie staat nog niet in de Historie.","Bestaat al ("&COUNTIFS(L_Versie,$C$7,L_Versiejaar,$C$8,L_Actief,1)&" actieve regels): Omzet, BOR en AK krijgen revisie +1 en vervangen die."))')
    s.set('I13', text='Let op', style=LBL)
    s.set('J13', is_str=True, style=TXT, formula='IF(OR($AG$7=0,$AG$8=0),"Kolommen BOB/BOR niet gevonden: plak het oude tabblad in A1 van Bron_Versie.",IF($G$10=0,"Geen WBS-elementen gevonden in kolom A van Bron_Versie.",""))')
    # hulpcellen
    for r, pat in ((7, 'BOB*'), (8, 'BOR*'), (9, 'AK*')):
        rows = ''.join(f'IF(ISNUMBER(MATCH("{pat}",Bron_Versie!${k}:${k},0)),{k},' for k in range(1, 7)) + '0' + ')' * 6
        s.set(f'AE{r}', formula=rows)
        s.set(f'AF{r}', formula=f'IF($AE{r}=0,0,MATCH("{pat}",INDEX(Bron_Versie!$A$1:$AZ$6,$AE{r},0),0))')
        s.set(f'AG{r}', formula=f'IF($AF{r}=0,0,$AF{r}+4)')
    s.set('AE10', is_str=True, formula='IF($AE$7=0,"",IFERROR(RIGHT(TRIM(INDEX(Bron_Versie!$A$1:$AZ$6,$AE$7,$AF$7)),4),""))')
    s.set('AE11', formula='IF($C$9<>"",$C$9,IFERROR(VALUE($AE$10),""))')
    s.set('AE12', formula='IF($C$12="EUR mln",1000000,IF($C$12="EUR",1,1000))')
    s.set('AE13', is_str=True, formula='"OUD-"&$C$7&"-"&$C$8&"-"&TEXT(TODAY(),"yyyymmdd")')
    s.set('AE14', formula='IF(LEFT($G$11,2)="Ja",3,2)')
    s.set('AD6', text='hulp (niet wijzigen)')
    # regel voor Invoer › 3
    s.set('B16', text="Officiële AK-dekking regio (optioneel): kopieer A18:L18 en plak als waarden in Invoer › 3 (eerste lege regel vanaf rij 136).", style=TXT)
    for i, h in enumerate(['Doeljaar','Versiejaar','Versie','Niveau','WBS-element','Rubriek','Soort','Waarde (EUR)','Eenheid','Status','Toelichting','Bron / eigenaar']):
        s.set(f'{chr(65+i)}17', text=h, style=HDR)
    g = 'IF($C$13="","",'
    for col, f in zip('ABCDEFGHIJKL', ['$AE$11','$C$8','$C$7','"Regio"','""','"AK-dekking"','"FY"','$C$13','"EUR"','"Vastgesteld"',
                                         '"Officieel AK-totaal "&$C$7&" "&$C$8&" uit oude RDFO"','$C$11']):
        s.set(f'{col}18', formula=g + f + ')', style=CELL, is_str=col not in 'ABH')
    # output-kop
    for i, h in enumerate(HIST_HDR): s.set(f'{chr(65+i)}{OUT0-1}', text=h, style=HDR)
    s.set(f'T{OUT0-1}', text='bron', style=HDR); s.set(f'U{OUT0-1}', text='kengetal', style=HDR)
    s.set(f'AA{OUT0-1}', text='nr'); s.set(f'AB{OUT0-1}', text='WBS plakblad')
    for i in range(1, ROWS_SRC + 1):
        r = OUT0 - 1 + i
        s.set(f'AB{r}', formula=f'TRIM(Bron_Versie!$A{i}&"")', is_str=True)
        s.set(f'AA{r}', formula=f'IF(AND(LEN(AB{r})>=8,ISNUMBER(FIND(".",AB{r})),COUNTIF($AB${OUT0}:AB{r},AB{r})=1),MAX($AA${OUT0-1}:AA{r-1})+1,"")')
    src = f'MATCH($J{{r}},Bron_Versie!$A$1:$A${ROWS_SRC},0)'
    for k in range(1, NOUT + 1):
        r = OUT0 - 1 + k
        m = src.format(r=r)
        s.set(f'T{r}', formula=f'IF($G$10=0,0,IFERROR(MATCH(INT(({k}-1)/$AE$14)+1,$AA${OUT0}:$AA${OUT0+ROWS_SRC-1},0),0))')
        s.set(f'U{r}', is_str=True, formula=f'IF($T{r}=0,"",CHOOSE(MOD({k}-1,$AE$14)+1,"Omzet FY","BOR FY","AK-dekking FY"))')
        z = f'IF($T{r}=0,"",'
        F = {
         'A': '$AE$13', 'B': '$C$11', 'C': 'TODAY()', 'D': '"Vastgesteld"', 'E': '$C$8', 'F': '$C$7',
         'G': f'IF(COUNTIF(L_Sleutel,$J{r}&"|"&$E{r}&"|"&$F{r}&"|"&$I{r}&"|"&$O{r})=0,0,_xlfn.MAXIFS(L_Revisie,L_Sleutel,$J{r}&"|"&$E{r}&"|"&$F{r}&"|"&$I{r}&"|"&$O{r})+1)',
         'H': 'IF($C$10="","",$C$10)', 'I': '$AE$11',
         'J': f'INDEX($AB${OUT0}:$AB${OUT0+ROWS_SRC-1},$T{r})',
         'K': f'IFERROR(LEFT($J{r},FIND(".",$J{r},3)-1),$J{r})',
         'L': f'IFERROR(INDEX(Bron_Versie!$C$1:$C${ROWS_SRC},{m})&"","")',
         'M': f'IF(IFERROR(INDEX(V_Type,MATCH($J{r},V_WBS,0))&"","")<>"",INDEX(V_Type,MATCH($J{r},V_WBS,0))&"",IF(ISNUMBER(SEARCH("deelnem",IFERROR(INDEX(Bron_Versie!$B$1:$B${ROWS_SRC},{m})&"",""))),"Deelneming","Opstal"))',
         'N': f'IF(IFERROR(INDEX(V_IFRS,MATCH($J{r},V_WBS,0))&"","")<>"",INDEX(V_IFRS,MATCH($J{r},V_WBS,0))&"","J")',
         'O': f'$U{r}',
         'P': f'IFERROR(SUMIFS(INDEX(Bron_Versie!$A$1:$AZ${ROWS_SRC},0,IF($U{r}="Omzet FY",$AG$7,IF($U{r}="BOR FY",$AG$8,$AG$9))),Bron_Versie!$A$1:$A${ROWS_SRC},$J{r})*$AE$12,0)',
         'Q': '"EUR"',
         'R': '"Overgezet uit oude RDFO: "&$C$11',
        }
        sty = {'C': DATE, 'H': DATE, 'P': NUM}
        for col, f in F.items():
            s.set(f'{col}{r}', formula=z + f + ')', style=sty.get(col, CELL), is_str=col not in 'CEGHIP')
    s.root.find(N+'dimension').set('ref', f'A1:AG{OUT0+NOUT-1}')

def main(src, dst):
    zin = zipfile.ZipFile(src); paths = sheet_paths(zin)
    names = zin.namelist()
    n = max(int(re.search(r'sheet(\d+)\.xml', p).group(1)) for p in names if re.match(r'xl/worksheets/sheet\d+\.xml$', p))
    wbx = zin.read('xl/workbook.xml').decode()
    if 'name="Oude versies"' in wbx: raise SystemExit('al aanwezig')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode()
    rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels))
    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wbx))
    # stijlen uit bestaande bladen
    inv = Sheet(zin.read(paths['Invoer'])); his = Sheet(zin.read(paths['Historie'])); akd = Sheet(zin.read(paths['AK-dekking']))
    st = dict(title=inv.style('A5'), text=inv.style('A6'), label=inv.style('C9'), inp_txt=inv.style('D14'),
              inp_date=inv.style('D10'), inp_num=inv.style('D9'), hdr=his.style('B5'), cell=his.style('A6'),
              num=his.style('P6'), date=his.style('C6'), help=inv.style('E9'))
    cols = ('<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/>'
            '<col min="3" max="3" width="13" customWidth="1"/><col min="4" max="9" width="12" customWidth="1"/>'
            '<col min="10" max="10" width="18" customWidth="1"/><col min="11" max="11" width="13" customWidth="1"/>'
            '<col min="12" max="12" width="34" customWidth="1"/><col min="13" max="15" width="13" customWidth="1"/>'
            '<col min="16" max="16" width="14" customWidth="1"/><col min="17" max="17" width="7" customWidth="1"/>'
            '<col min="18" max="18" width="34" customWidth="1"/><col min="20" max="21" width="9" hidden="1" customWidth="1"/>'
            '<col min="27" max="33" width="9" hidden="1" customWidth="1"/></cols>')
    dv = ('<dataValidations count="2">'
          '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C7"><formula1>"Budget,Q1,Q2,Q3,Q4"</formula1></dataValidation>'
          '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C12"><formula1>"EUR,EUR x 1.000,EUR mln"</formula1></dataValidation>'
          '</dataValidations>')
    s_ov = Sheet(skeleton(cols, dv, tab='FF2E75B6', freeze=OUT0)); build_oude(s_ov, st)
    s_bv = Sheet(skeleton(tab='FFFFC000'))
    s_bv.set('A1', text="Plak hier (A1) het oude tabblad 'Budget' of 'Prognose Qx' als waarden. Zie tabblad Oude versies.")
    new = [('Oude versies', n+1, s_ov.xml()), ('Bron_Versie', n+2, s_bv.xml())]
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            d = data.decode(); add = ''
            for i, (nm, num, _) in enumerate(new):
                add += f'<sheet name="{nm}" sheetId="{sid+1+i}" r:id="rId{rid+1+i}"/>'
            d = d.replace('</sheets>', add + '</sheets>'); data = d.encode()
        elif item.filename == 'xl/_rels/workbook.xml.rels':
            d = data.decode(); add = ''
            for i, (nm, num, _) in enumerate(new):
                add += (f'<Relationship Id="rId{rid+1+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                        f'Target="worksheets/sheet{num}.xml"/>')
            data = d.replace('</Relationships>', add + '</Relationships>').encode()
        elif item.filename == '[Content_Types].xml':
            d = data.decode(); add = ''.join(
                f'<Override PartName="/xl/worksheets/sheet{num}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for _, num, _ in new)
            data = d.replace('</Types>', add + '</Types>').encode()
        elif item.filename == 'docProps/app.xml':
            d = data.decode()
            m = re.search(r'(<vt:lpstr>Werkbladen</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)', d)
            m2 = re.search(r'<TitlesOfParts><vt:vector size="(\d+)"', d)
            if m and m2:
                k = int(m.group(2))
                d = d[:m.start(2)] + str(k+2) + d[m.end(2):]
                d = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)', lambda x: x.group(1)+str(int(x.group(2))+2), d)
                parts = list(re.finditer(r'<vt:lpstr>.*?</vt:lpstr>', d[d.index('<TitlesOfParts>'):]))
                off = d.index('<TitlesOfParts>') + parts[k-1].end()
                d = d[:off] + '<vt:lpstr>Oude versies</vt:lpstr><vt:lpstr>Bron_Versie</vt:lpstr>' + d[off:]
            data = d.encode()
        elif item.filename == paths['Handleiding']:
            h = Sheet(data)
            h.set('D82', style=h.style('D82'), text=(
                "Twee manieren. (1) Met macro: klik op Start op 'Oude prognoses overzetten' en kies je oude RDFO-bestand; een versie die er al in staat, wordt overgeslagen. "
                "(2) Zonder macro, of om een versie te vervangen: plak het oude tabblad 'Budget' of 'Prognose Qx' in Bron_Versie (A1), vul op 'Oude versies' versie en versiejaar in, "
                "controleer de aansluiting en plak het blok A:R als waarden op de eerste vrije rij van Historie. Bestaande versies krijgen revisie +1. "
                "Het officiële AK-totaal van de regio gaat via Invoer › 3 (regel 18 op 'Oude versies')."))
            data = h.xml()
        zout.writestr(item, data)
    for nm, num, xml in new:
        zout.writestr(f'xl/worksheets/sheet{num}.xml', xml)
    zout.close()

if __name__ == '__main__': main(sys.argv[1], sys.argv[2])
