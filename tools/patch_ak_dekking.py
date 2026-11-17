"""Patch AK-dekking logic in an RDFO workbook (.xlsm) by editing the sheet XML directly.
usage: python3 -I patch_ak.py in.xlsm out.xlsm"""
import sys, zipfile, re
from lxml import etree
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N='{%s}'%NS

def col2n(c):
    n=0
    for ch in c: n=n*26+ord(ch)-64
    return n
def n2col(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
def split(ref):
    m=re.match(r'([A-Z]+)(\d+)$',ref); return m.group(1),int(m.group(2))

class Sheet:
    def __init__(self,xml):
        self.root=etree.fromstring(xml)
        self.sd=self.root.find(N+'sheetData')
        self.rows={int(r.get('r')):r for r in self.sd.findall(N+'row')}
    def row(self,r):
        if r in self.rows: return self.rows[r]
        new=etree.Element(N+'row'); new.set('r',str(r))
        after=[k for k in self.rows if k<r]
        if after: self.rows[max(after)].addnext(new)
        else: self.sd.insert(0,new)
        self.rows[r]=new; return new
    def cell(self,ref):
        c,r=split(ref); row=self.row(r)
        if 'spans' in row.attrib: del row.attrib['spans']
        prev=None
        for el in row.findall(N+'c'):
            cc,_=split(el.get('r'))
            if cc==c: return el
            if col2n(cc)<col2n(c): prev=el
        el=etree.Element(N+'c'); el.set('r',ref)
        if prev is not None: prev.addnext(el)
        else: row.insert(0,el)
        return el
    def style(self,ref):
        c,r=split(ref)
        if r in self.rows:
            for el in self.rows[r].findall(N+'c'):
                if el.get('r')==ref: return el.get('s')
        return None
    def set(self,ref,formula=None,text=None,style=None,is_str=False):
        el=self.cell(ref)
        for ch in list(el): el.remove(ch)
        for a in ('t',): 
            if a in el.attrib: del el.attrib[a]
        if style is not None: el.set('s',style)
        if formula is not None:
            f=etree.SubElement(el,N+'f'); f.text=formula
            if is_str: el.set('t','str')
        elif text is not None:
            el.set('t','inlineStr'); is_=etree.SubElement(el,N+'is'); t=etree.SubElement(is_,N+'t'); t.text=text
    def xml(self):
        return etree.tostring(self.root,xml_declaration=True,encoding='UTF-8',standalone=True)

def sheet_paths(z):
    wb=z.read('xl/workbook.xml').decode()
    rels=z.read('xl/_rels/workbook.xml.rels').decode()
    rid={m.group(2):m.group(1) for m in re.finditer(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"',rels)}
    rid={v:k for k,v in rid.items()}
    out={}
    for m in re.finditer(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"',wb):
        name=m.group(1).replace('&apos;',"'"); out[name]='xl/'+rid[m.group(2)]
    return out

QC={1:'BM',2:'BN',3:'BO',4:'BP'}
AKc={1:'F',2:'G',3:'H',4:'I'}
Kc={1:'M',2:'N',3:'O',4:'P'}
Rc={1:'AD',2:'AE',3:'AF',4:'AG'}

def patch_ak(s):
    sty_hdr=s.style('AC20'); sty_dat=s.style('AC21')
    # header
    for q in range(1,5):
        s.set(f'{Rc[q]}20',formula=f'"Prognose-"&CHAR(10)&"kosten Q{q}"',style=sty_hdr,is_str=True)
    s.set('AD19',text='HULP: NOG NIET GEREALISEERDE KOSTEN PER KWARTAAL',style=s.style('Z19') or s.style('M19'))
    for r in range(21,121):
        v=r-15
        s.set(f'AB{r}',formula='+'.join(f'IF(AK_PeilQ<={q},N(VERW_Project!${QC[q]}{v}),0)' for q in range(1,5)),style=s.style(f'AB{r}'))
        for q in range(1,5):
            s.set(f'{Rc[q]}{r}',style=sty_dat,formula=
              f'IF(C{r}<>"Ja","",IF({q}<AK_PeilQ,0,IF(AB{r}<>0,N(VERW_Project!${QC[q]}{v}),IF(OR({q}>AK_PeilQ,AND({q}=4,AK_PeilQ>=4)),N(AC{r})/MAX(1,4-AK_PeilQ),0))))')
            s.set(f'{Kc[q]}{r}',style=s.style(f'{Kc[q]}{r}'),formula=
              f'IF(C{r}<>"Ja","",IF({q}<=AK_PeilQ,N(AA{r})/MAX(1,AK_PeilQ),0)+N({Rc[q]}{r}))')
            s.set(f'{AKc[q]}{r}',style=s.style(f'{AKc[q]}{r}'),formula=
              f'IF(C{r}<>"Ja","",IF(AND({q}<=AK_PeilQ,AKR_NJaar>0),SUMIFS(AKR_Bedrag,AKR_WBS,$A{r},AKR_Jaar,P_RJ,AKR_Kw,IF({q}=AK_PeilQ,">=","")&{q})+IF(D{r}="",0,N({Rc[q]}{r})/P_AKFactor*D{r}),IF(D{r}="","",{Kc[q]}{r}/P_AKFactor*D{r})))')
        s.set(f'Q{r}',style=s.style(f'Q{r}'),formula=
          f'IF(C{r}<>"Ja","",IF(AKR_NJaar>0,SUMIFS(AKR_Bedrag,AKR_WBS,$A{r},AKR_Jaar,P_RJ),IF(D{r}="",0,N(AA{r})/P_AKFactor*D{r})))')
    for q in range(1,5):
        col=chr(ord('E')+q)
        s.set(f'{col}6',style=s.style(f'{col}6'),is_str=True,formula=
          f'IF({q}<=AK_PeilQ,IF(AKR_NJaar>0,"geboekt","kosten LJ")&IF(SUM(${Rc[q]}$21:${Rc[q]}$120)<>0," + rest",""),"prognose")')
        k=f'{Kc[q]}'; a=AKc[q]; rr=Rc[q]
        m=f'MATCH($F$11,$B$21:$B$120,0)'
        s.set(f'B{11+q}',style=s.style(f'B{11+q}'),is_str=True,formula=
          f'IF($F$11="","",IFERROR(IF(INDEX($C$21:$C$120,{m})<>"Ja","Q{q}: dit project telt niet mee voor AK (Projectlijst)",'
          f'IF(AND({q}<=AK_PeilQ,AKR_NJaar>0),"Q{q} geboekt: € "&FIXED(SUMIFS(AKR_Bedrag,AKR_WBS,INDEX($A$21:$A$120,{m}),AKR_Jaar,P_RJ,AKR_Kw,IF({q}=AK_PeilQ,">=","")&{q}),0)&" (grootboek)"'
          f'&IF(N(INDEX(${rr}$21:${rr}$120,{m}))<>0," + nog te realiseren kosten € "&FIXED(N(INDEX(${rr}$21:${rr}$120,{m})),0)&" / "&FIXED(P_AKFactor,2)&" x "&FIXED(N(INDEX($D$21:$D$120,{m}))*100,1)&"%","")'
          f'&" = € "&FIXED(N(INDEX(${a}$21:${a}$120,{m})),0),'
          f'"Q{q} "&IF({q}<=AK_PeilQ,"berekend (geen grootboek, kosten uit Kosten LJ"&IF(N(INDEX(${rr}$21:${rr}$120,{m}))<>0," + rest kwartaal","")&")","prognose")&": kosten € "&FIXED(N(INDEX(${k}$21:${k}$120,{m})),0)&" / "&FIXED(P_AKFactor,2)&" x "&FIXED(N(INDEX($D$21:$D$120,{m}))*100,1)&"% = € "&FIXED(N(INDEX(${a}$21:${a}$120,{m})),0))),"Project niet gevonden"))')
    s.set('B2',style=s.style('B2'),is_str=True,formula=
      '"Per kwartaal: t/m Q"&AK_PeilQ&" de geboekte AK uit het grootboek (zonder dump: Kosten LJ / "&FIXED(P_AKFactor,2)&" x AK %). Kosten die nog niet gerealiseerd zijn (rest van het lopende kwartaal en later) / "&FIXED(P_AKFactor,2)&" x AK %. Jaar = Q1 + Q2 + Q3 + Q4."')
    s.set('B3',style=s.style('F18') or s.style('B2'),is_str=True,formula=
      'IF(AKR_NJaar=0,"LET OP: geen grootboekdump AK voor "&P_RJ&" ingelezen (Bron_AKR leeg). Q1 t/m Q"&AK_PeilQ&" zijn nu geschat met Kosten LJ x AK %; de werkelijk geboekte AK en AK-boekingen zonder WBS (fees) ontbreken. Lees de dump AK-realisatie in.",'
      'IF(AND(AKR_MaxPer>0,AKR_MaxPer<P_PeilMaand),"LET OP: de grootboekdump loopt t/m periode "&AKR_MaxPer&", de peildatum is maand "&P_PeilMaand&". AK op kosten van periode "&(AKR_MaxPer+1)&" t/m "&P_PeilMaand&" is nog niet geboekt en telt niet mee. Gebruik een dump t/m de peildatum.",""))')
    # dimension & cols
    dim=s.root.find(N+'dimension')
    if dim is not None: dim.set('ref','A1:AG122')
    cols=s.root.find(N+'cols')
    for c in cols.findall(N+'col'):
        if c.get('min')=='26' and c.get('max')=='29': c.set('max','33')

def patch_stg(s):
    s.set('E4',text='Periode',style=s.style('D4'))
    s.set('L2',text='Laatste periode in rapportagejaar')
    s.set('L3',formula='MAX(0,MAX($E$5:$E$3004))')
    for r in range(5,3005):
        i=r-4
        s.set(f'E{r}',formula=f'IF(OR($B{r}="",$F$3=0),"",IF(N($C{r})<>P_RJ,"",IFERROR(--INDEX(Bron_AKR!$A$2:$Z$3001,{i},$F$3),"")))')

def patch_ctl(s):
    s.set('D49',style=s.style('D49'),text='Grootboekdump AK: periode ontbreekt of loopt niet t/m de peildatum')
    s.set('E49',style=s.style('E49'),formula='IF(AKR_Rijen=0,0,IF(STG_AKR!$F$3=0,1,IF(AND(AKR_MaxPer>0,AKR_MaxPer<P_PeilMaand),1,0)))')
    s.set('H49',style=s.style('H49'),is_str=True,formula=
      'IF(AKR_Rijen=0,"",IF(STG_AKR!$F$3=0,"Kolom Boekingsperiode niet gevonden: alle geboekte AK telt in het lopende kwartaal (jaartotaal klopt wel).","Dump loopt t/m periode "&AKR_MaxPer&", peildatum is maand "&P_PeilMaand&". AK op kosten na periode "&AKR_MaxPer&" is nog niet geboekt en ontbreekt. Gebruik een dump t/m de peildatum."))')

def main(src,dst):
    zin=zipfile.ZipFile(src); paths=sheet_paths(zin)
    targets={paths['AK-dekking']:patch_ak, paths['STG_AKR']:patch_stg, paths['Controle']:patch_ctl}
    zout=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data=zin.read(item.filename)
        if item.filename in targets:
            s=Sheet(data); targets[item.filename](s); data=s.xml()
        elif item.filename=='xl/calcChain.xml':
            continue
        elif item.filename=='[Content_Types].xml':
            data=re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>',b'',data)
        elif item.filename=='xl/_rels/workbook.xml.rels':
            data=re.sub(rb'<Relationship [^>]*Target="calcChain.xml"/>',b'',data)
        elif item.filename=='xl/workbook.xml':
            d=data.decode()
            if 'name="AKR_MaxPer"' not in d:
                d=d.replace('<definedName name="AKR_N">','<definedName name="AKR_MaxPer">STG_AKR!$L$3</definedName><definedName name="AKR_N">',1)
                assert 'AKR_MaxPer' in d
            d=re.sub(r'<calcPr([^>]*?)/>',lambda m:'<calcPr'+re.sub(r' fullCalcOnLoad="\d"','',m.group(1))+' fullCalcOnLoad="1"/>',d)
            data=d.encode()
        zout.writestr(item,data)
    zout.close()
if __name__=='__main__': main(sys.argv[1],sys.argv[2])
