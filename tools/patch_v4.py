"""Ronde 4: de huidige upload is het versie-tabblad van 'Versie van de upload' (Invoer D14). Bron_ZPMW03 vervalt
(verborgen); staan daar nog gegevens, dan gaan die naar het lege versie-tabblad. usage: patch_v4.py in.xlsm out.xlsm"""
import sys, os, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_eenvoud import N, expand_shared
from patch_ak import split
from lxml import etree

VERS = ['Budget', 'Q1', 'Q2', 'Q3', 'Q4']
def main(src, dst):
    bk = Book(src)
    rngs = ','.join(f"'{v} ZPMW03'!$A$1:$ZZ$600" for v in VERS)
    bk.set_name('P_UploadIdx', 'IFERROR(MATCH(Invoer!$D$14,{"Budget","Q1","Q2","Q3","Q4"},0),4)')
    bk.set_name('B_ZPMW03', f'CHOOSE(P_UploadIdx,{rngs})')
    bk.set_name('P_ZHdr', 'MAX(1,INDEX(VERSIE_REKEN!$CE$6:$CE$10,P_UploadIdx))')
    # STG_ZPMW03 leest het gekozen versie-tabblad
    s = bk.sheet('STG_ZPMW03'); expand_shared(s, s.rows.keys())
    for row in s.rows.values():
        for el in row.findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.text and 'Bron_ZPMW03' in f.text:
                f.text = f.text.replace('Bron_ZPMW03!$A$1:$ZZ$501', 'B_ZPMW03').replace('Bron_ZPMW03!$1:$1', 'INDEX(B_ZPMW03,P_ZHdr,0)')
    for k in range(1, 501):
        s.set(f'A{5+k}', formula=f'P_ZHdr+{k}', style=s.style(f'A{5+k}'))
    s.set('A1', text='Intern veld (leest het versie-tabblad van Invoer › Versie van de upload)')
    # Controle
    c = bk.sheet('Controle')
    for row in c.rows.values():
        for el in row.findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.text and 'Bron_ZPMW03' in f.text:
                f.text = (f.text.replace('COUNTA(Bron_ZPMW03!$A$2:$A$501)', 'MAX(0,COUNTA(INDEX(B_ZPMW03,0,1))-P_ZHdr)')
                                 .replace('COUNTA(Bron_ZPMW03!$A$502:$A$20000)', '0').replace('Bron_ZPMW03!$A$502:$A$20000', 'Invoer!$ZZ$1'))
    X = 'MAX(0,COUNTIF(INDEX(B_ZPMW03,0,MAX(1,STG_ZPMW03!$B$3)),"?*")-1)'
    c.set('F9', is_str=True, style=c.style('F9'), formula=f'"verwacht "&{X}')
    c.set('G9', is_str=True, style=c.style('G9'), formula=f'IF(E9={X},"OK","FOUT")')
    c.set('D9', text='Regels ZPMW03 (huidige versie) gelezen', style=c.style('D9'))
    c.set('G14', is_str=True, style=c.style('G14'), formula='IF(INDEX(VERSIE_REKEN!$CE$6:$CE$10,P_UploadIdx)=0,"INFO",IF(E14=0,"OK","FOUT"))')
    c.set('H9', is_str=True, style=c.style('H9'), formula='"Leest tabblad \'"&Invoer!$D$14&" ZPMW03\' (Invoer › Versie van de upload)."')
    # Invoer
    inv = bk.sheet('Invoer')
    expand_shared(inv, range(31, 131))
    for i in range(100):   # lege soort telt als Opstal
        r = 31 + i; p = 8 + i
        inv.set(f'F{r}', is_str=True, style=inv.style(f'F{r}'), formula=f'IF(Projectlijst!$A${p}="","",IF(Projectlijst!$C${p}="","Opstal",Projectlijst!$C${p}))')
    inv.set('C14', text='Versie van de upload (= huidig kwartaal)', style=inv.style('C14'))
    inv.set('E14', text="Budget / Q1 / Q2 / Q3 / Q4. Bepaalt welk tabblad de huidige cijfers levert, bijv. Q3 → tabblad 'Q3 ZPMW03'.", style=inv.style('E14'))
    # Dashboard volgt standaard Invoer (kiezen blijft mogelijk)
    d = bk.sheet('Dashboard')
    d.set('E5', is_str=True, formula='Invoer!$D$14', style=d.style('E5'))
    d.set('B5', formula='P_RJ', style=d.style('B5'))
    # Versies: huidige upload markeren
    v = bk.sheet('Versies')
    for i, ver in enumerate(VERS):
        r = 6 + i; f = v.cell(f'O{r}').find(N+'f')
        f.text = f'IF($B{r}=Invoer!$D$14,"★ huidige upload · ","")&' + f.text
    v.set('B4', text="De versie die op Invoer als 'Versie van de upload' staat (★) levert ook de huidige cijfers voor alle rapporten.", style=v.style('B3'))
    # gegevens uit Bron_ZPMW03 naar het lege versie-tabblad van de upload
    up = None
    el = bk.sheet('Invoer').cell('D14'); t = el.get('t')
    if t == 's':
        idx = int(el.find(N+'v').text); sst = re.findall(r'<si>(.*?)</si>', bk.text('xl/sharedStrings.xml'), re.S)
        up = re.sub('<[^>]+>', '', sst[idx])
    elif t == 'inlineStr': up = ''.join(el.itertext())
    elif el.find(N+'v') is not None: up = el.find(N+'v').text
    bron = bk.sheet('Bron_ZPMW03')
    if up in VERS and len(bron.rows) > 1:
        doel = bk.sheet(f'{up} ZPMW03')
        sd = doel.root.find(N+'sheetData')
        for r in list(sd): sd.remove(r)
        for r in sorted(bron.rows):
            sd.append(copy.deepcopy(bron.rows[r]))
        doel.rows = {int(r.get('r')): r for r in sd.findall(N+'row')}
        dim = bron.root.find(N+'dimension'); doel.root.find(N+'dimension').set('ref', dim.get('ref') if dim is not None else 'A1')
        print('gegevens Bron_ZPMW03 ->', f'{up} ZPMW03', len(bron.rows), 'rijen')
    # Bron_ZPMW03 verbergen en leegmaken (alleen nog voor oude macro's)
    sd = bron.root.find(N+'sheetData')
    for r in list(sd): sd.remove(r)
    bron.rows = {}
    bron.set('A1', text="Niet meer gebruikt: plak de ZPMW03 in het versie-tabblad (bijv. 'Q3 ZPMW03').")
    wb = bk.text('xl/workbook.xml')
    wb = re.sub(r'(<sheet name="Bron_ZPMW03" sheetId="\d+")( state="\w+")?', r'\1 state="hidden"', wb)
    bk.settext('xl/workbook.xml', wb)
    bk.save(dst)

if __name__ == '__main__': main(sys.argv[1], sys.argv[2])
