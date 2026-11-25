"""Ronde 6: bandbreedte-grafiek zichtbaar (gegevens in verborgen kolommen), reserves per soort met eigen vrijvalkwartaal,
Projecten zonder lege projecten (opgeslagen filter), knop rekenbladen tonen/verbergen op Start, links naar de berekening
per regel op RDFO en Bandbreedte.
usage: patch_v6.py in.xlsm out.xlsm [herberekend.xlsx]"""
import sys, os, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import split, col2n, n2col, N
from patch_eenvoud import add_hyperlink, clear, get_or_make
from lxml import etree

def hidden_cols(s):
    out = set(); cols = s.root.find(N+'cols')
    if cols is not None:
        for c in cols:
            if c.get('hidden') == '1': out.update(range(int(c.get('min')), int(c.get('max')) + 1))
    return out

def charts_zichtbaar(bk):
    """Excel tekent standaard geen gegevens uit verborgen rijen/kolommen; zet plotVisOnly uit waar dat speelt."""
    names = {}
    for m in re.finditer(r'<sheet name="([^"]+)"', bk.text('xl/workbook.xml')): names[m.group(1).replace('&apos;', "'")] = 1
    for fn in [k for k in bk.data if re.match(r'xl/charts/chart\d+\.xml$', k)]:
        t = bk.text(fn); hit = False
        for sh, c1, r1, c2, r2 in re.findall(r"<c:f>'?([^'!<]+)'?!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?</c:f>", t):
            if sh not in names: continue
            s = bk.sheet(sh); hc = hidden_cols(s)
            a, b = col2n(c1), col2n(c2 or c1)
            if any(c in hc for c in range(a, b + 1)): hit = True
            for r in range(int(r1), int(r2 or r1) + 1):
                if r in s.rows and s.rows[r].get('hidden') == '1': hit = True
        if hit and '<c:plotVisOnly val="1"/>' in t:
            bk.settext(fn, t.replace('<c:plotVisOnly val="1"/>', '<c:plotVisOnly val="0"/>')); print('  grafiek toont verborgen cellen:', fn)

def reserves(bk):
    r = bk.sheet('Reserves')
    r.set('B2', text='Invullen op de Projectlijst: bedrag per soort (kolommen L–N) en het kwartaal waarin die vrijvalt (P–R). Leeg of "Later" = na ' + "volgend jaar.", style=r.style('B2'))
    for row, amt, kw in ((6, 'L', 'P'), (7, 'M', 'Q'), (8, 'N', 'R')):
        for c, y in (('D', 0), ('E', 1)):
            r.set(f'{c}{row}', style=r.style(f'{c}{row}'),
                  formula=f'SUMIFS(Projectlijst!${amt}$8:${amt}$107,PL_Telt,"Ja",Projectlijst!${kw}$8:${kw}$107,(P_RJ+{y})&"-*")')
    r.set('H13', text='Valt vrij in\n(comm. · oper. · nazorg)', style=r.style('H13'))
    for row in range(14, 200):
        if row not in r.rows: continue
        h = r.cell(f'H{row}').find(N+'f')
        if h is None or 'SMALL($Z$200:$Z$299,' not in (h.text or ''): continue
        k = re.search(r'SMALL\(\$Z\$200:\$Z\$299,(\d+)\)', h.text).group(1)
        ix = f'SMALL($Z$200:$Z$299,{k})'
        g = lambda c: f'IF(INDEX(Projectlijst!${c}$8:${c}$107,{ix})="","–",INDEX(Projectlijst!${c}$8:${c}$107,{ix}))'
        r.set(f'H{row}', is_str=True, style=r.style(f'C{row}'),
              formula=f'IFERROR(IF(INDEX(Projectlijst!$O$8:$O$107,{ix})="","",{g("P")}&" · "&{g("Q")}&" · "&{g("R")}),"")')
        clear(r, f'I{row}')
        for c, src in (('J', 'S'), ('K', 'T'), ('L', 'AR')):
            f = r.cell(f'{c}{row}').find(N+'f')
            if f is not None: f.text = re.sub(r'Projectlijst!\$[A-Z]+\$8:\$[A-Z]+\$107', f'Projectlijst!${src}$8:${src}$107', f.text)
    clear(r, 'I13')
    cols = r.root.find(N+'cols')
    for c in cols:
        if c.get('min') == '8' and c.get('max') == '9':
            e = copy.deepcopy(c); c.set('max', '8'); c.set('width', '24'); e.set('min', '9'); e.set('hidden', '1'); c.addnext(e)
    bk.set_name('RS_Effect0', 'Projectlijst!$S$8:$S$107'); bk.set_name('RS_Effect1', 'Projectlijst!$T$8:$T$107')
    bk.set_name('RS_EffectLater', 'Projectlijst!$AR$8:$AR$107'); bk.set_name('RS_Invoer', 'Projectlijst!$L$8:$R$107')
    # verouderde uitleg (% gereed) in teksten
    sst = bk.text('xl/sharedStrings.xml')
    sst = sst.replace('Zelf invullen op de Projectlijst (L–N) met % gereed (P–Q). Vrijval = reserve x % gereed.',
                      'Zelf invullen op de Projectlijst: bedrag per soort (L–N) en kwartaal van vrijval (P–R).')
    bk.settext('xl/sharedStrings.xml', sst)

def projecten(bk, herberekend):
    s = bk.sheet('Projecten')
    af = s.root.find(N+'autoFilter')
    for fc in af.findall(N+'filterColumn'): af.remove(fc)
    fc = etree.SubElement(af, N+'filterColumn'); fc.set('colId', str(col2n('X') - 1))
    fl = etree.SubElement(fc, N+'filters'); etree.SubElement(fl, N+'filter').set('val', 'met cijfers')
    pr = s.root.find(N+'sheetPr'); pr.set('filterMode', '1')
    toon = {}
    if herberekend:
        import openpyxl
        v = openpyxl.load_workbook(herberekend, data_only=True, read_only=True)
        for i, (x,) in enumerate(v['Projecten'].iter_rows(min_row=10, max_row=109, min_col=24, max_col=24, values_only=True)):
            toon[10 + i] = (x == 'met cijfers')
    for row in range(10, 110):
        el = s.row(row)
        if toon.get(row, False): el.attrib.pop('hidden', None)
        else: el.set('hidden', '1')
    s.set('B7', style=s.style('B7'), text="Geel = kwartaal gekozen op het Dashboard. Mutatie: groen = hoger, rood = lager. "
          "Projecten zonder cijfers worden verborgen (knop rechtsboven: alles tonen). Niet sorteren.")
    print('  Projecten zichtbaar:', sum(toon.values()), 'van', len([1 for r in range(10, 110) if r in toon]))

def start_knop(bk):
    rfn = bk.paths['Start'].replace('worksheets/', 'worksheets/_rels/') + '.rels'
    fn = 'xl/drawings/' + re.search(r'Target="\.\./drawings/([^"]+)"', bk.text(rfn)).group(1)
    d = bk.text(fn)
    if 'RDFO_ToonRekenbladen' in d: return
    tmpl = next(a for a in re.findall(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', d, re.S) if 'RDFO_MaakPowerPoint' in a)
    knop = tmpl.replace('RDFO_MaakPowerPoint', 'RDFO_ToonRekenbladen')
    knop = re.sub(r'(<xdr:(?:from|to)><xdr:col>)(\d+)', lambda m: m.group(1) + str(int(m.group(2)) + 1), knop)
    knop = re.sub(r'descr="[^"]*"', 'descr="Rekenbladen tonen / verbergen"', knop)
    knop = re.sub(r'name="Knop \d+"', 'name="Knop 1101"', knop); knop = re.sub(r'<xdr:cNvPr id="\d+"', '<xdr:cNvPr id="1101"', knop)
    knop = re.sub(r'(<a:t>)[^<]*(</a:t>)', r'\1Rekenbladen tonen / verbergen\2', knop, count=1)
    # andere kleur dan de PowerPoint-knop: wit met navy rand en tekst
    knop = re.sub(r'<a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"/></a:solidFill>', '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>', knop, count=1)
    knop = knop.replace('<a:ln><a:noFill/></a:ln>', '<a:ln w="19050"><a:solidFill><a:srgbClr val="14284B"/></a:solidFill></a:ln>', 1)
    knop = re.sub(r'<a:extLst>.*?</a:extLst>', '', knop, count=1, flags=re.S)
    d = d.replace(tmpl, tmpl + knop); bk.settext(fn, d)
    st = bk.sheet('Start')
    st.set('C12', text='Verbergt of toont de grijze rekenbladen achteraan.', style=st.style('B12'))

def rdfo_links(bk):
    s = bk.sheet('RDFO'); LNK = s.style('E3')
    s.set('R6', text='Berekening', style=s.style('O6'))
    doel = {}
    for row in range(7, 32):
        f = s.cell(f'C{row}').find(N+'f')
        if f is None or not f.text: continue
        m = re.search(r'CALC_Versies!\$([A-Z]+)\$5', f.text)
        if m: doel[row] = (f"'CALC_Versies'!{m.group(1)}4", 'CALC_Versies')
        elif 'SUM(' in f.text or 'IF(AND(ISNUMBER(' in f.text: doel[row] = None
    doel[18] = ("'AK-dekking'!A1", 'AK-dekking')
    for row in (34, 35, 36, 37): doel[row] = ("'Reserves'!B5", 'Reserves')
    for row, c in ((40, 'L'), (41, 'M'), (42, 'N')): doel[row] = (f"'Projectlijst'!{c}7", 'Projectlijst')
    for row, t in sorted(doel.items()):
        if t is None: continue
        loc, naam = t
        s.set(f'R{row}', text=f'{naam} →', style=LNK); add_hyperlink(s, f'R{row}', loc, naam)
    for c in s.root.find(N+'cols'):
        if c.get('min') == '18' and c.get('max') == '18': c.set('width', '15')

def bandbreedte_links(bk):
    b = bk.sheet('Bandbreedte'); LNK = bk.sheet('RDFO').style('E3')
    kol = {6: 'K', 7: 'F', 8: 'L', 9: 'M', 10: 'N'}
    for row, c in kol.items():
        b.set(f'F{row}', text='SCEN →', style=LNK); add_hyperlink(b, f'F{row}', f"'SCEN'!{c}3", 'SCEN')
    b.set('F12', text='Versies →', style=LNK); add_hyperlink(b, 'F12', "'Versies'!A1", 'Versies')
    b.set('F5', text='Berekening', style=b.style('E5'))

def main(src, dst, herberekend=None):
    bk = Book(src)
    charts_zichtbaar(bk); reserves(bk); projecten(bk, herberekend); start_knop(bk); rdfo_links(bk); bandbreedte_links(bk)
    bk.save(dst)

if __name__ == '__main__': main(*sys.argv[1:])
