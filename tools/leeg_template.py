"""Maakt een leeg template (zoals de macro 'Bestand leegmaken' + versie-plakbladen) en verwijdert opgeslagen
rekenresultaten, zodat er nergens oude bedragen zichtbaar zijn. usage: leeg_template.py in.xlsm out.xlsm"""
import sys, os, re
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import split, col2n, n2col, N

def name_ref(bk, nm):
    m = re.search(r'<definedName name="%s">([^<]+)</definedName>' % re.escape(nm), bk.text('xl/workbook.xml'))
    sh, rng = m.group(1).replace('$', '').rsplit('!', 1); sh = sh.strip("'").replace("''", "'")
    a, b = (rng.split(':') + [rng])[:2]
    (c0, r0), (c1, r1) = split(a), split(b)
    return sh, col2n(c0), r0, col2n(c1), r1

def clear_values(s, c0, r0, c1, r1, whole_rows=False):
    for r, row in s.rows.items():
        if r < r0 or r > r1: continue
        for el in row.findall(N+'c'):
            c = col2n(split(el.get('r'))[0])
            if not whole_rows and (c < c0 or c > c1): continue
            if el.find(N+'f') is not None: continue
            for ch in list(el): el.remove(ch)
            if 't' in el.attrib: del el.attrib['t']

def main(src, dst):
    bk = Book(src)
    for sh in ['Bron_ZPMW03', 'Bron_Woningetalage', 'Bron_AK', 'Bron_AKR']:
        clear_values(bk.sheet(sh), 1, 2, 10**6, 10**7, whole_rows=True)
    for v in ['Budget', 'Q1', 'Q2', 'Q3', 'Q4']:
        vs = bk.sheet(f'{v} ZPMW03'); clear_values(vs, 1, 1, 10**6, 10**7, whole_rows=True)
        vs.set('A1', text=f'Plak hier de ZPMW03-export van {v} (hele blad, vanaf A1).')
    ms = bk.sheet('Bron_Mijlpalen'); clear_values(ms, 1, 1, 10**6, 10**7, whole_rows=True)
    ms.set('A1', text='Plak hier de mijlpalen-export (hele blad, vanaf A1): rij 1 = mijlpalen, rij 2 = verwacht/werkelijk, kolom A = "WBS - naam".')
    clear_values(bk.sheet('Projecten'), 19, 10, 19, 109)
    for sh, ref in [('Invoer', 'D8')]:
        s = bk.sheet(sh); s.set(ref, text='Regio (invullen)', style=s.style(ref))
    d = bk.sheet('Dashboard'); d.set('N5', text='Alle', style=d.style('N5'))
    for sh, ref in [('Projectdetail', 'C4'), ('AK-dekking', 'F11'), ('Bandbreedte', 'AJ25')]:
        c, r = split(ref); clear_values(bk.sheet(sh), col2n(c), r, col2n(c), r)
    clear_values(bk.sheet('Versies'), 3, 6, 3, 10); clear_values(bk.sheet('Versies'), 11, 6, 11, 10)
    spec = [('PL_Invoer', 0, 0), ('PK_Keuzes', 0, 0), ('PK_AKEigen', 0, 0), ('RS_Invoer', 0, 0), ('Gebieden', 0, 0),
            ('MJ_Funnel', 0, 0), ('MJ_Target', 0, 0), ('DB_Omzet', 0, 7), ('KR_Project', 0, 7), ('A_Doel', 0, 12),
            ('T_Key', -8, 5), ('T_Txt', 0, 3), ('KT_WBS', 0, 6), ('KT_WBS', 10, 1), ('PL_Details', 0, 0), ('PL_Stoplicht', 0, 0)]
    for nm, off, width in spec:
        sh, c0, r0, c1, r1 = name_ref(bk, nm)
        c0 += off; c1 = c0 + width - 1 if width else c1 + off
        clear_values(bk.sheet(sh), c0, r0, c1, r1)
    h = bk.sheet('Historie'); clear_values(h, 1, 6, 18, 5005); clear_values(h, 24, 6, 34, 5005)
    for nm in ['P_BronZ', 'P_BronE', 'P_ImportID', 'P_ToelCorrectie', 'P_Importdatum', 'P_RegelsZ', 'P_RegelsE']:
        sh, c, r, _, _ = name_ref(bk, nm); s = bk.sheet(sh); clear_values(s, c, r, c, r)
        if nm in ('P_RegelsZ', 'P_RegelsE'): s.set(f'{n2col(c)}{r}', formula=None, style=s.style(f'{n2col(c)}{r}')); s.cell(f'{n2col(c)}{r}').append(__import__('lxml').etree.Element(N+'v')); s.cell(f'{n2col(c)}{r}').find(N+'v').text = '0'
    # opgeslagen rekenresultaten weg (worden bij openen opnieuw berekend)
    for name, p in list(bk.paths.items()):
        s = bk.sheet(name)
        for row in s.rows.values():
            for el in row.findall(N+'c'):
                f = el.find(N+'f')
                if f is None: continue
                v = el.find(N+'v')
                if v is not None: el.remove(v)
                if el.get('t') in ('str', 'e', 'b', 'n'): del el.attrib['t']
    for fn in list(bk.data):
        if fn.startswith('xl/charts/chart') and fn.endswith('.xml'):
            d = bk.text(fn)
            d = re.sub(r'<c:pt idx="\d+">.*?</c:pt>', '', d, flags=re.S)
            d = re.sub(r'<c:ptCount val="\d+"/>', '<c:ptCount val="0"/>', d)
            bk.settext(fn, d)
    bk.save(dst)

if __name__ == '__main__': main(sys.argv[1], sys.argv[2])
