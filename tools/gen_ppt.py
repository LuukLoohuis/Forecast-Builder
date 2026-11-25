"""Maakt de RDFO-presentatie zoals de macro RDFO_MaakPowerPoint, maar buiten Excel (python-pptx).
usage: gen_ppt.py bestand.xlsm doorgerekend.xlsx logo.png uit.pptx"""
import sys, re, datetime, zipfile
import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

NAVY = RGBColor(0x11, 0x1F, 0x39); GEEL = RGBColor(0xFE, 0xB9, 0x00); GRIJS = RGBColor(0x6B, 0x73, 0x85)
LICHT = RGBColor(0xF3, 0xF5, 0xF9); LIJN = RGBColor(0xD9, 0xDE, 0xE8); WIT = RGBColor(0xFF, 0xFF, 0xFF); ROOD = RGBColor(0xE3, 0x06, 0x13)
src, calc, logo, out = sys.argv[1:5]
F = openpyxl.load_workbook(src, keep_vba=False)              # formules, namen, voorwaardelijke opmaak
V = openpyxl.load_workbook(calc, data_only=True)             # doorgerekende waarden
def name(nm):
    sh, rng = F.defined_names[nm].attr_text.rsplit('!', 1)
    return sh.strip("'").replace("''", "'"), rng.replace('$', '')
def rng_cells(nm):
    sh, rng = name(nm); a, b = (rng.split(':') + [rng])[:2]
    m1 = re.match(r'([A-Z]+)(\d+)', a); m2 = re.match(r'([A-Z]+)(\d+)', b)
    return sh, CI(m1.group(1)), int(m1.group(2)), CI(m2.group(1)), int(m2.group(2))
def val(nm):
    sh, c, r, _, _ = rng_cells(nm); return V[sh].cell(r, c).value

def nl(x, dec, thou=True):
    s = f'{abs(x):,.{dec}f}'
    s = s.replace(',', 'X').replace('.', ',').replace('X', '.' if thou else '')
    return s
def fmt(v, nf):
    if v is None: return ''
    if isinstance(v, bool): return 'Ja' if v else 'Nee'
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%m-%Y') if 'yy' in nf and 'd' not in nf.replace('\\-', '') else v.strftime('%d-%m-%Y')
    if isinstance(v, str): return v
    secs = nf.split(';')
    if nf in ('General', '') or not secs[0]:
        return nl(v, 0) if abs(v - round(v)) < 1e-9 and abs(v) >= 1 or v == 0 else nl(v, 1 if abs(v) >= 10 else 2)
    sec = secs[0] if v > 0 or len(secs) == 1 else (secs[1] if v < 0 else (secs[2] if len(secs) > 2 else secs[0]))
    lit = re.sub(r'\[[^\]]*\]', '', sec)
    if '0' not in lit and '#' not in lit: return lit.replace('\\', '').replace('"', '')
    dec = len(re.search(r'\.(0+)', lit).group(1)) if re.search(r'\.(0+)', lit) else 0
    pct = '%' in lit; x = v * 100 if pct else v
    if ',' in lit and lit.rstrip().endswith(',') and not pct: x = x / 1000
    neg = v < 0 and (len(secs) == 1 or '-' in secs[1] if len(secs) > 1 else True)
    s = nl(x, dec, thou=(',' in lit))
    return ('-' if v < 0 else '') + s + ('%' if pct else '')

CFRULES = {}
def cf_color(sh, r, c, v):
    if sh not in CFRULES:
        rules = []
        for cf in F[sh].conditional_formatting:
            for rg in str(cf.sqref).split():
                a = rg.split(':')[0]; m = re.match(r'([A-Z]+)(\d+)', a); r0, c0 = int(m.group(2)), CI(m.group(1))
                for ru in cf.rules:
                    col = ru.dxf.font.color.rgb if ru.dxf and ru.dxf.font and ru.dxf.font.color and isinstance(ru.dxf.font.color.rgb, str) else None
                    if col and ru.formula: rules.append((rg, r0, c0, ru.formula[0], col))
        CFRULES[sh] = rules
    for rg, r0, c0, f, col in CFRULES[sh]:
        a, b = (rg.split(':') + [rg])[:2]
        m1 = re.match(r'([A-Z]+)(\d+)', a); m2 = re.match(r'([A-Z]+)(\d+)', b)
        if not (int(m1.group(2)) <= r <= int(m2.group(2)) and CI(m1.group(1)) <= c <= CI(m2.group(1))): continue
        m = re.match(r'AND\(ISNUMBER\(\$?[A-Z]+\$?\d+\),\$?[A-Z]+\$?\d+(<|>|=)(-?[\d.]+)\)$', f)
        if m and isinstance(v, (int, float)) and not isinstance(v, bool):
            op, n = m.group(1), float(m.group(2))
            if (op == '<' and v < n) or (op == '>' and v > n) or (op == '=' and v == n): return RGBColor.from_string(col[-6:])
        m = re.match(r'\$?[A-Z]+\$?\d+="(.*)"$', f)
        if m and v == m.group(1): return RGBColor.from_string(col[-6:])
    return None

prs = Presentation(); prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
BLANK = prs.slide_layouts[6]
def box(sld, x, y, w, h, kleur):
    s = sld.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    s.fill.solid(); s.fill.fore_color.rgb = kleur; s.line.fill.background(); s.shadow.inherit = False; return s
def text(sld, x, y, w, h, t, size, kleur, bold=False, align=None):
    tb = sld.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h)); tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(str(t).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.text = line
        if align: p.alignment = align
        for r in p.runs: r.font.name = 'Arial'; r.font.size = Pt(size); r.font.color.rgb = kleur; r.font.bold = bold
    return tb
def add_logo(sld, rechts, boven, hoogte):
    w = hoogte * 236 / 83; sld.shapes.add_picture(logo, Pt(rechts - w), Pt(boven), Pt(w), Pt(hoogte))
def titeldia():
    sld = prs.slides.add_slide(BLANK)
    box(sld, 0, 0, 960, 540, GEEL); box(sld, 0, 0, 470, 540, NAVY)
    text(sld, 48, 150, 390, 18, str(val('P_Regio')).upper(), 11, GEEL, True)
    text(sld, 48, 176, 390, 120, val('PPT_Titel'), 34, WIT, True)
    text(sld, 48, 300, 390, 30, val('PPT_Sub'), 18, GEEL)
    text(sld, 48, 492, 390, 16, 'Gemaakt op ' + datetime.date.today().strftime('%d-%m-%Y') + '  ·  bron: ' + src.split('/')[-1], 9, WIT)
    add_logo(sld, 920, 470, 40)
def nieuwedia(titel, kern):
    sld = prs.slides.add_slide(BLANK)
    box(sld, 0, 0, 960, 8, GEEL)
    text(sld, 36, 22, 880, 14, (titel or '').upper(), 9, NAVY, True)
    text(sld, 36, 38, 880, 30, kern or titel, 20, NAVY, True)
    ln = sld.shapes.add_connector(1, Pt(36), Pt(508), Pt(924), Pt(508)); ln.line.color.rgb = LIJN; ln.line.width = Pt(0.75)
    text(sld, 36, 516, 600, 14, val('PPT_Footer'), 8, GRIJS)
    text(sld, 780, 516, 40, 14, str(len(prs.slides)), 8, NAVY, True, PP_ALIGN.RIGHT)
    add_logo(sld, 924, 512, 18)
    return sld
def kop(sld, t, x, y, w):
    if t: text(sld, x, y - 20, w, 16, t, 11, NAVY, True)
def celtekst(sh, r, c):
    cell = V[sh].cell(r, c); return fmt(cell.value, cell.number_format or 'General').replace('\r', '')
def tegels(sld, nm, x, y, w, h):
    sh, c0, r0, _, _ = rng_cells(nm); tw = (w - 5 * 8) / 6
    for k in range(6):
        tx = x + k * (tw + 8); box(sld, tx, y, tw, h, LICHT); box(sld, tx, y, tw, 3, GEEL)
        text(sld, tx + 6, y + 8, tw - 12, 14, celtekst(sh, r0 + k, c0), 8.5, NAVY, True)
        text(sld, tx + 6, y + 24, tw - 12, 28, celtekst(sh, r0 + k, c0 + 1), 19, NAVY, True)
        text(sld, tx + 6, y + h - 18, tw - 12, 14, celtekst(sh, r0 + k, c0 + 2), 8, GRIJS)
def tabel(sld, nm, x, y, w, tkop):
    sh, c0, r0, c1, r1 = rng_cells(nm); ws = V[sh]
    kol = []; som = 0
    for c in range(c0, c1 + 1):
        g = ws.cell(r0 - 1, c).value
        if isinstance(g, (int, float)) and g > 0: kol.append((c, g)); som += g
    reg = [r for r in range(r0, r1 + 1) if r == r0 or celtekst(sh, r, c0) != '']
    if not kol or not reg: return
    size = 8 if len(reg) > 15 else 9; kop(sld, tkop, x, y, w)
    gs = sld.shapes.add_table(len(reg), len(kol), Pt(x), Pt(y), Pt(w), Pt((size + 6) * len(reg)))
    tbl = gs.table; tbl.first_row = True; tbl.horz_banding = False
    tblpr = gs._element.graphic.graphicData.tbl.tblPr
    for ch in list(tblpr): tblpr.remove(ch)        # geen thema-stijl: we kleuren zelf
    for j, (c, g) in enumerate(kol): tbl.columns[j].width = Pt(max(20, w * g / som))
    for i, r in enumerate(reg):
        tbl.rows[i].height = Pt(size + 6)
        for j, (c, g) in enumerate(kol):
            cell = tbl.cell(i, j); t = celtekst(sh, r, c); src_cell = ws.cell(r, c)
            bol = t in ('●', '○')
            cell.margin_left = cell.margin_right = Pt(3); cell.margin_top = cell.margin_bottom = Pt(1); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.text = t; tf.word_wrap = True
            al = PP_ALIGN.LEFT if j == 0 else (PP_ALIGN.CENTER if bol or i == 0 else PP_ALIGN.RIGHT if not isinstance(src_cell.value, str) else PP_ALIGN.LEFT)
            kleur = NAVY if i else WIT
            if i:
                cc = cf_color(sh, r, c, src_cell.value)
                if cc: kleur = cc
                elif isinstance(src_cell.value, (int, float)) and src_cell.value < 0 and '[RED]' in (src_cell.number_format or '').upper(): kleur = RGBColor(0xD7, 0x37, 0x2F)
            for p in tf.paragraphs:
                p.alignment = al
                for run in p.runs:
                    run.font.name = 'Arial'; run.font.size = Pt((size - 1 if i == 0 else size) + (5 if bol else 0)); run.font.color.rgb = kleur
                    run.font.bold = True if i == 0 else bool(src_cell.font and src_cell.font.b)
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY if i == 0 else WIT
def opsomming(sld, nm, x, y, w, h):
    sh, c0, r0, c1, r1 = rng_cells(nm)
    items = [celtekst(sh, r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1) if celtekst(sh, r, c).strip()]
    if not items: items = ['Geen bijzonderheden.']
    box(sld, x, y, w, h, LICHT)
    tb = sld.shapes.add_textbox(Pt(x + 24), Pt(y + 20), Pt(w - 48), Pt(h - 40)); tf = tb.text_frame; tf.word_wrap = True
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.space_after = Pt(14)
        r = p.add_run(); r.text = '■  '; r.font.color.rgb = GEEL; r.font.size = Pt(16); r.font.name = 'Arial'
        r = p.add_run(); r.text = t; r.font.color.rgb = NAVY; r.font.size = Pt(16); r.font.name = 'Arial'

# ---------------- grafieken uit het Dias-blad (gegevens uit de doorgerekende grafiekcache)
Z = zipfile.ZipFile(calc)
def zread(p): return Z.read(p).decode('utf8')
def rels_of(p):
    d = p.rsplit('/', 1); rp = d[0] + '/_rels/' + d[1] + '.rels'
    t = zread(rp); m = {}
    for e in re.finditer(r'<Relationship [^>]*>', t):
        i = re.search(r'Id="([^"]+)"', e.group(0)).group(1); tg = re.search(r'Target="([^"]+)"', e.group(0)).group(1); m[i] = tg
    return m
def resolve(base, tg):
    parts = base.rsplit('/', 1)[0].split('/')
    for seg in tg.split('/'):
        if seg == '..': parts.pop()
        else: parts.append(seg)
    return '/'.join(parts)
wbx = zread('xl/workbook.xml'); wrel = rels_of('xl/workbook.xml')
rid = re.search(r'<sheet name="Dias" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wbx).group(1)
dias = resolve('xl/workbook.xml', wrel[rid]); drw = resolve(dias, list(rels_of(dias).values())[0])
drels = rels_of(drw); charts = []
for m in re.finditer(r'<xdr:graphicFrame.*?</xdr:graphicFrame>', zread(drw), re.S):
    charts.append(resolve(drw, drels[re.search(r'r:id="([^"]+)"', m.group(0)).group(1)]))
NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
def cache(el):
    if el is None: return []
    pts = {}
    for pt in el.iter('{%s}pt' % NS['c']):
        v = pt.find('c:v', NS); pts[int(pt.get('idx'))] = v.text if v is not None else None
    n = el.find('.//c:ptCount', NS); n = int(n.get('val')) if n is not None else (max(pts) + 1 if pts else 0)
    return [pts.get(i) for i in range(n)]
def chart_info(k):
    x = etree.fromstring(Z.read(charts[k - 1])); series = []
    for plot in x.find('.//c:plotArea', NS):
        tag = etree.QName(plot).localname
        if not tag.endswith('Chart'): continue
        bd = plot.find('c:barDir', NS); gr = plot.find('c:grouping', NS)
        for s in plot.findall('c:ser', NS):
            nm = s.find('c:tx//c:v', NS); fill = s.find('c:spPr/a:solidFill/a:srgbClr', NS); nofill = s.find('c:spPr/a:noFill', NS)
            series.append(dict(plot=tag, dir=bd.get('val') if bd is not None else None, grp=gr.get('val') if gr is not None else None,
                               name=nm.text if nm is not None else '', cat=cache(s.find('c:cat', NS)), val=cache(s.find('c:val', NS)),
                               color=fill.get('val') if fill is not None else None, nofill=nofill is not None))
    return series
PALET = ['111F39', 'FEB900', 'A4B3BB', 'E30613', '3B7A1E', '5B9BD5']
def grafiek(sld, k, x, y, w, h, tkop):
    kop(sld, tkop, x, y, w)
    if k == 9:   # bandbreedte: scenario's per jaar uit het tabblad Bandbreedte
        ws = V['Bandbreedte']; cd = CategoryChartData(); cd.categories = [str(ws.cell(5, c).value) for c in (2, 3, 4)]
        for r in range(6, 11): cd.add_series(str(ws.cell(r, 1).value), [ws.cell(r, c).value if isinstance(ws.cell(r, c).value, (int, float)) else None for c in (2, 3, 4)])
        typ = XL_CHART_TYPE.COLUMN_CLUSTERED; ser = None
    else:
        ser = [s for s in chart_info(k) if s['plot'] == 'barChart' and any(v not in (None, '') for v in s['val'])]
        if not ser: return
        cats = ser[0]['cat'] or [str(i + 1) for i in range(len(ser[0]['val']))]
        cd = CategoryChartData(); cd.categories = [c if c is not None else '' for c in cats]
        for s in ser: cd.add_series(s['name'], [float(v) if v not in (None, '') else None for v in s['val']])
        horiz = ser[0]['dir'] == 'bar'; stacked = ser[0]['grp'] == 'stacked'
        typ = {(False, False): XL_CHART_TYPE.COLUMN_CLUSTERED, (False, True): XL_CHART_TYPE.COLUMN_STACKED,
               (True, False): XL_CHART_TYPE.BAR_CLUSTERED, (True, True): XL_CHART_TYPE.BAR_STACKED}[(horiz, stacked)]
    gf = sld.shapes.add_chart(typ, Pt(x), Pt(y), Pt(w), Pt(h), cd); ch = gf.chart
    ch.font.size = Pt(8); ch.font.name = 'Arial'; ch.font.color.rgb = NAVY
    ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
    for i, s in enumerate(ch.plots[0].series):
        info = ser[i] if ser else None
        if info and (info['nofill'] or info['name'] in ('Basis',)):
            s.format.fill.background(); s.format.line.fill.background(); continue
        s.format.fill.solid(); s.format.fill.fore_color.rgb = RGBColor.from_string((info['color'] if info and info['color'] else PALET[i % len(PALET)])[-6:])
    if ser and ser[0]['name'] == 'Basis':   # waterval: legenda zonder 'Basis'
        ch.has_legend = False
    ch.plots[0].gap_width = 60
    va = ch.value_axis; va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = LIJN; va.tick_labels.font.size = Pt(8)
    va.tick_labels.number_format = '#,##0.0' if k in (1, 2, 3, 4, 5, 7, 9, 10) else '#,##0'; va.tick_labels.number_format_is_linked = False
    ch.category_axis.tick_labels.font.size = Pt(8)

# ---------------- opbouw volgens PPT_Spec
titeldia()
sh, c0, r0, c1, r1 = rng_cells('PPT_Spec'); ws = V[sh]; vorige = None; sld = None
for r in range(r0, r1 + 1):
    row = [ws.cell(r, c).value for c in range(c0, c1 + 1)]
    if row[10] != 1 or row[0] is None: continue
    if row[0] != vorige:
        sld = nieuwedia(row[1] or '', row[2] or ''); vorige = row[0]
    soort, ref = row[3], row[4]; X, Y, B, H, kp = row[5], row[6], row[7], row[8], row[9] or ''
    try:
        if soort == 'kpi': tegels(sld, ref, X, Y, B, H)
        elif soort == 'tabel': tabel(sld, ref, X, Y, B, kp)
        elif soort == 'grafiek': grafiek(sld, int(ref), X, Y, B, H, kp)
        elif soort == 'opsomming': opsomming(sld, ref, X, Y, B, H)
    except Exception as e:
        print('overgeslagen: dia', row[0], soort, ref, e)
prs.save(out); print('dia\'s:', len(prs.slides), '->', out)
