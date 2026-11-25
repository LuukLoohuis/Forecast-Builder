"""Vereenvoudiging: Start, knoppen, Handleiding, Projectlijst (Meetellen), Invoer, AK-dekking (sortering + realisatie),
Projecten, Bandbreedte (rekenvoorbeeld). Draai na patch_versies.py. usage: patch_eenvoud.py in.xlsm out.xlsm"""
import sys, os, re
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
from patch_ak import split, col2n, n2col
from lxml import etree
NSM = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'; N = '{%s}' % NSM

# ------------------------------------------------------------------ hulpfuncties
def clear(s, ref):
    c, r = split(ref)
    if r in s.rows:
        for el in s.rows[r].findall(N+'c'):
            if el.get('r') == ref: s.rows[r].remove(el)
def clear_rows(s, r0, r1, keep_cols=()):
    for r in range(r0, r1+1):
        if r in s.rows:
            for el in s.rows[r].findall(N+'c'):
                if split(el.get('r'))[0] not in keep_cols: s.rows[r].remove(el)
def hide_rows(s, r0, r1):
    for r in range(r0, r1+1):
        row = s.row(r); row.set('hidden', '1')
def child_index(root, tag_order, tag):
    """voeg element in op de juiste plek volgens OOXML-volgorde"""
    pass
ORDER = ['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges',
         'scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting',
         'dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks',
         'customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture',
         'oleObjects','controls','webPublishItems','tableParts','extLst']
def get_or_make(s, tag):
    el = s.root.find(N+tag)
    if el is not None: return el
    el = etree.Element(N+tag); my = ORDER.index(tag)
    for i, ch in enumerate(s.root):
        t = etree.QName(ch).localname
        if t in ORDER and ORDER.index(t) > my: ch.addprevious(el); return el
    s.root.append(el); return el
def add_cf(s, sqref, formula, dxf, prio=1000):
    cf = etree.Element(N+'conditionalFormatting'); cf.set('sqref', sqref)
    ru = etree.SubElement(cf, N+'cfRule'); ru.set('type', 'expression'); ru.set('dxfId', str(dxf)); ru.set('priority', str(prio))
    f = etree.SubElement(ru, N+'formula'); f.text = formula
    existing = s.root.findall(N+'conditionalFormatting')
    if existing: existing[-1].addnext(cf)
    else:
        my = ORDER.index('conditionalFormatting')
        for ch in s.root:
            t = etree.QName(ch).localname
            if t in ORDER and ORDER.index(t) > my: ch.addprevious(cf); break
def add_dv_list(s, sqref, formula):
    dvs = get_or_make(s, 'dataValidations')
    dv = etree.SubElement(dvs, N+'dataValidation'); dv.set('type', 'list'); dv.set('allowBlank', '1'); dv.set('sqref', sqref)
    f = etree.SubElement(dv, N+'formula1'); f.text = formula; dvs.set('count', str(len(dvs)))
def hyperlinks_keep(s, keep):
    hl = s.root.find(N+'hyperlinks')
    if hl is None: return
    for h in list(hl):
        if h.get('ref') not in keep: hl.remove(h)
    if len(hl) == 0: s.root.remove(hl)
def add_hyperlink(s, ref, loc, disp):
    hl = get_or_make(s, 'hyperlinks')
    for h in hl.findall(N+'hyperlink'):
        if h.get('ref') == ref: hl.remove(h)
    h = etree.SubElement(hl, N+'hyperlink'); h.set('ref', ref); h.set('location', loc); h.set('display', disp)
def merges_drop(s, pred):
    mc = s.root.find(N+'mergeCells')
    if mc is None: return
    for m in list(mc):
        a, b = m.get('ref').split(':'); r0 = split(a)[1]; r1 = split(b)[1]
        if pred(r0, r1): mc.remove(m)
    mc.set('count', str(len(mc)))
    if len(mc) == 0: s.root.remove(mc)
def add_merge(s, ref):
    mc = get_or_make(s, 'mergeCells'); m = etree.SubElement(mc, N+'mergeCell'); m.set('ref', ref); mc.set('count', str(len(mc)))
def formula_of(s, ref):
    c, r = split(ref)
    if r not in s.rows: return None
    for el in s.rows[r].findall(N+'c'):
        if el.get('r') == ref:
            f = el.find(N+'f'); return f
    return None

REF = re.compile(r"(?<![A-Za-z_\.'!$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\(A-Za-z_\d])")
def shift_rows(formula, dr):
    out = []
    for i, part in enumerate(formula.split('"')):
        if i % 2: out.append(part); continue
        out.append(REF.sub(lambda m: m.group(1)+m.group(2)+m.group(3)+(m.group(4) if m.group(3) else str(int(m.group(4))+dr)), part))
    return '"'.join(out)
def expand_shared(s, rows):
    """maak gedeelde formules in de opgegeven rijen expliciet"""
    masters = {}
    for r, row in s.rows.items():
        for el in row.findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.get('t') == 'shared' and f.get('ref'):
                masters[f.get('si')] = (split(el.get('r'))[1], f.text)
    for r in rows:
        if r not in s.rows: continue
        for el in s.rows[r].findall(N+'c'):
            f = el.find(N+'f')
            if f is not None and f.get('t') == 'shared':
                r0, txt = masters[f.get('si')]
                f.text = shift_rows(txt, r - r0)
                for a in ('t', 'ref', 'si'):
                    if a in f.attrib: del f.attrib[a]

# ------------------------------------------------------------------ Start
def start(bk):
    s = bk.sheet('Start')
    sty = {k: s.style(k) for k in ['B2','B4','B5','B12','B14','B15','C15','G8']}
    s.set('B2', style=sty['B2'], is_str=True, formula='"Upload "&P_UploadVersie&" "&P_UploadVJ&" · peildatum "&P_PeilTxt&" · ingelezen "&IF(ISNUMBER(P_Importdatum),DAY(P_Importdatum)&"-"&MONTH(P_Importdatum)&"-"&YEAR(P_Importdatum),"")')
    s.set('C5', style=sty['B5'], is_str=True, formula='"Versies: "&'+'&"  ".join([])' if False else
          '"Versies: Budget "&IF(ISNUMBER(Versies!$I$6),"✔","–")&"  Q1 "&IF(ISNUMBER(Versies!$I$7),"✔","–")&"  Q2 "&IF(ISNUMBER(Versies!$I$8),"✔","–")&"  Q3 "&IF(ISNUMBER(Versies!$I$9),"✔","–")&"  Q4 "&IF(ISNUMBER(Versies!$I$10),"✔","–")')
    clear(s, 'E5')
    clear_rows(s, 12, 40); clear_rows(s, 7, 40, keep_cols=('B',))
    for c in ('G7','G8','G9','G10','G11','G13','G14','G15'): clear(s, c)
    s.set('B12', style=sty['B12'], text='Kiest de exports en plakt ze in de gele tabbladen (mag ook zelf plakken).')
    s.set('C12', style=sty['B12'], text='Presentatie in Heijmans-stijl.')
    s.set('B14', style=sty['B14'], text='Zo werkt het')
    steps = [('1. Plakken', 'Gele tabbladen achteraan: ZPMW03, Budget en Q1–Q4 ZPMW03, Woningetalage, AK-percentages, AK-realisatie, Mijlpalen.'),
             ('2. Projectlijst', "Knop 'Nieuwe projecten toevoegen'. Oranje = soort kiezen (Opstal / Deelneming / Buiten scope), deelgebied optioneel."),
             ('3. Invoer', 'Instellingen controleren (peildatum, versie, AK-factor). Kolom Status zegt wat nog ontbreekt.'),
             ('4. Bekijken', 'Dashboard: kies kwartaal en deelgebied. Elk rapport heeft een link naar de berekening.'),
             ('5. Presenteren', "Knop 'Maak PowerPoint'.")]
    merges_drop(s, lambda a, b: a >= 14)
    for i, (a, b) in enumerate(steps):
        r = 15 + i
        s.set(f'B{r}', style=sty['B15'], text=a); s.set(f'C{r}', style=sty['C15'], text=b); add_merge(s, f'C{r}:E{r}')
    s.set('B21', style=sty['B15'], text='Uitleg en begrippen: tabblad Handleiding →')
    hyperlinks_keep(s, {'B5', 'D5'})
    add_hyperlink(s, 'B21', "'Handleiding'!A1", 'Handleiding'); add_hyperlink(s, 'C5', "'Versies'!A1", 'Versies')
    # knoppen: Vastleggen, Leegmaken, Oude prognoses weg; PowerPoint en Toon één kolom naar links
    dfn = [k for k in bk.data if k.startswith('xl/drawings/drawing') and k.endswith('.xml')]
    for fn in dfn:
        d = bk.text(fn)
        if 'RDFO_BronnenInlezen' not in d or 'RDFO_PrognoseVastleggen' not in d: continue
        anchors = re.findall(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', d, re.S)
        for a in anchors:
            if any(m in a for m in ('RDFO_PrognoseVastleggen', 'RDFO_LeegMaken', 'RDFO_OudePrognosesOverzetten', 'RDFO_ToonRekenbladen')):
                d = d.replace(a, '')
            elif 'RDFO_MaakPowerPoint' in a:
                b = re.sub(r'(<xdr:(?:from|to)><xdr:col>)(\d+)', lambda m: m.group(1)+str(int(m.group(2))-1), a)
                d = d.replace(a, b)
        bk.settext(fn, d)

# ------------------------------------------------------------------ Handleiding
def handleiding(bk):
    s = bk.sheet('Handleiding')
    st = dict(title=s.style('B1'), text=s.style('B2'), sec=s.style('B5'), th=s.style('C12'), td=s.style('D13') or s.style('B2'), tl=s.style('C13') or s.style('B2'))
    sd = s.root.find(N+'sheetData')
    for r in list(sd): sd.remove(r)
    s.rows = {}
    for tag in ('mergeCells', 'conditionalFormatting', 'hyperlinks', 'dataValidations'):
        for el in s.root.findall(N+tag): s.root.remove(el)
    cols = s.root.find(N+'cols')
    for c in list(cols): cols.remove(c)
    for mn, mx, w in ((1,1,2),(2,2,24),(3,3,95)):
        e = etree.SubElement(cols, N+'col'); e.set('min',str(mn)); e.set('max',str(mx)); e.set('width',str(w)); e.set('customWidth','1')
    rows = []
    def sec(t): rows.append(('sec', t, None))
    def item(a, b): rows.append(('item', a, b))
    def gap(): rows.append(('gap', None, None))
    rows.append(('title', 'Handleiding', None)); rows.append(('text', 'Alles wat je nodig hebt op één plek. Gele cel = invullen. n.b. = niet beschikbaar.', None)); gap()
    sec('Elk kwartaal')
    item('1. Plakken', "Plak de exports in de gele tabbladen achteraan, telkens vanaf A1: Bron_ZPMW03 (huidig), Budget/Q1–Q4 ZPMW03 (versies), Bron_Woningetalage, Bron_AK (AK-percentages), Bron_AKR (AK-realisatie, grootboek) en Bron_Mijlpalen. Of gebruik Start › 'Bronnen inlezen'. Alles rekent meteen.")
    item('2. Projectlijst', "Knop 'Nieuwe projecten toevoegen' zet nieuwe WBS-elementen uit Bron_ZPMW03 onderaan. Oranje WBS = soort nog kiezen: Opstal, Deelneming of Buiten scope (telt nergens mee). Deelgebied kiezen is optioneel; leeg = geen deelgebied.")
    item('3. Invoer', 'Rapportagejaar, peildatum, versie, eenheid en AK-factor (standaard 1,06). De kolom Status geeft aan wat nog niet is ingevuld. Deelgebieden maak je hier aan (blok 3).')
    item('4. Bekijken', 'Dashboard: kies doeljaar, kwartaal (vergelijking met het vorige kwartaal), type, project en deelgebied. Het deelgebiedfilter geldt voor alle rapporten. Elk rapport heeft bovenaan een link naar de berekening.')
    item('5. Presenteren', "Start › 'Maak PowerPoint' (Heijmans-stijl, met logo).")
    gap(); sec('Versies en budget')
    item('Budget', 'De prognose die vorig jaar voor dit jaar is afgegeven, bijv. de export van Q3 2025 voor budget 2026. Plak die in Budget ZPMW03.')
    item('Q1 t/m Q4', 'De ZPMW03-export van dat kwartaal. Peildatum leeg = einde kwartaal; anders invullen op Versies.')
    item('Woningen', 'Alleen uit de huidige Woningetalage (één dump). Versies tonen omzet, BOR en AK-dekking.')
    item('Officiële AK', "Wijkt het gerapporteerde AK-bedrag af? Vul het in op Versies, kolom 'AK-dekking officieel'.")
    item('Mutatie', 'Projecten: verschil tussen het gekozen kwartaal en het vorige (Q3 t.o.v. Q2, Q1 t.o.v. Budget). Groen = hoger, rood = lager. Typ de uitleg in de gele kolom Toelichting mutatie. Niet sorteren, alleen filteren.')
    gap(); sec('Berekeningen (te volgen via de links)')
    item('Omzet FY', '− Bedrijfsopbrengsten LJ − OP.BOB van het jaar (volgend jaar: − OP.BOB volgend jaar). Per WBS opgeteld. Rekenblad: STG_ZPMW03 › VERW_Project; versies: VERSIE_REKEN.')
    item('BOR FY', '− Bruto (omzet) result. LJ − OP.BOR van het jaar.')
    item('AK-dekking', 'Kosten per kwartaal / factor (Invoer, standaard 1,06) x AK %. AK % uit Bron_AK (eigen AK % op de Projectlijst gaat voor). Geboekte AK uit Bron_AKR telt t/m de peildatum. Per project uitgeschreven op AK-dekking; kolom Bron AK % linkt naar de regel in Bron_AK.')
    item('AK-realisatie', 'AK-dekking onderaan: geboekt per jaar en kwartaal, en per kwartaal tegen het budget (onder- of overdekking).')
    item('Bandbreedte', "Vijf scenario's per jaar: basis, afslag op verkopen, optimistisch, gewogen met kans % en 0-100. Rekenvoorbeeld per project vanaf kolom AI; rekenblad SCEN.")
    item('Reserves', 'Zelf invullen op de Projectlijst (L–N) met % gereed (P–Q). Vrijval = reserve x % gereed.')
    item("Kansen en risico's", 'Zelf invullen per regel (project, jaar, kwartaal, bedrag positief). Risico telt als min. Saldo per jaar in Outlook en PowerPoint.')
    item('Mijlpalen', 'Projectdetail onderaan: per mijlpaal verwacht en werkelijk uit Bron_Mijlpalen, met status (gerealiseerd, later dan verwacht, verlopen, gepland).')
    item('Verlies', 'Projecten: BOR FY van de gekozen versie is negatief.')
    gap(); sec('Tabbladen')
    for a, b in [('Rapporten', 'Start, Dashboard, RDFO, Projecten, Projectdetail, Woningetalage, Outlook, Bandbreedte, AK-dekking.'),
                 ('Invoer', "Projectlijst, Invoer, Kansen en risico's, Reserves, Versies."),
                 ('Gele tabbladen', 'Plakbladen voor de exports.'),
                 ('Grijze tabbladen', 'Rekenbladen: alleen lezen, niets wijzigen. Te bereiken via de links op de rapporten.'),
                 ('Controle', 'Alleen nodig bij ✖ op Start; kolom H zegt wat te doen.')]:
        item(a, b)
    gap(); sec('Vragen')
    item('Versie toont n.b.', "Plakblad leeg, of de peildatum hoort bij een ander jaar. Zie Versies, kolom Status.")
    item('Nieuw project ontbreekt', "Projectlijst › 'Nieuwe projecten toevoegen' (leest Bron_ZPMW03), of typ de WBS onderaan de lijst.")
    item("Geen macro's", 'Alles werkt met plakken. Alleen nieuwe projecten toevoegen en PowerPoint maken gebruiken een macro.')
    r = 1
    for kind, a, b in rows:
        if kind == 'title': s.set(f'B{r}', text=a, style=st['title'])
        elif kind == 'text': s.set(f'B{r}', text=a, style=st['text'])
        elif kind == 'sec': s.set(f'B{r}', text=a, style=st['sec'])
        elif kind == 'item': s.set(f'B{r}', text=a, style=st['tl']); s.set(f'C{r}', text=b, style=st['td'])
        r += 1
    s.root.find(N+'dimension').set('ref', f'A1:C{r}')

# ------------------------------------------------------------------ Projectlijst + Invoer
def projectlijst(bk, dxf):
    s = bk.sheet('Projectlijst')
    s.set('C7', text='Soort', style=s.style('C7'))
    s.set('D7', text='Deelgebied\n(leeg = geen)', style=s.style('D7'))
    for dv in s.root.iter(N+'dataValidation'):
        if dv.get('sqref') == 'C8:C107':
            dv.find(N+'formula1').text = '"Opstal,Deelneming,Buiten scope"'
    s.set('AN6', text='FILTER', style=s.style('AG6'))
    s.set('AN7', text='Telt mee\n(soort + filter)', style=s.style('AG7'))
    for r in range(8, 108):
        s.set(f'AN{r}', is_str=True, style=s.style('O8'), formula=
              f'IF($A{r}="","",IF(OR($E{r}="Nee",$C{r}="Buiten scope"),"Nee",IF(OR(Sel_Deelgebied="Alle",Sel_Deelgebied="",AND(Sel_Deelgebied="(geen)",$D{r}=""),$D{r}=Sel_Deelgebied),"Ja","Nee")))')
    bk.set_name('PL_Telt', 'Projectlijst!$AN$8:$AN$107')
    cols = s.root.find(N+'cols')
    new = []
    for c in list(cols):
        mn, mx = int(c.get('min')), int(c.get('max'))
        if mn <= 8 and mx >= 5:   # splits rond E:H
            for a, b, hid in ((mn, min(mx, 4), False), (max(mn, 5), min(mx, 8), True), (max(mn, 9), mx, False)):
                if a <= b:
                    e = etree.Element(N+'col'); [e.set(k, v) for k, v in c.attrib.items()]; e.set('min', str(a)); e.set('max', str(b))
                    if hid: e.set('hidden', '1')
                    new.append(e)
            cols.remove(c)
    if not any(int(c.get('min')) <= 6 <= int(c.get('max')) for c in cols) and not new:
        e = etree.Element(N+'col'); e.set('min','5'); e.set('max','8'); e.set('width','9'); e.set('hidden','1'); e.set('customWidth','1'); new.append(e)
    for e in new: cols.append(e)
    items = sorted(cols, key=lambda c: int(c.get('min')))
    for c in list(cols): cols.remove(c)
    for c in items: cols.append(c)
    s.set('B2', style=s.style('B2'), text="Eén regel per project, gevuld vanuit ZPMW03 (knop 'Nieuwe projecten toevoegen'). Kies per project de soort: Opstal, Deelneming of Buiten scope. Deelgebied optioneel. Oranje WBS = soort nog kiezen.")
    s.set('B4', style=s.style('B4'), is_str=True, formula=
          'IF(COUNTIFS($A$8:$A$107,"?*",$C$8:$C$107,"")=0,"✔ alle projecten hebben een soort","⚠ "&COUNTIFS($A$8:$A$107,"?*",$C$8:$C$107,"")&" project(en): soort kiezen (oranje)")'
          '&"   ·   nieuw in ZPMW03, nog niet op de lijst: "&SUMPRODUCT((STG_ZPMW03!$B$6:$B$505<>"")*(COUNTIF($A$8:$A$107,STG_ZPMW03!$B$6:$B$505)=0)/COUNTIF(STG_ZPMW03!$B$6:$B$505,STG_ZPMW03!$B$6:$B$505&""))')
    add_cf(s, 'A8:A107', 'AND($A8<>"",$C8="")', dxf, prio=1)
    # AK % gebruikt: zoeken op WBS i.p.v. rij
    for r in range(8, 108):
        s.set(f'K{r}', style=s.style(f'K{r}'), formula=f'IFERROR(IF(INDEX(\'AK-dekking\'!$D$21:$D$120,MATCH($A{r},\'AK-dekking\'!$A$21:$A$120,0))="","",INDEX(\'AK-dekking\'!$D$21:$D$120,MATCH($A{r},\'AK-dekking\'!$A$21:$A$120,0))),"")')
    inv = bk.sheet('Invoer')
    expand_shared(inv, range(31, 131))
    for i in range(100):
        r = 31 + i; p = 8 + i
        inv.set(f'G{r}', style=inv.style(f'G{r}'), is_str=True, formula=f'IF(A{r}="","",IF(OR(Projectlijst!$AN${p}<>"Ja",Projectlijst!$F${p}="Nee"),"Nee","Ja"))')
        inv.set(f'U{r}', style=inv.style(f'U{r}'), is_str=True, formula=f'IF(A{r}="","",IF(OR(Projectlijst!$AN${p}<>"Ja",Projectlijst!$G${p}="Nee"),"Nee","Ja"))')
    hide_rows(inv, 28, 132)
    inv.set('E3', style=inv.style('E3'), text='2. Projecten: zie Projectlijst')
    add_hyperlink(inv, 'E3', "'Projectlijst'!A1", 'Projectlijst')

# ------------------------------------------------------------------ AK-dekking: meetellende projecten bovenaan + realisatie
XREF = re.compile(r"(Invoer|Projectlijst|VERW_Project)!(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d:])")
BASE = {'Invoer': (31, 130), 'Projectlijst': (8, 107), 'VERW_Project': (6, 105)}
def akdekking(bk):
    s = bk.sheet('AK-dekking')
    expand_shared(s, range(21, 121))
    for r in range(21, 121):
        i = r - 20
        for el in s.rows[r].findall(N+'c'):
            f = el.find(N+'f')
            if f is None or f.text is None or split(el.get('r'))[0] in ('AW', 'AX'): continue
            def rep(m):
                sh, d1, col, d2, row = m.groups(); b0, b1 = BASE[sh]
                if int(row) != b0 + i - 1: return m.group(0)
                return f'INDEX({sh}!${col}${b0}:${col}${b1},$AW{r})'
            parts = f.text.split('"')
            for k in range(0, len(parts), 2): parts[k] = XREF.sub(rep, parts[k])
            f.text = '"'.join(parts)
        s.set(f'AX{r}', formula=f'IF(Invoer!$A{30+i}="",3000,IF(OR(Projectlijst!$AN{7+i}<>"Ja",Projectlijst!$G{7+i}="Nee"),2000,IF(N(IF(Projectlijst!$I{7+i}<>"",Projectlijst!$I{7+i},Invoer!$V{30+i}))>0,0,1000)))+{i}')
        s.set(f'AW{r}', formula=f'MOD(SMALL($AX$21:$AX$120,{i}),1000)')
    s.set('AW20', text='rij-idx'); s.set('AX20', text='sorteersleutel')
    cols = s.root.find(N+'cols'); e = etree.SubElement(cols, N+'col'); e.set('min', '49'); e.set('max', '50'); e.set('width', '8'); e.set('hidden', '1'); e.set('customWidth', '1')
    s.set('B18', style=s.style('B18'), text='3. Per project (projecten die meetellen met AK % > 0 staan bovenaan)')
    # VERW_Project EE..EJ: zoeken op WBS
    v = bk.sheet('VERW_Project'); expand_shared(v, range(6, 106))
    for col, src in (('EE', 'Q'), ('EF', 'J'), ('EG', 'S'), ('EH', 'U'), ('EI', 'W'), ('EJ', 'Y')):
        for r in range(6, 106):
            m = f"MATCH($A{r},'AK-dekking'!$A$21:$A$120,0)"
            v.set(f'{col}{r}', style=v.style(f'{col}{r}'), formula=f"IFERROR(IF(INDEX('AK-dekking'!${src}$21:${src}$120,{m})=\"\",\"\",INDEX('AK-dekking'!${src}$21:${src}$120,{m})),\"\")")
    # ---- 4. AK-realisatie per jaar en kwartaal
    T, H, NUM, TOT, LBL = s.style('B4'), s.style('F5'), s.style('F7'), s.style('F9'), s.style('B7')
    r0 = 125
    s.set(f'B{r0}', text='4. AK-realisatie: geboekt per jaar en kwartaal (grootboek, Bron_AKR)', style=T)
    for j, h in enumerate(['Q1', 'Q2', 'Q3', 'Q4', 'Jaar']): s.set(f'{"FGHIJ"[j]}{r0+1}', text=h, style=H)
    for k in range(3):
        r = r0 + 2 + k
        s.set(f'B{r}', is_str=True, formula=f'"Geboekt "&(P_RJ-{2-k})', style=LBL)
        for q in range(4):
            s.set(f'{"FGHI"[q]}{r}', style=NUM, formula=f'IF(COUNTIFS(AKR_Jaar,P_RJ-{2-k})=0,"–",SUMIFS(AKR_Bedrag,AKR_Jaar,P_RJ-{2-k},AKR_Kw,{q+1}))')
        s.set(f'J{r}', style=TOT, formula=f'IF(COUNTIFS(AKR_Jaar,P_RJ-{2-k})=0,"–",SUM(F{r}:I{r}))')
    r1 = r0 + 6
    s.set(f'B{r1}', is_str=True, formula='"5. Onder- of overdekking "&P_RJ&" per kwartaal (geboekt tegen budget)"', style=T)
    for j, h in enumerate(['Q1', 'Q2', 'Q3', 'Q4', 'Jaar']): s.set(f'{"FGHIJ"[j]}{r1+1}', text=h, style=H)
    lab = ['Budget AK (Budget-export: kosten kwartaal / factor x AK %)', 'Geboekt (grootboek)', 'Verschil geboekt − budget', 'Oordeel', 'Prognose AK nu (sectie 1)']
    for k, t in enumerate(lab): s.set(f'B{r1+2+k}', text=t, style=LBL)
    kq = ['AO', 'AP', 'AQ', 'AR']
    for q in range(4):
        c = 'FGHI'[q]
        s.set(f'{c}{r1+2}', style=NUM, formula=f'IF(NOT(ISNUMBER(Versies!$I$6)),"n.b.",SUMPRODUCT(VERSIE_REKEN!${kq[q]}$6:${kq[q]}$155,VERSIE_REKEN!$BL$6:$BL$155)/P_AKFactor)')
        s.set(f'{c}{r1+3}', style=NUM, formula=f'IF(OR(AKR_N=0,{q+1}>AK_PeilQ),"–",SUMIFS(AKR_Bedrag,AKR_Jaar,P_RJ,AKR_Kw,{q+1}))')
        s.set(f'{c}{r1+4}', style=NUM, formula=f'IF(AND(ISNUMBER({c}{r1+2}),ISNUMBER({c}{r1+3})),{c}{r1+3}-{c}{r1+2},"")')
        s.set(f'{c}{r1+5}', style=LBL, is_str=True, formula=f'IF(ISNUMBER({c}{r1+4}),IF({c}{r1+4}<0,"onderdekking","overdekking"),"")')
        s.set(f'{c}{r1+6}', style=NUM, formula=f'{c}9')
    s.set(f'J{r1+2}', style=TOT, formula=f'IF(ISNUMBER(F{r1+2}),SUM(F{r1+2}:I{r1+2}),"n.b.")')
    s.set(f'J{r1+3}', style=TOT, formula=f'IF(AKR_N=0,"n.b.",SUM(F{r1+3}:I{r1+3}))')
    s.set(f'J{r1+4}', style=TOT, formula=f'IF(AND(ISNUMBER(J{r1+2}),ISNUMBER(J{r1+3})),J{r1+3}-SUMPRODUCT(F{r1+2}:I{r1+2},--(COLUMN(F{r1+2}:I{r1+2})-COLUMN(F{r1+2})+1<=AK_PeilQ)),"")')
    s.set(f'K{r1+4}', is_str=True, formula='"t/m Q"&AK_PeilQ', style=LBL)
    s.set(f'J{r1+6}', style=TOT, formula='J9')
    s.set(f'B{r1+8}', text='Geboekt = boekingen in het grootboek per kwartaal (periode / 3). Budget = per kwartaal uit de Budget-export (tabblad Budget ZPMW03). Negatief verschil = minder AK geboekt dan begroot.', style=s.style('F18') or LBL)
    s.root.find(N+'dimension').set('ref', f'A1:AX{r1+8}')

# ------------------------------------------------------------------ Projecten
def projecten(bk):
    s = bk.sheet('Projecten')
    s.set('J9', style=s.style('J9'), is_str=True, formula='"Nu"&CHAR(10)&"ingelezen"')
    s.set('R9', style=s.style('R9'), text='Verlies\n(BOR < 0)')
    s.set('B7', style=s.style('B7'), text='Budget, Q1–Q4 = versies (tabblad Versies). Nu ingelezen = laatste ZPMW03 (peildatum). Geselecteerd = versie gekozen op het Dashboard. Verlies = BOR FY van de geselecteerde versie is negatief. Marge % = BOR / omzet.')

# ------------------------------------------------------------------ Bandbreedte: rekenvoorbeeld
def bandbreedte(bk):
    s = bk.sheet('Bandbreedte')
    T, H, LBL, NUM, PCT, INP, TXT = s.style('A24'), s.style('A25'), s.style('A15'), s.style('O26'), bk.sheet('Projecten').style('O10'), s.style('D15'), s.style('A15')
    inp = bk.sheet('Invoer').style('D14')
    s.set('AI24', text='REKENVOORBEELD: zo komt het scenario tot stand', style=T)
    s.set('AI25', text='Project (kies)', style=H); s.set('AJ25', style=inp)
    add_dv_list(s, 'AJ25', '$A$26:$A$125')
    s.set('AK25', is_str=True, style=TXT, formula='IF(AJ25<>"",AJ25,IFERROR(INDEX($A$26:$A$125,MATCH(MAX($O$26:$O$125),$O$26:$O$125,0)),""))')
    s.set('AL25', formula='IFERROR(MATCH(AK25,$A$26:$A$125,0),0)')
    s.set('AI26', text='Jaar', style=LBL); s.set('AJ26', formula='SC_Jaar', style=NUM)
    o = 'SC_Jaar-P_RJ+1'
    def sc(c0, c1, c2): return f'IFERROR(CHOOSE({o},INDEX(SCEN!${c0}$6:${c0}$105,$AL$25),INDEX(SCEN!${c1}$6:${c1}$105,$AL$25),INDEX(SCEN!${c2}$6:${c2}$105,$AL$25)),"n.b.")'
    lines = [
     ('Basis omzet (prognose)', f'{sc("F","O","X")}/P_Factor', NUM, 'omzet uit de upload'),
     ('Al gerealiseerd (t/m peildatum)', 'IF(SC_Jaar=P_RJ,N(INDEX(VERW_Project!$W$6:$W$105,$AL$25))/P_Factor,0)', NUM, 'telt in elk scenario volledig mee'),
     ('Totaal woningen', 'IFERROR(INDEX(SCEN!$E$6:$E$105,$AL$25),"n.b.")', NUM, 'Woningetalage'),
     ('Verkocht vóór dit jaar', sc('G', 'P', 'Y'), NUM, ''),
     ('Gepland te verkopen dit jaar', sc('H', 'Q', 'Z'), NUM, ''),
     ('Verkocht % eind jaar (plan)', sc('I', 'R', 'AA'), PCT, '= (vóór + gepland) / totaal'),
     ('Verkocht % bij afslag', sc('J', 'S', 'AB'), PCT, '= (vóór + afslag x gepland) / totaal; afslag = '),
     ('Afslag-omzet', f'{sc("L","U","AD")}/P_Factor', NUM, '= gerealiseerd + (basis − gerealiseerd) x (verkocht % afslag / verkocht % plan)'),
     ('Optimistisch', f'{sc("K","T","AC")}/P_Factor', NUM, '= gerealiseerd + (basis − gerealiseerd) x min(max, optimistisch % / verkocht % plan), alleen als plan < optimistisch %'),
     ('Kans %', 'IFERROR(INDEX(SCEN!$B$6:$B$105,$AL$25),"n.b.")', PCT, 'Projectlijst (leeg = bron)'),
     ('Gewogen', f'{sc("M","V","AE")}/P_Factor', NUM, '= basis x kans %'),
     ('0-100', f'{sc("N","W","AF")}/P_Factor', NUM, '= basis als kans % ≥ grens (of 0-100 = Ja), anders 0'),
    ]
    for k, (a, f, st, uitleg) in enumerate(lines):
        r = 27 + k
        s.set(f'AI{r}', text=a, style=LBL); s.set(f'AJ{r}', formula=f'IF($AL$25=0,"",{f})', style=st)
        if 'afslag = ' in uitleg:
            s.set(f'AK{r}', is_str=True, style=TXT, formula=f'"{uitleg}"&TEXT(SC_Afslag,"0%")')
        elif uitleg: s.set(f'AK{r}', text=uitleg, style=TXT)
    s.set('AI40', is_str=True, style=TXT, formula='"Instellingen: afslag "&TEXT(SC_Afslag,"0%")&" · optimistisch "&TEXT(SC_Opt,"0%")&" verkocht, max "&SC_Max&" x · 0-100 vanaf kans "&TEXT(SC_Grens,"0%")&" (rij 15-18)."')
    s.set('A21', text='Rekenvoorbeeld per project: kolom AI →', style=TXT); add_hyperlink(s, 'A21', "'Bandbreedte'!AI24", 'Rekenvoorbeeld')
    cols = s.root.find(N+'cols')
    for mn, mx, w in ((34, 34, 3), (35, 35, 34), (36, 36, 14), (37, 37, 70), (38, 38, 6)):
        e = etree.SubElement(cols, N+'col'); e.set('min', str(mn)); e.set('max', str(mx)); e.set('width', str(w)); e.set('customWidth', '1')
        if mn == 38: e.set('hidden', '1')
    dim = s.root.find(N+'dimension'); a, b = dim.get('ref').split(':'); c, r = split(b)
    dim.set('ref', f'A1:{"AL" if col2n(c) < col2n("AL") else c}{max(r, 40)}')

def add_dxf(bk, rgb='FFFFC000'):
    st = bk.text('xl/styles.xml')
    m = re.search(r'<dxfs count="(\d+)"', st); n = int(m.group(1))
    dxf = f'<dxf><fill><patternFill><bgColor rgb="{rgb}"/></patternFill></fill></dxf>'
    st = st.replace(m.group(0), f'<dxfs count="{n+1}"', 1).replace('</dxfs>', dxf + '</dxfs>', 1)
    bk.settext('xl/styles.xml', st); return n

def rename_woningen(bk):
    for fn in list(bk.data):
        if fn.endswith('.xml') or fn.endswith('.rels'):
            t = bk.data[fn]
            if b'Woningen' not in t: continue
            d = t.decode()
            d2 = re.sub(r"(?<![A-Za-z_])'?Woningen'?!", 'Woningetalage!', d)
            if fn == 'xl/workbook.xml': d2 = d2.replace('<sheet name="Woningen"', '<sheet name="Woningetalage"')
            if fn == 'docProps/app.xml': d2 = d2.replace('<vt:lpstr>Woningen</vt:lpstr>', '<vt:lpstr>Woningetalage</vt:lpstr>')
            if d2 != d: bk.data[fn] = d2.encode()
    bk.paths['Woningetalage'] = bk.paths.pop('Woningen')

if __name__ == '__main__':
    bk = Book(sys.argv[1])
    dxf = add_dxf(bk)
    start(bk); handleiding(bk); projectlijst(bk, dxf); akdekking(bk); projecten(bk); bandbreedte(bk)
    for p, s in bk.sheets.items(): bk.data[p] = s.xml()
    bk.sheets = {}
    rename_woningen(bk)
    bk.save(sys.argv[2])
