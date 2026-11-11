"""Zet in het RDFO-sjabloon:
  1. op AK-dekking een blok 'Over- of onderdekking' (per kwartaal en cumulatief) met grafiek;
  2. een nieuw tabblad 'Vergelijk' (twee prognoses naast elkaar per project en kwartaal).

Gebruik:  python3 tools/bouw_ak_vergelijk.py IN.xlsm UIT.xlsm
Het script werkt rechtstreeks op de XML in het bestand, zodat knoppen, grafieken en macro's blijven staan.
Eenmalig: draait niet nog eens op een bestand dat het tabblad Vergelijk al heeft.
"""
import re
import sys
import zipfile
from xml.sax.saxutils import escape

SRC, DST = sys.argv[1], sys.argv[2]
AK = "'AK-dekking'!"
INS_AT, INS_N = 11, 31          # 31 nieuwe rijen vanaf rij 11 op AK-dekking

# ---------------------------------------------------------------- verwijzingen verschuiven
REF = re.compile(
    r'''"(?:[^"]|"")*"'''
    r'''|(?<![A-Za-z0-9_.$!'])(?P<sh>'(?:[^']|'')*'!|[A-Za-z_][A-Za-z0-9_.]*!)?'''
    r'''(?P<a>\$?[A-Z]{1,3}\$?[0-9]+)(?P<b>:\$?[A-Z]{1,3}\$?[0-9]+)?(?![A-Za-z0-9_(.!])''')
CELL = re.compile(r'(\$?[A-Z]{1,3}\$?)([0-9]+)')


def _cel(ref):
    m = CELL.fullmatch(ref)
    rij = int(m.group(2))
    return m.group(1) + str(rij + INS_N if rij >= INS_AT else rij)


def schuif(formule, eigen_blad):
    """Verhoogt rijnummers >= INS_AT in verwijzingen naar AK-dekking."""
    def r(m):
        if m.group('a') is None:
            return m.group(0)                       # tekst tussen aanhalingstekens
        sh = m.group('sh')
        if not (sh == AK or (sh is None and eigen_blad)):
            return m.group(0)
        b = m.group('b')
        return (sh or '') + _cel(m.group('a')) + ((':' + _cel(b[1:])) if b else '')
    return REF.sub(r, formule)


def schuif_sqref(sq):
    return ' '.join(':'.join(_cel(p) for p in deel.split(':')) for deel in sq.split())


def formules(xml, eigen_blad, tags=('f', 'formula', 'formula1', 'formula2')):
    for t in tags:
        xml = re.sub(r'(<%s(?: [^>]*)?>)([^<]*)(</%s>)' % (t, t),
                     lambda m: m.group(1) + schuif(m.group(2), eigen_blad) + m.group(3), xml)
    return xml


# ---------------------------------------------------------------- cellen schrijven
def c(ref, s, v=None, f=None, tekst=False):
    if f is not None:
        return '<c r="%s" s="%d"%s><f>%s</f></c>' % (ref, s, ' t="str"' if tekst else '', escape(f))
    if v is None:
        return '<c r="%s" s="%d"/>' % (ref, s)
    if isinstance(v, str):
        return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(v))
    return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, v)


def rij(nr, cellen, ht=None):
    return '<row r="%d"%s>%s</row>' % (nr, ' ht="%s" customHeight="1"' % ht if ht else '', ''.join(cellen))


def kol(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- 1. AK-dekking
def ak_blad(x):
    kop, rest = x.split('<sheetData>', 1)
    data, staart = rest.split('</sheetData>', 1)

    # bestaande rijen opschuiven
    data = re.sub(r'<row r="(\d+)"', lambda m: '<row r="%d"' % (int(m.group(1)) + (INS_N if int(m.group(1)) >= INS_AT else 0)), data)
    data = re.sub(r'<c r="([A-Z]+)(\d+)"', lambda m: '<c r="%s%d"' % (m.group(1), int(m.group(2)) + (INS_N if int(m.group(2)) >= INS_AT else 0)), data)
    data = formules(data, True)
    data = data.replace('AK projecten (zie tabel 3)', 'AK projecten (zie tabel 4)')
    data = data.replace('<t>2. Rekenvoorbeeld voor project:</t>', '<t>3. Rekenvoorbeeld voor project:</t>')
    data = data.replace('<t>3. Per project</t>', '<t>4. Per project</t>')

    # nieuw blok 2: over- of onderdekking (rijen 11..41)
    r0 = 13                                           # eerste periode
    P1, PN = 52, 151                                  # projectregels na het opschuiven
    perioden = [('P_RJ&" Q%d"' % q, '$%s$6' % 'FGHI'[q - 1], '$%s$9' % 'FGHI'[q - 1]) for q in range(1, 5)]
    perioden += [('(P_RJ+1)&" Q%d"' % q, '"prognose"', 'SUM($%s$%d:$%s$%d)' % (k, P1, k, PN)) for q, k in zip(range(1, 5), ('AE', 'AF', 'AG', 'AH'))]
    perioden += [('"Jaar "&(P_RJ+%d)' % j, '"prognose"', 'SUM($%s$%d:$%s$%d)' % (k, P1, k, PN)) for j, k in zip((2, 3, 4), ('U', 'W', 'Y'))]
    nieuw = [rij(11, [c('B11', 4, '2. Over- of onderdekking per kwartaal'),
                      c('G11', 32, 'AK-dekking uit de projecten tegenover de AK-kosten van de regio. Vul de AK-kosten in de gele cellen in.')])]
    koppen = ['Periode', '', '', 'Bron', 'AK-dekking', 'AK-kosten regio\n(invullen)', 'Over (+) /\nonder (−)',
              'Dekking\ncumulatief', 'Kosten\ncumulatief', 'Over / onder\ncumulatief', 'Status']
    nieuw.append(rij(12, [c('%s12' % k, 39 if k == 'B' else 40, t) for k, t in zip('BCDEFGHIJKL', koppen)], ht=34))
    for i, (label, bron, dekking) in enumerate(perioden):
        r = r0 + i
        jaarstart = i in (0, 4) or i >= 8             # cumulatief begint opnieuw per kalenderjaar
        eerste = r if jaarstart else (r0 if i < 4 else r0 + 4)
        cum_i = '=F%d' % r if jaarstart else '=I%d+F%d' % (r - 1, r)
        cum_j = 'IF(COUNT($G$%d:G%d)<ROWS($G$%d:G%d),"n.b.",SUM($G$%d:G%d))' % (eerste, r, eerste, r, eerste, r)
        nieuw.append(rij(r, [
            c('B%d' % r, 79, f=label, tekst=True), c('C%d' % r, 79), c('D%d' % r, 79),
            c('E%d' % r, 101, f=bron, tekst=True),
            c('F%d' % r, 71, f=dekking),
            c('G%d' % r, 144),
            c('H%d' % r, 98, f='IF(G%d="","n.b.",F%d-G%d)' % (r, r, r)),
            c('I%d' % r, 150, f=cum_i[1:]),
            c('J%d' % r, 150, f=cum_j),
            c('K%d' % r, 98, f='IF(ISNUMBER(J%d),I%d-J%d,"n.b.")' % (r, r, r)),
            c('L%d' % r, 79, f='IF(ISNUMBER(K%d),IF(K%d>0,"▲ over",IF(K%d<0,"▼ onder","● gelijk")),"")' % (r, r, r), tekst=True)]))
    nieuw.append(rij(24, [c('B24', 32, 'Cumulatief telt per kalenderjaar op; latere jaren staan als jaarbedrag. Groen = overdekking '
                                       '(dekking hoger dan kosten), rood = onderdekking. Zonder ingevulde AK-kosten staat er n.b.')]))
    pos = data.index('<row r="%d"' % (INS_AT + INS_N))
    data = data[:pos] + ''.join(nieuw) + data[pos:]

    # per project: AK-dekking volgend jaar per kwartaal (kolommen AE..AH, naast de bestaande hulpkolommen)
    def extra(m):
        r = int(m.group(1))
        if r == 50:
            cel = c('AE50', 132, 'VOLGEND JAAR PER KWARTAAL') + ''.join(c('%s50' % k, 133) for k in ('AF', 'AG', 'AH'))
        elif r == 51:
            cel = ''.join(c('%s51' % k, 40, f='"AK"&CHAR(10)&(P_RJ+1)&" Q%d"' % q, tekst=True) for q, k in enumerate(('AE', 'AF', 'AG', 'AH'), 1))
        elif P1 <= r <= PN:
            cel = ''
            for k, v in zip(('AE', 'AF', 'AG', 'AH'), ('AU', 'AV', 'AW', 'AX')):
                b = 'N(VERW_Project!$%s%d)/P_AKFactor*$D%d' % (v, r - 46, r)
                cel += c('%s%d' % (k, r), 150, f='IF(OR($C%d<>"Ja",$D%d=""),"",IF(AK_Negatief="Als 0",MAX(0,%s),%s))' % (r, r, b, b))
        else:
            return m.group(0)
        return m.group(0)[:-len('</row>')] + cel + '</row>'
    data = re.sub(r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*</row>', extra, data, flags=re.S)

    # kop: bereik, deelvensters, kolommen
    kop = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:AH%d"/>' % (122 + INS_N), kop)
    kop = re.sub(r'<pane [^>]*/>', '<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>', kop)
    kop = kop.replace('</cols>', '<col hidden="1" width="3" customWidth="1" min="30" max="30"/>'
                                 '<col hidden="1" outlineLevel="1" width="11" customWidth="1" min="31" max="34"/></cols>')

    # staart: samengevoegde cellen, opmaak, keuzelijsten
    staart = re.sub(r'(<mergeCell ref=")([^"]*)(")', lambda m: m.group(1) + schuif_sqref(m.group(2)) + m.group(3), staart)
    staart = re.sub(r'( sqref=")([^"]*)(")', lambda m: m.group(1) + schuif_sqref(m.group(2)) + m.group(3), staart)
    staart = formules(staart, True)
    cf = ''
    for n, (bereik, cel) in enumerate((('H13:H23', 'H13'), ('K13:K23', 'K13'), ('L13:L23', '$K13'))):
        cf += ('<conditionalFormatting sqref="%s"><cfRule type="expression" priority="%d" dxfId="15"><formula>AND(ISNUMBER(%s),%s&lt;0)</formula></cfRule>'
               '<cfRule type="expression" priority="%d" dxfId="16"><formula>AND(ISNUMBER(%s),%s&gt;=0)</formula></cfRule></conditionalFormatting>'
               % (bereik, 20 + 2 * n, cel, cel, 21 + 2 * n, cel, cel))
    staart = staart.replace('<dataValidations', cf + '<dataValidations', 1)
    staart = staart.replace('</worksheet>', '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdAK1"/></worksheet>')
    return kop + '<sheetData>' + data + '</sheetData>' + staart


def ak_grafiek():
    """Kolommen dekking en kosten per kwartaal, lijn = cumulatieve over- of onderdekking."""
    a = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    lettertype = ('<txPr><a:bodyPr %s/><a:p %s><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="6B7385"/></a:solidFill>'
                  '<a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></txPr>' % (a, a))

    def reeks(i, kol_, naam, vorm):
        return ('<ser><idx val="%d"/><order val="%d"/><tx><v>%s</v></tx>%s'
                '<cat><strRef><f>\'AK-dekking\'!$B$13:$B$20</f></strRef></cat>'
                '<val><numRef><f>\'AK-dekking\'!$%s$13:$%s$20</f></numRef></val>%s</ser>'
                % (i, i, naam, vorm, kol_, kol_, '<smooth val="0"/>' if i == 2 else ''))
    vlak = '<spPr><a:solidFill %s><a:srgbClr val="%%s"/></a:solidFill></spPr><invertIfNegative val="0"/>' % a
    lijn = ('<spPr><a:ln %s w="28575" cap="rnd"><a:solidFill><a:srgbClr val="12864E"/></a:solidFill></a:ln></spPr>'
            '<marker><symbol val="circle"/><size val="6"/><spPr><a:solidFill %s><a:srgbClr val="12864E"/></a:solidFill>'
            '<a:ln %s><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></spPr></marker>' % (a, a, a))
    as_lijn = '<spPr><a:ln %s w="6350"><a:solidFill><a:srgbClr val="BFC6D4"/></a:solidFill></a:ln></spPr>' % a
    raster = '<majorGridlines><spPr><a:ln %s w="6350"><a:solidFill><a:srgbClr val="E6E9EF"/></a:solidFill></a:ln></spPr></majorGridlines>' % a
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><roundedCorners val="0"/><chart>'
            '<autoTitleDeleted val="1"/><plotArea><layout/>'
            '<barChart><barDir val="col"/><grouping val="clustered"/><varyColors val="0"/>'
            + reeks(0, 'F', 'AK-dekking', vlak % '002060') + reeks(1, 'G', 'AK-kosten regio', vlak % 'F2B800') +
            '<gapWidth val="60"/><overlap val="-10"/><axId val="50010"/><axId val="50020"/></barChart>'
            '<lineChart><grouping val="standard"/><varyColors val="0"/>'
            + reeks(2, 'K', 'Over / onder cumulatief', lijn) +
            '<marker val="1"/><axId val="50010"/><axId val="50020"/></lineChart>'
            '<catAx><axId val="50010"/><scaling><orientation val="minMax"/></scaling><delete val="0"/><axPos val="b"/>'
            '<numFmt formatCode="General" sourceLinked="0"/><majorTickMark val="none"/><minorTickMark val="none"/><tickLblPos val="low"/>'
            + as_lijn + lettertype + '<crossAx val="50020"/><crosses val="autoZero"/><auto val="1"/><lblAlgn val="ctr"/><lblOffset val="100"/><noMultiLvlLbl val="0"/></catAx>'
            '<valAx><axId val="50020"/><scaling><orientation val="minMax"/></scaling><delete val="0"/><axPos val="l"/>' + raster +
            '<numFmt formatCode="#,##0" sourceLinked="0"/><majorTickMark val="none"/><minorTickMark val="none"/><tickLblPos val="nextTo"/>'
            '<spPr><a:ln %s><a:noFill/></a:ln></spPr>' % a + lettertype + '<crossAx val="50010"/><crosses val="autoZero"/><crossBetween val="between"/></valAx>'
            '</plotArea><legend><legendPos val="t"/><overlay val="0"/>' + lettertype + '</legend><plotVisOnly val="1"/><dispBlanksAs val="gap"/></chart>'
            '<spPr><a:solidFill %s><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln %s><a:noFill/></a:ln></spPr></chartSpace>' % (a, a))


TEKENING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><oneCellAnchor>'
            '<from><col>1</col><colOff>0</colOff><row>25</row><rowOff>0</rowOff></from><ext cx="9000000" cy="2850000"/>'
            '<graphicFrame><nvGraphicFramePr><cNvPr id="1" name="Grafiek over- of onderdekking"/><cNvGraphicFramePr/></nvGraphicFramePr><xfrm/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
            '</a:graphicData></a:graphic></graphicFrame><clientData/></oneCellAnchor></wsDr>')
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'


# ---------------------------------------------------------------- 2. Vergelijk
def vergelijk_blad():
    K1, KN = 13, 112                                   # projectregels
    OV = KN + 1                                        # overige (niet meer op de projectlijst)
    blokken = (('C', '$B$4', '$C$4'), ('I', '$B$5', '$C$5'))
    letters = lambda start: [kol(ord(start) - 64 + i) for i in range(5)]

    def waarde(r, q, versie, jaar):
        kw = 'SUMIFS(INDEX(KWL_Vals,0,$W$4+%d),KWL_WBS,$A%d,KWL_Versie,%s,KWL_VJ,%s,KWL_Actief,1)' % (q, r, versie, jaar)
        return 'IF($A%d="","",IF(%s="Upload",N(INDEX(Kwartalen!$C%d:$R%d,$W$4+%d)),%s)/P_Factor)' % (r, versie, r - 7, r - 7, q, kw)

    def status(versie, jaar):
        n = 'COUNTIFS(KWL_Versie,%s,KWL_VJ,%s,KWL_Actief,1)' % (versie, jaar)
        p = '_xlfn.MAXIFS(KWL_Peil,KWL_Versie,%s,KWL_VJ,%s,KWL_Actief,1)' % (versie, jaar)
        return ('IF(%s="Upload","huidige upload "&P_UploadVersie&" "&P_UploadVJ&" (nog niet vastgelegd of opnieuw ingelezen)",'
                'IF(%s=0,"✖ %s is niet vastgelegd: geen cijfers","✔ vastgelegd · "&%s&" projecten · peildatum "&DAY(%s)&"-"&MONTH(%s)&"-"&YEAR(%s)))'
                % (versie, n, '"&%s&" "&%s&"' % (versie, jaar), n, p, p, p))

    rijen = []
    rijen.append(rij(1, [c('A1', 19, f='"Vergelijk prognoses — "&P_Regio', tekst=True)] + [c('%s1' % kol(i), 1) for i in range(2, 21)], ht=28))
    rijen.append(rij(2, [c('A2', 3, f='"Twee prognoses naast elkaar, per project en per kwartaal  ·  bedragen in "&P_Eenheid&"  ·  kies hieronder in de gele cellen"', tekst=True)]))
    rijen.append(rij(4, [c('A4', 35, 'Prognose A (oud)'), c('B4', 142, 'Budget'), c('C4', 129, f='P_VJ'), c('E4', 32, f=status('$B$4', '$C$4'), tekst=True),
                         c('W4', 12, f='(MATCH($B$6,{"Omzet","BOR","Kosten","AK-dekking"},0)-1)*4')]))
    rijen.append(rij(5, [c('A5', 35, 'Prognose B (nieuw)'), c('B5', 142, 'Upload'), c('C5', 129, f='P_VJ'), c('E5', 32, f=status('$B$5', '$C$5'), tekst=True)]))
    rijen.append(rij(6, [c('A6', 35, 'Kengetal'), c('B6', 142, 'Omzet'), c('E6', 32, 'Omzet, BOR, Kosten of AK-dekking. Het jaar is het versiejaar: de kwartalen zijn Q1–Q4 van dat jaar.')]))
    rijen.append(rij(8, [c('A8', 4, f='$B$6&" totaal: "&FIXED($G$11,0)&" → "&FIXED($M$11,0)&"   ·   verschil "&IF($S$11>0,"+","")&FIXED($S$11,0)'
                                     '&IF($G$11=0,"","  ("&IF($S$11>0,"+","")&FIXED($S$11/ABS($G$11)*100,1)&"%)")', tekst=True)]))
    rijen.append(rij(9, [c('A9', 161, f='IFERROR(IF(COUNT($S$%d:$S$%d)=0,"","Grootste stijging: "&INDEX($B$%d:$B$%d,MATCH(MAX($S$%d:$S$%d),$S$%d:$S$%d,0))&" ("&FIXED(MAX($S$%d:$S$%d),0)&")'
                                        '   ·   grootste daling: "&INDEX($B$%d:$B$%d,MATCH(MIN($S$%d:$S$%d),$S$%d:$S$%d,0))&" ("&FIXED(MIN($S$%d:$S$%d),0)&")"),"")'
                                        % ((K1, KN) * 9), tekst=True)]))
    # groepskop
    g = [c('A10', 163), c('B10', 163)]
    g += [c('C10', 162, f='"PROGNOSE A · "&$B$4&" "&$C$4', tekst=True)] + [c('%s10' % k, 163) for k in 'DEFG'] + [c('H10', 163)]
    g += [c('I10', 162, f='"PROGNOSE B · "&$B$5&" "&$C$5', tekst=True)] + [c('%s10' % k, 163) for k in 'JKLM'] + [c('N10', 163)]
    g += [c('O10', 162, 'VERSCHIL B − A')] + [c('%s10' % k, 163) for k in 'PQRST']
    rijen.append(rij(10, g))
    # totaal en kolomkoppen
    tot = [c('A11', 55, 'Totaal'), c('B11', 55)]
    kopr = [c('A12', 39, 'WBS-element'), c('B12', 39, 'Project')]
    for start in ('C', 'I', 'O'):
        for i, k in enumerate(letters(start)):
            tot.append(c('%s11' % k, 73, f='SUM(%s%d:%s%d)' % (k, K1, k, OV)))
            kopr.append(c('%s12' % k, 40, 'Jaar' if i == 4 else 'Q%d' % (i + 1)))
    tot.append(c('T11', 164, f='IF($G$11=0,"",$S$11/ABS($G$11))'))
    kopr.append(c('T12', 40, 'Verschil\n%'))
    rijen.append(rij(11, tot))
    rijen.append(rij(12, kopr, ht=28))
    for r in range(K1, KN + 1):
        cl = [c('A%d' % r, 79, f='Kwartalen!$A%d' % (r - 7), tekst=True), c('B%d' % r, 79, f='IF($A%d="","",Kwartalen!$B%d)' % (r, r - 7), tekst=True)]
        for start, versie, jaar in blokken:
            ks = letters(start)
            for q, k in enumerate(ks[:4], 1):
                cl.append(c('%s%d' % (k, r), 71, f=waarde(r, q, versie, jaar)))
            cl.append(c('%s%d' % (ks[4], r), 98, f='IF($A%d="","",SUM(%s%d:%s%d))' % (r, ks[0], r, ks[3], r)))
        for i, k in enumerate(letters('O')):
            a, b = letters('C')[i], letters('I')[i]
            cl.append(c('%s%d' % (k, r), 98 if i == 4 else 71, f='IF($A%d="","",%s%d-%s%d)' % (r, b, r, a, r)))
        cl.append(c('T%d' % r, 164, f='IF(OR($A%d="",N(G%d)=0),"",S%d/ABS(G%d))' % (r, r, r, r)))
        rijen.append(rij(r, cl))
    # overige: vastgelegd in die versie, maar het project staat niet meer op de lijst
    cl = [c('A%d' % OV, 101, 'Overige'), c('B%d' % OV, 101, 'projecten die niet meer op de projectlijst staan')]
    for start, versie, jaar in blokken:
        ks = letters(start)
        for q, k in enumerate(ks[:4], 1):
            cl.append(c('%s%d' % (k, OV), 150, f='IF(%s="Upload",0,ROUND(SUMIFS(INDEX(KWL_Vals,0,$W$4+%d),KWL_Versie,%s,KWL_VJ,%s,KWL_Actief,1)/P_Factor-SUM(%s%d:%s%d),2))'
                        % (versie, q, versie, jaar, k, K1, k, KN)))
        cl.append(c('%s%d' % (ks[4], OV), 150, f='SUM(%s%d:%s%d)' % (ks[0], OV, ks[3], OV)))
    for i, k in enumerate(letters('O')):
        cl.append(c('%s%d' % (k, OV), 150, f='%s%d-%s%d' % (letters('I')[i], OV, letters('C')[i], OV)))
    rijen.append(rij(OV, cl))
    uitleg = ('Leeswijzer',
              '• Budget en Q1–Q4 zijn vastgelegde prognoses (knop "Prognose vastleggen" op Start). Upload = de cijfers die nu zijn ingelezen.',
              '• Groen = gunstiger in prognose B, rood = ongunstiger. Bij Kosten is lager gunstiger.',
              '• Een vastgelegde prognose bevat de vier kwartalen van het versiejaar; vergelijk daarom prognoses van hetzelfde jaar.')
    for i, t in enumerate(uitleg):
        rijen.append(rij(OV + 2 + i, [c('A%d' % (OV + 2 + i), 35 if i == 0 else 32, t)]))

    breedtes = [(1, 1, 16), (2, 2, 34), (3, 6, 10.5), (7, 7, 11.5), (8, 8, 1.5), (9, 12, 10.5), (13, 13, 11.5), (14, 14, 1.5),
                (15, 18, 10.5), (19, 19, 11.5), (20, 20, 9)]
    cols = ''.join('<col width="%s" customWidth="1" min="%d" max="%d"/>' % (w, a, b) for a, b, w in breedtes)
    cols += '<col hidden="1" width="9" customWidth="1" min="23" max="23"/>'
    teken = 'IF($B$6="Kosten",-1,1)'
    cf = ('<conditionalFormatting sqref="O11:T%d"><cfRule type="expression" priority="1" dxfId="4"><formula>AND(ISNUMBER(O11),O11*%s&lt;0)</formula></cfRule>'
          '<cfRule type="expression" priority="2" dxfId="5"><formula>AND(ISNUMBER(O11),O11*%s&gt;0)</formula></cfRule></conditionalFormatting>'
          '<conditionalFormatting sqref="E4:E5"><cfRule type="expression" priority="3" dxfId="0"><formula>LEFT(E4,1)="✖"</formula></cfRule></conditionalFormatting>'
          % (OV, teken, teken))
    dv = ('<dataValidations count="2">'
          '<dataValidation sqref="B4:B5" showErrorMessage="1" allowBlank="0" error="Kies een waarde uit de lijst." type="list"><formula1>"Budget,Q1,Q2,Q3,Q4,Upload"</formula1></dataValidation>'
          '<dataValidation sqref="B6" showErrorMessage="1" allowBlank="0" error="Kies een waarde uit de lijst." type="list"><formula1>"Omzet,BOR,Kosten,AK-dekking"</formula1></dataValidation>'
          '</dataValidations>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="00F2B800"/><pageSetUpPr fitToPage="1"/></sheetPr>'
            '<dimension ref="A1:W%d"/><sheetViews><sheetView showGridLines="0" zoomScale="100" workbookViewId="0">'
            '<pane xSplit="2" ySplit="12" topLeftCell="C13" activePane="bottomRight" state="frozen"/><selection pane="topRight"/><selection pane="bottomLeft"/>'
            '<selection pane="bottomRight" activeCell="B4" sqref="B4"/></sheetView></sheetViews><sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>'
            '<cols>%s</cols><sheetData>%s</sheetData>'
            '<mergeCells count="3"><mergeCell ref="C10:G10"/><mergeCell ref="I10:M10"/><mergeCell ref="O10:T10"/></mergeCells>%s%s'
            '<printOptions horizontalCentered="1"/><pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.5" footer="0.5"/>'
            '<pageSetup orientation="landscape" paperSize="9" fitToHeight="0" fitToWidth="1"/></worksheet>'
            % (OV + 5, cols, ''.join(rijen), cf, dv))


# ---------------------------------------------------------------- pakket bijwerken
def werkmap(x):
    x = re.sub(r'(<definedName [^>]*>)([^<]*)(</definedName>)', lambda m: m.group(1) + schuif(m.group(2), False) + m.group(3), x)
    # Vergelijk komt na RDFO (positie 3): bladgebonden namen van latere bladen schuiven één op
    x = re.sub(r'localSheetId="(\d+)"', lambda m: 'localSheetId="%d"' % (int(m.group(1)) + (1 if int(m.group(1)) >= 3 else 0)), x)
    blad = '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Vergelijk" sheetId="32" state="visible" r:id="rIdVergelijk"/>'
    i = x.index('<sheet ', x.index('name="RDFO"'))
    return x[:i] + blad + x[i:]


def main():
    zi = zipfile.ZipFile(SRC)
    if b'name="Vergelijk"' in zi.read('xl/workbook.xml'):
        sys.exit('Dit bestand heeft het tabblad Vergelijk al.')
    naam_ak = 'xl/worksheets/sheet12.xml'
    assert b'codeName="shAKdekking"' in zi.read(naam_ak)
    zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
    ct = ('<Override PartName="/xl/worksheets/sheet32.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '<Override PartName="/xl/drawings/drawing11.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
          '<Override PartName="/xl/charts/chart25.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    for it in zi.infolist():
        n = it.filename
        d = zi.read(n)
        if n == naam_ak:
            d = ak_blad(d.decode('utf8')).encode('utf8')
        elif n in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet16.xml'):      # Start en Handleiding: uitleg bijwerken
            d = formules(d.decode('utf8'), False)
            d = d.replace('Bovenaan het totaal per kwartaal, daarna een rekenvoorbeeld',
                          'Bovenaan het totaal per kwartaal en de over- of onderdekking (AK-kosten invullen in de gele cellen), daarna een rekenvoorbeeld')
            d = d.replace('Onderaan: per gebied.', 'Onderaan: per gebied. Tabblad Vergelijk: twee prognoses naast elkaar per project en kwartaal.').encode('utf8')
        elif n.startswith('xl/worksheets/sheet') and b"'AK-dekking'!" in d:
            d = formules(d.decode('utf8'), False).encode('utf8')
        elif n.startswith('xl/charts/') and b"'AK-dekking'!" in d:
            d = formules(d.decode('utf8'), False, tags=('f',)).encode('utf8')
        elif n == 'xl/workbook.xml':
            d = werkmap(d.decode('utf8')).encode('utf8')
        elif n == 'xl/_rels/workbook.xml.rels':
            d = d.decode('utf8').replace('</Relationships>', '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                                         'Target="/xl/worksheets/sheet32.xml" Id="rIdVergelijk"/></Relationships>').encode('utf8')
        elif n == '[Content_Types].xml':
            d = d.decode('utf8').replace('</Types>', ct + '</Types>').encode('utf8')
        zo.writestr(it, d)
    assert 'xl/worksheets/_rels/sheet12.xml.rels' not in zi.namelist()
    zo.writestr('xl/worksheets/sheet32.xml', vergelijk_blad())
    zo.writestr('xl/worksheets/_rels/sheet12.xml.rels', RELS % '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/drawing11.xml" Id="rIdAK1"/>')
    zo.writestr('xl/drawings/drawing11.xml', TEKENING)
    zo.writestr('xl/drawings/_rels/drawing11.xml.rels', RELS % '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/xl/charts/chart25.xml" Id="rId1"/>')
    zo.writestr('xl/charts/chart25.xml', ak_grafiek())
    zo.close()


if __name__ == '__main__':
    main()
