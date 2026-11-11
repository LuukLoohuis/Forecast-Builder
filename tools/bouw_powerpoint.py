"""PowerPoint-macro in de stijl van de cashflow-presentatie, plus twee nieuwe dia's
(AK-dekking over of onder, vergelijking van prognoses).

Gebruik:  python3 tools/bouw_powerpoint.py IN.xlsm UIT.xlsm
Draait na bouw_ak_vergelijk.py. Past de macrotekst aan in vbaProject.bin (alleen module modRDFO)
en vult de dia-specificatie (verborgen tabbladen PPT en Dias) aan.
"""
import io
import re
import struct
import sys
import zipfile
from xml.sax.saxutils import escape

SRC, DST = sys.argv[1], sys.argv[2]


# ================================================================ macrotekst
def vervang(tekst, oud, nieuw, aantal=1):
    assert tekst.count(oud) == aantal, (tekst.count(oud), oud[:70])
    return tekst.replace(oud, nieuw)


def blok(tekst, begin, eind, nieuw):
    """Vervangt een hele procedure: van de regel 'begin' tot en met de eerstvolgende regel 'eind'."""
    i = tekst.index(begin)
    assert tekst.count(begin) == 1, begin
    j = tekst.index(eind, i) + len(eind)
    return tekst[:i] + nieuw.strip('\n') + tekst[j:]


TITELDIA = '''
Private Sub TitelDia(ByVal pres As Object)
    ' donkerblauw vlak met witte titel, rechts een gele baan (zoals de cashflow-presentatie)
    Dim sld As Object, sh As Object
    Set sld = pres.Slides.Add(1, 12)
    Vlak sld, 0, 0, 960, 540, GEEL
    Vlak sld, 0, 0, 640, 540, NAVY
    Set sh = sld.Shapes.AddTextbox(1, 72, 150, 520, 16)
    Tekst sh, T("P_Regio"), 12, GEEL, True
    Set sh = sld.Shapes.AddTextbox(1, 72, 176, 540, 110)
    Tekst sh, T("PPT_Titel"), 36, WIT, True
    Set sh = sld.Shapes.AddTextbox(1, 72, 300, 540, 30)
    Tekst sh, T("PPT_Sub"), 15, GEEL, False
    Set sh = sld.Shapes.AddTextbox(1, 72, 470, 540, 16)
    Tekst sh, "Gemaakt op " & Format(Now, "dd-mm-yyyy") & "  ·  bron: " & ThisWorkbook.Name, 9, WIT, False
    Set sh = sld.Shapes.AddTextbox(1, 700, 464, 200, 26)
    Tekst sh, "heijmans", 18, NAVY, True
End Sub
'''

NIEUWEDIA = '''
Private Function NieuweDia(ByVal pres As Object, ByVal titel As String, ByVal kern As String) As Object
    ' lichte achtergrond, gele balk, onderwerp klein in oranje en de kernboodschap als kop
    Dim sld As Object, sh As Object, hoofd As String
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Vlak sld, 0, 0, 960, 540, ACHTER
    Vlak sld, 0, 0, 960, 8, GEEL
    Set sh = sld.Shapes.AddTextbox(1, 36, 24, 888, 14)
    Tekst sh, UCase(titel), 9, ORANJE, True
    hoofd = kern
    If hoofd = "" Then hoofd = titel
    Set sh = sld.Shapes.AddTextbox(1, 36, 40, 888, 46)
    If Len(hoofd) > 90 Then
        Tekst sh, hoofd, 14, NAVY, True
    Else
        Tekst sh, hoofd, 19, NAVY, True
    End If
    Set sh = sld.Shapes.AddTextbox(1, 36, 518, 600, 14)
    Tekst sh, T("PPT_Footer"), 8, GRIJS, False
    Set sh = sld.Shapes.AddTextbox(1, 770, 514, 110, 18)
    Tekst sh, "heijmans", 12, NAVY, True
    sh.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    Set sh = sld.Shapes.AddTextbox(1, 884, 518, 40, 14)
    Tekst sh, CStr(sld.SlideIndex), 8, GRIJS, False
    sh.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    Set NieuweDia = sld
End Function
'''

TEGELS = '''
Private Sub Tegels(ByVal sld As Object, ByVal blok As Range, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    ' witte kaarten met een gekleurde rand links
    Dim k As Long, tw As Single, tx As Single, sh As Object
    tw = (w - 5 * 8) / 6
    For k = 1 To 6
        tx = x + (k - 1) * (tw + 8)
        Vlak sld, tx, y, tw, h, WIT
        If k Mod 2 = 1 Then
            Vlak sld, tx, y, 4, h, NAVY
        Else
            Vlak sld, tx, y, 4, h, GEEL
        End If
        Set sh = sld.Shapes.AddTextbox(1, tx + 12, y + 8, tw - 18, 14)
        Tekst sh, UCase(CelTekst(blok.Cells(k, 1))), 7.5, GRIJS, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 12, y + 24, tw - 18, 28)
        Tekst sh, CelTekst(blok.Cells(k, 2)), 20, NAVY, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 12, y + h - 18, tw - 18, 14)
        Tekst sh, CelTekst(blok.Cells(k, 3)), 8, GRIJS, False
    Next k
End Sub
'''

GRAFIEK_KOP_OUD = '''Private Sub Grafiek(ByVal sld As Object, ByVal wsD As Worksheet, ByVal nr As Long, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal tekstKop As String)
    Dim co As ChartObject, sh As Object, poging As Integer, bestand As String
    Set co = wsD.ChartObjects(nr)
    Kop sld, tekstKop, x, y, w
'''
GRAFIEK_KOP_NIEUW = '''Private Sub Grafiek(ByVal sld As Object, ByVal wsD As Worksheet, ByVal grafRef As String, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal tekstKop As String)
    ' grafRef = nummer van de grafiek op het tabblad Dias, of "Tabblad|nummer" voor een grafiek op een ander tabblad
    Dim co As ChartObject, sh As Object, poging As Integer, bestand As String, delen() As String, nr As Long
    If InStr(grafRef, "|") > 0 Then
        delen = Split(grafRef, "|")
        nr = CLng(delen(1))
        Set co = ThisWorkbook.Worksheets(delen(0)).ChartObjects(nr)
        nr = nr + 100
    Else
        nr = CLng(grafRef)
        Set co = wsD.ChartObjects(nr)
    End If
    Kop sld, tekstKop, x, y, w
    Vlak sld, x, y, w, h, WIT
'''


def macro(t):
    t = vervang(t, 'NAVY = RGB(20, 40, 75)', 'NAVY = RGB(0, 32, 96)')
    t = vervang(t, 'GEEL = RGB(245, 184, 0)', 'GEEL = RGB(242, 184, 0)')
    t = vervang(t, 'Private Function WIT() As Long\n    WIT = RGB(255, 255, 255)\nEnd Function\n',
                'Private Function WIT() As Long\n    WIT = RGB(255, 255, 255)\nEnd Function\n'
                'Private Function ORANJE() As Long\n    ORANJE = RGB(240, 138, 0)\nEnd Function\n'
                'Private Function ACHTER() As Long\n    ACHTER = RGB(247, 248, 250)\nEnd Function\n')
    # de projectkeuze bij het rekenvoorbeeld staat sinds het blok over- of onderdekking op F42
    t = vervang(t, 'ThisWorkbook.Worksheets("AK-dekking").Range("F11").Value = ""', 'ThisWorkbook.Worksheets("AK-dekking").Range("F42").Value = ""')
    t = vervang(t, 'Case "grafiek": Grafiek sld, wsD, CLng(spec.Cells(r, 5).Value), ', 'Case "grafiek": Grafiek sld, wsD, spec.Cells(r, 5).Text, ')
    t = blok(t, 'Private Sub TitelDia(ByVal pres As Object)', 'End Sub\n', TITELDIA + '\n')
    t = blok(t, 'Private Function NieuweDia(', 'End Function\n', NIEUWEDIA + '\n')
    t = blok(t, 'Private Sub Tegels(', 'End Sub\n', TEGELS + '\n')
    t = vervang(t, GRAFIEK_KOP_OUD, GRAFIEK_KOP_NIEUW)
    # tabelkop: geel met blauwe letters
    t = vervang(t, '                    If i = 1 Then\n                        .Color.RGB = WIT\n', '                    If i = 1 Then\n                        .Color.RGB = NAVY\n')
    t = vervang(t, '                If i = 1 Then\n                    .Fill.ForeColor.RGB = NAVY\n', '                If i = 1 Then\n                    .Fill.ForeColor.RGB = GEEL\n')
    # opsomming op een witte kaart
    t = vervang(t, '    Vlak sld, x, y, w, h, LICHT\n    Set sh = sld.Shapes.AddTextbox(1, x + 24, y + 20, w - 48, h - 40)',
                '    Vlak sld, x, y, w, h, WIT\n    Set sh = sld.Shapes.AddTextbox(1, x + 24, y + 20, w - 48, h - 40)')
    t = vervang(t, "'   RDFO_MaakPowerPoint      : presentatie in huisstijl maken (geel / donkerblauw)",
                "'   RDFO_MaakPowerPoint      : presentatie in huisstijl maken (geel / blauw, opzet zoals de cashflow-presentatie)")
    return t


# ================================================================ MS-OVBA-compressie
def comprimeer(data):
    uit = bytearray(b'\x01')
    for start in range(0, len(data), 4096):
        stuk = data[start:start + 4096]
        buf = bytearray()
        pos = 0
        while pos < len(stuk):
            vlag = 0
            tokens = bytearray()
            for bit in range(8):
                if pos >= len(stuk):
                    break
                bits = max((pos - 1).bit_length(), 4) if pos > 0 else 4
                maxlen = (0xFFFF >> bits) + 3
                beste_len, beste_off = 0, 0
                if pos > 0 and pos + 3 <= len(stuk):
                    kandidaat = stuk.rfind(stuk[pos:pos + 3], 0, pos + 2)
                    pogingen = 0
                    while kandidaat >= 0 and pogingen < 48:
                        n = 3
                        grens = min(maxlen, len(stuk) - pos)
                        while n < grens and stuk[kandidaat + n] == stuk[pos + n]:
                            n += 1
                        if n > beste_len:
                            beste_len, beste_off = n, pos - kandidaat
                            if n == grens:
                                break
                        kandidaat = stuk.rfind(stuk[pos:pos + 3], 0, kandidaat + 2)
                        pogingen += 1
                if beste_len >= 3:
                    tokens += struct.pack('<H', ((beste_off - 1) << (16 - bits)) | (beste_len - 3))
                    vlag |= 1 << bit
                    pos += beste_len
                else:
                    tokens.append(stuk[pos])
                    pos += 1
            buf.append(vlag)
            buf += tokens
        assert len(buf) <= 4096 and len(buf) >= 1
        uit += struct.pack('<H', 0xB000 | (len(buf) + 2 - 3)) + buf
    return bytes(uit)


# ================================================================ stroom in het OLE-bestand vervangen
def ole_vervang(bin_, naam, nieuw):
    """Vervangt één (grote) stroom; de rest van het bestand blijft byte voor byte gelijk."""
    b = bytearray(bin_)
    assert struct.unpack_from('<H', b, 0x1E)[0] == 9 and struct.unpack_from('<I', b, 0x2C)[0] == 1
    sect = lambda i: 512 + i * 512
    fat_sector = struct.unpack_from('<I', b, 0x4C)[0]
    fat = list(struct.unpack_from('<128I', b, sect(fat_sector)))
    EINDE, VRIJ = 0xFFFFFFFE, 0xFFFFFFFF
    # map doorlopen
    s = struct.unpack_from('<I', b, 0x30)[0]
    ingang = None
    while s != EINDE:
        for k in range(4):
            o = sect(s) + k * 128
            n = struct.unpack_from('<H', b, o + 0x40)[0]
            if n and b[o:o + n - 2].decode('utf-16-le') == naam:
                ingang = o
        s = fat[s]
    assert ingang is not None
    start, grootte = struct.unpack_from('<II', b, ingang + 0x74)
    assert grootte >= 4096 and len(nieuw) >= 4096
    keten = [start]
    while fat[keten[-1]] != EINDE:
        keten.append(fat[keten[-1]])
    nodig = -(-len(nieuw) // 512)
    while len(keten) < nodig:                       # nieuwe sectoren achteraan
        assert (len(b) - 512) % 512 == 0
        i = (len(b) - 512) // 512
        assert i < 128, 'FAT vol'
        b += bytes(512)
        keten.append(i)
    for i in keten[nodig:]:
        fat[i] = VRIJ
    keten = keten[:nodig]
    for a, c in zip(keten, keten[1:] + [EINDE]):
        fat[a] = c
    data = nieuw + bytes(nodig * 512 - len(nieuw))
    for k, i in enumerate(keten):
        b[sect(i):sect(i) + 512] = data[k * 512:(k + 1) * 512]
    struct.pack_into('<128I', b, sect(fat_sector), *fat)
    struct.pack_into('<II', b, ingang + 0x74, keten[0], len(nieuw))
    return bytes(b)


def vba(bin_):
    import olefile
    from oletools import olevba
    o = olefile.OleFileIO(io.BytesIO(bin_))
    bron = bytes(olevba.decompress_stream(bytearray(o.openstream('VBA/modRDFO').read())))
    tekst = macro(bron.decode('cp1252').replace('\r\n', '\n')).replace('\n', '\r\n').encode('cp1252')
    gecomprimeerd = comprimeer(tekst)
    assert bytes(olevba.decompress_stream(bytearray(gecomprimeerd))) == tekst
    nieuw = ole_vervang(bin_, 'modRDFO', gecomprimeerd)
    # controle: alle andere stromen ongewijzigd, de module leest terug als de nieuwe tekst
    n = olefile.OleFileIO(io.BytesIO(nieuw))
    for st in o.listdir(streams=True):
        if st != ['VBA', 'modRDFO']:
            assert o.openstream(st).read() == n.openstream(st).read(), st
    assert bytes(olevba.decompress_stream(bytearray(n.openstream('VBA/modRDFO').read()))) == tekst
    return nieuw, tekst


# ================================================================ cellen
def c(ref, s, v=None, f=None, tekst=False):
    st = ' s="%d"' % s if s is not None else ''
    if f is not None:
        return '<c r="%s"%s%s><f>%s</f></c>' % (ref, st, ' t="str"' if tekst else '', escape(f))
    if v is None:
        return '<c r="%s"%s/>' % (ref, st)
    if isinstance(v, str):
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, st, escape(v))
    return '<c r="%s"%s><v>%s</v></c>' % (ref, st, v)


def rij(nr, cellen):
    return '<row r="%d">%s</row>' % (nr, ''.join(cellen))


AKD0, VGL0 = 283, 297          # kopregels van de twee nieuwe tabellen op Dias


def dias(x):
    r = [rij(AKD0 - 1, [c('A%d' % (AKD0 - 1), 120, 'gewicht')] + [c('%s%d' % (k, AKD0 - 1), 120, w) for k, w in zip('BCDEFGH', (1.5, 1.1, 1.2, 1.2, 1.2, 1.2, 1.1))])]
    r.append(rij(AKD0, [c('A%d' % AKD0, 120, 'T_AKD'), c('B%d' % AKD0, 39, 'Periode')] +
                 [c('%s%d' % (k, AKD0), 40, t) for k, t in zip('CDEFGH', ('Bron', 'AK-dekking', 'AK-kosten regio', 'Over / onder', 'Cumulatief', 'Status'))]))
    for i in range(8):
        n, a = AKD0 + 1 + i, 13 + i
        r.append(rij(n, [c('B%d' % n, 243, f="'AK-dekking'!$B$%d&\"\"" % a, tekst=True), c('C%d' % n, 243, f="'AK-dekking'!$E$%d&\"\"" % a, tekst=True),
                         c('D%d' % n, 262, f="'AK-dekking'!$F$%d" % a),
                         c('E%d' % n, 262, f="IF('AK-dekking'!$G$%d=\"\",\"n.b.\",'AK-dekking'!$G$%d)" % (a, a)),
                         c('F%d' % n, 262, f="'AK-dekking'!$H$%d" % a), c('G%d' % n, 262, f="'AK-dekking'!$K$%d" % a),
                         c('H%d' % n, 243, f="'AK-dekking'!$L$%d&\"\"" % a, tekst=True)]))
    # grootste stijgingen en dalingen tussen de twee prognoses op het tabblad Vergelijk
    V = 'Vergelijk!$%s$13:$%s$112'
    r.append(rij(VGL0 - 1, [c('A%d' % (VGL0 - 1), 120, 'gewicht')] + [c('%s%d' % (k, VGL0 - 1), 120, w) for k, w in zip('BCDEF', (3, 1.2, 1.2, 1.2, 0.9))]))
    r.append(rij(VGL0, [c('A%d' % VGL0, 120, 'T_VGL'), c('B%d' % VGL0, 39, f='"Project ("&Vergelijk!$B$6&", "&P_Eenheid&")"', tekst=True),
                        c('C%d' % VGL0, 40, f='Vergelijk!$B$4&" "&Vergelijk!$C$4', tekst=True), c('D%d' % VGL0, 40, f='Vergelijk!$B$5&" "&Vergelijk!$C$5', tekst=True),
                        c('E%d' % VGL0, 40, 'Verschil'), c('F%d' % VGL0, 40, '%')]))
    for i in range(10):
        n = VGL0 + 1 + i
        k, fn, teken = (i + 1, 'LARGE', '>') if i < 5 else (i - 4, 'SMALL', '<')
        idx = 'MATCH(%s(%s,%d),%s,0)' % (fn, V % ('V', 'V'), k, V % ('V', 'V'))
        r.append(rij(n, [c('A%d' % n, 120, f='IFERROR(%s,0)' % idx),
                         c('B%d' % n, 243, f='IF($A%d=0,"",IF(INDEX(%s,$A%d)%s0,INDEX(%s,$A%d),""))' % (n, V % ('S', 'S'), n, teken, V % ('B', 'B'), n), tekst=True)] +
                     [c('%s%d' % (kol, n), 164 if kol == 'F' else 262, f='IF($B%d="","",INDEX(%s,$A%d))' % (n, V % (bron, bron), n))
                      for kol, bron in zip('CDEF', 'GMST')]))
    n = VGL0 + 11
    r.append(rij(n, [c('B%d' % n, 55, 'Totaal alle projecten')] + [c('%s%d' % (kol, n), 164 if kol == 'F' else 73, f='Vergelijk!$%s$11' % bron) for kol, bron in zip('CDEF', 'GMST')]))
    for nieuw in r:                                   # de regels bestaan al als lege regel
        nr = re.match(r'<row r="(\d+)"', nieuw).group(1)
        leeg = '<row r="%s" ht="15" customHeight="1"/>' % nr
        assert x.count(leeg) == 1, nr
        x = x.replace(leeg, nieuw.replace('<row r="%s">' % nr, '<row r="%s" ht="15" customHeight="1">' % nr))
    teken = 'IF(Vergelijk!$B$6="Kosten",-1,1)'
    cf = ('<conditionalFormatting sqref="F{a}:G{b}"><cfRule type="expression" priority="901" dxfId="4"><formula>AND(ISNUMBER(F{a}),F{a}&lt;0)</formula></cfRule>'
          '<cfRule type="expression" priority="902" dxfId="5"><formula>AND(ISNUMBER(F{a}),F{a}&gt;0)</formula></cfRule></conditionalFormatting>'
          '<conditionalFormatting sqref="H{a}:H{b}"><cfRule type="expression" priority="903" dxfId="4"><formula>AND(ISNUMBER($G{a}),$G{a}&lt;0)</formula></cfRule>'
          '<cfRule type="expression" priority="904" dxfId="5"><formula>AND(ISNUMBER($G{a}),$G{a}&gt;0)</formula></cfRule></conditionalFormatting>'
          ).format(a=AKD0 + 1, b=AKD0 + 8)
    cf += ('<conditionalFormatting sqref="E{a}:F{b}"><cfRule type="expression" priority="905" dxfId="4"><formula>AND(ISNUMBER(E{a}),E{a}*{t}&lt;0)</formula></cfRule>'
           '<cfRule type="expression" priority="906" dxfId="5"><formula>AND(ISNUMBER(E{a}),E{a}*{t}&gt;0)</formula></cfRule></conditionalFormatting>'
           ).format(a=VGL0 + 1, b=VGL0 + 11, t=teken)
    i = x.index('<conditionalFormatting')
    return x[:i] + cf + x[i:]


def ppt(x):
    oud = re.search(r'<row r="66">.*?</row>', x, re.S).group(0)
    assert 'Aandachtspunten' in oud and '<row r="67"' not in x
    ak_kern = ('IF(COUNT(\'AK-dekking\'!$G$13:$G$16)<4,"AK-kosten "&P_RJ&" zijn nog niet voor alle kwartalen ingevuld (tabblad AK-dekking)",'
               'IF(\'AK-dekking\'!$K$16>=0,"Overdekking","Onderdekking")&" van € "&FIXED(ABS(\'AK-dekking\'!$K$16)/1000,0)&" duizend over "&P_RJ'
               '&"  ·  dekking € "&FIXED(\'AK-dekking\'!$I$16/1000,0)&" duizend, kosten € "&FIXED(\'AK-dekking\'!$J$16/1000,0)&" duizend")')
    ak_aan = "IF(COUNT('AK-dekking'!$G$13:$G$20)>0,1,0)"

    def regel(n, dia, titel, kern, soort, ref, x_, y, b, h, kop, actief):
        cel = [c('H%d' % n, None, dia), c('I%d' % n, None, titel)]
        cel.append(c('J%d' % n, None, f=kern, tekst=True) if kern else c('J%d' % n, None, ''))
        cel += [c('K%d' % n, None, soort), c('L%d' % n, None, ref), c('M%d' % n, None, x_), c('N%d' % n, None, y), c('O%d' % n, None, b), c('P%d' % n, None, h),
                c('Q%d' % n, None, kop)]
        cel.append(c('R%d' % n, None, f=actief) if isinstance(actief, str) else c('R%d' % n, None, actief))
        return rij(n, cel)
    nieuw = (regel(66, 18, 'AK-dekking: over of onder', ak_kern, 'tabel', 'T_AKD', 36, 120, 420, 0, 'Per kwartaal (EUR)', ak_aan) +
             regel(67, 18, '', None, 'grafiek', 'AK-dekking|1', 488, 120, 436, 370, 'Dekking, kosten en cumulatief saldo', ak_aan) +
             regel(68, 19, 'Vergelijking van prognoses', 'Vergelijk!$A$8', 'tabel', 'T_VGL', 36, 120, 888, 0,
                   'Grootste stijgingen en dalingen per project', 'IF(N(Vergelijk!$G$11)*N(Vergelijk!$M$11)<>0,1,0)') +
             oud.replace('r="66"', 'r="69"').replace('66"', '69"').replace('<v>18</v>', '<v>20</v>', 1))
    return x.replace(oud, nieuw)


def vergelijk(x):
    assert 'r="V13"' not in x
    x = x.replace('<col hidden="1" width="9" customWidth="1" min="23" max="23"/>', '<col hidden="1" width="9" customWidth="1" min="22" max="23"/>')

    def extra(m):
        n = int(m.group(1))
        if not 13 <= n <= 112:
            return m.group(0)
        return m.group(0)[:-len('</row>')] + c('V%d' % n, 12, f='IF(ISNUMBER(S%d),S%d+ROW()/10^9,"")' % (n, n)) + '</row>'
    return re.sub(r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*</row>', extra, x, flags=re.S)


def werkmap(x):
    x = x.replace('<definedName name="PPT_Spec">PPT!$H$40:$R$66</definedName>', '<definedName name="PPT_Spec">PPT!$H$40:$R$69</definedName>')
    assert 'PPT!$H$40:$R$69' in x and 'T_AKD' not in x
    namen = ('<definedName name="T_AKD">Dias!$B$%d:$H$%d</definedName><definedName name="T_VGL">Dias!$B$%d:$F$%d</definedName>'
             % (AKD0, AKD0 + 8, VGL0, VGL0 + 11))
    return x.replace('</definedNames>', namen + '</definedNames>')


def main():
    zi = zipfile.ZipFile(SRC)
    wm = zi.read('xl/workbook.xml').decode('utf8')
    assert 'name="Vergelijk"' in wm, 'draai eerst bouw_ak_vergelijk.py'
    stappen = {'xl/worksheets/sheet31.xml': dias, 'xl/worksheets/sheet30.xml': ppt, 'xl/worksheets/sheet32.xml': vergelijk, 'xl/workbook.xml': werkmap}
    assert b'codeName="shDias"' in zi.read('xl/worksheets/sheet31.xml') and b'codeName="shPPT"' in zi.read('xl/worksheets/sheet30.xml')
    zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename in stappen:
            d = stappen[it.filename](d.decode('utf8')).encode('utf8')
        elif it.filename == 'xl/vbaProject.bin':
            d, tekst = vba(d)
            open(DST + '.modRDFO.bas', 'wb').write(tekst)
        zo.writestr(it, d)
    zo.close()


if __name__ == '__main__':
    main()
