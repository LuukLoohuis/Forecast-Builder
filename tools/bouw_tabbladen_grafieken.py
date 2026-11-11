"""Minder zichtbare tabbladen en rustiger grafieken.

Gebruik:  python3 tools/bouw_tabbladen_grafieken.py IN.xlsm UIT.xlsm
Draait na bouw_powerpoint.py.

  - Outlook, Bandbreedte, Kansen en risico's, Reserves en Handleiding worden 'verdieping':
    standaard verborgen, met één knop op Start weer te tonen (macro RDFO_ToonVerdieping).
  - Grafieken: smallere kolommen, groen en rood uit de presentatie, budgetstreepje in oranje.
"""
import io
import re
import sys
import zipfile
from xml.sax.saxutils import escape

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from bouw_powerpoint import comprimeer, ole_vervang, vervang  # noqa: E402

SRC, DST = sys.argv[1], sys.argv[2]
VERDIEPING = ["Outlook", "Bandbreedte", "Kansen en risico's", "Reserves", "Handleiding"]

MACRO = '''
' ================================================================= 3b. VERDIEPING
' Outlook, Bandbreedte, Kansen en risico's, Reserves en Handleiding staan standaard verborgen; deze knop toont of verbergt ze.
Public Sub RDFO_ToonVerdieping()
    Dim naam As Variant, tonen As Boolean
    tonen = (ThisWorkbook.Worksheets("Outlook").Visible <> xlSheetVisible)
    For Each naam In Split(VERDIEPING, ",")
        If tonen Then
            ThisWorkbook.Worksheets(CStr(naam)).Visible = xlSheetVisible
        Else
            ThisWorkbook.Worksheets(CStr(naam)).Visible = xlSheetHidden
        End If
    Next naam
    If tonen Then
        ThisWorkbook.Worksheets("Outlook").Activate
    Else
        ThisWorkbook.Worksheets("Start").Activate
    End If
End Sub

'''


def macro(t):
    regel = [r for r in t.split('\n') if r.startswith('Private Const VERBORGEN As String')]
    assert len(regel) == 1
    t = vervang(t, regel[0] + '\n', regel[0] + '\nPrivate Const VERDIEPING As String = "%s"\n' % ','.join(VERDIEPING))
    kop = "' ================================================================= 4. POWERPOINT\n"
    return vervang(t, kop, MACRO.lstrip('\n') + kop)


def vba(bin_):
    import olefile
    from oletools import olevba
    o = olefile.OleFileIO(io.BytesIO(bin_))
    bron = bytes(olevba.decompress_stream(bytearray(o.openstream('VBA/modRDFO').read())))
    tekst = macro(bron.decode('cp1252').replace('\r\n', '\n')).replace('\n', '\r\n').encode('cp1252')
    gecomprimeerd = comprimeer(tekst)
    assert bytes(olevba.decompress_stream(bytearray(gecomprimeerd))) == tekst
    nieuw = ole_vervang(bin_, 'modRDFO', gecomprimeerd)
    n = olefile.OleFileIO(io.BytesIO(nieuw))
    for st in o.listdir(streams=True):
        if st != ['VBA', 'modRDFO']:
            assert o.openstream(st).read() == n.openstream(st).read(), st
    assert bytes(olevba.decompress_stream(bytearray(n.openstream('VBA/modRDFO').read()))) == tekst
    return nieuw


def werkmap(x):
    for naam in VERDIEPING:
        oud = 'name="%s" sheetId=' % escape(naam, {"'": '&apos;'})
        if oud not in x:
            oud = 'name="%s" sheetId=' % naam
        i = x.index(oud)
        j = x.index('/>', i)
        stuk = x[i:j]
        assert 'state="visible"' in stuk, naam
        x = x[:i] + stuk.replace('state="visible"', 'state="hidden"') + x[j:]
    return x


KNOP = '''<xdr:twoCellAnchor editAs="absolute">
<xdr:from><xdr:col>6</xdr:col><xdr:colOff>57150</xdr:colOff><xdr:row>29</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>31</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="[0]!RDFO_ToonVerdieping" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1009" name="Knop 1009" descr="Toon / verberg verdieping"/><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="540000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="F2B800"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>
<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="72000" tIns="36000" rIns="72000" bIns="36000" rtlCol="0" anchor="ctr"/>
<a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="nl-NL" sz="1100" b="1"><a:solidFill><a:srgbClr val="002060"/></a:solidFill>
<a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>Toon / verberg verdieping</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'''


def start(x):
    rij = re.search(r'<row r="32"[^>]*>.*?</row>', x, re.S).group(0)
    assert 'r="G32"' not in rij
    cel = ('<c r="G32" s="11" t="inlineStr"><is><t>Verdieping: Outlook, Bandbreedte, Kansen en risico\'s, Reserves en Handleiding. '
           'Staan verborgen; de koppelingen hiernaast werken pas na deze knop.</t></is></c>')
    return x.replace(rij, rij[:-len('</row>')] + cel + '</row>')


def grafiek(x):
    if '<barChart>' in x:
        richting = re.search(r'<barDir val="(\w+)"', x).group(1)
        groep = re.search(r'<barChart>.*?<grouping val="(\w+)"', x, re.S).group(1)
        oud = int(re.search(r'<gapWidth val="(\d+)"', x).group(1))
        if richting == 'col':
            nieuw = 50 if oud <= 30 else (100 if groep == 'stacked' else max(oud, 80))
        else:
            nieuw = 60
        x = x.replace('<gapWidth val="%d"/>' % oud, '<gapWidth val="%d"/>' % nieuw)
    for oud, nieuw in (('5AA02C', '12864E'), ('3B7A1E', '12864E'), ('D7372F', 'C0392F')):
        x = x.replace('val="%s"' % oud, 'val="%s"' % nieuw)
    # budgetstreepje: korter en oranje, zodat het totaal erboven leesbaar blijft
    x = re.sub(r'<marker><symbol val="dash"/><size val="20"/>.*?</marker>',
               lambda m: m.group(0).replace('<size val="20"/>', '<size val="15"/>').replace('val="6B7385"', 'val="F08A00"'), x, flags=re.S)
    return x


def main():
    zi = zipfile.ZipFile(SRC)
    assert b'RDFO_ToonVerdieping' not in zi.read('xl/drawings/drawing10.xml'), 'al gedaan'
    zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        n, d = it.filename, zi.read(it.filename)
        if n == 'xl/workbook.xml':
            d = werkmap(d.decode('utf8')).encode('utf8')
        elif n == 'xl/vbaProject.bin':
            d = vba(d)
        elif n == 'xl/drawings/drawing10.xml':
            d = vervang(d.decode('utf8'), '</xdr:wsDr>', KNOP + '</xdr:wsDr>').encode('utf8')
        elif n == 'xl/worksheets/sheet1.xml':
            d = start(d.decode('utf8')).encode('utf8')
        elif n.startswith('xl/charts/chart') and n != 'xl/charts/chart25.xml':
            d = grafiek(d.decode('utf8')).encode('utf8')
        zo.writestr(it, d)
    zo.close()


if __name__ == '__main__':
    main()
