"""Bandbreedte-grafiek (staand): per jaar een zwevende kolom van ondergrens tot bovengrens, rood onder basis,
groen boven basis, navy streep = basis; labels: ondergrens onder, basis rechts, bovengrens boven.
chart_xml()            -> grafiek op Bandbreedte!Y14:AF17 (Excel)
chart_xml(ppt=waarden) -> zelfde grafiek op Sheet1 van de PowerPoint-werkmap, met opgeslagen waarden."""
NAVY, RED, GREEN = '111F39', 'E30613', '3B7A1E'
ROLES = ['onder', 'down', 'lijn', 'up', 'laag', 'basis', 'hoog']
XL = dict(onder='Z', down='AA', lijn='AF', up='AB', laag='AC', basis='AD', hoog='AE')
NAMEN = dict(onder='onder', down='Downside (ondergrens tot basis)', lijn='Basis (prognose)', up='Upside (basis tot bovengrens)',
             laag='Ondergrens', basis='Basis', hoog='Bovengrens')

def _txt(sz, b, kleur):
    return (f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}" b="{1 if b else 0}"><a:solidFill><a:srgbClr val="{kleur}"/></a:solidFill>'
            f'</a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr>')

def _refs(role, ppt):
    if ppt is None:
        c = XL[role]
        return (f'<c:tx><c:strRef><c:f>Bandbreedte!${c}$14</c:f></c:strRef></c:tx>',
                '<c:cat><c:strRef><c:f>Bandbreedte!$Y$15:$Y$17</c:f></c:strRef></c:cat>',
                f'<c:val><c:numRef><c:f>Bandbreedte!${c}$15:${c}$17</c:f></c:numRef></c:val>')
    i = ROLES.index(role); col = 'BCDEFGH'[i]; cats, vals = ppt['cats'], ppt[role]; n = len(cats)
    sc = ''.join(f'<c:pt idx="{k}"><c:v>{v}</c:v></c:pt>' for k, v in enumerate(cats))
    nc = ''.join(f'<c:pt idx="{k}"><c:v>{v}</c:v></c:pt>' for k, v in enumerate(vals) if v is not None)
    return (f'<c:tx><c:strRef><c:f>Sheet1!${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{NAMEN[role]}</c:v></c:pt></c:strCache></c:strRef></c:tx>',
            f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${n + 1}</c:f><c:strCache><c:ptCount val="{n}"/>{sc}</c:strCache></c:strRef></c:cat>',
            f'<c:val><c:numRef><c:f>Sheet1!${col}$2:${col}${n + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{n}"/>{nc}</c:numCache></c:numRef></c:val>')

def _bar(i, role, fill, ppt):
    tx, cat, val = _refs(role, ppt)
    sp = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' if fill is None else f'<c:spPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
    return f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}{sp}<c:invertIfNegative val="0"/>{cat}{val}</c:ser>'

def _line(i, role, pos, kleur, sz, dx, ppt):
    tx, cat, val = _refs(role, ppt)
    lbl = lambda extra='': (f'<c:numFmt formatCode="#,##0.0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{_txt(sz, True, kleur)}'
                            f'<c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
                            '<c:showPercent val="0"/><c:showBubbleSize val="0"/>')
    per = ''.join(f'<c:dLbl><c:idx val="{k}"/><c:layout><c:manualLayout><c:x val="{dx}"/><c:y val="0"/></c:manualLayout></c:layout>{lbl()}</c:dLbl>'
                  for k in range(3)) if dx else ''
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker><c:dLbls>{per}{lbl()}</c:dLbls>{cat}{val}<c:smooth val="0"/></c:ser>')

def chart_xml(ppt=None):
    bars = _bar(0, 'onder', None, ppt) + _bar(1, 'down', RED, ppt) + _bar(2, 'lijn', NAVY, ppt) + _bar(3, 'up', GREEN, ppt)
    lines = (_line(4, 'laag', 'b', RED, 900, 0, ppt) + _line(5, 'basis', 'r', NAVY, 1000, 0.06, ppt)
             + _line(6, 'hoog', 't', GREEN, 900, 0, ppt))
    ax = ('<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>'
          '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="111F39"/></a:solidFill></a:ln></c:spPr>'
          + _txt(1000, True, NAVY) + '<c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
          '<c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9DEE8"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
          '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
          + _txt(800, False, '6B7385') + '<c:crossAx val="5001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    leg = ''.join(f'<c:legendEntry><c:idx val="{i}"/><c:delete val="1"/></c:legendEntry>' for i in (0, 4, 5, 6))
    ext = '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' if ppt is not None else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
            '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>' + bars +
            '<c:gapWidth val="170"/><c:overlap val="100"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>'
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + lines + '<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>'
            + ax + '</c:plotArea><c:legend><c:legendPos val="b"/>' + leg + '<c:overlay val="0"/>' + _txt(900, False, NAVY) + '</c:legend>'
            '<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr>' + ext + '</c:chartSpace>')
