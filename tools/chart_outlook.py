"""Outlook-grafiek omzet per jaar: variant A = prognose naast budget (geclusterd), variant B = alleen prognose."""
NAVY, GEEL, GRIJS, TXTGR = '111F39', 'FEB900', 'A4B3BB', '6B7385'
def _txt(sz, b, kleur):
    return (f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}" b="{1 if b else 0}"><a:solidFill><a:srgbClr val="{kleur}"/></a:solidFill>'
            f'</a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr>')
def _ser(i, tx, cat, val, kleur, lsz, lkleur, lb):
    dl = (f'<c:dLbls><c:numFmt formatCode="#,##0.0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{_txt(lsz, lb, lkleur)}'
          '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}<c:spPr><a:solidFill><a:srgbClr val="{kleur}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>{dl}{cat}{val}</c:ser>')
def chart_xml(series, legend=True, gap=60, overlap=-5):
    """series: lijst van (tx_xml, cat_xml, val_xml, kleur, labelgrootte, labelkleur, vet)"""
    sers = ''.join(_ser(i, *s) for i, s in enumerate(series))
    ax = ('<c:catAx><c:axId val="6001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>'
          '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9DEE8"/></a:solidFill></a:ln></c:spPr>'
          + _txt(900, False, NAVY) + '<c:crossAx val="6002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
          '<c:valAx><c:axId val="6002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="1"/><c:axPos val="l"/>'
          '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
          '<c:crossAx val="6001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    leg = ('<c:legend><c:legendPos val="t"/><c:overlay val="0"/>' + _txt(900, False, NAVY) + '</c:legend>') if legend else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
            f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{sers}<c:gapWidth val="{gap}"/><c:overlap val="{overlap}"/>'
            '<c:axId val="6001"/><c:axId val="6002"/></c:barChart>' + ax + '</c:plotArea>' + leg +
            '<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr></c:chartSpace>')
def refs(sh, name_cell, cat_rng, val_rng):
    return (f'<c:tx><c:strRef><c:f>{sh}!{name_cell}</c:f></c:strRef></c:tx>', f'<c:cat><c:strRef><c:f>{sh}!{cat_rng}</c:f></c:strRef></c:cat>',
            f'<c:val><c:numRef><c:f>{sh}!{val_rng}</c:f></c:numRef></c:val>')
