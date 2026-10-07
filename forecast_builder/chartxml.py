"""Grafiek-XML (DrawingML chart) voor het Dashboard, opgebouwd uit kleine bouwstenen.

Dezelfde bouwers worden gebruikt voor de echte grafieken (reeksen via benoemde bereiken [0]!g_*) en voor
previews (reeksen via directe celbereiken). cache.py herkent de vorm <c:strRef><c:f>REF</c:f></c:strRef>
en <c:numRef><c:f>REF</c:f></c:numRef> om cachewaarden toe te voegen; die vorm dus niet veranderen.
"""

NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
FONT = "Arial"

# kleuren (zelfde familie als het Dashboard)
NAVY = "002060"        # basis / cumulatieve cashflow / getransporteerd
GREEN = "2F9E63"       # opbrengsten / upside
RED_LINE = "C0392F"    # downside
RED_BAR = "E5603B"     # kosten (CVD-veilig naast het groen van opbrengsten)
AMBER = "EDA100"       # scenario (gele lijn)
BLUE = "2A78D6"        # verkocht (lijn)
GREY_LINE = "8A94A6"   # vorige prognose / totaal
BAND = "8FA9D6"        # bandbreedte (transparant)
AREA_NEG = "FCEAE8"    # voorfinanciering
AREA_POS = "E3F3EA"    # positief saldo
REALISATIE = "D9DEE6"  # waas over de gerealiseerde kwartalen (op een eigen 0-1 as, altijd de volle hoogte)
GRID = "EDEFF2"
AXIS = "9AA3AD"
TXT = "666666"
TXT_LIGHT = "7A7A7A"
LABEL_BORDER = "E3E6EA"

# vaste kleuren per woningtype (blok 1..10), ook gebruikt voor de cel-legenda op het Dashboard
TYPEKLEUREN = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948", "1F8A8A", "8C6D3F"]

PT = 12700             # EMU per punt
FMT_ABS = "0;0"        # geen minteken onder de as (vlindergrafiek)
FMT_MLN = '&quot;€&quot;0&quot;M&quot;;&quot;−€&quot;0&quot;M&quot;;&quot;€&quot;0&quot;M&quot;'
AX1, AX2, AX3, AX4 = 111111111, 222222222, 333333333, 444444444


def _rpr(sz, bold=False, color=TXT):
    return (f'<a:defRPr sz="{sz}" b="{1 if bold else 0}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:defRPr>')


def _txpr(sz, bold=False, color=TXT):
    return f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{_rpr(sz, bold, color)}</a:pPr><a:endParaRPr lang="nl-NL"/></a:p></c:txPr>'


def _solid(color, alpha=None):
    a = f'<a:alpha val="{int(alpha * 1000)}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:srgbClr val="{color}">{a}</a:srgbClr></a:solidFill>'


def _ln(w_pt, color, dash="solid"):
    return f'<a:ln w="{int(w_pt * PT)}" cap="rnd">{_solid(color)}<a:prstDash val="{dash}"/><a:round/></a:ln>'


def _tx(naam):
    """Reeksnaam: vaste tekst, of een celverwijzing (begint met '=')."""
    if naam.startswith("="):
        return f'<c:tx><c:strRef><c:f>{naam[1:]}</c:f></c:strRef></c:tx>'
    return f'<c:tx><c:v>{naam}</c:v></c:tx>'


def _cat(ref):
    return f'<c:cat><c:strRef><c:f>{ref}</c:f></c:strRef></c:cat>'


def _val(ref):
    return f'<c:val><c:numRef><c:f>{ref}</c:f></c:numRef></c:val>'


def _no_dlbls():
    return ('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
            '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')


def ser_area(idx, naam, cat, val, fill, alpha=None):
    f = _solid(fill, alpha) if fill else "<a:noFill/>"
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr>{f}<a:ln><a:noFill/></a:ln></c:spPr>'
            f'{_cat(cat)}{_val(val)}</c:ser>')


def _lichte_kleur(kleur):
    """Relatieve luminantie (sRGB) boven 0,4: donkere tekst leest beter dan witte (bv. op het amber van type 4)."""
    def lin(c):
        c = int(kleur[c:c + 2], 16) / 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(0) + 0.7152 * lin(2) + 0.0722 * lin(4) > 0.4


def ser_bar(idx, naam, cat, val, fill, alpha=None, labels=False):
    """Balkreeks; labels=True zet de waarde als klein cijfer midden in elk blokje (0 en 1 blijven leeg: te smal);
    wit op donkere kleuren, donker op lichte."""
    dl = ""
    if labels:
        kleur = "1F2937" if _lichte_kleur(fill) else "FFFFFF"
        dl = (f'<c:dLbls><c:numFmt formatCode="[&lt;2]&quot;&quot;;0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
              f'{_txpr(700, True, kleur)}<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
              f'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr>{_solid(fill, alpha)}<a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>{dl}{_cat(cat)}{_val(val)}</c:ser>')


def ser_line(idx, naam, cat, val, color, w_pt, dash="solid"):
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr>{_ln(w_pt, color, dash)}</c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker>{_cat(cat)}{_val(val)}<c:smooth val="0"/></c:ser>')


def ser_punt(idx, naam_ref, cat, val, color, pos="t", size=7):
    """Los punt met markering en een label (de reeksnaam, uit een cel) in een wit kader."""
    marker = (f'<c:marker><c:symbol val="circle"/><c:size val="{size}"/><c:spPr>{_solid(color)}'
              f'<a:ln w="{int(1.5 * PT)}">{_solid("FFFFFF")}</a:ln></c:spPr></c:marker>')
    dlbls = (f'<c:dLbls><c:spPr>{_solid("FFFFFF")}<a:ln w="{int(0.75 * PT)}">{_solid(LABEL_BORDER)}</a:ln></c:spPr>'
             f'{_txpr(750, True, NAVY)}<c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="0"/>'
             f'<c:showCatName val="0"/><c:showSerName val="1"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam_ref)}<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
            f'{marker}{dlbls}{_cat(cat)}{_val(val)}<c:smooth val="0"/></c:ser>')


def grp_area(sers, ax=(AX1, AX2), grouping="standard"):
    return (f'<c:areaChart><c:grouping val="{grouping}"/><c:varyColors val="0"/>{"".join(sers)}{_no_dlbls()}'
            f'<c:axId val="{ax[0]}"/><c:axId val="{ax[1]}"/></c:areaChart>')


def grp_bar(sers, ax=(AX1, AX2), grouping="clustered", gap=60, overlap=None):
    ov = f'<c:overlap val="{overlap}"/>' if overlap is not None else ""
    return (f'<c:barChart><c:barDir val="col"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>{"".join(sers)}{_no_dlbls()}'
            f'<c:gapWidth val="{gap}"/>{ov}<c:axId val="{ax[0]}"/><c:axId val="{ax[1]}"/></c:barChart>')


def grp_waas(idx, cat, val, naam="Realisatie"):
    """Grijze waas over de gerealiseerde kwartalen: vlak (waarde 1) op een onzichtbare 0-1 as, dus altijd de volle hoogte;
    eindigt precies op het punt 'nu' (het midden van het laatste gerealiseerde kwartaal)."""
    ser = ser_area(idx, naam, cat, val, REALISATIE, alpha=28)
    return grp_area([ser], ax=(AX3, AX4))


def assen_waas():
    return [cat_ax(AX3, AX4, deleted=True), val_ax(AX4, AX3, "General", pos="r", crosses="max", grid=False, vmin=0, vmax=1, zichtbaar=False)]


def grp_line(sers, ax=(AX1, AX2)):
    return (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{"".join(sers)}{_no_dlbls()}<c:marker val="1"/>'
            f'<c:axId val="{ax[0]}"/><c:axId val="{ax[1]}"/></c:lineChart>')


def cat_ax(ax_id, cross_id, deleted=False):
    if deleted:
        return (f'<c:catAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>'
                f'<c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                f'<c:crossAx val="{cross_id}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
                f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="1"/></c:catAx>')
    return (f'<c:catAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
            f'<c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
            f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr><a:ln w="{int(0.75 * PT)}">{_solid(AXIS)}</a:ln></c:spPr>'
            f'{_txpr(750, False, TXT)}<c:crossAx val="{cross_id}"/><c:crosses val="autoZero"/><c:auto val="1"/>'
            f'<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="1"/></c:catAx>')


def cat_ax_nul(ax_id, cross_id):
    """Categorie-as die op 0 ligt, als dunne stippellijn; de labels staan onderaan (tickLblPos low)."""
    return (f'<c:catAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
            f'<c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
            f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr><a:ln w="{int(0.75 * PT)}">{_solid(AXIS)}'
            f'<a:prstDash val="sysDash"/></a:ln></c:spPr>{_txpr(750, False, TXT)}<c:crossAx val="{cross_id}"/>'
            f'<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
            f'<c:noMultiLvlLbl val="1"/></c:catAx>')


def donker(kleur, f=0.72):
    """Zelfde tint, iets donkerder (transportkolommen)."""
    return "".join(f"{int(int(kleur[i:i + 2], 16) * f):02X}" for i in (0, 2, 4))


def val_ax(ax_id, cross_id, fmt, pos="l", crosses="autoZero", grid=True, deleted=False, vmin=None, vmax=None, zichtbaar=True):
    scaling = '<c:scaling><c:orientation val="minMax"/>'
    if vmax is not None:
        scaling += f'<c:max val="{vmax}"/>'
    if vmin is not None:
        scaling += f'<c:min val="{vmin}"/>'
    scaling += '</c:scaling>'
    g = f'<c:majorGridlines><c:spPr><a:ln w="{int(0.5 * PT)}">{_solid(GRID)}</a:ln></c:spPr></c:majorGridlines>' if grid else ""
    lbl = "nextTo" if zichtbaar else "none"
    return (f'<c:valAx><c:axId val="{ax_id}"/>{scaling}<c:delete val="{1 if deleted else 0}"/><c:axPos val="{pos}"/>{g}'
            f'<c:numFmt formatCode="{fmt}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            f'<c:tickLblPos val="{lbl}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>{_txpr(800, False, TXT_LIGHT)}'
            f'<c:crossAx val="{cross_id}"/><c:crosses val="{crosses}"/><c:crossBetween val="between"/></c:valAx>')


def legend(verborgen=()):
    entries = "".join(f'<c:legendEntry><c:idx val="{i}"/><c:delete val="1"/></c:legendEntry>' for i in verborgen)
    return f'<c:legend><c:legendPos val="t"/>{entries}<c:overlay val="0"/>{_txpr(800, False, TXT)}</c:legend>'


# Excel 2016+: #N/A als lege cel tonen, anders zet Excel bij ingeschakelde datalabels de tekst '#N/A' bij elk #N/A-punt
# (lege typeblokken). LibreOffice negeert de extensie.
NA_ALS_LEEG = ('<c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" '
               'xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">'
               '<c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst>')


def chart_space(groepen, assen, legenda="", na_als_leeg=False):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace {NS}><c:date1904 val="0"/><c:lang val="nl-NL"/>'
            f'<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>{"".join(groepen)}{"".join(assen)}'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>{legenda}<c:plotVisOnly val="0"/>'
            f'<c:dispBlanksAs val="gap"/>{NA_ALS_LEEG if na_als_leeg else ""}</c:chart><c:spPr>{_solid("FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="{FONT}"/></a:defRPr></a:pPr>'
            f'<a:endParaRPr lang="nl-NL"/></a:p></c:txPr></c:chartSpace>')


# ---------------------------------------------------------------------------------------------
#  De drie (vier) grafieken. `r` is een dict reeksnaam -> bereikverwijzing; `lbl` dict -> celverwijzing (tekst).
# ---------------------------------------------------------------------------------------------
def cashflow(r, lbl, scenario=True):
    """Cashflow per kwartaal: grijs vlak = gerealiseerde kwartalen, vlakken voorfinanciering/positief saldo,
    balken opbrengsten/kosten, lijnen cumulatieve cashflow, vorige prognose en (optioneel) de gele scenariolijn."""
    kw = r["kwartaal"]
    areas = [ser_area(0, "Voorfinanciering", kw, r["voorfinanciering"], AREA_NEG),
             ser_area(1, "Positief saldo", kw, r["positief_saldo"], AREA_POS)]
    waas = grp_waas(2, kw, r["realisatie"])
    bars = [ser_bar(3, "Opbrengsten", kw, r["opbrengsten"], GREEN), ser_bar(4, "Kosten", kw, r["kosten"], RED_BAR)]
    lines = [ser_line(5, "Cumulatieve cashflow", kw, r["stand"], NAVY, 2.5),
             ser_line(6, "Vorige prognose", kw, r["vorige"], GREY_LINE, 1.5, "sysDash")]
    verborgen = [0, 1, 2]
    nxt = 7
    if scenario:
        lines.append(ser_line(7, "=" + lbl["scenario_naam"], kw, r["scenario"], AMBER, 2.25))
        nxt = 8
    punten = [ser_punt(nxt, "=" + lbl["nu"], kw, r["punt_nu"], NAVY, "t"),
              ser_punt(nxt + 1, "=" + lbl["dal"], kw, r["punt_dal"], RED_LINE, "b"),
              ser_punt(nxt + 2, "=" + lbl["eind_basis"], kw, r["eind_basis"], NAVY, "l")]
    verborgen += [nxt, nxt + 1, nxt + 2]
    nxt += 3
    if scenario:
        punten.append(ser_punt(nxt, "=" + lbl["eind_scenario"], kw, r["eind_scenario"], AMBER, "t"))   # basis-label links, scenario boven
        verborgen.append(nxt)
    groepen = [grp_area(areas), waas, grp_bar(bars, grouping="stacked", gap=55, overlap=100), grp_line(lines + punten)]
    return chart_space(groepen, [cat_ax(AX1, AX2), val_ax(AX2, AX1, FMT_MLN)] + assen_waas(), legend(verborgen))


def band(r, lbl):
    """Scenario's: band tussen downside en upside rond de basis, met gerealiseerde kwartalen als grijs blok."""
    kw = r["kwartaal"]
    areas = [ser_area(0, "Band onder", kw, r["band_onder"], None),
             ser_area(1, "Bandbreedte upside–downside", kw, r["bandbreedte"], BAND, alpha=32)]
    lines = [
        ser_line(3, "Downside", kw, r["downside"], RED_LINE, 1.75),
        ser_line(4, "Upside", kw, r["upside"], GREEN, 1.75),
        ser_line(5, "Basis", kw, r["stand"], NAVY, 3),
        ser_punt(6, "=" + lbl["nu"], kw, r["punt_nu"], NAVY, "t"),
        ser_punt(7, "=" + lbl["dal"], kw, r["punt_dal"], RED_LINE, "b"),
        ser_punt(8, "=" + lbl["eind_up"], kw, r["eind_upside"], GREEN, "t"),
        ser_punt(9, "=" + lbl["eind_down"], kw, r["eind_downside"], RED_LINE, "b"),
    ]
    groepen = [grp_area(areas, grouping="stacked"), grp_waas(2, kw, r["realisatie"]), grp_line(lines)]
    return chart_space(groepen, [cat_ax(AX1, AX2), val_ax(AX2, AX1, FMT_MLN)] + assen_waas(), legend([0, 2, 6, 7, 8, 9]))


def verkoop_kwartaal(r):
    """Verkoop en transport per kwartaal (balken), met gerealiseerde kwartalen als grijs vlak."""
    kw = r["kwartaal"]
    bars = [ser_bar(0, "Verkocht", kw, r["verkocht"], BLUE), ser_bar(1, "Getransporteerd", kw, r["getransporteerd"], NAVY)]
    groepen = [grp_bar(bars, gap=60, overlap=-15), grp_waas(2, kw, r["realisatie"])]
    return chart_space(groepen, [cat_ax(AX1, AX2), val_ax(AX2, AX1, "0")] + assen_waas(), legend([2]))


def verkoop_cumulatief(r, lbl):
    """Verkocht en getransporteerd cumulatief, met het totaal aantal woningen als stippellijn."""
    kw = r["kwartaal"]
    lines = [
        ser_line(1, "Totaal woningen", kw, r["totaal"], GREY_LINE, 1.25, "sysDash"),
        ser_line(2, "Verkocht cumulatief", kw, r["verkocht_cum"], BLUE, 2.25),
        ser_line(3, "Getransporteerd cumulatief", kw, r["getransporteerd_cum"], NAVY, 2.25),
        ser_punt(4, "=" + lbl["uitverkocht"], kw, r["punt_uitverkocht"], BLUE, "t"),
        ser_punt(5, "=" + lbl["alles_transport"], kw, r["punt_alles_transport"], NAVY, "b"),
    ]
    groepen = [grp_waas(0, kw, r["realisatie"]), grp_line(lines)]
    return chart_space(groepen, [cat_ax(AX1, AX2), val_ax(AX2, AX1, "0")] + assen_waas(), legend([0, 4, 5]))


# ---------------------------------------------------------------------------------------------
#  Verkoop en transport per woningtype: twee panelen (verkocht / getransporteerd), zelfde kwartaal-as en typekleuren
# ---------------------------------------------------------------------------------------------
TITEL_VERKOCHT = "Verkocht per kwartaal"
TITEL_TRANSPORT = "Getransporteerd per kwartaal (notarieel)"
TXT_TITEL = "5B6169"       # zelfde grijs als de sectiekoppen op het Dashboard
PANEEL_PLOT_X = 0.030      # linkermarge van het plotgebied (fractie van de breedte): ruimte voor aslabels
PANEEL_PLOT_W = 0.962      # identiek in beide panelen, zodat de kwartalen exact onder elkaar staven
PANEEL_TOP_CM = 0.75       # ruimte voor de titel boven het plotgebied
PANEEL_H_CM = 6.35         # hoogte van één paneel (twaalf Dashboard-rijen van 15 pt)


def titel(tekst, sz=900, color=TXT_TITEL):
    """Grafiektitel linksboven in de grafiek, vaste tekst, Arial vet, zonder overlay."""
    rpr = (f'<a:rPr lang="nl-NL" sz="{sz}" b="1"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
           f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:rPr>')
    return (f'<c:title><c:tx><c:rich><a:bodyPr wrap="none" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="l">'
            f'<a:defRPr sz="{sz}" b="1"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{FONT}"/></a:defRPr></a:pPr>'
            f'<a:r>{rpr}<a:t>{tekst}</a:t></a:r></a:p></c:rich></c:tx>'
            f'<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.006"/><c:y val="0.02"/></c:manualLayout></c:layout>'
            f'<c:overlay val="0"/></c:title>')


def plot_layout(x, y, w, h):
    """Handmatige lay-out van het (binnen)plotgebied als fractie van de grafiek."""
    return (f'<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>'
            f'<c:x val="{x:.4f}"/><c:y val="{y:.4f}"/><c:w val="{w:.4f}"/><c:h val="{h:.4f}"/></c:manualLayout></c:layout>')


def ser_totaal(idx, naam, cat, val, sz=700, color=TXT):
    """Onzichtbare lijnreeks met de waarde als klein cijfer boven elk punt (= boven de stapel); nullen verborgen."""
    dlbls = (f'<c:dLbls><c:numFmt formatCode="0;;" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
             f'{_txpr(sz, False, color)}<c:dLblPos val="t"/><c:showLegendKey val="0"/><c:showVal val="1"/>'
             f'<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker>{dlbls}{_cat(cat)}{_val(val)}<c:smooth val="0"/></c:ser>')


def chart_space_paneel(groepen, assen, titel_xml, layout_xml):
    """chart_space zonder legenda, met titel en handmatig plotgebied; #N/A (lege typeblokken) zonder label."""
    xml = chart_space(groepen, assen, legenda="", na_als_leeg=True)
    oud = '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
    assert oud in xml
    return xml.replace(oud, f'<c:chart>{titel_xml}<c:autoTitleDeleted val="0"/><c:plotArea>{layout_xml}')


def verkoop_paneel(r, lbl, types, soort, hoogte_cm=PANEEL_H_CM, cat_labels=True, gap=45):
    """Eén paneel van de verkoopgrafiek: gestapelde kolommen per woningtype (vaste kleur per blok) met het aantal in elk
    blokje, totaal per kwartaal als klein cijfer boven de stapel, één mijlpaalpunt en de grijze waas over de gerealiseerde kwartalen. Geen legenda:
    de kleur per type staat als chip-cellen boven het eerste paneel (lege typeblokken zouden spookvermeldingen geven).
    soort = 'verkocht' (mijlpaal 'uitverkocht') of 'transport' (mijlpaal 'laatste transport'; getallen positief).
    `types` = [(naamcel, verkocht-bereik, getransporteerd-bereik)] per typeblok.

    Groepsvolgorde balken -> lijnen -> waas: LibreOffice tekent de categorie-as en de rasterlijnen verkeerd als een
    lijngroep op de primaire as na een vlakgroep op de tweede as staat."""
    kw = r["kwartaal"]
    n = len(types)
    if soort == "verkocht":
        bars = [ser_bar(k, "=" + naam, kw, v, TYPEKLEUREN[k % len(TYPEKLEUREN)], labels=True) for k, (naam, v, _) in enumerate(types)]
        totaal, punt, ptxt, tt = r["verkocht"], r["punt_uitverkocht"], lbl["uitverkocht"], TITEL_VERKOCHT
    else:
        bars = [ser_bar(k, "=" + naam, kw, t, TYPEKLEUREN[k % len(TYPEKLEUREN)], labels=True) for k, (naam, _, t) in enumerate(types)]
        totaal, punt, ptxt, tt = r["getransporteerd"], r["punt_alles_transport"], lbl["alles_transport"], TITEL_TRANSPORT
    mijlpaal = ser_punt(n + 1, "=" + ptxt, kw, punt, NAVY, "t", size=5)
    mijlpaal = mijlpaal.replace('<c:txPr><a:bodyPr/>', '<c:txPr><a:bodyPr wrap="none"/>')   # label op één regel
    lines = [ser_totaal(n, "Totaal", kw, totaal), mijlpaal]
    groepen = [grp_bar(bars, grouping="stacked", gap=gap, overlap=100), grp_line(lines), grp_waas(n + 2, kw, r["realisatie"])]
    top = PANEEL_TOP_CM / hoogte_cm
    bottom = (0.45 if cat_labels else 0.12) / hoogte_cm
    layout = plot_layout(PANEEL_PLOT_X, top, PANEEL_PLOT_W, 1 - top - bottom)
    catax = cat_ax(AX1, AX2) if cat_labels else cat_ax(AX1, AX2).replace('<c:tickLblPos val="low"/>', '<c:tickLblPos val="none"/>')
    assen = [catax, val_ax(AX2, AX1, "0", vmin=0)] + assen_waas()
    return chart_space_paneel(groepen, assen, titel(tt), layout)


# ---------------------------------------------------------------------------------------------
#  Bouwtermijnen per woningtype: tijdlijn (horizontale gestapelde balken), één rij per type × termijn
# ---------------------------------------------------------------------------------------------
VORIGE = "F3A6A0"      # vorige prognose (zacht rood)
BT_GAP = 15            # bijna aaneengesloten rijen, zoals een tabel met gekleurde cellen


def ser_bar_onzichtbaar(idx, naam, cat, val):
    """Offset-reeks: geen vulling, geen lijn; schuift het gestapelde blokje naar het juiste beginpunt."""
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>{_cat(cat)}{_val(val)}</c:ser>')


def met_order(ser_xml, order):
    """Geef een reeks een eigen stapelvolgorde (c:order) los van c:idx (legenda-index)."""
    i = ser_xml.index('<c:idx val="') + len('<c:idx val="')
    idx = ser_xml[i:ser_xml.index('"', i)]
    return ser_xml.replace(f'<c:order val="{idx}"/>', f'<c:order val="{order}"/>', 1)


def grp_hbar(sers, ax=(AX1, AX2), gap=BT_GAP):
    """Horizontale gestapelde balken (barDir=bar)."""
    return (f'<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>{"".join(sers)}{_no_dlbls()}'
            f'<c:gapWidth val="{gap}"/><c:overlap val="100"/><c:axId val="{ax[0]}"/><c:axId val="{ax[1]}"/></c:barChart>')


def cat_ax_rijen(ax_id, cross_id, deleted=False):
    """Verticale categorie-as (rijen), eerste rij bovenaan (maxMin); de tijd-as (jaartallen) staat als koprij bovenaan."""
    if deleted:
        return (f'<c:catAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/>'
                f'<c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                f'<c:crossAx val="{cross_id}"/><c:crosses val="max"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
                f'<c:lblOffset val="100"/><c:noMultiLvlLbl val="1"/></c:catAx>')
    return (f'<c:catAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="0"/>'
            f'<c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
            f'<c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
            f'{_txpr(750, False, TXT)}<c:crossAx val="{cross_id}"/><c:crosses val="max"/><c:auto val="1"/>'
            f'<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="1"/></c:catAx>')


def val_ax_tijd(ax_id, cross_id, vmin, vmax=None, zichtbaar=True, grid=True):
    """Horizontale tijd-as in jaren (waarde = jaar + (kw-1)/4): vaste min, stap 1 jaar met lichte rasterlijn, kwartaalraster
    nog lichter; max automatisch (alle stapels lopen tot het einde van de laatste periode) of vast."""
    g = ""
    if grid:
        g = (f'<c:majorGridlines><c:spPr><a:ln w="{int(0.75 * PT)}">{_solid(AXIS)}</a:ln></c:spPr></c:majorGridlines>'
             f'<c:minorGridlines><c:spPr><a:ln w="{int(0.5 * PT)}">{_solid(GRID)}</a:ln></c:spPr></c:minorGridlines>')
    lbl = "nextTo" if zichtbaar else "none"
    mx = f'<c:max val="{vmax}"/>' if vmax is not None else ""
    return (f'<c:valAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="minMax"/>{mx}<c:min val="{vmin}"/></c:scaling>'
            f'<c:delete val="0"/><c:axPos val="b"/>{g}<c:numFmt formatCode="0" sourceLinked="0"/>'
            f'<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="{lbl}"/>'
            f'<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>{_txpr(800, False, TXT_LIGHT)}'
            f'<c:crossAx val="{cross_id}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>'
            f'<c:majorUnit val="1"/><c:minorUnit val="0.25"/></c:valAx>')


def bouwtermijnen(r, jaar_min, jaar_max=None):
    """Tijdlijn van de bouwtermijnen: per rij (één bouwtermijn, gedeeld door de typen die hem hebben) een blauw blokje van het
    vroegste tot het laatste kwartaal waarin de termijn in de huidige planning vervalt (één kwartaal als alle typen gelijk
    lopen), een rood blokje voor de vorige prognose (zelfde plek als blauw zonder vorige prognose: dan onzichtbaar) en een
    grijze waas over de gerealiseerde kwartalen. Twee gestapelde balkgroepen: stapel 1 (eerste as) onzichtbaar tot de
    eerste periode, grijs, onzichtbaar, rood, grijs, onzichtbaar tot het einde; stapel 2 (tweede as, bovenop) onzichtbaar,
    blauw, onzichtbaar tot het einde. Alle stapels zijn even lang, zodat beide assen gelijk schalen. De jaartallen staan
    bovenaan als koprij (zoals in een planningstabel): de categorie-as loopt van boven naar beneden (maxMin) en de tijd-as
    kruist bij de eerste rij.

    r: dict met bereikverwijzingen 'label' (categorie) en 'seg0' … 'seg8' (waarden in jaren). jaar_min: eerste jaar
    (vaste as-ondergrens; balken beginnen op 0); jaar_max: vaste bovengrens of None (automatisch).
    Reeksnummering: XML-positie p = stapelvolgorde (c:order), c:idx = 8 − p. Excel koppelt legendEntry-idx aan c:idx,
    LibreOffice legt de legenda van gestapelde balken in omgekeerde XML-volgorde aan; met idx = 8 − p verbergen beide
    dezelfde reeksen. Zichtbaar in de legenda: huidige planning (idx 1), vorige prognose (idx 5), gerealiseerd (idx 7)."""
    cat = r["label"]
    stapel1 = [met_order(ser_bar_onzichtbaar(8, "·", cat, r["seg0"]), 0),
               met_order(ser_bar(7, "Gerealiseerd", cat, r["seg1"], REALISATIE, alpha=60), 1),
               met_order(ser_bar_onzichtbaar(6, "·", cat, r["seg2"]), 2),
               met_order(ser_bar(5, "Vorige prognose", cat, r["seg3"], VORIGE), 3),
               met_order(ser_bar(4, "Gerealiseerd", cat, r["seg4"], REALISATIE, alpha=60), 4),
               met_order(ser_bar_onzichtbaar(3, "·", cat, r["seg5"]), 5)]
    stapel2 = [met_order(ser_bar_onzichtbaar(2, "·", cat, r["seg6"]), 6),
               met_order(ser_bar(1, "Huidige planning", cat, r["seg7"], BLUE), 7),
               met_order(ser_bar_onzichtbaar(0, "·", cat, r["seg8"]), 8)]
    assen = [cat_ax_rijen(AX1, AX2), val_ax_tijd(AX2, AX1, jaar_min, jaar_max),
             cat_ax_rijen(AX3, AX4, deleted=True), val_ax_tijd(AX4, AX3, jaar_min, jaar_max, zichtbaar=False, grid=False)]
    groepen = [grp_hbar(stapel1), grp_hbar(stapel2, ax=(AX3, AX4))]
    return chart_space(groepen, assen, legend([0, 2, 3, 4, 6, 8]), na_als_leeg=True)
