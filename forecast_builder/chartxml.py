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


def ser_bar(idx, naam, cat, val, fill, alpha=None):
    return (f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{_tx(naam)}<c:spPr>{_solid(fill, alpha)}<a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:invertIfNegative val="0"/>{_cat(cat)}{_val(val)}</c:ser>')


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


def chart_space(groepen, assen, legenda=""):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace {NS}><c:date1904 val="0"/><c:lang val="nl-NL"/>'
            f'<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>{"".join(groepen)}{"".join(assen)}'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>{legenda}<c:plotVisOnly val="0"/>'
            f'<c:dispBlanksAs val="gap"/></c:chart><c:spPr>{_solid("FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>'
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
        punten.append(ser_punt(nxt, "=" + lbl["eind_scenario"], kw, r["eind_scenario"], AMBER, "b"))
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


def verkoop_types(r, lbl, types):
    """Vlinder per woningtype: verkocht per kwartaal als gestapelde kolommen omhoog, getransporteerd (negatief in het
    Model) als gestapelde kolommen omlaag in een donkerder tint van dezelfde typekleur. Geen legenda in de grafiek
    (lege typeblokken zouden spookvermeldingen geven): de legenda staat als gekleurde cellen boven de grafiek.
    `types` = [(naamcel, verkocht-bereik, transport-bereik)] per typeblok, vaste kleur per blok.

    Groepsvolgorde balken -> punten -> waas: LibreOffice tekent de categorie-as en de rasterlijnen verkeerd als een
    lijngroep op de primaire as na een vlakgroep op de tweede as staat."""
    kw = r["kwartaal"]
    n = len(types)
    bars = [ser_bar(k, "=" + naam, kw, v, TYPEKLEUREN[k % len(TYPEKLEUREN)]) for k, (naam, v, _) in enumerate(types)]
    bars += [ser_bar(n + k, "=" + naam, kw, t, donker(TYPEKLEUREN[k % len(TYPEKLEUREN)])) for k, (naam, _, t) in enumerate(types)]
    punten = [ser_punt(2 * n, "=" + lbl["uitverkocht"], kw, r["punt_uitverkocht"], NAVY, "t"),
              ser_punt(2 * n + 1, "=" + lbl["alles_transport"], kw, r["punt_alles_transport"], NAVY, "b")]
    groepen = [grp_bar(bars, grouping="stacked", gap=55, overlap=100), grp_line(punten), grp_waas(2 * n + 2, kw, r["realisatie"])]
    return chart_space(groepen, [cat_ax_nul(AX1, AX2), val_ax(AX2, AX1, FMT_ABS)] + assen_waas())
