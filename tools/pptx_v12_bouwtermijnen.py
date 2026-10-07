#!/usr/bin/env python3
"""Kwartaalsjabloon v11 -> v12: bouwtermijnen-tijdlijn op dia 5, nieuwe dia 6 'Verkoop en transport per kwartaal'.

Gebruik:
    python3 tools/pptx_v12_bouwtermijnen.py INVOER.pptx UITVOER.pptx [UITVOER2.pptx ...] [--voorbeeld v7.xlsx]

Bewerking (python-pptx + lxml, zelfde werkwijze als tools/pptx_rente_card.py en tools/pptx_scenario_line.py):
  * Dia 5: de grafiek VT_GRAFIEK (verkocht/getransporteerd per kwartaal) wordt verwijderd (vorm, relatie, chart-part en
    ingebed werkboek) en op dezelfde plek en maat vervangen door BT_GRAFIEK: de bouwtermijnen-tijdlijn uit
    forecast_builder.chartxml.bouwtermijnen (horizontale gestapelde balken, een rij per woningtype x bouwtermijn, blauw =
    huidige planning, rood = vorige prognose, grijs = gerealiseerd; vaste tijd-as 2023..2031 als voorbeeld). De grafiek
    wordt eerst met python-pptx aangemaakt (chart-part, rels en ingebed werkboek kloppen dan), daarna wordt de chart-XML
    vervangen en het ingebedde werkboek (Sheet1) met openpyxl geschreven: rij 1 de koppen van het blok AA7:AJ7 van tab
    PowerPoint (A Termijn, B..J seg0..seg8), 20 voorbeeldrijen (4 typen x 5 termijnen), leeg t/m rij 101.
    VT_TITEL/VT_SUBTITEL heten nu BT_TITEL/BT_SUBTITEL, VT_KICKER heet BT_KICKER; VT_KAART, VT_TABELKAART en VT_TABEL
    blijven (SchikDia5 in de VBA gebruikt die namen).
  * Nieuwe dia 6 (kopie van de tekstvormen van dia 5, zelfde lay-out): VP_KICKER, VP_TITEL, VP_SUBTITEL en VP_KAART
    (grafiekkaart over de volle hoogte, tot de onderkant van de oude tabelkaart), met twee grafieken boven elkaar:
    VP1_GRAFIEK (verkocht per woningtype) en VP2_GRAFIEK (getransporteerd per woningtype) uit chartxml.verkoop_paneel.
    Beide lezen hetzelfde blok AK7:BJ.. (26 kolommen); het ingebedde werkboek krijgt de voorbeeldgegevens uit het
    doorgerekende werkboek (--voorbeeld, tab Model, 32 kwartalen), leeg t/m rij 61.
  * De oude dia 6 (VO_*) wordt dia 7. Dianummers in de voettekst zijn velden (<a:fld type="slidenum">); de cachetekst
    wordt op het nieuwe dianummer gezet.
  * De str-/numCache van elke reeks wordt gevuld met de voorbeeldwaarden uit het ingebedde werkboek (ptCount = aantal
    voorbeeldrijen), zodat de dia zonder herberekening een kloppende grafiek toont (ook in LibreOffice).

Afspraak met de VBA-macro (VulGrafiek): de macro kopieert kop + n rijen naar A1 van het gegevensblad en zet per reeks het
X-bereik (kolom A) en het Y-bereik op de kolomletter uit de SERIES-formule; voor BT_GRAFIEK moet de macro daarna beide
waarde-assen (eerste en tweede as) op Model!BY100 (ondergrens) en Model!BY101 + 1 (bovengrens) zetten.
"""
import argparse
import copy
import io
import os
import re
import shutil
import subprocess
import sys
import zipfile

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.util import Emu

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from forecast_builder import chartxml as cx  # noqa: E402

NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EMU_CM = 360000
DIA_BT = 4                      # dia 5 (0-gebaseerd)
DIA_VP_POS = 5                  # nieuwe dia 6 komt op positie 5 in de sldIdLst

# ---- dia 5: bouwtermijnen ---------------------------------------------------------------------------------------------
BT_KOPPEN = ["Termijn", "", "Gerealiseerd", "", "Vorige prognose", "Gerealiseerd", "", "", "Huidige planning", ""]
BT_RIJEN = 100                  # rijen 2..101 in het gegevensblad (Model!BY96 = aantal gevulde rijen)
BT_JAAR_MIN, BT_JAAR_MAX = 2023, 2031           # voorbeeld: Model!BY100 en Model!BY101 + 1
BT_T_FIRST, BT_T_END, BT_T_NU_END = 2023.0, 2031.0, 2026.75   # eerste periode, einde laatste periode, einde Q3 '26
BT_BLOK = 0.25
BT_TERMIJNEN = [("Start bouw", 0), ("Fundering gereed", 1), ("Casco gereed", 3), ("Wind- en waterdicht", 5), ("Oplevering", 8)]
# (type, start huidige planning in jaren, verschuiving vorige prognose in kwartalen: negatief = vorige prognose eerder)
BT_TYPEN = [("Rijwoning", 2027.0, -1), ("Tweekapper", 2027.5, 0), ("Vrijstaand", 2028.25, 2), ("Appartement", 2025.5, 1)]
BT_TEKST = {"BT_KICKER": "BOUWTERMIJNEN PER WONINGTYPE", "BT_TITEL": "Bouwtermijnen per woningtype",
            "BT_SUBTITEL": "blauw = huidige planning · rood = vorige prognose · grijs = gerealiseerd"}
BT_REFS = {"label": f"Sheet1!$A$2:$A${BT_RIJEN + 1}"}
BT_REFS.update({f"seg{i}": f"Sheet1!${get_column_letter(i + 2)}$2:${get_column_letter(i + 2)}${BT_RIJEN + 1}" for i in range(9)})

# ---- dia 6: verkoop en transport per kwartaal -----------------------------------------------------------------------
VP_RIJEN = 60                   # rijen 2..61 (MAX_RIJEN periodes)
VP_TYPEN = ["Rijwoning", "Tweekapper", "Vrijstaand", "Appartement"] + [f"Type {k}" for k in range(5, 11)]
VP_KOPPEN = ["Kwartaal"] + VP_TYPEN + VP_TYPEN + ["Verkocht", "Getransporteerd", "Uitverkocht Q1 ’30", "Laatste transport Q3 ’30", "Realisatie"]
VP_TEKST = {"VP_KICKER": "VERKOOP EN TRANSPORT", "VP_TITEL": "Verkoop en transport per kwartaal",
            "VP_SUBTITEL": "per woningtype · cijfer boven de stapel = totaal"}
_c = {n: get_column_letter(i + 1) for i, n in enumerate(["kwartaal"] + [f"v{k}" for k in range(1, 11)] + [f"t{k}" for k in range(1, 11)]
                                                         + ["verkocht", "getransporteerd", "punt_uitverkocht", "punt_alles_transport", "realisatie"])}
assert _c["realisatie"] == "Z" and _c["v1"] == "B" and _c["t1"] == "L" and _c["verkocht"] == "V"


def _rng(col):
    return f"Sheet1!${col}$2:${col}${VP_RIJEN + 1}"


VP_REFS = {n: _rng(_c[n]) for n in ["kwartaal", "verkocht", "getransporteerd", "punt_uitverkocht", "punt_alles_transport", "realisatie"]}
VP_LBL = {"uitverkocht": f"Sheet1!${_c['punt_uitverkocht']}$1", "alles_transport": f"Sheet1!${_c['punt_alles_transport']}$1"}
VP_TYPES = [(f"Sheet1!${_c[f'v{k}']}$1", _rng(_c[f"v{k}"]), _rng(_c[f"t{k}"])) for k in range(1, 11)]

VOORBEELD_STANDAARD = ("/tmp/claude-0/-home-user-Forecast-Builder/f57f376f-c2b1-5dd0-aed4-c5cc02f799d8/scratchpad/out/v7.xlsx")


def q(tag):
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


def shapes_by_name(slide):
    return {sh.name: sh for sh in slide.shapes}


def set_text_keep_format(shape, text):
    """Vervang de tekst van de eerste run; overige runs/alinea's verwijderen (zoals pptx_rente_card.py)."""
    tf = shape.text_frame
    p = tf.paragraphs[0]
    runs = p.runs
    if not runs:
        p.text = text
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)


def hernoem(shape, naam):
    shape._element.find(".//" + q("p:cNvPr")).set("name", naam)


# ---------------------------------------------------------------------------------------------------------------------
#  grafiek-parts
# ---------------------------------------------------------------------------------------------------------------------
def verwijder_grafiek(slide, naam):
    """Verwijdert het grafiekvak `naam` (vorm + relatie); chart-part en ingebed werkboek zijn daarna onbereikbaar en
    worden bij het opslaan niet meer weggeschreven. Geeft (x, y, w, h) terug."""
    sh = shapes_by_name(slide).get(naam)
    if sh is None or not sh.has_chart:
        sys.exit(f"{naam} niet gevonden als grafiek op dia {slide.slide_id}.")
    box = (sh.left, sh.top, sh.width, sh.height)
    rId = sh._element.find(".//" + q("c:chart")).get(q("r:id"))
    oud_part = slide.part.related_part(rId)
    sh._element.getparent().remove(sh._element)
    slide.part.rels.pop(rId)
    return box, oud_part.partname


def nieuwe_grafiek(slide, naam, box, chart_type, n_cat, n_ser, partnaam, werkbladnaam):
    """Maakt met python-pptx een grafiek met plaatsvervangende gegevens (chart-part, rels, ingebed werkboek) en geeft
    het vak een naam en vaste partnamen."""
    cd = CategoryChartData()
    cd.categories = [f"r{i}" for i in range(n_cat)]
    for s in range(n_ser):
        cd.add_series(f"s{s}", [1] * n_cat)
    gf = slide.shapes.add_chart(chart_type, Emu(box[0]), Emu(box[1]), Emu(box[2]), Emu(box[3]), cd)
    hernoem(gf, naam)
    part = gf.chart.part
    part.partname = PackURI(f"/ppt/charts/{partnaam}.xml")
    part.chart_workbook.xlsx_part.partname = PackURI(f"/ppt/embeddings/{werkbladnaam}.xlsx")
    return gf


def splits_ref(f):
    """'Sheet1!$B$2:$B$101' -> ('Sheet1', 'B', 2, 101); 'Sheet1!$B$1' -> ('Sheet1', 'B', 1, 1)"""
    blad, bereik = f.rsplit("!", 1)
    delen = bereik.replace("$", "").split(":")
    m = re.fullmatch(r"([A-Z]+)(\d+)", delen[0])
    kol, r1 = m.group(1), int(m.group(2))
    r2 = int(re.fullmatch(r"([A-Z]+)(\d+)", delen[-1]).group(2))
    return blad, kol, r1, r2


def vul_caches(cs, ws, n):
    """Zet per <c:strRef>/<c:numRef> een cache met de voorbeeldwaarden uit het werkblad: één cel -> ptCount 1,
    kolombereik -> ptCount n (de voorbeeldrijen), #N/A en lege cellen zonder punt."""
    for ref in list(cs.iter(q("c:strRef"), q("c:numRef"))):
        f = ref.find(q("c:f")).text
        blad, kol, r1, r2 = splits_ref(f)
        assert blad == ws.title, (f, ws.title)
        rijen = range(r1, r1 + 1) if r1 == r2 else range(r1, r1 + n)
        waarden = [ws[f"{kol}{r}"].value for r in rijen]
        for oud in ref.findall(q("c:strCache")) + ref.findall(q("c:numCache")):
            ref.remove(oud)
        if ref.tag == q("c:strRef"):
            cache = etree.SubElement(ref, q("c:strCache"))
            etree.SubElement(cache, q("c:ptCount")).set("val", str(len(waarden)))
            for i, w in enumerate(waarden):
                if w in (None, "") or (isinstance(w, str) and w.startswith("#")):
                    continue
                pt = etree.SubElement(cache, q("c:pt"))
                pt.set("idx", str(i))
                etree.SubElement(pt, q("c:v")).text = str(w)
        else:
            cache = etree.SubElement(ref, q("c:numCache"))
            etree.SubElement(cache, q("c:formatCode")).text = "General"
            etree.SubElement(cache, q("c:ptCount")).set("val", str(len(waarden)))
            for i, w in enumerate(waarden):
                if not isinstance(w, (int, float)) or isinstance(w, bool):
                    continue
                pt = etree.SubElement(cache, q("c:pt"))
                pt.set("idx", str(i))
                etree.SubElement(pt, q("c:v")).text = repr(float(w)) if isinstance(w, float) else str(w)


def zet_chart_xml(gf, xml, ws, n):
    """Vervangt de chart-XML van grafiekvak `gf` door `xml` (chartxml), behoudt de koppeling met het ingebedde werkboek
    (<c:externalData r:id>), vult de caches en schrijft het werkblad `ws` als ingebed werkboek."""
    part = gf.chart.part
    oud = part._element
    ext = oud.find(q("c:externalData"))
    assert ext is not None, "python-pptx-grafiek zonder externalData"
    cs = parse_xml(xml.encode("utf-8"))
    nieuw_ext = etree.SubElement(cs, q("c:externalData"))
    nieuw_ext.set(q("r:id"), ext.get(q("r:id")))
    etree.SubElement(nieuw_ext, q("c:autoUpdate")).set("val", "0")
    vul_caches(cs, ws, n)
    part._element = cs
    buf = io.BytesIO()
    ws.parent.save(buf)
    part.chart_workbook.update_from_xlsx_blob(buf.getvalue())


# ---------------------------------------------------------------------------------------------------------------------
#  voorbeeldgegevens
# ---------------------------------------------------------------------------------------------------------------------
def bt_rij(t_blue, t_red):
    """Negen segmenten (jaren) voor één rij; stapel 1 = seg0..seg5 (eerste as), stapel 2 = seg6..seg8 (tweede as).
    Beide stapels lopen van 0 tot BT_T_END."""
    a, e, nu, b = BT_T_FIRST, BT_T_END, BT_T_NU_END, BT_BLOK
    r = t_red if t_red is not None else t_blue
    rood = b if t_red is not None and abs(t_red - t_blue) > 1e-9 else 0.0   # geen vorige prognose of gelijk: onzichtbaar
    seg1 = max(0.0, min(r, nu) - a)                 # grijs vóór het rode blokje
    seg2 = r - a - seg1                             # onzichtbaar tot het rode blokje
    seg3 = rood
    seg4 = max(0.0, min(nu, e) - (r + rood))        # grijs na het rode blokje
    seg5 = e - (r + rood) - seg4                    # onzichtbaar tot het einde
    seg6 = t_blue
    seg7 = b
    seg8 = e - t_blue - b
    segs = [a, seg1, seg2, seg3, seg4, seg5, seg6, seg7, seg8]
    assert abs(sum(segs[:6]) - e) < 1e-9 and abs(sum(segs[6:]) - e) < 1e-9, segs
    return [round(s, 6) for s in segs]


def werkboek_bt():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(BT_KOPPEN)
    for naam, start, schuif in BT_TYPEN:
        for termijn, kw in BT_TERMIJNEN:
            t_blue = start + kw * BT_BLOK
            t_red = t_blue + schuif * BT_BLOK if schuif else t_blue
            if naam == "Appartement" and t_blue < BT_T_NU_END:
                t_red = t_blue                      # gerealiseerde termijnen: vorige prognose gelijk
            ws.append([f"{naam} · {termijn}"] + bt_rij(t_blue, t_red))
    n = ws.max_row - 1
    ws.column_dimensions["A"].width = 28
    return ws, n


def werkboek_vp(pad_voorbeeld):
    """Voorbeeldgegevens uit het doorgerekende werkboek (tab Model): 32 kwartalen (rijen 8..39)."""
    from forecast_builder import layout as LY
    wb_m = load_workbook(pad_voorbeeld, data_only=True)
    ws_m = wb_m["Model"]
    n = int(ws_m[LY.h("n").replace("$", "")].value)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    koppen = list(VP_KOPPEN)
    for lbl, kop in (("lbl_uitverkocht", 23), ("lbl_alles_transport", 24)):
        w = ws_m[LY.h(lbl).replace("$", "")].value
        if isinstance(w, str) and w:
            koppen[kop] = w
    ws.append(koppen)

    def cel(col, r):
        v = ws_m[f"{col}{r}"].value
        return "#N/A" if v is None else v

    for i in range(n):
        r = LY.ROW1 + i
        rij = [cel(LY.FIX["kwartaal"], r)]
        rij += [cel(LY.m_col("gv", k), r) for k in range(1, LY.N_TYPES + 1)]
        rij += [cel(LY.m_col("gt", k), r) for k in range(1, LY.N_TYPES + 1)]
        rij += [cel(LY.M[naam], r) for naam in ("g_verkocht", "g_getransporteerd", "g_punt_uitverkocht", "g_punt_alles_transport", "g_realisatie")]
        ws.append(rij)
    assert ws.max_column == 26 and ws.max_row == n + 1, (ws.max_column, ws.max_row)
    return ws, n


# ---------------------------------------------------------------------------------------------------------------------
#  dia's
# ---------------------------------------------------------------------------------------------------------------------
def dia5_bouwtermijnen(prs):
    slide = prs.slides[DIA_BT]
    by_name = shapes_by_name(slide)
    for oud, nieuw in (("VT_KICKER", "BT_KICKER"), ("VT_TITEL", "BT_TITEL"), ("VT_SUBTITEL", "BT_SUBTITEL")):
        hernoem(by_name[oud], nieuw)
        set_text_keep_format(by_name[oud], BT_TEKST[nieuw])
    box, oud_part = verwijder_grafiek(slide, "VT_GRAFIEK")
    gf = nieuwe_grafiek(slide, "BT_GRAFIEK", box, XL_CHART_TYPE.BAR_STACKED, 20, 9, "chart_bt", "Werkblad_bt")
    # grafiek vóór de tabelkaart in de spTree houden (zelfde z-volgorde als VT_GRAFIEK had)
    anker = by_name["VT_TABELKAART"]._element
    anker.addprevious(gf._element)
    ws, n = werkboek_bt()
    zet_chart_xml(gf, cx.bouwtermijnen(BT_REFS, BT_JAAR_MIN, BT_JAAR_MAX), ws, n)
    return box, oud_part, n


def dia6_verkoop(prs, pad_voorbeeld):
    bron = prs.slides[DIA_BT]
    slide = prs.slides.add_slide(bron.slide_layout)
    for sh in list(slide.shapes):
        sh._element.getparent().remove(sh._element)          # lege plaatsvervangers van de lay-out
    tree = slide.shapes._spTree
    bron_namen = shapes_by_name(bron)
    kaart_onder = bron_namen["VT_TABELKAART"].top + bron_namen["VT_TABELKAART"].height
    for el in bron.shapes._spTree:
        if el.tag != q("p:sp"):
            continue                                        # grafiek en tabel niet meenemen
        naam = el.find(".//" + q("p:cNvPr")).get("name")
        if naam == "VT_TABELKAART":
            continue
        tree.append(copy.deepcopy(el))
    by_name = shapes_by_name(slide)
    for oud, nieuw in (("VT_KICKER", "VP_KICKER"), ("VT_TITEL", "VP_TITEL"), ("VT_SUBTITEL", "VP_SUBTITEL"), ("VT_KAART", "VP_KAART")):
        hernoem(by_name[oud], nieuw)
        if nieuw in VP_TEKST:
            set_text_keep_format(by_name[oud], VP_TEKST[nieuw])
    kaart = by_name["VT_KAART"]
    kaart.height = kaart_onder - kaart.top
    # twee panelen: zelfde binnenmarges als VT_GRAFIEK in VT_KAART had (boven 54864, onder 36576 EMU)
    g = bron_namen["VT_GRAFIEK"]
    boven = g.top - bron_namen["VT_KAART"].top
    onder = (bron_namen["VT_KAART"].top + bron_namen["VT_KAART"].height) - (g.top + g.height)
    y0 = kaart.top + boven
    hoogte = (kaart.top + kaart.height - onder - y0) // 2
    ws, n = werkboek_vp(pad_voorbeeld)
    frames = []
    for i, (naam, soort, part, wbnaam) in enumerate((("VP1_GRAFIEK", "verkocht", "chart_vp1", "Werkblad_vp1"),
                                                     ("VP2_GRAFIEK", "transport", "chart_vp2", "Werkblad_vp2"))):
        box = (g.left, y0 + i * hoogte, g.width, hoogte)
        gf = nieuwe_grafiek(slide, naam, box, XL_CHART_TYPE.COLUMN_STACKED, n, 13, part, wbnaam)
        xml = cx.verkoop_paneel(VP_REFS, VP_LBL, VP_TYPES, soort, hoogte_cm=hoogte / EMU_CM)
        zet_chart_xml(gf, xml, ws_kopie(ws), n)
        frames.append(gf)
    # grafieken vóór het dianummer in de spTree
    nummer = by_name["Slide Number Placeholder 0"]._element
    for gf in frames:
        nummer.addprevious(gf._element)
    # dia op positie 6 (na dia 5)
    lst = prs.slides._sldIdLst
    el = lst[-1]
    lst.remove(el)
    lst.insert(DIA_VP_POS, el)
    return slide, n, hoogte


def ws_kopie(ws):
    """Elke grafiek krijgt een eigen werkboek-blob; openpyxl-werkbladen zijn niet deelbaar tussen opslagbeurten,
    dus de gegevens worden in een nieuw werkboek overgenomen."""
    wb = Workbook()
    w = wb.active
    w.title = ws.title
    for rij in ws.iter_rows(values_only=True):
        w.append(list(rij))
    return w


def dianummers_bijwerken(prs):
    """De dianummers zijn velden; de cachetekst op het echte nummer zetten (PowerPoint werkt velden zelf bij)."""
    velden = 0
    for i, slide in enumerate(prs.slides, 1):
        for fld in slide.shapes._spTree.iter(q("a:fld")):
            if fld.get("type") == "slidenum":
                t = fld.find(q("a:t"))
                if t is not None:
                    t.text = str(i)
                velden += 1
    return velden


# ---------------------------------------------------------------------------------------------------------------------
#  controle
# ---------------------------------------------------------------------------------------------------------------------
def controleer(pad, verwacht):
    fouten = []
    with zipfile.ZipFile(pad) as z:
        slecht = z.testzip()
        if slecht:
            fouten.append(f"zip: {slecht} kapot")
        namen = z.namelist()
        ct = z.read("[Content_Types].xml").decode("utf-8")
        for naam in namen:
            if naam.startswith("ppt/charts/chart") and naam.endswith(".xml"):
                etree.fromstring(z.read(naam))
                if f'PartName="/{naam}"' not in ct:
                    fouten.append(f"{naam} niet in [Content_Types].xml")
            if naam.endswith(".rels"):
                etree.fromstring(z.read(naam))
        for oud in ("ppt/charts/chart_vt.xml", "ppt/embeddings/Werkblad_vt.xlsx"):
            if oud in namen:
                fouten.append(f"{oud} is niet verwijderd")
        if 'Extension="xlsx"' not in ct:
            fouten.append("geen Default voor xlsx in [Content_Types].xml")
        charts = sorted(n for n in namen if n.startswith("ppt/charts/") and n.endswith(".xml"))
        emb = sorted(n for n in namen if n.startswith("ppt/embeddings/"))
    prs = Presentation(pad)
    for dia, namen_verwacht in verwacht.items():
        slide = prs.slides[dia - 1]
        aanwezig = {sh.name: sh for sh in slide.shapes}
        for naam in namen_verwacht:
            if naam not in aanwezig:
                fouten.append(f"dia {dia}: {naam} ontbreekt")
        n_charts = sum(1 for sh in slide.shapes if sh.has_chart)
        print(f"  dia {dia}: {n_charts} grafiek(en); vormen: " + ", ".join(aanwezig))
    return fouten, charts, emb


def render(pad, map_uit):
    """Dia's als PNG via LibreOffice + pdftoppm; geeft de lijst PNG-paden (leeg als de hulpprogramma's ontbreken)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not shutil.which("pdftoppm"):
        return []
    os.makedirs(map_uit, exist_ok=True)
    profiel = os.path.join(map_uit, "lo_profiel")
    subprocess.run([soffice, "--headless", f"-env:UserInstallation=file://{profiel}", "--convert-to", "pdf", "--outdir", map_uit, pad],
                   capture_output=True, timeout=300)
    pdf = os.path.join(map_uit, os.path.splitext(os.path.basename(pad))[0] + ".pdf")
    if not os.path.exists(pdf):
        return []
    stam = os.path.join(map_uit, "dia")
    subprocess.run(["pdftoppm", "-png", "-r", "110", pdf, stam], capture_output=True, timeout=300)
    return sorted(os.path.join(map_uit, f) for f in os.listdir(map_uit) if f.startswith("dia") and f.endswith(".png"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("invoer")
    ap.add_argument("uitvoer")
    ap.add_argument("extra", nargs="*", help="extra kopieën van de uitvoer")
    ap.add_argument("--voorbeeld", default=VOORBEELD_STANDAARD, help="doorgerekend werkboek (.xlsx) voor de voorbeeldgegevens van dia 6")
    ap.add_argument("--render", default="", help="map voor PNG's van alle dia's (LibreOffice + pdftoppm); leeg = niet renderen")
    args = ap.parse_args()

    prs = Presentation(args.invoer)
    if len(prs.slides) != 7:
        sys.exit(f"Verwacht 7 dia's in v11, gevonden {len(prs.slides)}.")
    if "BT_GRAFIEK" in shapes_by_name(prs.slides[DIA_BT]):
        sys.exit("BT_GRAFIEK bestaat al in de invoer; niets gedaan.")

    slide6, n_vp, hoogte = dia6_verkoop(prs, args.voorbeeld)      # eerst: kopieert de VT_*-tekstvormen van dia 5
    box, oud_part, n_bt = dia5_bouwtermijnen(prs)
    velden = dianummers_bijwerken(prs)

    prs.save(args.uitvoer)
    for extra in args.extra:
        shutil.copyfile(args.uitvoer, extra)
    print(f"Opgeslagen: {args.uitvoer}" + "".join(f"\n            {e}" for e in args.extra))
    print(f"  dia 5: {oud_part} verwijderd; BT_GRAFIEK op x={box[0]} y={box[1]} w={box[2]} h={box[3]} ({n_bt} voorbeeldrijen, as {BT_JAAR_MIN}..{BT_JAAR_MAX})")
    print(f"  dia 6: VP1_GRAFIEK en VP2_GRAFIEK, elk {hoogte} EMU = {hoogte / EMU_CM:.2f} cm hoog ({n_vp} kwartalen uit {os.path.basename(args.voorbeeld)})")
    print(f"  dianummervelden bijgewerkt: {velden}")

    print("Controle:")
    fouten, charts, emb = controleer(args.uitvoer, {
        5: ["BT_KICKER", "BT_TITEL", "BT_SUBTITEL", "VT_KAART", "BT_GRAFIEK", "VT_TABELKAART", "VT_TABEL"],
        6: ["VP_KICKER", "VP_TITEL", "VP_SUBTITEL", "VP_KAART", "VP1_GRAFIEK", "VP2_GRAFIEK"],
        7: ["VO_TITEL", "VO_GRAFIEK"],
    })
    print("  chart-parts: " + ", ".join(charts))
    print("  ingebedde werkboeken: " + ", ".join(emb))
    if fouten:
        sys.exit("FOUTEN:\n  " + "\n  ".join(fouten))
    print("  geen fouten")
    if args.render:
        pngs = render(args.uitvoer, args.render)
        print("  PNG's: " + (", ".join(pngs) if pngs else "niet gerenderd (LibreOffice/pdftoppm ontbreekt)"))


if __name__ == "__main__":
    main()
