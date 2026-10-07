#!/usr/bin/env python3
"""Kwartaalsjabloon v12 -> v13: bouwtermijnen per woningtype als banen (dia 5).

Gebruik:
    python3 tools/pptx_v13_banen.py INVOER_v12.pptx UITVOER_v13.pptx [UITVOER2.pptx ...] --voorbeeld doorgerekend.xlsx [--render MAP]

Bewerking (python-pptx + lxml, zelfde werkwijze als tools/pptx_v12_bouwtermijnen.py, waarvan de hulpfuncties worden gebruikt):
  * Dia 5: BT_GRAFIEK (v12: één rij per termijn, negen segmenten) wordt verwijderd en op dezelfde plek en maat vervangen
    door de nieuwe BT_GRAFIEK uit forecast_builder.chartxml.bouwtermijnen: koprij jaartallen, koprij kwartaalnummers 1-4,
    daarna per bouwtermijn een baan per woningtype (blokje in de typekleur = huidige planning, lichte tint = vorige
    prognose, grijs = gerealiseerd), een lege rij tussen de termijnen; tijd-as in kwartaalindex (jaar x 4 + kwartaal).
    Legenda onderaan met de typen (huidige planning) en 'Gerealiseerd'.
  * Ingebed werkboek (Sheet1): rij 1 de koppen van het blok AA7.. van tab PowerPoint (A Termijn, daarna de reeksen uit
    layout.BT_REEKSEN: 107 kolommen), daaronder de voorbeeldrijen uit het doorgerekende werkboek (--voorbeeld, tab Model,
    bouwtermijnentabel, Model!BY96 rijen). Reeksnamen van de typen en de jaren verwijzen naar rij 1 (zo volgen ze de
    koppen die de macro plakt); de kwartaalreeksen heten vast '1'..'4'.
  * BT_SUBTITEL krijgt de nieuwe legenda-tekst.

Afspraak met de VBA-macro (VulGrafiek): kop + n rijen naar A1 van het gegevensblad, per reeks X-bereik kolom A en Y-bereik
de kolomletter uit de SERIES-formule; reeksen met kop '(uit)' gaan weg; daarna beide waarde-assen op Model!BY97 (ondergrens)
en Model!BY99 (bovengrens).
"""
import argparse
import os
import shutil
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from forecast_builder import chartxml as cx, layout as LY  # noqa: E402
from pptx_v12_bouwtermijnen import (DIA_BT, controleer, nieuwe_grafiek, render, set_text_keep_format,  # noqa: E402
                                    shapes_by_name, verwijder_grafiek, zet_chart_xml)

SUBTITEL = "typekleur = huidige planning · lichte tint = vorige prognose · grijs = gerealiseerd"
KOLOMMEN = ["label"] + LY.BT_REEKSEN                      # kolom A = label, dan de reeksen (zelfde volgorde als tab PowerPoint)
KOL = {naam: get_column_letter(i + 1) for i, naam in enumerate(KOLOMMEN)}


def refs(n):
    r = {naam: f"Sheet1!${KOL[naam]}$2:${KOL[naam]}${n + 1}" for naam in KOLOMMEN}
    r.update({f"naam_{naam}": f"Sheet1!${KOL[naam]}$1" for naam in LY.BT_PREV + LY.BT_CUR + LY.BT_JR})
    return r


def werkboek_bt(pad_voorbeeld):
    """Sheet1 met de koppen en de voorbeeldrijen uit de bouwtermijnentabel van het doorgerekende werkboek."""
    m = load_workbook(pad_voorbeeld, data_only=True)["Model"]
    n = int(m[f"{LY.M_HULP}{LY.H['bt_n']}"].value or 0)
    t_first, t_end = m[f"{LY.M_HULP}{LY.H['t_first']}"].value, m[f"{LY.M_HULP}{LY.H['t_end']}"].value
    if n == 0:
        sys.exit("Voorbeeldwerkboek zonder bouwtermijnen (Model!BY96 = 0).")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    koppen = ["Termijn"]
    for naam in LY.BT_REEKSEN:
        koppen.append(LY.BT_SEG_NAMEN.get(naam) or m[f"{LY.BT[naam]}7"].value)
    ws.append([str(k) for k in koppen])
    for i in range(n):
        rij = []
        for naam in KOLOMMEN:
            v = m[f"{LY.BT[naam]}{LY.BT_ROW1 + i}"].value
            rij.append(None if (isinstance(v, str) and v.startswith("#")) else v)
        ws.append(rij)
    ws.column_dimensions["A"].width = 28
    return ws, n, int(t_first), int(t_end)


def dia5(prs, pad_voorbeeld):
    """Vervangt BT_GRAFIEK op dia 5 en geeft de banen meer hoogte."""
    slide = prs.slides[DIA_BT]
    by_name = shapes_by_name(slide)
    set_text_keep_format(by_name["BT_SUBTITEL"], SUBTITEL)
    box, oud_part = verwijder_grafiek(slide, "BT_GRAFIEK")
    # meer hoogte voor de banen: grafiekkaart en grafiek hoger, tabelkaart en tabel evenveel omlaag; de tabelkaart blijft
    # 30 pt (381000 EMU) boven de dia-rand, dezelfde grens als SchikDia5 in de VBA aanhoudt
    kaart, tkaart, tabel = by_name["VT_KAART"], by_name["VT_TABELKAART"], by_name["VT_TABEL"]
    onder = prs.slide_height - 381000
    delta = max(0, min(560000, onder - (tkaart.top + tkaart.height)))
    kaart.height += delta
    for vorm in (tkaart, tabel):
        vorm.top += delta
    box = (box[0], box[1], box[2], box[3] + delta)
    ws, n, t_first, t_end = werkboek_bt(pad_voorbeeld)
    gf = nieuwe_grafiek(slide, "BT_GRAFIEK", box, XL_CHART_TYPE.BAR_STACKED, n, len(LY.BT_REEKSEN), "chart_bt13", "Werkblad_bt13")
    by_name["VT_TABELKAART"]._element.addprevious(gf._element)      # zelfde z-volgorde als de oude grafiek
    zet_chart_xml(gf, cx.bouwtermijnen(refs(n), t_first, t_end, legenda=True), ws, n)
    return box, oud_part, n, t_first, t_end


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("invoer")
    ap.add_argument("uitvoer")
    ap.add_argument("extra", nargs="*", help="extra kopieën van de uitvoer")
    ap.add_argument("--voorbeeld", required=True, help="doorgerekend werkboek (.xlsx) met de voorbeeldgegevens (tab Model)")
    ap.add_argument("--render", default="", help="map voor PNG's van alle dia's (LibreOffice + pdftoppm); leeg = niet renderen")
    args = ap.parse_args()

    prs = Presentation(args.invoer)
    if len(prs.slides) != 8:
        sys.exit(f"Verwacht 8 dia's in v12, gevonden {len(prs.slides)}.")
    box, oud_part, n, t_first, t_end = dia5(prs, args.voorbeeld)
    prs.save(args.uitvoer)
    for extra in args.extra:
        shutil.copyfile(args.uitvoer, extra)
    print(f"Opgeslagen: {args.uitvoer}" + "".join(f"\n            {e}" for e in args.extra))
    print(f"  dia 5: {oud_part} verwijderd; BT_GRAFIEK op x={box[0]} y={box[1]} w={box[2]} h={box[3]} "
          f"({n} voorbeeldrijen, {len(LY.BT_REEKSEN)} reeksen, as {t_first}..{t_end})")
    print("Controle:")
    fouten, charts, emb = controleer(args.uitvoer, {
        5: ["BT_KICKER", "BT_TITEL", "BT_SUBTITEL", "VT_KAART", "BT_GRAFIEK", "VT_TABELKAART", "VT_TABEL"],
        6: ["VP_KICKER", "VP_TITEL", "VP_SUBTITEL", "VP_KAART", "VP1_GRAFIEK", "VP2_GRAFIEK"],
        7: ["VO_TITEL", "VO_GRAFIEK"],
    })
    if "ppt/charts/chart_bt.xml" in charts or "ppt/embeddings/Werkblad_bt.xlsx" in emb:
        fouten.append("oude chart_bt / Werkblad_bt niet verwijderd")
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
