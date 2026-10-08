#!/usr/bin/env python3
"""Bouwt Cashflow_scenario (.xlsx en, met een bestaand .xlsm als VBA-bron, ook .xlsm).

Gebruik:
  python build.py                                   # voorbeeldproject -> out/Cashflow_scenario.xlsx
  python build.py --from Cashflow_scenario.xlsm     # neemt invoer en knoppen over uit een bestaand werkboek
                                                    # (oude indeling met vijf types, of de nieuwe indeling)
                                                    # en de VBA uit dat .xlsm -> out/Cashflow_scenario.xlsm
  python build.py --from oud.xlsm --out map/naam    # eigen bestandsnaam (zonder extensie)
  python build.py --vba ander.xlsm                  # VBA uit een ander .xlsm halen

De VBA-module (vba/CashflowNaarPowerPoint_v10.bas) wordt in het .xlsm geschreven: importeren is niet nodig.
Het .xlsm wordt altijd gemaakt; als VBA-sjabloon dient vba/vbaProject_template.bin (of --vba <ander .xlsm>).
"""
import argparse
import os
import tempfile

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from forecast_builder import cache, charts, data as D, layout as LY, package, sheets

HIER = os.path.dirname(os.path.abspath(__file__))
VBA_BAS = os.path.join(HIER, "vba", "CashflowNaarPowerPoint_v10.bas")
VBA_TEMPLATE = os.path.join(HIER, "vba", "vbaProject_template.bin")
VBA_MODULE = "CashflowNaarPowerPoint"


def vba_bin_maken(template_pad=None, bas_pad=VBA_BAS):
    """vbaProject.bin met de module uit het .bas, op basis van het sjabloon (document-modules, projectinstellingen)."""
    from forecast_builder import vbabuild
    pad = template_pad or VBA_TEMPLATE
    if pad.lower().endswith(".xlsm"):
        template = package.lees_vba(pad)
    else:
        with open(pad, "rb") as f:
            template = f.read()
    with open(bas_pad, encoding="utf-8", newline="") as f:
        bron = f.read().replace("\r\n", "\n").replace("\n", "\r\n")   # VBA-bron wordt met CRLF opgeslagen
    codenames = ["shDashboard", "shInvoer", "shWoningtypes", "shModel", "shPowerPoint"]   # werkbladen in het werkboek
    return vbabuild.build_vbaproject(template, {VBA_MODULE: bron}, documentmodules=codenames)


def jaar_min_van(project):
    """Eerste jaar van de periodes (vaste ondergrens van de tijd-as in de bouwtermijnengrafiek)."""
    jaren = []
    for rij in project.periodes:
        try:
            jaren.append(int(float(str(rij.get("jaar")).strip())))
        except (TypeError, ValueError):
            pass
    return min(jaren) if jaren else 2020


def bouw_werkboek(project):
    wb = Workbook()
    wb.remove(wb.active)
    wb.code_name = "ThisWorkbook"
    ws_dash = sheets.bouw_dashboard(wb, project)
    sheets.bouw_invoer(wb, project)
    sheets.bouw_woningtypes(wb, project)
    sheets.bouw_model(wb)
    sheets.bouw_powerpoint(wb, project)
    # benoemde bereiken voor de grafieken: zo lang als er periodes zijn
    for naam, col in LY.M.items():
        if naam.startswith("g_"):
            wb.defined_names[naam] = DefinedName(naam, attr_text=f"OFFSET(Model!${col}${LY.ROW1},0,0,MAX(1,Model!{LY.h('n')}),1)")
    wb.defined_names["g_kwartaal"] = DefinedName("g_kwartaal", attr_text=f"OFFSET(Model!${LY.FIX['kwartaal']}${LY.ROW1},0,0,MAX(1,Model!{LY.h('n')}),1)")
    for k in range(1, LY.N_TYPES + 1):   # verkoopgrafiek: per typeblok verkocht (g_v<k>) en getransporteerd (g_t<k>) per kwartaal
        for naam, blok in ((f"g_v{k}", "gv"), (f"g_t{k}", "gt")):
            wb.defined_names[naam] = DefinedName(naam, attr_text=f"OFFSET(Model!${LY.m_col(blok, k)}${LY.ROW1},0,0,MAX(1,Model!{LY.h('n')}),1)")
    for key in ["label"] + LY.BT_REEKSEN:   # bouwtermijnengrafiek: zo lang als de bouwtermijnentabel rijen heeft
        naam = f"g_bt_{key}"
        wb.defined_names[naam] = DefinedName(naam, attr_text=f"OFFSET(Model!${LY.BT[key]}${LY.BT_ROW1},0,0,MAX(1,Model!{LY.h('bt_n')}),1)")
    for key in LY.PT_REEKSEN:   # grafiek per woningtype (selectieblok): zo lang als er periodes zijn
        naam = f"g_pt_{key}"
        wb.defined_names[naam] = DefinedName(naam, attr_text=f"OFFSET(Model!${LY.PT[key]}${LY.ROW1},0,0,MAX(1,Model!{LY.h('n')}),1)")
    charts.plaats_grafieken(ws_dash, wb, jaar_min_van(project))
    wb.calculation.fullCalcOnLoad = True
    return wb


def bouw(project, pad_uit_basis, vba_bin=None, caches=True):
    """Schrijft <basis>.xlsx en, als vba_bin is meegegeven, <basis>.xlsm. Geeft de geschreven paden terug.

    Met caches=True (en LibreOffice aanwezig) worden de uitkomsten van alle formules en de grafiekreeksen als
    cachewaarde in het bestand gezet, zodat Excel ook in de beveiligde weergave meteen cijfers toont.
    """
    wb = bouw_werkboek(project)
    xmls = charts.chart_xmls(jaar_min_van(project))
    os.makedirs(os.path.dirname(os.path.abspath(pad_uit_basis)), exist_ok=True)
    paden = []
    with tempfile.TemporaryDirectory() as tmp:
        ruw = os.path.join(tmp, "ruw.xlsx")
        wb.save(ruw)
        xlsx = pad_uit_basis + ".xlsx"
        package.nabewerken(ruw, xlsx, xmls)
        paden.append(xlsx)
        if vba_bin is not None:
            xlsm = pad_uit_basis + ".xlsm"
            package.nabewerken(ruw, xlsm, xmls, vba_bin)
            paden.append(xlsm)
        if caches:
            if not cache.caches_vullen(paden, tmp):
                print("let op: LibreOffice niet gevonden of mislukt; bestand zonder cachewaarden (Excel rekent bij openen)")
    return paden


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--from", dest="bron", help="bestaand werkboek waaruit invoer en knoppen worden overgenomen")
    ap.add_argument("--fo", help="FO-werkboek (Financieel Overzicht) waaruit cashflow, woningtypes, termijnen en verkoop worden gelezen (forecast_builder/fo.py)")
    ap.add_argument("--vba", help="VBA-sjabloon: een .xlsm of vbaProject.bin waaruit de projectstructuur en document-modules komen "
                                  "(standaard vba/vbaProject_template.bin); de module zelf komt altijd uit vba/CashflowNaarPowerPoint_v10.bas")
    ap.add_argument("--geen-vba", action="store_true", help="alleen een .xlsx maken")
    ap.add_argument("--geen-cache", action="store_true", help="niet doorrekenen met LibreOffice (geen cachewaarden)")
    ap.add_argument("--out", default="out/Cashflow_scenario", help="uitvoerpad zonder extensie (standaard out/Cashflow_scenario)")
    args = ap.parse_args()

    if args.fo:
        from forecast_builder import fo
        project = fo.lees_fo(args.fo)
        if args.bron:                                   # knoppen (Dashboard) uit een bestaand werkboek overnemen
            uit_fo = ("actuals_jaar", "actuals_kw") if "actuals" in project.fo else ()
            project.params.update({k: v for k, v in D.lees_werkboek(args.bron).params.items() if k not in uit_fo})
        for melding in project.fo.get("waarschuwingen", []):
            print("let op:", melding)
    else:
        project = D.lees_werkboek(args.bron) if args.bron else D.voorbeeld()
    vba_bin = None if args.geen_vba else vba_bin_maken(args.vba)
    for pad in bouw(project, args.out, vba_bin, caches=not args.geen_cache):
        print("geschreven:", pad)


if __name__ == "__main__":
    main()
