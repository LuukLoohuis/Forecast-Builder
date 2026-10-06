#!/usr/bin/env python3
"""Bouwt Cashflow_scenario (.xlsx en, met een bestaand .xlsm als VBA-bron, ook .xlsm).

Gebruik:
  python build.py                                   # voorbeeldproject -> out/Cashflow_scenario.xlsx
  python build.py --from Cashflow_scenario.xlsm     # neemt invoer en knoppen over uit een bestaand werkboek
                                                    # (oude indeling met vijf types, of de nieuwe indeling)
                                                    # en de VBA uit dat .xlsm -> out/Cashflow_scenario.xlsm
  python build.py --from oud.xlsm --out map/naam    # eigen bestandsnaam (zonder extensie)
  python build.py --vba ander.xlsm                  # VBA uit een ander .xlsm halen

Daarna in Excel: open het .xlsm, importeer vba/CashflowNaarPowerPoint_v10.bas over de oude module
(VBA-editor: module CashflowNaarPowerPoint verwijderen, Bestand > Bestand importeren) en bewaar.
"""
import argparse
import os
import tempfile

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from forecast_builder import charts, data as D, layout as LY, package, sheets


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
    wb.defined_names["g_kwartaal"] = DefinedName("g_kwartaal", attr_text=f"OFFSET(Model!$C${LY.ROW1},0,0,MAX(1,Model!{LY.h('n')}),1)")
    charts.plaats_grafieken(ws_dash, wb)
    wb.calculation.fullCalcOnLoad = True
    return wb


def bouw(project, pad_uit_basis, vba_bin=None):
    """Schrijft <basis>.xlsx en, als vba_bin is meegegeven, <basis>.xlsm. Geeft de geschreven paden terug."""
    wb = bouw_werkboek(project)
    os.makedirs(os.path.dirname(os.path.abspath(pad_uit_basis)), exist_ok=True)
    paden = []
    with tempfile.TemporaryDirectory() as tmp:
        ruw = os.path.join(tmp, "ruw.xlsx")
        wb.save(ruw)
        xlsx = pad_uit_basis + ".xlsx"
        package.nabewerken(ruw, xlsx, charts.chart_xmls())
        paden.append(xlsx)
        if vba_bin is not None:
            xlsm = pad_uit_basis + ".xlsm"
            package.nabewerken(ruw, xlsm, charts.chart_xmls(), vba_bin)
            paden.append(xlsm)
    return paden


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--from", dest="bron", help="bestaand werkboek waaruit invoer en knoppen worden overgenomen")
    ap.add_argument("--vba", help=".xlsm waaruit de VBA (vbaProject.bin) wordt overgenomen; standaard het --from-bestand als dat een .xlsm is")
    ap.add_argument("--out", default="out/Cashflow_scenario", help="uitvoerpad zonder extensie (standaard out/Cashflow_scenario)")
    args = ap.parse_args()

    project = D.lees_werkboek(args.bron) if args.bron else D.voorbeeld()
    vba_pad = args.vba or (args.bron if args.bron and args.bron.lower().endswith(".xlsm") else None)
    vba_bin = package.lees_vba(vba_pad) if vba_pad else None
    for pad in bouw(project, args.out, vba_bin):
        print("geschreven:", pad)


if __name__ == "__main__":
    main()
