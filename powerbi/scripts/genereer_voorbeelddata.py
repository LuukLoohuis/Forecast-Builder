"""Genereert Cashflow_Template.xlsx met voorbeelddata (Q2 2023 t/m Q4 2030).

Gebruik:  python genereer_voorbeelddata.py [uitvoermap]
Schrijft ook voorbeeld_deneb_dataset.json: de rijen zoals Deneb ze van Power BI
ontvangt, zodat de Vega-Lite-spec buiten Power BI te testen is.
"""
import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

REALISATIE_TM = date(2026, 7, 1)  # Q3 2026

# (opbrengsten, kosten) in euro's per kwartaal, kosten als positief bedrag
ACTUEEL = [
    (0, 20_000), (0, 15_000), (0, 60_000), (0, 30_000), (0, 25_000), (0, 30_000),
    (0, 60_000), (0, 80_000), (0, 230_000), (0, 120_000), (480_000, 0),
    (500_000, 170_000), (0, 280_000), (0, 120_000), (0, 230_000), (650_000, 640_000),
    (2_280_000, 640_000), (80_000, 290_000), (0, 550_000), (0, 260_000),
    (3_650_000, 1_150_000), (1_180_000, 750_000), (450_000, 1_090_000), (0, 0),
    (2_430_000, 2_030_000), (2_100_000, 1_030_000), (2_420_000, 2_040_000),
    (960_000, 1_290_000), (950_000, 810_000), (0, 1_660_000), (0, 0),
]


def kwartalen(start=date(2023, 4, 1), n=len(ACTUEEL)):
    y, m = start.year, start.month
    for _ in range(n):
        yield date(y, m, 1)
        m += 3
        if m > 12:
            y, m = y + 1, m - 12


def versies():
    rows = []
    qs = list(kwartalen())
    for i, (q, (o, k)) in enumerate(zip(qs, ACTUEEL)):
        rows.append((q, "Actueel", o, k))
        # Vorige prognose: zelfde patroon, opbrengsten iets eerder gespreid
        vo = o * (0.9 if i % 3 else 1.15)
        rows.append((q, "Vorige prognose", round(vo, -3), k))
        if q > REALISATIE_TM:
            rows.append((q, "Scenario 1", round(o * 0.85, -3), k))            # lagere opbrengsten
            rows.append((q, "Scenario 2", o, round(k * 1.10, -3)))            # hogere kosten
            rows.append((q, "Scenario 3", round(o * 1.05, -3), round(k * 0.98, -3)))
    return rows


def schrijf_excel(pad: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cashflow"
    ws.append(["Kwartaal", "Versie", "Opbrengsten", "Kosten"])
    for r in rows:
        ws.append(list(r))
    for c in ws["A"][1:]:
        c.number_format = "dd-mm-yyyy"
    for col in "CD":
        for c in ws[col][1:]:
            c.number_format = '€ #,##0'
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    t = Table(displayName="tblCashflow", ref=f"A1:D{len(rows) + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)

    inst = wb.create_sheet("Instellingen")
    inst.append(["Instelling", "Waarde"])
    inst.append(["RealisatieTotEnMet", REALISATIE_TM])
    inst["B2"].number_format = "dd-mm-yyyy"
    inst.column_dimensions["A"].width = 22
    inst.column_dimensions["B"].width = 14
    t2 = Table(displayName="tblInstellingen", ref="A1:B2")
    t2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    inst.add_table(t2)

    uitleg = wb.create_sheet("Uitleg")
    regels = [
        ("Cashflow-template voor Power BI", True),
        ("", False),
        ("Tabblad Cashflow (tabel tblCashflow) - één rij per kwartaal per versie:", True),
        ("  Kwartaal     eerste dag van het kwartaal (andere datums worden in Power Query naar kwartaalstart gezet)", False),
        ("  Versie       Actueel | Vorige prognose | Scenario 1 | Scenario 2 | Scenario 3", False),
        ("  Opbrengsten  bedrag in euro's (positief)", False),
        ("  Kosten       bedrag in euro's; positief of negatief invoeren mag, het model maakt het negatief", False),
        ("", False),
        ("Actueel = realisatie t/m de realisatiedatum, daarna de huidige prognose.", False),
        ("Vorige prognose = volledige reeks van de vorige prognoseversie (stippellijn).", False),
        ("Scenario's = alleen kwartalen ná de realisatiedatum; ze starten vanaf de actuele stand.", False),
        ("", False),
        ("Tabblad Instellingen: RealisatieTotEnMet = eerste dag van het laatste gerealiseerde kwartaal.", True),
    ]
    for tekst, vet in regels:
        uitleg.append([tekst])
        if vet:
            uitleg.cell(row=uitleg.max_row, column=1).font = Font(bold=True)
    uitleg.column_dimensions["A"].width = 110
    uitleg["A1"].fill = PatternFill("solid", fgColor="DDEBF7")
    wb.move_sheet("Uitleg", offset=-2)
    wb.active = 1
    wb.save(pad)


def deneb_dataset(rows):
    """Bootst de DAX-measures na: één rij per kwartaal zoals Deneb die ontvangt."""
    qs = sorted({r[0] for r in rows})
    per = {(r[0], r[1]): r[2] - abs(r[3]) for r in rows}
    out, cum = [], {}
    for q in qs:
        rec = {
            "Kwartaal": f"{q.year}-Q{(q.month - 1) // 3 + 1}",
            "Realisatie": 1 if q <= REALISATIE_TM else 0,
        }
        o = next(r[2] for r in rows if r[0] == q and r[1] == "Actueel")
        k = next(r[3] for r in rows if r[0] == q and r[1] == "Actueel")
        rec["Opbrengsten"] = o
        rec["Kosten"] = -abs(k)
        for v in ("Actueel", "Vorige prognose"):
            cum[v] = cum.get(v, 0) + per[(q, v)]
        rec["Cumulatieve cashflow"] = cum["Actueel"]
        rec["Vorige prognose"] = cum["Vorige prognose"]
        for s in ("Scenario 1", "Scenario 2", "Scenario 3"):
            if q < REALISATIE_TM:
                rec[s] = None
            elif q == REALISATIE_TM:
                cum[s] = cum["Actueel"]
                rec[s] = cum[s]
            else:
                cum[s] += per[(q, s)]
                rec[s] = cum[s]
        out.append(rec)
    return out


if __name__ == "__main__":
    doel = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[1]
    rows = versies()
    schrijf_excel(doel / "Cashflow_Template.xlsx", rows)
    (doel / "deneb" / "voorbeeld_deneb_dataset.json").write_text(
        json.dumps(deneb_dataset(rows), indent=1), encoding="utf-8"
    )
    print("Geschreven:", doel / "Cashflow_Template.xlsx")
