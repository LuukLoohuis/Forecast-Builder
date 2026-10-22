"""Bouwt het Power BI-project (PBIP) in powerbi/pbip/ uit de bronbestanden.

Bronnen (enige bron van waarheid):
  power_query.m                 -> queries Cashflow en Instellingen
  dax/model.dax                 -> berekende tabellen en measures
  deneb/cashflow_vegalite.json  -> Deneb-visual

Gebruik:  python genereer_pbip.py
Open daarna pbip/Cashflow.pbip in Power BI Desktop en sla op als .pbix.
"""
import json
import re
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "pbip"
NAAM = "Cashflow"
DENEB_GUID = "deneb7E15AEF80B9E4D4F8E12924291ECE89A"
STANDAARD_PAD = r"C:\Data\Cashflow_Template.xlsx"
MEASURE_TABEL = "Calculations"
LEGE_TABEL_M = ["let", "    Bron = #table({}, {})", "in", "    Bron"]
EURO = "\u20ac #,##0;-\u20ac #,##0;\u20ac 0"


# ---------------------------------------------------------------- bronnen --
def lees_m():
    tekst = (ROOT / "power_query.m").read_text(encoding="utf-8")
    queries = {}
    for naam in ("Cashflow", "Instellingen"):
        blok = tekst.split(f"// ---------- Query: {naam} ----------", 1)[1]
        blok = blok.split("// ----------", 1)[0].strip()
        # BronPad komt uit de parameter van het model
        blok = re.sub(r"^\s*BronPad\s*=.*\n", "", blok, flags=re.M)
        queries[naam] = blok.splitlines()
    return queries


def lees_dax():
    """Splitst model.dax in blokken 'Naam = expressie'."""
    blokken, naam, regels = {}, None, []
    for regel in (ROOT / "dax" / "model.dax").read_text(encoding="utf-8").splitlines():
        m = re.match(r"^(?!VAR )([^\s/].*?) =\s*$", regel)
        if m:
            if naam:
                blokken[naam] = regels
            naam, regels = m.group(1), []
        elif naam is not None:
            regels.append(regel)
    blokken[naam] = regels

    def opschonen(r):
        r = [x for x in r if not x.lstrip().startswith("//")]
        while r and not r[-1].strip():
            r.pop()
        return [re.sub(r"\s+//.*$", "", x) for x in r]

    return {k: opschonen(v) for k, v in blokken.items()}


# ------------------------------------------------------------ semantisch --
def kolom(naam, type_, fmt=None, bron=None, verborgen=False, **extra):
    k = {"name": naam, "dataType": type_, "sourceColumn": bron or naam,
         "summarizeBy": "none"}
    if fmt:
        k["formatString"] = fmt
    if verborgen:
        k["isHidden"] = True
    k.update(extra)
    return k


def calc_kolom(naam, type_, **extra):
    return kolom(naam, type_, bron=f"[{naam}]", type="calculatedTableColumn",
                 isNameInferred=True, isDataTypeInferred=True, **extra)


def model_bim(m, dax):
    measures_fmt = {
        "Opbrengsten": EURO, "Kosten": EURO, "Cumulatieve cashflow": EURO,
        "Vorige prognose": EURO, "Scenario 1": EURO, "Scenario 2": EURO,
        "Scenario 3": EURO, "Dieptepunt": EURO, "Stand realisatie": EURO,
        "Eindsaldo": EURO, "_Netto": EURO, "_Realisatie t/m": "dd-mm-yyyy",
    }
    tabellen_dax = {"Kalender", "Scenario keuze"}
    measures = []
    for naam, regels in dax.items():
        if naam in tabellen_dax:
            continue
        ms = {"name": naam, "expression": regels}
        if naam in measures_fmt:
            ms["formatString"] = measures_fmt[naam]
        if naam.startswith("_"):
            ms["isHidden"] = True
            ms["displayFolder"] = "Hulp"
        measures.append(ms)

    tables = [
        {
            "name": MEASURE_TABEL,
            "partitions": [{"name": MEASURE_TABEL, "mode": "import",
                            "source": {"type": "m", "expression": LEGE_TABEL_M}}],
            "measures": measures,
        },
        {
            "name": "Cashflow",
            "columns": [
                kolom("Kwartaal", "dateTime", "dd-mm-yyyy", verborgen=True),
                kolom("Versie", "string"),
                kolom("Opbrengsten", "double", EURO, verborgen=True, summarizeBy="sum"),
                kolom("Kosten", "double", EURO, verborgen=True, summarizeBy="sum"),
            ],
            "partitions": [{"name": "Cashflow", "mode": "import",
                            "source": {"type": "m", "expression": m["Cashflow"]}}],
        },
        {
            "name": "Instellingen",
            "columns": [kolom("RealisatieTotEnMet", "dateTime", "dd-mm-yyyy")],
            "partitions": [{"name": "Instellingen", "mode": "import",
                            "source": {"type": "m", "expression": m["Instellingen"]}}],
        },
        {
            "name": "Kalender",
            "dataCategory": "Time",
            "columns": [
                calc_kolom("Datum", "dateTime", formatString="dd-mm-yyyy", isKey=True),
                calc_kolom("Jaar", "int64", formatString="0"),
                calc_kolom("Kwartaal", "string"),
                calc_kolom("Kwartaal label", "string", sortByColumn="Datum"),
                calc_kolom("Realisatie", "int64", formatString="0"),
            ],
            "partitions": [{"name": "Kalender", "mode": "import",
                            "source": {"type": "calculated", "expression": dax["Kalender"]}}],
        },
        {
            "name": "Scenario keuze",
            "columns": [calc_kolom("Scenario", "string")],
            "partitions": [{"name": "Scenario keuze", "mode": "import",
                            "source": {"type": "calculated", "expression": dax["Scenario keuze"]}}],
        },
    ]
    return {
        "compatibilityLevel": 1550,
        "model": {
            "culture": "nl-NL",
            "sourceQueryCulture": "nl-NL",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "tables": tables,
            "relationships": [{
                "name": "Kalender_Cashflow",
                "fromTable": "Cashflow", "fromColumn": "Kwartaal",
                "toTable": "Kalender", "toColumn": "Datum",
            }],
            "expressions": [{
                "name": "BronPad",
                "kind": "m",
                "expression": f'"{STANDAARD_PAD}" meta [IsParameterQuery=true, '
                              'Type="Text", IsParameterQueryRequired=true]',
            }],
            "annotations": [
                {"name": "PBI_QueryOrder", "value": '["BronPad","Calculations","Cashflow","Instellingen"]'},
                {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
            ],
        },
    }


# ---------------------------------------------------------------- rapport --
def lit(v):
    """Power BI-literal: tekst tussen enkele quotes, quotes verdubbeld."""
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": "'" + str(v).replace("'", "''") + "'"}}}


def veld(entiteit, alias, naam, measure=False):
    soort = "Measure" if measure else "Column"
    return {
        soort: {"Expression": {"SourceRef": {"Source": alias}}, "Property": naam},
        "Name": f"{entiteit}.{naam}",
        "NativeReferenceName": naam,
    }


def container(x, y, w, h, z, single):
    cfg = {"name": uuid.uuid4().hex[:20],
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h}}],
           "singleVisual": single}
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(cfg, ensure_ascii=False), "filters": "[]"}


def titel(tekst):
    return {"title": [{"properties": {"show": lit(True), "text": lit(tekst)}}]}


def deneb_visual(spec):
    kolommen = [("Kalender", "k", "Kwartaal", False), ("Kalender", "k", "Realisatie", False)]
    kolommen += [(MEASURE_TABEL, "c", n, True) for n in (
        "Opbrengsten", "Kosten", "Cumulatieve cashflow", "Vorige prognose",
        "Scenario 1", "Scenario 2", "Scenario 3")]
    select = [veld(e, a, n, m) for e, a, n, m in kolommen]
    return {
        "visualType": DENEB_GUID,
        "projections": {"dataset": [{"queryRef": s["Name"]} for s in select]},
        "prototypeQuery": {
            "Version": 2,
            "From": [{"Name": "k", "Entity": "Kalender", "Type": 0},
                     {"Name": "c", "Entity": MEASURE_TABEL, "Type": 0}],
            "Select": select,
            "OrderBy": [{"Direction": 1, "Expression": {"Column": {
                "Expression": {"SourceRef": {"Source": "k"}}, "Property": "Kwartaal"}}}],
        },
        "drillFilterOtherVisuals": True,
        "objects": {"vega": [{"properties": {
            "jsonSpec": lit(json.dumps(spec, indent=2, ensure_ascii=False)),
            "jsonConfig": lit("{}"),
            "provider": lit("vegaLite"),
            "renderMode": lit("svg"),
            "enableTooltips": lit(True),
        }}]},
        "vcObjects": titel("Cashflow per kwartaal \u00b7 balken: opbrengsten en kosten \u00b7 "
                           "lijnen: cumulatieve cashflow, vorige prognose (stippel), scenario's"),
    }


def kaart(measure, label):
    s = veld(MEASURE_TABEL, "c", measure, True)
    return {
        "visualType": "card",
        "projections": {"Values": [{"queryRef": s["Name"]}]},
        "prototypeQuery": {"Version": 2,
                           "From": [{"Name": "c", "Entity": MEASURE_TABEL, "Type": 0}],
                           "Select": [s]},
        "objects": {"labels": [{"properties": {"fontSize": lit(20)}}],
                    "categoryLabels": [{"properties": {"show": lit(False)}}]},
        "vcObjects": titel(label),
    }


def slicer():
    s = veld("Scenario keuze", "s", "Scenario")
    return {
        "visualType": "slicer",
        "projections": {"Values": [{"queryRef": s["Name"], "active": True}]},
        "prototypeQuery": {"Version": 2,
                           "From": [{"Name": "s", "Entity": "Scenario keuze", "Type": 0}],
                           "Select": [s]},
        "objects": {"data": [{"properties": {"mode": lit("Basic")}}],
                    "selection": [{"properties": {"singleSelect": lit(False)}}]},
        "vcObjects": titel("Scenario's tonen"),
    }


def report_json(spec):
    vis = [
        container(16, 16, 230, 90, 1000, kaart("Dieptepunt", "Dieptepunt")),
        container(256, 16, 160, 90, 1001, kaart("Dieptepunt kwartaal", "Dieptepunt in")),
        container(426, 16, 230, 90, 1002, kaart("Stand realisatie", "Stand einde realisatie")),
        container(666, 16, 230, 90, 1003, kaart("Eindsaldo", "Eindsaldo")),
        container(1060, 16, 204, 150, 1004, slicer()),
        container(16, 176, 1248, 528, 2000, deneb_visual(spec)),
    ]
    config = {
        "version": "5.43",
        "activeSectionIndex": 0,
        "defaultDrillFilterOtherVisuals": True,
        "publicCustomVisuals": [DENEB_GUID],
        "settings": {"useStylableVisualContainerHeader": True},
    }
    return {
        "config": json.dumps(config),
        "layoutOptimization": 0,
        "sections": [{
            "name": "ReportSectionCashflow",
            "displayName": "Cashflow",
            "displayOption": 1,
            "width": 1280,
            "height": 720,
            "config": "{}",
            "filters": "[]",
            "visualContainers": vis,
        }],
        "filters": "[]",
    }


# ------------------------------------------------------------------ main --
def schrijf(pad: Path, data):
    pad.parent.mkdir(parents=True, exist_ok=True)
    tekst = data if isinstance(data, str) else json.dumps(data, indent=2, ensure_ascii=False)
    pad.write_text(tekst + "\n", encoding="utf-8")


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    spec = json.loads((ROOT / "deneb" / "cashflow_vegalite.json").read_text(encoding="utf-8"))
    sm, rp = OUT / f"{NAAM}.SemanticModel", OUT / f"{NAAM}.Report"

    schrijf(OUT / f"{NAAM}.pbip", {
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAAM}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    schrijf(sm / "definition.pbism", {"version": "1.0", "settings": {}})
    schrijf(sm / "model.bim", model_bim(lees_m(), lees_dax()))
    schrijf(rp / "definition.pbir", {
        "version": "1.0",
        "datasetReference": {"byPath": {"path": f"../{NAAM}.SemanticModel"}, "byConnection": None},
    })
    schrijf(rp / "report.json", report_json(spec))
    schrijf(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print("Geschreven:", OUT)


if __name__ == "__main__":
    main()
