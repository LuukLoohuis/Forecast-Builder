"""Voegt de cashflowpagina toe aan een bestaand Power BI-sjabloon (.pbit).

Gebruik:  python voeg_toe_aan_pbit.py <invoer.pbit> <uitvoer.pbit>

Werkt voor sjablonen in het PBIR-formaat (Report/definition/...) met een
pagina-indeling zoals de Heijmans-template: de pagina '1-KPI' wordt als basis
gekopieerd (titelbalk, filterbalk, KPI-blok) en gevuld met:
  - de Deneb-cashflowgrafiek
  - vier KPI-kaarten (dieptepunt, kwartaal, stand realisatie, eindsaldo)
  - de scenario-slicer (eerste slicer in de filterbalk)
Aan het datamodel worden de queries, tabellen, relatie en measures uit
power_query.m en dax/model.dax toegevoegd. Measures komen in 'Calculations'.
"""
import copy
import json
import re
import secrets
import sys
import uuid
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from genereer_pbip import (  # noqa: E402
    DENEB_GUID, MEASURE_TABEL, ROOT, lees_dax, lees_m, lit, model_bim,
)

BRON_PAGINA = "1-KPI"
NIEUWE_MAP = "Cashflow"
PAGINA_TITEL = "Cashflow"
QUERY_GROEP = "Cashflow"
VC_SCHEMA = ("https://developer.microsoft.com/json-schemas/fabric/item/report/"
             "definition/visualContainer/2.12.0/schema.json")


def nieuwe_naam():
    return secrets.token_hex(10)


def tag():
    return str(uuid.uuid4())


# ------------------------------------------------------------- datamodel --
def met_lineage(obj):
    obj.setdefault("lineageTag", tag())
    for c in obj.get("columns", []):
        c.setdefault("lineageTag", tag())
    for ms in obj.get("measures", []):
        ms.setdefault("lineageTag", tag())
    return obj


def breid_model_uit(dms):
    eigen = model_bim(lees_m(), lees_dax())["model"]
    model = dms["model"]
    bestaand = {t["name"]: t for t in model["tables"]}

    for t in eigen["tables"]:
        if t["name"] == MEASURE_TABEL:
            doel = bestaand.get(MEASURE_TABEL)
            if doel is None:
                model["tables"].append(met_lineage(t))
                continue
            namen = {ms["name"] for ms in doel.get("measures", [])}
            for ms in t["measures"]:
                if ms["name"] in namen:
                    raise SystemExit(f"Measure '{ms['name']}' bestaat al in {MEASURE_TABEL}")
                ms["lineageTag"] = tag()
                doel.setdefault("measures", []).append(ms)
            continue
        if t["name"] in bestaand:
            raise SystemExit(f"Tabel '{t['name']}' bestaat al in het sjabloon")
        for p in t["partitions"]:
            if p["source"]["type"] == "m":
                p["queryGroup"] = QUERY_GROEP
        t["annotations"] = [{"name": "PBI_ResultType", "value": "Table"}]
        model["tables"].append(met_lineage(t))

    model.setdefault("relationships", []).extend(
        dict(r, name=tag()) for r in eigen["relationships"])

    for e in eigen["expressions"]:
        e.update(queryGroup=QUERY_GROEP, lineageTag=tag(), annotations=[
            {"name": "PBI_NavigationStepName", "value": "Navigatie"},
            {"name": "PBI_ResultType", "value": "Text"}])
        model.setdefault("expressions", []).append(e)

    groepen = model.setdefault("queryGroups", [])
    if not any(g["folder"] == QUERY_GROEP for g in groepen):
        groepen.append({"folder": QUERY_GROEP, "annotations": [
            {"name": "PBI_QueryGroupOrder", "value": str(len(groepen))}]})

    for a in model.get("annotations", []):
        if a["name"] == "PBI_QueryOrder":
            volgorde = json.loads(a["value"])
            volgorde += [n for n in ("BronPad", "Cashflow", "Instellingen") if n not in volgorde]
            a["value"] = json.dumps(volgorde)
    return dms


# --------------------------------------------------------------- rapport --
def veld(entiteit, naam, measure=False):
    soort = "Measure" if measure else "Column"
    return {"field": {soort: {"Expression": {"SourceRef": {"Entity": entiteit}}, "Property": naam}},
            "queryRef": f"{entiteit}.{naam}", "nativeQueryRef": naam}


def titel(tekst, tonen=True):
    return {"title": [{"properties": {"show": lit(tonen), "text": lit(tekst)}}]}


def deneb_visual(spec):
    velden = [veld("Kalender", "Kwartaal"), veld("Kalender", "Realisatie")]
    velden += [veld(MEASURE_TABEL, n, True) for n in (
        "Opbrengsten", "Kosten", "Cumulatieve cashflow", "Vorige prognose",
        "Scenario 1", "Scenario 2", "Scenario 3")]
    return {
        "visualType": DENEB_GUID,
        "query": {
            "queryState": {"dataset": {"projections": velden}},
            "sortDefinition": {"sort": [{"field": velden[0]["field"], "direction": "Ascending"}],
                               "isDefaultSort": True},
        },
        "objects": {"vega": [{"properties": {
            "jsonSpec": lit(json.dumps(spec, indent=2, ensure_ascii=False)),
            "jsonConfig": lit("{}"),
            "provider": lit("vegaLite"),
            "renderMode": lit("svg"),
            "enableTooltips": lit(True),
        }}]},
        "visualContainerObjects": titel("Cashflow-grafiek", tonen=False),
        "drillFilterOtherVisuals": True,
    }


def kaart(measure, label):
    return {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [veld(MEASURE_TABEL, measure, True)]}}},
        "objects": {"labels": [{"properties": {"fontSize": lit(22)}}],
                    "categoryLabels": [{"properties": {"show": lit(False)}}]},
        "visualContainerObjects": titel(label),
        "drillFilterOtherVisuals": True,
    }


def container(naam, pos, visual, groep):
    return {"$schema": VC_SCHEMA, "name": naam, "position": pos,
            "visual": visual, "parentGroupName": groep}


def zet_tekst(visual_json, tekst):
    runs = visual_json["visual"]["objects"]["general"][0]["properties"]["paragraphs"][0]["textRuns"]
    runs[0]["value"] = tekst
    del runs[1:]


def bouw_pagina(bestanden, spec):
    """Kopieert pagina BRON_PAGINA naar NIEUWE_MAP en vult hem."""
    pre = f"Report/definition/pages/{BRON_PAGINA}/"
    bron = {k[len(pre):]: v for k, v in bestanden.items() if k.startswith(pre)}
    page = json.loads(bron["page.json"])

    # Alle namen (pagina, visuals) krijgen een nieuwe unieke waarde
    namen = {page["name"]: nieuwe_naam()}
    visuals = {}
    for pad, inhoud in bron.items():
        if pad.endswith("/visual.json"):
            v = json.loads(inhoud)
            namen[v["name"]] = nieuwe_naam()
            visuals[pad.split("/")[1]] = v

    # Bookmarks van de filterbalk voor deze pagina kopiëren
    bm_pre = "Report/definition/bookmarks/"
    bookmarks = {}
    for k, v in bestanden.items():
        if k.startswith(bm_pre) and k.endswith(".bookmark.json"):
            b = json.loads(v)
            if b["explorationState"].get("activeSection") == page["name"]:
                namen[b["name"]] = nieuwe_naam()
                bookmarks[k] = b

    def hernoem(obj):
        tekst = json.dumps(obj, ensure_ascii=False)
        for oud, nieuw in namen.items():
            tekst = tekst.replace(oud, nieuw)
        return json.loads(tekst)

    nieuw = {}
    page = hernoem(page)
    page["displayName"] = PAGINA_TITEL
    nieuw[f"Report/definition/pages/{NIEUWE_MAP}/page.json"] = page

    visuals = {k: hernoem(v) for k, v in visuals.items()}
    grafiek = visuals["KPI-1-Grafiek"]
    groep = grafiek["parentGroupName"]

    zet_tekst(visuals["Titelbalk-Titel"], "HNL Dashboard Titel | Cashflow")
    zet_tekst(visuals["KPI-1-Titel"], "Cashflow per kwartaal")

    # Scenario-slicer in de eerste filterbalk-slicer
    sl = visuals["Filterbalk-Slicer-1"]["visual"]
    sl["query"] = {"queryState": {"Values": {"projections": [
        dict(veld("Scenario keuze", "Scenario"), active=True)]}}}
    sl["objects"] = {"data": [{"properties": {"mode": lit("Basic")}}],
                     "selection": [{"properties": {"singleSelect": lit(False)}}]}
    sl["visualContainerObjects"] = titel("Scenario's")

    # KPI-kaarten boven de grafiek, grafiek eronder (posities relatief aan de groep)
    z, tab = grafiek["position"]["z"], grafiek["position"]["tabOrder"]
    kaarten = [("Dieptepunt", "Dieptepunt"), ("Dieptepunt kwartaal", "Dieptepunt in"),
               ("Stand realisatie", "Stand einde realisatie"), ("Eindsaldo", "Eindsaldo")]
    for i, (measure, label) in enumerate(kaarten):
        pos = {"x": 50 + i * 330, "y": 80, "z": z + 1 + i, "height": 110, "width": 310,
               "tabOrder": tab + 1 + i}
        visuals[f"Cashflow-KPI-{i + 1}"] = container(nieuwe_naam(), pos, kaart(measure, label), groep)
    grafiek["position"].update(y=205, height=680)
    grafiek["visual"] = deneb_visual(spec)

    for map_naam, v in visuals.items():
        nieuw[f"Report/definition/pages/{NIEUWE_MAP}/visuals/{map_naam}/visual.json"] = v

    nieuwe_bm = []
    for k, b in bookmarks.items():
        b = hernoem(b)
        b["displayName"] = b["displayName"].replace("1 KPI", PAGINA_TITEL)
        bestand = Path(k).name.replace("1-KPI", NIEUWE_MAP)
        nieuw[bm_pre + bestand] = b
        nieuwe_bm.append(b["name"])
    return nieuw, page["name"], nieuwe_bm


# ------------------------------------------------------------------ main --
def dump(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def main(invoer, uitvoer):
    spec = json.loads((ROOT / "deneb" / "cashflow_vegalite.json").read_text(encoding="utf-8"))
    with zipfile.ZipFile(invoer) as z:
        volgorde = [i.filename for i in z.infolist()]
        bestanden = {n: z.read(n) for n in volgorde}

    if f"Report/definition/pages/{BRON_PAGINA}/page.json" not in bestanden:
        raise SystemExit("Sjabloon heeft geen PBIR-pagina " + BRON_PAGINA)

    # Datamodel (UTF-16LE zonder BOM)
    dms = json.loads(bestanden["DataModelSchema"].decode("utf-16-le"))
    bestanden["DataModelSchema"] = json.dumps(
        breid_model_uit(dms), indent=2, ensure_ascii=False).encode("utf-16-le")

    # Credentials horen bij het oude model; zonder dit bestand vraagt Power BI ze opnieuw
    bestanden.pop("SecurityBindings", None)
    volgorde = [n for n in volgorde if n != "SecurityBindings"]
    ct = bestanden["[Content_Types].xml"].decode("utf-8-sig")
    ct = re.sub(r'<Override PartName="/SecurityBindings"[^>]*/>', "", ct)
    bestanden["[Content_Types].xml"] = b"\xef\xbb\xbf" + ct.encode("utf-8")

    # Rapport
    nieuw, pagina, bms = bouw_pagina({k: v.decode("utf-8-sig") for k, v in bestanden.items()
                                       if k.startswith("Report/definition/")}, spec)
    for k, v in nieuw.items():
        bestanden[k] = dump(v)
        volgorde.append(k)

    pages = json.loads(bestanden["Report/definition/pages/pages.json"])
    pages["pageOrder"].append(pagina)
    bestanden["Report/definition/pages/pages.json"] = dump(pages)

    bm_pad = "Report/definition/bookmarks/bookmarks.json"
    if bm_pad in bestanden and bms:
        bm = json.loads(bestanden[bm_pad])
        bm["items"][0].setdefault("children", []).extend(bms)
        bestanden[bm_pad] = dump(bm)

    rep_pad = "Report/definition/report.json"
    rep = json.loads(bestanden[rep_pad])
    if DENEB_GUID not in rep.setdefault("publicCustomVisuals", []):
        rep["publicCustomVisuals"].append(DENEB_GUID)
    bestanden[rep_pad] = dump(rep)

    with zipfile.ZipFile(uitvoer, "w", zipfile.ZIP_DEFLATED) as z:
        for n in volgorde:
            z.writestr(n, bestanden[n])
    print("Geschreven:", uitvoer)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
