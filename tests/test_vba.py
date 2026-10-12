"""Statische controles op de VBA-module (geen Excel nodig)."""
import os
import re
import sys

HIER = os.path.dirname(os.path.abspath(__file__))
BAS = os.path.join(os.path.dirname(HIER), "vba", "CashflowNaarPowerPoint_v10.bas")


def _procedures(src):
    """{naam: body} van alle Subs en Functions."""
    uit = {}
    for m in re.finditer(r"^(?:Private |Public )?(?:Sub|Function) (\w+)\(.*?\n(.*?)^End (?:Sub|Function)", src, flags=re.M | re.S):
        uit[m.group(1)] = m.group(2)
    return uit


def test_geen_functie_roept_zichzelf_aan():
    """Een functie die haar eigen naam aanroept (Tekst = Tekst(v)) loopt in Excel vast op 'Out of stack space'."""
    src = open(BAS, encoding="utf-8").read()
    procs = _procedures(src)
    assert len(procs) > 40
    recursief = [n for n, body in procs.items() if re.search(rf"^\s*(?!')[^'\n]*\b{n}\s*\(", body, flags=re.M | re.I)]
    assert recursief == [], recursief


def test_foutafhandelaars_bewaren_err():
    """Een On Error-statement wist het Err-object: in een handler moet Err.Number vóór 'On Error Resume Next' zijn vastgelegd."""
    src = open(BAS, encoding="utf-8").read()
    for m in re.finditer(r"^Fout:\n(.*?)^End Sub", src, flags=re.M | re.S):
        body = m.group(1)
        i = body.find("On Error Resume Next")
        if i >= 0:
            assert "Err.Number" in body[:i], "Err.Number pas na On Error Resume Next gelezen:\n" + body
            assert "Err.Number" not in body[i:], "Err.Number na On Error Resume Next gelezen (is dan altijd 0):\n" + body


def test_geen_procedure_heet_als_parameter():
    """Een procedure met dezelfde naam als een parameter of variabele elders is verwarrend en gaf in LibreOffice een crash."""
    src = open(BAS, encoding="utf-8").read()
    procs = set(n.lower() for n in _procedures(src))
    namen = set()
    for m in re.finditer(r"\b(?:Dim|Const)\s+(.+)$", src, flags=re.M):
        namen |= {w.group(1).lower() for part in m.group(1).split(",") if (w := re.match(r"\s*(\w+)", part))}
    for m in re.finditer(r"^(?:Private |Public )?(?:Sub|Function) \w+\((.*?)\)", src, flags=re.M):
        namen |= {w.group(1).lower() for part in m.group(1).split(",") if (w := re.match(r"\s*(?:ByVal |ByRef |Optional )*(\w+)", part)) and w.group(1)}
    assert procs & namen == set(), procs & namen


def test_blokconstanten_volgen_de_layout():
    """De celverwijzingen in de macro (blokken op tab PowerPoint, hulpcellen op Model) horen bij forecast_builder/layout.py."""
    sys.path.insert(0, os.path.dirname(HIER))
    from forecast_builder import layout as LY
    src = open(BAS, encoding="utf-8").read()

    def const(naam):
        m = re.search(rf'^Private Const {naam} As (?:String|Long) = "?([A-Z0-9]+)"?', src, flags=re.M)
        assert m, f"constante {naam} niet gevonden"
        return m.group(1)

    assert const("CEL_N") == f"{LY.M_HULP}{LY.H['n']}"
    assert const("BLOK_PT") == LY.PP_BLOK["pt"] + "7"
    assert int(const("KOL_PT")) == LY.PP_PT_KOL
    assert int(const("KOL_VP")) == LY.PP_VP_KOL
    assert const("CEL_PT_N") == f"{LY.M_HULP}{LY.H['pt_n']}"
    assert const("CEL_VP_N") == f"{LY.M_HULP}{LY.H['vp_n']}"
    assert const("CEL_AS_MIN") == f"{LY.M_HULP}{LY.H['t_first']}" and const("CEL_AS_MAX") == f"{LY.M_HULP}{LY.H['t_end']}"
    assert const("KOL_BT") == str(LY.PP_BT_KOL)
    for regel in (f'Array("CF_GRAFIEK", "{LY.PP_BLOK["cf"]}7", 8, "")', f'Array("SC_GRAFIEK", "{LY.PP_BLOK["sc"]}7", 10, "")',
                  f'Array("BT_GRAFIEK", "{LY.PP_BLOK["bt"]}7", KOL_BT, "{LY.M_HULP}{LY.H["bt_n"]}")',
                  f'Array("VP1_GRAFIEK", "{LY.PP_BLOK["vp"]}7", KOL_VP, CEL_VP_N)', f'Array("VP2_GRAFIEK", "{LY.PP_BLOK["vp"]}7", KOL_VP, CEL_VP_N)',
                  f'Array("VO_GRAFIEK", "{LY.PP_BLOK["vo"]}7", 5, "")', f'Array("SC_TABEL", "{LY.PP_TABEL_SC}", 9, 4, 0)',
                  f'Array("VT_TABEL", "{LY.PP_TABEL_SC.rstrip("7")}{LY.PP_TABEL_VT_ROW}", MAX_TYPES + 1, 8, 1)'):
        assert regel in src, regel


def test_dia2_geschikt_na_de_teksten():
    """NaarPowerPoint schikt dia 2 direct na het vullen van de teksten: kleine kaarten zonder regel op tab PowerPoint gaan weg
    (SchikDia2, op basis van kolom F), de rest verdeelt de breedte; de voetnoot gaat van de dia."""
    src = open(BAS, encoding="utf-8").read()
    procs = _procedures(src)
    body = procs["NaarPowerPoint"]
    assert body.index("VulTeksten pres, wsP, mist") < body.index("SchikDia2 pres, wsP") < body.index("t = Tabellen()")
    assert 'RegelVanVorm(wsP, "KPI" & k & "_WAARDE")' in procs["SchikDia2"] and "shp.Delete" in procs["SchikDia2"]
    assert 'VormOpDia(dia, "VOETNOOT")' in procs["SchikDia2"] and '"Stip"' in procs["SchikDia2"]
