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
    # v20: keuzecel en schakelaar van de grafiek per woningtype (totaal zonder DAEB, per type of beide)
    assert const("CEL_PT_KEUZE") == LY.D["pt_keuze"] and const("CEL_PT_PP") == LY.D["pt_pp"]
    assert f'Private Const PT_TOTAAL As String = "{LY.PT_TOTAAL}"' in src
    assert LY.PT_PP_KEUZES[0].startswith("totaal") and LY.PT_PP_KEUZES[1] == "per type" and LY.PT_PP_KEUZES[2].endswith("en per type")
    assert const("CEL_VP_N") == f"{LY.M_HULP}{LY.H['vp_n']}"
    # dia 2 (sjabloon v17): grafiekblok, as-grenzen, positie actuals, instellingen en het tekenvlak van KPI_GRAFIEK
    assert const("BLOK_KC") == LY.PP_BLOK["kc"] + "7" and int(const("KOL_KC")) == LY.PP_KC_KOL
    assert const("CEL_KC_MIN") == f"{LY.M_HULP}{LY.H['kc_min']}" and const("CEL_KC_MAX") == f"{LY.M_HULP}{LY.H['kc_max']}"
    assert const("CEL_POS_ACTUALS") == f"{LY.M_HULP}{LY.H['pos_actuals']}"
    assert const("CEL_SJABLOON") == LY.PP_CEL_SJABLOON and const("CEL_NAAM") == LY.PP_CEL_NAAM
    for naam, waarde in (("KC_PLOT_X", LY.KC_PLOT_X), ("KC_PLOT_W", LY.KC_PLOT_W)):
        m = re.search(rf"^Private Const {naam} As Double = ([0-9.]+)", src, flags=re.M)
        assert m and abs(float(m.group(1)) - waarde) < 1e-4, naam
    assert 'Array("KPI_GRAFIEK", BLOK_KC, KOL_KC, "")' in src
    assert const("CEL_AS_MIN") == f"{LY.M_HULP}{LY.H['t_first']}" and const("CEL_AS_MAX") == f"{LY.M_HULP}{LY.H['t_end']}"
    assert const("KOL_BT") == str(LY.PP_BT_KOL)
    # v18: cashflowblok met de drie scenariolijnen, blok van dia 4 en de hulpcellen van tab Scenario's
    assert const("BLOK_CF") == LY.PP_BLOK["cf"] + "7" and int(const("KOL_CF")) == LY.PP_CF_KOL
    assert const("BLOK_VS") == LY.PP_BLOK["vs"] + "7" and int(const("KOL_VS")) == LY.PP_VS_KOL and int(const("N_VS")) == LY.N_VS
    for naam, hulp in (("CEL_VS_PP", "vs_pp_aan"), ("CEL_VS_MIN", "vs_min"), ("CEL_VS_MAX", "vs_max"), ("CEL_VS_STAP", "vs_stap"),
                       ("CEL_VS_NU", "vs_nu_pos")):
        assert const(naam) == f"{LY.M_HULP}{LY.H[hulp]}", naam
    for naam, waarde in (("VS_PLOT_X", LY.VS_PLOT_X), ("VS_PLOT_W", LY.VS_PLOT_W)):
        m = re.search(rf"^Private Const {naam} As Double = ([0-9.]+)", src, flags=re.M)
        assert m and abs(float(m.group(1)) - waarde) < 1e-4, naam
    assert 'Array("CF_GRAFIEK", BLOK_CF, KOL_CF, "")' in src and 'Array("VS_GRAFIEK", BLOK_VS, KOL_VS, CEL_VP_N)' in src
    for regel in (f'Array("SC_GRAFIEK", "{LY.PP_BLOK["sc"]}7", 10, "")',
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
    # v17: balken na het schikken, de grafiek van dia 2 na het vullen geschikt, '(uit)'-reeksen ook in KPI_GRAFIEK
    assert body.index("SchikDia2 pres, wsP") < body.index("VulBalken pres, wsP") < body.index("t = Tabellen()")
    assert 'If vorm = "KPI_GRAFIEK" Then SchikGrafiekDia2 shp, ch' in procs["VulGrafiekVorm"]
    assert 'Or vorm = "KPI_GRAFIEK" Or vorm = "VS_GRAFIEK" Then VerwijderUitgezetteReeksen' in procs["VulGrafiekVorm"]
    assert "If IsBalkregel(vorm) Then" in procs["VulTeksten"] and "RichtingKleur(richting)" in procs["VulBalken"]
    assert 'RegelVanVorm(wsP, "KPI" & k & "_NAAM1")' in procs["SchikDia2"] and "GRAFIEK_LEGENDA_VORIG_LIJN" in procs["SchikDia2"]


def test_dia_scenarios_alleen_als_ze_aan_staan():
    """Dia 4 'Verkoopscenario's' (sjabloon v18): zonder ingevuld scenario of met Naar PowerPoint = nee (Model!vs_pp_aan = 0) haalt
    de macro de dia weg en slaat hij de VS_-teksten en VS_GRAFIEK over; anders verdwijnen de lege kaarten en legenda-regels en
    schikt SchikGrafiekVS de as en het grijze vlak. De scenariolijnen in de cashflowgrafiek gaan weg via de kop '(uit)'."""
    src = open(BAS, encoding="utf-8").read()
    procs = _procedures(src)
    body = procs["NaarPowerPoint"]
    assert body.index("VulTeksten pres, wsP, mist") < body.index("SchikDiaScenarios pres, wsP") < body.index("t = Tabellen()")
    assert 'If CStr(g(i)(0)) <> "VS_GRAFIEK" Or VsDiaAan() Then' in body
    assert 'Left$(vorm, 3) = "VS_" And Not VsDiaAan()' in procs["VulTeksten"]
    assert "VsDiaAan()" in procs["SchikDiaScenarios"] and ".Delete" in procs["SchikDiaScenarios"]
    assert '"VS_K" & k & "_WAARDE"' in procs["SchikDiaScenarios"] and '"VS_LEG" & k' in procs["SchikDiaScenarios"]
    assert 'If vorm = "VS_GRAFIEK" Then SchikGrafiekVS shp, ch' in procs["VulGrafiekVorm"]
    assert '"VS_REALISATIE"' in procs["SchikGrafiekVS"] and "CEL_VS_NU" in procs["SchikGrafiekVS"]
    assert "CEL_VS_PP" in procs["VsDiaAan"]


def test_vervallen_vormen_weg():
    """Sinds v21 krijgen een paar vormen geen tekst meer (ondertitels, regels onderaan de scenariokaarten, voetnoot dia 4): de macro
    haalt precies die vormen vóór het vullen van de dia (ook uit sjabloon v18), en tab PowerPoint heeft er geen regel meer voor."""
    from forecast_builder import sheets
    import inspect
    src = open(BAS, encoding="utf-8").read()
    procs = _procedures(src)
    body = procs["NaarPowerPoint"]
    assert body.index("HaalVervallenVormenWeg pres") < body.index("VulTeksten pres, wsP, mist")
    weg = re.findall(r'"([A-Z0-9_]+)"', procs["HaalVervallenVormenWeg"])
    sys.path.insert(0, os.path.join(os.path.dirname(BAS), "..", "tools"))
    import pptx_v19_kort
    assert sorted(weg) == sorted(pptx_v19_kort.WEG)
    tekst = inspect.getsource(sheets)
    assert not any(f'"{v}"' in tekst for v in weg)
