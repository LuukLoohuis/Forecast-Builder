"""Regressietests voor het gebouwde werkboek.

Draaien: python3 -m pytest tests/ -q   (of: python3 tests/test_model.py)
Heeft LibreOffice (soffice) nodig om formules door te rekenen; zonder LibreOffice worden de rekentests overgeslagen.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

import pytest
from openpyxl import load_workbook

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))

import build  # noqa: E402
from forecast_builder import data as D, layout as LY  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def herbereken(pad):
    """Rekent het werkboek door met LibreOffice (in-place) en geeft het pad terug."""
    map_ = os.path.dirname(pad)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir",
                    os.path.join(map_, "calc"), pad], check=True, capture_output=True, timeout=300)
    return os.path.join(map_, "calc", os.path.basename(pad))


def bouw_en_bereken(project, map_, naam):
    paden = build.bouw(project, os.path.join(map_, naam), None)
    return load_workbook(herbereken(paden[0]), data_only=True)


@pytest.fixture(scope="module")
def tmp():
    with tempfile.TemporaryDirectory() as d:
        yield d


@pytest.fixture(scope="module")
def basis(tmp):
    if not SOFFICE:
        pytest.skip("LibreOffice niet gevonden")
    return bouw_en_bereken(D.voorbeeld(), tmp, "basis")


def foutcellen(wb):
    """Foutwaarden buiten de plekken waar NA() hoort (grafiekkolommen voorbij de laatste periode, vorige prognose)."""
    n = wb["Model"]["BY8"].value
    chart = {LY.M[k] for k in LY.M if k.startswith("g_")}
    na_ok = {LY.FIX["vorig"], LY.M["g_vorige"], LY.M["g_punt_nu"], LY.M["g_punt_dal"], LY.M["g_eind_upside"], LY.M["g_eind_downside"], LY.M["g_eind_basis"],
             LY.M["g_realisatie"], LY.M["g_scenario"], LY.M["g_eind_scenario"], LY.M["g_punt_uitverkocht"], LY.M["g_punt_alles_transport"]}
    na_ok |= {LY.m_col(blok, k) for blok in ("gv", "gt") for k in range(1, LY.N_TYPES + 1)}   # lege typeblokken: NA()
    na_ok |= set(LY.BT.values())                                                            # bouwtermijnentabel: NA() buiten de gebruikte rijen
    na_ok |= set(LY.PT.values())                                                            # grafiek per type: NA() voorbij de laatste periode
    fouten = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("#"):
                    col = "".join(ch for ch in c.coordinate if ch.isalpha())
                    # PowerPoint: NA() buiten de periodes/rijen en in de puntkolommen, de scenariolijn, lege typeblokken en de realisatie
                    ok = v == "#N/A" and ((ws.title == "Model" and (col in na_ok or (col in chart and c.row > 7 + n)))
                                          or ws.title == "PowerPoint")
                    if not ok:
                        fouten.append(f"{ws.title}!{c.coordinate}={v}")
    return fouten


def test_geen_onverwachte_fouten(basis):
    assert foutcellen(basis) == []
    assert basis["Model"]["BY8"].value == len(D.voorbeeld().periodes)


def test_controles_ok(basis):
    d = basis["Dashboard"]
    assert [d[f"W{r}"].value for r in range(LY.D_ROW_CONTROLES + 1, LY.D_ROW_CONTROLES + 7)] == ["OK"] * 6


def test_rente_nagerekend(basis):
    m, d = basis["Model"], basis["Dashboard"]
    n = m["BY8"].value
    rente = d["W6"].value
    start_idx = m[f"BY{LY.H['start_idx']}"].value
    eind = {"basis": start_idx, "up": start_idx + d["X18"].value, "down": start_idx + d["W18"].value}
    kol = {"basis": ("I", "H", LY.M["stand_basis_rente"]), "up": (LY.M["opbr_up"], LY.M["kosten_up"], LY.M["stand_up_raw"]),
           "down": (LY.M["opbr_down"], LY.M["kosten_down"], LY.M["stand_down_raw"])}
    for naam, (opb, kos, stand_col) in kol.items():
        stand = 0.0
        for r in range(8, 8 + n):
            ri = max(0.0, -stand) * rente / 4 if (m[f"G{r}"].value == "Prognose" and m[f"F{r}"].value < eind[naam]) else 0.0
            stand += (m[f"{opb}{r}"].value or 0) - (m[f"{kos}{r}"].value or 0) - ri
            assert abs(m[f"{stand_col}{r}"].value - stand) < 1e-6, (naam, r)


def test_knoppen_op_nul_geeft_basis(tmp):
    if not SOFFICE:
        pytest.skip("LibreOffice niet gevonden")
    p = D.voorbeeld()
    for k in ("shift_down", "shift_up", "uitstel_down", "uitstel_up", "opbr_down", "opbr_up", "kosten_down", "kosten_up"):
        p.params[k] = 0
    wb = bouw_en_bereken(p, tmp, "knoppen0")
    m = wb["Model"]
    n = m["BY8"].value
    for r in range(8, 8 + n):
        b = m[f"{LY.M['stand_basis']}{r}"].value
        assert abs(m[f"{LY.M['stand_up']}{r}"].value - b) < 1e-6
        assert abs(m[f"{LY.M['stand_down']}{r}"].value - b) < 1e-6


def test_lege_rijen_tellen_niet_mee(tmp, basis):
    if not SOFFICE:
        pytest.skip("LibreOffice niet gevonden")
    p = D.voorbeeld()
    leeg = {k: None for k in p.periodes[0]}
    p.periodes = [dict(leeg)] * 5 + p.periodes[:3] + [dict(leeg)] + p.periodes[3:]
    p.verkocht = [[None] * 10] * 5 + p.verkocht[:3] + [[None] * 10] + p.verkocht[3:]
    p.transport = [[None] * 10] * 5 + p.transport[:3] + [[None] * 10] + p.transport[3:]
    wb = bouw_en_bereken(p, tmp, "legerijen")
    for r in range(8, 70):
        a, b = basis["Model"][f"BY{r}"].value, wb["Model"][f"BY{r}"].value
        assert a == b or (isinstance(a, float) and abs(a - b) < 1e-6), f"BY{r}: {a!r} != {b!r}"
    assert [wb["Model"][f"B{r}"].value for r in range(8, 12)] == [7, 8, 9, 11]   # positie binnen Invoer!$B$7:$B$500 (koprij = 1)


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))


def test_bouwtermijnen_als_pct_van_aanneemsom(tmp, basis):
    """Bouwtermijnen die samen 100% zijn (van de aanneemsom) geven hetzelfde model als termijnen die samen 100% − grond zijn."""
    p = D.voorbeeld()
    for t in p.types:
        som = sum(float(pct or 0) for _, pct, _ in t.termijnen)
        if som:
            t.termijnen = [(naam, (float(pct or 0) / som) if pct is not None else None, kw) for naam, pct, kw in t.termijnen]
    wb = bouw_en_bereken(p, tmp, "aanneemsom")
    m, mb = wb["Model"], basis["Model"]
    n = basis["Model"]["BY8"].value
    for naam in ("model_basis", "model_up", "stand_basis", "stand_up", "stand_down"):
        for r in range(LY.ROW1, LY.ROW1 + n):
            a, b = m[f"{LY.M[naam]}{r}"].value, mb[f"{LY.M[naam]}{r}"].value
            assert abs(float(a) - float(b)) < 1e-3, (naam, r, a, b)
    assert wb["Dashboard"][f"W{LY.D_ROW_CONTROLES + 9}"].value == "OK"
    assert all(abs(float(wb["Woningtypes"].cell(LY.WT_R_TOTAAL, LY.wt_col(k, 1)).value) - 1) < 1e-9
               for k in range(1, len(p.types) + 1) if p.types[k - 1].termijnen)


def test_extra_opbrengsten_per_woning(tmp, basis):
    """Extra's (€ per woning): bij transport en per bouwkwartaal tellen ze op bij de modelopbrengst, de koopsomknop niet."""
    p = D.voorbeeld()
    t = p.types[0]
    t.extras = [("Kadastrale kosten", 1000, None), ("Kopersmeerwerk", 2000, 3)]
    wb = bouw_en_bereken(p, tmp, "extras")
    m, mb = wb["Model"], basis["Model"]
    n = basis["Model"]["BY8"].value
    laatste = LY.ROW1 + n - 1
    extra = float(t.aantal) * 3000
    assert abs(float(m[f"{LY.M['model_basis']}{laatste}"].value) - float(mb[f"{LY.M['model_basis']}{laatste}"].value) - extra) < 1e-3
    assert abs(float(m[f"BY{LY.H['model_opbr']}"].value) - float(mb[f"BY{LY.H['model_opbr']}"].value) - extra) < 1e-3
    # de koopsomknop upside (+2 % opbrengsten staat al aan in het voorbeeld) schaalt de extra's niet mee
    assert abs(float(m[f"{LY.M['model_up']}{laatste}"].value) - float(mb[f"{LY.M['model_up']}{laatste}"].value) - extra) < 1e-3
    # tussentijds: in elke rij ligt de extra tussen 0 en aantal × 3000 en loopt hij op
    vorige = 0.0
    for r in range(LY.ROW1, laatste + 1):
        d = float(m[f"{LY.M['model_basis']}{r}"].value) - float(mb[f"{LY.M['model_basis']}{r}"].value)
        assert -1e-6 <= d <= extra + 1e-6 and d >= vorige - 1e-6, (r, d)
        vorige = d
    assert wb["Woningtypes"].cell(LY.WT_R_X_TOTAAL, LY.wt_col(1, 1)).value == 3000


def test_daeb_fees(tmp, basis):
    """DAEB-type (voorbeeldtype 5): fees tellen als bedragen voor het hele type (bedrag of % van het componenttotaal), met het kwartaal
    uit 'jaar Qk', een bouwkwartaal of 'actuals'; koopsom en bouwtermijnen tellen niet; de knoppen raken de fees niet."""
    mb = basis["Model"]
    n = mb["BY8"].value
    laatste = LY.ROW1 + n - 1
    daeb = D.voorbeeld().types[4]
    assert daeb.soort == "DAEB"
    fee = LY.m_col("fee", 5)
    assert mb[f"{fee}5"].value == 1
    totaal = 1192252 * (0.30 + 0.30 + 0.38 + 0.02) + 1636540 * (0.40 + 0.40 + 0.20)
    assert abs(float(mb[f"{fee}6"].value) - totaal) < 1e-3
    assert abs(float(mb[f"{fee}{laatste}"].value) - totaal) < 1e-3
    # 'actuals' valt in het laatste gerealiseerde kwartaal (Q4 '26 = rij van idx 8108), '2027 Q1' een kwartaal later,
    # bouwkwartaal 1 = start bouw (2027 Q3), 5 = vier kwartalen later
    idx = {r: mb[f"{LY.FIX['idx']}{r}"].value for r in range(LY.ROW1, laatste + 1)}
    rij = {v: r for r, v in idx.items()}
    ak = 1192252
    assert abs(float(mb[f"{fee}{rij[8108]}"].value) - ak * 0.98) < 1e-3
    assert abs(float(mb[f"{fee}{rij[8109]}"].value) - (ak + 1636540 * 0.40)) < 1e-3
    assert abs(float(mb[f"{fee}{rij[8111]}"].value) - (ak + 1636540 * 0.80)) < 1e-3
    assert abs(float(mb[f"{fee}{rij[8115]}"].value) - totaal) < 1e-3
    # de fees zitten in model_basis en in alle scenario-kolommen: scenario's verschillen niet door de fees
    p = D.voorbeeld()
    p.types[4].fee_termijnen = []
    p.types[4].fee_comp = []
    wb = bouw_en_bereken(p, tmp, "daeb_zonder_fees")
    m = wb["Model"]
    for r in range(LY.ROW1, laatste + 1):
        d = float(mb[f"{LY.M['model_basis']}{r}"].value) - float(m[f"{LY.M['model_basis']}{r}"].value)
        assert abs(d - float(mb[f"{fee}{r}"].value)) < 1e-3, (r, d)
        for naam in ("model_up", "model_down", "model_scn"):
            a = float(mb[f"{LY.M[naam]}{r}"].value) - float(mb[f"{LY.M['model_basis']}{r}"].value)
            b = float(m[f"{LY.M[naam]}{r}"].value) - float(m[f"{LY.M['model_basis']}{r}"].value)
            assert abs(a - b) < 1e-3, (naam, r, a, b)
        for naam in ("stand_up", "stand_down", "stand_scn"):
            assert abs(float(mb[f"{LY.M[naam]}{r}"].value) - float(m[f"{LY.M[naam]}{r}"].value)) < 1e-3, (naam, r)
    # vorige prognose van de fees (blok 6): 'na omgevingsvergunning' stond in 2026 Q4 (nu 2027 Q1), 'bij gevelsluiting' in bouwkwartaal 4
    # (nu 5, geen vorige start bouw: t.o.v. de huidige start 2027 Q3); zonder vorig kwartaal ('na akkoord SO', actuals) staat vorig = nu
    BT = LY.BT
    rij_bt = lambda i: LY.BT_ROW1 + 4 * LY.N_TERMIJNEN + i - 1
    assert (mb[f"{BT['idxb']}{rij_bt(4)}"].value, mb[f"{BT['idxr']}{rij_bt(4)}"].value) == (8109, 8108)
    assert (mb[f"{BT['idxb']}{rij_bt(8)}"].value, mb[f"{BT['idxr']}{rij_bt(8)}"].value) == (8115, 8114)
    assert mb[f"{BT['idxr']}{rij_bt(1)}"].value == mb[f"{BT['idxb']}{rij_bt(1)}"].value == 8108
    assert mb[f"{BT['fout']}{rij_bt(4)}"].value == 0
    # controles en labels
    d = basis["Dashboard"]
    assert d[f"W{LY.D_ROW_CONTROLES + 12}"].value == "OK" and d[f"W{LY.D_ROW_CONTROLES + 13}"].value == "OK"
    assert mb[f"BY{LY.H['daeb_types']}"].value == 1
    assert mb[f"BY{LY.H['verk_voor_start']}"].value == sum(float(mb[f"{LY.m_col('vcum', k)}5"].value) for k in range(1, 5))
    assert mb[f"BY{LY.H['won_met_start']}"].value == 40 + 20 + 8 + 24
    # fee-component zonder volledige planning: controle slaat aan
    p = D.voorbeeld()
    p.types[4].fee_termijnen[3] = ("na omgevingsvergunning", None, None)         # 98% gepland
    wb = bouw_en_bereken(p, tmp, "daeb_98pct")
    assert str(wb["Dashboard"][f"W{LY.D_ROW_CONTROLES + 12}"].value).startswith("LET OP")
    assert abs(float(wb["Model"][f"{fee}6"].value) - (totaal - ak * 0.02)) < 1e-3
    # ongeldig vorig kwartaal (blok 6): controle 'geldig kwartaal' slaat aan, de huidige planning blijft staan
    p = D.voorbeeld()
    p.types[4].fee_vorig[3] = "ooit"
    wb = bouw_en_bereken(p, tmp, "daeb_vorig_fout")
    assert str(wb["Dashboard"][f"W{LY.D_ROW_CONTROLES + 13}"].value).startswith("LET OP")
    assert wb["Model"][f"{LY.BT['idxr']}{rij_bt(4)}"].value == 99999 and wb["Model"][f"{LY.BT['idxb']}{rij_bt(4)}"].value == 8109
    assert abs(float(wb["Model"][f"{fee}6"].value) - totaal) < 1e-3


def test_grafiek_per_type(basis):
    """Selectieblok van de grafiek per woningtype: type 1 (keuzecel), som van de onderdelen = koopsom x aantal."""
    m = basis["Model"]
    n = m["BY8"].value
    laatste = LY.ROW1 + n - 1
    assert m[f"BY{LY.H['pt_keuze']}"].value == 1 and m[f"BY{LY.H['pt_naam']}"].value == "Rijwoning"
    assert abs(float(m[f"{LY.PT['cum']}{laatste}"].value) - 40 * 385000 / 1e6) < 1e-6
    som = sum(float(m[f"{LY.PT[x]}{r}"].value) for x in ("grond", "bouw", "extra", "fee1", "fee2") for r in range(LY.ROW1, laatste + 1))
    assert abs(som - 40 * 385000 / 1e6) < 1e-6
    assert m[f"BY{LY.H['pt_lbl_grond']}"].value == "Grondtermijn" and (m[f"BY{LY.H['pt_lbl_fee1']}"].value in (None, ""))


def test_geen_kringverwijzingen(tmp):
    """Excel meldt een kringverwijzing zodra een formule (ook via INDEX over een bereik) afhangt van een cel die van haar afhangt."""
    sys.path.insert(0, os.path.join(os.path.dirname(HIER), "tools"))
    import kringverwijzing
    paden = build.bouw(D.voorbeeld(), os.path.join(tmp, "kring"), None, caches=False)
    uit = kringverwijzing.kringen(paden[0])
    assert uit == [], [[kringverwijzing.adres(n) for n in scc[:8]] for scc, _ in uit]


def test_tijdas_twee_jaar_voor_eerste_termijn(basis):
    """De tijd-as van de bouwtermijnengrafiek begint op 1 januari van het jaar twee jaar vóór de eerste termijn (huidig of vorige
    prognose), maar niet vóór de eerste periode; de Excel-grafiek krijgt dezelfde ondergrens als het Model (BY97)."""
    m = basis["Model"]
    BT = LY.BT
    eerste = min(min(m[f"{BT['idxb']}{r}"].value, m[f"{BT['idxr']}{r}"].value)
                 for r in range(LY.BT_ROW1, LY.BT_ROWN + 1) if m[f"{BT['gebruikt']}{r}"].value == 1)
    assert m[f"BY{LY.H['bt_min']}"].value == eerste
    jaar_periodes = int((m[f"{LY.FIX['idx']}{LY.ROW1}"].value - 1) // 4)
    verwacht = max(jaar_periodes, (eerste - 1) // 4 - 2)
    assert m[f"BY{LY.H['jaar_first']}"].value == verwacht and m[f"BY{LY.H['t_first']}"].value == verwacht * 4 + 1
    assert eerste == 2026 * 4 + 3 and verwacht == 2024          # voorbeeld: Rijwoning volgens de vorige prognose vanaf 2026 Q3
