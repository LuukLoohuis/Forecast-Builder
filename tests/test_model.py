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
    na_ok = {LY.FIX["vorig"], LY.M["g_vorige"], LY.M["g_punt_nu"], LY.M["g_punt_dal"], LY.M["g_eind_upside"], LY.M["g_eind_downside"], LY.M["g_eind_basis"]}
    fouten = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("#"):
                    col = "".join(ch for ch in c.coordinate if ch.isalpha())
                    ok = v == "#N/A" and ((ws.title == "Model" and (col in na_ok or (col in chart and c.row > 7 + n)))
                                          or (ws.title == "PowerPoint" and (c.row > 7 + n or col in ("M", "U", "V", "W", "X"))))
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
