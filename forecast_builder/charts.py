"""Grafieken op het Dashboard.

De grafiek-XML komt uit chartxml.py: reeksen op de benoemde bereiken g_* (zo lang als er periodes zijn) en
reeksnamen uit de labelcellen in het Model (kolom BY). openpyxl zet eerst een plaatsvervangende grafiek neer,
zodat de tekening en relaties kloppen; daarna vervangt package.py de inhoud door deze XML.
"""
from openpyxl.chart import LineChart, Reference

from . import chartxml as cx, layout as LY


class _Namen(dict):
    """reeksnaam -> benoemd bereik [0]!g_<naam>; ontbrekende namen zijn een fout bij het bouwen, niet in Excel."""
    def __missing__(self, key):
        naam = f"g_{key}"
        if naam not in LY.M and key != "kwartaal":
            raise KeyError(f"geen grafiekkolom {naam} in layout.M")
        return f"[0]!{naam}"


def _lbl(**namen):
    return {k: f"Model!{LY.h(v)}" for k, v in namen.items()}


def _xml():
    r = _Namen()
    cashflow = cx.cashflow(r, _lbl(scenario_naam="lbl_scenario", nu="lbl_nu_cf", dal="lbl_dal", eind_basis="lbl_eind_basis_cf",
                                   eind_scenario="lbl_eind_scn"))
    band = cx.band(r, _lbl(nu="lbl_nu", dal="lbl_dal", eind_up="lbl_eind_up", eind_down="lbl_eind_down"))
    verkoop = cx.verkoop_types(r, _lbl(uitverkocht="lbl_uitverkocht", alles_transport="lbl_alles_transport"), _type_namen())
    return cashflow, band, verkoop


def _type_namen():
    """Per type: (naamcel, verkocht per kwartaal, getransporteerd per kwartaal) via benoemde bereiken g_v<k> / g_t<k>."""
    return [(f"Model!${LY.m_col('gv', k)}$7", f"[0]!g_v{k}", f"[0]!g_t{k}") for k in range(1, LY.N_TYPES + 1)]


def grafieken():
    """[(naam, anker, breedte cm, hoogte cm, xml)] in de volgorde chart1.xml, chart2.xml, chart3.xml."""
    cashflow, band, verkoop = _xml()
    A = LY.D_CHART_ANCHORS
    return [("scenario", A["scenario"], 30.3, 13.0, band), ("cashflow", A["cashflow"], 30.3, 11.5, cashflow),
            ("verkoop", A["verkoop"], 30.3, 10.0, verkoop)]


def plaats_grafieken(ws_dashboard, wb):
    """Zet drie plaatsvervangende grafieken op het Dashboard (zelfde plek en maat als de echte)."""
    ws_model = wb["Model"]
    for naam, anker, w, h, _ in grafieken():
        ch = LineChart()
        ch.add_data(Reference(ws_model, min_col=3, min_row=8, max_row=9), titles_from_data=False)
        ch.width, ch.height = w, h
        ch.anchor = anker
        ws_dashboard.add_chart(ch)


def chart_xmls():
    return [g[4] for g in grafieken()]
