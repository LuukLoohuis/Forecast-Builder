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


def _xml(jaar_min, jaar_max=None):
    r = _Namen()
    r_cf = _Namen({f"vs{s}": f"[0]!g_vs{s}" for s in range(1, LY.N_VS + 1)})      # plus de scenario's van tab Scenario's
    cashflow = cx.cashflow(r_cf, _lbl(scenario_naam="lbl_scenario", nu="lbl_nu_cf", dal="lbl_dal", eind_basis="lbl_eind_basis_cf",
                                      eind_scenario="lbl_eind_scn"))
    band = cx.band(r, _lbl(nu="lbl_nu", dal="lbl_dal", eind_up="lbl_eind_up", eind_down="lbl_eind_down"))
    lbl_vt = _lbl(uitverkocht="lbl_uitverkocht", alles_transport="lbl_alles_transport", daeb_v="vp_daeb_lbl_v", daeb_t="vp_daeb_lbl_t",
                  titel_v="vp_titel_v", titel_t="vp_titel_t")
    r_vk = {naam[5:]: f"[0]!{naam}" for naam in LY.VK_NAMEN}            # kwartaal, realisatie, verkocht, ... in het tijdvak
    verkoop = cx.verkoop_paneel(r_vk, lbl_vt, _type_namen(), "verkocht")
    transport = cx.verkoop_paneel(r_vk, lbl_vt, _type_namen(), "transport")
    r_bt = {naam: f"[0]!g_bt_{naam}" for naam in ["label"] + LY.BT_REEKSEN}
    r_bt.update({f"naam_{naam}": f"Model!${LY.BT[naam]}$7" for naam in LY.BT_PREV + LY.BT_CUR + LY.BT_JR + LY.BT_FL})
    bouwtermijnen = cx.bouwtermijnen(r_bt, jaar_min * 4 + 1, (jaar_max + 1) * 4 + 1 if jaar_max else None)
    r_pt = {naam: f"[0]!g_pt_{naam}" for naam in LY.PT_REEKSEN}
    r_pt.update({"kwartaal": "[0]!g_pt_kwartaal", "realisatie": "[0]!g_pt_realisatie"})
    r_pt.update({f"naam_{naam}": f"Model!{LY.h('pt_lbl_' + naam)}" for naam in LY.PT_REEKSEN})
    pertype = cx.type_paneel(r_pt)
    # tab Scenario's: tijdvak vanaf het actuals-jaar (benoemde bereiken g_vsw_*), labels van de dieptepunten uit het Model
    r_vs = {naam: f"[0]!g_vsw_{naam}" for naam in VSW_NAMEN}
    lbl_vs = {"dal0": f"Model!{LY.h('vs0_lbl_dal')}"}
    lbl_vs.update({f"dal{s}": f"Model!{LY.h(f'vs{s}_lbl_dal')}" for s in range(1, LY.N_VS + 1)})
    scenarios = cx.scenarios(r_vs, lbl_vs)
    return cashflow, band, verkoop, transport, bouwtermijnen, pertype, scenarios


def _type_namen():
    """Per type: (naamcel, verkocht per kwartaal, getransporteerd per kwartaal) via benoemde bereiken g_v<k> / g_t<k>."""
    return [(f"Model!${LY.m_col('gv', k)}$7", f"[0]!g_v{k}", f"[0]!g_t{k}") for k in range(1, LY.N_TYPES + 1)]


BT_H_CM = 10.05   # bouwtermijnengrafiek: 19 Dashboard-rijen (koprijen + per termijn een baan per type)


# tab Scenario's: benoemde bereiken g_vsw_<naam> (tijdvak vanaf 1 januari van het actuals-jaar, zoals de verkoopgrafiek) -> Modelkolom
VSW_NAMEN = {"kwartaal": LY.FIX["kwartaal"], "realisatie": LY.M["g_realisatie"], "basis": LY.M["g_stand"], "dal0": LY.M["g_punt_dal"]}
VSW_NAMEN.update({f"vs{s}": LY.vs_col(s, "g") for s in range(1, LY.N_VS + 1)})
VSW_NAMEN.update({f"dal{s}": LY.vs_col(s, "gdal") for s in range(1, LY.N_VS + 1)})
VS_NAMEN = {f"g_vs{s}": LY.vs_col(s, "gcf") for s in range(1, LY.N_VS + 1)}   # cashflowgrafiek op het Dashboard (alle periodes; schakelaar)


def grafieken(jaar_min=2020, jaar_max=None):
    """[(naam, anker, breedte cm, hoogte cm, xml)] in de volgorde chart1.xml … chart7.xml (zes op het Dashboard, dan tab Scenario's).
    jaar_min/jaar_max = eerste en laatste jaar van de periodes: voorlopige grenzen van de tijd-as in de bouwtermijnengrafiek
    (kwartaalindex jaar × 4 + 1); cache.py zet ze na het doorrekenen op Model!BY97 (twee jaar vóór de eerste termijn) en BY99."""
    cashflow, band, verkoop, transport, bouwtermijnen, pertype, scenarios = _xml(jaar_min, jaar_max)
    A = LY.D_CHART_ANCHORS
    return [("scenario", A["scenario"], 30.3, 13.0, band), ("cashflow", A["cashflow"], 30.3, 11.5, cashflow),
            ("verkoop", A["verkoop"], 30.3, cx.PANEEL_H_CM, verkoop), ("transport", A["transport"], 30.3, cx.PANEEL_H_CM, transport),
            ("bouwtermijnen", A["bouwtermijnen"], 30.3, BT_H_CM, bouwtermijnen), ("pertype", A["pertype"], 30.3, cx.PT_H_CM, pertype),
            ("scenarios", SCENARIO_ANKER, 20.2, 11.6, scenarios)]


SCENARIO_ANKER = "L16"     # tab Scenario's (scenarios.GRAFIEK_ANKER)


def plaats_grafieken(ws_dashboard, wb, jaar_min=2020, jaar_max=None):
    """Zet plaatsvervangende grafieken op het Dashboard en op tab Scenario's (zelfde plek en maat als de echte). openpyxl nummert
    de grafieken per tabblad in de volgorde van de tabbladen (Dashboard, Scenario's, ...), dus chart1..6 = Dashboard, chart7 = tab
    Scenario's: dezelfde volgorde als grafieken()."""
    ws_model = wb["Model"]
    for naam, anker, w, h, _ in grafieken(jaar_min, jaar_max):
        ch = LineChart()
        ch.add_data(Reference(ws_model, min_col=3, min_row=8, max_row=9), titles_from_data=False)
        ch.width, ch.height = w, h
        ch.anchor = anker
        (wb[LY.SC_TAB] if naam == "scenarios" else ws_dashboard).add_chart(ch)


def chart_xmls(jaar_min=2020, jaar_max=None):
    return [g[4] for g in grafieken(jaar_min, jaar_max)]
