"""Vaste indeling van het werkboek.

Alles wat per woningtype is, staat in blokken van vaste breedte:
  Woningtypes : drie kolommen per type (naam/%, %, kw), eerste blok in kolom C
  Invoer      : twee kolommen per type (verkocht, transport), eerste paar in kolom L
  Model       : vijf blokken van N_TYPES kolommen achter de hulpcellen (vanaf kolom CA)

Het model leest de typeblokken positioneel (INDEX op een vast bereik), dus wie op tabblad
Woningtypes drie kolommen en op tabblad Invoer twee kolommen invoegt of verwijdert, ziet
alles meeschuiven zonder formules aan te passen.
"""
from openpyxl.utils import get_column_letter as L

N_TYPES = 10          # aantal typeblokken (capaciteit)
N_PERIODS = 60        # modelrijen 8 t/m 67 (maximaal 60 periodes)
ROW1 = 8              # eerste periode-rij op Invoer en in het Model
ROWN = ROW1 + N_PERIODS - 1   # 67
IN_ROWMAX = 500       # Invoer-bereiken lopen van de koprij (7) tot rij 500: rijen invoegen/verwijderen kan ze niet
IN_ANCHOR = ROW1 - 1  # breken (ook niet de eerste of de laatste datarij); het Model leest de i-de gevulde rij
                      # (kolom Bronrij, positie binnen dat bereik; de koprij 'Jaar' is positie 1), dus lege rijen
                      # worden overgeslagen
N_TERMIJNEN = 10      # bouwtermijnen per type (naast de grondtermijn)

# ---- Woningtypes -----------------------------------------------------------
WT_COL1 = 3           # kolom C = eerste typeblok
WT_W = 3              # kolommen per type
WT_R_HEADER = 6       # "TYPE 1" ...
WT_R_NAAM = 7
WT_R_AANTAL = 8
WT_R_KOOPSOM = 9
WT_R_STARTJAAR = 10
WT_R_STARTKW = 11
WT_R_STARTTEKST = 12
WT_R_TERMIJNTITEL = 13
WT_R_TERMIJNKOP = 14
WT_R_GROND = 15
WT_R_T1 = 16                              # eerste bouwtermijn
WT_R_TN = WT_R_T1 + N_TERMIJNEN - 1       # 25
WT_R_TOTAAL = 26
# blok 3: extra opbrengsten per woning in euro's (kopersmeerwerk, kadastrale kosten, overige); kw leeg of 0 = bij transport
N_EXTRA = 6
WT_R_X_TITEL = 31
WT_R_X_KOP = 32
WT_R_X1 = 33
WT_R_XN = WT_R_X1 + N_EXTRA - 1          # 38
WT_R_X_TOTAAL = 39                       # totaal extra per woning
WT_R_X_TOTAAL2 = 40                      # koopsom + extra per woning
# blok 4: vorige prognose (start bouw volgens de vorige prognose) voor de rode blokjes in de bouwtermijnengrafiek
WT_R_V_TITEL = 42
WT_R_VSTARTJAAR = 43
WT_R_VSTARTKW = 44
WT_R_VSTARTTEKST = 45
# blok 5: soort en fees (DAEB = sociale huur met gescheiden koop-/aannemingsovereenkomst: geen koopsom en geen
# bouwtermijnen, maar fees). Twee componenten (AK fee, bijkomende kosten) met een totaalbedrag en elk vijf termijnen; per termijn
# een bedrag in € (> 1) óf een percentage van het componenttotaal (<= 1), en een kwartaal: 'jaar Qk', een bouwkwartaal (getal,
# 1 = start bouw) of 'actuals' (zit al in de eigen cashflow). De tien fee-termijnen nemen in de tijdlijn de plek van de bouwtermijnen in.
N_FEE_COMP = 2
N_FEE_PER_COMP = N_TERMIJNEN // N_FEE_COMP     # 5
WT_R_D_TITEL = 47
WT_R_SOORT = 48
WT_R_D_NOTE = 49
WT_R_F_KOP = 50
WT_R_FC1 = 51                                  # component 1: naam, totaal €, gepland
WT_R_FCN = WT_R_FC1 + N_FEE_COMP - 1           # 52
WT_R_FT_KOP = 53
WT_R_FT1 = 54                                  # tien fee-termijnen: rijen 54-58 component 1, 59-63 component 2
WT_R_FTN = WT_R_FT1 + N_TERMIJNEN - 1          # 63
WT_R_F_TOTAAL = 64
# blok 6: vorige prognose van de fees (DAEB): per mijlpaal het kwartaal volgens de vorige prognose (lichtrood in de tijdlijn)
WT_R_VF_TITEL = 66
WT_R_VF_KOP = 67
WT_R_VF1 = 68
WT_R_VFN = WT_R_VF1 + N_TERMIJNEN - 1          # 77
WT_LASTCOL = WT_COL1 + N_TYPES * WT_W - 1  # AH bij 10 types
WT_RANGE_END = "ZZ"   # positionele bereiken lopen tot ZZ: invoegen/verwijderen van kolommen kan ze niet breken


def wt_col(k, offset=0):
    """Kolomnummer van blok k (1..N) op tabblad Woningtypes; offset 0 = naam/%, 1 = %, 2 = kw."""
    return WT_COL1 + (k - 1) * WT_W + offset


def wt_ref(row, k, offset=0, rows=None):
    """Positionele verwijzing naar een cel (of kolomvector) in blok k op Woningtypes.

    rows=None  -> INDEX(Woningtypes!$A$row:$ZZ$row,1,col)
    rows=(a,b) -> INDEX(Woningtypes!$A$a:$ZZ$b,0,col)   (hele kolom van dat rijbereik)
    """
    col = wt_col(k, offset)
    if rows is None:
        return f"INDEX(Woningtypes!$A${row}:${WT_RANGE_END}${row},1,{col})"
    a, b = rows
    return f"INDEX(Woningtypes!$A${a}:${WT_RANGE_END}${b},0,{col})"


# ---- Invoer ------------------------------------------------------------------
IN_COL1 = 12          # kolom L = eerste paar (verkocht, transport)
IN_W = 2
IN_R_TYPE = 6         # typenaam boven het paar
IN_R_KOP = 7          # 'Verkocht' / 'Transport'
IN_LASTCOL = IN_COL1 + N_TYPES * IN_W - 1   # AE bij 10 types


def in_col(k, offset=0):
    """Kolomnummer op tabblad Invoer van type k; offset 0 = verkocht, 1 = transport."""
    return IN_COL1 + (k - 1) * IN_W + offset


def in_ref(row, k, offset=0):
    """Positionele verwijzing naar de invoercel van type k voor modelrij `row` (via de bronrij in kolom B)."""
    return f"INDEX(Invoer!$A${IN_ANCHOR}:${WT_RANGE_END}${IN_ROWMAX},${FIX['bron']}{row},{in_col(k, offset)})"


def in_col_ref(col_letter, row):
    """Invoercel in kolom `col_letter` (B..J) voor modelrij `row`, via de bronrij."""
    return f"INDEX(Invoer!${col_letter}${IN_ANCHOR}:${col_letter}${IN_ROWMAX},${FIX['bron']}{row})"


def in_rng(col_letter, col2=None):
    """Invoerbereik van de koprij tot rij IN_ROWMAX, bv. Invoer!$B$7:$B$500."""
    return f"Invoer!${col_letter}${IN_ANCHOR}:${col2 or col_letter}${IN_ROWMAX}"


# ---- Model -------------------------------------------------------------------
# vaste kolommen A..L, modelkolommen M..BF, hulpcellen BX:BY, typeblokken vanaf CA
FIX = {"nr": "A", "bron": "B", "kw": "C", "kwartaal": "D", "kwartaal_vol": "E", "idx": "F", "fase": "G",
       "kosten": "H", "opbr": "I", "stand": "J", "cum_opbr": "K", "vorig": "L"}
M = {}
for i, name in enumerate([
    "model_basis", "model_up", "model_down",            # L M N
    "cum_opbr_up", "cum_opbr_down",                     # O P
    "opbr_up", "opbr_down", "kosten_up", "kosten_down", # Q R S T
    "rente_basis", "stand_basis_rente",                 # U V
    "rente_up", "stand_up_raw", "rente_down", "stand_down_raw",   # W X Y Z
    "stand_basis", "kosten_basis", "stand_up", "stand_down",       # AA AB AC AD  (getoond)
    "pos_basis", "pos_up", "pos_down", "pos_vorig",     # AE AF AG AH
    # scenariolijn (gele lijn in de cashflowgrafiek, eigen knoppen in de derde kolom op het Dashboard)
    "model_scn", "cum_opbr_scn", "opbr_scn", "kosten_scn", "rente_scn", "stand_scn_raw", "stand_scn", "pos_scn",
    # grafiekkolommen
    "g_voorfinanciering", "g_positief_saldo", "g_opbrengsten", "g_kosten", "g_stand", "g_vorige",
    "g_band_onder", "g_bandbreedte", "g_downside", "g_upside",
    "g_punt_nu", "g_punt_dal", "g_eind_upside", "g_eind_downside", "g_eind_basis",
    "g_verkocht", "g_getransporteerd", "g_verkocht_cum", "g_getransporteerd_cum",
    "g_verkocht_pct", "g_getransporteerd_pct", "g_opbrengsten_pct", "g_kosten_pct",
    "g_verkocht_pct_lbl", "g_getransporteerd_pct_lbl",           # label boven de stapel: cumulatief % (alleen in kwartalen met een stapel)
    "g_scenario", "g_eind_scenario", "g_realisatie",            # scenariolijn, eindpunt, waas (1 in gerealiseerde kwartalen)
    "g_punt_uitverkocht", "g_punt_alles_transport",             # mijlpalen in de verkoopgrafiek (top/onderkant van de stapel)
    "stap_ok",                                          # controle: idx loopt met de juiste stap op
], start=13):
    M[name] = L(i)

M_HULP_LABEL = "BX"
M_HULP = "BY"
M_SPACER = "BW"
M_BLOK1 = 79          # kolom CA
# verkocht cum, transport cum, transport cum up/down, termijnen vervallen, transport cum scenariolijn,
# grafiek: verkocht per kwartaal (gv) en getransporteerd per kwartaal als negatief getal (gt); benoemde bereiken g_v<k>, g_t<k>
# vx: extra opbrengsten per woning (€) vervallen per bouwkwartaal; rij 6 = extra's bij transport
# fee: rij 5 = DAEB (1/0), rij 6 = totaal fees (gepland, €), rijen = fees vervallen t/m de periode (cumulatief, €)
BLOKKEN = ["vcum", "tcum", "tup", "tdown", "verv", "tscn", "gv", "gt", "vx", "fee"]


# ---- bouwtermijnentabel in het Model: tijdlijn met per bouwtermijn een baan (lane) per woningtype ----
# Raster (rijen BT_ROW1..): alle (type k, termijn i)-combinaties met hun kwartaalindex (huidig en vorige prognose), de rang van de
# termijnnaam (volgorde van eerste voorkomen; dezelfde naam bij meerdere typen = één groep) en de positie in de compacte lijst.
# Compacte lijst (zelfde rijen, kolommen vanaf 'nr'): rij 1 = koprij jaartallen, rij 2 = koprij kwartaalnummers, daarna per termijn
# de banen van de typen die hem hebben (in typevolgorde) en een lege scheidingsrij tussen de termijnen. Tijd = kwartaalindex
# (jaar × 4 + kwartaal), dus één eenheid per kwartaal; de as loopt van 1 januari van het eerste jaar tot het einde van het laatste.
BT_KOP = 2                              # koprijen (jaar, kwartaal)
BT_LANES = N_TYPES * N_TERMIJNEN        # 100 banen maximaal
BT_N = BT_KOP + BT_LANES + 30           # 132 rijen: koprijen, banen en tot 30 scheidingsrijen
BT_ROW1 = ROW1
BT_ROWN = BT_ROW1 + BT_N - 1            # 139
BT_COL1 = M_BLOK1 + len(BLOKKEN) * N_TYPES + 1
BT_NJ = 16                              # jaarreeksen (koprij 1): blokjes van vier kwartaalen met het jaartal als label
BT_NK = 4 * BT_NJ                       # kwartaalreeksen (koprij 2): blokjes van één kwartaal met '1'..'4' als label
BT_PREV = [f"prev{k}" for k in range(1, N_TYPES + 1)]     # vorige prognose per type (lichtrood)
BT_CUR = [f"cur{k}" for k in range(1, N_TYPES + 1)]       # huidige planning per type (typekleur)
BT_JR = [f"jr{y}" for y in range(1, BT_NJ + 1)]
BT_KW = [f"kw{q}" for q in range(1, BT_NK + 1)]
# reeksen in stapelvolgorde: stapel 1 (eerste as) en stapel 2 (tweede as, bovenop)
BT_STAPEL1 = ["s0", "s1", "s2"] + BT_PREV + ["s4", "s5"] + BT_JR
BT_STAPEL2 = ["s6"] + BT_CUR + ["s8"] + BT_KW
BT_REEKSEN = BT_STAPEL1 + BT_STAPEL2    # kolommen van het PowerPoint-blok (na 'label') en de benoemde bereiken g_bt_*
BT = {}
for _i, _name in enumerate([
    "j", "k", "i", "gebruikt",                        # raster: alle (type k, termijn i)-combinaties, gebruikt = 1/0
    "knaam", "tnaam", "mnaam",                        # raster: typenaam, groepsnaam (termijn, of fee-component bij DAEB), baannaam
    "comp", "bedrag", "waarde", "kw", "kwr", "fout",  # raster: fee-component (0 = bouwtermijn), € van de baan (fee), invoer, kwartaalcel,
                                                      # kwartaalcel vorige prognose (blok 6), fout
    "idxb", "idxr",                                   # raster: kwartaalindex huidig/vorig (99999 = geen)
    "uniek", "rang", "pos", "eerste",                 # raster: eerste gebruikte rij per groepsnaam; rang; positie in de lijst;
                                                      # eerste = 1 bij de eerste baan van de groep (die krijgt het label)
    "nr", "bron", "soort", "label",                   # compact: nr = j van de r-de unieke termijn; bron = rasterrij van de baan;
                                                      # soort 1 jaar / 2 kwartaal / 3 baan / 4 scheiding / 0 leeg
    "kb", "ib", "ir",                                 # compact: type, kwartaalindex huidig en vorig van de baan
] + BT_REEKSEN):
    BT[_name] = L(BT_COL1 + _i)
BT_SEG_NAMEN = {"s0": "·", "s1": "Gerealiseerd", "s2": "·", "s4": "Gerealiseerd", "s5": "·", "s6": "·", "s8": "·"}   # vaste reeksnamen

# benoemde bereiken van de verkoop- en transportgrafiek (tijdvak vanaf het actuals-jaar): naam -> Modelkolom
VK_NAMEN = {"g_vk_kwartaal": FIX["kwartaal"], "g_vk_realisatie": M["g_realisatie"], "g_vk_verkocht": M["g_verkocht"],
            "g_vk_getransporteerd": M["g_getransporteerd"], "g_vk_punt_uitverkocht": M["g_punt_uitverkocht"],
            "g_vk_punt_alles_transport": M["g_punt_alles_transport"],
            "g_vk_pct": M["g_verkocht_pct"], "g_vk_tpct": M["g_getransporteerd_pct"],
            "g_vk_pct_lbl": M["g_verkocht_pct_lbl"], "g_vk_tpct_lbl": M["g_getransporteerd_pct_lbl"]}


# ---- grafiek per woningtype: selectieblok (het type uit de keuzecel op het Dashboard), € mln per kwartaal ----
PT_COL1 = BT_COL1 + len(BT) + 1
PT = {}
for _i, _name in enumerate(["grond", "bouw", "extra", "fee1", "fee2", "cum", "totaal", "eind"]):
    PT[_name] = L(PT_COL1 + _i)
PT_REEKSEN = ["grond", "bouw", "extra", "fee1", "fee2", "cum", "totaal", "eind"]   # totaal = som per kwartaal (label), eind = laatste punt
PT_KLEUREN = {"grond": "002060", "bouw": "2A78D6", "extra": "EDA100", "fee1": "4A3AA7", "fee2": "1F8A8A"}


def m_col(blok, k):
    """Kolomletter in het Model van blok `blok` voor type k."""
    return L(M_BLOK1 + BLOKKEN.index(blok) * N_TYPES + (k - 1))


def m_range(blok, row):
    """Rijbereik (type 1..N) van blok `blok` in rij `row`, bv. CA8:CJ8."""
    return f"{m_col(blok, 1)}{row}:{m_col(blok, N_TYPES)}{row}"


def m_range_abs(blok, row):
    return f"${m_col(blok, 1)}${row}:${m_col(blok, N_TYPES)}${row}"


# hulpcellen in kolom BY (rij -> omschrijving); de formules staan in sheets.py
H = {
    "n": 8, "idx_actuals": 9, "pos_actuals": 10, "kw_actuals": 11, "stand_nu": 12,
    "model_aan": 13, "vorig_aanwezig": 14, "opbr_totaal": 15, "kosten_totaal": 16, "cum_opbr_actuals": 17,
    "dal_basis": 18, "dal_basis_kw": 19, "eind_basis": 20, "be_basis": 21, "be_basis_pos": 22,
    "dal_up": 23, "dal_up_kw": 24, "eind_up": 25, "be_up": 26, "be_up_pos": 27,
    "dal_down": 28, "dal_down_kw": 29, "eind_down": 30, "be_down": 31, "be_down_pos": 32,
    "laatste_kw": 33, "dal_vorig": 34, "eind_vorig": 35, "vorig_nu": 36, "be_vorig": 37, "be_vorig_pos": 38,
    "woningen": 39, "verkocht_plan": 40, "transport_plan": 41, "verkocht_actuals": 42, "transport_actuals": 43,
    "model_opbr": 44, "verk_bij_start": 45, "won_met_start": 46, "verk_bij_start_pct": 47,
    "uitverkocht": 48, "alles_transport": 49,
    "lbl_nu": 50, "lbl_dal": 51, "lbl_eind_up": 52, "lbl_eind_down": 53, "lbl_eind_basis": 54,
    "rente": 55, "rente_tm_start": 56, "rente_in_basis": 57, "start_idx": 58, "start_tekst": 59,
    "rente_eind_basis": 60, "rente_eind_up": 61, "rente_eind_down": 62,
    "rente_basis_tot": 63, "rente_up_tot": 64, "rente_down_tot": 65, "rente_up_extra": 66, "rente_down_extra": 67,
    "rente_tekst": 68, "aantal_types": 69,
    # scenariolijn
    "scn_aan": 70, "scn_naam": 71, "eind_scn": 72, "dal_scn": 73, "dal_scn_kw": 74, "be_scn": 75, "be_scn_pos": 76,
    "rente_eind_scn": 77, "rente_scn_tot": 78, "rente_scn_extra": 79, "lbl_scenario": 80, "lbl_eind_scn": 81,
    "lbl_nu_cf": 82, "lbl_eind_basis_cf": 83,
    # verkoopgrafiek
    "uitverkocht_pos": 84, "alles_transport_pos": 85, "lbl_uitverkocht": 86, "lbl_alles_transport": 87,
    "nog_verkopen": 88, "nog_transport": 89,
    # eigen jaarrente en koopsomknop per scenario
    "rente_pct_down": 90, "rente_pct_up": 91, "rente_pct_scn": 92, "koopsom_down": 93, "koopsom_up": 94, "koopsom_scn": 95,
    # bouwtermijnengrafiek (de VBA leest bt_n voor het aantal rijen en t_first/t_end voor de tijd-as): tijd = kwartaalindex
    "bt_n": 96, "t_first": 97, "t_nu_end": 98, "t_end": 99, "jaar_first": 100, "jaar_last": 101, "lbl_realisatie": 102,
    "bt_types": 103, "bt_jaren": 104, "bt_lanes": 105, "bt_uniek": 106,
    # fees (DAEB) en de grafiek per woningtype (keuzecel op het Dashboard)
    "fee_fout_gepland": 107, "fee_fout_termijn": 108, "daeb_types": 109,
    "pt_keuze": 110, "pt_naam": 111, "pt_daeb": 112, "pt_aantal": 113, "pt_totaal": 114, "pt_nu": 115, "pt_titel": 116, "pt_subtitel": 117,
    "pt_lbl_grond": 118, "pt_lbl_bouw": 119, "pt_lbl_extra": 120, "pt_lbl_fee1": 121, "pt_lbl_fee2": 122, "pt_lbl_cum": 123,
    "bt_min": 124,          # eerste termijn in de bouwtermijnengrafiek (kwartaalindex, huidig of vorig; 99999 = geen)
    "pt_first": 125, "pt_jaar_first": 126, "pt_start": 127, "pt_n": 128,   # grafiek per type: eerste opbrengst, eerste jaar, startpositie, aantal
    "pt_lbl_totaal": 129, "pt_lbl_eind": 130,
    "vp_start": 131, "vp_n": 132,   # verkoopgrafiek: eerste getoonde periode (1 januari van het actuals-jaar) en aantal periodes
    "vp_tekst": 133,                # verkoopgrafiek: voortgang per type als tekst ('Rijwoning 25%, Tweekapper 40%, ...')
    # marge (dia 2 en Dashboard): DAEB heeft geen marge, dus marge niet-DAEB = projectresultaat / (opbrengsten - fees DAEB)
    "opbr_daeb": 134, "woningen_daeb": 135, "marge": 136, "marge_nd": 137, "marge_tekst": 138, "marge_split": 139,
    # dia 2 (sjabloon v17): as-grenzen van de kleine cashflowgrafiek (€ mln; de VBA leest kc_min/kc_max) en de grootste lasten
    "top_basis": 140, "top_vorig": 141, "kc_min": 142, "kc_max": 143, "lasten_max": 144,
    # scenario's (tab Scenario's, sinds v18): planning, schakelaar, as van de scenariografiek, teksten; per scenario hieronder
    "vs_vnd_act": 145, "vs_tnd_act": 146, "vs_vv_pos": 147, "vs_vv_plan": 148, "vs_tr_pos": 149, "vs_tr_plan": 150, "vs_gap": 151,
    "vs_aan": 152, "vs_pp": 153, "vs_pp_aan": 154, "vs_min": 155, "vs_max": 156, "vs_stap": 157, "vs_titel": 158, "vs_subtitel": 159,
    "vs_eind_tekst": 160, "vs_slechtst": 161, "vs_sb_mee": 162, "vs_sel": 163, "vs_typesel": 164, "vs0_leg": 165, "vs0_sub": 166,
    "vs0_delta": 167, "vs0_lbl_dal": 168, "vs_kicker": 169, "vs0_dal_pos": 170, "vs_nu_pos": 171, "vs_tr_tekst": 172,
    "vs_cf": 173,                   # scenariolijnen ook in de cashflowgrafiek (Dashboard en dia 3), schakelaar op tab Scenario's
}


def h(name):
    """Absolute verwijzing naar een hulpcel in het Model, bv. $BY$8."""
    return f"${M_HULP}${H[name]}"


def hm(name):
    """Zelfde, met tabbladnaam (voor andere tabbladen)."""
    return f"Model!{h(name)}"


# ---- Dashboard (parameters in kolom V/W/X/Y, toelichting in Z) ----------------
D = {
    "actuals_jaar": "W5", "actuals_kw": "X5",
    "rente": "W6", "rente_tm": "W7", "rente_basis": "W8", "start_project": "W9",
    "model_aan": "W12", "dekking": "W13",
    "shift_down": "W17", "shift_up": "X17",
    "uitstel_down": "W18", "uitstel_up": "X18",
    "opbr_down": "W19", "opbr_up": "X19",
    "kosten_down": "W20", "kosten_up": "X20",
    "rente_pct_down": "W21", "rente_pct_up": "X21",          # eigen jaarrente per scenario (leeg = algemene jaarrente)
    "koopsom_down": "W22", "koopsom_up": "X22",              # koopsom (VON-prijs) in % via het verkooptempo-model
    "rente_down": "W23", "rente_up": "X23",                  # info: rente totaal
    # scenariolijn (derde kolom Y; de koptekst 'Scenario' in Y16 laat data.py zien dat deze kolom bestaat)
    "shift_scn": "Y17", "uitstel_scn": "Y18", "opbr_scn": "Y19", "kosten_scn": "Y20", "rente_pct_scn": "Y21", "koopsom_scn": "Y22",
    "rente_scn": "Y23", "scn_aan": "Y25", "scn_naam": "Y26",
    "norm": "W29",
    "pt_keuze": "B120",                                     # grafiek per woningtype: gekozen type (naam; keuzelijst uit Model rij 7)
    # FO-koppeling (macro 'Ophalen uit FO' op tab Invoer, of build.py --fo): bestand, actuals t/m, totalen uit '1. Cashflow', datum
    "fo_bestand": "W69", "fo_actuals": "W70", "fo_opbr": "W71", "fo_kosten": "W72", "fo_datum": "W73",
    # grootste lasten (dia 2, kaart KPI9): kop van de kaart en drie regels naam (W:X samengevoegd) en bedrag in € (Y)
    "lasten_kop": "W76", "last1_naam": "W77", "last1_bedrag": "Y77", "last2_naam": "W78", "last2_bedrag": "Y78",
    "last3_naam": "W79", "last3_bedrag": "Y79",
}
N_LASTEN = 3
# data.py leest de knoppen op het label in kolom V (zo blijven oudere bestanden met andere rijnummers leesbaar)
D_LABELS = {
    "actuals_jaar": "Actuals t/m (jaar · kwartaal)", "rente": "Jaarrente", "rente_tm": "Rente t/m", "rente_basis": "Rente ook in de basis",
    "model_aan": "Woningtypes en termijnen gebruiken",
    "shift_down": "Verkoop en transport verschuiven (kw)", "uitstel_down": "Uitstel start bouw (kw)", "opbr_down": "Opbrengsten",
    "kosten_down": "Kosten", "rente_pct_down": "Jaarrente per scenario", "koopsom_down": "Koopsom (VON-prijs)",
    "scn_aan": "Scenariolijn tonen in de cashflowgrafiek", "scn_naam": "Naam van de scenariolijn", "norm": "Norm verkocht bij start bouw",
    "fo_bestand": "FO-bestand", "fo_actuals": "Actuals in het FO t/m", "fo_opbr": "Opbrengsten totaal volgens het FO",
    "fo_kosten": "Kosten totaal volgens het FO", "fo_datum": "Opgehaald op",
}
# oudere labels van dezelfde knop (data.py zoekt ook hierop, zodat een bestand van vóór de hernoeming leesbaar blijft)
D_LABELS_OUD = {"norm": ["Norm verkocht vóór start bouw"]}
D_ROW_SCENARIO = 15        # sectie SCENARIO'S (kop Downside/Upside/Scenario in rij 16)
D_ROW_VERKOOP = 28
D_ROW_POWERPOINT = 34      # knop staat op W35:X36 (vast: de VBA-macro zet hem daar)
D_ROW_CONTROLES = 40
D_ROW_TYPES = 56           # overzicht woningtypes (na vijftien controles)
D_ROW_FO = 68              # FO-koppeling: bestand, actuals, totalen (cellen in LY.D fo_*)
D_ROW_LASTEN = 75          # grootste lasten voor dia 2 (sectiekop; data.py zoekt de sectie op dit label)
D_LABEL_LASTEN = "GROOTSTE LASTEN (dia 2)"
D_ROW_LEGENDA = 70         # cel-legenda van de verkoopgrafiek (gekleurde cellen per woningtype)
D_ROW_BT = 95              # sectiekop bouwtermijnengrafiek; chips (typekleuren) in rij 96, grafiek op B97 (19 rijen)
D_ROW_LEGENDA_BT = 96
D_ROW_PT = 118             # sectiekop grafiek per woningtype; keuzecel en chips in rij 120, grafiek op B121 (12 rijen)
D_ROW_LEGENDA_PT = 120
D_CHART_ANCHORS = {"scenario": "B15", "cashflow": "B44", "verkoop": "B71", "transport": "B83", "bouwtermijnen": "B97", "pertype": "B121"}
D_LEGENDA_CELLEN = ["B", "C", "D", "E", "G", "H", "I", "J", "L", "M"]   # chip per typeblok 1..N_TYPES
D_LEGENDA_PT_CELLEN = ["G", "H", "I", "J", "L"]                          # chips van de grafiek per type (PT_REEKSEN zonder cum)
D_ROW_UITLEG = 136


def d(name):
    return f"Dashboard!${D[name][0]}${D[name][1:]}" if len(D[name]) == 2 else f"Dashboard!${D[name][:-2]}${D[name][-2:]}"


def dabs(name):
    """Absolute verwijzing met tabbladnaam, bv. Dashboard!$W$17."""
    cell = D[name]
    col = "".join(ch for ch in cell if ch.isalpha())
    row = "".join(ch for ch in cell if ch.isdigit())
    return f"Dashboard!${col}${row}"


# ---- PowerPoint ----------------------------------------------------------------
PP_KPI_ROW1 = 8
# grafiekblokken (kop in rij 7; de VBA leest dezelfde kopcellen): cf dia 3 (8 kolommen), sc dia 4 (10), bt dia 5 bouwtermijnen
# (PP_BT_KOL kolommen: label + alle reeksen van BT_REEKSEN, BT_N rijen), vp dia 6 verkoop/transport per type (26 kolommen: kwartaal,
# 10x verkocht, 10x transport, totalen, mijlpalen, realisatie; twee grafieken lezen hetzelfde blok), vo dia 7 (5)
PP_BT_KOL = 1 + len(BT_REEKSEN)        # 108
PP_PT_KOL = 2 + len(PT_REEKSEN)        # 10: kwartaal, acht reeksen, realisatie
PP_VP_KOL = 1 + 2 * N_TYPES + 5 + 4    # 30: kwartaal, verkocht en getransporteerd per type, verkocht, getransporteerd, twee mijlpalen, realisatie,
                                       # cumulatief % en label (verkocht), cumulatief % en label (transport); de macro: KOL_VP
_vp = 27 + PP_BT_KOL + 1               # 136 (EF)
_vo = _vp + PP_VP_KOL + 1              # 167 (FK)
_pt = _vo + 6                          # 173 (FQ): dia 8, grafiek per woningtype (de macro vult per type een kopie van de dia)
PP_BLOK = {"cf": "G", "sc": "P", "bt": "AA", "vp": L(_vp), "vo": L(_vo), "pt": L(_pt)}
PP_TABEL_SC = f"{L(_pt + PP_PT_KOL + 1)}7"                      # GB7: 9 rijen x 4 kolommen
PP_TABEL_VT_ROW = 18           # kop in BR18, daaronder N_TYPES rijen
# dia 2 (sjabloon v17): kleine cashflowgrafiek KPI_GRAFIEK, blok 'kc' rechts van de tabellen (8 kolommen: jaarlabel, cum.
# cashflow, vorige prognose, nu, dieptepunt, break-even, eindsaldo, €0; de laatste zes als één-punts-reeksen)
PP_KC_KOL = 8
_kc = _pt + PP_PT_KOL + 1 + 9                                    # GK: na de tabel woningtypes (8 kolommen vanaf GB)
PP_BLOK["kc"] = L(_kc)
PP_CEL_SJABLOON = "D130"       # (de VBA leest dezelfde cellen: CEL_SJABLOON / CEL_NAAM; data.py zoekt op het label in kolom C)
PP_CEL_NAAM = "D131"           # (tot v17: D76/D77; sinds v18 lager, er zijn meer KPI-regels door de dia Verkoopscenario's)
PP_LABEL_SJABLOON = "Sjabloon (pptx)"
# dia 2: grafiekkader en tekenvlak van KPI_GRAFIEK in inch (links, boven, rechts, onder), zoals tools/pptx_v17_kerncijfers.py ze
# zet (manualLayout); de VBA rekent met de fracties KC_PLOT_X/KC_PLOT_W de x van het 'nu'-kwartaal uit (grijs vlak, nu-blok)
KC_KADER = (0.30, 2.58, 9.50, 3.98)
KC_PLOT = (0.50, 2.76, 8.98, 3.98)
KC_PLOT_X = round((KC_PLOT[0] - KC_KADER[0]) / (KC_KADER[2] - KC_KADER[0]), 4)
KC_PLOT_W = round((KC_PLOT[2] - KC_PLOT[0]) / (KC_KADER[2] - KC_KADER[0]), 4)


# ---- Scenario's (sinds v18): start bouw, verkoop en transport per scenario --------------------------------------------------------
# Tab 'Scenario's': per scenario alleen invullen wat anders is dan de planning (leeg = automatisch). Het rekenwerk staat in het Model
# vanaf VS_COL1: per scenario verschoven verkoop (vc) en transport (tc) per type, bouwtermijnen (vv) en extra's (vx) met de start
# bouw van het scenario, en de kolommen opbrengsten, kosten, rente en stand zoals bij de scenariolijn.
N_VS = 3                                                  # scenario's naast de Basis
VS_KLEUREN = ["002060", "6250D6", "C2399B", "7A2E6E"]     # Basis (navy) en scenario 1..3: zelfde kleur op tab, Dashboard en dia's
SC_TAB = "Scenario's"
SC_BLAD = "'Scenario''s'"                                 # tabbladnaam in formules
SC_COL = [3, 5, 7, 9]                                     # Basis, scenario 1..3: twee kolommen per scenario (C:D, E:F, G:H, I:J)
SC_R_NAAM, SC_R_SB, SC_R_VV, SC_R_TV = 17, 18, 19, 20     # invoerrijen
SC_R_PP = 22                                              # Naar PowerPoint (ja/nee): dia 4
SC_CEL_PP = "C22"
SC_R_CF = 23                                              # lijnen in de cashflowgrafiek (ja/nee): Dashboard en dia 3
SC_CEL_CF = "C23"
SC_R_FOUT = 24                                            # controle per scenario (tekst)
SC_CEL_SEL = "C42"                                        # tabel per termijn: scenario (Basis / Scenario 1..3)
SC_CEL_TYPE = "I42"                                       # tabel per termijn: woningtype (alle types / naam)
SC_N_KW = 24                                              # kwartalen in de tabel per termijn (vanaf 1 januari van het actuals-jaar)
SC_LABELS = {"naam": "Naam", "sb": "Start bouw (hele project)", "vv": "Verkoop vanaf · het verkooptempo schuift mee",
             "tv": "Transport vanaf · nooit vóór de verkoop", "pp": "Naar PowerPoint (dia Verkoopscenario's)",
             "cf": "Lijnen in de cashflowgrafiek"}


def sc_in(s, veld):
    """Absolute verwijzing naar de invoercel van scenario s (1..N_VS) op tab Scenario's; veld: naam, sb, vv, tv."""
    rij = {"naam": SC_R_NAAM, "sb": SC_R_SB, "vv": SC_R_VV, "tv": SC_R_TV}[veld]
    return f"{SC_BLAD}!${L(SC_COL[s])}${rij}"


VS_COL1 = PT_COL1 + len(PT) + 1
VS_G = {}                                                 # gedeelde kolommen
_c = VS_COL1
# kosten vóór / vanaf start bouw (cumulatief), verkocht / transport niet-DAEB; prognosekosten van de posten op tab CF kosten die
# meeschuiven met de start bouw (cksb) of met de verkoop (ckv), cumulatief
for _naam in ["kpre", "kpost", "vnd", "tnd", "cksb", "ckv"]:
    VS_G[_naam] = L(_c)
    _c += 1
_c += 1
VS_AGG = ["pv", "pt", "psb", "qf", "model", "cumo", "opbr", "kcum", "kosten", "rente", "sraw", "stand", "pos", "tnd", "g", "gdal", "gcf"]
VS_TB = ["vc", "tc", "vv", "vx"]                          # per type: verkocht cum, transport cum, termijnen vervallen, extra's vervallen
_VS = {}
for _s in range(1, N_VS + 1):
    for _naam in VS_AGG:
        _VS[(_s, _naam)] = L(_c)
        _c += 1
    for _blok in VS_TB:
        for _k in range(1, N_TYPES + 1):
            _VS[(_s, _blok, _k)] = L(_c)
            _c += 1
    _c += 1
# selectie voor de tabel per termijn (keuze scenario en woningtype op tab Scenario's): per type transport (ts; rij 6 = start bouw),
# extra's (xs) en fees (fs) van het gekozen scenario; cumulatief per termijn (tj1..), grond, extra's, fees, verkocht, getransporteerd;
# opbrengsten, kosten en stand van het gekozen scenario; inhaal per termijn (ih1..: termijn die al vervallen was bij transport)
SEL_TB = ["ts", "xs", "fs"]
SEL = {}
for _blok in SEL_TB:
    for _k in range(1, N_TYPES + 1):
        SEL[(_blok, _k)] = L(_c)
        _c += 1
for _naam in (["vj", "trj", "gj", "xj", "fj", "osel", "ksel", "ssel"] + [f"tj{_j}" for _j in range(1, N_TERMIJNEN + 1)]
              + [f"ih{_j}" for _j in range(1, N_TERMIJNEN + 1)]):
    SEL[_naam] = L(_c)
    _c += 1
_c += 1
# termijnmatrix (rijen onder de periodes): per termijn j en type k de naam, het bedrag per woning (€) en het bouwkwartaal;
# kolom 'label' = rijnaam in de tabel; rij TM_R_SEL = 1 voor de gekozen typen, TM_R_AG = grondtermijn per woning (€)
TM = {_k: L(_c + _k - 1) for _k in range(1, N_TYPES + 1)}
TM["label"] = L(_c + N_TYPES)
_c += N_TYPES + 2
VS_LAST = _c - 1
TM_R_NAAM, TM_R_A, TM_R_KW = 75, 85, 95                   # j = 1..N_TERMIJNEN onder elkaar
TM_R_SEL, TM_R_AG = 106, 107


def vs_col(s, naam, k=None):
    """Kolomletter in het Model: aggregaat `naam` van scenario s, of (met k) typeblok `naam` (vc/tc/vv/vx) voor type k."""
    return _VS[(s, naam)] if k is None else _VS[(s, naam, k)]


def vs_rng(s, blok, row, abs_=False):
    """Rijbereik van typeblok `blok` van scenario s in rij `row` (type 1..N)."""
    a, b = vs_col(s, blok, 1), vs_col(s, blok, N_TYPES)
    return f"${a}${row}:${b}${row}" if abs_ else f"{a}{row}:{b}{row}"


def sel_rng(blok, row, abs_=False):
    a, b = SEL[(blok, 1)], SEL[(blok, N_TYPES)]
    return f"${a}${row}:${b}${row}" if abs_ else f"{a}{row}:{b}{row}"


def tm_rng(row):
    return f"${TM[1]}${row}:${TM[N_TYPES]}${row}"


# hulpcellen per scenario (vs1_*, vs2_*, ...): invoer (kwartaalindex, 0 = leeg of ongeldig), verschuivingen, naam, uitkomst, teksten
VS_VELDEN = ["in_sb", "in_vv", "in_tv", "fout", "aan", "d_sb", "sb", "vv", "d_v", "tv", "d_t", "naam", "kop", "leg", "dal", "dal_pos",
             "dal_kw", "be_pos", "be", "eind", "rente", "vbsb", "vbsb_pct", "et_pos", "et", "d_fin", "d_be", "sub", "delta", "lbl_dal",
             "kaart", "pp", "cf"]
_r = 180
for _s in range(1, N_VS + 1):
    for _veld in VS_VELDEN:
        H[f"vs{_s}_{_veld}"] = _r
        _r += 1

# PowerPoint: cashflowgrafiek (dia 3) met de scenariolijnen: blok van 11 kolommen (tot v17 8 kolommen vanaf G7); dia 4
# 'Verkoopscenario's': grafiekblok (kwartaal/jaartal, Basis, scenario 1..3, dieptepunt Basis en scenario 1..3), tijdvak vanaf
# 1 januari van het actuals-jaar (vp_start, vp_n), zoals de verkoopgrafiek
PP_CF_KOL = 11
PP_VS_KOL = 9
PP_BLOK["cf"] = L(_kc + PP_KC_KOL + 1)                    # GT
PP_BLOK["vs"] = L(_kc + PP_KC_KOL + 1 + PP_CF_KOL + 1)     # HF
# dia 4: grafiekkader en tekenvlak van VS_GRAFIEK in inch (zoals tools/pptx_v18_scenarios.py ze zet); de VBA rekent met de fracties
# VS_PLOT_X/VS_PLOT_W de x van het 'nu'-kwartaal uit (grijs vlak VS_REALISATIE)
VS_KADER = (0.40, 2.74, 9.55, 5.00)
VS_PLOT = (0.80, 2.82, 9.42, 4.70)
VS_PLOT_X = round((VS_PLOT[0] - VS_KADER[0]) / (VS_KADER[2] - VS_KADER[0]), 4)
VS_PLOT_W = round((VS_PLOT[2] - VS_PLOT[0]) / (VS_KADER[2] - VS_KADER[0]), 4)

# tab CF kosten (sinds v19): kosten per post en per periode, zoals het FO-tabblad 'CF - kosten' (blok 'Onderdelen'); per post
# 'schuift mee met' (start bouw / verkoop / vast) voor de scenario's op tab Scenario's, en de kosten t/m het dieptepunt (grootste
# lasten op dia 2). Kolom = periode van het Model (1..60, zoals de rijen op tab Invoer); gerealiseerd = één bedrag in kolom E.
KF_TAB = "CF kosten"
KF_BLAD = "'CF kosten'"
N_KF = 20                                                 # kostenposten
KF_R_POS = 6                                              # verborgen: positie van de periode boven elke kolom
KF_R_KOP = 7
KF_R1 = 8
KF_RN = KF_R1 + N_KF - 1
KF_R_TOT, KF_R_INV, KF_R_VERSCHIL, KF_R_CONTROLE = KF_RN + 1, KF_RN + 2, KF_RN + 3, KF_RN + 4
KF_C = {"naam": "B", "mee": "C", "totaal": "D", "real": "E", "dal": "F", "code": "G", "sleutel": "H"}
KF_COL1 = 9                                               # kolom I = periode 1
KF_MEE = ["start bouw", "verkoop", "vast"]
KF_LABEL_KOP = "Kostenpost"


def kf_rng(veld):
    """Absoluut bereik van een kolom van de posten op tab CF kosten (met tabbladnaam)."""
    c = KF_C[veld]
    return f"{KF_BLAD}!${c}${KF_R1}:${c}${KF_RN}"


def kf_blok():
    """Absoluut bereik van alle periodekolommen van de posten (N_KF rijen x 60 kolommen)."""
    return f"{KF_BLAD}!${L(KF_COL1)}${KF_R1}:${L(KF_COL1 + ROWN - ROW1)}${KF_RN}"


_r = 285
for _naam in ["kf_aan", "kf_sb_tot", "kf_v_tot", "vs_v_mee", "kf_tekst", "kf_tekst_kort", "kf_dal_tot", "kf_verschil", "kf_controle",
              "kl_bron", "kl1_naam", "kl1_bedrag", "kl2_naam", "kl2_bedrag", "kl3_naam", "kl3_bedrag",
              "kf_dal_real", "kf_dal_lbl", "kf_real_gap", "kf_buiten"]:
    H[_naam] = _r
    _r += 1
