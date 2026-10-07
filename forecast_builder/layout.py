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
BLOKKEN = ["vcum", "tcum", "tup", "tdown", "verv", "tscn", "gv", "gt", "vx"]


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
BT_PREV = [f"prev{k}" for k in range(1, N_TYPES + 1)]     # vorige prognose per type (lichte tint)
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
    "knaam", "tnaam", "idxb", "idxr",                 # raster: typenaam, termijnnaam, kwartaalindex huidig/vorig
    "uniek", "rang", "pos", "eerste",                 # raster: eerste gebruikte rij per naam; rang van de naam; positie in de lijst;
                                                      # eerste = 1 bij de eerste baan van de groep (die krijgt het label)
    "nr", "bron", "soort", "label",                   # compact: nr = j van de r-de unieke termijn; bron = rasterrij van de baan;
                                                      # soort 1 jaar / 2 kwartaal / 3 baan / 4 scheiding / 0 leeg
    "kb", "ib", "ir",                                 # compact: type, kwartaalindex huidig en vorig van de baan
] + BT_REEKSEN):
    BT[_name] = L(BT_COL1 + _i)
BT_SEG_NAMEN = {"s0": "·", "s1": "Gerealiseerd", "s2": "·", "s4": "Gerealiseerd", "s5": "·", "s6": "·", "s8": "·"}   # vaste reeksnamen


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
    "model_opbr": 44, "verk_voor_start": 45, "won_met_start": 46, "verk_voor_start_pct": 47,
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
}
# data.py leest de knoppen op het label in kolom V (zo blijven oudere bestanden met andere rijnummers leesbaar)
D_LABELS = {
    "actuals_jaar": "Actuals t/m (jaar · kwartaal)", "rente": "Jaarrente", "rente_tm": "Rente t/m", "rente_basis": "Rente ook in de basis",
    "model_aan": "Woningtypes en termijnen gebruiken",
    "shift_down": "Verkoop en transport verschuiven (kw)", "uitstel_down": "Uitstel start bouw (kw)", "opbr_down": "Opbrengsten",
    "kosten_down": "Kosten", "rente_pct_down": "Jaarrente per scenario", "koopsom_down": "Koopsom (VON-prijs)",
    "scn_aan": "Scenariolijn tonen in de cashflowgrafiek", "scn_naam": "Naam van de scenariolijn", "norm": "Norm verkocht vóór start bouw",
}
D_ROW_SCENARIO = 15        # sectie SCENARIO'S (kop Downside/Upside/Scenario in rij 16)
D_ROW_VERKOOP = 28
D_ROW_POWERPOINT = 34      # knop staat op W35:X36 (vast: de VBA-macro zet hem daar)
D_ROW_CONTROLES = 40
D_ROW_TYPES = 53           # overzicht woningtypes (na twaalf controles)
D_ROW_LEGENDA = 70         # cel-legenda van de verkoopgrafiek (gekleurde cellen per woningtype)
D_ROW_BT = 95              # sectiekop bouwtermijnengrafiek; chips (typekleuren) in rij 96, grafiek op B97 (19 rijen)
D_ROW_LEGENDA_BT = 96
D_CHART_ANCHORS = {"scenario": "B15", "cashflow": "B44", "verkoop": "B71", "transport": "B83", "bouwtermijnen": "B97"}
D_LEGENDA_CELLEN = ["B", "C", "D", "E", "G", "H", "I", "J", "L", "M"]   # chip per typeblok 1..N_TYPES
D_ROW_UITLEG = 118


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
PP_BT_KOL = 1 + len(BT_REEKSEN)        # 109
PP_BLOK = {"cf": "G", "sc": "P", "bt": "AA", "vp": L(27 + PP_BT_KOL + 1), "vo": L(27 + PP_BT_KOL + 1 + 27)}   # vp EG, vo FH
PP_TABEL_SC = f"{L(27 + PP_BT_KOL + 1 + 27 + 6)}7"            # FN7: 9 rijen x 4 kolommen
PP_TABEL_VT_ROW = 18           # kop in BR18, daaronder N_TYPES rijen
PP_CEL_SJABLOON = "D56"        # (de VBA leest dezelfde cellen: CEL_SJABLOON / CEL_NAAM)
PP_CEL_NAAM = "D57"
