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
N_PERIODS = 60        # rijen 8 t/m 67
ROW1 = 8              # eerste periode-rij (Invoer en Model lopen gelijk)
ROWN = ROW1 + N_PERIODS - 1   # 67
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
    """Positionele verwijzing naar de invoercel van type k in periode-rij `row`."""
    return f"INDEX(Invoer!$A${ROW1}:${WT_RANGE_END}${ROWN},{row - ROW1 + 1},{in_col(k, offset)})"


# ---- Model -------------------------------------------------------------------
# vaste kolommen A..K, modelkolommen L..BE, hulpcellen BX:BY, typeblokken vanaf CA
M = {}
for i, name in enumerate([
    "model_basis", "model_up", "model_down",            # L M N
    "cum_opbr_up", "cum_opbr_down",                     # O P
    "opbr_up", "opbr_down", "kosten_up", "kosten_down", # Q R S T
    "rente_basis", "stand_basis_rente",                 # U V
    "rente_up", "stand_up_raw", "rente_down", "stand_down_raw",   # W X Y Z
    "stand_basis", "kosten_basis", "stand_up", "stand_down",       # AA AB AC AD  (getoond)
    "pos_basis", "pos_up", "pos_down", "pos_vorig",     # AE AF AG AH
    # grafiekkolommen
    "g_voorfinanciering", "g_positief_saldo", "g_opbrengsten", "g_kosten", "g_stand", "g_vorige",
    "g_band_onder", "g_bandbreedte", "g_downside", "g_upside",
    "g_punt_nu", "g_punt_dal", "g_eind_upside", "g_eind_downside", "g_eind_basis",
    "g_verkocht", "g_getransporteerd", "g_verkocht_cum", "g_getransporteerd_cum",
    "g_verkocht_pct", "g_getransporteerd_pct", "g_opbrengsten_pct", "g_kosten_pct",
], start=12):
    M[name] = L(i)

M_HULP_LABEL = "BX"
M_HULP = "BY"
M_SPACER = "BW"
M_BLOK1 = 79          # kolom CA
BLOKKEN = ["vcum", "tcum", "tup", "tdown", "verv"]   # verkocht cum, transport cum, transport cum up/down, termijnen vervallen


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
}


def h(name):
    """Absolute verwijzing naar een hulpcel in het Model, bv. $BY$8."""
    return f"${M_HULP}${H[name]}"


def hm(name):
    """Zelfde, met tabbladnaam (voor andere tabbladen)."""
    return f"Model!{h(name)}"


# ---- Dashboard (parameters in kolom V/W/X, toelichting in Y) ----------------
D = {
    "actuals_jaar": "W5", "actuals_kw": "X5",
    "rente": "W6", "rente_tm": "W7", "rente_basis": "W8", "start_project": "W9",
    "model_aan": "W12", "dekking": "W13",
    "shift_down": "W17", "shift_up": "X17",
    "uitstel_down": "W18", "uitstel_up": "X18",
    "opbr_down": "W19", "opbr_up": "X19",
    "kosten_down": "W20", "kosten_up": "X20",
    "rente_down": "W21", "rente_up": "X21",
    "norm": "W24",
}
D_ROW_POWERPOINT = 31      # knop staat op W32:X33 (vast: de VBA-macro zet hem daar)
D_ROW_CONTROLES = 37
D_ROW_TYPES = 45           # overzicht woningtypes
D_CHART_ANCHORS = {"scenario": "B15", "cashflow": "B44", "verkoop": "B70"}
D_ROW_UITLEG = 93


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
PP_BLOK = {"cf": "G", "sc": "O", "vt": "Z", "vo": "AF"}     # kopcellen rij 7 (VBA leest G7, O7, Z7, AF7)
PP_TABEL_SC = "AL7"            # 9 rijen x 4 kolommen
PP_TABEL_VT_ROW = 18           # kop in AL18, daaronder N_TYPES rijen
PP_CEL_SJABLOON = "D50"
PP_CEL_NAAM = "D51"
