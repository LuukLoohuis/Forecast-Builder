"""Projectgegevens: wat er in de invoercellen komt te staan.

De builder vult het werkboek met een ProjectData. Die komt uit:
  - een bestaand Cashflow_scenario-werkboek (oude indeling met vijf types naast elkaar,
    of de nieuwe indeling met typeblokken), via `lees_werkboek`;
  - of het ingebouwde voorbeeld (`voorbeeld`), zodat een lege build altijd werkt.
"""
from dataclasses import dataclass, field

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L

from . import layout as LY

PERIODE_KOLOMMEN = ["jaar", "kw", "kosten", "pct_kosten", "omzet", "pct_omzet", "cf", "cf1000", "cf_vorig"]  # Invoer B..J

STANDAARD_TERMIJNEN = [
    "Start bouw",
    "Na het leggen van de fundering",
    "Casco gereed",
    "Na gereedkomen buitenmetselwerk",
    "Oplevering woning",
]

STANDAARD_EXTRA = ["Kopersmeerwerk 25%", "Kopersmeerwerk 75%", "Kadastrale kosten / rentes", "Overige opbrengsten", "Verschillen", ""]
# DAEB (blok 5): twee fee-componenten met elk vijf termijnen (mijlpalen uit het fee-overzicht van de gebruiker)
STANDAARD_FEE_COMP = ["AK fee", "Bijkomende kosten (excl. AK)"]
STANDAARD_FEE_MIJLPALEN = ["na akkoord SO", "na akkoord VO", "na akkoord DO", "na omgevingsvergunning", "",
                           "na tekenen TKO", "bij start bouw", "bij gevelsluiting", "bij oplevering", ""]

PARAM_STANDAARD = {
    "actuals_jaar": 2026, "actuals_kw": 3,
    "rente": 0.03, "rente_tm": "start bouw", "rente_basis": "nee",
    "model_aan": "ja",
    "shift_down": 2, "shift_up": -2,
    "uitstel_down": 4, "uitstel_up": 0,
    "opbr_down": -0.03, "opbr_up": 0.02,
    "kosten_down": 0.03, "kosten_up": 0.0,
    # scenariolijn (gele lijn in de cashflowgrafiek): standaard verkoop en transport twee kwartalen eerder
    "shift_scn": -2, "uitstel_scn": 0, "opbr_scn": 0.0, "kosten_scn": 0.0, "scn_aan": "ja", "scn_naam": "Scenario",
    # eigen jaarrente per scenario (None = algemene jaarrente) en koopsom (VON-prijs) in %
    "rente_pct_down": None, "rente_pct_up": None, "rente_pct_scn": None, "koopsom_down": 0.0, "koopsom_up": 0.0, "koopsom_scn": 0.0,
    "norm": 0.7,
    "sjabloon": None,
    # grootste lasten voor dia 2 (Dashboard-sectie): kop van de kaart en tot drie (naam, bedrag in €); leeg = kaart niet op de dia
    "lasten_kop": "GROOTSTE LASTEN",
    "lasten": (),
    # scenario's (tab Scenario's, sinds v18): per scenario een dict met naam, sb (start bouw), vv (verkoop vanaf), tv (transport vanaf)
    # als tekst ('Q3 2027'); leeg = zoals de planning. vs_pp = Naar PowerPoint (ja/nee); vs_sel / vs_type = keuze in de tabel per termijn
    "vs": (),
    "vs_pp": "ja", "vs_cf": "nee",
    "vs_sel": "Scenario 1",
    "vs_type": "alle types",
    # grafiek per woningtype naar PowerPoint (Dashboard W120, sinds v20): totaal zonder DAEB, per type of beide
    "pt_pp": "totaal (zonder DAEB)",
    "projectnaam": "",          # titeldia en voettekst van de dia's; bij Ophalen uit FO uit de bestandsnaam als hij leeg is
}


def standaard_mee(naam):
    """Tab CF kosten: standaardkeuze 'schuift mee met' op de naam van de post (zoals het FO ze noemt). Bouwkosten, meerwerk,
    aansluitkosten en de reserve na oplevering schuiven mee met de start bouw; promotie- en verkoopkosten met de verkoop; de rest
    (grond, honoraria, leges, rente, AK) blijft staan."""
    n = (naam or "").lower()
    if any(w in n for w in ("promotie", "verkoop", "verhuur", "makelaar", "marketing")):
        return "verkoop"
    if any(w in n for w in ("aanneemsom", "bouwkosten", "meerwerk", "aansluit", "reserve", "bouwplaats", "nutsvoorziening")):
        return "start bouw"
    return "vast"


@dataclass
class TypeData:
    naam: str = ""
    aantal: object = None
    koopsom: object = None
    start_jaar: object = None
    start_kw: object = None
    grond_pct: object = None
    termijnen: list = field(default_factory=list)   # [(naam, pct, kw), ...]
    extras: list = field(default_factory=list)      # [(naam, euro per woning, kw of None = bij transport), ...]
    vorig_start_jaar: object = None                 # start bouw volgens de vorige prognose (lichtrode blokjes in de bouwtermijnengrafiek)
    vorig_start_kw: object = None
    soort: object = None                            # None/'niet-DAEB' = koopwoningen; 'DAEB' = alleen fees (blok 5)
    fee_comp: list = field(default_factory=list)    # [(componentnaam, totaal €)] (maximaal N_FEE_COMP)
    fee_termijnen: list = field(default_factory=list)   # [(mijlpaal, bedrag € (> 1) of % (<= 1), kwartaal)] per rij (N_TERMIJNEN rijen, None = leeg)
    fee_vorig: list = field(default_factory=list)   # kwartaal per fee-termijn volgens de vorige prognose (blok 6; None = zoals nu)
    bouwkosten: list = field(default_factory=list)  # blok 7: [(termijn, bedrag € voor het hele type, bouwkwartaal)] (wat de aannemer krijgt)


@dataclass
class ProjectData:
    periodes: list = field(default_factory=list)    # dicts met PERIODE_KOLOMMEN
    verkocht: list = field(default_factory=list)    # per periode: lijst per type
    transport: list = field(default_factory=list)
    types: list = field(default_factory=list)       # TypeData
    params: dict = field(default_factory=lambda: dict(PARAM_STANDAARD))
    termijn_namen: list = field(default_factory=lambda: list(STANDAARD_TERMIJNEN))
    extra_namen: list = field(default_factory=lambda: list(STANDAARD_EXTRA))
    fee_comp_namen: list = field(default_factory=lambda: list(STANDAARD_FEE_COMP))
    fee_mijlpalen: list = field(default_factory=lambda: list(STANDAARD_FEE_MIJLPALEN))
    fo: dict = field(default_factory=dict)          # koppeling met een FO-werkboek: bestand, actuals, opbrengsten, kosten (fo.py)
    # tab CF kosten (sinds v19): per post een dict met naam, mee ('start bouw' / 'verkoop' / 'vast'), gerealiseerd (€ t/m actuals) en
    # bedragen (lijst per periode van het Model, None = leeg)
    kosten_posten: list = field(default_factory=list)

    def type_(self, k):
        """TypeData van blok k (1-based); leeg type als het blok niet gevuld is."""
        return self.types[k - 1] if k - 1 < len(self.types) else TypeData()


def _v(ws, ref):
    v = ws[ref].value
    return None if v == "" else v


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


# ---------------------------------------------------------------------------
def leeg():
    """Leeg sjabloon: standaardknoppen, geen periodes, geen woningtypes; standaardnamen voor termijnen, extra's en fees blijven staan."""
    p = ProjectData()
    p.params = dict(PARAM_STANDAARD)
    return p


def voorbeeld(veel_daeb=False):
    """Klein fictief project: vier woningtypes, twaalf kwartalen gerealiseerd. veel_daeb: drie koop-typen met samen 36
    woningen en 187 DAEB-woningen die in één kwartaal verkocht en geleverd worden (fictief; voor de weergave met veel DAEB:
    verkoopgrafiek alleen koop met DAEB als markering, bouwtermijnen met DAEB-fees)."""
    p = ProjectData()
    kosten = [250000, 180000, 160000, 220000, 300000, 450000, 900000, 1200000, 2600000, 3100000, 3400000, 2900000,
              4100000, 4600000, 4300000, 3900000, 3200000, 2400000, 1500000, 600000]
    omzet = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3800000, 4900000, 6200000, 6900000, 7400000, 7100000, 5200000, 2600000]
    jaar, kw = 2024, 1
    for i in range(len(kosten)):
        p.periodes.append({"jaar": jaar, "kw": f"Q{kw}", "kosten": kosten[i], "omzet": omzet[i],
                           "pct_kosten": None, "pct_omzet": None, "cf": None, "cf1000": None, "cf_vorig": None})
        kw += 1
        if kw == 5:
            jaar, kw = jaar + 1, 1
    p.types = [
        TypeData("Rijwoning", 40, 385000, 2027, 1, 0.30, [("Start bouw", 0.035, 1), ("Na het leggen van de fundering", 0.07, 2),
                                                           ("Casco gereed", 0.28, 4), ("Na gereedkomen buitenmetselwerk", 0.245, 5), ("Oplevering woning", 0.07, 7)]),
        TypeData("Tweekapper", 20, 495000, 2027, 2, 0.32, [("Start bouw", 0.034, 1), ("Na het leggen van de fundering", 0.068, 2),
                                                            ("Casco gereed", 0.272, 4), ("Na gereedkomen buitenmetselwerk", 0.238, 5), ("Oplevering woning", 0.068, 7)]),
        TypeData("Vrijstaand", 8, 685000, 2027, 3, 0.38, [("Start bouw", 0.031, 1), ("Na het leggen van de fundering", 0.062, 2),
                                                           ("Casco gereed", 0.248, 3), ("Na gereedkomen buitenmetselwerk", 0.217, 4), ("Oplevering woning", 0.062, 6)]),
        TypeData("Appartement", 24, 370000, 2027, 1, 0.22, [("Start bouw", 0.039, 1), ("Na het leggen van de fundering", 0.078, 2),
                                                             ("Casco gereed", 0.312, 4), ("Na gereedkomen buitenmetselwerk", 0.273, 6), ("Oplevering woning", 0.078, 7)]),
    ]
    if veel_daeb:
        p.types = p.types[:3]
        for t, aantal in zip(p.types, (12, 14, 10)):
            t.aantal = aantal
    p.types[0].vorig_start_jaar, p.types[0].vorig_start_kw = 2026, 3   # vorige prognose: twee kwartalen eerder (lichtrode blokjes)
    # vijfde type: DAEB (sociale huur, gescheiden koop-/aannemingsovereenkomst): geen koopsom en bouwtermijnen, wel fees (blok 5)
    daeb = TypeData("DAEB", 30, None, 2027, 3, None, [])
    daeb.soort = "DAEB"
    daeb.fee_comp = [("AK fee", 1192252), ("Bijkomende kosten (excl. AK)", 1636540)]
    daeb.fee_termijnen = [("na akkoord SO", 0.30, "actuals"), ("na akkoord VO", 0.30, "actuals"), ("na akkoord DO", 0.38, "actuals"),
                          ("na omgevingsvergunning", 0.02, "2027 Q1"), ("", None, None),
                          ("na tekenen TKO", 0.40, "2027 Q1"), ("bij start bouw", 0.40, 1), ("bij gevelsluiting", 0.20, 5), ("bij oplevering", None, None),
                          ("", None, None)]
    daeb.fee_vorig = [None, None, None, "2026 Q4", None, "2026 Q4", 1, 4, None, None]   # vorige prognose: twee termijnen een kwartaal eerder
    if veel_daeb:
        daeb.aantal = 187
        daeb.fee_comp = [("AK fee", 3100000), ("Bijkomende kosten (excl. AK)", 4300000)]
    p.types.append(daeb)
    n = len(p.periodes)
    tempo = [(8, 10), (4, 5), (2, 4), (6, 4), (30, 12)]   # type -> (per kwartaal, vanaf periode-index); DAEB: alles in één kwartaal
    if veel_daeb:
        tempo = [(6, 13), (6, 13), (4, 14), (187, 11)]       # koop vanaf Q2 ’27, DAEB in één keer in Q4 ’26
    cum = [0, 0, 0, 0, 0]
    for i in range(n):
        rij_v, rij_t = [], []
        for t, (per, vanaf) in enumerate(tempo):
            aantal = p.types[t].aantal
            v = min(per, aantal - cum[t]) if i >= vanaf else 0
            cum[t] += v
            rij_v.append(v or None)
        p.verkocht.append(rij_v)
    for i in range(n):
        rij_t = []
        for t in range(len(tempo)):
            v = p.verkocht[i - 2][t] if i >= 2 else None
            rij_t.append(v)
        p.transport.append(rij_t)
    p.params["actuals_jaar"], p.params["actuals_kw"] = 2026, 4
    p.params["projectnaam"] = "Voorbeeldproject DAEB" if veel_daeb else "Voorbeeldproject"
    p.params["lasten"] = [("Grondkosten", 7900000), ("Honoraria", 3200000), ("Aansluitkosten", 800000)]
    # twee scenario's: de resterende verkoop een jaar later, en de bouw een half jaar later starten
    p.params["vs"] = [{"vv": "Q1 2028"}, {"naam": "Later starten", "sb": "Q3 2027"}]
    p.kosten_posten = _voorbeeld_kostenposten(kosten, 12)
    _voorbeeld_bouwkosten(p, next(k for k in p.kosten_posten if k["naam"] == "Aanneemsom")["bedragen"])
    return p


BK_NAMEN = ["Aanvang bouw", "Fundering", "Ruwe vloer", "Casco", "Gevel", "Dak waterdicht", "Afbouw", "Oplevering"]


def _voorbeeld_bouwkosten(p, aanneemsom):
    """Blok 7 (bouwkosten per termijn) voor het voorbeeld: de post Aanneemsom van tab CF kosten per kwartaal verdeeld over de types
    die dan al gestart zijn (naar koopsom x aantal x (1 - grondtermijn)), per type één termijn per bouwkwartaal. Per kwartaal is de
    som precies de post Aanneemsom, zoals in een FO."""
    idx = [per["jaar"] * 4 + int(per["kw"][1]) for per in p.periodes]
    koop = [t for t in p.types if t.soort != "DAEB" and t.koopsom]
    gewicht = {id(t): t.aantal * t.koopsom * (1 - (t.grond_pct or 0)) for t in koop}
    for i, b in enumerate(aanneemsom):
        if not b:
            continue
        actief = [t for t in koop if t.start_jaar * 4 + t.start_kw <= idx[i]]
        totaal = sum(gewicht[id(t)] for t in actief)
        rest = b
        for n, t in enumerate(actief):
            deel = rest if n == len(actief) - 1 else round(b * gewicht[id(t)] / totaal)
            rest -= deel
            kw = idx[i] - (t.start_jaar * 4 + t.start_kw) + 1
            t.bouwkosten.append((BK_NAMEN[kw - 1] if kw <= len(BK_NAMEN) else f"Bouwkwartaal {kw}", deel, kw))


def _voorbeeld_kostenposten(kosten, n_act):
    """Tab CF kosten voor het voorbeeld: de prognosekosten (na de eerste n_act periodes) verdeeld over posten zoals in een FO, per
    kwartaal precies gelijk aan de kosten op tab Invoer; de bouwkosten zijn de rest. Gerealiseerd: grond, honoraria, overig."""
    n = len(kosten)
    leeg = lambda: [None] * n  # noqa: E731
    hon, aansl, verk, ak, bouw = leeg(), leeg(), leeg(), leeg(), leeg()
    for i in range(n_act, n):
        hon[i] = 200000 if i < n_act + 4 else None
        verk[i] = 100000 if i < n_act + 4 else None
        ak[i] = 150000
    aansl[n_act], aansl[n - 2] = 400000, 400000
    for i in range(n_act, n):
        bouw[i] = kosten[i] - sum(x[i] or 0 for x in (hon, aansl, verk, ak))
    real = sum(kosten[:n_act])
    return [
        {"naam": "Grondkosten", "mee": "vast", "gerealiseerd": 7900000, "bedragen": leeg()},
        {"naam": "Aanneemsom", "mee": "start bouw", "gerealiseerd": None, "bedragen": bouw},
        {"naam": "Honoraria", "mee": "vast", "gerealiseerd": 2400000, "bedragen": hon},
        {"naam": "Promotiekosten", "mee": "verkoop", "gerealiseerd": None, "bedragen": verk},
        {"naam": "Aansluitkosten", "mee": "start bouw", "gerealiseerd": None, "bedragen": aansl},
        {"naam": "Direct doorbelaste projectkosten", "mee": "vast", "gerealiseerd": real - 7900000 - 2400000, "bedragen": ak},
    ]


# ---------------------------------------------------------------------------
def lees_werkboek(pad):
    """Leest de invoer uit een bestaand werkboek (oude of nieuwe indeling)."""
    wb = load_workbook(pad, data_only=True, keep_vba=pad.lower().endswith(".xlsm"))
    if "Woningtypes" in wb.sheetnames:
        return _lees_v2(wb)
    return _lees_v1(wb)


def _lees_periodes(ws, p):
    for r in range(LY.ROW1, LY.ROWN + 1):
        rij = {naam: _v(ws, f"{L(2 + i)}{r}") for i, naam in enumerate(PERIODE_KOLOMMEN)}
        if rij["jaar"] is None:
            rij = {naam: None for naam in PERIODE_KOLOMMEN}
        p.periodes.append(rij)


def _lees_v1(wb):
    """Oude indeling: Invoer L:P verkocht, R:V transport, X8:AB12 types, X18:AH27 termijnen; Dashboard W5.. parameters."""
    p = ProjectData()
    ws = wb["Invoer"]
    _lees_periodes(ws, p)
    for r in range(LY.ROW1, LY.ROWN + 1):
        p.verkocht.append([_num(_v(ws, f"{L(12 + t)}{r}")) for t in range(5)])    # L..P
        p.transport.append([_num(_v(ws, f"{L(18 + t)}{r}")) for t in range(5)])   # R..V
    namen = [_v(ws, f"X{r}") for r in range(19, 28)]
    p.termijn_namen = [n for n in namen if n] or list(STANDAARD_TERMIJNEN)
    for t in range(5):
        naam = _v(ws, f"X{8 + t}")
        aantal = _v(ws, f"Y{8 + t}")
        if naam is None and aantal is None:
            continue
        pc, kc = L(25 + 2 * t), L(26 + 2 * t)      # Y/Z, AA/AB, ...
        termijnen = []
        for r in range(19, 28):
            tn, pct, kw = _v(ws, f"X{r}"), _v(ws, f"{pc}{r}"), _v(ws, f"{kc}{r}")
            if tn is None and pct is None and kw is None:
                continue
            termijnen.append((tn, pct, kw))
        p.types.append(TypeData(naam or "", aantal, _v(ws, f"Z{8 + t}"), _v(ws, f"AA{8 + t}"), _v(ws, f"AB{8 + t}"),
                                _v(ws, f"{pc}18"), termijnen))
    d = wb["Dashboard"]
    p.params.update({
        "actuals_jaar": _v(d, "W5"), "actuals_kw": _v(d, "X5"), "model_aan": _v(d, "W8") or "ja",
        "shift_down": _v(d, "W13"), "shift_up": _v(d, "X13"), "opbr_down": _v(d, "W14"), "opbr_up": _v(d, "X14"),
        "kosten_down": _v(d, "W15"), "kosten_up": _v(d, "X15"), "norm": _v(d, "W18"),
    })
    if "PowerPoint" in wb.sheetnames:
        p.params["sjabloon"] = _lees_sjabloon(wb["PowerPoint"])
    for naam in ("fo_bestand", "fo_actuals", "fo_opbr", "fo_kosten", "fo_datum"):       # FO-koppeling (sinds v8), op label
        label = LY.D_LABELS.get(naam)
        rij = next((r for r in range(1, 120) if _v(d, f"V{r}") == label), None)
        if rij is not None and _v(d, f"W{rij}") not in (None, ""):
            p.fo[{"opbr": "opbrengsten"}.get(naam[3:], naam[3:])] = _v(d, f"W{rij}")
    return p


def _lees_sjabloon(pp):
    """Sjabloonpad op tab PowerPoint: de cel naast het label 'Sjabloon (pptx)' in kolom C (de rij verschilt per versie)."""
    rij = next((r for r in range(1, 200) if _v(pp, f"C{r}") == LY.PP_LABEL_SJABLOON), None)
    return _v(pp, f"D{rij}") if rij is not None else _v(pp, LY.PP_CEL_SJABLOON)


def _lees_v2(wb):
    """Nieuwe indeling: typeblokken op Woningtypes, paren op Invoer, parameters op Dashboard."""
    p = ProjectData()
    ws = wb["Invoer"]
    _lees_periodes(ws, p)
    for r in range(LY.ROW1, LY.ROWN + 1):
        p.verkocht.append([_num(_v(ws, f"{L(LY.in_col(k, 0))}{r}")) for k in range(1, LY.N_TYPES + 1)])
        p.transport.append([_num(_v(ws, f"{L(LY.in_col(k, 1))}{r}")) for k in range(1, LY.N_TYPES + 1)])
    wt = wb["Woningtypes"]
    heeft_extras = isinstance(_v(wt, f"B{LY.WT_R_X_TITEL}"), str) and _v(wt, f"B{LY.WT_R_X_TITEL}").startswith("3 ·")   # blok 3 bestaat sinds v6
    heeft_vorig = isinstance(_v(wt, f"B{LY.WT_R_V_TITEL}"), str) and _v(wt, f"B{LY.WT_R_V_TITEL}").startswith("4 ·")     # blok 4 sinds v7
    heeft_fees = isinstance(_v(wt, f"B{LY.WT_R_D_TITEL}"), str) and _v(wt, f"B{LY.WT_R_D_TITEL}").startswith("5 ·")      # blok 5 sinds v8
    heeft_vf = isinstance(_v(wt, f"B{LY.WT_R_VF_TITEL}"), str) and _v(wt, f"B{LY.WT_R_VF_TITEL}").startswith("6 ·")      # blok 6 sinds v9
    heeft_bk = isinstance(_v(wt, f"B{LY.WT_R_B_TITEL}"), str) and _v(wt, f"B{LY.WT_R_B_TITEL}").startswith("7 ·")        # blok 7 sinds v19
    for k in range(1, LY.N_TYPES + 1):
        c0, c1, c2 = (L(LY.wt_col(k, o)) for o in range(3))
        naam = _v(wt, f"{c0}{LY.WT_R_NAAM}")
        aantal = _v(wt, f"{c0}{LY.WT_R_AANTAL}")
        termijnen = []
        for r in range(LY.WT_R_T1, LY.WT_R_TN + 1):
            tn, pct, kw = _v(wt, f"{c0}{r}"), _v(wt, f"{c1}{r}"), _v(wt, f"{c2}{r}")
            if tn is None and pct is None and kw is None:
                continue
            termijnen.append((tn, pct, kw))
        td = TypeData(naam or "", aantal, _v(wt, f"{c0}{LY.WT_R_KOOPSOM}"), _v(wt, f"{c0}{LY.WT_R_STARTJAAR}"),
                      _v(wt, f"{c0}{LY.WT_R_STARTKW}"), _v(wt, f"{c1}{LY.WT_R_GROND}"), termijnen)
        if heeft_vorig:
            td.vorig_start_jaar = _v(wt, f"{c0}{LY.WT_R_VSTARTJAAR}")
            td.vorig_start_kw = _v(wt, f"{c0}{LY.WT_R_VSTARTKW}")
        if heeft_extras:
            for r in range(LY.WT_R_X1, LY.WT_R_XN + 1):
                xn, eur, kw = _v(wt, f"{c0}{r}"), _v(wt, f"{c1}{r}"), _v(wt, f"{c2}{r}")
                if xn is None and eur is None and kw is None:
                    continue
                td.extras.append((xn, eur, kw))
        if heeft_fees:
            td.soort = _v(wt, f"{c0}{LY.WT_R_SOORT}")
            td.fee_comp = [(_v(wt, f"{c0}{r}"), _v(wt, f"{c1}{r}")) for r in range(LY.WT_R_FC1, LY.WT_R_FCN + 1)]
            td.fee_termijnen = [(_v(wt, f"{c0}{r}"), _v(wt, f"{c1}{r}"), _v(wt, f"{c2}{r}")) for r in range(LY.WT_R_FT1, LY.WT_R_FTN + 1)]
        if heeft_vf:
            td.fee_vorig = [_v(wt, f"{c2}{r}") for r in range(LY.WT_R_VF1, LY.WT_R_VFN + 1)]
        if heeft_bk:
            for r in range(LY.WT_R_B1, LY.WT_R_BN + 1):
                bn, eur, kw = _v(wt, f"{c0}{r}"), _v(wt, f"{c1}{r}"), _v(wt, f"{c2}{r}")
                if bn is None and eur is None and kw is None:
                    continue
                td.bouwkosten.append((bn, eur, kw))
        p.types.append(td)
    # lege blokken aan het eind weglaten
    while p.types and not p.types[-1].naam and p.types[-1].aantal is None and not p.types[-1].bouwkosten:
        p.types.pop()
    namen = [t.termijnen[i][0] for t in p.types for i in range(len(t.termijnen)) if t.termijnen[i][0]]
    if namen:
        p.termijn_namen = list(dict.fromkeys(namen))[:LY.N_TERMIJNEN]
    xnamen = [x[0] for t in p.types for x in t.extras if x[0]]
    if xnamen:
        p.extra_namen = list(dict.fromkeys(xnamen))[:LY.N_EXTRA]
    # fee-componentnamen en mijlpalen: per rij de eerste niet-lege naam over de typen (de rijen zijn vast per component)
    for i in range(LY.N_FEE_COMP):
        namen_i = [t.fee_comp[i][0] for t in p.types if i < len(t.fee_comp) and t.fee_comp[i][0]]
        if namen_i:
            p.fee_comp_namen[i] = namen_i[0]
    for i in range(LY.N_TERMIJNEN):
        namen_i = [t.fee_termijnen[i][0] for t in p.types if i < len(t.fee_termijnen) and t.fee_termijnen[i][0]]
        if namen_i:
            p.fee_mijlpalen[i] = namen_i[0]
    d = wb["Dashboard"]
    # knoppen worden op het label in kolom V gezocht (de rijnummers verschilden per versie); de kolom (W/X/Y) komt uit LY.D.
    # Een knop die het bestand nog niet heeft (oudere versie), houdt de standaardwaarde.
    rijen = {}
    for r in range(1, LY.D_ROW_CONTROLES + 1):
        t = _v(d, f"V{r}")
        if isinstance(t, str) and t.strip() and t.strip() not in rijen:
            rijen[t.strip()] = r
    scn_kolom = _v(d, f"Y{LY.D_ROW_SCENARIO + 1}") == "Scenario"       # oudere v2-bestanden hebben de kolom Scenario nog niet
    paren = {"shift_up": "shift_down", "uitstel_up": "uitstel_down", "opbr_up": "opbr_down", "kosten_up": "kosten_down",
             "rente_pct_up": "rente_pct_down", "koopsom_up": "koopsom_down", "shift_scn": "shift_down", "uitstel_scn": "uitstel_down",
             "opbr_scn": "opbr_down", "kosten_scn": "kosten_down", "rente_pct_scn": "rente_pct_down", "koopsom_scn": "koopsom_down",
             "actuals_kw": "actuals_jaar"}
    for naam, cel in LY.D.items():
        if naam in ("start_project", "dekking", "rente_down", "rente_up", "rente_scn"):
            continue
        if naam.endswith("_scn") and not scn_kolom or naam in ("scn_aan", "scn_naam") and not scn_kolom:
            continue
        label = LY.D_LABELS.get(paren.get(naam, naam))
        rij = next((rijen[lb] for lb in [label] + LY.D_LABELS_OUD.get(naam, []) if lb in rijen), None)
        if rij is None:
            continue
        kol = "".join(ch for ch in cel if ch.isalpha())
        p.params[naam] = _v(d, f"{kol}{rij}")
    # grootste lasten (dia 2, sinds sjabloon v17): sectiekop in kolom V, daaronder de kop van de kaart en drie regels naam/bedrag
    rij_l = next((r for r in range(LY.D_ROW_CONTROLES, 200) if _v(d, f"V{r}") == LY.D_LABEL_LASTEN), None)
    if rij_l is not None:
        kop = _v(d, f"W{rij_l + 1}")
        if kop not in (None, ""):
            p.params["lasten_kop"] = kop
        p.params["lasten"] = [(_v(d, f"W{rij_l + 1 + i}"), _v(d, f"Y{rij_l + 1 + i}")) for i in range(1, LY.N_LASTEN + 1)
                              if _v(d, f"W{rij_l + 1 + i}") not in (None, "") or _v(d, f"Y{rij_l + 1 + i}") not in (None, "")]
    rij_pt = next((r for r in range(LY.D_ROW_CONTROLES, 200) if _v(d, f"V{r}") == LY.D_LABEL_PT_PP), None)
    if rij_pt is not None and _v(d, f"W{rij_pt}") not in (None, ""):
        p.params["pt_pp"] = _v(d, f"W{rij_pt}")
    if "PowerPoint" in wb.sheetnames:
        p.params["sjabloon"] = _lees_sjabloon(wb["PowerPoint"])
    for naam in ("fo_bestand", "fo_actuals", "fo_opbr", "fo_kosten", "fo_datum"):       # FO-koppeling (sinds v8), op label
        label = LY.D_LABELS.get(naam)
        rij = next((r for r in range(1, 120) if _v(d, f"V{r}") == label), None)
        if rij is not None and _v(d, f"W{rij}") not in (None, ""):
            p.fo[{"opbr": "opbrengsten"}.get(naam[3:], naam[3:])] = _v(d, f"W{rij}")
    if LY.SC_TAB in wb.sheetnames:
        _lees_scenarios(wb[LY.SC_TAB], p)
    if LY.KF_TAB in wb.sheetnames:
        _lees_kosten(wb[LY.KF_TAB], p)
    return p


def _lees_kosten(ws, p):
    """Tab CF kosten (sinds v19): de posten onder de koprij 'Kostenpost' (kolom B): naam, schuift mee met, gerealiseerd en een
    bedrag per periode (kolom = periode van het Model)."""
    C = LY.KF_C
    kop = next((r for r in range(1, 40) if _v(ws, f"{C['naam']}{r}") == LY.KF_LABEL_KOP), None)
    if kop is None:
        return
    posten = []
    for r in range(kop + 1, kop + 1 + LY.N_KF):
        naam = _v(ws, f"{C['naam']}{r}")
        if naam is None or not str(naam).strip():
            continue
        mee = _v(ws, f"{C['mee']}{r}")
        mee = str(mee).strip() if mee is not None else ""
        posten.append({"naam": str(naam).strip(), "mee": mee or "vast",                  # leeg rekent het Model als vast
                       "gerealiseerd": _num(_v(ws, f"{C['real']}{r}")),
                       "bedragen": [_num(_v(ws, f"{L(LY.KF_COL1 + i)}{r}")) for i in range(LY.ROWN - LY.ROW1 + 1)]})
    p.kosten_posten = posten


def _lees_scenarios(ws, p):
    """Tab Scenario's (sinds v18): invoer per scenario op het label in kolom B (naam, start bouw, verkoop vanaf, transport vanaf),
    de schakelaar Naar PowerPoint en de keuze in de tabel per termijn."""
    rijen = {}
    for r in range(1, 60):
        t = _v(ws, f"B{r}")
        if isinstance(t, str) and t.strip() and t.strip() not in rijen:
            rijen[t.strip()] = r
    tekst = lambda v: None if v in (None, "") else str(v).strip()  # noqa: E731
    vs = []
    for s in range(1, LY.N_VS + 1):
        col = L(LY.SC_COL[s])
        d = {}
        for veld in ("naam", "sb", "vv", "tv", "gr"):
            r = rijen.get(LY.SC_LABELS[veld])
            if r is not None and tekst(_v(ws, f"{col}{r}")):
                d[veld] = tekst(_v(ws, f"{col}{r}"))
        vs.append(d)
    while vs and not vs[-1]:
        vs.pop()
    p.params["vs"] = vs
    for naam, labels in (("vs_pp", (LY.SC_LABELS["pp"], "Naar PowerPoint")), ("vs_cf", (LY.SC_LABELS["cf"],))):
        r = next((rijen[t] for t in labels if t in rijen), None)
        if r is not None and tekst(_v(ws, f"C{r}")):
            p.params[naam] = tekst(_v(ws, f"C{r}"))
    # keuze in de tabel per termijn: rij met 'Scenario' in kolom B (scenario in C, woningtype in I); werkboeken tot nu: C41/I41
    r = rijen.get("Scenario")
    for naam, cel in (("vs_sel", f"C{r}" if r else LY.SC_CEL_SEL), ("vs_type", f"I{r}" if r else LY.SC_CEL_TYPE)):
        if tekst(_v(ws, cel)):
            p.params[naam] = tekst(_v(ws, cel))
