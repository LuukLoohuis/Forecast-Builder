"""Bouwt de vijf tabbladen. Alle formules staan hier."""
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

from . import chartxml as CX, layout as LY
from .layout import FIX, H, M, N_TYPES, ROW1, ROWN, h, hm, dabs, m_col, m_range_abs
from .styles import *  # noqa: F401,F403

MINUS = "−"       # typografisch minteken, zoals in het bestaande werkboek
AMBER_TXT = "B87400"   # tekstkleur van de scenariolijn (donkerder dan de lijn zelf, leesbaar op wit)
VERKOOP_TITEL = ("Verkoop en transport per woningtype en per kwartaal · aantal woningen · cijfer boven de stapel = totaal per kwartaal · "
                 "grijs vlak = gerealiseerde kwartalen")
APOS = "’"
AUTEUR = "Claude"


def _comment(text):
    c = Comment(text, AUTEUR)
    c.width, c.height = 320, 110
    return c


def eur_m(ref, dec=1):
    """Formulefragment: bedrag als '−€1,2M' (M = mln)."""
    return f'IF({ref}<0,"{MINUS}","")&"€"&FIXED(ABS({ref})/1000000,{dec})&"M"'


def eur_mln(ref, dec=1):
    return f'IF({ref}<0,"{MINUS}","")&"€"&FIXED(ABS({ref})/1000000,{dec})&" mln"'


# =============================================================================
#  WONINGTYPES
# =============================================================================
def bouw_woningtypes(wb, data):
    ws = wb.create_sheet("Woningtypes")
    ws.sheet_properties.codeName = "shWoningtypes"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"{L(LY.WT_COL1)}{LY.WT_R_NAAM}"
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 36
    for k in range(1, N_TYPES + 1):
        ws.column_dimensions[L(LY.wt_col(k, 0))].width = 27
        ws.column_dimensions[L(LY.wt_col(k, 1))].width = 9
        ws.column_dimensions[L(LY.wt_col(k, 2))].width = 7
    ws.row_dimensions[LY.WT_R_TERMIJNKOP].height = 18
    ws.row_dimensions[4].height = 22      # ruimte voor de knoppen Type invoegen / Type verwijderen (.xlsm)

    put(ws, "B1", "Woningtypes — per type één blok van drie kolommen", f=F_TITLE)
    put(ws, "B2", "Blauw op lichtblauw = invoer. Naam en aantal altijd; koopsom, start bouw en termijnen alleen als je het "
                  "verkooptempo-model gebruikt (schakelaar op tab Dashboard). Elk type heeft zijn eigen termijnen.", f=F_NOTE)
    put(ws, "B3", f"Type toevoegen: vul het eerstvolgende lege blok, of voeg op de gewenste plek drie kolommen in (en twee kolommen op tab "
                  f"Invoer). Type verwijderen: haal hier drie kolommen weg en op tab Invoer twee. Alles schuift mee; maximaal {N_TYPES} types. "
                  "Met macro's: knoppen 'Type invoegen' en 'Type verwijderen' (alleen in de .xlsm).", f=F_NOTE)
    put(ws, f"B{LY.WT_R_HEADER - 1}", "1 · WONINGTYPES · naam, aantal, koopsom en start bouw", f=F_SECTION)
    labels = {
        LY.WT_R_NAAM: "Type (naam)",
        LY.WT_R_AANTAL: "Aantal woningen",
        LY.WT_R_KOOPSOM: "Koopsom per woning (€)",
        LY.WT_R_STARTJAAR: "Start bouw · jaar",
        LY.WT_R_STARTKW: "Start bouw · kwartaal (1-4)",
        LY.WT_R_STARTTEKST: "Start bouw (controle)",
        LY.WT_R_GROND: "Grondtermijn (bij notarieel transport)",
        LY.WT_R_TOTAAL: "Totaal bouwtermijnen (100% van de aanneemsom, óf 100% − grondtermijn) · kw = bouwtijd",
    }
    for r, t in labels.items():
        put(ws, f"B{r}", t, f=F_NOTE, al=AL_VCENTER)
    for i in range(LY.N_TERMIJNEN):
        put(ws, f"B{LY.WT_R_T1 + i}", f"Bouwtermijn {i + 1}", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_X_TITEL}", "3 · EXTRA OPBRENGSTEN PER WONING · € per woning, bovenop de koopsom (kopersmeerwerk, kadastrale kosten, "
                                   "overige) · kw = bouwkwartaal waarin het binnenkomt · leeg of 0 = bij notarieel transport", f=F_SECTION)
    ws[f"B{LY.WT_R_X_KOP}"].comment = _comment(
        "Extra opbrengsten per woning in euro's, naast de koopsom: kopersmeerwerk, kadastrale kosten/rentes, overige opbrengsten, "
        "verschillen. Per regel een bedrag per woning en het bouwkwartaal waarin het binnenkomt (zoals bij de termijnen); leeg of 0 = "
        "in het kwartaal van notarieel transport. Het model telt ze mee in de opbrengst per woning; de koopsomknop op het Dashboard "
        "werkt er niet op.")
    for i in range(LY.N_EXTRA):
        put(ws, f"B{LY.WT_R_X1 + i}", f"Extra opbrengst {i + 1}", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_X_TOTAAL}", "Totaal extra per woning (€)", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_X_TOTAAL2}", "Koopsom + extra per woning (€)", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_V_TITEL}", "4 · VORIGE PROGNOSE · start bouw volgens de vorige prognose: de blokjes in lichte tint in de bouwtermijnengrafiek "
                                   "(Dashboard en dia 5) · leeg = geen vorige prognose", f=F_SECTION)
    put(ws, f"B{LY.WT_R_VSTARTJAAR}", "Start bouw vorige prognose · jaar", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_VSTARTKW}", "Start bouw vorige prognose · kwartaal (1-4)", f=F_NOTE, al=AL_VCENTER)
    put(ws, f"B{LY.WT_R_VSTARTTEKST}", "Vorige prognose (controle)", f=F_NOTE, al=AL_VCENTER)
    ws[f"B{LY.WT_R_VSTARTJAAR}"].comment = _comment("Start bouw zoals die in de vorige prognose stond. De bouwtermijnen schuiven dan in de grafiek "
                                                     "van de lichte tint (vorige prognose) naar de volle typekleur (huidige planning). Leeg laten als er "
                                                     "geen vorige prognose is: dan staan er alleen blokjes in de volle kleur.")
    put(ws, f"B{LY.WT_R_TERMIJNTITEL}", "2 · TERMIJNEN PER TYPE · % van de koopsom · kw = bouwkwartaal waarin de termijn vervalt "
                                        "(1 = kwartaal van start bouw van dat type) · naam vrij", f=F_SECTION)
    ws[f"B{LY.WT_R_TERMIJNKOP}"].comment = _comment(
        "Bouwtermijnen: per type een eigen lijst. Naam vrij (bijvoorbeeld 'Casco gereed'), percentage van de koopsom en het "
        "bouwkwartaal waarin de termijn vervalt, geteld vanaf de start bouw van dat type. Niet van toepassing: leeg laten.")
    ws[f"B{LY.WT_R_GROND}"].comment = _comment(
        "Eerste termijn van de koopsom: de grondtermijn. Komt binnen in de periode van notarieel transport en schuift dus mee "
        "als je met het verkooptempo speelt.")

    dv_kw4 = DataValidation(type="whole", operator="between", formula1="1", formula2="4", allow_blank=True, error="Kwartaal 1 t/m 4.")
    dv_kw = DataValidation(type="whole", operator="between", formula1="1", formula2="60", allow_blank=True, error="Bouwkwartaal: heel getal 1 t/m 60.")
    dv_kwx = DataValidation(type="whole", operator="between", formula1="0", formula2="60", allow_blank=True,
                            error="Bouwkwartaal: heel getal 1 t/m 60; leeg of 0 = bij notarieel transport.")
    dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True, error="Percentage tussen 0% en 100%.")
    for dv in (dv_kw4, dv_kw, dv_kwx, dv_pct):
        dv.showErrorMessage = True
        dv.errorTitle = "Ongeldige invoer"
        ws.add_data_validation(dv)

    totaal_cellen = []
    for k in range(1, N_TYPES + 1):
        c0, c1, c2 = (L(LY.wt_col(k, o)) for o in range(3))
        t = data.type_(k)
        # kop: positioneel ("TYPE 1" volgt de plek van het blok, ook na invoegen of verwijderen van kolommen)
        put(ws, f"{c0}{LY.WT_R_HEADER}", f'="TYPE "&COLUMN()/{LY.WT_W}', f=F_HDR, fl=FL_HDR, al=AL_CENTER_CONT)
        for c in (c1, c2):
            style(ws, f"{c}{LY.WT_R_HEADER}", f=F_HDR, fl=FL_HDR, al=AL_CENTER_CONT)
        put(ws, f"{c0}{LY.WT_R_NAAM}", t.naam or None, f=F_INPUT, fl=FL_INPUT, al=AL_LEFT_TOP)
        put(ws, f"{c0}{LY.WT_R_AANTAL}", t.aantal, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT, al=AL_LEFT_TOP)
        put(ws, f"{c0}{LY.WT_R_KOOPSOM}", t.koopsom, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT, al=AL_LEFT_TOP)
        put(ws, f"{c0}{LY.WT_R_STARTJAAR}", t.start_jaar, f=F_INPUT, fl=FL_INPUT, nf="0", al=AL_LEFT_TOP)
        put(ws, f"{c0}{LY.WT_R_STARTKW}", t.start_kw, f=F_INPUT, fl=FL_INPUT, nf="0", al=AL_LEFT_TOP)
        dv_kw4.add(f"{c0}{LY.WT_R_STARTKW}")
        put(ws, f"{c0}{LY.WT_R_STARTTEKST}",
            f'=IF(OR(N({c0}{LY.WT_R_AANTAL})=0,N({c0}{LY.WT_R_STARTJAAR})=0),"–","Q"&N({c0}{LY.WT_R_STARTKW})&" "&N({c0}{LY.WT_R_STARTJAAR}))',
            f=F_NOTE, al=AL_LEFT_TOP)
        # termijnen
        put(ws, f"{c0}{LY.WT_R_TERMIJNKOP}", "termijn", f=F_HDR, fl=FL_HDR, al=AL_LEFT)
        put(ws, f"{c1}{LY.WT_R_TERMIJNKOP}", "%", f=F_HDR, fl=FL_HDR, al=AL_RIGHT)
        put(ws, f"{c2}{LY.WT_R_TERMIJNKOP}", "kw", f=F_HDR, fl=FL_HDR, al=AL_RIGHT)
        put(ws, f"{c0}{LY.WT_R_GROND}", "Grondtermijn (bij transport)", f=F_NOTE)
        put(ws, f"{c1}{LY.WT_R_GROND}", t.grond_pct, f=F_INPUT, fl=FL_INPUT, nf=FMT_PCT)
        put(ws, f"{c2}{LY.WT_R_GROND}", "transport", f=F_NOTE8, al=Alignment(horizontal="right"))
        dv_pct.add(f"{c1}{LY.WT_R_GROND}")
        for i in range(LY.N_TERMIJNEN):
            r = LY.WT_R_T1 + i
            if i < len(t.termijnen):
                naam, pct, kw = t.termijnen[i]
            else:
                naam, pct, kw = (data.termijn_namen[i] if i < len(data.termijn_namen) else None), None, None
            put(ws, f"{c0}{r}", naam, f=F_INPUT, fl=FL_INPUT)
            put(ws, f"{c1}{r}", pct, f=F_INPUT, fl=FL_INPUT, nf=FMT_PCT)
            put(ws, f"{c2}{r}", kw, f=F_INPUT, fl=FL_INPUT, nf="0")
            dv_pct.add(f"{c1}{r}")
            dv_kw.add(f"{c2}{r}")
        put(ws, f"{c1}{LY.WT_R_TOTAAL}", f"=SUM({c1}{LY.WT_R_T1}:{c1}{LY.WT_R_TN})", f=F_BOLD9, nf=FMT_PCT, al=AL_RIGHT)
        # blok 3: extra opbrengsten per woning (€)
        put(ws, f"{c0}{LY.WT_R_X_KOP}", "omschrijving", f=F_HDR, fl=FL_HDR, al=AL_LEFT)
        put(ws, f"{c1}{LY.WT_R_X_KOP}", "€/won", f=F_HDR, fl=FL_HDR, al=AL_RIGHT)
        put(ws, f"{c2}{LY.WT_R_X_KOP}", "kw", f=F_HDR, fl=FL_HDR, al=AL_RIGHT)
        for i in range(LY.N_EXTRA):
            r = LY.WT_R_X1 + i
            if i < len(t.extras):
                xn, eur, kw = t.extras[i]
            else:
                xn, eur, kw = (data.extra_namen[i] if i < len(data.extra_namen) else None) or None, None, None
            put(ws, f"{c0}{r}", xn, f=F_INPUT, fl=FL_INPUT)
            put(ws, f"{c1}{r}", eur, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT)
            put(ws, f"{c2}{r}", kw, f=F_INPUT, fl=FL_INPUT, nf="0")
            dv_kwx.add(f"{c2}{r}")
        put(ws, f"{c1}{LY.WT_R_X_TOTAAL}", f"=SUM({c1}{LY.WT_R_X1}:{c1}{LY.WT_R_XN})", f=F_BOLD9, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{c1}{LY.WT_R_X_TOTAAL2}", f"=N({c0}{LY.WT_R_KOOPSOM})+{c1}{LY.WT_R_X_TOTAAL}", f=F_BOLD9, nf=FMT_INT, al=AL_RIGHT)
        # blok 4: vorige prognose
        put(ws, f"{c0}{LY.WT_R_VSTARTJAAR}", t.vorig_start_jaar, f=F_INPUT, fl=FL_INPUT, nf="0", al=AL_LEFT_TOP)
        put(ws, f"{c0}{LY.WT_R_VSTARTKW}", t.vorig_start_kw, f=F_INPUT, fl=FL_INPUT, nf="0", al=AL_LEFT_TOP)
        dv_kw4.add(f"{c0}{LY.WT_R_VSTARTKW}")
        put(ws, f"{c0}{LY.WT_R_VSTARTTEKST}",
            f'=IF(OR(N({c0}{LY.WT_R_VSTARTJAAR})=0,N({c0}{LY.WT_R_VSTARTKW})=0),"–","Q"&N({c0}{LY.WT_R_VSTARTKW})&" "&N({c0}{LY.WT_R_VSTARTJAAR}))',
            f=F_NOTE, al=AL_LEFT_TOP)
        put(ws, f"{c2}{LY.WT_R_TOTAAL}", f'=IF(COUNT({c2}{LY.WT_R_T1}:{c2}{LY.WT_R_TN})=0,"",MAX({c2}{LY.WT_R_T1}:{c2}{LY.WT_R_TN}))',
            f=F_BOLD9, nf="0", al=AL_RIGHT)
        totaal_cellen.append(f"{c1}{LY.WT_R_TOTAAL}")
    # rood als de bouwtermijnen niet optellen tot 100% (van de aanneemsom) of tot 100% − grondtermijn (van de koopsom); per blok, relatief
    c0, c1 = L(LY.wt_col(1, 0)), L(LY.wt_col(1, 1))
    ws.conditional_formatting.add(
        " ".join(totaal_cellen),
        FormulaRule(formula=[f"AND(N({c0}{LY.WT_R_AANTAL})>0,{c1}{LY.WT_R_TOTAAL}<>0,ABS({c1}{LY.WT_R_TOTAAL}-1)>=0.00005,"
                             f"ABS({c1}{LY.WT_R_TOTAAL}-(1-N({c1}{LY.WT_R_GROND})))>=0.00005)"],
                    font=Font(name=ARIAL, bold=True, color=RED)))
    put(ws, f"B{LY.WT_R_TOTAAL + 2}", "Bouwtermijnen vul je in zoals ze bij jullie staan: als % van de aanneemsom (samen 100%) of als % van de koopsom "
                                      "(samen 100% − grondtermijn). Het model schaalt ze zelf naar 100% − grondtermijn, zodat grondtermijn + bouwtermijnen "
                                      "altijd de hele koopsom vormen; terugrekenen hoeft niet.", f=F_NOTE8)
    put(ws, f"B{LY.WT_R_TOTAAL + 3}", "Termijnen die bij transport al vervallen zijn, komen in het transportkwartaal in één keer binnen, "
                                      "bovenop de grondtermijn. Een type zonder start bouw telt alleen de grondtermijn mee.", f=F_NOTE8)
    return ws


# =============================================================================
#  INVOER
# =============================================================================
def bouw_invoer(wb, data):
    ws = wb.create_sheet("Invoer")
    ws.sheet_properties.codeName = "shInvoer"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D8"
    for col, w in zip("ABCDEFGHIJK", (2, 7, 6, 13, 9, 13, 9, 11, 11, 12, 3)):
        ws.column_dimensions[col].width = w
    for k in range(1, N_TYPES + 1):
        ws.column_dimensions[L(LY.in_col(k, 0))].width = 11
        ws.column_dimensions[L(LY.in_col(k, 1))].width = 11
    ws.row_dimensions[7].height = 30

    put(ws, "B1", "Invoer — cashflow uit je eigen bestand, verkoop en transport per periode", f=F_TITLE)
    put(ws, "B2", "Blauw op lichtblauw = invoer. Grijs = wordt niet gebruikt, mag gewoon meegeplakt worden. "
                  "Woningtypes, koopsommen en termijnen staan op tab Woningtypes.", f=F_NOTE)
    put(ws, "B3", "Kopieer in tab '4. Cashflow' het blok B4:J(laatste rij) en plak hier in B8 met Plakken speciaal → Waarden. "
                  "Kosten en omzet in euro's, 'CF vorig kwartaal' in x € 1.000.", f=F_NOTE)
    put(ws, "B4", "Je eigen cashflowbestand blijft ongewijzigd: je kopieert er alleen cijfers uit. Rijen (kwartalen) mag je hier gewoon "
                  "verwijderen of invoegen; lege rijen tellen niet mee. Maximaal 60 kwartalen.", f=F_NOTE)
    put(ws, "B5", "1 · CASHFLOW UIT JE EIGEN BESTAND", f=F_SECTION)
    # titel in de smalle kolom K: loopt over de paren heen en schuift niet mee als er kolommen bij L worden ingevoegd
    put(ws, f"{L(LY.IN_COL1 - 1)}5", "2 · VERKOCHT EN GETRANSPORTEERD · aantal woningen per periode, per woningtype · "
                                     "verkocht naast notarieel transport · leeg = 0", f=F_SECTION)
    ws[f"{L(LY.IN_COL1 - 1)}5"].comment = _comment(
        "Per type twee kolommen naast elkaar: verkocht en notarieel getransporteerd. Bij transport komt de grondtermijn binnen. "
        "Type toevoegen of verwijderen: twee kolommen invoegen of verwijderen (zie tab Woningtypes).")

    koppen = ["Jaar", "Q", "Kosten", "% Kosten", "Omzet", "% Omzet", "CF", "CF (x 1000)", "CF vorig kwartaal"]
    grijs = {"E", "G", "H", "I"}
    for i, kop in enumerate(koppen):
        col = L(2 + i)
        put(ws, f"{col}7", kop, f=F_HDR, fl=FL_HDR_GREY if col in grijs else FL_HDR,
            al=AL_LEFT_WRAP if col in ("B", "C") else AL_RIGHT_WRAP)
    ws["C7"].comment = _comment("Q1 t/m Q4 (of 1 t/m 4). Staat hier een jaartal of iets anders, dan telt de regel als een heel jaar.")
    ws["J7"].comment = _comment("Cumulatieve cashflow volgens de vorige prognose, in x € 1.000 (kolom J in je eigen bestand). "
                                "Alles nul of leeg = geen vergelijking.")
    for k in range(1, N_TYPES + 1):
        cv, ct = L(LY.in_col(k, 0)), L(LY.in_col(k, 1))
        # typenaam boven het paar: positioneel, zodat de kop meeschuift met invoegen/verwijderen van kolommen
        kk = f"((COLUMN()-{LY.IN_COL1})/{LY.IN_W}+1)"
        naam = f"INDEX(Woningtypes!$A${LY.WT_R_NAAM}:${LY.WT_RANGE_END}${LY.WT_R_NAAM},1,{LY.WT_W}*{kk})"
        put(ws, f"{cv}{LY.IN_R_TYPE}", f'=IF({naam}="","type "&{kk},{naam})', f=F_BOLD9, al=AL_CENTER_CONT)
        style(ws, f"{ct}{LY.IN_R_TYPE}", f=F_BOLD9, al=AL_CENTER_CONT)
        put(ws, f"{cv}{LY.IN_R_KOP}", "Verkocht", f=F_HDR, fl=FL_HDR, al=AL_RIGHT_WRAP)
        put(ws, f"{ct}{LY.IN_R_KOP}", "Transport", f=F_HDR, fl=FL_HDR, al=AL_RIGHT_WRAP)

    for i in range(LY.N_PERIODS):
        r = ROW1 + i
        rij = data.periodes[i] if i < len(data.periodes) else {}
        for j, naam in enumerate(["jaar", "kw", "kosten", "pct_kosten", "omzet", "pct_omzet", "cf", "cf1000", "cf_vorig"]):
            col = L(2 + j)
            v = rij.get(naam)
            if col in ("B", "C"):
                put(ws, f"{col}{r}", v, f=F_INPUT, fl=FL_INPUT, al=AL_LEFT_TOP)
            elif col in grijs:
                put(ws, f"{col}{r}", v, f=F_GREY, nf=FMT_INT)
            else:
                put(ws, f"{col}{r}", v, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT)
        for k in range(1, N_TYPES + 1):
            vv = data.verkocht[i][k - 1] if i < len(data.verkocht) and k - 1 < len(data.verkocht[i]) else None
            tv = data.transport[i][k - 1] if i < len(data.transport) and k - 1 < len(data.transport[i]) else None
            put(ws, f"{L(LY.in_col(k, 0))}{r}", vv, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT)
            put(ws, f"{L(LY.in_col(k, 1))}{r}", tv, f=F_INPUT, fl=FL_INPUT, nf=FMT_INT)
    return ws


# =============================================================================
#  MODEL
# =============================================================================
def bouw_model(wb):
    ws = wb.create_sheet("Model")
    ws.sheet_properties.codeName = "shModel"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"{FIX['idx']}{ROW1}"
    laatste = LY.M_BLOK1 + len(LY.BLOKKEN) * N_TYPES - 1
    for c in range(1, laatste + 1):
        ws.column_dimensions[L(c)].width = 11
    ws.column_dimensions[LY.M_SPACER].width = 3
    ws.column_dimensions[LY.M_HULP_LABEL].width = 36
    ws.column_dimensions[LY.M_HULP].width = 30
    ws.column_dimensions[L(LY.M_BLOK1 - 1)].width = 3
    ws.row_dimensions[7].height = 34
    NR, BRON, KW, KWT, KWV, IDX, FASE, KOS, OPB, STAND, CUM, VORIG = (FIX[k] for k in
        ("nr", "bron", "kw", "kwartaal", "kwartaal_vol", "idx", "fase", "kosten", "opbr", "stand", "cum_opbr", "vorig"))
    inv_b = LY.in_rng("B")
    inv_j = LY.in_rng("J")

    put(ws, "A1", "Model — rekenblad, niets invullen", f=F_TITLE)
    put(ws, "A2", "Rij i van dit blad is de i-de gevulde rij op tab Invoer (kolom Bronrij): lege rijen tellen niet mee en rijen invoegen of "
                  "verwijderen op Invoer kan gewoon. Elke kolom rekent met de invoer, de typeblokken (vanaf kolom CA), kolommen links ervan of de "
                  "vorige rij: geen kringverwijzingen. #N/B in grafiekkolommen is normaal. Basis = je eigen cashflow (kolom J); scenario's = basis + "
                  "effect van de knoppen; rente = negatieve stand van het vorige kwartaal × jaarrente / 4.", f=F_NOTE)
    put(ws, f"{M['model_basis']}4", "modelkolommen: opbrengsten volgens woningtypes en termijnen, scenario's, rente en de 'getoonde' standen (met of zonder rente in de basis)", f=F_NOTE8)
    put(ws, f"{M['g_voorfinanciering']}4", "grafiekkolommen (x € 1 mln)", f=F_NOTE8)
    bloklabels = {
        "vcum": "rij 5: verkocht vóór start bouw · rij 6: aantal woningen",
        "tcum": "rij 5: aantal (types met start bouw) · rij 6: koopsom per woning",
        "tup": "rij 5: schaal bouwtermijnen = (100% − grond) / som bouwtermijnen · rij 6: grondtermijn %",
        "tdown": "rij 6: getransporteerd t/m actuals",
        "verv": "rij 6: index start bouw (jaar × 4 + kwartaal; 99999 = geen)",
        "tscn": "transport cumulatief in de scenariolijn (verschuiving uit de derde knoppenkolom)",
        "gv": "grafiek: verkocht per kwartaal per type (rij 7 = typenaam)",
        "gt": "grafiek: getransporteerd per kwartaal per type",
        "vx": "rij 6: extra opbrengsten per woning bij transport (€) · rijen: extra's vervallen per bouwkwartaal (€ per woning)",
    }
    for blok, t in bloklabels.items():
        put(ws, f"{m_col(blok, 1)}4", t, f=font(8, True, TXT2))

    koppen = {
        NR: "Nr", BRON: "Bronrij (Invoer)", KW: "Kw", KWT: "Kwartaal", KWV: "Kwartaal volledig", IDX: "Idx", FASE: "Fase",
        KOS: "Kosten", OPB: "Opbrengsten", STAND: "Stand", CUM: "Cum opbrengsten", VORIG: "Vorige prognose",
        M["model_basis"]: "Model basis", M["model_up"]: "Model upside", M["model_down"]: "Model downside",
        M["cum_opbr_up"]: "Cum opbr upside", M["cum_opbr_down"]: "Cum opbr downside",
        M["opbr_up"]: "Opbrengsten upside", M["opbr_down"]: "Opbrengsten downside",
        M["kosten_up"]: "Kosten upside", M["kosten_down"]: "Kosten downside",
        M["rente_basis"]: "Rente basis (berekend)", M["stand_basis_rente"]: "Stand basis incl. rente",
        M["rente_up"]: "Rente upside", M["stand_up_raw"]: "Stand upside (berekend)",
        M["rente_down"]: "Rente downside", M["stand_down_raw"]: "Stand downside (berekend)",
        M["stand_basis"]: "Stand basis (getoond)", M["kosten_basis"]: "Kosten basis (getoond)",
        M["stand_up"]: "Stand upside (getoond)", M["stand_down"]: "Stand downside (getoond)",
        M["pos_basis"]: "Positief", M["pos_up"]: "Positief upside", M["pos_down"]: "Positief downside", M["pos_vorig"]: "Positief vorig",
        M["model_scn"]: "Model scenariolijn", M["cum_opbr_scn"]: "Cum opbr scenariolijn", M["opbr_scn"]: "Opbrengsten scenariolijn",
        M["kosten_scn"]: "Kosten scenariolijn", M["rente_scn"]: "Rente scenariolijn", M["stand_scn_raw"]: "Stand scenariolijn (berekend)",
        M["stand_scn"]: "Stand scenariolijn (getoond)", M["pos_scn"]: "Positief scenariolijn",
        M["g_voorfinanciering"]: "Voorfinanciering", M["g_positief_saldo"]: "Positief saldo", M["g_opbrengsten"]: "Opbrengsten",
        M["g_kosten"]: "Kosten", M["g_stand"]: "Stand", M["g_vorige"]: "Vorige", M["g_band_onder"]: "Band onder",
        M["g_bandbreedte"]: "Bandbreedte", M["g_downside"]: "Downside", M["g_upside"]: "Upside", M["g_punt_nu"]: "Punt nu",
        M["g_punt_dal"]: "Punt dal", M["g_eind_upside"]: "Eind upside", M["g_eind_downside"]: "Eind downside",
        M["g_eind_basis"]: "Eind basis", M["g_verkocht"]: "Verkocht", M["g_getransporteerd"]: "Getransporteerd",
        M["g_verkocht_cum"]: "Verkocht cum", M["g_getransporteerd_cum"]: "Getransporteerd cum",
        M["g_verkocht_pct"]: "Verkocht pct", M["g_getransporteerd_pct"]: "Getransporteerd pct",
        M["g_opbrengsten_pct"]: "Opbrengsten pct", M["g_kosten_pct"]: "Kosten pct",
        M["g_scenario"]: "Scenariolijn", M["g_eind_scenario"]: "Eind scenariolijn", M["g_realisatie"]: "Realisatie (waas)",
        M["g_punt_uitverkocht"]: "Punt uitverkocht", M["g_punt_alles_transport"]: "Punt laatste transport",
        M["stap_ok"]: "Stap ok (controle)",
    }
    for k in range(1, N_TYPES + 1):
        koppen[m_col("vcum", k)] = f"Verkocht cum {k}"
        koppen[m_col("tcum", k)] = f"Transport cum {k}"
        koppen[m_col("tup", k)] = f"Transport cum upside {k}"
        koppen[m_col("tdown", k)] = f"Transport cum downside {k}"
        koppen[m_col("verv", k)] = f"Bouwtermijnen vervallen {k}"
        koppen[m_col("tscn", k)] = f"Transport cum scenariolijn {k}"
        koppen[m_col("vx", k)] = f"Extra's vervallen {k} (€/won)"
    for col, t in koppen.items():
        put(ws, f"{col}7", t, f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    for k in range(1, N_TYPES + 1):   # typenaam als reeksnaam van de verkoopgrafiek (vaste cel, schuift niet mee met Woningtypes)
        naam = LY.wt_ref(LY.WT_R_NAAM, k)
        for blok in ("gv", "gt"):
            put(ws, f"{m_col(blok, k)}7", f'=IF({naam}="","Type {k}",{naam})', f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    put(ws, f"{LY.M_HULP_LABEL}7", "Hulpcellen", f=F_BOLD9)

    # ---- rij 5 en 6: per type ------------------------------------------------
    idx_rng = f"${IDX}${ROW1}:${IDX}${ROWN}"
    for k in range(1, N_TYPES + 1):
        vc, tc, tu, td, bv, ts, gv, gt, vx = (m_col(b, k) for b in LY.BLOKKEN)
        aantal = f"N({LY.wt_ref(LY.WT_R_AANTAL, k)})"
        # extra opbrengsten per woning die bij transport binnenkomen (kw leeg of 0)
        x_eur = LY.wt_ref(0, k, 1, rows=(LY.WT_R_X1, LY.WT_R_XN))
        x_kw = LY.wt_ref(0, k, 2, rows=(LY.WT_R_X1, LY.WT_R_XN))
        put(ws, f"{vx}6", f'=SUMPRODUCT((N(+{x_kw})<1)*N(+{x_eur}))', f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{vc}5", f'=IF(OR({bv}$6>=99999,COUNTIF({idx_rng},"<"&{bv}$6)=0),0,'
                          f'N(INDEX({vc}${ROW1}:{vc}${ROWN},COUNTIF({idx_rng},"<"&{bv}$6))))', f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{vc}6", f"={aantal}", f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{tc}5", f"=IF({bv}$6>=99999,0,{vc}$6)", f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{tc}6", f"=N({LY.wt_ref(LY.WT_R_KOOPSOM, k)})", f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        put(ws, f"{tu}6", f"=N({LY.wt_ref(LY.WT_R_GROND, k, 1)})", f=F_CALC, nf=FMT_PCT, al=AL_RIGHT)
        pct_k = LY.wt_ref(0, k, 1, rows=(LY.WT_R_T1, LY.WT_R_TN))
        put(ws, f"{tu}5", f'=IF(SUM({pct_k})=0,0,MAX(0,1-{tu}$6)/SUM({pct_k}))', f=F_CALC, nf="0.0000", al=AL_RIGHT)
        put(ws, f"{td}6", f"=IF({h('pos_actuals')}=0,0,N(INDEX({tc}${ROW1}:{tc}${ROWN},{h('pos_actuals')})))", f=F_CALC, nf=FMT_INT, al=AL_RIGHT)
        sj, sk = LY.wt_ref(LY.WT_R_STARTJAAR, k), LY.wt_ref(LY.WT_R_STARTKW, k)
        put(ws, f"{bv}6", f"=IF(AND({aantal}>0,N({sj})>0),N({sj})*4+N({sk}),99999)", f=F_CALC, nf="0", al=AL_RIGHT)

    # ---- periode-rijen ------------------------------------------------------
    up_shift, down_shift, scn_shift = dabs("shift_up"), dabs("shift_down"), dabs("shift_scn")
    for r in range(ROW1, ROWN + 1):
        p = r - 1
        i = r - ROW1 + 1
        leeg = f'${NR}{r}=""'
        jaar = LY.in_col_ref("B", r)
        kw = LY.in_col_ref("C", r)
        f = {}
        # bronrij: de i-de gevulde rij op Invoer (matrixformule); daarna is alles positioneel
        # koprij 'Jaar' is altijd de eerste gevulde cel van het bereik, dus periode i = de (i+1)-de gevulde cel;
        # geen verwijzing naar één cel (die zou #REF! worden als de gebruiker die rij verwijdert)
        f[BRON] = ArrayFormula(f"{BRON}{r}", f'=IFERROR(SMALL(IF({inv_b}<>"",ROW({inv_b})-MIN(ROW({inv_b}))+1),{i + 1}),"")')
        f[NR] = f'=IF(${BRON}{r}="","",{i})'
        kwv = f'VALUE({kw}&"")'
        f[KW] = (f'=IF({leeg},"",IF(LEFT({kw}&"",1)="Q",IFERROR(VALUE(MID({kw},2,1)),0),'
                 f'IFERROR(IF(AND({kwv}>=1,{kwv}<=4),{kwv},0),0)))')
        f[KWT] = f'=IF({leeg},"",IF({KW}{r}=0,{jaar}&"",IF({KW}{r}=1,"Q1 {APOS}"&RIGHT({jaar},2),"Q"&{KW}{r})))'
        f[KWV] = f'=IF({leeg},"",IF({KW}{r}=0,{jaar}&"","Q"&{KW}{r}&" {APOS}"&RIGHT({jaar},2)))'
        f[IDX] = f'=IF({leeg},"",IFERROR(VALUE(TRIM({jaar}&"")),0)*4+IF({KW}{r}=0,4,{KW}{r}))'
        f[FASE] = f'=IF({leeg},"",IF({IDX}{r}<={h("idx_actuals")},"Realisatie","Prognose"))'
        f[KOS] = f'=IF({leeg},"",IFERROR(N({LY.in_col_ref("D", r)}),0))'
        f[OPB] = f'=IF({leeg},"",IFERROR(N({LY.in_col_ref("F", r)}),0))'
        f[STAND] = f'=IF({leeg},"",N({STAND}{p})+{OPB}{r}-{KOS}{r})'
        f[CUM] = f'=IF({leeg},"",N({CUM}{p})+{OPB}{r})'
        vorig = LY.in_col_ref("J", r)
        f[VORIG] = f'=IF(OR({leeg},{h("vorig_aanwezig")}=0),NA(),IF(ISNUMBER({vorig}),{vorig}*1000,NA()))'
        # typeblokken
        for k in range(1, N_TYPES + 1):
            vc, tc, tu, td, bv, ts, gv, gt, vx = (m_col(b, k) for b in LY.BLOKKEN)
            f[vc] = f'=IF({leeg},"",N({vc}{p})+IFERROR(N({LY.in_ref(r, k, 0)}),0))'
            f[tc] = f'=IF({leeg},"",N({tc}{p})+IFERROR(N({LY.in_ref(r, k, 1)}),0))'
            f[gv] = f'=IF(OR({leeg},{vc}$6=0),NA(),{vc}{r}-N({vc}{p}))'
            f[gt] = f'=IF(OR({leeg},{vc}$6=0),NA(),{tc}{r}-N({tc}{p}))'
            for col, shift in ((tu, up_shift), (td, down_shift), (ts, scn_shift)):
                f[col] = (f'=IF({leeg},"",IF(OR({FASE}{r}="Realisatie",${NR}{r}={h("n")}),{tc}{r},'
                          f'MAX({td}$6,N(INDEX({tc}${ROW1}:{tc}${ROWN},MAX(1,MIN({h("n")},${NR}{r}-N({shift}))))))))')
            kw_rng = LY.wt_ref(0, k, 2, rows=(LY.WT_R_T1, LY.WT_R_TN))
            pct_rng = LY.wt_ref(0, k, 1, rows=(LY.WT_R_T1, LY.WT_R_TN))
            # bouwtermijnen geschaald naar 100% − grondtermijn (rij 5 van blok tup), zodat termijnen als % van de aanneemsom mogen
            f[bv] = f'=IF({leeg},"",IF({IDX}{r}<{bv}$6,0,SUMIF({kw_rng},"<="&({IDX}{r}-{bv}$6+1),{pct_rng})*{tu}$5))'
            # extra's (€ per woning) met een bouwkwartaal >= 1: vervallen zoals termijnen; bij transport (rij 6) telt het model apart
            x_kw = LY.wt_ref(0, k, 2, rows=(LY.WT_R_X1, LY.WT_R_XN))
            x_eur = LY.wt_ref(0, k, 1, rows=(LY.WT_R_X1, LY.WT_R_XN))
            f[vx] = f'=IF({leeg},"",IF({IDX}{r}<{bv}$6,0,SUMIFS({x_eur},{x_kw},">=1",{x_kw},"<="&({IDX}{r}-{bv}$6+1))))'
        tc1, tcN = m_col("tcum", 1), m_col("tcum", N_TYPES)
        koopsom = m_range_abs("tcum", 6)
        grond = m_range_abs("tup", 6)
        verv = f"{m_col('verv', 1)}{r}:{m_col('verv', N_TYPES)}{r}"
        x_transport = m_range_abs("vx", 6)
        x_verv = f"{m_col('vx', 1)}{r}:{m_col('vx', N_TYPES)}{r}"
        # opbrengst per getransporteerde woning = koopsom × (grond + vervallen termijnen) [× koopsomknop per scenario]
        #                                        + extra's bij transport + extra's vervallen per bouwkwartaal
        for naam, blok, ks in (("model_basis", "tcum", None), ("model_up", "tup", "koopsom_up"), ("model_down", "tdown", "koopsom_down"),
                               ("model_scn", "tscn", "koopsom_scn")):
            rng = f"{m_col(blok, 1)}{r}:{m_col(blok, N_TYPES)}{r}"
            factor = f"*(1+{h(ks)})" if ks else ""
            f[M[naam]] = f'=IF({leeg},"",SUMPRODUCT({rng}*({koopsom}*({grond}+{verv}){factor}+{x_transport}+{x_verv})))'
        cum_rng = f"${CUM}${ROW1}:${CUM}${ROWN}"
        for naam, model, shift in (("cum_opbr_up", "model_up", up_shift), ("cum_opbr_down", "model_down", down_shift),
                                   ("cum_opbr_scn", "model_scn", scn_shift)):
            f[M[naam]] = (f'=IF({leeg},"",IF({h("model_aan")}=1,{CUM}{r}+{M[model]}{r}-{M["model_basis"]}{r},'
                          f'IF(OR({FASE}{r}="Realisatie",${NR}{r}={h("n")}),{CUM}{r},MAX({h("cum_opbr_actuals")},'
                          f'N(INDEX({cum_rng},MAX(1,MIN({h("n")},${NR}{r}-N({shift})))))))))')
        f[M["opbr_up"]] = f'=IF({leeg},"",IF({FASE}{r}="Realisatie",{OPB}{r},({M["cum_opbr_up"]}{r}-N({M["cum_opbr_up"]}{p}))*(1+N({dabs("opbr_up")}))))'
        f[M["opbr_down"]] = f'=IF({leeg},"",IF({FASE}{r}="Realisatie",{OPB}{r},({M["cum_opbr_down"]}{r}-N({M["cum_opbr_down"]}{p}))*(1+N({dabs("opbr_down")}))))'
        f[M["kosten_up"]] = f'=IF({leeg},"",IF({FASE}{r}="Prognose",{KOS}{r}*(1+N({dabs("kosten_up")})),{KOS}{r}))'
        f[M["kosten_down"]] = f'=IF({leeg},"",IF({FASE}{r}="Prognose",{KOS}{r}*(1+N({dabs("kosten_down")})),{KOS}{r}))'
        f[M["opbr_scn"]] = f'=IF({leeg},"",IF({FASE}{r}="Realisatie",{OPB}{r},({M["cum_opbr_scn"]}{r}-N({M["cum_opbr_scn"]}{p}))*(1+N({dabs("opbr_scn")}))))'
        f[M["kosten_scn"]] = f'=IF({leeg},"",IF({FASE}{r}="Prognose",{KOS}{r}*(1+N({dabs("kosten_scn")})),{KOS}{r}))'

        # rente: over de negatieve stand van het vorige kwartaal, alleen in prognosekwartalen, t/m de rente-eindindex
        def rente(stand_col, eind, pct):
            return (f'=IF({leeg},"",IF(AND({FASE}{r}="Prognose",{pct}>0,{IDX}{r}<{eind}),'
                    f'MAX(0,-N({stand_col}{p}))*{pct}/4,0))')
        f[M["rente_basis"]] = rente(M["stand_basis_rente"], h("rente_eind_basis"), h("rente"))
        f[M["stand_basis_rente"]] = f'=IF({leeg},"",N({M["stand_basis_rente"]}{p})+{OPB}{r}-{KOS}{r}-{M["rente_basis"]}{r})'
        f[M["rente_up"]] = rente(M["stand_up_raw"], h("rente_eind_up"), h("rente_pct_up"))
        f[M["stand_up_raw"]] = f'=IF({leeg},"",N({M["stand_up_raw"]}{p})+{M["opbr_up"]}{r}-{M["kosten_up"]}{r}-{M["rente_up"]}{r})'
        f[M["rente_down"]] = rente(M["stand_down_raw"], h("rente_eind_down"), h("rente_pct_down"))
        f[M["stand_down_raw"]] = f'=IF({leeg},"",N({M["stand_down_raw"]}{p})+{M["opbr_down"]}{r}-{M["kosten_down"]}{r}-{M["rente_down"]}{r})'
        f[M["rente_scn"]] = rente(M["stand_scn_raw"], h("rente_eind_scn"), h("rente_pct_scn"))
        f[M["stand_scn_raw"]] = f'=IF({leeg},"",N({M["stand_scn_raw"]}{p})+{M["opbr_scn"]}{r}-{M["kosten_scn"]}{r}-{M["rente_scn"]}{r})'
        # getoond: met rente in de basis (ja) of alleen het verschil in rente in de scenario's (nee)
        rb = h("rente_in_basis")
        f[M["stand_basis"]] = f'=IF({leeg},"",IF({rb}=1,{M["stand_basis_rente"]}{r},{STAND}{r}))'
        f[M["kosten_basis"]] = f'=IF({leeg},"",{KOS}{r}+IF({rb}=1,{M["rente_basis"]}{r},0))'
        f[M["stand_up"]] = f'=IF({leeg},"",IF({rb}=1,{M["stand_up_raw"]}{r},{M["stand_up_raw"]}{r}+({STAND}{r}-{M["stand_basis_rente"]}{r})))'
        f[M["stand_down"]] = f'=IF({leeg},"",IF({rb}=1,{M["stand_down_raw"]}{r},{M["stand_down_raw"]}{r}+({STAND}{r}-{M["stand_basis_rente"]}{r})))'
        f[M["stand_scn"]] = f'=IF({leeg},"",IF({rb}=1,{M["stand_scn_raw"]}{r},{M["stand_scn_raw"]}{r}+({STAND}{r}-{M["stand_basis_rente"]}{r})))'
        for naam, bron in (("pos_basis", M["stand_basis"]), ("pos_up", M["stand_up"]), ("pos_down", M["stand_down"]), ("pos_vorig", VORIG),
                           ("pos_scn", M["stand_scn"])):
            f[M[naam]] = f'=IF(ISNUMBER({bron}{r}),IF({bron}{r}>0,1,0),0)'
        sb, su, sd, kb = M["stand_basis"], M["stand_up"], M["stand_down"], M["kosten_basis"]
        na = f'=IF({leeg},NA(),'
        f[M["g_voorfinanciering"]] = f'{na}MIN({sb}{r}/1000000,0))'
        f[M["g_positief_saldo"]] = f'{na}MAX({sb}{r}/1000000,0))'
        f[M["g_opbrengsten"]] = f'{na}{OPB}{r}/1000000)'
        f[M["g_kosten"]] = f'{na}-{kb}{r}/1000000)'
        f[M["g_stand"]] = f'{na}{sb}{r}/1000000)'
        f[M["g_vorige"]] = f'{na}{VORIG}{r}/1000000)'
        f[M["g_band_onder"]] = f'{na}MIN({su}{r},{sd}{r})/1000000)'
        f[M["g_bandbreedte"]] = f'{na}ABS({su}{r}-{sd}{r})/1000000)'
        f[M["g_downside"]] = f'{na}{sd}{r}/1000000)'
        f[M["g_upside"]] = f'{na}{su}{r}/1000000)'
        f[M["g_punt_nu"]] = f'{na}IF(${NR}{r}={h("pos_actuals")},{sb}{r}/1000000,NA()))'
        f[M["g_punt_dal"]] = f'{na}IF({sb}{r}={h("dal_basis")},{sb}{r}/1000000,NA()))'
        f[M["g_eind_upside"]] = f'{na}IF(${NR}{r}={h("n")},{su}{r}/1000000,NA()))'
        f[M["g_eind_downside"]] = f'{na}IF(${NR}{r}={h("n")},{sd}{r}/1000000,NA()))'
        f[M["g_eind_basis"]] = f'{na}IF(${NR}{r}={h("n")},{sb}{r}/1000000,NA()))'
        # scenariolijn: alleen als de knop 'tonen' op ja staat (anders overal #N/B: de lijn verdwijnt)
        ss = M["stand_scn"]
        # de lijn begint bij het punt 'nu' (laatste gerealiseerde kwartaal): daarvóór valt hij samen met de basis
        f[M["g_scenario"]] = f'{na}IF(AND({h("scn_aan")}=1,OR({FASE}{r}="Prognose",${NR}{r}={h("pos_actuals")})),{ss}{r}/1000000,NA()))'
        f[M["g_eind_scenario"]] = f'{na}IF(AND({h("scn_aan")}=1,${NR}{r}={h("n")}),{ss}{r}/1000000,NA()))'
        f[M["g_realisatie"]] = f'{na}IF({FASE}{r}="Realisatie",1,NA()))'
        vcum, tcum = M["g_verkocht_cum"], M["g_getransporteerd_cum"]
        f[M["g_verkocht"]] = f'{na}{vcum}{r}-N({vcum}{p}))'
        f[M["g_getransporteerd"]] = f'{na}{tcum}{r}-N({tcum}{p}))'
        f[vcum] = f'{na}SUM({m_col("vcum", 1)}{r}:{m_col("vcum", N_TYPES)}{r}))'
        f[tcum] = f'{na}SUM({tc1}{r}:{tcN}{r}))'
        f[M["g_verkocht_pct"]] = f'{na}IF({h("woningen")}=0,0,{vcum}{r}/{h("woningen")}))'
        f[M["g_getransporteerd_pct"]] = f'{na}IF({h("woningen")}=0,0,{tcum}{r}/{h("woningen")}))'
        f[M["g_opbrengsten_pct"]] = f'{na}IF({h("opbr_totaal")}=0,0,{CUM}{r}/{h("opbr_totaal")}))'
        f[M["g_kosten_pct"]] = f'{na}IF({h("kosten_totaal")}=0,0,({CUM}{r}-{sb}{r})/{h("kosten_totaal")}))'
        f[M["g_punt_uitverkocht"]] = f'{na}IF(${NR}{r}={h("uitverkocht_pos")},{M["g_verkocht"]}{r},NA()))'
        f[M["g_punt_alles_transport"]] = f'{na}IF(${NR}{r}={h("alles_transport_pos")},{M["g_getransporteerd"]}{r},NA()))'
        if r == ROW1:
            f[M["stap_ok"]] = f'=IF({leeg},1,1)'
        else:
            f[M["stap_ok"]] = f'=IF(OR(${NR}{r}="",${NR}{p}=""),1,IF({IDX}{r}-{IDX}{p}=1+3*({KW}{r}=0),1,0))'
        for col, formule in f.items():
            put(ws, f"{col}{r}", formule, f=F_CALC, nf=FMT_INT)
        for col in (KWT, KWV, FASE):
            ws[f"{col}{r}"].alignment = AL_LEFT_TOP
        for col in range(LY.M_BLOK1 + 4 * N_TYPES, LY.M_BLOK1 + 5 * N_TYPES):
            ws[f"{L(col)}{r}"].number_format = FMT_PCT
        for naam in ("g_voorfinanciering", "g_positief_saldo", "g_opbrengsten", "g_kosten", "g_stand", "g_vorige", "g_band_onder",
                     "g_bandbreedte", "g_downside", "g_upside", "g_punt_nu", "g_punt_dal", "g_eind_upside", "g_eind_downside", "g_eind_basis",
                     "g_scenario", "g_eind_scenario"):
            ws[f"{M[naam]}{r}"].number_format = "0.00"
        for naam in ("g_verkocht_pct", "g_getransporteerd_pct", "g_opbrengsten_pct", "g_kosten_pct"):
            ws[f"{M[naam]}{r}"].number_format = FMT_PCT

    # ---- bouwtermijnentabel (tijdlijn): raster van alle type × termijn-combinaties, daarna de compacte lijst voor de grafiek:
    #      twee koprijen (jaartallen, kwartaalnummers), per bouwtermijn een baan per woningtype dat hem heeft, een lege rij
    #      tussen de termijnen. Tijd = kwartaalindex (jaar × 4 + kwartaal): één eenheid per kwartaal. ----------------------------
    BT = LY.BT
    put(ws, f"{BT['j']}4", "bouwtermijnentabel (grafiek): raster van alle type × termijn-combinaties met rang (volgorde van de termijnnaam; dezelfde "
                           "naam bij meerdere typen = één groep) en positie; vanaf 'nr' de compacte lijst: koprij jaartallen, koprij kwartaalnummers, "
                           "per termijn een baan per type (lichte tint = vorige prognose, typekleur = huidige planning, grijs = gerealiseerd) en een "
                           "lege rij tussen de termijnen; tijd in kwartaalindex", f=F_NOTE8)
    bt_koppen = {"j": "j", "k": "type", "i": "termijn", "gebruikt": "Gebruikt", "knaam": "Typenaam", "tnaam": "Termijnnaam", "idxb": "Idx huidig",
                 "idxr": "Idx vorig", "uniek": "Uniek", "rang": "Rang termijn", "pos": "Positie", "eerste": "Eerste baan",
                 "nr": "Nr (uniek)", "bron": "Bronrij", "soort": "Soort", "label": "Termijn", "kb": "Type", "ib": "Idx huidig", "ir": "Idx vorig"}
    bt_koppen.update(LY.BT_SEG_NAMEN)
    for naam, kop in bt_koppen.items():
        put(ws, f"{BT[naam]}7", kop, f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    # reeksnamen van de typereeksen, jaarreeksen en kwartaalreeksen (de Dashboard-grafiek en het PowerPoint-blok lezen deze cellen;
    # '(uit)' = de macro haalt de reeks uit de grafiek in de presentatie)
    k_rng, gebruikt_rng = (f"${BT[x]}${LY.BT_ROW1}:${BT[x]}${LY.BT_ROWN}" for x in ("k", "gebruikt"))
    for k in range(1, N_TYPES + 1):
        naam_k = LY.wt_ref(LY.WT_R_NAAM, k)
        knaam = f'IF({naam_k}="","Type {k}",{naam_k}&"")'
        aan = f"COUNTIFS({k_rng},{k},{gebruikt_rng},1)>0"
        put(ws, f"{BT[f'cur{k}']}7", f'=IF({aan},{knaam},"(uit)")', f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
        put(ws, f"{BT[f'prev{k}']}7", f'=IF({aan},{knaam}&" · vorige prognose","(uit)")', f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    for y in range(1, LY.BT_NJ + 1):
        put(ws, f"{BT[f'jr{y}']}7", f'=IF({y}<={h("bt_jaren")},TEXT({h("jaar_first")}+{y - 1},"0"),"(uit)")', f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    for q_ in range(1, LY.BT_NK + 1):
        put(ws, f"{BT[f'kw{q_}']}7", str((q_ - 1) % 4 + 1), f=F_HDR8, fl=FL_HDR, al=AL_RIGHT_WRAP)
    ws.column_dimensions[L(LY.BT_COL1 - 1)].width = 3
    ws.column_dimensions[BT["label"]].width = 30
    for naam in BT:
        if naam != "label":
            ws.column_dimensions[BT[naam]].width = 11 if not naam.startswith(("jr", "kw")) else 5
    rng_bt = lambda naam: f"${BT[naam]}${LY.BT_ROW1}:${BT[naam]}${LY.BT_ROWN}"  # noqa: E731
    uniek_rng, tnaam_rng, idxb_rng, idxr_rng = rng_bt("uniek"), rng_bt("tnaam"), rng_bt("idxb"), rng_bt("idxr")
    j_rng, rang_rng, pos_rng, nr_rng, eerste_rng = rng_bt("j"), rng_bt("rang"), rng_bt("pos"), rng_bt("nr"), rng_bt("eerste")
    tf, te, tn = h("t_first"), h("t_end"), h("t_nu_end")
    for r in range(LY.BT_ROW1, LY.BT_ROWN + 1):
        j = r - LY.BT_ROW1 + 1
        f = {}
        f["j"] = j
        # ---- raster (alleen de eerste N_TYPES × N_TERMIJNEN rijen) ----
        if j <= LY.BT_LANES:
            k = (j - 1) // LY.N_TERMIJNEN + 1
            i = (j - 1) % LY.N_TERMIJNEN + 1
            kw_ki = LY.wt_ref(LY.WT_R_T1 + i - 1, k, 2)
            pct_ki = LY.wt_ref(LY.WT_R_T1 + i - 1, k, 1)
            tn_ki = LY.wt_ref(LY.WT_R_T1 + i - 1, k, 0)
            naam_k = LY.wt_ref(LY.WT_R_NAAM, k)
            vj_k, vk_k = LY.wt_ref(LY.WT_R_VSTARTJAAR, k), LY.wt_ref(LY.WT_R_VSTARTKW, k)
            gebruikt = f"${BT['gebruikt']}{r}"
            rang = f"${BT['rang']}{r}"
            f["k"], f["i"] = k, i
            f["gebruikt"] = (f'=IF(AND(N({LY.wt_ref(LY.WT_R_AANTAL, k)})>0,{m_col("verv", k)}$6<99999,ISNUMBER({kw_ki}),N({kw_ki})>=1,'
                             f'N({pct_ki})<>0),1,0)')
            f["knaam"] = f'=IF({naam_k}="","Type {k}",{naam_k}&"")'
            f["tnaam"] = f'=IF({tn_ki}="","termijn {i}",{tn_ki}&"")'
            f["idxb"] = f'=IF({gebruikt}=1,{m_col("verv", k)}$6+N({kw_ki})-1,"")'
            f["idxr"] = f'=IF({gebruikt}=1,IF(AND(N({vj_k})>0,N({vk_k})>=1),N({vj_k})*4+N({vk_k})+N({kw_ki})-1,${BT["idxb"]}{r}),"")'
            if r == LY.BT_ROW1:
                f["uniek"] = f'={gebruikt}'
            else:
                eerder_t = f"${BT['tnaam']}${LY.BT_ROW1}:${BT['tnaam']}${r - 1}"
                eerder_g = f"${BT['gebruikt']}${LY.BT_ROW1}:${BT['gebruikt']}${r - 1}"
                f["uniek"] = f'=IF(AND({gebruikt}=1,SUMPRODUCT(--({eerder_t}=${BT["tnaam"]}{r}),--({eerder_g}=1))=0),1,0)'
            # rang van de termijnnaam = positie (in 'nr') van de eerste gebruikte rij met deze naam
            zelfde = f"({tnaam_rng}=${BT['tnaam']}{r})*({gebruikt_rng}=1)"
            f["rang"] = ArrayFormula(f"{BT['rang']}{r}", f'=IF({gebruikt}=1,MATCH(MIN(IF({zelfde},{j_rng})),{nr_rng},0),"")')
            # positie in de compacte lijst: banen van lagere rangen + scheidingsrijen + eerdere typen met deze termijn
            f["pos"] = (f'=IF({gebruikt}=1,SUMPRODUCT(({gebruikt_rng}=1)*({rang_rng}<{rang}))+{rang}-1'
                        f'+SUMPRODUCT(({gebruikt_rng}=1)*({rang_rng}={rang})*({k_rng}<{k}))+1,"")')
            f["eerste"] = f'=IF({gebruikt}=1,IF(SUMPRODUCT(({gebruikt_rng}=1)*({rang_rng}={rang})*({k_rng}<{k}))=0,1,0),"")'
        else:
            for naam_ in ("k", "i", "gebruikt", "knaam", "tnaam", "idxb", "idxr", "uniek", "rang", "pos", "eerste"):
                f[naam_] = None
        # ---- compacte lijst ----
        f["nr"] = ArrayFormula(f"{BT['nr']}{r}", f'=IFERROR(SMALL(IF({uniek_rng}=1,{j_rng}),{j}),"")')
        bron, soort = f"${BT['bron']}{r}", f"${BT['soort']}{r}"
        if j <= LY.BT_KOP:
            f["bron"] = None
            f["soort"] = j
            f["label"] = " "
            f["kb"] = f["ib"] = f["ir"] = None
        else:
            f["bron"] = f'=IFERROR(MATCH({j - LY.BT_KOP},{pos_rng},0),"")'
            f["soort"] = f'=IF({bron}<>"",3,IF({j - LY.BT_KOP}<={h("bt_lanes")}+{h("bt_uniek")}-1,4,0))'
            f["label"] = f'=IF({soort}=3,IF(INDEX({eerste_rng},{bron})=1,INDEX({tnaam_rng},{bron})," "),IF({soort}=4," ",""))'
            f["kb"] = f'=IF({soort}=3,INDEX({k_rng},{bron}),"")'
            f["ib"] = f'=IF({soort}=3,INDEX({idxb_rng},{bron}),"")'
            f["ir"] = f'=IF({soort}=3,INDEX({idxr_rng},{bron}),"")'
        kb, ib, ir = f"${BT['kb']}{r}", f"${BT['ib']}{r}", f"${BT['ir']}{r}"
        baan = f"{soort}=3"
        # blokjes van één kwartaal, afgekapt op de as [t_first, t_end]; buiten de as: breedte 0
        ibc, irc = f"MIN(MAX({ib},{tf}),{te})", f"MIN(MAX({ir},{tf}),{te})"
        blauw, rood = f"MAX(0,MIN({ib}+1,{te})-{ibc})", f"MAX(0,MIN({ir}+1,{te})-{irc})"
        s1, s2, s4 = (f"${BT[x]}{r}" for x in ("s1", "s2", "s4"))
        # stapel 1 (eerste as): onzichtbaar tot de as-ondergrens, grijs, onzichtbaar, vorige prognose (per type), grijs, onzichtbaar; jaarblokjes in koprij 1
        f["s0"] = f'=IF(OR({baan},{soort}=1),{tf},NA())'
        f["s1"] = f'=IF({baan},MAX(0,MIN({irc},{tn})-{tf}),NA())'
        f["s2"] = f'=IF({baan},MAX(0,{irc}-{tn}),NA())'
        for k in range(1, N_TYPES + 1):
            f[f"prev{k}"] = f'=IF(AND({baan},{kb}={k}),{rood},NA())'
        f["s4"] = f'=IF({baan},MAX(0,{tn}-({irc}+{rood})),NA())'
        f["s5"] = f'=IF({baan},MAX(0,{te}-({tf}+{s1}+{s2}+{rood}+{s4})),NA())'
        # stapel 2 (tweede as, bovenop): onzichtbaar, huidige planning (per type), onzichtbaar; kwartaalblokjes in koprij 2
        f["s6"] = f'=IF({baan},{ibc},IF({soort}=2,{tf},NA()))'
        for k in range(1, N_TYPES + 1):
            f[f"cur{k}"] = f'=IF(AND({baan},{kb}={k}),{blauw},NA())'
        f["s8"] = f'=IF({baan},MAX(0,{te}-{ibc}-{blauw}),NA())'
        for y in range(1, LY.BT_NJ + 1):
            f[f"jr{y}"] = f'=IF({y}<={h("bt_jaren")},4,NA())' if j == 1 else "#N/A"
        for q_ in range(1, LY.BT_NK + 1):
            f[f"kw{q_}"] = f'=IF({q_}<=4*{h("bt_jaren")},1,NA())' if j == 2 else "#N/A"
        for naam_, formule in f.items():
            if formule is None:
                continue
            nf = "0.00" if naam_ in ("s1", "s2", "s4", "s5", "s8") or naam_.startswith(("prev", "cur")) else "0"
            put(ws, f"{BT[naam_]}{r}", formule, f=F_CALC, nf=nf, al=AL_LEFT_TOP if naam_ in ("label", "knaam", "tnaam") else AL_RIGHT)
            if formule == "#N/A":
                ws[f"{BT[naam_]}{r}"].data_type = "e"

    # ---- hulpcellen -----------------------------------------------------------
    sb, su, sd, ss = M["stand_basis"], M["stand_up"], M["stand_down"], M["stand_scn"]
    pb, pu, pd, pv, ps = M["pos_basis"], M["pos_up"], M["pos_down"], M["pos_vorig"], M["pos_scn"]
    rng = lambda col: f"${col}${ROW1}:${col}${ROWN}"  # noqa: E731
    D_ = rng(KWV)
    bron_van = lambda pos: f"INDEX({rng(BRON)},{pos})"  # noqa: E731  bronrij (Invoer) van modelpositie pos
    hulp = {
        "n": ("Aantal periodes", f"=COUNT({rng(NR)})"),
        "idx_actuals": ("Index actuals t/m", f"={dabs('actuals_jaar')}*4+{dabs('actuals_kw')}"),
        "pos_actuals": ("Positie actuals", f'=COUNTIF({rng(IDX)},"<="&{h("idx_actuals")})'),
        "kw_actuals": ("Kwartaal actuals", f'=IF({h("pos_actuals")}=0,"–",INDEX({D_},{h("pos_actuals")}))'),
        "stand_nu": ("Stand nu", f'=IF({h("pos_actuals")}=0,0,INDEX({rng(sb)},{h("pos_actuals")}))'),
        "model_aan": ("Verkooptempo-model aan (1/0)", f'=IF(LOWER({dabs("model_aan")}&"")="ja",1,0)'),
        "vorig_aanwezig": ("Vorige prognose aanwezig (1/0)", f'=IF(SUMIF({inv_j},">0")-SUMIF({inv_j},"<0")>0,1,0)'),
        "opbr_totaal": ("Opbrengsten totaal", f"=SUM({rng(OPB)})"),
        "kosten_totaal": ("Kosten totaal (incl. rente als die in de basis zit)", f"=SUM({rng(M['kosten_basis'])})"),
        "cum_opbr_actuals": ("Cum. opbrengsten t/m actuals", f'=IF({h("pos_actuals")}=0,0,INDEX({rng(CUM)},{h("pos_actuals")}))'),
        "dal_basis": ("Dieptepunt basis", f"=MIN({rng(sb)})"),
        "dal_basis_kw": ("Dieptepunt basis kwartaal", f'=IFERROR(INDEX({D_},MATCH({h("dal_basis")},{rng(sb)},0)),"–")'),
        "eind_basis": ("Eindsaldo basis", f'=IF({h("n")}=0,0,INDEX({rng(sb)},{h("n")}))'),
        "be_basis": ("Break-even basis", f'=IFERROR(INDEX({D_},MATCH(1,{rng(pb)},0)),"")'),
        "be_basis_pos": ("Positie break-even basis", f'=IFERROR(MATCH(1,{rng(pb)},0),0)'),
        "dal_up": ("Dieptepunt upside", f"=MIN({rng(su)})"),
        "dal_up_kw": ("Dieptepunt upside kwartaal", f'=IFERROR(INDEX({D_},MATCH({h("dal_up")},{rng(su)},0)),"–")'),
        "eind_up": ("Eindsaldo upside", f'=IF({h("n")}=0,0,INDEX({rng(su)},{h("n")}))'),
        "be_up": ("Break-even upside", f'=IFERROR(INDEX({D_},MATCH(1,{rng(pu)},0)),"")'),
        "be_up_pos": ("Positie break-even upside", f'=IFERROR(MATCH(1,{rng(pu)},0),0)'),
        "dal_down": ("Dieptepunt downside", f"=MIN({rng(sd)})"),
        "dal_down_kw": ("Dieptepunt downside kwartaal", f'=IFERROR(INDEX({D_},MATCH({h("dal_down")},{rng(sd)},0)),"–")'),
        "eind_down": ("Eindsaldo downside", f'=IF({h("n")}=0,0,INDEX({rng(sd)},{h("n")}))'),
        "be_down": ("Break-even downside", f'=IFERROR(INDEX({D_},MATCH(1,{rng(pd)},0)),"")'),
        "be_down_pos": ("Positie break-even downside", f'=IFERROR(MATCH(1,{rng(pd)},0),0)'),
        "laatste_kw": ("Laatste kwartaal", f'=IF({h("n")}=0,"–",INDEX({D_},{h("n")}))'),
        "dal_vorig": ("Dieptepunt vorige prognose (€)", f'=IF({h("vorig_aanwezig")}=0,0,MIN({inv_j})*1000)'),
        "eind_vorig": ("Eindsaldo vorige prognose (€)", f'=IF(OR({h("vorig_aanwezig")}=0,{h("n")}=0),0,N(INDEX({inv_j},{bron_van(h("n"))}))*1000)'),
        "vorig_nu": ("Vorige prognose nu (€)", f'=IF(OR({h("vorig_aanwezig")}=0,{h("pos_actuals")}=0),0,N(INDEX({inv_j},{bron_van(h("pos_actuals"))}))*1000)'),
        "be_vorig": ("Break-even vorige prognose", f'=IFERROR(INDEX({D_},MATCH(1,{rng(pv)},0)),"")'),
        "be_vorig_pos": ("Positie break-even vorige prognose", f'=IFERROR(MATCH(1,{rng(pv)},0),0)'),
        "woningen": ("Woningen totaal", f"=SUM({m_range_abs('vcum', 6)})"),
        "verkocht_plan": ("Verkocht totaal (planning)", f'=IF({h("n")}=0,0,INDEX({rng(M["g_verkocht_cum"])},{h("n")}))'),
        "transport_plan": ("Getransporteerd totaal (planning)", f'=IF({h("n")}=0,0,INDEX({rng(M["g_getransporteerd_cum"])},{h("n")}))'),
        "verkocht_actuals": ("Verkocht t/m actuals", f'=IF({h("pos_actuals")}=0,0,INDEX({rng(M["g_verkocht_cum"])},{h("pos_actuals")}))'),
        "transport_actuals": ("Getransporteerd t/m actuals", f'=IF({h("pos_actuals")}=0,0,INDEX({rng(M["g_getransporteerd_cum"])},{h("pos_actuals")}))'),
        "model_opbr": ("Model-opbrengst totaal", f'=IF({h("n")}=0,0,INDEX({rng(M["model_basis"])},{h("n")}))'),
        "verk_voor_start": ("Verkocht vóór start bouw (aantal)", f"=SUM({m_range_abs('vcum', 5)})"),
        "won_met_start": ("Woningen in types met start bouw", f"=SUM({m_range_abs('tcum', 5)})"),
        "verk_voor_start_pct": ("Verkocht vóór start bouw (%)", f'=IFERROR({h("verk_voor_start")}/{h("won_met_start")},0)'),
        "uitverkocht": ("Uitverkocht in",
                        f'=IF(OR({h("woningen")}=0,(COUNTIF({rng(M["g_verkocht_pct"])},"<"&0.99999)+1)>{h("n")}),"niet binnen de looptijd",'
                        f'INDEX({D_},(COUNTIF({rng(M["g_verkocht_pct"])},"<"&0.99999)+1)))'),
        "alles_transport": ("Alles getransporteerd in",
                            f'=IF(OR({h("woningen")}=0,(COUNTIF({rng(M["g_getransporteerd_pct"])},"<"&0.99999)+1)>{h("n")}),"niet binnen de looptijd",'
                            f'INDEX({D_},(COUNTIF({rng(M["g_getransporteerd_pct"])},"<"&0.99999)+1)))'),
        "lbl_nu": ("Label punt nu", f'="nu · "&{h("kw_actuals")}'),
        "lbl_dal": ("Label dieptepunt", f'="Dieptepunt "&{h("dal_basis_kw")}&"  "&{eur_m(h("dal_basis"))}'),
        "lbl_eind_up": ("Label eind upside", f'="Upside "&{eur_m(h("eind_up"))}'),
        "lbl_eind_down": ("Label eind downside", f'="Downside "&{eur_m(h("eind_down"))}'),
        "lbl_eind_basis": ("Label eind basis", f'="Basis "&{eur_m(h("eind_basis"))}'),
        "rente": ("Jaarrente", f"=N({dabs('rente')})"),
        "rente_tm_start": ("Rente t/m start bouw (1/0)", f'=IF(LOWER({dabs("rente_tm")}&"")="start bouw",1,0)'),
        "rente_in_basis": ("Rente ook in de basis (1/0)", f'=IF(LOWER({dabs("rente_basis")}&"")="ja",1,0)'),
        "start_idx": ("Index start bouw project (vroegste type)", f"=IF(MIN({m_range_abs('verv', 6)})>=99999,99999,MIN({m_range_abs('verv', 6)}))"),
        "start_tekst": ("Start bouw project", f'=IF({h("start_idx")}>=99999,"–","Q"&(MOD({h("start_idx")}-1,4)+1)&" "&INT(({h("start_idx")}-1)/4))'),
        "rente_eind_basis": ("Rente-eindindex basis (rente zolang idx < deze)", f'=IF({h("rente_tm_start")}=1,{h("start_idx")},99999)'),
        "rente_eind_up": ("Rente-eindindex upside", f'=IF(OR({h("rente_tm_start")}=0,{h("start_idx")}>=99999),99999,{h("start_idx")}+N({dabs("uitstel_up")}))'),
        "rente_eind_down": ("Rente-eindindex downside", f'=IF(OR({h("rente_tm_start")}=0,{h("start_idx")}>=99999),99999,{h("start_idx")}+N({dabs("uitstel_down")}))'),
        "rente_basis_tot": ("Rente basis totaal (berekend)", f"=SUM({rng(M['rente_basis'])})"),
        "rente_up_tot": ("Rente upside totaal", f"=SUM({rng(M['rente_up'])})"),
        "rente_down_tot": ("Rente downside totaal", f"=SUM({rng(M['rente_down'])})"),
        "rente_up_extra": ("Extra rente upside t.o.v. basis", f'={h("rente_up_tot")}-{h("rente_basis_tot")}'),
        "rente_down_extra": ("Extra rente downside t.o.v. basis", f'={h("rente_down_tot")}-{h("rente_basis_tot")}'),
        "rente_tekst": ("Rente-tekst", f'=IF({h("rente")}=0,"geen rente","rente "&FIXED({h("rente")}*100,1)&"%"&IF({h("rente_tm_start")}=1," t/m start bouw",'
                                       f'" over de hele looptijd")&IF({h("rente_in_basis")}=1,", ook in de basis",""))'
                                       f'&IF(OR({h("rente_pct_down")}<>{h("rente")},{h("rente_pct_up")}<>{h("rente")},{h("rente_pct_scn")}<>{h("rente")}),'
                                       f'" · per scenario "&FIXED({h("rente_pct_down")}*100,1)&"% / "&FIXED({h("rente_pct_up")}*100,1)&"% / "'
                                       f'&FIXED({h("rente_pct_scn")}*100,1)&"%","")'),
        "aantal_types": ("Aantal woningtypes (met aantal)", f'=COUNTIF({m_range_abs("vcum", 6)},">0")'),
        # scenariolijn (gele lijn in de cashflowgrafiek)
        "scn_aan": ("Scenariolijn tonen (1/0)", f'=IF(LOWER({dabs("scn_aan")}&"")="nee",0,1)'),
        "scn_naam": ("Naam scenariolijn", f'=IF(TRIM({dabs("scn_naam")}&"")="","Scenario",TRIM({dabs("scn_naam")}&""))'),
        "eind_scn": ("Eindsaldo scenariolijn", f'=IF({h("n")}=0,0,INDEX({rng(ss)},{h("n")}))'),
        "dal_scn": ("Dieptepunt scenariolijn", f"=MIN({rng(ss)})"),
        "dal_scn_kw": ("Dieptepunt scenariolijn kwartaal", f'=IFERROR(INDEX({D_},MATCH({h("dal_scn")},{rng(ss)},0)),"–")'),
        "be_scn": ("Break-even scenariolijn", f'=IFERROR(INDEX({D_},MATCH(1,{rng(ps)},0)),"")'),
        "be_scn_pos": ("Positie break-even scenariolijn", f'=IFERROR(MATCH(1,{rng(ps)},0),0)'),
        "rente_eind_scn": ("Rente-eindindex scenariolijn", f'=IF(OR({h("rente_tm_start")}=0,{h("start_idx")}>=99999),99999,{h("start_idx")}+N({dabs("uitstel_scn")}))'),
        "rente_scn_tot": ("Rente scenariolijn totaal", f"=SUM({rng(M['rente_scn'])})"),
        "rente_scn_extra": ("Extra rente scenariolijn t.o.v. basis", f'={h("rente_scn_tot")}-{h("rente_basis_tot")}'),
        "lbl_scenario": ("Legenda scenariolijn", f'={h("scn_naam")}&IF({h("scn_aan")}=1,""," (uit)")'),
        "lbl_eind_scn": ("Label eind scenariolijn", f'={h("scn_naam")}&"  "&{eur_m(h("eind_scn"))}'),
        "lbl_nu_cf": ("Label stand nu (cashflowgrafiek)", f'="Stand "&{h("kw_actuals")}&"  "&{eur_m(h("stand_nu"))}'),
        "lbl_eind_basis_cf": ("Label eindsaldo (cashflowgrafiek)", f'="Eindsaldo "&{h("laatste_kw")}&"  "&{eur_m(h("eind_basis"))}'),
        # verkoopgrafiek
        "uitverkocht_pos": ("Positie uitverkocht (0 = niet binnen de looptijd)",
                            f'=IF(OR({h("woningen")}=0,(COUNTIF({rng(M["g_verkocht_pct"])},"<"&0.99999)+1)>{h("n")}),0,COUNTIF({rng(M["g_verkocht_pct"])},"<"&0.99999)+1)'),
        "alles_transport_pos": ("Positie alles getransporteerd (0 = niet binnen de looptijd)",
                                f'=IF(OR({h("woningen")}=0,(COUNTIF({rng(M["g_getransporteerd_pct"])},"<"&0.99999)+1)>{h("n")}),0,COUNTIF({rng(M["g_getransporteerd_pct"])},"<"&0.99999)+1)'),
        "lbl_uitverkocht": ("Label uitverkocht", f'="Uitverkocht "&{h("uitverkocht")}'),
        "lbl_alles_transport": ("Label laatste transport", f'="Laatste transport "&{h("alles_transport")}'),
        "nog_verkopen": ("Nog te verkopen (prognose)", f'=MAX(0,{h("woningen")}-{h("verkocht_actuals")})'),
        "nog_transport": ("Nog te transporteren (prognose)", f'=MAX(0,{h("woningen")}-{h("transport_actuals")})'),
        # eigen jaarrente per scenario (lege knop = de algemene jaarrente) en koopsomknop (VON-prijs) per scenario
        "rente_pct_down": ("Jaarrente downside", f'=IF({dabs("rente_pct_down")}="",{h("rente")},N({dabs("rente_pct_down")}))'),
        "rente_pct_up": ("Jaarrente upside", f'=IF({dabs("rente_pct_up")}="",{h("rente")},N({dabs("rente_pct_up")}))'),
        "rente_pct_scn": ("Jaarrente scenariolijn", f'=IF({dabs("rente_pct_scn")}="",{h("rente")},N({dabs("rente_pct_scn")}))'),
        "koopsom_down": ("Koopsom downside (%)", f'=N({dabs("koopsom_down")})'),
        "koopsom_up": ("Koopsom upside (%)", f'=N({dabs("koopsom_up")})'),
        "koopsom_scn": ("Koopsom scenariolijn (%)", f'=N({dabs("koopsom_scn")})'),
        # bouwtermijnengrafiek: aantal rijen en de tijd-as in kwartaalindex (jaar × 4 + kwartaal): van 1 januari van het eerste jaar
        # van de periodes tot het einde van het laatste jaar; de VBA leest bt_n (rijen), t_first en t_end (as-grenzen)
        "bt_lanes": ("Bouwtermijnentabel: aantal banen (type × termijn)", f"=SUM(${LY.BT['gebruikt']}${LY.BT_ROW1}:${LY.BT['gebruikt']}${LY.BT_ROWN})"),
        "bt_uniek": ("Bouwtermijnentabel: aantal unieke termijnnamen", f"=SUM(${LY.BT['uniek']}${LY.BT_ROW1}:${LY.BT['uniek']}${LY.BT_ROWN})"),
        "bt_n": ("Bouwtermijnentabel: aantal rijen (koprijen, banen, scheidingsrijen)",
                 f'=IF({h("bt_lanes")}=0,0,MIN({LY.BT_N},{LY.BT_KOP}+{h("bt_lanes")}+{h("bt_uniek")}-1))'),
        "jaar_first": ("Tijd-as: eerste jaar", f'=IF({h("n")}=0,0,INT(({IDX}{ROW1}-1)/4))'),
        "jaar_last": ("Tijd-as: laatste jaar", f'=IF({h("n")}=0,0,INT((INDEX({rng(IDX)},{h("n")})-1)/4))'),
        "bt_jaren": ("Tijd-as: aantal jaren (koprij)", f'=MAX(0,MIN({LY.BT_NJ},{h("jaar_last")}-{h("jaar_first")}+1))'),
        "t_first": ("Tijd-as: ondergrens (kwartaalindex 1 jan eerste jaar)", f'={h("jaar_first")}*4+1'),
        "t_end": ("Tijd-as: bovengrens (kwartaalindex na het laatste jaar)", f'=({h("jaar_last")}+1)*4+1'),
        "t_nu_end": ("Tijd einde laatste gerealiseerde kwartaal (kwartaalindex)",
                     f'=IF({h("pos_actuals")}=0,{h("t_first")},MIN(MAX({h("idx_actuals")}+1,{h("t_first")}),{h("t_end")}))'),
        "lbl_realisatie": ("Legenda gerealiseerd", f'="Gerealiseerd t/m "&{h("kw_actuals")}'),
        "bt_types": ("Aantal typen met bouwtermijnen (aantal > 0 en start bouw)", f"=SUMPRODUCT(--({m_range_abs('verv', 6)}<99999),--({m_range_abs('vcum', 6)}>0))"),
    }
    for naam, (label, formule) in hulp.items():
        r = H[naam]
        put(ws, f"{LY.M_HULP_LABEL}{r}", label, f=F_NOTE8)
        put(ws, f"{LY.M_HULP}{r}", formule, f=F_CALC, al=AL_LEFT_TOP)
    for naam in ("verk_voor_start_pct", "rente", "rente_pct_down", "rente_pct_up", "rente_pct_scn", "koopsom_down", "koopsom_up", "koopsom_scn"):
        ws[f"{LY.M_HULP}{H[naam]}"].number_format = FMT_PCT
    for naam in ("stand_nu", "opbr_totaal", "kosten_totaal", "cum_opbr_actuals", "dal_basis", "eind_basis", "dal_up", "eind_up",
                 "dal_down", "eind_down", "dal_vorig", "eind_vorig", "vorig_nu", "model_opbr", "rente_basis_tot", "rente_up_tot",
                 "rente_down_tot", "rente_up_extra", "rente_down_extra", "eind_scn", "dal_scn", "rente_scn_tot", "rente_scn_extra"):
        ws[f"{LY.M_HULP}{H[naam]}"].number_format = FMT_INT
    return ws


# =============================================================================
#  DASHBOARD
# =============================================================================
def bouw_dashboard(wb, data):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_properties.codeName = "shDashboard"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    for col, w in (("A", 2), ("B", 9), ("F", 2), ("G", 9), ("K", 2), ("L", 9), ("P", 2), ("Q", 9), ("V", 36), ("W", 12), ("X", 12), ("Y", 12)):
        ws.column_dimensions[col].width = w
    for r, hgt in ((1, 15.6), (2, 25.95), (3, 20), (5, 16.05), (6, 28.05), (7, 13.95), (8, 15), (10, 16.05), (11, 24), (12, 13.95), (13, 15)):
        ws.row_dimensions[r].height = hgt
    P = data.params
    rb = hm("rente_in_basis")

    # ---- kopregels ---------------------------------------------------------------
    put(ws, "B1", f'="CASHFLOW · STAND PER "&UPPER({hm("kw_actuals")})', f=F_NOTE)
    ws.merge_cells("B2:T2")
    put(ws, "B2", f'="Financiering piekt op €"&FIXED(ABS({hm("dal_basis")})/1000000,1)&" mln in "&{hm("dal_basis_kw")}'
                  f'&IF({hm("be_basis")}="",", geen break-even",", break-even in "&{hm("be_basis")})', f=font(14, True, TXT), al=AL_VCENTER)
    ws.merge_cells("B3:T3")
    put(ws, "B3", f'="Projectresultaat "&{eur_mln(hm("eind_basis"))}&" · marge "&IFERROR(FIXED({hm("eind_basis")}/{hm("opbr_totaal")}*100,1),"0")'
                  f'&"% · opbrengsten €"&FIXED({hm("opbr_totaal")}/1000000,1)&" mln · kosten €"&FIXED({hm("kosten_totaal")}/1000000,1)&" mln · basis = je eigen cashflow"'
                  f'&IF({rb}=1," + rente ("&FIXED({hm("rente_basis_tot")}/1000000,1)&" mln)","")', f=font(10, False, TXT2))

    # ---- KPI-kaarten -------------------------------------------------------------
    def kaart(col0, r0, titel, waarde, sub, verschil, nf=FMT_MLN, kleur=TXT, size=18):
        c0, c3 = L(col0), L(col0 + 3)
        for r in (r0, r0 + 1, r0 + 2, r0 + 3):
            ws.merge_cells(f"{c0}{r}:{c3}{r}")
        put(ws, f"{c0}{r0}", titel, f=F_NOTE, fl=FL_CARD, al=AL_LEFT,
            bd=Border(left=side(), right=side(), top=side()))
        for c in range(col0 + 1, col0 + 3):
            style(ws, f"{L(c)}{r0}", bd=Border(top=side()))
        style(ws, f"{c3}{r0}", bd=Border(top=side(), right=side()))
        put(ws, f"{c0}{r0 + 1}", waarde, f=font(size, True, kleur), fl=FL_CARD, nf=nf, al=AL_LEFT, bd=Border(left=side(), right=side()))
        style(ws, f"{c3}{r0 + 1}", bd=Border(right=side()))
        put(ws, f"{c0}{r0 + 2}", sub, f=font(9, False, TXT2), fl=FL_CARD, al=AL_LEFT, bd=Border(left=side(), right=side()))
        style(ws, f"{c3}{r0 + 2}", bd=Border(right=side()))
        put(ws, f"{c0}{r0 + 3}", verschil, f=F_NOTE, fl=FL_CARD, al=AL_LEFT,
            bd=Border(left=side(), right=side(), bottom=side("thin", LINE2)))
        for c in range(col0 + 1, col0 + 3):
            style(ws, f"{L(c)}{r0 + 3}", bd=Border(bottom=side("thin", LINE2)))
        style(ws, f"{c3}{r0 + 3}", bd=Border(right=side(), bottom=side("thin", LINE2)))

    dal, dal_v, eind, eind_v = hm("dal_basis"), hm("dal_vorig"), hm("eind_basis"), hm("eind_vorig")
    kaart(2, 5, "MAX. FINANCIERINGSBEHOEFTE", f"={dal}/1000000", f'="dieptepunt in "&{hm("dal_basis_kw")}',
          f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose","was "&{eur_m(dal_v)}&IF(ABS(({dal}-{dal_v}))<50000," · gelijk",'
          f'" · €"&FIXED(ABS(({dal}-{dal_v}))/1000000,1)&"M "&IF(({dal}-{dal_v})<0,"dieper","minder diep")))', kleur=RED)
    kaart(7, 5, "BREAK-EVEN", f'=IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})', "eerste kwartaal met positief saldo",
          f'=IF(OR({hm("be_basis_pos")}=0,{hm("be_vorig_pos")}=0),"geen vergelijking","was "&{hm("be_vorig")}&IF(({hm("be_basis_pos")}-{hm("be_vorig_pos")})=0," · gelijk",'
          f'" · "&ABS(({hm("be_basis_pos")}-{hm("be_vorig_pos")}))&" kw "&IF(({hm("be_basis_pos")}-{hm("be_vorig_pos")})>0,"later","eerder")))', nf="General")
    kaart(12, 5, "PROJECTRESULTAAT", f"={eind}/1000000",
          f'="eindsaldo in "&{hm("laatste_kw")}&" · marge "&IFERROR(FIXED({eind}/{hm("opbr_totaal")}*100,1),"0")&"%"',
          f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose","was "&{eur_m(eind_v)}&IF(ABS(({eind}-{eind_v}))<50000," · gelijk",'
          f'" · €"&FIXED(ABS(({eind}-{eind_v}))/1000000,1)&"M "&IF(({eind}-{eind_v})>0,"hoger","lager")))', kleur=GREEN)
    kaart(17, 5, "CUMULATIEVE CASHFLOW", f'={hm("stand_nu")}/1000000', f'="stand per "&{hm("kw_actuals")}',
          f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose","vorige prognose "&{eur_m(hm("vorig_nu"))})')
    kaart(2, 10, "DOWNSIDE · EINDSALDO", f'={hm("eind_down")}/1000000',
          f'="diepste dal "&{eur_m(hm("dal_down"))}&" in "&{hm("dal_down_kw")}',
          f'="break-even "&IF({hm("be_down")}="","niet bereikt",{hm("be_down")})', kleur=RED, size=16)
    kaart(7, 10, "BASIS · EINDSALDO", f"={eind}/1000000", f'="diepste dal "&{eur_m(dal)}&" in "&{hm("dal_basis_kw")}',
          f'="break-even "&IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})', kleur=BLUE_DARK, size=16)
    kaart(12, 10, "UPSIDE · EINDSALDO", f'={hm("eind_up")}/1000000',
          f'="diepste dal "&{eur_m(hm("dal_up"))}&" in "&{hm("dal_up_kw")}',
          f'="break-even "&IF({hm("be_up")}="","niet bereikt",{hm("be_up")})', kleur=GREEN, size=16)
    kaart(17, 10, "BANDBREEDTE EINDSALDO", f'=ABS({hm("eind_up")}-{hm("eind_down")})/1000000',
          f'="van "&{eur_m(hm("eind_down"))}&" tot "&{eur_m(hm("eind_up"))}',
          f'="diepste dal van "&{eur_m("MIN(" + hm("dal_up") + "," + hm("dal_down") + ")")}&" tot "&{eur_m("MAX(" + hm("dal_up") + "," + hm("dal_down") + ")")}', size=16)

    put(ws, "B14", "Scenario's · cumulatieve cashflow: basis (donkerblauw) met de band tussen downside en upside · grijs vlak = gerealiseerde kwartalen · € mln", f=F_SECTION)
    put(ws, "B43", "Cashflow per kwartaal · balken: opbrengsten en kosten · lijn: cumulatieve cashflow · stippellijn: vorige prognose · "
                   "gele lijn: scenariolijn (knoppen rechts, kolom Scenario) · grijs vlak = gerealiseerd · € mln", f=F_SECTION)
    put(ws, "B69", VERKOOP_TITEL, f=F_SECTION)
    put(ws, f"B{LY.D_ROW_BT}", "Bouwtermijnen per woningtype · kwartaal waarin de termijn vervalt (tab Woningtypes) · per termijn een baan per type "
                               "in de typekleur · lichte tint = vorige prognose (blok 4) · grijs = gerealiseerd", f=F_SECTION)

    # cel-legenda: één gekleurde chip per typeblok (zelfde kleur als in de grafiek), leeg en wit als het blok niet meedoet
    def chips(r, meedoen, tekst):
        ws.row_dimensions[r].height = 16
        for k, col in enumerate(LY.D_LEGENDA_CELLEN, start=1):
            naam = LY.wt_ref(LY.WT_R_NAAM, k)
            cel = f"{col}{r}"
            put(ws, cel, f'=IF({meedoen(k)},IF({naam}="","Type {k}",{naam}),"")', f=font(8, True, "FFFFFF"),
                fl=fill(CX.TYPEKLEUREN[k - 1]), al=Alignment(horizontal="center", vertical="center", shrink_to_fit=True))
            ws.conditional_formatting.add(cel, FormulaRule(formula=[f'{cel}=""'], fill=PatternFill(fill_type="solid", bgColor="FFFFFF", fgColor="FFFFFF")))
        ws.merge_cells(f"O{r}:T{r}")
        put(ws, f"O{r}", tekst, f=F_NOTE, al=AL_RIGHT)

    chips(LY.D_ROW_LEGENDA, lambda k: f'N(Model!{m_col("vcum", k)}$6)>0', "zelfde kleuren in beide panelen · ● = mijlpaal")
    chips(LY.D_ROW_LEGENDA_BT, lambda k: f'Model!${LY.BT[f"cur{k}"]}$7<>"(uit)"',
          f'="lichte tint = vorige prognose · grijs = gerealiseerd t/m "&{hm("kw_actuals")}')

    # ---- parameters (kolom V/W/X/Y, toelichting in Z) --------------------------------
    put(ws, "V1", "KNOPPEN", f=F_TITLE)
    put(ws, "V2", "Geel = invoer. Cijfers vul je in op tab Invoer, woningtypes en termijnen op tab Woningtypes.", f=F_NOTE)

    def sectie(r, t):
        put(ws, f"V{r}", t, f=F_SECTION, bd=BD_SECTION)
        for c in "WXYZ":
            style(ws, f"{c}{r}", bd=BD_SECTION)

    def label(r, t, note=None):
        put(ws, f"V{r}", t, f=F_NOTE, al=AL_VCENTER)
        if note:
            put(ws, f"Z{r}", note, f=F_NOTE8, al=AL_VCENTER)

    def geel(ref, v, nf="General"):
        put(ws, ref, v, f=F_YELLOW, fl=FL_YELLOW, nf=nf, al=AL_CENTER)

    def info(ref, v, nf="General", kleur=TXT):
        put(ws, ref, v, f=font(10, True, kleur), nf=nf, al=AL_CENTER)

    sectie(4, "ALGEMEEN")
    label(5, LY.D_LABELS["actuals_jaar"], "t/m dit kwartaal veranderen de scenario's niets")
    geel("W5", P["actuals_jaar"], "0")
    geel("X5", P["actuals_kw"], "0")
    label(6, LY.D_LABELS["rente"], "over de negatieve stand van het vorige kwartaal, alleen in prognosekwartalen · 0% = geen rente")
    geel("W6", P["rente"], FMT_PCT)
    ws["W6"].comment = _comment("Rente per jaar over de negatieve stand (voorfinanciering) van het vorige kwartaal, gedeeld door 4. "
                                "Alleen in prognosekwartalen. 0% = geen rente.")
    label(7, LY.D_LABELS["rente_tm"], "start bouw = t/m het kwartaal vóór start bouw van het project (+ uitstel) · hele looptijd = alle prognosekwartalen")
    geel("W7", P["rente_tm"])
    label(8, LY.D_LABELS["rente_basis"], "nee = je eigen cashflow bevat al rente; scenario's tellen alleen de extra rente · ja = basis en scenario's krijgen rente")
    geel("W8", P["rente_basis"])
    ws["W8"].comment = _comment("nee: de basis blijft je eigen cashflow; downside en upside krijgen alleen het verschil in rente ten opzichte "
                                "van de basis (door verschuiven, uitstel of andere kosten). ja: ook de basis krijgt rente over de negatieve stand; "
                                "de kaarten en dia's rekenen daar dan mee.")
    label(9, "Start bouw project", "vroegste start bouw van de woningtypes (tab Woningtypes)")
    info("W9", f"={hm('start_tekst')}")

    sectie(11, "VERKOOPTEMPO-MODEL")
    label(12, LY.D_LABELS["model_aan"], "ja = grondtermijn schuift mee met transport · nee = alle prognose-opbrengsten schuiven")
    geel("W12", P["model_aan"])
    ws["W12"].comment = _comment("ja: het model gebruikt je woningtypes, grondtermijn en bouwtermijnen om alleen het verschil door een ander "
                                 "verkooptempo uit te rekenen. nee: de verschuiving verplaatst al je prognose-opbrengsten.")
    label(13, "Model dekt van je opbrengsten", "ter controle: aantal × koopsom volgens je planning, gedeeld door je eigen opbrengsten")
    info("W13", f'=IFERROR({hm("model_opbr")}/{hm("opbr_totaal")},0)', FMT_PCT0)

    r = LY.D_ROW_SCENARIO
    sectie(r, "SCENARIO'S")
    put(ws, f"W{r + 1}", "Downside", f=font(9, True, RED), al=AL_CENTER)
    put(ws, f"X{r + 1}", "Upside", f=font(9, True, GREEN), al=AL_CENTER)
    put(ws, f"Y{r + 1}", "Scenario", f=font(9, True, AMBER_TXT), al=AL_CENTER)   # data.py herkent deze kop
    put(ws, f"Z{r + 1}", "Scenario = de gele lijn in de cashflowgrafiek, met eigen knoppen; downside en upside vormen de band in de scenariografiek",
        f=F_NOTE8, al=AL_VCENTER)
    label(17, LY.D_LABELS["shift_down"], "+ = later, − = eerder")
    geel("W17", P["shift_down"], FMT_KW)
    geel("X17", P["shift_up"], FMT_KW)
    geel("Y17", P["shift_scn"], FMT_KW)
    ws["W17"].comment = _comment("Aantal kwartalen dat verkoop en notarieel transport later (+) of eerder (−) vallen dan in de basis. "
                                 "Geldt alleen voor prognosekwartalen.")
    label(18, LY.D_LABELS["uitstel_down"], "verlengt alleen de renteperiode (bij rente t/m start bouw); kosten en termijnen schuiven niet")
    geel("W18", P["uitstel_down"], FMT_KW)
    geel("X18", P["uitstel_up"], FMT_KW)
    geel("Y18", P["uitstel_scn"], FMT_KW)
    ws["W18"].comment = _comment("Start bouw zoveel kwartalen later dan gepland. Net als in het oude template verlengt dit alleen de periode "
                                 "waarover rente loopt: de kosten in je eigen cashflow en de bouwtermijnen blijven staan.")
    label(19, LY.D_LABELS["opbr_down"], "op alle prognose-opbrengsten (indexatie), ook buiten het verkooptempo-model")
    geel("W19", P["opbr_down"], FMT_PCT_SIGN)
    geel("X19", P["opbr_up"], FMT_PCT_SIGN)
    geel("Y19", P["opbr_scn"], FMT_PCT_SIGN)
    label(20, LY.D_LABELS["kosten_down"], "op alle prognosekosten (kostenindexatie)")
    geel("W20", P["kosten_down"], FMT_PCT_SIGN)
    geel("X20", P["kosten_up"], FMT_PCT_SIGN)
    geel("Y20", P["kosten_scn"], FMT_PCT_SIGN)
    label(21, LY.D_LABELS["rente_pct_down"], "leeg = de algemene jaarrente hierboven · zelfde regels: negatieve stand × rente / 4, alleen prognose")
    geel("W21", P["rente_pct_down"], FMT_PCT)
    geel("X21", P["rente_pct_up"], FMT_PCT)
    geel("Y21", P["rente_pct_scn"], FMT_PCT)
    ws["W21"].comment = _comment("Eigen jaarrente voor dit scenario. Leeg = de algemene jaarrente (knop Jaarrente bovenaan). "
                                 "De renteperiode (t/m start bouw + uitstel, of hele looptijd) blijft gelijk.")
    label(22, LY.D_LABELS["koopsom_down"], "hogere of lagere VON-prijs, dus grondtermijn én bouwtermijnen · werkt via het verkooptempo-model (knop 'ja')")
    geel("W22", P["koopsom_down"], FMT_PCT_SIGN)
    geel("X22", P["koopsom_up"], FMT_PCT_SIGN)
    geel("Y22", P["koopsom_scn"], FMT_PCT_SIGN)
    ws["W22"].comment = _comment("Koopsom (VON-prijs) van alle woningen zoveel procent hoger (+) of lager (−) dan op tab Woningtypes. "
                                 "Grondtermijn en bouwtermijnen schalen mee, op de momenten van het verkooptempo-model. Alleen met "
                                 "'Woningtypes en termijnen gebruiken = ja'; de knop Opbrengsten hieronder werkt op alle prognose-opbrengsten.")
    label(23, "Rente in de scenario's (totaal)")
    info("W23", f'={hm("rente_down_tot")}/1000000', FMT_MLN, RED)
    info("X23", f'={hm("rente_up_tot")}/1000000', FMT_MLN, GREEN)
    info("Y23", f'={hm("rente_scn_tot")}/1000000', FMT_MLN, AMBER_TXT)
    put(ws, "Z23", f'=IF(AND({hm("rente")}=0,{hm("rente_pct_down")}=0,{hm("rente_pct_up")}=0,{hm("rente_pct_scn")}=0),"zet een jaarrente om rente mee te rekenen",'
                   f'"rente over de basis zou €"&FIXED({hm("rente_basis_tot")}/1000000,2)&" mln zijn"&'
                   f'IF({rb}=1," (zit in de basis)"," (niet in de basis: scenario\'s tellen alleen het verschil)"))', f=F_NOTE8, al=AL_VCENTER)
    label(24, "Eindsaldo")
    info("W24", f'={hm("eind_down")}/1000000', FMT_MLN, RED)
    info("X24", f'={hm("eind_up")}/1000000', FMT_MLN, GREEN)
    info("Y24", f'={hm("eind_scn")}/1000000', FMT_MLN, AMBER_TXT)
    put(ws, "Z24", f'="basis "&{eur_m(hm("eind_basis"))}&" · scenario: dieptepunt "&{eur_m(hm("dal_scn"))}&" in "&{hm("dal_scn_kw")}'
                   f'&", break-even "&IF({hm("be_scn")}="","niet bereikt",{hm("be_scn")})', f=F_NOTE8, al=AL_VCENTER)
    label(25, LY.D_LABELS["scn_aan"], "nee = de gele lijn verdwijnt uit de grafiek (de knoppen blijven staan)")
    geel(LY.D["scn_aan"], P["scn_aan"])
    label(26, LY.D_LABELS["scn_naam"], "staat in de legenda en bij het eindpunt, bv. 'Versnelde verkoop' · leeg = Scenario")
    geel(LY.D["scn_naam"], P["scn_naam"])

    r = LY.D_ROW_VERKOOP
    sectie(r, "VERKOOP")
    label(r + 1, LY.D_LABELS["norm"])
    geel(LY.D["norm"], P["norm"], FMT_PCT0)
    label(r + 2, "Verkocht vóór start bouw")
    info(f"W{r + 2}", f'={hm("verk_voor_start_pct")}', FMT_PCT0)
    put(ws, f"Z{r + 2}", f'={hm("verk_voor_start")}&" van "&{hm("won_met_start")}&" woningen (types met start bouw)"', f=F_NOTE8, al=AL_VCENTER)
    label(r + 3, "Verkocht t/m nu")
    info(f"W{r + 3}", f'={hm("verkocht_actuals")}&" van "&{hm("woningen")}')
    put(ws, f"Z{r + 3}", f'="nog te verkopen in de prognose: "&{hm("nog_verkopen")}&" · uitverkocht in "&{hm("uitverkocht")}', f=F_NOTE8, al=AL_VCENTER)
    label(r + 4, "Getransporteerd t/m nu")
    info(f"W{r + 4}", f'={hm("transport_actuals")}&" van "&{hm("woningen")}')
    put(ws, f"Z{r + 4}", f'="nog te transporteren in de prognose: "&{hm("nog_transport")}&" · alles getransporteerd in "&{hm("alles_transport")}', f=F_NOTE8, al=AL_VCENTER)

    r = LY.D_ROW_POWERPOINT
    sectie(r, "POWERPOINT")
    label(r + 1, "Dia's maken met één klik")
    put(ws, f"V{r + 4}", "De knop staat hier in de .xlsm (macro's inschakelen). Zonder macro: blokken plakken vanaf tab PowerPoint.", f=F_NOTE8, al=AL_VCENTER)

    r = LY.D_ROW_CONTROLES
    sectie(r, "CONTROLES")
    vcum1, vcumN = m_col("vcum", 1), m_col("vcum", N_TYPES)
    tcum1, tcumN = m_col("tcum", 1), m_col("tcum", N_TYPES)
    b_rng = LY.in_rng("B")
    checks = [
        ("Periodes oplopend, zonder gaten",
         f'=IF(COUNTIF(Model!${M["stap_ok"]}${ROW1}:${M["stap_ok"]}${ROWN},0)=0,"OK","LET OP: jaar/kwartaal loopt niet op of er ontbreekt een kwartaal")'),
        ("Niet meer dan 60 periodes",
         f'=IF(COUNTA({b_rng})-1<={LY.N_PERIODS},"OK","LET OP: "&(COUNTA({b_rng})-1)&" gevulde rijen, alleen de eerste {LY.N_PERIODS} tellen mee")'),
        ("Elke gevulde rij heeft een jaar",
         f'=IF(SUMPRODUCT(({b_rng}="")*({LY.in_rng("C", LY.WT_RANGE_END)}<>""))=0,"OK","LET OP: rij met cijfers maar zonder jaar telt niet mee")'),
        ("Jaar is een getal tussen 1990 en 2100",
         f'=IF(SUMPRODUCT(({b_rng}<>"")*ISTEXT({b_rng}))-COUNTIF({b_rng},"Jaar")+COUNTIF({b_rng},"<1990")+COUNTIF({b_rng},">2100")=0,"OK",'
         f'"LET OP: jaar als tekst, datum of buiten 1990-2100 (tekst werkt, maar maak er een getal van)")'),
        ("Koprij 'Jaar' staat boven de periodes",
         f'=IF(IFERROR(INDEX({b_rng},MATCH(TRUE,INDEX({b_rng}<>"",0),0)),"")="Jaar","OK",'
         f'"LET OP: koprij \'Jaar\' op tab Invoer ontbreekt; de eerste gevulde rij telt dan niet mee")'),
        ("Verkocht = aantal woningen",
         f'=IF({hm("verkocht_plan")}={hm("woningen")},"OK","LET OP: "&{hm("verkocht_plan")}&" verkocht tegen "&{hm("woningen")}&" woningen")'),
        ("Getransporteerd = aantal woningen",
         f'=IF({hm("transport_plan")}={hm("woningen")},"OK","LET OP: "&{hm("transport_plan")}&" getransporteerd tegen "&{hm("woningen")}&" woningen")'),
        ("Geen transport vóór verkoop",
         f'=IF(SUMPRODUCT(--(Model!${tcum1}${ROW1}:${tcumN}${ROWN}>Model!${vcum1}${ROW1}:${vcumN}${ROWN}))=0,"OK",'
         f'"LET OP: in minstens één periode meer getransporteerd dan verkocht")'),
    ]
    tot = []
    kwc = []
    for k in range(1, N_TYPES + 1):
        aantal = f"N(Model!{m_col('vcum', k)}$6)"
        totaal = LY.wt_ref(LY.WT_R_TOTAAL, k, 1)
        grond_k = LY.wt_ref(LY.WT_R_GROND, k, 1)
        tot.append(f"OR({aantal}=0,N({totaal})=0,ABS(N({totaal})-1)<0.00005,ABS(N({totaal})-(1-N({grond_k})))<0.00005)")
        kw_rng = LY.wt_ref(0, k, 2, rows=(LY.WT_R_T1, LY.WT_R_TN))
        pct_rng = LY.wt_ref(0, k, 1, rows=(LY.WT_R_T1, LY.WT_R_TN))
        kwc.append(f'ABS(SUMIF({kw_rng},">=1",{pct_rng})-SUM({pct_rng}))<0.00005')
    checks.append(("Bouwtermijnen per type 100% (of 100% − grond)",
                   f'=IF({hm("model_aan")}=0,"n.v.t. (model staat uit)",IF(AND({",".join(tot)}),"OK",'
                   f'"LET OP: bouwtermijnen tellen bij minstens één type niet op tot 100% of tot 100% − grondtermijn"))'))
    checks.append(("Elke bouwtermijn heeft een bouwkwartaal",
                   f'=IF({hm("model_aan")}=0,"n.v.t. (model staat uit)",IF(AND({",".join(kwc)}),"OK","LET OP: percentage zonder bouwkwartaal"))'))
    bt_ib = f"Model!${LY.BT['idxb']}${LY.BT_ROW1}:${LY.BT['idxb']}${LY.BT_ROWN}"
    bt_g = f"Model!${LY.BT['gebruikt']}${LY.BT_ROW1}:${LY.BT['gebruikt']}${LY.BT_ROWN}"
    idx_first = f"Model!${FIX['idx']}${ROW1}"
    idx_last = f"INDEX(Model!${FIX['idx']}${ROW1}:${FIX['idx']}${ROWN},MAX(1,{hm('n')}))"
    buiten = f'(COUNTIFS({bt_g},1,{bt_ib},"<"&{idx_first})+COUNTIFS({bt_g},1,{bt_ib},">"&{idx_last}))'
    checks.append(("Bouwtermijnen vallen binnen de periodes",
                   f'=IF({buiten}=0,"OK","LET OP: "&{buiten}&" bouwtermijn(en) vallen buiten de periodes op tab Invoer; in de grafiek staan ze "'
                   f'&"alleen als ze in de getoonde jaren vallen")'))
    checks.append(("Koopsomknop: verkooptempo-model op ja",
                   f'=IF(AND({hm("model_aan")}=0,OR({hm("koopsom_down")}<>0,{hm("koopsom_up")}<>0,{hm("koopsom_scn")}<>0)),'
                   f'"LET OP: zet \'Woningtypes en termijnen gebruiken\' op ja, anders doet de koopsomknop niets","OK")'))
    for i, (t, formule) in enumerate(checks):
        label(r + 1 + i, t)
        put(ws, f"W{r + 1 + i}", formule, f=font(9, True, GREEN), al=AL_LEFT)
    ws.conditional_formatting.add(f"W{r + 1}:W{r + len(checks)}",
                                  FormulaRule(formula=[f'LEFT($W{r + 1},6)="LET OP"'], font=Font(name=ARIAL, bold=True, color=RED)))

    r = LY.D_ROW_TYPES
    sectie(r, "WONINGTYPES (tab Woningtypes)")
    put(ws, f"W{r}", "aantal", f=font(8, True, TXT2), al=AL_CENTER, bd=BD_SECTION)
    put(ws, f"X{r}", "start bouw", f=font(8, True, TXT2), al=AL_CENTER, bd=BD_SECTION)
    for k in range(1, N_TYPES + 1):
        naam = LY.wt_ref(LY.WT_R_NAAM, k)
        bv = f"Model!{m_col('verv', k)}$6"
        put(ws, f"V{r + k}", f'=IF({naam}="","–",{naam})', f=F_NOTE, al=AL_VCENTER)
        put(ws, f"W{r + k}", f'=IF(Model!{m_col("vcum", k)}$6=0,"",Model!{m_col("vcum", k)}$6)', f=F_CALC9, nf=FMT_INT, al=AL_CENTER)
        put(ws, f"X{r + k}", f'=IF({bv}>=99999,"–","Q"&(MOD({bv}-1,4)+1)&" {APOS}"&RIGHT(INT(({bv}-1)/4),2))', f=F_CALC9, al=AL_CENTER)

    # ---- validaties -------------------------------------------------------------------
    for dv, cells in (
        (DataValidation(type="list", formula1='"ja,nee"', allow_blank=False, error="Kies ja of nee."), ["W12", "W8", LY.D["scn_aan"]]),
        (DataValidation(type="list", formula1='"start bouw,hele looptijd"', allow_blank=False, error="Kies 'start bouw' of 'hele looptijd'."), ["W7"]),
        (DataValidation(type="whole", operator="between", formula1="1", formula2="4", allow_blank=True, error="Kwartaal 1 t/m 4."), ["X5"]),
        (DataValidation(type="whole", operator="between", formula1="-20", formula2="20", allow_blank=True, error="Heel aantal kwartalen tussen -20 en 20."), ["W17:Y17"]),
        (DataValidation(type="whole", operator="between", formula1="0", formula2="40", allow_blank=True, error="Heel aantal kwartalen tussen 0 en 40."), ["W18:Y18"]),
        (DataValidation(type="decimal", operator="between", formula1="0", formula2="0.5", allow_blank=True, error="Jaarrente tussen 0% en 50%."), ["W6", "W21:Y21"]),
        (DataValidation(type="decimal", operator="between", formula1="-0.5", formula2="0.5", allow_blank=True, error="Koopsom tussen −50% en +50%."), ["W22:Y22"]),
    ):
        dv.showErrorMessage = True
        dv.errorTitle = "Ongeldige invoer"
        ws.add_data_validation(dv)
        for c in cells:
            dv.add(c)

    # ---- uitleg ----------------------------------------------------------------------------
    r = LY.D_ROW_UITLEG
    put(ws, f"B{r}", "Zo werkt het", f=font(11, True, TXT))
    uitleg = [
        ("Werkwijze", "1. tab Invoer: plak kosten en omzet uit je eigen bestand en vul verkoop en transport in · 2. tab Woningtypes: types, koopsom, "
                      "start bouw en termijnen · 3. zet hier rechts de knoppen · 4. tab PowerPoint: blokken naar de dia's"),
        ("Basis", "het basispad is altijd je eigen cashflow; dit bestand rekent je cijfers niet opnieuw uit (behalve rente, als je die ook in de basis zet)"),
        ("Scenario's", "upside en downside = basis + effect van de knoppen: verkoop/transport eerder of later, uitstel start bouw, opbrengsten %, kosten %, rente "
                       "(alleen op prognosekwartalen)"),
        ("Scenariolijn", "de gele lijn in de cashflowgrafiek is een derde scenario met eigen knoppen (kolom Scenario rechts); zet hem aan of uit en geef hem een naam; "
                         "hij staat los van de band in de scenariografiek"),
        ("Verkooptempo-model", "ja: bij verschuiven schuift de grondtermijn mee met het transport en blijven de bouwtermijnen bij de bouw; alleen dat verschil "
                               "komt bovenop je eigen opbrengsten"),
        ("", "nee: verschuiven verplaatst al je prognose-opbrengsten; woningtypes en termijnen zijn dan niet nodig"),
        ("Rente", "rente = negatieve stand van het vorige kwartaal × jaarrente / 4, t/m het kwartaal vóór start bouw van het project (of de hele looptijd). "
                  "'Rente ook in de basis' = nee: de scenario's tellen alleen de extra rente ten opzichte van de basis"),
        ("Uitstel start bouw", "verlengt alleen de renteperiode van dat scenario; kosten en bouwtermijnen blijven staan (zoals in het oude template)"),
        ("Jaarrente per scenario", "elk scenario kan een eigen jaarrente krijgen; leeg = de algemene jaarrente"),
        ("Koopsom", "per scenario de VON-prijs hoger of lager: grondtermijn en bouwtermijnen schalen mee via het verkooptempo-model (alleen met 'ja')"),
        ("Woningtypes", f"maximaal {N_TYPES} types; elk type heeft op tab Woningtypes een eigen blok van drie kolommen met koopsom, start bouw en eigen termijnen, "
                        "en op tab Invoer twee kolommen (verkocht | transport)"),
        ("Bouwtermijnen", "invullen als % van de aanneemsom (samen 100%) of van de koopsom (samen 100% − grondtermijn); het model schaalt ze naar "
                          "100% − grondtermijn, dus grondtermijn + bouwtermijnen = de koopsom"),
        ("Extra opbrengsten", "per type in € per woning (kopersmeerwerk, kadastrale kosten, overige) met een bouwkwartaal, of leeg = bij transport; "
                              "tellen mee in de modelopbrengst, niet in de koopsomknop"),
        ("Bouwtermijnengrafiek", "per bouwtermijn een baan per woningtype dat hem heeft (zelfde naam bij meerdere typen = één termijn); het blokje "
                                 "staat in het kwartaal waarin de termijn vervalt (start bouw + bouwkwartaal − 1) in de typekleur; lichte tint = "
                                 "hetzelfde met de start bouw uit de vorige prognose (tab Woningtypes blok 4); grijs = gerealiseerde kwartalen; "
                                 "bovenaan de jaartallen en de kwartaalnummers"),
        ("Type toevoegen/verwijderen", "kolommen invoegen of verwijderen op tab Woningtypes (drie) en tab Invoer (twee), of de knoppen op tab Woningtypes (.xlsm); "
                                       "het model leest de blokken op positie, dus alles schuift mee"),
        ("Naar PowerPoint", "knop rechts (alleen in de .xlsm): opent het sjabloon, vult teksten, tabellen en grafieken en bewaart een nieuwe presentatie naast dit bestand"),
        ("Rijen (kwartalen)", "op tab Invoer mag je rijen verwijderen of invoegen; lege rijen tellen niet mee, het model pakt de gevulde rijen op volgorde (maximaal 60)"),
        ("Ander project", "tab Invoer en tab Woningtypes leegmaken en opnieuw vullen; grafieken en KPI's volgen het aantal periodes vanzelf"),
        ("Kringverwijzingen", "geen; elke kolom rekent met de invoer, de typeblokken, kolommen links ervan of de vorige rij"),
    ]
    for i, (kop, tekst) in enumerate(uitleg):
        put(ws, f"B{r + 1 + i}", kop or None, f=F_SECTION)
        put(ws, f"E{r + 1 + i}", tekst, f=F_NOTE)
    return ws


# =============================================================================
#  POWERPOINT
# =============================================================================
def bouw_powerpoint(wb, data):
    ws = wb.create_sheet("PowerPoint")
    ws.sheet_properties.codeName = "shPowerPoint"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B7"
    c_t0 = ws[f"{''.join(ch for ch in LY.PP_TABEL_SC if ch.isalpha())}1"].column
    widths = {"A": 2, "B": 5, "C": 30, "D": 70, "E": 11, "F": 20, "O": 3, "Z": 3, "BK": 3, "BQ": 3, L(c_t0): 30}
    widths[LY.PP_BLOK["bt"]] = 30                       # label 'type · termijn'
    for c in range(7, c_t0):
        widths.setdefault(L(c), 13)
    for c in range(c_t0 + 1, c_t0 + 8):
        widths[L(c)] = 14
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[7].height = 32
    put(ws, "B1", "Gegevens voor PowerPoint", f=F_TITLE)
    put(ws, "B2", "Grafiek bijwerken: blauw kader selecteren → kopiëren → in PowerPoint rechtsklik op de grafiek → Gegevens bewerken → "
                  "klik cel A1 → Plakken speciaal → Waarden. Geen macro nodig.", f=F_NOTE)
    put(ws, "B3", "Sneller: de knop 'Naar PowerPoint' (alleen in de .xlsm) doet dit in één keer: sjabloon openen, teksten, tabellen en grafieken "
                  "vullen, nieuwe presentatie bewaren. Kolom F = de vorm op de dia.", f=F_NOTE)
    put(ws, "B5", "KPI-TEKSTEN VOOR DE DIA'S", f=F_SECTION)
    for col, t in (("B", "Dia"), ("C", "Onderdeel"), ("D", "Tekst voor de dia"), ("E", "Richting"), ("F", "Vorm op de dia")):
        put(ws, f"{col}7", t, f=F_HDR, fl=FL_HDR, al=AL_LEFT_WRAP)

    dal, dal_v, eind, eind_v = hm("dal_basis"), hm("dal_vorig"), hm("eind_basis"), hm("eind_vorig")
    bep, bevp = hm("be_basis_pos"), hm("be_vorig_pos")
    rb = hm("rente_in_basis")
    kpi = [
        (2, "Kopje", f'="KERNCIJFERS · STAND PER "&UPPER({hm("kw_actuals")})', None, "KPI_KICKER"),
        (2, "Titel", "=Dashboard!$B$2", None, "KPI_TITEL"),
        (2, "Ondertitel", f'="Projectresultaat "&{eur_mln(eind)}&" · marge "&IFERROR(FIXED({eind}/{hm("opbr_totaal")}*100,1),"0")&"% · opbrengsten €"'
                          f'&FIXED({hm("opbr_totaal")}/1000000,1)&" mln · kosten €"&FIXED({hm("kosten_totaal")}/1000000,1)&" mln · "&{hm("woningen")}&" woningen"'
                          f'&IF({rb}=1," · incl. rente","")', None, "KPI_SUBTITEL"),
        (2, "Max. financieringsbehoefte", f"={eur_m(dal)}", None, "KPI1_WAARDE"),
        (2, "  toelichting", f'="dieptepunt in "&{hm("dal_basis_kw")}', None, "KPI1_SUB"),
        (2, "  verschil met vorige prognose",
         f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose",IF(ABS(({dal}-{dal_v}))<50000,"gelijk aan vorige prognose",'
         f'"€"&FIXED(ABS(({dal}-{dal_v}))/1000000,1)&"M "&IF(({dal}-{dal_v})<0,"dieper","minder diep")&" dan vorige prognose"))',
         f'=IF(OR({hm("vorig_aanwezig")}=0,ABS(({dal}-{dal_v}))<50000),"gelijk",IF(({dal}-{dal_v})<0,"slechter","beter"))', "KPI1_VERSCHIL"),
        (2, "  vorige prognose", f'=IF({hm("vorig_aanwezig")}=0,"","was "&{eur_m(dal_v)})', None, "KPI1_WAS"),
        (2, "Break-even", f'=IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})', None, "KPI2_WAARDE"),
        (2, "  toelichting", "eerste kwartaal met positief saldo", None, "KPI2_SUB"),
        (2, "  verschil met vorige prognose",
         f'=IF(OR({bep}=0,{bevp}=0),"geen vergelijking met vorige prognose",IF(({bep}-{bevp})=0,"gelijk aan vorige prognose",'
         f'ABS(({bep}-{bevp}))&" kw "&IF(({bep}-{bevp})>0,"later","eerder")&" dan vorige prognose"))',
         f'=IF(OR({bep}=0,{bevp}=0,({bep}-{bevp})=0),"gelijk",IF(({bep}-{bevp})>0,"slechter","beter"))', "KPI2_VERSCHIL"),
        (2, "  vorige prognose", f'=IF({bevp}=0,"","was "&{hm("be_vorig")})', None, "KPI2_WAS"),
        (2, "Projectresultaat", f"={eur_m(eind)}", None, "KPI3_WAARDE"),
        (2, "  toelichting", f'="marge "&IFERROR(FIXED({eind}/{hm("opbr_totaal")}*100,1),"0")&"% · eindsaldo in "&{hm("laatste_kw")}', None, "KPI3_SUB"),
        (2, "  verschil met vorige prognose",
         f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose",IF(ABS(({eind}-{eind_v}))<50000,"gelijk aan vorige prognose",'
         f'"€"&FIXED(ABS(({eind}-{eind_v}))/1000000,1)&"M "&IF(({eind}-{eind_v})>0,"hoger","lager")&" dan vorige prognose"))',
         f'=IF(OR({hm("vorig_aanwezig")}=0,ABS(({eind}-{eind_v}))<50000),"gelijk",IF(({eind}-{eind_v})>0,"beter","slechter"))', "KPI3_VERSCHIL"),
        (2, "  vorige prognose", f'=IF({hm("vorig_aanwezig")}=0,"","was "&{eur_m(eind_v)})', None, "KPI3_WAS"),
        (2, "Cumulatieve cashflow · label", f'="CUM. CASHFLOW PER "&UPPER({hm("kw_actuals")})', None, "KPI4_LABEL"),
        (2, "Cumulatieve cashflow", f"={eur_m(hm('stand_nu'))}", None, "KPI4_WAARDE"),
        (2, "  toelichting", f'=IF({hm("vorig_aanwezig")}=0,"geen vorige prognose","vorige prognose "&{eur_m(hm("vorig_nu"))})', None, "KPI4_SUB"),
        (2, "Bandbreedte eindsaldo", f'={eur_m(hm("eind_down"))}&" tot "&{eur_m(hm("eind_up"))}', None, "KPI5_WAARDE"),
        (2, "  toelichting", "downside tot upside", None, "KPI5_SUB"),
        (2, "Verkocht vóór start bouw", f'=FIXED({hm("verk_voor_start_pct")}*100,0)&"%"',
         f'=IF({hm("verk_voor_start_pct")}>={dabs("norm")},"op norm","onder norm")', "KPI6_WAARDE"),
        (2, "  toelichting", f'="norm "&FIXED({dabs("norm")}*100,0)&"% · "&{hm("verk_voor_start")}&" van "&{hm("won_met_start")}&" woningen"', None, "KPI6_SUB"),
        (2, "Woningen getransporteerd", f'={hm("transport_actuals")}&" van "&{hm("woningen")}', None, "KPI7_WAARDE"),
        (2, "  toelichting", f'="t/m "&{hm("kw_actuals")}&" · verkocht "&{hm("verkocht_actuals")}', None, "KPI7_SUB"),
        (3, "Titel", f'="Dieptepunt "&{eur_mln(dal)}&" in "&{hm("dal_basis_kw")}&", eindsaldo "&{eur_mln(eind)}&" in "&{hm("laatste_kw")}', None, "CF_TITEL"),
        (3, "Ondertitel", f'="Stand per "&{hm("kw_actuals")}&": "&{eur_mln(hm("stand_nu"))}&" · balken: opbrengsten en kosten per kwartaal · lijn: cumulatieve cashflow"'
                          f'&IF({rb}=1," · incl. "&{hm("rente_tekst")},"")', None, "CF_SUBTITEL"),
        (3, "Blokje stand · label", f'="STAND "&UPPER({hm("kw_actuals")})', None, "CF_NU_LABEL"),
        (3, "Blokje stand", f"={eur_m(hm('stand_nu'))}", None, "CF_NU_WAARDE"),
        (3, "Blokje dieptepunt", f'={eur_m(dal)}&" · "&{hm("dal_basis_kw")}', None, "CF_DAL_WAARDE"),
        (3, "Blokje eindsaldo", f'={eur_m(eind)}&" · "&{hm("laatste_kw")}', None, "CF_EIND_WAARDE"),
        (3, "Blokje break-even", f'=IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})', None, "CF_BE_WAARDE"),
        (4, "Titel", f'="Eindsaldo tussen "&{eur_mln(hm("eind_down"))}&" en "&{eur_mln(hm("eind_up"))}&", basis "&{eur_mln(eind)}', None, "SC_TITEL"),
        (4, "Ondertitel", f'="Basis = eigen cashflow"&IF({rb}=1," + rente","")&" · band = ruimte tussen downside en upside · "'
                          f'&IF({hm("model_aan")}=1,"verschuiving via grondtermijn en bouwtermijnen per woningtype","verschuiving van alle prognose-opbrengsten")'
                          f'&IF({hm("rente")}>0," · "&{hm("rente_tekst")},"")', None, "SC_SUBTITEL"),
        (5, "Titel", f'="Bouwtermijnen per woningtype · start bouw "&{hm("start_tekst")}&" (vroegste type)"', None, "BT_TITEL"),
        (5, "Ondertitel", f'="Typekleur = huidige planning · lichte tint = vorige prognose · grijs = gerealiseerd t/m "&{hm("kw_actuals")}'
                          f'&" · "&{hm("bt_uniek")}&" termijnen, "&{hm("bt_types")}&" woningtypes"', None, "BT_SUBTITEL"),
        (6, "Titel", f'={hm("verkocht_actuals")}&" van "&{hm("woningen")}&" woningen verkocht, "&{hm("transport_actuals")}&" getransporteerd"', None, "VP_TITEL"),
        (6, "Ondertitel", f'="Stand per "&{hm("kw_actuals")}&" · per kwartaal en woningtype · uitverkocht in "&{hm("uitverkocht")}'
                          f'&" · laatste transport in "&{hm("alles_transport")}', None, "VP_SUBTITEL"),
        (7, "Titel", f'="Uitverkocht in "&{hm("uitverkocht")}&", alles getransporteerd in "&{hm("alles_transport")}', None, "VO_TITEL"),
        (7, "Uitverkocht in", f"={hm('uitverkocht')}", None, "VO_STAT1_WAARDE"),
        (7, "Alles getransporteerd in", f"={hm('alles_transport')}", None, "VO_STAT2_WAARDE"),
        (7, "Laatste kwartaal", f"={hm('laatste_kw')}", None, "VO_STAT3_WAARDE"),
        # rente: vijfde kleine kaart op dia 2 (sjabloon v10)
        (2, "Rentelasten · label", f'=IF({hm("rente")}=0,"RENTE",IF({rb}=1,"RENTELASTEN (IN BASIS)","RENTELASTEN (BEREKEND)"))', None, "KPI8_LABEL"),
        (2, "Rentelasten", f'=IF({hm("rente")}=0,"geen",{eur_m(hm("rente_basis_tot"))})', None, "KPI8_WAARDE"),
        (2, "  toelichting", f'=IF({hm("rente")}=0,"zet een jaarrente op het Dashboard",FIXED({hm("rente")}*100,1)&"% "&IF({hm("rente_tm_start")}=1,"t/m start bouw","hele looptijd")'
                             f'&" · down "&{eur_m(hm("rente_down_tot"))}&" · up "&{eur_m(hm("rente_up_tot"))})', None, "KPI8_SUB"),
    ]
    for i, (dia, onderdeel, tekst, richting, vorm) in enumerate(kpi):
        r = LY.PP_KPI_ROW1 + i
        put(ws, f"B{r}", dia, f=F_NOTE, bd=BD_ROW)
        put(ws, f"C{r}", onderdeel, f=F_NOTE, bd=BD_ROW)
        put(ws, f"D{r}", tekst, f=F_BOLD9, al=AL_LEFT_TOP, bd=BD_ROW)
        put(ws, f"E{r}", richting, f=F_NOTE, bd=BD_ROW)
        put(ws, f"F{r}", vorm, f=F_NOTE8, bd=BD_ROW)
    r_s = int(LY.PP_CEL_SJABLOON[1:])
    assert LY.PP_KPI_ROW1 + len(kpi) - 1 <= r_s - 2, "KPI-regels lopen tot in de instellingen (de rij erboven moet leeg blijven)"
    put(ws, f"C{r_s}", "Sjabloon (pptx)", f=F_NOTE)
    put(ws, LY.PP_CEL_SJABLOON, data.params.get("sjabloon") or None, f=F_INPUT, fl=FL_INPUT, al=AL_LEFT_TOP)
    put(ws, f"E{r_s}", "leeg = Kwartaal_Template_cashflow_v13.pptx in dezelfde map als dit bestand; anders het volledige pad", f=F_NOTE8)
    put(ws, f"C{r_s + 1}", "Naam nieuwe presentatie", f=F_NOTE)
    put(ws, LY.PP_CEL_NAAM, f'="Cashflow update Q"&{dabs("actuals_kw")}&" "&{dabs("actuals_jaar")}', f=F_CALC9, al=AL_LEFT_TOP)
    put(ws, f"E{r_s + 1}", "de knop zet er datum en tijd achter en bewaart naast dit bestand", f=F_NOTE8)

    # ---- grafiekblokken (kop in rij 7, één rij per periode) ------------------------------
    blokken = [
        (LY.PP_BLOK["cf"], "DIA 3 · CASHFLOW PER KWARTAAL · € mln", "grafiek cashflow · plakken: G7 t/m N, laatste periode (kolom N = scenariolijn; sjabloon v11)",
         ["Kwartaal", "Voorfinanciering", "Positief saldo", "Opbrengsten", "Kosten", "Cumulatieve cashflow", "Vorige prognose", f"={hm('lbl_scenario')}"],
         [FIX["kwartaal"], M["g_voorfinanciering"], M["g_positief_saldo"], M["g_opbrengsten"], M["g_kosten"], M["g_stand"], M["g_vorige"], M["g_scenario"]], "0.0"),
        (LY.PP_BLOK["sc"], "DIA 4 · SCENARIO'S · € mln", "grafiek scenario's · plakken: P7 t/m Y, laatste periode",
         ["Kwartaal", "Band onder", "Bandbreedte", "Downside", "Upside", "Basis", f"={hm('lbl_nu')}", f"={hm('lbl_dal')}", f"={hm('lbl_eind_up')}", f"={hm('lbl_eind_down')}"],
         [FIX["kwartaal"], M["g_band_onder"], M["g_bandbreedte"], M["g_downside"], M["g_upside"], M["g_stand"], M["g_punt_nu"], M["g_punt_dal"], M["g_eind_upside"], M["g_eind_downside"]], "0.0"),
        (LY.PP_BLOK["bt"], "DIA 5 · BOUWTERMIJNEN · tijd in kwartaalindex (jaar × 4 + kwartaal)",
         f"grafiek bouwtermijnen · koprijen jaar/kwartaal, daarna per termijn een baan per type · plakken: AA7 t/m {L(26 + LY.PP_BT_KOL)}, "
         f"laatste rij ({hm('bt_n')} rijen) · '(uit)' = reeks niet in de presentatie",
         ["Termijn"] + [LY.BT_SEG_NAMEN.get(x, f"=Model!${LY.BT[x]}$7") for x in LY.BT_REEKSEN],
         [LY.BT["label"]] + [LY.BT[x] for x in LY.BT_REEKSEN], "0.00"),
        (LY.PP_BLOK["vp"], "DIA 6 · VERKOOP EN TRANSPORT PER WONINGTYPE · aantal woningen", "twee grafieken (verkocht / getransporteerd) lezen dit blok · plakken: AK7 t/m BJ, laatste periode",
         ["Kwartaal"] + [f"=Model!${m_col('gv', k)}$7" for k in range(1, N_TYPES + 1)] + [f"=Model!${m_col('gt', k)}$7" for k in range(1, N_TYPES + 1)]
         + ["Verkocht", "Getransporteerd", f"={hm('lbl_uitverkocht')}", f"={hm('lbl_alles_transport')}", "Realisatie"],
         [FIX["kwartaal"]] + [m_col("gv", k) for k in range(1, N_TYPES + 1)] + [m_col("gt", k) for k in range(1, N_TYPES + 1)]
         + [M["g_verkocht"], M["g_getransporteerd"], M["g_punt_uitverkocht"], M["g_punt_alles_transport"], M["g_realisatie"]], FMT_INT),
        (LY.PP_BLOK["vo"], "DIA 7 · VAN VERKOOP NAAR OMZET · cumulatief %", "grafiek verkoop naar omzet · plakken: BL7 t/m BP, laatste periode",
         ["Kwartaal", "Verkocht", "Getransporteerd", "Opbrengsten ontvangen", "Kosten gemaakt"],
         [FIX["kwartaal"], M["g_verkocht_pct"], M["g_getransporteerd_pct"], M["g_opbrengsten_pct"], M["g_kosten_pct"]], FMT_PCT0),
    ]
    blauw = side("medium", PP_BORDER)
    for col0, titel, sub, koppen, bronnen, nf in blokken:
        c0 = ws[f"{col0}1"].column
        put(ws, f"{col0}4", titel, f=F_SECTION)
        put(ws, f"{col0}5", sub, f=F_NOTE8)
        n = len(koppen)
        bt = col0 == LY.PP_BLOK["bt"]                                   # bouwtermijnenblok: BT_N rijen uit de bouwtermijnentabel
        rijen = LY.BT_N if bt else LY.N_PERIODS
        n_ref = hm("bt_n") if bt else hm("n")
        r1, rN = (LY.BT_ROW1, LY.BT_ROWN) if bt else (ROW1, ROWN)
        for j, kop in enumerate(koppen):
            c = L(c0 + j)
            bd = Border(top=blauw, left=blauw if j == 0 else None, right=blauw if j == n - 1 else None)
            put(ws, f"{c}7", kop or None, f=F_HDR, fl=FL_HDR, al=AL_LEFT_WRAP if j == 0 else AL_RIGHT_WRAP, bd=bd)
        for i in range(rijen):
            r = LY.PP_KPI_ROW1 + i
            nr = i + 1
            for j, bron in enumerate(bronnen):
                c = L(c0 + j)
                if j == 0:
                    formule = f'=IF({nr}>{n_ref},"",INDEX(Model!${bron}${r1}:${bron}${rN},{nr}))'
                else:
                    formule = f'=IF({nr}>{n_ref},NA(),INDEX(Model!${bron}${r1}:${bron}${rN},{nr}))'
                bd = Border(left=blauw if j == 0 else None, right=blauw if j == n - 1 else None, bottom=blauw if i == rijen - 1 else None)
                put(ws, f"{c}{r}", formule, f=F_CALC9, nf=None if j == 0 else nf, al=AL_LEFT_TOP if j == 0 else None, bd=bd)
    ws.conditional_formatting.add(f"G{LY.PP_KPI_ROW1}:{L(c_t0 - 1)}{LY.PP_KPI_ROW1 + LY.BT_N - 1}",
                                  FormulaRule(formula=[f"ISERROR(G{LY.PP_KPI_ROW1})"], font=Font(name=ARIAL, color="D0D3D8")))

    # ---- tabellen (kolom T0 = eerste tabelkolom) --------------------------------------------
    T0 = "".join(ch for ch in LY.PP_TABEL_SC if ch.isalpha())
    T1, T2, T3, T4 = (L(c_t0 + j) for j in (1, 2, 3, 4))
    put(ws, f"{T0}4", "DIA 4 · TABEL SCENARIO'S", f=F_SECTION)
    for col, t in ((T0, None), (T1, "Downside"), (T2, "Basis"), (T3, "Upside")):
        put(ws, f"{col}7", t, f=F_HDR, fl=FL_HDR, al=AL_LEFT_WRAP if col == T0 else AL_RIGHT_WRAP)

    def knop(ref, nf_kw=False):
        if nf_kw:
            return f'=IF({ref}=0,"0 kw",IF({ref}>0,"+","{MINUS}")&ABS({ref})&" kw")'
        return f'=IF({ref}=0,"0%",IF({ref}>0,"+","{MINUS}")&FIXED(ABS({ref})*100,0)&"%")'

    sc = [
        ("Eindsaldo", f"={eur_m(hm('eind_down'))}", f"={eur_m(eind)}", f"={eur_m(hm('eind_up'))}"),
        ("Diepste dal", f"={eur_m(hm('dal_down'))}", f"={eur_m(dal)}", f"={eur_m(hm('dal_up'))}"),
        ("Diepste dal in", f"={hm('dal_down_kw')}", f"={hm('dal_basis_kw')}", f"={hm('dal_up_kw')}"),
        ("Break-even", f'=IF({hm("be_down")}="","niet bereikt",{hm("be_down")})', f'=IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})',
         f'=IF({hm("be_up")}="","niet bereikt",{hm("be_up")})'),
        ("KNOPPEN", None, None, None),
        ("Verschuiving", knop(dabs("shift_down"), True), "–", knop(dabs("shift_up"), True)),
        ("Opbrengsten", knop(dabs("opbr_down")), "–", knop(dabs("opbr_up"))),
        ("Kosten", knop(dabs("kosten_down")), "–", knop(dabs("kosten_up"))),
    ]
    for i, (t, a, b, c) in enumerate(sc):
        r = 8 + i
        put(ws, f"{T0}{r}", t, f=F_NOTE)
        for col, v in ((T1, a), (T2, b), (T3, c)):
            put(ws, f"{col}{r}", v, f=F_CALC9, al=AL_RIGHT)
    # extra knoppen en de scenariolijn: niet in de tabel op dia 4 (die heeft 9 rijen), wel om over te nemen
    r_extra = LY.PP_TABEL_VT_ROW + N_TYPES + 3
    put(ws, f"{T0}{r_extra}", "DIA 4 · EXTRA KNOPPEN EN SCENARIOLIJN (niet in de tabel op de dia)", f=F_SECTION)
    for col, t in ((T0, None), (T1, "Downside"), (T2, "Basis"), (T3, "Upside"), (T4, f"={hm('scn_naam')}")):
        put(ws, f"{col}{r_extra + 1}", t, f=F_HDR, fl=FL_HDR, al=AL_LEFT_WRAP if col == T0 else AL_RIGHT_WRAP)
    dal_scn = f'{eur_m(hm("dal_scn"))}&" · "&{hm("dal_scn_kw")}'
    extra = [
        ("Verschuiving", knop(dabs("shift_down"), True), "–", knop(dabs("shift_up"), True), knop(dabs("shift_scn"), True)),
        ("Uitstel start bouw", knop(dabs("uitstel_down"), True), "–", knop(dabs("uitstel_up"), True), knop(dabs("uitstel_scn"), True)),
        ("Opbrengsten", knop(dabs("opbr_down")), "–", knop(dabs("opbr_up")), knop(dabs("opbr_scn"))),
        ("Kosten", knop(dabs("kosten_down")), "–", knop(dabs("kosten_up")), knop(dabs("kosten_scn"))),
        ("Koopsom (VON)", knop(dabs("koopsom_down")), "–", knop(dabs("koopsom_up")), knop(dabs("koopsom_scn"))),
        ("Jaarrente", f'=FIXED({hm("rente_pct_down")}*100,1)&"%"', f'=FIXED({hm("rente")}*100,1)&"%"', f'=FIXED({hm("rente_pct_up")}*100,1)&"%"',
         f'=FIXED({hm("rente_pct_scn")}*100,1)&"%"'),
        ("Rente", f'={eur_m(hm("rente_down_tot"), 2)}', f'=IF({rb}=1,{eur_m(hm("rente_basis_tot"), 2)},"–")', f'={eur_m(hm("rente_up_tot"), 2)}',
         f'={eur_m(hm("rente_scn_tot"), 2)}'),
        ("Eindsaldo", f"={eur_m(hm('eind_down'))}", f"={eur_m(eind)}", f"={eur_m(hm('eind_up'))}", f"={eur_m(hm('eind_scn'))}"),
        ("Diepste dal", f"={eur_m(hm('dal_down'))}", f"={eur_m(dal)}", f"={eur_m(hm('dal_up'))}", f"={dal_scn}"),
        ("Break-even", f'=IF({hm("be_down")}="","niet bereikt",{hm("be_down")})', f'=IF({hm("be_basis")}="","niet bereikt",{hm("be_basis")})',
         f'=IF({hm("be_up")}="","niet bereikt",{hm("be_up")})', f'=IF({hm("be_scn")}="","niet bereikt",{hm("be_scn")})'),
        ("Rente-instelling", f'={hm("rente_tekst")}', None, None, f'=IF({hm("scn_aan")}=1,"lijn aan","lijn uit")'),
    ]
    for i, (t, a, b, c, d_) in enumerate(extra):
        r = r_extra + 2 + i
        put(ws, f"{T0}{r}", t, f=F_NOTE)
        for col, v in ((T1, a), (T2, b), (T3, c), (T4, d_)):
            put(ws, f"{col}{r}", v, f=F_CALC9, al=AL_RIGHT)

    put(ws, f"{T0}{LY.PP_TABEL_VT_ROW - 1}", "DIA 5 · TABEL WONINGTYPES", f=F_SECTION)
    r0 = LY.PP_TABEL_VT_ROW
    for j, t in enumerate(["Type", "Aantal", "Verkocht t/m nu", "Getransporteerd t/m nu", "Start bouw", "Vóór start bouw", "In %", "Norm"]):
        put(ws, f"{L(c_t0 + j)}{r0}", t, f=F_HDR, fl=FL_HDR, al=AL_LEFT_WRAP if j == 0 else AL_RIGHT_WRAP)
    for k in range(1, N_TYPES + 1):
        r = r0 + k
        vc, tc, bv = m_col("vcum", k), m_col("tcum", k), m_col("verv", k)
        aantal = f"N(Model!{vc}$6)"
        leeg = f'IF({aantal}=0,"",'
        start = f"Model!{bv}$6"
        pos = hm("pos_actuals")
        cel = lambda j: f"{L(c_t0 + j)}{r}"  # noqa: E731
        put(ws, cel(0), f'={leeg}{LY.wt_ref(LY.WT_R_NAAM, k)})', f=F_CALC9, al=AL_LEFT_TOP)
        put(ws, cel(1), f'={leeg}{aantal})', f=F_CALC9, nf=FMT_INT, al=AL_RIGHT)
        put(ws, cel(2), f'={leeg}IF({pos}=0,0,N(INDEX(Model!${vc}${ROW1}:${vc}${ROWN},{pos}))))', f=F_CALC9, nf=FMT_INT, al=AL_RIGHT)
        put(ws, cel(3), f'={leeg}IF({pos}=0,0,N(INDEX(Model!${tc}${ROW1}:${tc}${ROWN},{pos}))))', f=F_CALC9, nf=FMT_INT, al=AL_RIGHT)
        put(ws, cel(4), f'={leeg}IF({start}>=99999,"–","Q"&(MOD({start}-1,4)+1)&" {APOS}"&RIGHT(INT(({start}-1)/4),2)))', f=F_CALC9, al=AL_RIGHT)
        put(ws, cel(5), f'={leeg}IF({start}>=99999,"–",Model!{vc}$5))', f=F_CALC9, nf=FMT_INT, al=AL_RIGHT)
        put(ws, cel(6), f'={leeg}IF({start}>=99999,"–",FIXED(Model!{vc}$5/{aantal}*100,0)&"%"))', f=F_CALC9, al=AL_RIGHT)
        put(ws, cel(7), f'={leeg}IF({start}>=99999,"–",FIXED({dabs("norm")}*100,0)&"%"))', f=F_CALC9, al=AL_RIGHT)
    return ws
