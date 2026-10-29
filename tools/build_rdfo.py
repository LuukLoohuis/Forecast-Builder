"""Bouwt het RDFO-prototype (Excel) uit de twee operationele bronnen en de fictieve historie.

Gebruik:  python tools/build_rdfo.py [uitvoerbestand.xlsx]

Lagen in de werkmap (links naar rechts in de tabvolgorde):
  rapportage  -> Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Controle
  invoer      -> Parameters, Invoer_Aanvullend, Invoer_Toelichting, Invoer_Koperstype, Mapping
  versies     -> Vastlegging (onveranderlijk grootboek), Vastleggen (concept uit upload), Importlog
  berekening  -> VERW_Project, CALC_Versies
  staging     -> STG_ZPMW03, STG_Woningetalage  (kolommen gezocht op header, niet op positie)
  bron        -> SRC_ZPMW03, SRC_Woningetalage (ongewijzigde export, alleen plakken)
  documentatie-> Veldwoordenboek, Rapportagemapping, Productievragen, Handleiding, Testresultaten
"""
import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlhelpers import *  # noqa: E402,F403
import docs_content as DOC  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SRC_Z_FILE = "ZPMW03_FICTIEF_20_projecten.xlsx"
SRC_E_FILE = "Woningetalage_FICTIEF_20_projecten.xlsx"
HIST_FILE = "RDFO_FICTIEVE_historie_en_mapping.xlsx"

PEIL = dt.datetime(2026, 9, 30)
IMPORTDATUM = dt.datetime(2026, 10, 29)

N_SRC = 500                      # capaciteit bronregels per bron
STG0, STG1 = 6, 6 + N_SRC - 1    # stagingrijen
MAP0, MAP1 = 5, 104              # projectmapping (100 regels)
V0, V1 = 6, 105                  # verwerkingsrijen (1 per mappingregel)
L0, L1 = 6, 5005                 # grootboek vastlegging (5000 regels)
A0, A1 = 6, 505                  # aanvullende invoer
T0, T1 = 6, 505                  # toelichtingen
K0, K1 = 6, 505                  # koperstype-invoer
VERSIES = ["Budget", "Q1", "Q2", "Q3", "Q4"]

K_OMZ, K_BOR, K_OMZA, K_BORA = "Omzet FY", "BOR FY", "Omzet actual YTD", "BOR actual YTD"
K_WFY, K_WVJ, K_WACT = "Woningverkopen FY", "Woningen verkocht t/m vorig jaar", "Woningverkopen actual YTD"
K_WQ = [f"Woningverkopen Q{q}" for q in range(1, 5)]

RUBRIEKEN = ["Taakstelling omzet", "Interne eliminaties", "IFRS-eliminaties", "AK-dekking",
             "Fees", "Overige resultaten", "Taakstelling EBIT"]

SHEETS = [
    ("Start", NAVY), ("Dashboard", YELLOW), ("RDFO", YELLOW), ("Projecten", YELLOW),
    ("Projectdetail", YELLOW), ("Woningen", YELLOW), ("Controle", RED),
    ("Parameters", "1F4E9A"), ("Invoer_Aanvullend", "1F4E9A"), ("Invoer_Toelichting", "1F4E9A"),
    ("Invoer_Koperstype", "1F4E9A"), ("Mapping", "1F4E9A"),
    ("Vastlegging", GREEN), ("Vastleggen", GREEN), ("Importlog", GREEN),
    ("VERW_Project", GREY), ("CALC_Versies", GREY), ("STG_ZPMW03", GREY), ("STG_Woningetalage", GREY),
    ("SRC_ZPMW03", "000000"), ("SRC_Woningetalage", "000000"),
    ("Veldwoordenboek", MID), ("Rapportagemapping", MID), ("Productievragen", MID),
    ("Handleiding", MID), ("Testresultaten", MID),
]


def rng(sheet, c, r0, r1):
    return f"{sheet}!${c}${r0}:${c}${r1}"


# ---------------------------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------------------------
def build_parameters(wb, ws):
    base_sheet(ws, zoom=100)
    title(ws, "Parameters", "Gele cellen zijn invoer. Alle jaartallen in kolomkoppen en rapporttitels zijn hiervan afgeleid.", 6)
    widths(ws, {"B": 34, "C": 22, "D": 90, "E": 4, "F": 4})
    header(ws, 4, 2, ["Parameter", "Waarde", "Toelichting"], height=20)
    params = [
        ("P_Regio", "Regio / rapportage-eenheid", "Demo Regio", None, "Alleen voor titels."),
        ("P_RJ", "Rapportagejaar (lopend jaar in upload)", 2026, "0", "Bepaalt welke ZPMW03-kolommen gelezen worden: 'OP.BOB <jaar>', 'OP.BOB <jaar+1>-Q1', ... en de Woningetalage-jaarkolom <jaar+2>."),
        ("P_Peildatum", "Peildatum upload", PEIL, FMT_DATE, "Datum t/m welke 'Bedrijfsopbrengsten LJ' en 'Huidig jr' actual zijn. Prototype: 30-09-2026."),
        ("P_PeilMaand", "Peilmaand (berekend)", "=MONTH(P_Peildatum)", "0", "hjr-maanden t/m deze maand = actual, daarna = prognose."),
        ("P_PeilTxt", "Peildatum als tekst (berekend)", '=RIGHT("0"&DAY(P_Peildatum),2)&"-"&RIGHT("0"&MONTH(P_Peildatum),2)&"-"&YEAR(P_Peildatum)', None, "Locale-onafhankelijk (geen TEXT-functie)."),
        ("P_UploadVJ", "Versiejaar van de upload", 2026, "0", "In welk versiejaar hoort de huidige upload thuis."),
        ("P_UploadVersie", "Versie van de upload", "Q3", None, "Budget/Q1/Q2/Q3/Q4. Wordt gecontroleerd tegen de peildatum (Controle)."),
        ("P_VJ", "Versiejaar voor rapportage", 2026, "0", "De kolommen Budget/Q1..Q4 in de rapporten tonen vastgelegde versies uit dit versiejaar."),
        ("P_Eenheid", "Eenheid rapportage RDFO/Projecten", "EUR x 1.000", None, "EUR / EUR x 1.000 / EUR mln. Bronprecisie blijft volledig in EUR bewaard."),
        ("P_Factor", "Deelfactor (berekend)", '=IF(P_Eenheid="EUR mln",1000000,IF(P_Eenheid="EUR x 1.000",1000,1))', "#,##0", ""),
        ("P_Demo", "Fictieve demo-aannames meenemen?", "Nee", None, "Ja = regels met status 'Fictieve aanname' in Invoer_Aanvullend tellen mee in RDFO. Standaard Nee."),
        ("P_ImportID", "Import-ID huidige upload", "IMP-2026-Q3-01", None, "Wordt meegegeven aan elke vastgelegde regel (Vastleggen)."),
        ("P_BronZ", "Bestandsnaam ZPMW03", SRC_Z_FILE, None, "Documentatie / audittrail."),
        ("P_BronE", "Bestandsnaam Woningetalage", SRC_E_FILE, None, "Documentatie / audittrail."),
        ("P_Importdatum", "Importdatum", IMPORTDATUM, FMT_DATE, "Datum waarop de bronnen in deze werkmap zijn geplakt."),
        ("P_RegelsZ", "Verwacht aantal regels ZPMW03", 20, "0", "Aantal dataregels volgens het exportbestand. Controle vergelijkt met wat werkelijk geplakt is."),
        ("P_RegelsE", "Verwacht aantal regels Woningetalage", 20, "0", "Idem."),
        ("P_StatusVastleg", "Status bij vastleggen", "Vastgesteld", None, "Concept of Vastgesteld. Alleen 'Vastgesteld' telt in de versiekolommen."),
        ("P_ToelCorrectie", "Toelichting bij correctie (revisie > 0)", "", None, "Verplicht als je een al vastgestelde versie opnieuw vastlegt."),
        ("P_Tol", "Tolerantie aansluitingen (EUR / aantal)", 1, "0.00", "Afrondingsmarge voor controles."),
    ]
    inputs = {"P_Regio", "P_RJ", "P_Peildatum", "P_UploadVJ", "P_UploadVersie", "P_VJ", "P_Eenheid", "P_Demo",
              "P_ImportID", "P_BronZ", "P_BronE", "P_Importdatum", "P_RegelsZ", "P_RegelsE", "P_StatusVastleg",
              "P_ToelCorrectie", "P_Tol"}
    r = 5
    pos = {}
    for name, lab, val, fmt, toel in params:
        pos[name] = r
        put(ws, f"B{r}", lab)
        c = put(ws, f"C{r}", val, fmt=fmt, align="right")
        if name in inputs:
            input_style(c, fmt)
        else:
            helper_style(c)
        put(ws, f"D{r}", toel, color=GREY, italic=True)
        add_name(wb, name, f"Parameters!$C${r}")
        r += 1

    # lijsten voor keuzelijsten
    r += 1
    section(ws, r, 2, "Keuzelijsten (gebruikt door validatie en formules)", 3)
    r += 1
    header(ws, r, 2, ["Versie", "Volgorde", "Peilmaand (verwacht)"], height=20)
    v0 = r + 1
    for i, (v, m) in enumerate(zip(VERSIES, [12, 3, 6, 9, 12])):
        put(ws, f"B{v0+i}", v)
        put(ws, f"C{v0+i}", i, align="right")
        put(ws, f"D{v0+i}", m, align="left")
    add_name(wb, "V_List", f"Parameters!$B${v0}:$B${v0+4}")
    add_name(wb, "V_Maand", f"Parameters!$D${v0}:$D${v0+4}")
    r = v0 + 6
    lists = [
        ("L_TypeList", ["Alle", "Opstal", "Deelneming"]),
        ("L_EenheidList", ["EUR", "EUR x 1.000", "EUR mln"]),
        ("L_JaNee", ["Ja", "Nee"]),
        ("L_StatusList", ["Concept", "Vastgesteld"]),
        ("L_AanvStatus", ["Vastgesteld", "Concept", "Fictieve aanname", "Voorbeeld"]),
        ("L_Rubrieken", RUBRIEKEN),
        ("L_Soort", ["FY", "Actual YTD"]),
        ("L_Niveau", ["Regio", "Project"]),
    ]
    for name, vals in lists:
        put(ws, f"B{r}", name.replace("L_", "Lijst: "), bold=True)
        for i, v in enumerate(vals):
            put(ws, f"{col(3+i)}{r}", v)
        add_name(wb, name, f"Parameters!${col(3)}${r}:${col(3+len(vals)-1)}${r}")
        r += 1
    put(ws, f"B{r}", "Lijst: doeljaren", bold=True)
    put(ws, f"C{r}", "=P_RJ")
    put(ws, f"D{r}", "=P_RJ+1")
    add_name(wb, "L_Doeljaar", f"Parameters!$C${r}:$D${r}")
    r += 2
    # projectlijst voor filter
    put(ws, f"B{r}", "Lijst: projecten (Alle + mapping)", bold=True)
    put(ws, f"C{r}", "Alle")
    for i in range(MAP1 - MAP0 + 1):
        put(ws, f"C{r+1+i}", f'=IF(Mapping!$C${MAP0+i}="","",Mapping!$C${MAP0+i})', color=GREY)
    add_name(wb, "L_Projecten", f"Parameters!$C${r}:$C${r+1+MAP1-MAP0}")
    add_name(wb, "L_ProjectenDet", f"Parameters!$C${r+1}:$C${r+1+MAP1-MAP0}")

    dv_list(ws, f"C{pos['P_UploadVersie']}", "V_List")
    dv_list(ws, f"C{pos['P_Eenheid']}", "L_EenheidList")
    dv_list(ws, f"C{pos['P_Demo']}", "L_JaNee")
    dv_list(ws, f"C{pos['P_StatusVastleg']}", "L_StatusList")
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------------------------------------
# Bronnen (ongewijzigd)
# ---------------------------------------------------------------------------------------------
def copy_source(ws, path, sheetname):
    src = load_workbook(path)[sheetname]
    base_sheet(ws, grid=True, zoom=85)
    n = 0
    for row in src.iter_rows():
        for c in row:
            if c.value is None:
                continue
            d = ws.cell(row=c.row, column=c.column, value=c.value)
            if c.number_format and c.number_format != "General":
                d.number_format = c.number_format
            elif hasattr(c.value, "year"):
                d.number_format = FMT_DATE
        n = max(n, row[0].row)
    for c in ws[1]:
        c.font = Font(name=FONT, size=9, bold=True)
        c.fill = fill(MID)
    ws.freeze_panes = "A2"
    return n - 1


# ---------------------------------------------------------------------------------------------
# Staging: kolommen op header zoeken, rijen 1-op-1 met de bron, managementteken in afgeleide velden
# ---------------------------------------------------------------------------------------------
def z_fields():
    f = [
        dict(n="WBS", h="WBS-element", k="key"),
        dict(n="IFRS", h="IFRS J/N", k="txt"),
        dict(n="Bedrijfsnummer", h="Bedrijfsnummer", k="raw"),
        dict(n="Bedrijfsnaam", alias=["Naam maatschappij", "Bedrijfsnaam"], k="txt"),
        dict(n="ProfitCenter", h="Profit center", k="raw"),
        dict(n="ProfitCenterNaam", h="Profitcenter naam", k="txt"),
        dict(n="Marktsegment", h="Marktsegment", k="txt"),
        dict(n="Projectdefinitie", h="Projectdefinitie", k="txt"),
        dict(n="Omschrijving", h="Omschrijving", k="txt"),
        dict(n="StatusWBS", h="Status WBS", k="txt"),
        dict(n="Kans", h="Kans%", k="num"),
        dict(n="PEW_BOB", h="PEW opbrengst", k="num"),
        dict(n="BOB_VJ", h="Bedrijfsopbrengsten VJ", k="num"),
        dict(n="BOB_LJ", h="Bedrijfsopbrengsten LJ", k="num"),
        dict(n="Kosten_LJ", h="Kosten LJ", k="num"),
        dict(n="BOR_LJ", h="Bruto (omzet) result. LJ", k="num"),
        dict(n="Port_BOB", h="Portefeuille BOB", k="num"),
        dict(n="Port_BOR", h="Portefeuille BOR", k="num"),
    ]
    for tag, off in (("RJ", 0), ("RJ1", 1)):
        yr = f'(P_RJ+{off})' if off else "P_RJ"
        for q in range(1, 5):
            f.append(dict(n=f"BOB_{tag}_Q{q}", h=f'="OP.BOB "&{yr}&"-Q{q}"', k="num"))
            f.append(dict(n=f"BOR_{tag}_Q{q}", h=f'="OP.BOR "&{yr}&"-Q{q}"', k="num"))
        f.append(dict(n=f"BOB_{tag}", h=f'="OP.BOB "&{yr}', k="num"))
        f.append(dict(n=f"Kosten_{tag}", h=f'="Kosten "&{yr}', k="num"))
        f.append(dict(n=f"BOR_{tag}", h=f'="OP.BOR "&{yr}', k="num"))
    for o in range(2, 9):
        f.append(dict(n=f"BOB_RJ{o}", h=f'="OP.BOB "&(P_RJ+{o})', k="num"))
        f.append(dict(n=f"BOR_RJ{o}", h=f'="OP.BOR "&(P_RJ+{o})', k="num"))
    f += [
        dict(n="BOB_ONV", h="OP.BOB onverd.", k="num"),
        dict(n="BOR_ONV", h="OP.BOR onverd.", k="num"),
        dict(n="POC", h="Huidige POC %", k="num"),
        dict(n="Orderintake", h="Orderintake", k="raw"),
        dict(n="StartBouw", h="Start bouw", k="raw"),
        dict(n="Oplevering", h="Maand oplevering", k="raw"),
        dict(n="Risico", h="Risicoklasse", k="raw"),
    ]
    return f


def z_derived():
    neg = lambda x: f'=IFERROR(-[{x}],"n.b.")'  # noqa: E731
    d = [
        ("M_OmzYTD", neg("BOB_LJ"), "Omzet actual YTD = -Bedrijfsopbrengsten LJ"),
        ("M_BorYTD", neg("BOR_LJ"), "BOR actual YTD = -Bruto (omzet) result. LJ"),
        ("M_OmzRest", neg("BOB_RJ"), "Omzet restant RJ = -OP.BOB RJ (demo: alleen Q4)"),
        ("M_BorRest", neg("BOR_RJ"), "BOR restant RJ = -OP.BOR RJ"),
        ("M_OmzFY", '=IFERROR([M_OmzYTD]+[M_OmzRest],"n.b.")', "FY omzet RJ = -BOB LJ - OP.BOB RJ"),
        ("M_BorFY", '=IFERROR([M_BorYTD]+[M_BorRest],"n.b.")', "FY BOR RJ = -BOR LJ - OP.BOR RJ"),
        ("M_OmzFY1", neg("BOB_RJ1"), "FY omzet RJ+1 = -OP.BOB RJ+1 (volledig jaar)"),
        ("M_BorFY1", neg("BOR_RJ1"), "FY BOR RJ+1 = -OP.BOR RJ+1"),
        ("M_OmzQ4", neg("BOB_RJ_Q4"), "Omzet Q4 RJ"),
        ("M_BorQ4", neg("BOR_RJ_Q4"), "BOR Q4 RJ"),
    ]
    for q in range(1, 5):
        d.append((f"M_OmzRJ1_Q{q}", neg(f"BOB_RJ1_Q{q}"), f"Omzet RJ+1 Q{q}"))
        d.append((f"M_BorRJ1_Q{q}", neg(f"BOR_RJ1_Q{q}"), f"BOR RJ+1 Q{q}"))
    for o in range(2, 9):
        d.append((f"M_OmzRJ{o}", neg(f"BOB_RJ{o}"), f"Omzet RJ+{o}"))
        d.append((f"M_BorRJ{o}", neg(f"BOR_RJ{o}"), f"BOR RJ+{o}"))
    d += [
        ("M_OmzONV", neg("BOB_ONV"), "Omzet onverdeeld"),
        ("M_BorONV", neg("BOR_ONV"), "BOR onverdeeld"),
        ("C_Kwartaal", '=IFERROR(IF(OR(ABS([BOB_RJ_Q1]+[BOB_RJ_Q2]+[BOB_RJ_Q3]+[BOB_RJ_Q4]-[BOB_RJ])>P_Tol,'
                       'ABS([BOR_RJ_Q1]+[BOR_RJ_Q2]+[BOR_RJ_Q3]+[BOR_RJ_Q4]-[BOR_RJ])>P_Tol,'
                       'ABS([BOB_RJ1_Q1]+[BOB_RJ1_Q2]+[BOB_RJ1_Q3]+[BOB_RJ1_Q4]-[BOB_RJ1])>P_Tol,'
                       'ABS([BOR_RJ1_Q1]+[BOR_RJ1_Q2]+[BOR_RJ1_Q3]+[BOR_RJ1_Q4]-[BOR_RJ1])>P_Tol),1,0),1)',
         "1 = som kwartalen wijkt af van jaartotaal (RJ of RJ+1)"),
        ("C_BobKostBor", '=IFERROR(IF(OR(ABS([BOB_RJ]+[Kosten_RJ]-[BOR_RJ])>P_Tol,ABS([BOB_RJ1]+[Kosten_RJ1]-[BOR_RJ1])>P_Tol,'
                         'ABS([BOB_LJ]+[Kosten_LJ]-[BOR_LJ])>P_Tol),1,0),1)', "1 = OP.BOB + Kosten <> OP.BOR"),
        ("C_Portefeuille", '=IFERROR(IF(ABS([Port_BOB]-([BOB_RJ]+[BOB_RJ1]+[BOB_RJ2]+[BOB_RJ3]+[BOB_RJ4]+[BOB_RJ5]+[BOB_RJ6]+[BOB_RJ7]+[BOB_RJ8]+[BOB_ONV]))>P_Tol,1,0),1)',
         "1 = Portefeuille BOB <> som OP.BOB RJ..RJ+8 + onverdeeld"),
        ("C_PEW", '=IFERROR(IF(ABS([PEW_BOB]-([BOB_VJ]+[BOB_LJ]+[Port_BOB]))>P_Tol,1,0),1)',
         "1 = PEW opbrengst <> BOB VJ + BOB LJ + portefeuille"),
        ("InMapping", f'=IF(COUNTIF(Mapping!$A${MAP0}:$A${MAP1},[WBS])>0,1,0)', "1 = WBS staat in Mapping"),
        ("Dubbel", f'=COUNTIF($B${STG0}:$B${STG1},[WBS])', "Aantal keer dat WBS voorkomt (>1 = dubbel)"),
        ("InEtalage", f'=COUNTIF(STG_Woningetalage!$B${STG0}:$B${STG1},[WBS])', "Aantal keer in Woningetalage"),
        ("NieuwNr", f'=IF([InMapping]=0,COUNTIF($[InMapping:col]${STG0}:[InMapping],0),"")', "Volgnummer niet-gemapt project"),
    ]
    return d


def e_fields():
    f = [
        dict(n="Stufe3", h="Stufe 3", k="key"),
        dict(n="Bedr", h="Bedr.", k="raw"),
        dict(n="Bedrijfsnaam", alias=["Bedrijfsnaam", "Naam maatschappij"], k="txt"),
        dict(n="Aandeel_txt", h="Aandeel Heijm.", k="raw"),
        dict(n="ProjSrt_code", h="Proj.srt", k="raw"),
        dict(n="ProjSrt_oms", h="Proj.srt", k="txt", occ2="ProjSrt_code"),
        dict(n="PC_code", h="Profitctr.", k="raw"),
        dict(n="PC_oms", h="Profitctr.", k="txt", occ2="PC_code"),
        dict(n="Project", h="Project", k="txt"),
        dict(n="Omschrijving", h="Omschrijving", k="txt"),
        dict(n="OmschrStufe3", h="Omschr.Stufe3", k="txt"),
        dict(n="Kans_txt", h="Kans", k="raw"),
        dict(n="StartVerk", h="Start verk.", k="raw"),
        dict(n="StartBouw", h="Start bouw", k="raw"),
        dict(n="EindOpl", h="Eind opl.", k="raw"),
        dict(n="Aantal", h="Aantal woningen", k="num"),
        dict(n="VJ_B", h="t/m vorig jr Beleggers", k="num"),
        dict(n="VJ_P", h="t/m vorig jr Particulier", k="num"),
        dict(n="VJ_T", h="t/m vorig jr Totaal", k="num"),
        dict(n="HJ_B", h="Huidig jr Beleggers", k="num"),
        dict(n="HJ_P", h="Huidig jr Particulier", k="num"),
        dict(n="HJ_T", h="Huidig jr Totaal", k="num"),
        dict(n="Cum_B", h="Cumulatief Beleggers", k="num"),
        dict(n="Cum_P", h="Cumulatief Particulier", k="num"),
        dict(n="Cum_T", h="Cumulatief Totaal", k="num"),
        dict(n="Etalage", h="Etalage Totaal", k="num"),
    ]
    for m in range(1, 13):
        f.append(dict(n=f"M{m:02d}", h=f"hjr {m:02d}", k="num", month=m))
    for q in range(1, 5):
        f.append(dict(n=f"VQ{q}", h=f"Volgend jr Q{q}", k="num"))
    f += [
        dict(n="HuidigJaar", h="Huidig jaar", k="num"),
        dict(n="VolgendJaar", h="Volgend jaar", k="num"),
    ]
    for intern in ["Rond", "Y_RJ2", "Y_RJ3", "Y_RJ4", "Y_RJ5", "Y_RJ6", "Y_NEG1", "Y_NEG2", "Y_NEG3"]:
        f.append(dict(n=intern, ym=intern, k="num"))
    return f


def e_derived():
    return [
        ("Aandeel_num", '=IFERROR(_xlfn.NUMBERVALUE([Aandeel_txt]&"",",","."),"n.b.")', "Aandeel als getal (alleen informatief, NIET toepassen)"),
        ("Kans_num", '=IFERROR(_xlfn.NUMBERVALUE([Kans_txt]&"",",","."),"n.b.")', "Kans als getal"),
        ("HQ1", '=IFERROR([M01]+[M02]+[M03],"n.b.")', "Verkopen RJ Q1 (maanden 1-3)"),
        ("HQ2", '=IFERROR([M04]+[M05]+[M06],"n.b.")', "Verkopen RJ Q2"),
        ("HQ3", '=IFERROR([M07]+[M08]+[M09],"n.b.")', "Verkopen RJ Q3"),
        ("HQ4", '=IFERROR([M10]+[M11]+[M12],"n.b.")', "Verkopen RJ Q4"),
        ("ActYTD", '=IFERROR(SUMIF($[M01:col]$5:$[M12:col]$5,"<="&P_PeilMaand,[M01]:[M12]),"n.b.")', "Verkopen actual t/m peilmaand"),
        ("PrognoseRest", '=IFERROR(SUM([M01]:[M12])-[ActYTD],"n.b.")', "Verkopen prognose rest RJ"),
        ("C_Aantal", '=IFERROR(IF(ABS([Aantal]-([VJ_T]+SUM([M01]:[M12])+[VolgendJaar]+[Y_RJ2]))>P_Tol,1,0),1)', "1 = aantal <> vorig + 12 mnd + volgend + RJ+2"),
        ("C_Etalage", '=IFERROR(IF(ABS([Etalage]-([Aantal]-[Cum_T]))>P_Tol,1,0),1)', "1 = etalage <> aantal - cumulatief"),
        ("C_Cum", '=IFERROR(IF(ABS([Cum_T]-([VJ_T]+[HJ_T]))>P_Tol,1,0),1)', "1 = cumulatief <> t/m vorig jr + huidig jr"),
        ("C_HJ", '=IFERROR(IF(ABS([HJ_T]-[ActYTD])>P_Tol,1,0),1)', "1 = Huidig jr Totaal <> som maanden t/m peilmaand"),
        ("C_BP", '=IFERROR(IF(OR(ABS([VJ_B]+[VJ_P]-[VJ_T])>P_Tol,ABS([HJ_B]+[HJ_P]-[HJ_T])>P_Tol,ABS([Cum_B]+[Cum_P]-[Cum_T])>P_Tol),1,0),1)', "1 = Beleggers + Particulier <> Totaal"),
        ("C_Jaren", '=IFERROR(IF(OR(ABS([HuidigJaar]-SUM([M01]:[M12]))>P_Tol,ABS([VolgendJaar]-([VQ1]+[VQ2]+[VQ3]+[VQ4]))>P_Tol),1,0),1)', "1 = Huidig/Volgend jaar <> som maanden/kwartalen"),
        ("C_NietGeinterp", '=IFERROR(IF(ABS([Rond])+ABS([Y_RJ3])+ABS([Y_RJ4])+ABS([Y_RJ5])+ABS([Y_RJ6])+ABS([Y_NEG1])+ABS([Y_NEG2])+ABS([Y_NEG3])>0,1,0),1)', "1 = niet-geïnterpreteerde kolom bevat waarde"),
        ("InZPMW03", f'=COUNTIF(STG_ZPMW03!$B${STG0}:$B${STG1},[Stufe3])', "Aantal keer in ZPMW03"),
        ("Dubbel", f'=COUNTIF($B${STG0}:$B${STG1},[Stufe3])', "Aantal keer dat Stufe 3 voorkomt"),
        ("C_Projdef", '=IF([InZPMW03]=0,0,IF(INDEX(STG_ZPMW03!$[Z:Projectdefinitie]$%d:$[Z:Projectdefinitie]$%d,MATCH([Stufe3],STG_ZPMW03!$B$%d:$B$%d,0))=[Project]&"",0,1))' % (STG0, STG1, STG0, STG1),
         "1 = Projectdefinitie (ZPMW03) <> Project (Woningetalage) voor dezelfde WBS"),
        ("C_MapAantal", f'=IFERROR(IF(ABS(INDEX(Mapping!$D${MAP0}:$D${MAP1},MATCH([Stufe3],Mapping!$A${MAP0}:$A${MAP1},0))-[Aantal])>P_Tol,1,0),0)', "1 = aantal woningen wijkt af van Mapping"),
        ("C_Meetellen", f'=IFERROR(IF(AND(INDEX(Mapping!$G${MAP0}:$G${MAP1},MATCH([Stufe3],Mapping!$A${MAP0}:$A${MAP1},0))="Nee",ABS([Aantal])+ABS([Cum_T])+ABS([HuidigJaar])>0),1,0),0)', "1 = woningen bij project met 'Woningen meetellen = Nee'"),
        ("InMapping", f'=IF(COUNTIF(Mapping!$A${MAP0}:$A${MAP1},[Stufe3])>0,1,0)', "1 = Stufe 3 staat in Mapping"),
        ("NietInZNr", f'=IF([InZPMW03]=0,COUNTIF($[InZPMW03:col]${STG0}:[InZPMW03],0),"")', "Volgnummer: niet in ZPMW03"),
    ]


def build_stg(wb, ws, src, fields, derived, cols_out):
    base_sheet(ws, zoom=80)
    SRCR = f"{src}!$A$1:$ZZ${N_SRC + 1}"
    put(ws, "A1", "Intern veld", bold=True)
    put(ws, "A2", "Gezochte bronheader", bold=True)
    put(ws, "A3", "Kolomnr in bron", bold=True)
    put(ws, "A4", "Status", bold=True)
    put(ws, "A5", "(hulp)", color=GREY)
    for i, f in enumerate(fields):
        cols_out[f["n"]] = col(2 + i)
    nfields = len(fields)
    for j, (n, _, _) in enumerate(derived):
        cols_out[n] = col(2 + nfields + j)

    def subst(tpl, r):
        out = tpl
        # [X:col] -> kolomletter, [Z:veld] -> kolom in STG_ZPMW03, [X] -> cel
        import re
        out = re.sub(r"\[Z:(\w+)\]", lambda m: ZCOLS[m.group(1)], out)
        out = re.sub(r"\[(\w+):col\]", lambda m: cols_out[m.group(1)], out)
        out = re.sub(r"\[(\w+)\]", lambda m: f"{cols_out[m.group(1)]}{r}", out)
        return out

    for i, f in enumerate(fields):
        c = col(2 + i)
        put(ws, f"{c}1", f["n"], bold=True, fillc=MID)
        if "alias" in f:
            put(ws, f"{c}2", " | ".join(f["alias"]), color=GREY)
            a, b = f["alias"]
            ws[f"{c}3"] = f'=IFERROR(MATCH("{a}",{src}!$1:$1,0),IFERROR(MATCH("{b}",{src}!$1:$1,0),0))'
        elif "occ2" in f:
            put(ws, f"{c}2", f["h"], color=GREY)
            b = cols_out[f["occ2"]]
            ws[f"{c}3"] = (f'=IF({b}3=0,0,IFERROR(MATCH({c}2,INDEX({src}!$1:$1,1,{b}3+1):{src}!$ZZ$1,0)+{b}3,0))')
        elif "ym" in f:
            ws[f"{c}2"] = f'=INDEX(YM_Header,MATCH("{f["ym"]}",YM_Intern,0))'
            ws[f"{c}3"] = f'=IFERROR(MATCH({c}2,{src}!$1:$1,0),0)'
        else:
            put(ws, f"{c}2", f["h"], color=GREY)
            ws[f"{c}3"] = f'=IFERROR(MATCH({c}2,{src}!$1:$1,0),0)'
        ws[f"{c}4"] = f'=IF({c}3=0,"ONTBREEKT","OK")'
        if "month" in f:
            put(ws, f"{c}5", f["month"], color=GREY)
        for rr in (2, 3, 4):
            ws[f"{c}{rr}"].font = font(8, False, GREY)
        for r in range(STG0, STG1 + 1):
            if f["k"] == "key":
                ws[f"{c}{r}"] = (f'=IF({c}$3=0,"",IF(INDEX({SRCR},$A{r},{c}$3)&""="","",INDEX({SRCR},$A{r},{c}$3)&""))')
            else:
                ws[f"{c}{r}"] = f'=IF($B{r}="","",IF({c}$3=0,"n.b.",INDEX({SRCR},$A{r},{c}$3)))'
            if f["k"] == "num":
                ws[f"{c}{r}"].number_format = FMT_RAW
    for j, (n, tpl, desc) in enumerate(derived):
        c = cols_out[n]
        put(ws, f"{c}1", n, bold=True, fillc=YELLOW_LT)
        put(ws, f"{c}2", desc, color=GREY, size=8)
        put(ws, f"{c}4", "afgeleid", color=GREY, size=8)
        for r in range(STG0, STG1 + 1):
            body = subst(tpl, r)
            ws[f"{c}{r}"] = f'=IF($B{r}="","",{body[1:]})'
            ws[f"{c}{r}"].number_format = FMT_RAW
    for r in range(STG0, STG1 + 1):
        ws[f"A{r}"] = r - STG0 + 2
        ws[f"A{r}"].font = font(8, False, GREY)
    ws.freeze_panes = f"C{STG0}"
    ws.column_dimensions["A"].width = 18
    return cols_out


ZCOLS = {}
ECOLS = {}


# ---------------------------------------------------------------------------------------------
# Veldwoordenboek incl. expliciete mapping van jaarheaders in Woningetalage
# ---------------------------------------------------------------------------------------------
def build_veldwoordenboek(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, "Veldwoordenboek", "Interne unieke veldnamen per bron, hoe ze gevonden worden en de expliciete interpretatie van jaarheaders.", 8)
    widths(ws, {"B": 22, "C": 30, "D": 14, "E": 12, "F": 70, "G": 14})
    r = 4
    section(ws, r, 2, "A. Jaarheaders Woningetalage (expliciete mapping — wijzig hier, niet in formules)", 6)
    r += 1
    header(ws, r, 2, ["Intern veld", "Bronheader (waarde)", "Gevonden in bron?", "Gebruikt?", "Interpretatie", "Status"], height=22)
    ym = [
        ("Y_RJ2", "=P_RJ+2", "Ja", "Verkopen in jaar RJ+2 (prototype: 2028). Onderdeel van 'Aantal woningen = vorig jaar + 12 mnd + volgend jaar + RJ+2'.", "Bevestigd (demo)"),
        ("Y_RJ3", "=P_RJ+3", "Nee", "Vermoedelijk verkopen RJ+3. Alleen gecontroleerd op 0.", "Aanname"),
        ("Y_RJ4", "=P_RJ+4", "Nee", "Vermoedelijk verkopen RJ+4. Alleen gecontroleerd op 0.", "Aanname"),
        ("Y_RJ5", "=P_RJ+5", "Nee", "Vermoedelijk verkopen RJ+5. Alleen gecontroleerd op 0.", "Aanname"),
        ("Y_RJ6", "=P_RJ+6", "Nee", "Vermoedelijk verkopen RJ+6. Alleen gecontroleerd op 0.", "Aanname"),
        ("Y_NEG1", "=-(P_RJ+7)", "Nee", "Negatieve jaarheader. Betekenis onbekend — NIET geïnterpreteerd; waarde <> 0 geeft waarschuwing.", "Niet geïnterpreteerd"),
        ("Y_NEG2", "=-(P_RJ+8)", "Nee", "Idem.", "Niet geïnterpreteerd"),
        ("Y_NEG3", "=-(P_RJ+9)", "Nee", "Idem.", "Niet geïnterpreteerd"),
        ("Rond", "Rond", "Nee", "Betekenis onbekend — NIET geïnterpreteerd; waarde <> 0 geeft waarschuwing.", "Niet geïnterpreteerd"),
    ]
    y0 = r + 1
    for i, (n, h, used, interp, st) in enumerate(ym):
        rr = y0 + i
        put(ws, f"B{rr}", n)
        c = put(ws, f"C{rr}", h, align="left")
        input_style(c)
        ws[f"D{rr}"] = f'=IF(ISNUMBER(MATCH(C{rr},SRC_Woningetalage!$1:$1,0)),"Ja","NEE")'
        put(ws, f"E{rr}", used)
        put(ws, f"F{rr}", interp, wrap=True)
        put(ws, f"G{rr}", st)
    add_name(wb, "YM_Intern", f"Veldwoordenboek!$B${y0}:$B${y0+len(ym)-1}")
    add_name(wb, "YM_Header", f"Veldwoordenboek!$C${y0}:$C${y0+len(ym)-1}")
    r = y0 + len(ym) + 1

    for label, sheet, fields, cols in (("B. ZPMW03", "STG_ZPMW03", z_fields(), ZCOLS),
                                        ("C. Woningetalage", "STG_Woningetalage", e_fields(), ECOLS)):
        section(ws, r, 2, f"{label}: velden en vindplaats (gezocht op header, kolomvolgorde in bron mag wijzigen)", 6)
        r += 1
        header(ws, r, 2, ["Intern veld", "Gezochte header", "Kolomnr", "Status", "Methode", ""], height=22)
        for f in fields:
            r += 1
            c = cols[f["n"]]
            put(ws, f"B{r}", f["n"])
            ws[f"C{r}"] = f"={sheet}!{c}2"
            ws[f"D{r}"] = f"={sheet}!{c}3"
            ws[f"E{r}"] = f"={sheet}!{c}4"
            if "alias" in f:
                m = "Alias: eerste die bestaat van " + " / ".join(f["alias"])
            elif "occ2" in f:
                m = "2e voorkomen van dubbele header (omschrijving); 1e voorkomen = code"
            elif f["n"] in ("ProjSrt_code", "PC_code"):
                m = "1e voorkomen van dubbele header (code)"
            elif "ym" in f:
                m = "Via jaarheader-mapping (tabel A)"
            elif f["k"] == "key":
                m = "Koppelsleutel"
            else:
                m = "Exacte header"
            put(ws, f"F{r}", m, color=GREY)
        r += 2
    ws.conditional_formatting.add(f"E5:E{r}", FormulaRule(formula=['E5="ONTBREEKT"'], fill=fill(ORANGE_LT), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"D{y0}:D{y0+len(ym)-1}", FormulaRule(formula=[f'D{y0}="NEE"'], font=Font(color=RED, bold=True)))


# ---------------------------------------------------------------------------------------------
# Mapping
# ---------------------------------------------------------------------------------------------
def build_mapping(wb, ws, hist):
    base_sheet(ws, zoom=100)
    title(ws, "Projectmapping", "Eén regel per WBS-element. Koppelsleutel = WBS-element (ZPMW03) = Stufe 3 (Woningetalage). Nooit op omschrijving koppelen.", 9)
    hdr = ["WBS-element", "Projectdefinitie", "Projectnaam", "Aantal woningen op aandeelbasis", "Aandeel procent",
           "Type", "Woningen meetellen", "Segment demo"]
    header(ws, 4, 1, hdr, height=32)
    rows = list(hist["Projectmapping"].iter_rows(min_row=2, values_only=True))
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.cell(row=MAP0 + i, column=1 + j, value=v)
            input_style(c)
    last = MAP0 + len(rows) - 1
    t = Table(displayName="T_Projectmapping", ref=f"A4:H{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    widths(ws, {"A": 18, "B": 15, "C": 30, "D": 14, "E": 10, "F": 12, "G": 11, "H": 15, "I": 4, "J": 70})
    ws.column_dimensions["A"].width = 18
    notes = [
        "Bron: RDFO_FICTIEVE_historie_en_mapping.xlsx (tab Projectmapping).",
        "Nieuw project in ZPMW03? Voeg hier een regel toe (de tabel groeit mee; formules lopen t/m rij 104).",
        "Type (Opstal/Deelneming) en Segment zijn mappingdimensies — niet afgeleid uit IFRS J/N.",
        "'Aandeel procent' is informatief. Bedragen en woningaantallen zijn in de demo al op aandeelbasis: NIET opnieuw toepassen.",
        "'Segment demo' is een fictieve demo-indeling, geen broninformatie.",
        "Historische vastleggingen bewaren hun eigen projectnaam/type (kolommen L en M in Vastlegging), dus wijzigingen hier herschrijven de historie niet.",
    ]
    put(ws, "J4", "Toelichting", bold=True, color=NAVY)
    for i, n in enumerate(notes):
        put(ws, f"J{5+i}", "• " + n, color=GREY, wrap=True)
    dv_list(ws, f"F{MAP0}:F{MAP1}", '"Opstal,Deelneming"')
    dv_list(ws, f"G{MAP0}:G{MAP1}", '"Ja,Nee"')
    for c, n in zip("ABCDEFGH", ["WBS", "Projdef", "Naam", "Aantal", "Aandeel", "Type", "Meetellen", "Segment"]):
        add_name(wb, f"Map_{n}", f"Mapping!${c}${MAP0}:${c}${MAP1}")
    ws.freeze_panes = "A5"
    return {row[0]: row for row in rows}


# ---------------------------------------------------------------------------------------------
# Vastlegging: onveranderlijk grootboek met versies, revisies en audittrail
# ---------------------------------------------------------------------------------------------
LCOLS = ["ImportID", "Bronbestand", "Importdatum", "Status", "Versiejaar", "Versie", "Revisie", "Peildatum",
         "Doeljaar", "WBS", "Projectdefinitie", "Projectnaam", "Type", "IFRS", "Kengetal", "Waarde", "Eenheid",
         "Toelichting"]
LHELP = ["Sleutel", "MaxRev", "Actief", "Dubbel"]


def hist_rows(hist, mapping, ifrs):
    out = []
    ph = list(hist["Prognosehistorie"].iter_rows(min_row=2, values_only=True))
    wh = list(hist["Woninghistorie"].iter_rows(min_row=2, values_only=True))
    for versie in VERSIES[:4]:
        for (wbs, naam, vj, v, peil, doel, omz, bor, omza, bora, herk) in ph:
            if v != versie:
                continue
            m = mapping[wbs]
            base = [f"HIST-{vj}-{v}", HIST_FILE, IMPORTDATUM, "Vastgesteld", vj, v, 0, peil, doel, wbs, m[1], naam,
                    m[5], ifrs.get(wbs, "n.b.")]
            for k, val in ((K_OMZ, omz), (K_BOR, bor), (K_OMZA, omza), (K_BORA, bora)):
                if val is not None:
                    out.append(base + [k, val, "EUR", herk])
        for (wbs, naam, vj, v, peil, doel, vorig, fy, q1, q2, q3, q4, act, herk) in wh:
            if v != versie:
                continue
            m = mapping[wbs]
            base = [f"HIST-{vj}-{v}", HIST_FILE, IMPORTDATUM, "Vastgesteld", vj, v, 0, peil, doel, wbs, m[1], naam,
                    m[5], ifrs.get(wbs, "n.b.")]
            for k, val in ((K_WFY, fy), (K_WVJ, vorig), (K_WQ[0], q1), (K_WQ[1], q2), (K_WQ[2], q3), (K_WQ[3], q4), (K_WACT, act)):
                if val is not None:
                    out.append(base + [k, val, "aantal", herk])
    return out


def build_ledger(wb, ws, rows):
    base_sheet(ws, zoom=85)
    title(ws, "Vastlegging — versiegrootboek (alleen toevoegen, nooit overschrijven)",
          "Plak nieuwe regels als WAARDEN onder de laatste regel (zie Vastleggen). Bestaande regels niet wijzigen: correcties = nieuwe regel met hogere revisie + toelichting.", 22)
    put(ws, "B3", "Eerste vrije rij:", bold=True)
    ws["D3"] = f'=SUMPRODUCT(MAX((L_WBS<>"")*ROW(L_WBS)))+1'
    ws["D3"].font = font(10, True, NAVY)
    put(ws, "F3", "Actieve regels:", bold=True)
    ws["H3"] = "=SUM(L_Actief)"
    put(ws, "J3", "Genegeerde dubbele regels:", bold=True)
    ws["M3"] = "=SUM(L_Dubbel)"
    header(ws, 5, 1, LCOLS + LHELP, height=30)
    for c in range(len(LCOLS) + 1, len(LCOLS) + len(LHELP) + 1):
        ws.cell(row=5, column=c).fill = fill(GREY)
    for i, row in enumerate(rows):
        r = L0 + i
        for j, v in enumerate(row):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.font = font(8)
            if j in (2, 7):
                c.number_format = FMT_DATE
            if j == 15:
                c.number_format = FMT_RAW
    for r in range(L0, L1 + 1):
        ws[f"S{r}"] = f'=IF(J{r}="","",J{r}&"|"&E{r}&"|"&F{r}&"|"&I{r}&"|"&O{r})'
        ws[f"T{r}"] = f'=IF(OR(S{r}="",D{r}<>"Vastgesteld"),"",_xlfn.MAXIFS($G${L0}:$G${L1},$S${L0}:$S${L1},S{r},$D${L0}:$D${L1},"Vastgesteld"))'
        ws[f"U{r}"] = f'=IF(T{r}="",0,IF(AND(G{r}=T{r},COUNTIFS($S${L0}:S{r},S{r},$G${L0}:G{r},G{r},$D${L0}:D{r},"Vastgesteld")=1),1,0))'
        ws[f"V{r}"] = f'=IF(T{r}="",0,IF(AND(G{r}=T{r},COUNTIFS($S${L0}:S{r},S{r},$G${L0}:G{r},G{r},$D${L0}:D{r},"Vastgesteld")>1),1,0))'
        for c in "STUV":
            ws[f"{c}{r}"].font = font(8, False, GREY)
    for j, n in enumerate(LCOLS + LHELP):
        add_name(wb, f"L_{n}", f"Vastlegging!${col(1+j)}${L0}:${col(1+j)}${L1}")
    widths(ws, {"A": 16, "B": 22, "C": 10, "D": 11, "E": 8, "F": 7, "G": 6, "H": 10, "I": 7, "J": 16, "K": 10,
                "L": 24, "M": 10, "N": 5, "O": 24, "P": 13, "Q": 7, "R": 40, "S": 30, "T": 6, "U": 6, "V": 6})
    ws.auto_filter.ref = f"A5:V{L0 + len(rows) - 1}"
    ws.freeze_panes = f"A{L0}"


# ---------------------------------------------------------------------------------------------
# Aanvullende invoer, toelichtingen, koperstype
# ---------------------------------------------------------------------------------------------
def build_aanvullend(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, "Aanvullende invoer RDFO", "Posten die NIET uit ZPMW03/Woningetalage komen. Leeg = 'n.b.' in de rapportage (nooit stil nul). Bedragen in EUR, managementteken (+ verhoogt omzet/resultaat).", 15)
    hdr = ["Doeljaar", "Versiejaar", "Versie", "Niveau", "WBS-element", "Rubriek", "Soort", "Waarde", "Eenheid",
           "Status", "Toelichting", "Bron / eigenaar", "Telt vastgesteld", "Telt concept", "Conflict regio/project"]
    header(ws, 5, 1, hdr, height=32)
    for c in (13, 14, 15):
        ws.cell(row=5, column=c).fill = fill(GREY)
    demo = [
        (2026, 2026, "Q3", "Regio", None, "Taakstelling omzet", "FY", 2000000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Interne eliminaties", "FY", -350000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "AK-dekking", "FY", -1400000, "EUR", "Fictieve aanname", "DEMO-scenario: vast bedrag, GEEN generiek AK-percentage", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Fees", "FY", 300000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Taakstelling EBIT", "FY", 150000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Project", "G.900004.8.9090", "IFRS-eliminaties", "FY", -100000, "EUR", "Voorbeeld", "VOORBEELDREGEL (formaat) — telt nooit mee", "Voorbeeld"),
    ]
    for i, row in enumerate(demo):
        for j, v in enumerate(row):
            c = ws.cell(row=A0 + i, column=1 + j, value=v)
            input_style(c, FMT_EUR if j == 7 else None)
    for r in range(A0, A1 + 1):
        ws[f"M{r}"] = f'=IF(H{r}="","",IF(OR(J{r}="Vastgesteld",AND(J{r}="Fictieve aanname",P_Demo="Ja")),1,0))'
        ws[f"N{r}"] = f'=IF(H{r}="","",IF(OR(J{r}="Vastgesteld",J{r}="Concept",AND(J{r}="Fictieve aanname",P_Demo="Ja")),1,0))'
        ws[f"O{r}"] = (f'=IF(H{r}="","",IF(COUNTIFS($A${A0}:$A${A1},A{r},$B${A0}:$B${A1},B{r},$C${A0}:$C${A1},C{r},'
                       f'$F${A0}:$F${A1},F{r},$G${A0}:$G${A1},G{r},$D${A0}:$D${A1},"<>"&D{r},$H${A0}:$H${A1},"<>")>0,1,0))')
        for c in "MNO":
            helper_style(ws[f"{c}{r}"])
    t = Table(displayName="T_Aanvullend", ref=f"A5:O{A0 + len(demo) - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    for c, n in zip("ABCDEFGHIJKMNO", ["Doel", "VJ", "Versie", "Niveau", "WBS", "Rubriek", "Soort", "Waarde", "Eenheid",
                                        "Status", "Toel", "TeltVast", "TeltConcept", "Conflict"]):
        add_name(wb, f"A_{n}", f"Invoer_Aanvullend!${c}${A0}:${c}${A1}")
    dv_list(ws, f"C{A0}:C{A1}", "V_List")
    dv_list(ws, f"D{A0}:D{A1}", "L_Niveau")
    dv_list(ws, f"F{A0}:F{A1}", "L_Rubrieken")
    dv_list(ws, f"G{A0}:G{A1}", "L_Soort")
    dv_list(ws, f"J{A0}:J{A1}", "L_AanvStatus")
    widths(ws, {"A": 8, "B": 9, "C": 7, "D": 8, "E": 16, "F": 20, "G": 10, "H": 13, "I": 7, "J": 15, "K": 48, "L": 14, "M": 9, "N": 9, "O": 10})
    put(ws, "B3", "Legenda status: Vastgesteld = telt in vastgelegde versiekolommen · Concept = telt alleen in kolom 'Upload (concept)' · "
        "Fictieve aanname = telt alleen als Parameters › P_Demo = Ja · Voorbeeld = telt nooit.", color=GREY, italic=True)
    ws.freeze_panes = f"A{A0}"


def build_toelichting(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, "Toelichtingen op verschillen", "Eén toelichting per project, doeljaar en versie. Verschijnt in Projecten en Projectdetail.", 9)
    hdr = ["Doeljaar", "Versiejaar", "Versie", "WBS-element", "Kengetal", "Toelichting", "Auteur", "Datum", "Sleutel"]
    header(ws, 5, 1, hdr, height=24)
    ws.cell(row=5, column=9).fill = fill(GREY)
    ex = [(2026, 2026, "Q3", "G.900008.8.9090", K_BOR, "[VOORBEELD – fictief] Verliesproject: voorbeeldtekst om het formaat te tonen.", "Demo", IMPORTDATUM)]
    for i, row in enumerate(ex):
        for j, v in enumerate(row):
            c = ws.cell(row=T0 + i, column=1 + j, value=v)
            input_style(c, FMT_DATE if j == 7 else None)
    for r in range(T0, T1 + 1):
        ws[f"I{r}"] = f'=IF(D{r}="","",D{r}&"|"&A{r}&"|"&B{r}&"|"&C{r})'
        helper_style(ws[f"I{r}"])
    t = Table(displayName="T_Toelichting", ref=f"A5:I{T0 + len(ex) - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    add_name(wb, "T_Key", f"Invoer_Toelichting!$I${T0}:$I${T1}")
    add_name(wb, "T_Txt", f"Invoer_Toelichting!$F${T0}:$F${T1}")
    dv_list(ws, f"C{T0}:C{T1}", "V_List")
    widths(ws, {"A": 8, "B": 9, "C": 7, "D": 16, "E": 16, "F": 70, "G": 12, "H": 11, "I": 30})
    ws.freeze_panes = f"A{T0}"


def build_koperstype(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, "Aanvullende invoer koperstype (toekomstige verkopen)", "Woningetalage kent Beleggers/Particulier alleen voor actuals. Prognoses zonder invoer hier blijven 'Onbekend'.", 9)
    hdr = ["WBS-element", "Doeljaar", "Periode", "Koperstype", "Aantal", "Status", "Toelichting", "Telt", "Type (mapping)"]
    header(ws, 5, 1, hdr, height=24)
    ws.cell(row=5, column=8).fill = fill(GREY)
    ws.cell(row=5, column=9).fill = fill(GREY)
    ex = [("G.900005.8.9090", 2026, "Q4", "Belegger", 4, "Voorbeeld", "VOORBEELDREGEL (formaat) — telt nooit mee")]
    for i, row in enumerate(ex):
        for j, v in enumerate(row):
            input_style(ws.cell(row=K0 + i, column=1 + j, value=v))
    for r in range(K0, K1 + 1):
        ws[f"H{r}"] = f'=IF(E{r}="","",IF(OR(F{r}="Vastgesteld",F{r}="Concept"),1,0))'
        ws[f"I{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(Map_Type,MATCH(A{r},Map_WBS,0)),"n.b."))'
        helper_style(ws[f"H{r}"])
        helper_style(ws[f"I{r}"])
    t = Table(displayName="T_Koperstype", ref=f"A5:I{K0 + len(ex) - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    for c, n in zip("ABCDEFHI", ["WBS", "Doel", "Periode", "Type", "Aantal", "Status", "Telt", "TypeMap"]):
        add_name(wb, f"KT_{n}", f"Invoer_Koperstype!${c}${K0}:${c}${K1}")
    dv_list(ws, f"C{K0}:C{K1}", '"Q1,Q2,Q3,Q4"')
    dv_list(ws, f"D{K0}:D{K1}", '"Belegger,Particulier"')
    dv_list(ws, f"F{K0}:F{K1}", '"Vastgesteld,Concept,Voorbeeld"')
    widths(ws, {"A": 16, "B": 8, "C": 8, "D": 12, "E": 8, "F": 12, "G": 50, "H": 6})


def build_importlog(wb, ws, hist_counts, nz, ne):
    base_sheet(ws, zoom=90)
    title(ws, "Importlog", "Elke inlees- of vastlegronde krijgt een Import-ID. 'Regels vastgelegd' telt live in Vastlegging.", 12)
    hdr = ["Import-ID", "Bron", "Bestandsnaam", "Importdatum", "Peildatum", "Versiejaar", "Versie",
           "Regels in bronbestand", "Regels vastgelegd (live)", "Status", "Toelichting"]
    header(ws, 5, 1, hdr, height=30)
    rows = []
    for v, peil in zip(VERSIES[:4], [dt.datetime(2025, 12, 31), dt.datetime(2026, 3, 31), dt.datetime(2026, 6, 30), PEIL]):
        rows.append((f"HIST-2026-{v}", "Historie (fictief)", HIST_FILE, IMPORTDATUM, peil, 2026, v, hist_counts[v],
                     "Vastgesteld", "Synthetische snapshot uit testhistorie; geen echte historie. Prognose- en woninghistorie samen."))
    rows.append(("IMP-2026-Q3-01", "ZPMW03", SRC_Z_FILE, IMPORTDATUM, PEIL, 2026, "Q3", nz,
                 "Ingelezen", "Huidige bron in SRC_ZPMW03. Q3 is al vastgesteld via historie; concept dient als aansluitcontrole (niet opnieuw vastgelegd)."))
    rows.append(("IMP-2026-Q3-01", "Woningetalage", SRC_E_FILE, IMPORTDATUM, PEIL, 2026, "Q3", ne,
                 "Ingelezen", "Huidige bron in SRC_Woningetalage."))
    for i, row in enumerate(rows):
        r = 6 + i
        for j, v in enumerate(row[:8]):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.font = font(9)
            if j in (3, 4):
                c.number_format = FMT_DATE
        ws.cell(row=r, column=9, value=f'=COUNTIF(L_ImportID,A{r})')
        ws.cell(row=r, column=10, value=row[8])
        ws.cell(row=r, column=11, value=row[9]).alignment = Alignment(wrap_text=True, vertical="top")
    t = Table(displayName="T_Importlog", ref=f"A5:K{5 + len(rows)}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    widths(ws, {"A": 17, "B": 16, "C": 38, "D": 11, "E": 11, "F": 9, "G": 7, "H": 11, "I": 11, "J": 12, "K": 70})
    put(ws, "B3", "Na elke vastlegging: voeg hier een regel toe met hetzelfde Import-ID als in Parameters › P_ImportID.", color=GREY, italic=True)


# ---------------------------------------------------------------------------------------------
# VERW_Project: één regel per gemapt WBS-element; elke bron eerst per WBS geaggregeerd (geen many-to-many)
# ---------------------------------------------------------------------------------------------
VCOLS = {}


def verw_columns():
    zr = lambda f: f"STG_ZPMW03!${ZCOLS[f]}${STG0}:${ZCOLS[f]}${STG1}"  # noqa: E731
    er = lambda f: f"STG_Woningetalage!${ECOLS[f]}${STG0}:${ECOLS[f]}${STG1}"  # noqa: E731
    zkey = f"STG_ZPMW03!$B${STG0}:$B${STG1}"
    ekey = f"STG_Woningetalage!$B${STG0}:$B${STG1}"
    c = []
    m = lambda n, mc, lab: c.append((n, lab, "=IF(Mapping!$%s{mr}=\"\",\"\",Mapping!$%s{mr})" % (mc, mc), None))  # noqa: E731
    m("WBS", "A", "WBS-element")
    m("Projdef", "B", "Projectdefinitie")
    m("Naam", "C", "Projectnaam")
    m("Type", "F", "Type")
    m("Segment", "H", "Segment (demo)")
    m("Aandeel", "E", "Aandeel % (info)")
    m("Meetellen", "G", "Woningen meetellen")
    m("MapAantal", "D", "Aantal woningen (mapping)")
    c.append(("IFRS", "IFRS J/N", f'=IF($A{{r}}="","",IFERROR(INDEX({zr("IFRS")},MATCH($A{{r}},{zkey},0)),"n.b."))', None))
    c.append(("nZ", "# regels ZPMW03", f'=IF($A{{r}}="","",COUNTIF({zkey},$A{{r}}))', "0"))
    c.append(("nE", "# regels Woningetalage", f'=IF($A{{r}}="","",COUNTIF({ekey},$A{{r}}))', "0"))
    money = [("OmzYTD", "M_OmzYTD", "Omzet actual YTD"), ("BorYTD", "M_BorYTD", "BOR actual YTD"),
             ("OmzRest", "M_OmzRest", "Omzet restant RJ"), ("BorRest", "M_BorRest", "BOR restant RJ"),
             ("OmzFY", "M_OmzFY", "Omzet FY RJ"), ("BorFY", "M_BorFY", "BOR FY RJ"),
             ("OmzFY1", "M_OmzFY1", "Omzet FY RJ+1"), ("BorFY1", "M_BorFY1", "BOR FY RJ+1"),
             ("OmzQ4", "M_OmzQ4", "Omzet Q4 RJ"), ("BorQ4", "M_BorQ4", "BOR Q4 RJ")]
    money += [(f"OmzRJ1_Q{q}", f"M_OmzRJ1_Q{q}", f"Omzet RJ+1 Q{q}") for q in range(1, 5)]
    money += [(f"BorRJ1_Q{q}", f"M_BorRJ1_Q{q}", f"BOR RJ+1 Q{q}") for q in range(1, 5)]
    money += [(f"OmzRJ{o}", f"M_OmzRJ{o}", f"Omzet RJ+{o}") for o in range(2, 9)] + [("OmzONV", "M_OmzONV", "Omzet onverdeeld")]
    money += [(f"BorRJ{o}", f"M_BorRJ{o}", f"BOR RJ+{o}") for o in range(2, 9)] + [("BorONV", "M_BorONV", "BOR onverdeeld")]
    for n, f, lab in money:
        c.append((n, lab, f'=IF($A{{r}}="","",IF($J{{r}}=0,"n.b.",SUMIFS({zr(f)},{zkey},$A{{r}})))', FMT_EUR))
    won = ["Aantal", "VJ_B", "VJ_P", "VJ_T", "HJ_B", "HJ_P", "HJ_T", "Cum_B", "Cum_P", "Cum_T", "Etalage"] + \
          [f"M{i:02d}" for i in range(1, 13)] + ["HQ1", "HQ2", "HQ3", "HQ4", "ActYTD", "PrognoseRest", "HuidigJaar",
                                                  "VolgendJaar", "VQ1", "VQ2", "VQ3", "VQ4", "Y_RJ2"]
    for n in won:
        c.append(("W_" + n, "Woningen " + n, f'=IF($A{{r}}="","",IF($K{{r}}=0,"n.b.",SUMIFS({er(n)},{ekey},$A{{r}})))', FMT_INT))
    for n, lab in [("Kans", "Kans% (ZPMW03)"), ("StatusWBS", "Status WBS"), ("StartBouw", "Start bouw"),
                   ("Oplevering", "Maand oplevering"), ("Risico", "Risicoklasse"), ("POC", "POC %"),
                   ("Orderintake", "Orderintake"), ("Marktsegment", "Marktsegment"), ("Bedrijfsnaam", "Bedrijfsnaam"),
                   ("ProfitCenterNaam", "Profitcenter")]:
        c.append(("I_" + n, lab, f'=IF($A{{r}}="","",IFERROR(INDEX({zr(n)},MATCH($A{{r}},{zkey},0)),"n.b."))', None))
    for n, lab in [("StartVerk", "Start verkoop"), ("EindOpl", "Eind oplevering")]:
        c.append(("I_" + n, lab, f'=IF($A{{r}}="","",IFERROR(INDEX({er(n)},MATCH($A{{r}},{ekey},0)),"n.b."))', None))
    c.append(("W_Bruto", "Woningen bruto (afgeleid)", '=IF($A{r}="","",IF(AND(ISNUMBER([W_Aantal]),ISNUMBER([Aandeel])),IF([Aandeel]>0,[W_Aantal]/([Aandeel]/100),"n.b."),"n.b."))', "#,##0.0"))
    c.append(("Verlies", "Verlies (BOR FY RJ < 0)", '=IF($A{r}="","",IF(ISNUMBER([BorFY]),IF([BorFY]<0,1,0),0))', "0"))
    c.append(("VerliesNr", "Volgnr verlies", f'=IF([Verlies]=1,COUNTIF($[Verlies:col]${V0}:[Verlies],1),"")', "0"))
    c.append(("VervallenNr", "Volgnr ontbreekt in ZPMW03", f'=IF(AND($A{{r}}<>"",$J{{r}}=0),COUNTIFS($A${V0}:$A{{r}},"?*",$J${V0}:$J{{r}},0),"")', "0"))
    c.append(("BorPct", "BOR % FY RJ", '=IF($A{r}="","",IF(AND(ISNUMBER([BorFY]),ISNUMBER([OmzFY])),IF([OmzFY]<>0,[BorFY]/[OmzFY],"n.b."),"n.b."))', FMT_PCT))
    return c


def build_verw(wb, ws):
    import re
    base_sheet(ws, zoom=80)
    put(ws, "A1", "VERW_Project — berekeningslaag (niet invullen). Eén regel per mappingregel; bronnen per WBS geaggregeerd; bedragen in EUR met managementteken (omzet/winst positief, verlies negatief).", bold=True, color=NAVY)
    put(ws, "A2", "RJ = Parameters › P_RJ. 'n.b.' = project ontbreekt in de betreffende bron (zie Controle).", color=GREY, italic=True)
    cols = verw_columns()
    for i, (n, *_r) in enumerate(cols):
        VCOLS[n] = col(1 + i)
    assert VCOLS["nZ"] == "J" and VCOLS["nE"] == "K"
    for i, (n, lab, tpl, fmt) in enumerate(cols):
        cl = col(1 + i)
        put(ws, f"{cl}4", n, color=GREY, size=8)
        h = ws.cell(row=5, column=1 + i, value=lab)
        h.fill = fill(NAVY)
        h.font = font(8, True, WHITE)
        h.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        for r in range(V0, V1 + 1):
            mr = MAP0 + (r - V0)
            f = tpl.replace("{mr}", str(mr)).replace("{r}", str(r))
            f = re.sub(r"\[(\w+):col\]", lambda mm: VCOLS[mm.group(1)], f)
            f = re.sub(r"\[(\w+)\]", lambda mm: f"{VCOLS[mm.group(1)]}{r}", f)
            cell = ws.cell(row=r, column=1 + i, value=f)
            cell.font = font(8)
            if fmt:
                cell.number_format = fmt
        add_name(wb, f"V_{n}", f"VERW_Project!${cl}${V0}:${cl}${V1}")
        ws.column_dimensions[cl].width = 11
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["C"].width = 26
    ws.row_dimensions[5].height = 42
    ws.freeze_panes = f"D{V0}"


# ---------------------------------------------------------------------------------------------
# CALC_Versies: totalen per versie (regio en gefilterd), geselecteerde en vorige versie
# ---------------------------------------------------------------------------------------------
CC = {}   # metric -> kolom in CALC_Versies
CR = {"Budget": 5, "Q1": 6, "Q2": 7, "Q3": 8, "Q4": 9, "Upload": 10, "Sel": 11, "Vorige": 12}


def build_calc(wb, ws):
    base_sheet(ws, zoom=80)
    put(ws, "A1", "CALC_Versies — totalen per versie (niet invullen). Versierijen lezen uitsluitend actieve, vastgestelde regels uit Vastlegging; 'Upload' leest de huidige bron.", bold=True, color=NAVY)
    put(ws, "A2", "Selectie komt van het Dashboard (doeljaar, versie, type, project). Regio-kolommen negeren type/projectfilter; gefilterde kolommen niet.", color=GREY, italic=True)
    led_base = 'L_Actief,1,L_Doeljaar,Sel_Doeljaar,L_Versiejaar,P_VJ,L_Versie,$A{r}'
    metrics = [
        # naam, label, kengetal, extra criteria ledger, upload-formule
        ("OmzOps", "Omzet opstal FY", K_OMZ, ',L_Type,"Opstal"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_OmzFY,V_Type,"Opstal"),SUMIFS(V_OmzFY1,V_Type,"Opstal"))', "C"),
        ("OmzDln", "Omzet deelnemingen FY", K_OMZ, ',L_Type,"Deelneming"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_OmzFY,V_Type,"Deelneming"),SUMIFS(V_OmzFY1,V_Type,"Deelneming"))', "C"),
        ("BorOps", "BOR opstal FY", K_BOR, ',L_Type,"Opstal"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_BorFY,V_Type,"Opstal"),SUMIFS(V_BorFY1,V_Type,"Opstal"))', "C"),
        ("BorDln", "BOR deelnemingen FY", K_BOR, ',L_Type,"Deelneming"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_BorFY,V_Type,"Deelneming"),SUMIFS(V_BorFY1,V_Type,"Deelneming"))', "C"),
        ("ActOmzOps", "Omzet opstal actual YTD", K_OMZA, ',L_Type,"Opstal"', 'SUMIFS(V_OmzYTD,V_Type,"Opstal")', "D"),
        ("ActOmzDln", "Omzet deelnemingen actual YTD", K_OMZA, ',L_Type,"Deelneming"', 'SUMIFS(V_OmzYTD,V_Type,"Deelneming")', "D"),
        ("ActBorOps", "BOR opstal actual YTD", K_BORA, ',L_Type,"Opstal"', 'SUMIFS(V_BorYTD,V_Type,"Opstal")', "D"),
        ("ActBorDln", "BOR deelnemingen actual YTD", K_BORA, ',L_Type,"Deelneming"', 'SUMIFS(V_BorYTD,V_Type,"Deelneming")', "D"),
        ("OmzIFRSN", "Omzet projecten IFRS=N (memo)", K_OMZ, ',L_IFRS,"N"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_OmzFY,V_IFRS,"N"),SUMIFS(V_OmzFY1,V_IFRS,"N"))', "C"),
        ("BorIFRSN", "BOR projecten IFRS=N (memo)", K_BOR, ',L_IFRS,"N"', 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_BorFY,V_IFRS,"N"),SUMIFS(V_BorFY1,V_IFRS,"N"))', "C"),
        ("WonFY", "Woningverkopen FY", K_WFY, "", 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_W_HuidigJaar,V_Meetellen,"Ja"),SUMIFS(V_W_VolgendJaar,V_Meetellen,"Ja"))', "E"),
        ("WonAct", "Woningverkopen actual YTD", K_WACT, "", 'SUMIFS(V_W_ActYTD,V_Meetellen,"Ja")', "W"),
        # gefilterd (dashboard)
        ("F_OmzFY", "Omzet FY (filter)", K_OMZ, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_OmzFY,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS),SUMIFS(V_OmzFY1,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS))', "F"),
        ("F_BorFY", "BOR FY (filter)", K_BOR, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_BorFY,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS),SUMIFS(V_BorFY1,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS))', "F"),
        ("F_ActOmz", "Omzet actual YTD (filter)", K_OMZA, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'SUMIFS(V_OmzYTD,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)', "FA"),
        ("F_ActBor", "BOR actual YTD (filter)", K_BORA, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'SUMIFS(V_BorYTD,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)', "FA"),
        ("F_WonFY", "Woningverkopen FY (filter)", K_WFY, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'IF(Sel_Doeljaar=P_RJ,SUMIFS(V_W_HuidigJaar,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS),SUMIFS(V_W_VolgendJaar,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS))', "FW"),
        ("F_WonAct", "Woningverkopen actual YTD (filter)", K_WACT, ",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS", 'SUMIFS(V_W_ActYTD,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)', "FWA"),
    ]
    heads = ["Versie", "Bron", "# regels omzet FY", "# regels actual", "# regels woningen", "# correcties (rev>0)"]
    for i, h in enumerate(heads):
        CC[h] = col(1 + i)
    c0 = len(heads) + 1
    allm = [(m[0], m[1]) for m in metrics] + [("AV_" + rb, rb + " FY (aanvullend)") for rb in RUBRIEKEN] + \
           [("AA_" + rb, rb + " actual YTD (aanvullend)") for rb in RUBRIEKEN]
    for i, (n, _l) in enumerate(allm):
        CC[n] = col(c0 + i)
    header(ws, 4, 1, heads + [lab for _n, lab in allm], height=44)
    actual_kinds = {"D", "W", "FA", "FWA"}

    def ledger_exists(r, k, extra=""):
        return f'COUNTIFS({led_base.format(r=r)},L_Kengetal,"{k}"{extra})'

    for v in VERSIES:
        r = CR[v]
        put(ws, f"A{r}", v, bold=True)
        ws[f"B{r}"] = f'=IF(C{r}>0,"Vastgesteld","n.b.")'
        ws[f"C{r}"] = "=" + ledger_exists(r, K_OMZ)
        ws[f"D{r}"] = "=" + ledger_exists(r, K_OMZA)
        ws[f"E{r}"] = "=" + ledger_exists(r, K_WFY)
        ws[f"F{r}"] = f'=COUNTIFS({led_base.format(r=r)},L_Revisie,">0")'
        for n, _lab, k, extra, _up, kind in metrics:
            cl = CC[n]
            s = f'SUMIFS(L_Waarde,{led_base.format(r=r)},L_Kengetal,"{k}"{extra})'
            if kind in actual_kinds:
                ex = ledger_exists(r, k, extra if kind in ("FA", "FWA") else "")
                ws[f"{cl}{r}"] = f'=IF(Sel_Doeljaar>P_VJ,"n.v.t.",IF({ex}=0,"n.b.",{s}))'
            elif kind in ("F", "FW"):
                ws[f"{cl}{r}"] = f'=IF({ledger_exists(r, k, extra)}=0,"n.b.",{s})'
            else:
                ws[f"{cl}{r}"] = f'=IF(${kind}{r}=0,"n.b.",{s})'
        for rb in RUBRIEKEN:
            for pre, soort in (("AV_", "FY"), ("AA_", "Actual YTD")):
                crit = f'A_Doel,Sel_Doeljaar,A_VJ,P_VJ,A_Versie,$A{r},A_Rubriek,"{rb}",A_Soort,"{soort}",A_TeltVast,1'
                ws[f"{CC[pre+rb]}{r}"] = f'=IF(COUNTIFS({crit})=0,"n.b.",SUMIFS(A_Waarde,{crit}))'
    # Upload (concept) uit huidige bronnen
    r = CR["Upload"]
    put(ws, f"A{r}", "Upload", bold=True)
    ws[f"B{r}"] = '=IF(C{0}>0,"Concept","n.b.")'.format(r)
    ws[f"C{r}"] = '=IF(OR(Sel_Doeljaar=P_RJ,Sel_Doeljaar=P_RJ+1),COUNT(V_OmzFY),0)'
    ws[f"D{r}"] = '=IF(Sel_Doeljaar=P_RJ,COUNT(V_OmzYTD),0)'
    ws[f"E{r}"] = '=IF(OR(Sel_Doeljaar=P_RJ,Sel_Doeljaar=P_RJ+1),COUNT(V_W_Aantal),0)'
    ws[f"F{r}"] = 0
    for n, _lab, k, extra, up, kind in metrics:
        cl = CC[n]
        if kind in actual_kinds:
            ws[f"{cl}{r}"] = f'=IF(Sel_Doeljaar<>P_RJ,"n.v.t.",IF($D{r}=0,"n.b.",{up}))'
        elif kind in ("E", "FW"):
            ws[f"{cl}{r}"] = f'=IF($E{r}=0,"n.b.",{up})'
        else:
            ws[f"{cl}{r}"] = f'=IF($C{r}=0,"n.b.",{up})'
    for rb in RUBRIEKEN:
        for pre, soort in (("AV_", "FY"), ("AA_", "Actual YTD")):
            crit = f'A_Doel,Sel_Doeljaar,A_VJ,P_UploadVJ,A_Versie,P_UploadVersie,A_Rubriek,"{rb}",A_Soort,"{soort}",A_TeltConcept,1'
            ws[f"{CC[pre+rb]}{r}"] = f'=IF(COUNTIFS({crit})=0,"n.b.",SUMIFS(A_Waarde,{crit}))'
    # Geselecteerd en vorige versie
    rs, rp = CR["Sel"], CR["Vorige"]
    ws[f"A{rs}"] = "=Sel_Versie"
    ws[f"B{rs}"] = (f'=IF(INDEX($C$5:$C$9,MATCH(Sel_Versie,$A$5:$A$9,0))>0,"Vastgesteld",'
                    f'IF(AND(Sel_Versie=P_UploadVersie,P_VJ=P_UploadVJ,$C$10>0),"Concept","n.b."))')
    ws[f"A{rp}"] = '=IF(Sel_Prev="","(geen)",Sel_Prev)'
    ws[f"B{rp}"] = f'=IF(Sel_Prev="","n.b.",IF(INDEX($C$5:$C$9,MATCH(Sel_Prev,$A$5:$A$9,0))>0,"Vastgesteld","n.b."))'
    last = col(c0 + len(allm) - 1)
    from openpyxl.utils import column_index_from_string as ci
    for cidx in range(3, ci(last) + 1):
        cl = col(cidx)
        ws[f"{cl}{rs}"] = (f'=IF($B${rs}="Vastgesteld",INDEX({cl}$5:{cl}$9,MATCH(Sel_Versie,$A$5:$A$9,0)),'
                           f'IF($B${rs}="Concept",{cl}$10,"n.b."))')
        ws[f"{cl}{rp}"] = f'=IF($B${rp}="Vastgesteld",INDEX({cl}$5:{cl}$9,MATCH(Sel_Prev,$A$5:$A$9,0)),"n.b.")'
    for r in range(5, 13):
        for cidx in range(1, ci(last) + 1):
            ws.cell(row=r, column=cidx).font = font(8, r >= 11)
            if cidx >= 3:
                ws.cell(row=r, column=cidx).number_format = FMT_EUR
        if r >= 11:
            row_border(ws, r, 1, ci(last), fillc=YELLOW_LT)
    # afgeleide selectienamen
    put(ws, "A15", "Afgeleide selectie", bold=True, color=NAVY)
    sel = [
        ("Sel_WBS", '=IF(Sel_Project="Alle","*",IFERROR(INDEX(Map_WBS,MATCH(Sel_Project,Map_Naam,0)),"(onbekend)"))', "WBS-criterium (* = alle)"),
        ("Sel_TypeCrit", '=IF(Sel_Type="Alle","*",Sel_Type)', "Type-criterium (* = alle)"),
        ("Sel_Prev", '=IF(Sel_Versie="Budget","",INDEX(V_List,MATCH(Sel_Versie,V_List,0)-1))', "Vorige versie in de reeks"),
        ("Sel_Bron", f"=$B${rs}", "Bron geselecteerde versie"),
        ("Sel_Label", f'=Sel_Versie&" "&P_VJ&" ("&LOWER($B${rs})&")"', "Label voor titels"),
        ("Sel_FiltLabel", '=IF(AND(Sel_Type="Alle",Sel_Project="Alle"),"geen filter",IF(Sel_Project<>"Alle",Sel_Project,Sel_Type))', "Filterlabel"),
    ]
    for i, (n, f, d) in enumerate(sel):
        r = 16 + i
        put(ws, f"A{r}", n)
        ws[f"B{r}"] = f
        put(ws, f"C{r}", d, color=GREY)
        add_name(wb, n, f"CALC_Versies!$B${r}")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.freeze_panes = "C5"


# ---------------------------------------------------------------------------------------------
# Vastleggen: concept van de huidige upload in exact het grootboekformaat (aaneengesloten blok)
# ---------------------------------------------------------------------------------------------
KT = [  # k: (doeljaar-offset, kengetal, VERW-kolom, eenheid, bron)
    (0, K_OMZ, "OmzFY", "EUR", "Z"), (0, K_BOR, "BorFY", "EUR", "Z"), (0, K_OMZA, "OmzYTD", "EUR", "Z"),
    (0, K_BORA, "BorYTD", "EUR", "Z"), (1, K_OMZ, "OmzFY1", "EUR", "Z"), (1, K_BOR, "BorFY1", "EUR", "Z"),
    (0, K_WFY, "W_HuidigJaar", "aantal", "E"), (0, K_WVJ, "W_VJ_T", "aantal", "E"),
    (0, K_WQ[0], "W_HQ1", "aantal", "E"), (0, K_WQ[1], "W_HQ2", "aantal", "E"), (0, K_WQ[2], "W_HQ3", "aantal", "E"),
    (0, K_WQ[3], "W_HQ4", "aantal", "E"), (0, K_WACT, "W_ActYTD", "aantal", "E"), (1, K_WFY, "W_VolgendJaar", "aantal", "E"),
]
C0, C1 = 8, 8 + (MAP1 - MAP0 + 1) * len(KT) - 1


def build_vastleggen(wb, ws):
    base_sheet(ws, zoom=85)
    title(ws, "Vastleggen — concept uit huidige upload", None, 22)
    put(ws, "B2", "HANDMATIGE STAP (geen knop/macro): kopieer A8:R<laatste regel> en plak als WAARDEN in Vastlegging op de 'eerste vrije rij'. Daarna Importlog bijwerken.",
        bold=True, color=RED)
    put(ws, "B3", "Te plakken regels:", bold=True)
    ws["E3"] = f'=COUNTIF(J{C0}:J{C1},"?*")'
    put(ws, "F3", "→ rijen 8 t/m", color=GREY)
    ws["H3"] = f'=7+E3'
    put(ws, "J3", "Al vastgesteld (zelfde versie):", bold=True)
    ws["N3"] = '=COUNTIFS(L_Versiejaar,P_UploadVJ,L_Versie,P_UploadVersie,L_Status,"Vastgesteld")'
    ws["B4"] = ('=IF(N3>0,"LET OP: "&P_UploadVersie&" "&P_UploadVJ&" is al vastgesteld. Opnieuw plakken = CORRECTIE (revisie +1, toelichting verplicht in Parameters › P_ToelCorrectie). '
                'Alleen doen bij een bewuste correctie.","Versie "&P_UploadVersie&" "&P_UploadVJ&" is nog niet vastgesteld: plakken legt revisie 0 vast.")')
    ws["B4"].font = font(9, True, NAVY)
    ws["B5"] = ('=IF(Controle!$D$4>0,"Controle bevat "&Controle!$D$4&" FOUT(en): eerst oplossen vóór vastleggen.",'
                'IF(AND(N3>0,P_ToelCorrectie=""),"Controle: geen fouten. Let op: bij een correctie eerst P_ToelCorrectie invullen.","Controle: geen fouten."))')
    ws["B5"].font = font(9, True, RED)
    ws["B6"] = f'=IF(SUM(U{C0}:U{C1})=0,"Concept sluit volledig aan op vastgelegde waarden (of er is nog niets vastgelegd).","Concept wijkt af van vastgelegde waarden: zie kolom U (verschil).")'
    ws["B6"].font = font(9, False, GREY, True)
    header(ws, 7, 1, LCOLS + ["Sleutel", "Vastgelegd (actief)", "Verschil", "", "", "", "",
                               "map-idx", "k", "waarde", "geldig", "volgnr", "rij-idx"], height=30)
    # kengetallentabel
    kt0 = 8
    put(ws, "AJ6", "k", bold=True)
    for j, h in enumerate(["offset", "kengetal", "eenheid", "bron"]):
        put(ws, f"{col(37+j)}6", h, bold=True)
    for k, (off, kg, vc, eh, br) in enumerate(KT):
        put(ws, f"AJ{kt0+k}", k)
        put(ws, f"AK{kt0+k}", off)
        put(ws, f"AL{kt0+k}", kg)
        put(ws, f"AM{kt0+k}", eh)
        put(ws, f"AN{kt0+k}", br)
    ktr = lambda c: f"${c}${kt0}:${c}${kt0+len(KT)-1}"  # noqa: E731
    # ruwe regels (AA..AF): per mappingregel x kengetal
    r = C0
    for mi in range(MAP1 - MAP0 + 1):
        vr = V0 + mi
        for k, (off, kg, vc, eh, br) in enumerate(KT):
            ws[f"AB{r}"] = mi + 1
            ws[f"AC{r}"] = k
            ws[f"AD{r}"] = f"=VERW_Project!{VCOLS[vc]}{vr}"
            ws[f"AE{r}"] = f'=IF(AND(VERW_Project!$A{vr}<>"",ISNUMBER(AD{r})),1,0)'
            ws[f"AF{r}"] = f'=IF(AE{r}=1,SUM($AE${C0}:AE{r}),"")'
            for c in ("AB", "AC", "AD", "AE", "AF"):
                ws[f"{c}{r}"].font = font(7, False, GREY)
            r += 1
    for r in range(C0, C1 + 1):
        i = r - C0 + 1
        ws[f"AG{r}"] = f'=IFERROR(MATCH({i},$AF${C0}:$AF${C1},0),"")'
        g = f"$AG{r}"
        mi = f"INDEX($AB${C0}:$AB${C1},{g})"
        kk = f"INDEX($AC${C0}:$AC${C1},{g})+1"
        vv = lambda n: f"INDEX(V_{n},{mi})"  # noqa: E731
        f = {
            "A": "P_ImportID",
            "B": f'IF(INDEX({ktr("AN")},{kk})="Z",P_BronZ,P_BronE)',
            "C": "P_Importdatum", "D": "P_StatusVastleg", "E": "P_UploadVJ", "F": "P_UploadVersie",
            "G": (f'IF(COUNTIFS(L_Sleutel,$S{r},L_Status,"Vastgesteld")=0,0,'
                  f'_xlfn.MAXIFS(L_Revisie,L_Sleutel,$S{r},L_Status,"Vastgesteld")+1)'),
            "H": "P_Peildatum",
            "I": f'P_RJ+INDEX({ktr("AK")},{kk})',
            "J": vv("WBS"), "K": vv("Projdef"), "L": vv("Naam"), "M": vv("Type"), "N": vv("IFRS"),
            "O": f'INDEX({ktr("AL")},{kk})',
            "P": f"INDEX($AD${C0}:$AD${C1},{g})",
            "Q": f'INDEX({ktr("AM")},{kk})',
            "R": 'IF($G{r}>0,P_ToelCorrectie&"","")'.format(r=r),
        }
        for c, body in f.items():
            ws[f"{c}{r}"] = f'=IF({g}="","",{body})'
            ws[f"{c}{r}"].font = font(8)
        ws[f"C{r}"].number_format = FMT_DATE
        ws[f"H{r}"].number_format = FMT_DATE
        ws[f"P{r}"].number_format = FMT_RAW
        ws[f"S{r}"] = f'=IF(J{r}="","",J{r}&"|"&E{r}&"|"&F{r}&"|"&I{r}&"|"&O{r})'
        ws[f"T{r}"] = f'=IF(S{r}="","",IF(COUNTIFS(L_Sleutel,S{r},L_Actief,1)=0,"",SUMIFS(L_Waarde,L_Sleutel,S{r},L_Actief,1)))'
        ws[f"U{r}"] = f'=IF(OR(S{r}="",T{r}=""),0,P{r}-T{r})'
        for c in "STU":
            ws[f"{c}{r}"].font = font(8, False, GREY)
        ws[f"T{r}"].number_format = FMT_RAW
        ws[f"U{r}"].number_format = FMT_RAW
    widths(ws, {"A": 15, "B": 22, "C": 10, "D": 11, "E": 8, "F": 7, "G": 6, "H": 10, "I": 7, "J": 16, "K": 10,
                "L": 24, "M": 10, "N": 5, "O": 24, "P": 13, "Q": 7, "R": 20, "S": 28, "T": 13, "U": 10})
    add_name(wb, "C_Verschil", f"Vastleggen!$U${C0}:$U${C1}")
    add_name(wb, "C_Kengetal", f"Vastleggen!$O${C0}:$O${C1}")
    add_name(wb, "C_Doeljaar", f"Vastleggen!$I${C0}:$I${C1}")
    add_name(wb, "C_Vastgelegd", f"Vastleggen!$T${C0}:$T${C1}")
    ws.freeze_panes = f"A{C0}"


# ---------------------------------------------------------------------------------------------
# RDFO
# ---------------------------------------------------------------------------------------------
RDFO_ROWS = {}


def build_rdfo(wb, ws):
    base_sheet(ws, zoom=100)
    title(ws, None, None, 16)
    ws["B1"] = '="RDFO "&P_Regio&" — doeljaar "&Sel_Doeljaar&"  ·  bedragen in "&P_Eenheid'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = ('="Versiejaar "&P_VJ&"  ·  geselecteerde versie: "&Sel_Label&"  ·  vergeleken met "&IF(Sel_Prev="","–",Sel_Prev)'
                '&"  ·  upload peildatum "&P_PeilTxt&"  ·  selectie wijzigen op Dashboard"')
    ws["B2"].font = font(9, False, GREY, True)
    vcols = {"C": "Budget", "D": "Q1", "E": "Q2", "F": "Q3", "G": "Q4", "H": "Upload", "J": "Sel", "O": "Vorige"}
    for c, v in vcols.items():
        put(ws, f"{c}5", v, color=WHITE, size=7)
    labels = {
        "B": "Rubriek",
        "C": '="Budget"&CHAR(10)&"FY "&Sel_Doeljaar',
        "H": '="Upload "&P_PeilTxt&CHAR(10)&"concept"',
        "J": '="Geselecteerd"&CHAR(10)&Sel_Versie&" FY"',
        "K": '="Realisatie YTD"&CHAR(10)&"zelfde peildatum"',
        "L": '="To go"&CHAR(10)&"(FY − YTD)"',
        "M": '="Δ geselect."&CHAR(10)&"vs Budget"',
        "N": '="Δ geselect."&CHAR(10)&"vs "&IF(Sel_Prev="","vorige",Sel_Prev)',
        "O": "Vorige versie (hulp)",
        "P": "Definitie / bron",
    }
    for i, v in enumerate(VERSIES[1:]):
        c = "DEFG"[i]
        labels[c] = f'="{v} "&P_VJ&CHAR(10)&"FY-prognose"&IF(CALC_Versies!$F${CR[v]}>0," ●","")'
    for c in "BCDEFGHIJKLMNOP":
        cell = ws[f"{c}6"]
        cell.value = labels.get(c, "")
        cell.fill = fill(NAVY if c not in "IK" else ("FFFFFF" if c == "I" else NAVY2))
        cell.font = font(9, True, WHITE)
        cell.alignment = Alignment(horizontal="center" if c != "B" else "left", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 36
    F = lambda m: CC[m]  # noqa: E731
    lines = [
        ("omz_ops", "Omzet opstal", "base", F("OmzOps"), F("ActOmzOps"), "ZPMW03, Type=Opstal (Mapping). FY RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ; RJ+1 = −OP.BOB RJ+1"),
        ("omz_dln", "Omzet deelnemingen", "base", F("OmzDln"), F("ActOmzDln"), "ZPMW03, Type=Deelneming (Mapping). Al op aandeelbasis; geen aandeel toegepast"),
        ("taak_omz", "Taakstelling omzet", "base", F("AV_Taakstelling omzet"), F("AA_Taakstelling omzet"), "Invoer_Aanvullend (rubriek Taakstelling omzet)"),
        ("int_elim", "Interne eliminaties", "base", F("AV_Interne eliminaties"), F("AA_Interne eliminaties"), "Invoer_Aanvullend"),
        ("omzet_old", "Omzet old", "sub", ["omz_ops", "omz_dln", "taak_omz", "int_elim"], None, "Subtotaal: opstal + deelnemingen + taakstelling + interne eliminaties"),
        ("ifrs_elim", "IFRS-eliminaties", "base", F("AV_IFRS-eliminaties"), F("AA_IFRS-eliminaties"), "Invoer_Aanvullend. IFRS J/N wordt NIET automatisch geëlimineerd (zie memo)"),
        ("omzet_ifrs", "Omzet IFRS", "sub", ["omzet_old", "ifrs_elim"], None, "Omzet old + IFRS-eliminaties"),
        (None,),
        ("bor_ops", "Projectresultaat (BOR) opstal", "base", F("BorOps"), F("ActBorOps"), "ZPMW03: FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ. Verlies blijft negatief (geen ABS)"),
        ("bor_dln", "Projectresultaat (BOR) deelnemingen", "base", F("BorDln"), F("ActBorDln"), "Idem, Type=Deelneming"),
        ("bor_tot", "Projectresultaten / BOR", "sub", ["bor_ops", "bor_dln"], None, "Subtotaal BOR"),
        ("ak", "AK-dekking", "base", F("AV_AK-dekking"), F("AA_AK-dekking"), "Invoer_Aanvullend. Geen generiek AK-percentage toegepast"),
        ("fees", "Fees", "base", F("AV_Fees"), F("AA_Fees"), "Invoer_Aanvullend"),
        ("overig", "Overige resultaten", "base", F("AV_Overige resultaten"), F("AA_Overige resultaten"), "Invoer_Aanvullend"),
        ("taak_ebit", "Taakstelling EBIT", "base", F("AV_Taakstelling EBIT"), F("AA_Taakstelling EBIT"), "Invoer_Aanvullend"),
        ("ebit", "EBIT-bijdrage", "sub", ["bor_tot", "ak", "fees", "overig", "taak_ebit"], None, "BOR + AK-dekking + fees + overige resultaten + taakstelling EBIT"),
        (None,),
        ("bor_pct", "BOR %", "pct", ("bor_tot", "omz_proj"), None, "Totale BOR / totale projectomzet (opstal + deelnemingen). Geen gemiddelde van projectpercentages"),
        ("ebit_pct", "EBIT %", "pct", ("ebit", "omzet_old"), None, "EBIT-bijdrage / Omzet old (definitie te bevestigen, zie Productievragen)"),
        (None,),
        ("memo", "Memo", "hdr"),
        ("omz_proj", "Omzet projecten (noemer BOR %)", "sub", ["omz_ops", "omz_dln"], None, "Omzet opstal + deelnemingen"),
        ("omz_ifrsn", "waarvan projecten met IFRS = N", "base", F("OmzIFRSN"), None, "Informatief: IFRS J/N is een dimensie, geen automatische eliminatie- of consolidatieregel"),
        ("bor_ifrsn", "BOR projecten met IFRS = N", "base", F("BorIFRSN"), None, "Informatief"),
        ("won_fy", "Woningverkopen (aantal, aandeelbasis)", "count", F("WonFY"), F("WonAct"), "Woningetalage 'Huidig jaar' (RJ) / 'Volgend jaar' (RJ+1). Verkopen ≠ opleveringen/transporten"),
    ]
    vc_rows = {"C": 5, "D": 6, "E": 7, "F": 8, "G": 9, "H": 10, "J": 11, "O": 12}
    r = 7
    for ln in lines:
        if ln[0] is not None:
            RDFO_ROWS[ln[0]] = r
        r += 1
    togo_keys = {"omz_ops", "omz_dln", "omzet_old", "omzet_ifrs", "bor_ops", "bor_dln", "bor_tot", "ebit", "omz_proj"}
    r = 7
    for ln in lines:
        if ln[0] is None:
            r += 1
            continue
        key, lab, kind = ln[0], ln[1], ln[2]
        if kind == "hdr":
            put(ws, f"B{r}", lab, bold=True, color=NAVY)
            ws[f"B{r}"].border = Border(bottom=thin)
            r += 1
            continue
        src, act, desc = ln[3], ln[4], ln[5]
        put(ws, f"B{r}", lab, bold=kind in ("sub", "pct"), size=9 if key not in ("omz_ifrsn", "bor_ifrsn") else 8,
            color="000000" if key not in ("omz_ifrsn", "bor_ifrsn") else GREY)
        if key in ("omz_ifrsn", "bor_ifrsn"):
            ws[f"B{r}"].alignment = Alignment(indent=1)
        put(ws, f"P{r}", desc, color=GREY, size=8)
        cols_ = list(vc_rows.keys()) + ["K"]
        for c in cols_:
            if kind in ("base", "count"):
                if c == "K":
                    if act is None:
                        ws[f"K{r}"] = '=""'
                        continue
                    ref = f"CALC_Versies!${act}${CR['Sel']}"
                else:
                    ref = f"CALC_Versies!${src}${vc_rows[c]}"
                if kind == "count":
                    ws[f"{c}{r}"] = f"={ref}"
                else:
                    ws[f"{c}{r}"] = f"=IF(ISNUMBER({ref}),{ref}/P_Factor,{ref})"
            elif kind == "sub":
                cells = [f"{c}{RDFO_ROWS[k]}" for k in src]
                ws[f"{c}{r}"] = f'=IF(ISNUMBER({cells[0]}),SUM({",".join(cells)}),{cells[0]})'
            elif kind == "pct":
                n_, d_ = f"{c}{RDFO_ROWS[src[0]]}", f"{c}{RDFO_ROWS[src[1]]}"
                ws[f"{c}{r}"] = f'=IF(AND(ISNUMBER({n_}),ISNUMBER({d_})),IF({d_}<>0,{n_}/{d_},"n.b."),"n.b.")'
        # to go en delta's
        if key in togo_keys:
            ws[f"L{r}"] = f'=IF(K{r}="n.v.t.",J{r},IF(AND(ISNUMBER(J{r}),ISNUMBER(K{r})),J{r}-K{r},"n.b."))'
        else:
            ws[f"L{r}"] = '=""'
        ws[f"M{r}"] = f'=IF(AND(ISNUMBER(J{r}),ISNUMBER(C{r})),J{r}-C{r},"n.b.")'
        ws[f"N{r}"] = f'=IF(AND(ISNUMBER(J{r}),ISNUMBER(O{r})),J{r}-O{r},"n.b.")'
        fmt = FMT_PCT if kind == "pct" else (FMT_INT if kind == "count" else FMT_EUR)
        for c in "CDEFGHJKLMNO":
            cell = ws[f"{c}{r}"]
            cell.number_format = fmt if not (kind == "pct" and c in "MN") else '+0.0%;[Red]-0.0%;"–"'
            cell.alignment = Alignment(horizontal="right")
            cell.font = font(9 if key not in ("omz_ifrsn", "bor_ifrsn") else 8, kind in ("sub", "pct"),
                             GREY if key in ("omz_ifrsn", "bor_ifrsn") else "000000")
            cell.border = Border(bottom=thin, top=thick_navy if kind == "sub" and key not in ("omz_proj",) else None)
            if c == "K":
                cell.fill = fill("EEF3EA")
                cell.font = font(9, kind in ("sub", "pct"), "3B5B33", True)
            if c == "O":
                cell.font = font(7, False, GREY)
        if kind == "sub":
            ws[f"B{r}"].border = Border(top=thick_navy, bottom=thin)
            # oranje = onvolledig (een onderliggende post is n.b.)
            leaves = []

            def walk(k):
                ln_ = next(x for x in lines if x[0] == k)
                if ln_[2] == "sub":
                    for kk in ln_[3]:
                        walk(kk)
                else:
                    leaves.append(RDFO_ROWS[k])
            walk(key)
            cond = f"AND(ISNUMBER(C{r}),OR(" + ",".join(f'C${lr}="n.b."' for lr in leaves) + "))"
            ws.conditional_formatting.add(f"C{r}:N{r}", FormulaRule(formula=[cond], fill=fill(ORANGE_LT)))
        r += 1
    last = r - 1
    # geselecteerde versiekolom geel markeren
    ws.conditional_formatting.add(f"C7:G{last}", FormulaRule(formula=["C$5=Sel_Versie"], fill=fill(YELLOW_LT)))
    ws.conditional_formatting.add("C6:G6", FormulaRule(formula=["C$5=Sel_Versie"], fill=fill(YELLOW), font=Font(color=NAVY, bold=True)))
    ws.conditional_formatting.add(f"J7:J{last}", FormulaRule(formula=["TRUE"], fill=fill(YELLOW_LT)))
    r = last + 2
    notes = [
        "Leeswijzer",
        "• Budget/Q1–Q4 = vastgelegde (onveranderlijke) FY-jaarprognoses uit Vastlegging. Een nieuwe upload wijzigt deze kolommen niet. ● = versie bevat een correctie (revisie > 0).",
        "• 'Upload … concept' = live berekend uit SRC_ZPMW03/SRC_Woningetalage; wordt pas een versie na handmatige vastlegging (tab Vastleggen).",
        "• 'Realisatie YTD' (groen, cursief) is géén prognose: actual t/m de peildatum van de geselecteerde versie. To go = geselecteerde FY − realisatie van dezelfde peildatum.",
        "• n.b. = niet beschikbaar (geen vastgelegde of aanvullende invoer) — nooit stil als nul. n.v.t. = doeljaar ligt na het versiejaar, dus nog geen realisatie.",
        "• Oranje subtotaal = onvolledig: minstens één onderliggende post is n.b. De som bevat alleen de beschikbare posten.",
        "• Tekens: omzet en resultaat positief, verlies negatief (rood). Aanvullende invoer in managementteken (+ verhoogt omzet/resultaat).",
    ]
    for i, t in enumerate(notes):
        put(ws, f"B{r+i}", t, bold=i == 0, color=NAVY if i == 0 else GREY, size=8 if i else 9)
    widths(ws, {"B": 36, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 1.5, "J": 12, "K": 12,
                "L": 11, "M": 11, "N": 11, "O": 9, "P": 70})
    ws.column_dimensions["O"].hidden = True
    ws.row_dimensions[5].hidden = True
    ws.freeze_panes = "C7"
    landscape(ws, f"A1:N{r+len(notes)}")


# ---------------------------------------------------------------------------------------------
# Projecten
# ---------------------------------------------------------------------------------------------
P0, P1 = 10, 10 + (V1 - V0)


def ledger_proj(wbs, ver, k, extra=""):
    crit = f'L_Actief,1,L_WBS,{wbs},L_Doeljaar,Sel_Doeljaar,L_Versiejaar,P_VJ,L_Versie,{ver},L_Kengetal,"{k}"{extra}'
    return f'IF(COUNTIFS({crit})=0,"n.b.",SUMIFS(L_Waarde,{crit}))'


def build_projecten(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, None, None, 20)
    ws["B1"] = '="Projectoverzicht "&P_Regio&" — omzet doeljaar "&Sel_Doeljaar&" ("&P_Eenheid&")"'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = '="Geselecteerde versie: "&Sel_Label&"  ·  realisatie en to go van dezelfde peildatum  ·  filter met de pijltjes in rij 9"'
    ws["B2"].font = font(9, False, GREY, True)
    codes = {"E": "Budget", "F": "Q1", "G": "Q2", "H": "Q3", "I": "Q4"}
    for c, v in codes.items():
        put(ws, f"{c}8", v, color=WHITE, size=7)
    hdr = ["WBS-element", "Projectnaam", "Type", "IFRS", '="Budget"&CHAR(10)&"FY "&Sel_Doeljaar',
           '="Q1 "&P_VJ&CHAR(10)&"FY"', '="Q2 "&P_VJ&CHAR(10)&"FY"', '="Q3 "&P_VJ&CHAR(10)&"FY"', '="Q4 "&P_VJ&CHAR(10)&"FY"',
           '="Upload"&CHAR(10)&"concept"', '="Geselect."&CHAR(10)&Sel_Versie&" FY"', "Realisatie\nYTD", "Omzet\nto go",
           '="BOR FY"&CHAR(10)&Sel_Versie', "Marge\n%", "Δ vs\nBudget", '="Δ vs "&IF(Sel_Prev="","vorige",Sel_Prev)',
           "Verlies", "Toelichting (Invoer_Toelichting)"]
    header(ws, 9, 1, hdr, height=36)
    ws["L9"].fill = fill(NAVY2)
    for r in range(P0, P1 + 1):
        vr = V0 + (r - P0)
        A = f"$A{r}"
        ws[f"A{r}"] = f'=IF(VERW_Project!$A{vr}="","",VERW_Project!$A{vr})'
        ws[f"B{r}"] = f'=IF({A}="","",VERW_Project!$C{vr})'
        ws[f"C{r}"] = f'=IF({A}="","",VERW_Project!$D{vr})'
        ws[f"D{r}"] = f'=IF({A}="","",VERW_Project!$I{vr})'
        for c in "EFGHI":
            ws[f"{c}{r}"] = f'=IF({A}="","",IF(ISNUMBER(_V),_V/P_Factor,_V))'.replace("_V", ledger_proj(A, f"{c}$8", K_OMZ))
        up = f'IF(Sel_Doeljaar=P_RJ,VERW_Project!${VCOLS["OmzFY"]}{vr},IF(Sel_Doeljaar=P_RJ+1,VERW_Project!${VCOLS["OmzFY1"]}{vr},"n.b."))'
        ws[f"J{r}"] = f'=IF({A}="","",IF(ISNUMBER({up}),{up}/P_Factor,{up}))'
        ws[f"K{r}"] = f'=IF({A}="","",IF(Sel_Bron="Vastgesteld",INDEX($E{r}:$I{r},MATCH(Sel_Versie,$E$8:$I$8,0)),IF(Sel_Bron="Concept",J{r},"n.b.")))'
        act_led = ledger_proj(A, "Sel_Versie", K_OMZA)
        act_up = f'IF(Sel_Doeljaar=P_RJ,VERW_Project!${VCOLS["OmzYTD"]}{vr},"n.v.t.")'
        ws[f"L{r}"] = (f'=IF({A}="","",IF(Sel_Doeljaar>P_VJ,"n.v.t.",IF(Sel_Bron="Vastgesteld",IF(ISNUMBER({act_led}),{act_led}/P_Factor,{act_led}),'
                       f'IF(Sel_Bron="Concept",IF(ISNUMBER({act_up}),{act_up}/P_Factor,{act_up}),"n.b."))))')
        ws[f"M{r}"] = f'=IF({A}="","",IF(L{r}="n.v.t.",K{r},IF(AND(ISNUMBER(K{r}),ISNUMBER(L{r})),K{r}-L{r},"n.b.")))'
        bor_led = ledger_proj(A, "Sel_Versie", K_BOR)
        bor_up = f'IF(Sel_Doeljaar=P_RJ,VERW_Project!${VCOLS["BorFY"]}{vr},IF(Sel_Doeljaar=P_RJ+1,VERW_Project!${VCOLS["BorFY1"]}{vr},"n.b."))'
        ws[f"N{r}"] = (f'=IF({A}="","",IF(Sel_Bron="Vastgesteld",IF(ISNUMBER({bor_led}),{bor_led}/P_Factor,{bor_led}),'
                       f'IF(Sel_Bron="Concept",IF(ISNUMBER({bor_up}),{bor_up}/P_Factor,{bor_up}),"n.b.")))')
        ws[f"O{r}"] = f'=IF({A}="","",IF(AND(ISNUMBER(N{r}),ISNUMBER(K{r})),IF(K{r}<>0,N{r}/K{r},"n.b."),"n.b."))'
        ws[f"P{r}"] = f'=IF({A}="","",IF(AND(ISNUMBER(K{r}),ISNUMBER(E{r})),K{r}-E{r},"n.b."))'
        prev = ledger_proj(A, "Sel_Prev", K_OMZ)
        ws[f"W{r}"] = f'=IF(OR({A}="",Sel_Prev=""),"n.b.",IF(ISNUMBER({prev}),{prev}/P_Factor,{prev}))'
        ws[f"Q{r}"] = f'=IF({A}="","",IF(AND(ISNUMBER(K{r}),ISNUMBER(W{r})),K{r}-W{r},"n.b."))'
        ws[f"R{r}"] = f'=IF({A}="","",IF(ISNUMBER(N{r}),IF(N{r}<0,"Verlies",""),""))'
        ws[f"S{r}"] = f'=IF({A}="","",IFERROR(INDEX(T_Txt,MATCH({A}&"|"&Sel_Doeljaar&"|"&P_VJ&"|"&Sel_Versie,T_Key,0)),""))'
        match = f'AND({A}<>"",OR(Sel_Type="Alle",$C{r}=Sel_Type),OR(Sel_Project="Alle",$B{r}=Sel_Project))'
        ws[f"U{r}"] = f'=IF(AND({match},ISNUMBER(K{r})),K{r}+ROW()/1000000000,"")'
        ws[f"V{r}"] = f'=IF(AND({match},ISNUMBER(Q{r})),ABS(Q{r})+ROW()/1000000000,"")'
        for c in "ABCDEFGHIJKLMNOPQRS":
            cell = ws[f"{c}{r}"]
            cell.font = font(9)
            cell.border = Border(bottom=thin)
            if c in "EFGHIJKLMNPQ":
                cell.number_format = FMT_EUR
                cell.alignment = Alignment(horizontal="right")
            if c == "O":
                cell.number_format = FMT_PCT
                cell.alignment = Alignment(horizontal="right")
        ws[f"L{r}"].font = font(9, False, "3B5B33", True)
        ws[f"R{r}"].font = font(9, True, RED)
        for c in "UVW":
            helper_style(ws[f"{c}{r}"])
    # totalen boven de kop (blijven staan bij filteren)
    put(ws, "B6", "Totaal (zichtbare regels)", bold=True)
    for c in "EFGHIJKLMNPQ":
        ws[f"{c}6"] = f'=IF(SUBTOTAL(102,{c}{P0}:{c}{P1})=0,"n.b.",SUBTOTAL(109,{c}{P0}:{c}{P1}))'
    ws["O6"] = '=IF(AND(ISNUMBER(N6),ISNUMBER(K6)),IF(K6<>0,N6/K6,"n.b."),"n.b.")'
    ws["R6"] = f'=COUNTIF(R{P0}:R{P1},"Verlies")&" verliesproject(en)"'
    for c in "BCDEFGHIJKLMNOPQR":
        cell = ws[f"{c}6"]
        cell.font = font(9, True, NAVY if c != "R" else RED)
        cell.fill = fill(LIGHT)
        cell.border = Border(top=thick_navy, bottom=thick_navy)
        if c in "EFGHIJKLMNPQ":
            cell.number_format = FMT_EUR
            cell.alignment = Alignment(horizontal="right")
    ws["O6"].number_format = FMT_PCT
    put(ws, "B7", "Marge % totaal = som BOR / som omzet (geen gemiddelde). Realisatie en to go horen bij de peildatum van de geselecteerde versie.", color=GREY, italic=True, size=8)
    ws.conditional_formatting.add(f"E{P0}:I{P1}", FormulaRule(formula=[f'AND($A{P0}<>"",E$8=Sel_Versie)'], fill=fill(YELLOW_LT)))
    ws.conditional_formatting.add(f"K{P0}:K{P1}", FormulaRule(formula=[f'$A{P0}<>""'], fill=fill(YELLOW_LT)))
    ws.conditional_formatting.add(f"L{P0}:L{P1}", FormulaRule(formula=[f'$A{P0}<>""'], fill=fill("EEF3EA")))
    ws.auto_filter.ref = f"A9:S{P1}"
    widths(ws, {"A": 15, "B": 27, "C": 10, "D": 5, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 10, "L": 10,
                "M": 9, "N": 9, "O": 7, "P": 9, "Q": 9, "R": 7, "S": 50, "T": 2, "U": 8, "V": 8, "W": 8})
    ws.row_dimensions[8].hidden = True
    for c in "UVW":
        ws.column_dimensions[c].hidden = True
    ws.freeze_panes = f"C{P0}"
    landscape(ws, f"A1:S{P0 + 30}")


# ---------------------------------------------------------------------------------------------
# Grafiek- en tegelhulpjes
# ---------------------------------------------------------------------------------------------
from openpyxl.chart.shapes import GraphicalProperties  # noqa: E402
from openpyxl.drawing.line import LineProperties  # noqa: E402


def style_chart(ch, colors, ytitle=None, numfmt="#,##0.0", w=16.5, h=7.2, labels=False, legend=True):
    ch.width, ch.height = w, h
    ch.style = 2
    ch.title = None
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.solidFill = c
        s.graphicalProperties.line.noFill = True
    ch.y_axis.numFmt = numfmt
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E7EE"))
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.tickLblPos = "low"
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    if ytitle:
        ch.y_axis.title = ytitle
    if legend:
        ch.legend.position = "b"
    else:
        ch.legend = None
    if labels:
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showVal = True
        ch.dataLabels.showSerName = False
        ch.dataLabels.showCatName = False
        ch.dataLabels.showLegendKey = False
        ch.dataLabels.showPercent = False
        ch.dataLabels.numFmt = numfmt + ";-" + numfmt + ";"
    ch.gapWidth = 60
    return ch


def bar(ws, data_ref, cats_ref, anchor, colors, kind="col", grouping="clustered", **kw):
    ch = BarChart()
    ch.type = kind
    ch.grouping = grouping
    if grouping == "stacked":
        ch.overlap = 100
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    style_chart(ch, colors, **kw)
    ws.add_chart(ch, anchor)
    return ch


def tile(ws, c1, c2, r, label, value, fmt, sub1, sub2=None):
    for rr in range(r, r + 4):
        for cc in range(c1, c2 + 1):
            ws.cell(row=rr, column=cc).fill = fill(LIGHT)
    for cc in range(c1, c2 + 1):
        ws.cell(row=r, column=cc).border = Border(top=Side(style="thick", color=YELLOW))
    a = ws.cell(row=r, column=c1, value=label)
    a.font = font(8, True, NAVY)
    ws.merge_cells(start_row=r + 1, start_column=c1, end_row=r + 1, end_column=c2)
    v = ws.cell(row=r + 1, column=c1, value=value)
    v.font = font(17, True, NAVY)
    v.number_format = fmt
    v.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r + 1].height = 28
    s = ws.cell(row=r + 2, column=c1, value=sub1)
    s.font = font(8, False, GREY)
    if sub2 is not None:
        s2 = ws.cell(row=r + 3, column=c1, value=sub2)
        s2.font = font(8, False, GREY)
    return v


from openpyxl.styles import Side  # noqa: E402

MLN = '#,##0.0" mln";[Red]-#,##0.0" mln";"–"'


def fx_mln(ref):
    return f'IF(ISNUMBER({ref}),FIXED({ref}/1000000,1),"n.b.")'


# ---------------------------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------------------------
def build_dashboard(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, None, None, 19)
    ws["B1"] = '="Regiodashboard "&P_Regio&" — doeljaar "&Sel_Doeljaar&" · versie "&Sel_Label'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = '="Bedragen in EUR mln (bron: EUR, volledige precisie) · upload peildatum "&P_PeilTxt&" · filter: "&Sel_FiltLabel'
    ws["B2"].font = font(9, False, GREY, True)
    # filters
    flt = [("B", "Doeljaar", 2026, "L_Doeljaar", "Sel_Doeljaar"), ("E", "Versie", "Q3", "V_List", "Sel_Versie"),
           ("H", "Opstal / deelneming", "Alle", "L_TypeList", "Sel_Type"), ("K", "Project", "Alle", "L_Projecten", "Sel_Project")]
    for c, lab, val, lst, name in flt:
        put(ws, f"{c}4", lab, bold=True, color=NAVY, size=8)
        cell = ws[f"{c}5"]
        cell.value = val
        input_style(cell)
        cell.font = font(10, True, "1F4E9A")
        dv_list(ws, f"{c}5", lst)
        add_name(wb, name, f"Dashboard!${c}$5")
    ws.merge_cells("K5:N5")
    ws.merge_cells("H5:I5")
    ws.merge_cells("E5:F5")
    ws.merge_cells("B5:C5")
    put(ws, "P4", "Controlestatus", bold=True, color=NAVY, size=8)
    ws["P5"] = '=IF(Controle!$D$4>0,"✖ "&Controle!$D$4&" fout(en)",IF(Controle!$D$5>0,"▲ "&Controle!$D$5&" waarschuwing(en)","✔ geen fouten"))'
    ws["P5"].font = font(10, True, NAVY)
    ws["P5"].hyperlink = "#'Controle'!A1"
    ws.conditional_formatting.add("P5", FormulaRule(formula=["Controle!$D$4>0"], font=Font(color=RED, bold=True)))
    S = lambda m, row="Sel": f"CALC_Versies!${CC[m]}${CR[row]}"  # noqa: E731
    r = 7
    tile(ws, 2, 3, r, '="Omzet FY "&Sel_Doeljaar&" · "&Sel_Versie', f'=IF(ISNUMBER({S("F_OmzFY")}),{S("F_OmzFY")}/1000000,"n.b.")', MLN,
         f'="Δ budget: "&IF(AND(ISNUMBER({S("F_OmzFY")}),ISNUMBER({S("F_OmzFY","Budget")})),FIXED(({S("F_OmzFY")}-{S("F_OmzFY","Budget")})/1000000,1)&" mln","n.b.")',
         f'="Δ "&IF(Sel_Prev="","vorige",Sel_Prev)&": "&IF(AND(ISNUMBER({S("F_OmzFY")}),ISNUMBER({S("F_OmzFY","Vorige")})),FIXED(({S("F_OmzFY")}-{S("F_OmzFY","Vorige")})/1000000,1)&" mln","n.b.")')
    tile(ws, 5, 6, r, '="BOR FY "&Sel_Doeljaar&" · "&Sel_Versie', f'=IF(ISNUMBER({S("F_BorFY")}),{S("F_BorFY")}/1000000,"n.b.")', MLN,
         f'="BOR %: "&IF(AND(ISNUMBER({S("F_BorFY")}),ISNUMBER({S("F_OmzFY")})),IF({S("F_OmzFY")}<>0,FIXED({S("F_BorFY")}/{S("F_OmzFY")}*100,1)&"%","n.b."),"n.b.")',
         f'="Δ budget: "&IF(AND(ISNUMBER({S("F_BorFY")}),ISNUMBER({S("F_BorFY","Budget")})),FIXED(({S("F_BorFY")}-{S("F_BorFY","Budget")})/1000000,1)&" mln","n.b.")')
    tile(ws, 8, 9, r, "Realisatie YTD (zelfde peildatum)", f'=IF(ISNUMBER({S("F_ActOmz")}),{S("F_ActOmz")}/1000000,{S("F_ActOmz")})', MLN,
         f'=IF(AND(ISNUMBER({S("F_ActOmz")}),ISNUMBER({S("F_OmzFY")})),IF({S("F_OmzFY")}<>0,FIXED({S("F_ActOmz")}/{S("F_OmzFY")}*100,0)&"% van FY-prognose","n.b."),"voortgang n.b.")',
         f'=IF(AND(ISNUMBER({S("F_ActOmz")}),ISNUMBER({S("F_OmzFY")})),IF({S("F_OmzFY")}<>0,{S("F_ActOmz")}/{S("F_OmzFY")},0),0)')
    ws.cell(row=r + 3, column=8).number_format = '0%'
    ws.conditional_formatting.add(f"H{r+3}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=YELLOW, showValue=True))
    tile(ws, 11, 12, r, "Omzet to go (FY − realisatie)",
         f'=IF({S("F_ActOmz")}="n.v.t.",IF(ISNUMBER({S("F_OmzFY")}),{S("F_OmzFY")}/1000000,"n.b."),IF(AND(ISNUMBER({S("F_OmzFY")}),ISNUMBER({S("F_ActOmz")})),({S("F_OmzFY")}-{S("F_ActOmz")})/1000000,"n.b."))', MLN,
         f'="BOR to go: "&IF({S("F_ActBor")}="n.v.t.",{fx_mln(S("F_BorFY"))},IF(AND(ISNUMBER({S("F_BorFY")}),ISNUMBER({S("F_ActBor")})),FIXED(({S("F_BorFY")}-{S("F_ActBor")})/1000000,1),"n.b."))&" mln"',
         '="nooit oude prognose − nieuwe actual"')
    tile(ws, 14, 15, r, '="Woningverkopen "&Sel_Doeljaar&" · "&Sel_Versie', f'={S("F_WonFY")}', '#,##0" st.";[Red]-#,##0" st.";"–"',
         f'="actual YTD: "&{S("F_WonAct")}', f'="budget: "&{S("F_WonFY","Budget")}')
    etal = 'SUMIFS(V_W_Etalage,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)'
    tile(ws, 17, 18, r, "Etalage (upload, onverkocht)", f"={etal}", '#,##0" st.";[Red]-#,##0" st.";"–"',
         '="verkocht: "&IF(SUMIFS(V_W_Aantal,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)=0,"n.b.",FIXED(SUMIFS(V_W_Cum_T,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)/SUMIFS(V_W_Aantal,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)*100,0)&"%")',
         '="verliesprojecten (upload): "&SUMIFS(V_Verlies,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)')
    # grafiekdata
    D0 = 50
    section(ws, D0 - 2, 2, "Grafiekdata (gekoppeld aan CALC_Versies / Projecten / VERW_Project — niet overschrijven)", 17)
    put(ws, f"B{D0-1}", "A. Jaarprognoseverloop (EUR mln)", bold=True, color=NAVY)
    header(ws, D0, 2, ["Versie", "Label", "Omzet FY", "BOR FY", "BOR %"], height=18)
    for i, v in enumerate(VERSIES + ["Upload"]):
        rr = D0 + 1 + i
        put(ws, f"B{rr}", v)
        o, b = S("F_OmzFY", v), S("F_BorFY", v)
        ws[f"C{rr}"] = f'=B{rr}&IF(ISNUMBER({o}),""," (n.b.)")' if v != "Upload" else f'="Upload "&P_PeilTxt&IF(ISNUMBER({o}),""," (n.b.)")'
        ws[f"D{rr}"] = f"=IF(ISNUMBER({o}),{o}/1000000,0)"
        ws[f"E{rr}"] = f"=IF(ISNUMBER({b}),{b}/1000000,0)"
        ws[f"F{rr}"] = f'=IF(AND(ISNUMBER({o}),ISNUMBER({b})),IF({o}<>0,{b}/{o},"n.b."),"n.b.")'
        for c in "DE":
            ws[f"{c}{rr}"].number_format = FMT_EUR1
        ws[f"F{rr}"].number_format = FMT_PCT
    put(ws, f"H{D0-1}", "B. Top 10 projecten: realisatie vs to go (EUR mln, geselecteerde versie)", bold=True, color=NAVY)
    header(ws, D0, 8, ["#", "Project", "Realisatie YTD", "To go", "FY"], height=18)
    for k in range(1, 11):
        rr = D0 + k
        put(ws, f"H{rr}", k)
        ix = f"IFERROR(MATCH(LARGE(Projecten!$U${P0}:$U${P1},{k}),Projecten!$U${P0}:$U${P1},0),0)"
        ws[f"Q{rr}"] = f"={ix}"
        helper_style(ws[f"Q{rr}"])
        q = f"$Q{rr}"
        ws[f"I{rr}"] = f'=IF({q}=0,"",INDEX(Projecten!$B${P0}:$B${P1},{q}))'
        ws[f"J{rr}"] = f'=IF({q}=0,0,IF(ISNUMBER(INDEX(Projecten!$L${P0}:$L${P1},{q})),INDEX(Projecten!$L${P0}:$L${P1},{q})*P_Factor/1000000,0))'
        ws[f"L{rr}"] = f'=IF({q}=0,0,INDEX(Projecten!$K${P0}:$K${P1},{q})*P_Factor/1000000)'
        ws[f"K{rr}"] = f'=IF({q}=0,0,IF(ISNUMBER(INDEX(Projecten!$M${P0}:$M${P1},{q})),INDEX(Projecten!$M${P0}:$M${P1},{q})*P_Factor/1000000,0))'
        for c in "JKL":
            ws[f"{c}{rr}"].number_format = FMT_EUR1
    put(ws, f"N{D0-1}", "C. Grootste mutaties vs vorige versie (EUR mln)", bold=True, color=NAVY)
    header(ws, D0, 14, ["#", "Project", "Δ omzet FY"], height=18)
    for k in range(1, 9):
        rr = D0 + k
        put(ws, f"N{rr}", k)
        ix = f"IFERROR(MATCH(LARGE(Projecten!$V${P0}:$V${P1},{k}),Projecten!$V${P0}:$V${P1},0),0)"
        ws[f"R{rr}"] = f"={ix}"
        helper_style(ws[f"R{rr}"])
        q = f"$R{rr}"
        ws[f"O{rr}"] = f'=IF({q}=0,"",INDEX(Projecten!$B${P0}:$B${P1},{q}))'
        ws[f"P{rr}"] = f'=IF({q}=0,0,INDEX(Projecten!$Q${P0}:$Q${P1},{q})*P_Factor/1000000)'
        ws[f"P{rr}"].number_format = FMT_EUR2
    put(ws, f"N{D0+10}", '=IF(Sel_Prev="","Budget heeft geen vorige versie: geen mutaties.",IF(COUNT(Projecten!$V$%d:$V$%d)=0,"Geen mutaties beschikbaar (vorige versie n.b.).",""))' % (P0, P1), color=GREY, italic=True, size=8)
    M0 = D0 + 13
    put(ws, f"B{M0-1}", "D. Meerjarendoorkijk huidige upload (EUR mln, niet vastgelegd)", bold=True, color=NAVY)
    header(ws, M0, 2, ["Jaar", "Omzet", "BOR"], height=18)
    yrs = [("=P_RJ&\" FY\"", "OmzFY", "BorFY"), ("=(P_RJ+1)&\"\"", "OmzFY1", "BorFY1")] + \
          [(f"=(P_RJ+{o})&\"\"", f"OmzRJ{o}", f"BorRJ{o}") for o in range(2, 9)] + [("Onverdeeld", "OmzONV", "BorONV")]
    for i, (lab, o, b) in enumerate(yrs):
        rr = M0 + 1 + i
        ws[f"B{rr}"] = lab
        ws[f"C{rr}"] = f"=SUMIFS(V_{o},V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)/1000000"
        ws[f"D{rr}"] = f"=SUMIFS(V_{b},V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)/1000000"
        for c in "CD":
            ws[f"{c}{rr}"].number_format = FMT_EUR1
    Mlast = M0 + len(yrs)
    # grafieken met dynamische titels in cellen
    gt = [("B12", '="Jaarprognoseverloop "&Sel_Doeljaar&": FY-omzet en BOR per vastgelegde versie (EUR mln)"'),
          ("K12", '="Realisatie vs resterende prognose — top 10 projecten, "&Sel_Label&" (EUR mln)"'),
          ("B30", '="Grootste projectmutaties "&Sel_Versie&" vs "&IF(Sel_Prev="","–",Sel_Prev)&" (Δ FY-omzet, EUR mln)"'),
          ("K30", '="Meerjarendoorkijk upload "&P_PeilTxt&" (EUR mln)"')]
    for ref, f in gt:
        ws[ref] = f
        ws[ref].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=4, min_row=D0, max_col=5, max_row=D0 + 6), Reference(ws, min_col=3, min_row=D0 + 1, max_row=D0 + 6),
        "B13", [NAVY, YELLOW], labels=True, w=17, h=8)
    bar(ws, Reference(ws, min_col=10, min_row=D0, max_col=11, max_row=D0 + 10), Reference(ws, min_col=9, min_row=D0 + 1, max_row=D0 + 10),
        "K13", [NAVY, YELLOW], kind="bar", grouping="stacked", w=17, h=8)
    ch = bar(ws, Reference(ws, min_col=16, min_row=D0, max_col=16, max_row=D0 + 8), Reference(ws, min_col=15, min_row=D0 + 1, max_row=D0 + 8),
             "B31", [NAVY2], kind="bar", w=17, h=8, legend=False, numfmt="#,##0.00", labels=True)
    ch.series[0].invertIfNegative = False
    bar(ws, Reference(ws, min_col=3, min_row=M0, max_col=4, max_row=Mlast), Reference(ws, min_col=2, min_row=M0 + 1, max_row=Mlast),
        "K31", [NAVY, YELLOW], w=17, h=8)
    put(ws, "B47", "Kolommen met '(n.b.)' in het label hebben geen vastgelegde versie: de staaf ontbreekt bewust (niet nul). Realisatie (donkerblauw) en resterende prognose (geel) zijn gescheiden.",
        color=GREY, italic=True, size=8)
    widths(ws, {c: 9.5 for c in "BCDEFGHIJKLMNOPQRS"})
    for c in "DGJMP":
        pass
    ws.column_dimensions["I"].width = 22
    ws.column_dimensions["O"].width = 22
    ws.column_dimensions["C"].width = 14
    landscape(ws, "A1:S47")


# ---------------------------------------------------------------------------------------------
# Woningen
# ---------------------------------------------------------------------------------------------
def build_woningen(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, None, None, 19)
    ws["B1"] = '="Woningverkopen "&P_Regio&" — "&P_RJ&" en "&(P_RJ+1)&" (aantallen op aandeelbasis)"'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = '="Bron: Woningetalage, peildatum "&P_PeilTxt&" · filter: "&Sel_FiltLabel&" · verkopen ≠ opleveringen/transporten/startbouw · projecten met Woningen meetellen = Nee uitgesloten"'
    ws["B2"].font = font(9, False, GREY, True)
    crit = 'V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS'
    sw = lambda n: f"SUMIFS(V_W_{n},{crit})"  # noqa: E731
    r = 4
    ST = '#,##0" st.";[Red]-#,##0" st.";"–"'
    tile(ws, 2, 3, r, "Totaal woningen (aandeelbasis)", f"={sw('Aantal')}", ST, f'="bruto (afgeleid): "&FIXED({sw("Bruto")},0)', '="niet nogmaals × aandeel"')
    tile(ws, 5, 6, r, '="Verkocht t/m "&(P_RJ-1)', f"={sw('VJ_T')}", ST, "actual")
    tile(ws, 8, 9, r, '="Verkocht "&P_RJ&" t/m "&P_PeilTxt', f"={sw('ActYTD')}", ST, "actual (hjr t/m peilmaand)")
    tile(ws, 11, 12, r, '="Jaarverwachting "&P_RJ', f"={sw('HuidigJaar')}", ST, '="waarvan prognose: "&' + sw("PrognoseRest"), '="volgend jaar: "&' + sw("VolgendJaar"))
    tile(ws, 14, 15, r, "Cumulatief verkocht", f"={sw('Cum_T')}", ST, f'=IF({sw("Aantal")}=0,"n.b.",FIXED({sw("Cum_T")}/{sw("Aantal")}*100,0)&"% van totaal")',
         f'=IF({sw("Aantal")}=0,0,{sw("Cum_T")}/{sw("Aantal")})')
    ws.cell(row=r + 3, column=14).number_format = "0%"
    ws.conditional_formatting.add(f"N{r+3}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=YELLOW, showValue=True))
    tile(ws, 17, 18, r, "Resterende etalage", f"={sw('Etalage')}", ST, "totaal − cumulatief actual")
    # per maand
    r = 10
    section(ws, r, 2, "Verkopen per maand en kwartaal (actual t/m peilmaand, daarna prognose)", 8)
    header(ws, r + 1, 2, ["Periode", "Actual", "Prognose", "Totaal", "Status"], height=18)
    m0 = r + 2
    for m in range(1, 13):
        rr = m0 + m - 1
        ws[f"B{rr}"] = f'=P_RJ&"-"&RIGHT("0{m}",2)'
        ws[f"C{rr}"] = f'=IF({m}<=P_PeilMaand,{sw(f"M{m:02d}")},0)'
        ws[f"D{rr}"] = f'=IF({m}>P_PeilMaand,{sw(f"M{m:02d}")},0)'
        ws[f"E{rr}"] = f"=C{rr}+D{rr}"
        ws[f"F{rr}"] = f'=IF({m}<=P_PeilMaand,"actual","prognose")'
    q0 = m0 + 13
    header(ws, q0 - 1, 2, ["Kwartaal", "Actual", "Prognose", "Totaal", "Status"], height=18)
    for q in range(1, 5):
        rr = q0 + q - 1
        ws[f"B{rr}"] = f'=P_RJ&" Q{q}"'
        ws[f"C{rr}"] = f"=SUM(C{m0+3*(q-1)}:C{m0+3*q-1})"
        ws[f"D{rr}"] = f"=SUM(D{m0+3*(q-1)}:D{m0+3*q-1})"
        ws[f"E{rr}"] = f"=C{rr}+D{rr}"
        ws[f"F{rr}"] = f'=IF(D{rr}=0,"actual",IF(C{rr}=0,"prognose","deels actual"))'
    for q in range(1, 5):
        rr = q0 + 3 + q
        ws[f"B{rr}"] = f'=(P_RJ+1)&" Q{q}"'
        ws[f"C{rr}"] = 0
        ws[f"D{rr}"] = f"={sw(f'VQ{q}')}"
        ws[f"E{rr}"] = f"=C{rr}+D{rr}"
        ws[f"F{rr}"] = "prognose"
    qlast = q0 + 7
    for rr in list(range(m0, m0 + 12)) + list(range(q0, qlast + 1)):
        for c in "CDE":
            ws[f"{c}{rr}"].number_format = FMT_INT
        ws[f"F{rr}"].font = font(8, False, GREY)
    ws[f"H{r+1}"] = '="Verkopen per kwartaal "&P_RJ&"–"&(P_RJ+1)&" (aantal): actual vs prognose"'
    ws[f"H{r+1}"].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=3, min_row=q0 - 1, max_col=4, max_row=qlast), Reference(ws, min_col=2, min_row=q0, max_row=qlast),
        f"H{r+2}", [NAVY, YELLOW], grouping="stacked", numfmt="#,##0", w=17, h=7.5)
    # versievergelijking
    r = qlast + 2
    section(ws, r, 2, "Jaarverwachting woningverkopen per prognoseversie (doeljaar volgens Dashboard)", 8)
    header(ws, r + 1, 2, ["Versie", "Label", "Verkopen FY", "Actual YTD", "Δ vs Budget"], height=18)
    S = lambda m, row: f"CALC_Versies!${CC[m]}${CR[row]}"  # noqa: E731
    v0 = r + 2
    for i, v in enumerate(VERSIES + ["Upload"]):
        rr = v0 + i
        put(ws, f"B{rr}", v)
        ws[f"C{rr}"] = f'=B{rr}&IF(ISNUMBER({S("F_WonFY", v)}),""," (n.b.)")'
        ws[f"D{rr}"] = f"={S('F_WonFY', v)}"
        ws[f"E{rr}"] = f"={S('F_WonAct', v)}"
        ws[f"F{rr}"] = f'=IF(AND(ISNUMBER(D{rr}),ISNUMBER($D${v0})),D{rr}-$D${v0},"n.b.")'
        ws[f"Q{rr}"] = f"=IF(ISNUMBER(D{rr}),D{rr},0)"
        helper_style(ws[f"Q{rr}"])
        for c in "DEF":
            ws[f"{c}{rr}"].number_format = FMT_INT
            ws[f"{c}{rr}"].alignment = Alignment(horizontal="right")
    put(ws, f"Q{v0-1}", "Verkopen FY", color=GREY, size=8)
    ws[f"H{r+1}"] = '="Jaarverwachting verkopen "&Sel_Doeljaar&" per versie (aantal)"'
    ws[f"H{r+1}"].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=17, min_row=v0 - 1, max_row=v0 + 5), Reference(ws, min_col=3, min_row=v0, max_row=v0 + 5),
        f"H{r+2}", [NAVY], numfmt="#,##0", w=17, h=6, labels=True, legend=False)
    put(ws, f"B{v0+7}", "Budget/Q1/Q2 zijn fictieve snapshots; hun realisatie (actual YTD) is niet vastgelegd en blijft n.b.", color=GREY, italic=True, size=8)
    # koperstype
    r = v0 + 14
    section(ws, r, 2, "Koperstype (Beleggers/Particulier alleen bekend voor actuals; prognose = aanvullende invoer of Onbekend)", 8)
    header(ws, r + 1, 2, ["Periode", "Beleggers", "Particulier", "Onbekend", "Totaal"], height=18)
    kt = lambda per, typ, off: (f'SUMIFS(KT_Aantal,KT_Telt,1,KT_Doel,P_RJ+{off},KT_Periode,"{per}",KT_Type,"{typ}",'  # noqa: E731
                                f'KT_TypeMap,Sel_TypeCrit,KT_WBS,Sel_WBS)')
    rows = [
        ('="t/m "&(P_RJ-1)&" (actual)"', sw("VJ_B"), sw("VJ_P"), "0", sw("VJ_T")),
        ('=P_RJ&" t/m "&P_PeilTxt&" (actual)"', sw("HJ_B"), sw("HJ_P"), "0", sw("HJ_T")),
        ('="Cumulatief (actual)"', sw("Cum_B"), sw("Cum_P"), "0", sw("Cum_T")),
    ]
    for i, (lab, b, p, o, t) in enumerate(rows):
        rr = r + 2 + i
        ws[f"B{rr}"] = lab
        ws[f"C{rr}"] = f"={b}"
        ws[f"D{rr}"] = f"={p}"
        ws[f"E{rr}"] = f"={o}"
        ws[f"F{rr}"] = f"={t}"
    rr = r + 5
    ws[f"B{rr}"] = '="Prognose rest "&P_RJ'
    bel = "+".join(kt(f"Q{q}", "Belegger", 0) for q in range(1, 5))
    par = "+".join(kt(f"Q{q}", "Particulier", 0) for q in range(1, 5))
    ws[f"C{rr}"] = f"={bel}"
    ws[f"D{rr}"] = f"={par}"
    ws[f"F{rr}"] = f"={sw('PrognoseRest')}"
    ws[f"E{rr}"] = f"=F{rr}-C{rr}-D{rr}"
    rr += 1
    ws[f"B{rr}"] = '="Prognose "&(P_RJ+1)'
    ws[f"C{rr}"] = "=" + "+".join(kt(f"Q{q}", "Belegger", 1) for q in range(1, 5))
    ws[f"D{rr}"] = "=" + "+".join(kt(f"Q{q}", "Particulier", 1) for q in range(1, 5))
    ws[f"F{rr}"] = f"={sw('VolgendJaar')}"
    ws[f"E{rr}"] = f"=F{rr}-C{rr}-D{rr}"
    for x in range(r + 2, rr + 1):
        for c in "CDEF":
            ws[f"{c}{x}"].number_format = FMT_INT
    # segment
    put(ws, f"H{r}", "Segment (uit Mapping › 'Segment demo' — fictieve indeling)", bold=True, color=NAVY)
    header(ws, r + 1, 8, ["Segment", "Totaal", "Cumulatief", "Etalage", '="Verw. "&P_RJ', '="Verw. "&(P_RJ+1)'], height=18)
    for i, seg in enumerate(["Grondgebonden", "Appartement"]):
        rr = r + 2 + i
        put(ws, f"H{rr}", seg)
        for c, n in zip("IJKLM", ["Aantal", "Cum_T", "Etalage", "HuidigJaar", "VolgendJaar"]):
            ws[f"{c}{rr}"] = f'=SUMIFS(V_W_{n},{crit},V_Segment,"{seg}")'
            ws[f"{c}{rr}"].number_format = FMT_INT
    rr = r + 4
    put(ws, f"H{rr}", "Overig / onbekend")
    for c, n in zip("IJKLM", ["Aantal", "Cum_T", "Etalage", "HuidigJaar", "VolgendJaar"]):
        ws[f"{c}{rr}"] = f"={sw(n)}-{c}{rr-2}-{c}{rr-1}"
        ws[f"{c}{rr}"].number_format = FMT_INT
    # projectmatrix
    r = r + 9
    section(ws, r, 2, "Projectmatrix woningen (upload) — % verkocht als voortgangsbalk", 17)
    hdr = ["Project", "Type", "Segment", "Aandeel %", "Totaal (aandeel)", "Bruto (afgeleid)", '="t/m "&(P_RJ-1)',
           '="Actual "&P_RJ', '="Verw. "&P_RJ', '="Verw. "&(P_RJ+1)', '=(P_RJ+2)&""', "Cumulatief", "Etalage", "% verkocht",
           '="Budget FY "&Sel_Doeljaar', '="Δ upload vs budget"']
    header(ws, r + 1, 2, hdr, height=30)
    p0 = r + 2
    for i in range(V1 - V0 + 1):
        rr = p0 + i
        vr = V0 + i
        a = f"VERW_Project!$A{vr}"
        V = lambda n: f"VERW_Project!${VCOLS[n]}{vr}"  # noqa: E731
        ws[f"B{rr}"] = f'=IF({a}="","",{V("Naam")})'
        ws[f"C{rr}"] = f'=IF({a}="","",{V("Type")})'
        ws[f"D{rr}"] = f'=IF({a}="","",{V("Segment")})'
        ws[f"E{rr}"] = f'=IF({a}="","",{V("Aandeel")})'
        nvt = f'{V("Meetellen")}="Nee"'
        for c, n in zip("FHIJKLMN", ["W_Aantal", "W_VJ_T", "W_ActYTD", "W_HuidigJaar", "W_VolgendJaar", "W_Y_RJ2", "W_Cum_T", "W_Etalage"]):
            ws[f"{c}{rr}"] = f'=IF({a}="","",IF({nvt},"n.v.t.",{V(n)}))'
        ws[f"G{rr}"] = f'=IF({a}="","",IF({nvt},"n.v.t.",IF(AND(ISNUMBER(F{rr}),ISNUMBER(E{rr})),IF(E{rr}>0,F{rr}/(E{rr}/100),"n.b."),"n.b.")))'
        ws[f"O{rr}"] = f'=IF({a}="","",IF(AND(ISNUMBER(F{rr}),ISNUMBER(M{rr})),IF(F{rr}>0,M{rr}/F{rr},""),""))'
        bud = ledger_proj(a, '"Budget"', K_WFY)
        ws[f"P{rr}"] = f'=IF({a}="","",IF({nvt},"n.v.t.",{bud}))'
        ws[f"Q{rr}"] = f'=IF({a}="","",IF(AND(ISNUMBER(P{rr}),Sel_Doeljaar=P_RJ,ISNUMBER(J{rr})),J{rr}-P{rr},IF(AND(ISNUMBER(P{rr}),Sel_Doeljaar=P_RJ+1,ISNUMBER(K{rr})),K{rr}-P{rr},"n.b.")))'
        for c in "BCDEFGHIJKLMNOPQ":
            ws[f"{c}{rr}"].font = font(8)
            ws[f"{c}{rr}"].border = Border(bottom=thin)
            if c not in "BCD":
                ws[f"{c}{rr}"].number_format = FMT_INT
                ws[f"{c}{rr}"].alignment = Alignment(horizontal="right")
        ws[f"O{rr}"].number_format = "0%"
        ws[f"G{rr}"].font = font(8, False, GREY, True)
    ws.conditional_formatting.add(f"O{p0}:O{p0+V1-V0}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=YELLOW, showValue=True))
    put(ws, f"B{p0-3}", "Bruto (afgeleid) = aandeelbasis ÷ aandeel% — alleen informatief; alle totalen gebruiken de aandeelbasis zonder nogmaals te wegen.", color=GREY, italic=True, size=8)
    widths(ws, {"B": 26, "C": 10, "D": 12, "E": 9, "F": 9, "G": 9, "H": 18, "I": 9, "J": 10, "K": 9, "L": 9, "M": 9, "N": 9, "O": 10, "P": 9, "Q": 9, "R": 9})
    landscape(ws, f"A1:R{p0+25}")


# ---------------------------------------------------------------------------------------------
# Projectdetail
# ---------------------------------------------------------------------------------------------
def build_projectdetail(wb, ws, first_name):
    base_sheet(ws, zoom=90)
    title(ws, None, None, 19)
    ws["B1"] = '="Projectdetail — "&Det_Project'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = '="Doeljaar "&Sel_Doeljaar&" · versiejaar "&P_VJ&" · bedragen in "&P_Eenheid&" · upload peildatum "&P_PeilTxt'
    ws["B2"].font = font(9, False, GREY, True)
    put(ws, "B4", "Kies project:", bold=True, color=NAVY)
    c = ws["C4"]
    c.value = first_name
    input_style(c)
    c.font = font(10, True, "1F4E9A")
    ws.merge_cells("C4:F4")
    dv_list(ws, "C4", "L_ProjectenDet")
    add_name(wb, "Det_Project", "Projectdetail!$C$4")
    ws["H4"] = "=IFERROR(MATCH(Det_Project,V_Naam,0),0)"
    helper_style(ws["H4"])
    add_name(wb, "Det_Idx", "Projectdetail!$H$4")
    ws["I4"] = '=IF(Det_Idx=0,"(project niet gevonden)",INDEX(V_WBS,Det_Idx))'
    helper_style(ws["I4"])
    add_name(wb, "Det_WBS", "Projectdetail!$I$4")
    V = lambda n: f'IF(Det_Idx=0,"n.b.",INDEX(V_{n},Det_Idx))'  # noqa: E731
    card = [("WBS-element", V("WBS"), None), ("Projectdefinitie", V("Projdef"), None), ("Type", V("Type"), None),
            ("IFRS J/N", V("IFRS"), None), ("Segment (demo)", V("Segment"), None), ("Aandeel % (info)", V("Aandeel"), "0"),
            ("Marktsegment", V("I_Marktsegment"), None), ("Status WBS", V("I_StatusWBS"), None),
            ("Kans %", V("I_Kans"), "0"), ("Risicoklasse", V("I_Risico"), "0"), ("POC %", V("I_POC"), "0.0"),
            ("Orderintake", V("I_Orderintake"), FMT_DATE), ("Start verkoop", V("I_StartVerk"), FMT_DATE),
            ("Start bouw", V("I_StartBouw"), FMT_DATE), ("Maand oplevering (bron)", V("I_Oplevering"), "0"),
            ("Eind oplevering", V("I_EindOpl"), FMT_DATE)]
    for i, (lab, f, fmt) in enumerate(card):
        rr = 6 + i % 8
        cc = 2 if i < 8 else 5
        put(ws, (rr, cc), lab, color=GREY, size=8)
        cell = ws.cell(row=rr, column=cc + 1, value="=" + f)
        cell.font = font(9, True, NAVY)
        cell.alignment = Alignment(horizontal="left")
        if fmt:
            cell.number_format = fmt
        for x in range(cc, cc + 2):
            ws.cell(row=rr, column=x).border = Border(bottom=thin)
    # versieverloop
    r = 15
    section(ws, r, 2, "Jaarprognoseverloop (vastgelegde versies) en upload", 8)
    header(ws, r + 1, 2, ["Versie", "Omzet FY", "BOR FY", "Marge %", "Realisatie YTD", "To go", "Woningverk. FY", "Toelichting"], height=26)
    d0 = r + 2
    for i, v in enumerate(VERSIES + ["Upload"]):
        rr = d0 + i
        put(ws, f"B{rr}", v if v != "Upload" else "Upload (concept)", bold=v == "Upload")
        if v == "Upload":
            o = f'IF(Sel_Doeljaar=P_RJ,{V("OmzFY")},IF(Sel_Doeljaar=P_RJ+1,{V("OmzFY1")},"n.b."))'
            b = f'IF(Sel_Doeljaar=P_RJ,{V("BorFY")},IF(Sel_Doeljaar=P_RJ+1,{V("BorFY1")},"n.b."))'
            a = f'IF(Sel_Doeljaar=P_RJ,{V("OmzYTD")},"n.v.t.")'
            w = f'IF(Sel_Doeljaar=P_RJ,{V("W_HuidigJaar")},IF(Sel_Doeljaar=P_RJ+1,{V("W_VolgendJaar")},"n.b."))'
            toel = '""'
        else:
            o = ledger_proj("Det_WBS", f'"{v}"', K_OMZ)
            b = ledger_proj("Det_WBS", f'"{v}"', K_BOR)
            a = f'IF(Sel_Doeljaar>P_VJ,"n.v.t.",{ledger_proj("Det_WBS", chr(34) + v + chr(34), K_OMZA)})'
            w = ledger_proj("Det_WBS", f'"{v}"', K_WFY)
            toel = f'IFERROR(INDEX(T_Txt,MATCH(Det_WBS&"|"&Sel_Doeljaar&"|"&P_VJ&"|{v}",T_Key,0)),"")'
        ws[f"C{rr}"] = f"=IF(ISNUMBER({o}),{o}/P_Factor,{o})"
        ws[f"D{rr}"] = f"=IF(ISNUMBER({b}),{b}/P_Factor,{b})"
        ws[f"E{rr}"] = f'=IF(AND(ISNUMBER(C{rr}),ISNUMBER(D{rr})),IF(C{rr}<>0,D{rr}/C{rr},"n.b."),"n.b.")'
        ws[f"F{rr}"] = f"=IF(ISNUMBER({a}),{a}/P_Factor,{a})"
        ws[f"G{rr}"] = f'=IF(F{rr}="n.v.t.",C{rr},IF(AND(ISNUMBER(C{rr}),ISNUMBER(F{rr})),C{rr}-F{rr},"n.b."))'
        ws[f"H{rr}"] = f"={w}"
        ws[f"I{rr}"] = f"={toel}"
        ws[f"Q{rr}"] = f"=IF(ISNUMBER(C{rr}),C{rr},0)"
        ws[f"R{rr}"] = f"=IF(ISNUMBER(D{rr}),D{rr},0)"
        ws[f"P{rr}"] = f'=B{rr}&IF(ISNUMBER(C{rr}),""," (n.b.)")'
        for cc_ in "PQR":
            helper_style(ws[f"{cc_}{rr}"])
        for cc_ in "CDFGH":
            ws[f"{cc_}{rr}"].number_format = FMT_EUR if cc_ != "H" else FMT_INT
            ws[f"{cc_}{rr}"].alignment = Alignment(horizontal="right")
        ws[f"E{rr}"].number_format = FMT_PCT
        ws[f"E{rr}"].alignment = Alignment(horizontal="right")
        ws[f"F{rr}"].fill = fill("EEF3EA")
        ws[f"F{rr}"].font = font(9, False, "3B5B33", True)
        for cc_ in "BCDEFGHI":
            ws[f"{cc_}{rr}"].border = Border(bottom=thin)
    ws.conditional_formatting.add(f"B{d0}:I{d0+4}", FormulaRule(formula=[f"$B{d0}=Sel_Versie"], fill=fill(YELLOW_LT)))
    put(ws, "Q%d" % (d0 - 1), "Omzet FY", color=GREY, size=8)
    put(ws, "R%d" % (d0 - 1), "BOR FY", color=GREY, size=8)
    ws[f"K{r+1}"] = '="Omzet en BOR "&Sel_Doeljaar&" per versie ("&P_Eenheid&")"'
    ws[f"K{r+1}"].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=17, min_row=d0 - 1, max_col=18, max_row=d0 + 5), Reference(ws, min_col=16, min_row=d0, max_row=d0 + 5),
        f"K{r+2}", [NAVY, YELLOW], numfmt="#,##0", w=14, h=6)
    # planning
    r = d0 + 14
    section(ws, r, 2, "Planning omzet en BOR volgens upload (managementteken; realisatie apart van prognose)", 17)
    labs = ['="Realisatie t/m "&P_PeilTxt', '="Q4 "&P_RJ&" (rest)"'] + [f'="Q{q} "&(P_RJ+1)' for q in range(1, 5)] + \
           [f'=(P_RJ+{o})&""' for o in range(2, 9)] + ["Onverdeeld", "Totaal resterend"]
    header(ws, r + 1, 2, ["Kengetal"] + labs, height=30)
    for i, (kg, pre) in enumerate([("Omzet", "Omz"), ("BOR", "Bor")]):
        rr = r + 2 + i
        put(ws, f"B{rr}", kg, bold=True)
        fields = [f"{pre}YTD", f"{pre}Q4"] + [f"{pre}RJ1_Q{q}" for q in range(1, 5)] + [f"{pre}RJ{o}" for o in range(2, 9)] + [f"{pre}ONV"]
        for j, fld in enumerate(fields):
            cell = ws.cell(row=rr, column=3 + j, value=f'=IF(ISNUMBER({V(fld)}),{V(fld)}/P_Factor,{V(fld)})')
            cell.number_format = FMT_EUR
            cell.alignment = Alignment(horizontal="right")
        last = col(3 + len(fields))
        ws[f"{last}{rr}"] = f"=SUM(D{rr}:{col(2+len(fields))}{rr})"
        ws[f"{last}{rr}"].number_format = FMT_EUR
        ws[f"{last}{rr}"].font = font(9, True)
        ws[f"C{rr}"].fill = fill("EEF3EA")
        ws[f"C{rr}"].font = font(9, False, "3B5B33", True)
    # woningverkopen per maand
    r = r + 6
    section(ws, r, 2, "Woningverkopen project (aantal, aandeelbasis)", 17)
    labs = [f'=P_RJ&"-{m:02d}"' for m in range(1, 13)] + [f'="Q{q} "&(P_RJ+1)' for q in range(1, 5)] + ['=(P_RJ+2)&""']
    header(ws, r + 1, 2, ["Periode"] + labs, height=30)
    put(ws, f"B{r+2}", "Actual", bold=True)
    put(ws, f"B{r+3}", "Prognose", bold=True)
    for j in range(17):
        cc = 3 + j
        if j < 12:
            fld = f"W_M{j+1:02d}"
            ws.cell(row=r + 2, column=cc, value=f'=IF({j+1}<=P_PeilMaand,{V(fld)},0)')
            ws.cell(row=r + 3, column=cc, value=f'=IF({j+1}>P_PeilMaand,{V(fld)},0)')
        else:
            fld = f"W_VQ{j-11}" if j < 16 else "W_Y_RJ2"
            ws.cell(row=r + 2, column=cc, value=0)
            ws.cell(row=r + 3, column=cc, value=f"={V(fld)}")
        for rr in (r + 2, r + 3):
            ws.cell(row=rr, column=cc).number_format = FMT_INT
    # transponeren voor grafiek
    g0 = r + 6
    put(ws, f"T{g0-1}", "grafiekdata", color=GREY, size=8)
    put(ws, f"U{g0-1}", "Actual", color=GREY, size=8)
    put(ws, f"V{g0-1}", "Prognose", color=GREY, size=8)
    for j in range(17):
        ws[f"T{g0+j}"] = f"={col(3+j)}{r+1}"
        ws[f"U{g0+j}"] = f"={col(3+j)}{r+2}"
        ws[f"V{g0+j}"] = f"={col(3+j)}{r+3}"
        for cc_ in "TUV":
            helper_style(ws[f"{cc_}{g0+j}"])
    ws[f"B{r+5}"] = '="Woningverkopen per periode (aantal): actual t/m "&P_PeilTxt&", daarna prognose"'
    ws[f"B{r+5}"].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=21, min_row=g0 - 1, max_col=22, max_row=g0 + 16), Reference(ws, min_col=20, min_row=g0, max_row=g0 + 16),
        f"B{r+6}", [NAVY, YELLOW], grouping="stacked", numfmt="#,##0", w=20, h=6.5)
    # woningsamenvatting
    put(ws, f"N{r+5}", "Samenvatting woningen", bold=True, color=NAVY)
    summ = [("Totaal (aandeelbasis)", "W_Aantal"), ("Verkocht t/m vorig jaar", "W_VJ_T"), ("Verkocht huidig jaar (actual)", "W_HJ_T"),
            ("Cumulatief verkocht", "W_Cum_T"), ("Etalage", "W_Etalage"), ("Beleggers cumulatief", "W_Cum_B"), ("Particulier cumulatief", "W_Cum_P")]
    for i, (lab, fld) in enumerate(summ):
        rr = r + 6 + i
        put(ws, f"N{rr}", lab, color=GREY, size=8)
        ws[f"Q{rr}"] = f"={V(fld)}"
        ws[f"Q{rr}"].number_format = FMT_INT
        ws[f"Q{rr}"].font = font(9, True, NAVY)
    widths(ws, {"B": 20, "C": 13, "D": 10, "E": 10, "F": 13, "G": 10, "H": 10, "I": 10, "J": 9, "K": 9, "L": 9, "M": 9,
                "N": 9, "O": 9, "P": 9, "Q": 9, "R": 9, "S": 9, "T": 9, "U": 7, "V": 7})
    landscape(ws, f"A1:S{r+22}")


# ---------------------------------------------------------------------------------------------
# Controle
# ---------------------------------------------------------------------------------------------
YM_ROWS = (6, 14)


def build_controle(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, "Controleblad", "Status per controle: FOUT = eerst oplossen (niet vastleggen) · WAARSCHUWING = beoordelen · OK · INFO.", 9)
    zc = lambda n: f"STG_ZPMW03!${ZCOLS[n]}${STG0}:${ZCOLS[n]}${STG1}"  # noqa: E731
    ec = lambda n: f"STG_Woningetalage!${ECOLS[n]}${STG0}:${ECOLS[n]}${STG1}"  # noqa: E731
    put(ws, "B4", "Aantal FOUT", bold=True)
    put(ws, "B5", "Aantal WAARSCHUWING", bold=True)
    put(ws, "B6", "Selectie bij controle", color=GREY)
    ws["D6"] = '="doeljaar "&Sel_Doeljaar&", versie "&Sel_Label&", filter "&Sel_FiltLabel'
    ws["D6"].font = font(8, False, GREY)
    R = lambda k: RDFO_ROWS[k]  # noqa: E731
    zlast = col(1 + len(z_fields()))
    elast = col(1 + len(e_fields()))
    upFY = lambda f: f'IF(Sel_Doeljaar=P_RJ,SUM(V_{f}),SUM(V_{f}1))'  # noqa: E731
    diff = lambda kg, dj: f'SUMPRODUCT(ABS(C_Verschil)*(C_Kengetal="{kg}")*(C_Doeljaar={dj}))'  # noqa: E731
    frozen = "Vastleggen!$N$3"
    checks = [
        ("Bron", "Regels ZPMW03 geplakt", f'=COUNTIF({zc("WBS")},"?*")', '="verwacht "&P_RegelsZ', 'IF({v}=P_RegelsZ,"OK","FOUT")',
         "Vergelijk met aantal regels in het exportbestand (Parameters › P_RegelsZ)."),
        ("Bron", "Regels Woningetalage geplakt", f'=COUNTIF({ec("Stufe3")},"?*")', '="verwacht "&P_RegelsE', 'IF({v}=P_RegelsE,"OK","FOUT")', "Idem P_RegelsE."),
        ("Bron", "ZPMW03-regels zonder WBS-element", f'=COUNTA(SRC_ZPMW03!$A$2:$A${N_SRC+1})-COUNTIF({zc("WBS")},"?*")', "0", 'IF({v}=0,"OK","FOUT")',
         "Regels zonder sleutel kunnen niet gekoppeld worden en vallen anders stil weg."),
        ("Bron", "Woningetalage-regels zonder Stufe 3", f'=COUNTA(SRC_Woningetalage!$A$2:$A${N_SRC+1})-COUNTIF({ec("Stufe3")},"?*")', "0", 'IF({v}=0,"OK","FOUT")', "Idem."),
        ("Bron", f"Regels voorbij capaciteit ({N_SRC} per bron)", f'=COUNTA(SRC_ZPMW03!$A${N_SRC+2}:$A$20000)+COUNTA(SRC_Woningetalage!$A${N_SRC+2}:$A$20000)', "0",
         'IF({v}=0,"OK","FOUT")', "Formulebereiken STG_* uitbreiden (zie Handleiding)."),
        ("Kolommen", "ZPMW03: ontbrekende kolommen", f'=COUNTIF(STG_ZPMW03!$B$4:${zlast}$4,"ONTBREEKT")', "0", 'IF({v}=0,"OK","FOUT")',
         "Zie Veldwoordenboek (B). Aliassen Bedrijfsnaam / Naam maatschappij worden beide ondersteund."),
        ("Kolommen", "Woningetalage: ontbrekende kolommen", f'=COUNTIF(STG_Woningetalage!$B$4:${elast}$4,"ONTBREEKT")', "0", 'IF({v}=0,"OK","FOUT")',
         "Zie Veldwoordenboek (C)."),
        ("Kolommen", "Woningetalage: jaarheaders uit mapping niet gevonden", f'=COUNTIF(Veldwoordenboek!$D${YM_ROWS[0]}:$D${YM_ROWS[1]},"NEE")', "0",
         'IF({v}=0,"OK","WAARSCHUWING")', "Jaarheader-mapping (Veldwoordenboek A) bijwerken."),
        ("Sleutels", "Dubbele WBS-elementen in ZPMW03", f'=COUNTIF({zc("Dubbel")},">1")', "0", 'IF({v}=0,"OK","FOUT")',
         "Een WBS mag één regel hebben; dubbel = dubbeltelling in de aggregatie."),
        ("Sleutels", "Dubbele Stufe 3 in Woningetalage", f'=COUNTIF({ec("Dubbel")},">1")', "0", 'IF({v}=0,"OK","FOUT")', "Idem."),
        ("Sleutels", "Dubbele WBS-elementen in Mapping", '=SUMPRODUCT((Map_WBS<>"")*(COUNTIF(Map_WBS,Map_WBS)>1))', "0", 'IF({v}=0,"OK","FOUT")', "Mapping moet uniek zijn."),
        ("Sleutels", "Dubbele projectnamen in Mapping", '=SUMPRODUCT((Map_Naam<>"")*(COUNTIF(Map_Naam,Map_Naam)>1))', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Projectfilters werken op naam; maak namen uniek."),
        ("Koppeling", "Nieuwe projecten: ZPMW03-WBS niet in Mapping", f'=COUNTIF({zc("InMapping")},0)', "0", 'IF({v}=0,"OK","FOUT")',
         "Toevoegen aan Mapping, anders ontbreekt het project in rapportage en vastlegging (lijst hieronder)."),
        ("Koppeling", "Vervallen / ontbrekend: Mapping-WBS niet in ZPMW03", '=COUNTIFS(V_WBS,"?*",V_nZ,0)', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Project afgesloten of ontbreekt in export? Historie blijft zichtbaar (lijst hieronder)."),
        ("Koppeling", "Woningetalage-Stufe 3 niet in ZPMW03", f'=COUNTIF({ec("InZPMW03")},0)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Lijst hieronder."),
        ("Koppeling", "Woningetalage-Stufe 3 niet in Mapping", f'=COUNTIF({ec("InMapping")},0)', "0", 'IF({v}=0,"OK","FOUT")', "Woningen worden anders niet gerapporteerd."),
        ("Koppeling", "ZPMW03-WBS niet in Woningetalage", f'=COUNTIF({zc("InEtalage")},0)', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Toegestaan voor projecten zonder woningen (Woningen meetellen = Nee)."),
        ("Koppeling", "Projectdefinitie (ZPMW03) ≠ Project (Woningetalage) bij zelfde WBS", f'=COUNTIF({ec("C_Projdef")},1)', "0", 'IF({v}=0,"OK","FOUT")',
         "Koppeling op WBS/Stufe 3 klopt niet met projectniveau."),
        ("Bedragen", "Som kwartalen ≠ jaartotaal (OP.BOB/OP.BOR RJ en RJ+1)", f'=COUNTIF({zc("C_Kwartaal")},1)', "0", 'IF({v}=0,"OK","FOUT")',
         "Kwartaalvelden en jaartotaal worden nooit bij elkaar opgeteld; ze moeten aansluiten."),
        ("Bedragen", "OP.BOB + Kosten ≠ OP.BOR (LJ, RJ, RJ+1)", f'=COUNTIF({zc("C_BobKostBor")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Bedragen", "OP.BOB RJ bevat kwartalen vóór peildatum (geen restant?)",
         f'=SUMPRODUCT((P_PeilMaand>=3)*((COUNTIF({zc("BOB_RJ_Q1")},"<0")+COUNTIF({zc("BOB_RJ_Q1")},">0"))))'
         f'+(P_PeilMaand>=6)*(COUNTIF({zc("BOB_RJ_Q2")},"<0")+COUNTIF({zc("BOB_RJ_Q2")},">0"))'
         f'+(P_PeilMaand>=9)*(COUNTIF({zc("BOB_RJ_Q3")},"<0")+COUNTIF({zc("BOB_RJ_Q3")},">0"))', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Demo-aanname: OP.BOB RJ = restant na peildatum. Afwijking → definitie in productie controleren (dubbeltelling met LJ!)."),
        ("Bedragen", "Portefeuille BOB ≠ som OP.BOB RJ..RJ+8 + onverdeeld", f'=COUNTIF({zc("C_Portefeuille")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', ""),
        ("Bedragen", "PEW opbrengst ≠ BOB VJ + BOB LJ + portefeuille", f'=COUNTIF({zc("C_PEW")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', ""),
        ("Woningen", "Aantal ≠ vorig jaar + 12 mnd + volgend jaar + RJ+2", f'=COUNTIF({ec("C_Aantal")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Woningen", "Etalage ≠ aantal − cumulatief actual", f'=COUNTIF({ec("C_Etalage")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Woningen", "Cumulatief ≠ t/m vorig jaar + huidig jaar actual", f'=COUNTIF({ec("C_Cum")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Woningen", "Huidig jr Totaal ≠ som hjr-maanden t/m peilmaand", f'=COUNTIF({ec("C_HJ")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Controleert ook de peilmaand."),
        ("Woningen", "Beleggers + Particulier ≠ Totaal", f'=COUNTIF({ec("C_BP")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Woningen", "Huidig/Volgend jaar ≠ som maanden/kwartalen", f'=COUNTIF({ec("C_Jaren")},1)', "0", 'IF({v}=0,"OK","FOUT")', ""),
        ("Woningen", "Niet-geïnterpreteerde kolommen bevatten waarden", f'=COUNTIF({ec("C_NietGeinterp")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Rond, RJ+3..RJ+6, negatieve jaarheaders: eerst betekenis vastleggen in Veldwoordenboek."),
        ("Woningen", "Aantal woningen wijkt af van Mapping", f'=COUNTIF({ec("C_MapAantal")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', ""),
        ("Woningen", "Woningen bij project met 'Woningen meetellen = Nee'", f'=COUNTIF({ec("C_Meetellen")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', ""),
        ("Weging", "Woningen niet dubbel gewogen: rapporttotaal = brontotaal", '=SUMIFS(V_W_Aantal,V_Meetellen,"Ja")', f'=SUM({ec("Aantal")})',
         'IF(ABS({v}-{n})<=P_Tol,"OK","FOUT")', '="Met aandeel nogmaals gewogen zou het "&FIXED(SUMPRODUCT(IFERROR(V_W_Aantal*V_Aandeel/100,0)),0)&" zijn — wordt bewust NIET gedaan."'),
        ("Verlies", "Verliesprojecten in upload (BOR FY RJ < 0)", "=SUM(V_Verlies)", "info", '"INFO"', "Zie lijst hieronder; verlies blijft negatief in alle rapporten."),
        ("Verlies", "RDFO-BOR upload = signed som projecten (geen ABS)", f'=IF(ISNUMBER(RDFO!$H${R("bor_tot")}),RDFO!$H${R("bor_tot")}*P_Factor,"n.b.")',
         f'={upFY("BorFY")}', 'IF(ISNUMBER({v}),IF(ABS({v}-{n})<=P_Tol,"OK","FOUT"),"INFO")', '="Met ABS zou het "&FIXED(SUMPRODUCT(ABS(N(+V_BorFY))),0)&" zijn."'),
        ("Versies", "Dubbele regels in Vastlegging (genegeerd)", "=SUM(L_Dubbel)", "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Dubbel geplakt; telt niet dubbel. Documenteer in Importlog."),
        ("Versies", "Correctie (revisie > 0) zonder toelichting", '=COUNTIFS(L_Revisie,">0",L_Toelichting,"")', "0", 'IF({v}=0,"OK","FOUT")', "Toelichting is verplicht."),
        ("Versies", "Ongeldige status in Vastlegging", '=COUNTIF(L_WBS,"?*")-COUNTIFS(L_WBS,"?*",L_Status,"Vastgesteld")-COUNTIFS(L_WBS,"?*",L_Status,"Concept")', "0",
         'IF({v}=0,"OK","FOUT")', "Alleen Concept of Vastgesteld."),
        ("Versies", "Ongeldige versie in Vastlegging", '=COUNTIF(L_WBS,"?*")-' + "-".join(f'COUNTIFS(L_WBS,"?*",L_Versie,"{v}")' for v in VERSIES), "0",
         'IF({v}=0,"OK","FOUT")', "Alleen Budget/Q1/Q2/Q3/Q4."),
        ("Versies", "Concept-regels in Vastlegging (tellen niet mee)", '=COUNTIFS(L_WBS,"?*",L_Status,"Concept")', "info", '"INFO"', ""),
        ("Versies", "Actieve correcties (revisie > 0)", '=COUNTIFS(L_Revisie,">0",L_Actief,1)', "info", '"INFO"', "Gemarkeerd met ● in RDFO."),
        ("Versies", '="Uploadversie "&P_UploadVersie&" "&P_UploadVJ&" al vastgesteld (regels)"', f"={frozen}", "info", '"INFO"',
         "Ja: concept dient als aansluitcontrole; opnieuw plakken = correctie."),
        ("Aansluiting", '="Vastgelegde "&P_UploadVersie&"-snapshot ↔ upload: Omzet FY "&P_RJ', f"={diff(K_OMZ, 'P_RJ')}", "0",
         f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', "Som |verschil| over projecten (EUR). Vastgelegde waarden blijven leidend."),
        ("Aansluiting", '="Vastgelegde "&P_UploadVersie&"-snapshot ↔ upload: BOR FY "&P_RJ', f"={diff(K_BOR, 'P_RJ')}", "0",
         f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', ""),
        ("Aansluiting", '="Vastgelegde "&P_UploadVersie&"-snapshot ↔ upload: Omzet FY "&(P_RJ+1)', f"={diff(K_OMZ, 'P_RJ+1')}", "0",
         f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', ""),
        ("Aansluiting", '="Vastgelegde "&P_UploadVersie&"-snapshot ↔ upload: realisatie YTD "&P_RJ',
         f"={diff(K_OMZA, 'P_RJ')}+{diff(K_BORA, 'P_RJ')}", "0", f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', ""),
        ("Aansluiting", '="Vastgelegde "&P_UploadVersie&"-snapshot ↔ upload: woningen (alle kengetallen)"',
         "=" + "+".join(diff(k, "P_RJ") for k in [K_WFY, K_WVJ, K_WACT] + K_WQ) + f"+{diff(K_WFY, 'P_RJ+1')}", "0",
         f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', ""),
        ("Aansluiting", "Totaal |verschil| alle kengetallen", "=SUMPRODUCT(ABS(C_Verschil))", "0",
         f'IF({frozen}=0,"INFO",IF({{v}}<=P_Tol,"OK","WAARSCHUWING"))', "Details: Vastleggen kolom U."),
        ("Parameters", "Peildatum past bij uploadversie", '=P_PeilMaand', '=INDEX(V_Maand,MATCH(P_UploadVersie,V_List,0))',
         'IF(AND({v}={n},DAY(P_Peildatum+1)=1),"OK","WAARSCHUWING")', "Q1=31-03, Q2=30-06, Q3=30-09, Q4/Budget=31-12."),
        ("Parameters", "Jaar peildatum past bij versiejaar upload", '=YEAR(P_Peildatum)', '=IF(P_UploadVersie="Budget",P_UploadVJ-1,P_UploadVJ)',
         'IF({v}={n},"OK","WAARSCHUWING")', "Budget heeft peildatum eind vorig jaar."),
        ("Parameters", "Rapportagejaar = jaar van peildatum", "=P_RJ", "=YEAR(P_Peildatum)", 'IF({v}={n},"OK","WAARSCHUWING")', ""),
        ("Aanvullend", "Fictieve demo-aannames tellen mee", '=IF(P_Demo="Ja",COUNTIF(A_Status,"Fictieve aanname"),0)', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Rapport bevat zelfgekozen demowaarden. Zet P_Demo op Nee voor een rapport zonder aannames."),
        ("Aanvullend", "Zowel regio- als projectniveau voor dezelfde rubriek/versie", "=SUM(A_Conflict)", "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Risico op dubbeltelling."),
        ("Aanvullend", "Aanvullende invoer met eenheid ≠ EUR", '=COUNTIFS(A_Waarde,"<>",A_Eenheid,"<>EUR")', "0", 'IF({v}=0,"OK","FOUT")', "Alleen hele euro's invoeren."),
        ("Aanvullend", "RDFO-posten n.b. in geselecteerde versie", f'=COUNTIF(RDFO!$J$7:$J${R("ebit")},"n.b.")', "info", '"INFO"',
         "Ontbrekende aanvullende invoer wordt n.b. getoond, nooit nul."),
        ("Rapport", "RDFO upload-omzet = som VERW_Project", f'=IF(ISNUMBER(RDFO!$H${R("omz_proj")}),RDFO!$H${R("omz_proj")}*P_Factor,"n.b.")',
         f'={upFY("OmzFY")}', 'IF(ISNUMBER({v}),IF(ABS({v}-{n})<=P_Tol,"OK","FOUT"),"INFO")', ""),
        ("Rapport", "Projecten totaal (geselecteerd) = RDFO geselecteerd", f"=IFERROR(Projecten!$K$6*P_Factor,\"n.b.\")",
         f'=IFERROR(RDFO!$J${R("omz_proj")}*P_Factor,"n.b.")', 'IF(AND(ISNUMBER({v}),ISNUMBER({n})),IF(ABS({v}-{n})<=P_Tol,"OK","FOUT"),"INFO")',
         "FOUT kan ook komen door een actief autofilter op Projecten."),
        ("Rapport", "Dashboard-omzet = RDFO (alleen zonder filter)", f'=CALC_Versies!${CC["F_OmzFY"]}${CR["Sel"]}',
         f'=IFERROR(RDFO!$J${R("omz_proj")}*P_Factor,"n.b.")', 'IF(Sel_FiltLabel<>"geen filter","INFO",IF(AND(ISNUMBER({v}),ISNUMBER({n})),IF(ABS({v}-{n})<=P_Tol,"OK","FOUT"),"INFO"))', ""),
        ("Rapport", "BOR % = totale BOR / totale omzet", f'=RDFO!$J${R("bor_pct")}',
         f'=IFERROR((CALC_Versies!${CC["BorOps"]}${CR["Sel"]}+CALC_Versies!${CC["BorDln"]}${CR["Sel"]})/(CALC_Versies!${CC["OmzOps"]}${CR["Sel"]}+CALC_Versies!${CC["OmzDln"]}${CR["Sel"]}),"n.b.")',
         'IF(AND(ISNUMBER({v}),ISNUMBER({n})),IF(ABS({v}-{n})<0.000001,"OK","FOUT"),"INFO")', "Geen gemiddelde van projectpercentages."),
        ("Historie", "Vastgelegde versies zonder realisatie (blijft n.b.)", "=COUNTIFS(CALC_Versies!$C$5:$C$9,\">0\",CALC_Versies!$D$5:$D$9,0)", "info", '"INFO"',
         "Budget/Q1/Q2 hebben geen vastgelegde actuals; die worden niet als 0 getoond."),
    ]
    header(ws, 8, 2, ["Nr", "Categorie", "Controle", "Resultaat", "Norm", "Status", "Toelichting / actie"], height=22)
    r0 = 9
    for i, (cat, txt, res, norm, st, act) in enumerate(checks):
        r = r0 + i
        put(ws, f"B{r}", i + 1, align="center")
        put(ws, f"C{r}", cat, color=GREY)
        ws[f"D{r}"] = txt
        ws[f"E{r}"] = res
        ws[f"F{r}"] = norm if norm.startswith("=") else norm
        ws[f"G{r}"] = "=" + st.replace("{v}", f"E{r}").replace("{n}", f"F{r}")
        ws[f"H{r}"] = act
        for c in "BCDEFGH":
            ws[f"{c}{r}"].font = font(9, c == "G", GREY if c in "CH" else "000000")
            ws[f"{c}{r}"].border = Border(bottom=thin)
            ws[f"{c}{r}"].alignment = Alignment(vertical="top", wrap_text=c in "DH", horizontal="right" if c in "EF" else None)
        ws[f"E{r}"].number_format = '#,##0.##;[Red]-#,##0.##;0'
    rl = r0 + len(checks) - 1
    ws["D4"] = f'=COUNTIF($G${r0}:$G${rl},"FOUT")'
    ws["D5"] = f'=COUNTIF($G${r0}:$G${rl},"WAARSCHUWING")'
    for c in ("D4", "D5"):
        ws[c].font = font(14, True, NAVY)
        ws[c].alignment = Alignment(horizontal="left")
    ws.conditional_formatting.add("D4", FormulaRule(formula=["$D$4>0"], font=Font(color=RED, bold=True, size=14)))
    ws.conditional_formatting.add(f"G{r0}:G{rl}", FormulaRule(formula=[f'G{r0}="FOUT"'], fill=fill("F8D7DA"), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"G{r0}:G{rl}", FormulaRule(formula=[f'G{r0}="WAARSCHUWING"'], fill=fill(ORANGE_LT), font=Font(color="9C5700", bold=True)))
    ws.conditional_formatting.add(f"G{r0}:G{rl}", FormulaRule(formula=[f'G{r0}="OK"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"G{r0}:G{rl}", FormulaRule(formula=[f'G{r0}="INFO"'], font=Font(color=GREY, bold=True)))
    # uitzonderingslijsten
    r = rl + 3
    section(ws, r, 2, "Uitzonderingen (eerste 15 per lijst) — niets verdwijnt stil", 7)
    lists = [
        ("Nieuw: in ZPMW03, niet in Mapping", f"STG_ZPMW03!${ZCOLS['NieuwNr']}${STG0}:${ZCOLS['NieuwNr']}${STG1}",
         zc("WBS"), zc("Omschrijving"), "B"),
        ("Vervallen/ontbreekt: in Mapping, niet in ZPMW03", "V_VervallenNr", "V_WBS", "V_Naam", "D"),
        ("Woningetalage, niet in ZPMW03", f"STG_Woningetalage!${ECOLS['NietInZNr']}${STG0}:${ECOLS['NietInZNr']}${STG1}",
         ec("Stufe3"), ec("OmschrStufe3"), "F"),
        ("Verliesprojecten (BOR FY RJ, EUR)", "V_VerliesNr", "V_Naam", "V_BorFY", "H"),
    ]
    for j, (lab, nr, a, b, c0) in enumerate(lists):
        cA = c0
        cB = col(ord(c0) - 64 + 1) if c0 != "H" else "I"
        put(ws, f"{cA}{r+1}", lab, bold=True, color=NAVY, size=8)
        for k in range(1, 16):
            rr = r + 1 + k
            ws[f"{cA}{rr}"] = f'=IFERROR(INDEX({a},MATCH({k},{nr},0)),"")'
            ws[f"{cB}{rr}"] = f'=IFERROR(INDEX({b},MATCH({k},{nr},0)),"")'
            ws[f"{cA}{rr}"].font = font(8)
            ws[f"{cB}{rr}"].font = font(8, False, GREY)
            if c0 == "H":
                ws[f"{cB}{rr}"].number_format = FMT_EUR
        ws[f"{cA}{r+17}"] = f'=IF(COUNT({nr})=0,"(geen)","")'
        ws[f"{cA}{r+17}"].font = font(8, False, GREEN, True)
    widths(ws, {"B": 30, "C": 22, "D": 44, "E": 16, "F": 30, "G": 26, "H": 34, "I": 16})
    ws.column_dimensions["B"].width = 18
    ws.freeze_panes = "A9"
    landscape(ws, f"A1:H{r+18}")


# ---------------------------------------------------------------------------------------------
# Start en documentatiebladen
# ---------------------------------------------------------------------------------------------
def build_start(wb, ws):
    base_sheet(ws, zoom=100)
    title(ws, "RDFO-prototype — fictieve demodata", "Werkend prototype op uitsluitend fictieve data. Demo-aannames zijn geen vastgestelde SAP-definities.", 8)
    widths(ws, {"B": 24, "C": 70, "D": 4, "E": 26, "F": 22})
    put(ws, "B4", "Status", bold=True, color=NAVY, size=11)
    ws["C4"] = '=IF(Controle!$D$4>0,"✖ "&Controle!$D$4&" FOUT(en) — zie Controle",IF(Controle!$D$5>0,"▲ geen fouten, "&Controle!$D$5&" waarschuwing(en)","✔ geen fouten of waarschuwingen"))'
    ws["C4"].font = font(11, True, NAVY)
    ws["C4"].hyperlink = "#'Controle'!A1"
    ws["C5"] = '="Upload: "&P_UploadVersie&" "&P_UploadVJ&" · peildatum "&P_PeilTxt&" · Import-ID "&P_ImportID&" · rapportage: doeljaar "&Sel_Doeljaar&", "&Sel_Label'
    ws["C5"].font = font(9, False, GREY)
    nav = [
        ("Rapportage", None),
        ("Dashboard", "Regiodashboard: filters (doeljaar, versie, type, project), KPI-tegels, prognoseverloop, mutaties, doorkijk."),
        ("RDFO", "Volledige RDFO-opbouw: Budget, Q1–Q4 vastgelegd, upload-concept, realisatie YTD, to go, delta’s."),
        ("Projecten", "Per project: versies, realisatie, to go, BOR, marge, verschillen + toelichting."),
        ("Projectdetail", "Eén project: kerngegevens, versieverloop, planning omzet/BOR, woningverkopen."),
        ("Woningen", "Verkopen per maand/kwartaal, jaarverwachting per versie, koperstype, segment, projectmatrix."),
        ("Controle", "Alle controles met status en uitzonderingslijsten."),
        ("Invoer", None),
        ("Parameters", "Rapportagejaar, peildatum, uploadversie, eenheid, Import-ID, status bij vastleggen."),
        ("Invoer_Aanvullend", "Posten die niet uit de bronnen komen (taakstelling, eliminaties, AK-dekking, fees …)."),
        ("Invoer_Toelichting", "Toelichtingen op verschillen per project en versie."),
        ("Invoer_Koperstype", "Optionele koperstype-verdeling voor toekomstige verkopen."),
        ("Mapping", "Projectmapping WBS → projectdefinitie, type, segment, woningen meetellen."),
        ("Versiebeheer", None),
        ("Vastlegging", "Onveranderlijk versiegrootboek (alleen toevoegen)."),
        ("Vastleggen", "Concept van de huidige upload in grootboekformaat — HANDMATIG kopiëren/plakken als waarden."),
        ("Importlog", "Audittrail per import/vastlegging."),
        ("Bron en verwerking (niet wijzigen)", None),
        ("SRC_ZPMW03", "Ongewijzigde export ZPMW03 — hier plakken."),
        ("SRC_Woningetalage", "Ongewijzigde export Woningetalage — hier plakken."),
        ("STG_ZPMW03", "Staging: velden gezocht op kolomnaam, tekenomzetting, regelcontroles."),
        ("STG_Woningetalage", "Staging: dubbele headers op positie, jaarheaders via mapping, regelcontroles."),
        ("VERW_Project", "Eén regel per gemapt WBS; elke bron eerst per WBS geaggregeerd."),
        ("CALC_Versies", "Totalen per versie (regio en gefilterd), geselecteerde en vorige versie."),
        ("Documentatie", None),
        ("Handleiding", "Bronvervanging, verversen, kwartaalvastlegging, correcties."),
        ("Rapportagemapping", "Elke rapportageregel → bron / berekening / aanvullende invoer + open definities."),
        ("Productievragen", "Openstaande vragen vóór productie."),
        ("Veldwoordenboek", "Interne veldnamen, aliassen, dubbele headers, jaarheader-mapping."),
        ("Testresultaten", "Uitgevoerde scenario-tests (nieuwe upload, dubbele import, correctie, weging, verlies)."),
    ]
    r = 7
    for n, d in nav:
        if d is None:
            r += 1
            put(ws, f"B{r}", n, bold=True, color=NAVY, size=10)
            ws[f"B{r}"].border = Border(bottom=thick_navy)
            ws[f"C{r}"].border = Border(bottom=thick_navy)
        else:
            c = put(ws, f"B{r}", n, color="1F4E9A")
            c.hyperlink = f"#'{n}'!A1"
            c.font = Font(name=FONT, size=9, color="1F4E9A", underline="single")
            put(ws, f"C{r}", d, color=GREY)
        r += 1
    put(ws, "E7", "Legenda", bold=True, color=NAVY, size=10)
    leg = [(YELLOW_LT, "Invoercel (blauw cijfer, lichtgeel)"), (YELLOW_LT, "Geselecteerde versie (geel vlak)"), ("EEF3EA", "Realisatie / actual (groen, cursief)"),
           (ORANGE_LT, "Subtotaal onvolledig (bevat n.b.)"), (LIGHT, "KPI-tegel / totaal"), (MID, "Bronkop")]
    for i, (fc, t) in enumerate(leg):
        cell = ws[f"E{8+i}"]
        cell.fill = fill(fc)
        cell.value = t
        cell.font = font(8)
    put(ws, "E15", "n.b. = niet beschikbaar · n.v.t. = niet van toepassing", color=GREY, size=8)
    put(ws, "E16", "Bedragen bron: EUR, volle precisie", color=GREY, size=8)
    put(ws, "E17", "Omzet/winst positief, verlies negatief", color=GREY, size=8)
    r = max(r, 20) + 1
    section(ws, r, 2, "Ontwerpkeuze", 5)
    for i, (k, v) in enumerate(DOC.ONTWERPKEUZE):
        put(ws, f"B{r+1+i}", k, bold=True)
        put(ws, f"C{r+1+i}", v, wrap=True, color=GREY)
        ws.row_dimensions[r + 1 + i].height = 15 * max(1, len(v) // 95 + 1)
    landscape(ws, f"A1:F{r + len(DOC.ONTWERPKEUZE) + 1}")


def build_docs(ws_hand, ws_map, ws_vr):
    for ws in (ws_hand, ws_map, ws_vr):
        base_sheet(ws, zoom=100)
    title(ws_hand, "Handleiding", "Bronvervanging · verversen · kwartaalvastlegging (HANDMATIG) · correcties · rapporteren", 4)
    widths(ws_hand, {"B": 130})
    r = 4
    for sec, lines in DOC.HANDLEIDING:
        section(ws_hand, r, 2, sec, 1)
        r += 1
        for ln in lines:
            put(ws_hand, f"B{r}", ln, wrap=True, color="000000" if not ln.startswith("Nooit") else RED, bold=ln.startswith("Nooit"))
            ws_hand.row_dimensions[r].height = 14 * max(1, len(ln) // 150 + 1)
            r += 1
        r += 1
    title(ws_map, "Rapportagemapping", "Elke rapportageregel naar bron, berekening en aanvullende invoer, met openstaande definities.", 7)
    widths(ws_map, {"B": 34, "C": 30, "D": 70, "E": 18, "F": 50})
    text_block(ws_map, 4, [("Rapportageregel", "Bron", "Berekening", "Aanvullende invoer", "Openstaande definitie")] + DOC.MAPPING)
    title(ws_vr, "Productievragen", "Te beantwoorden vóór productie. Het prototype draait op expliciete demo-aannames.", 5)
    widths(ws_vr, {"B": 5, "C": 12, "D": 95, "E": 40})
    text_block(ws_vr, 4, [("Nr", "Thema", "Vraag", "Raakt")] + [(i + 1,) + q for i, q in enumerate(DOC.PRODUCTIEVRAGEN)])


def build_tests(ws, results):
    base_sheet(ws, zoom=100)
    title(ws, "Testresultaten", "Automatisch uitgevoerd met tools/test_rdfo.py (LibreOffice-herberekening van gewijzigde kopieën van deze werkmap).", 6)
    widths(ws, {"B": 5, "C": 48, "D": 70, "E": 12})
    if not results:
        put(ws, "B4", "Nog geen testrun opgenomen.", color=GREY)
        return
    text_block(ws, 4, [("Nr", "Test", "Uitkomst / bewijs", "Resultaat")] + [(i + 1,) + tuple(t) for i, t in enumerate(results)])
    for r in range(5, 5 + len(results)):
        c = ws[f"E{r}"]
        c.font = font(9, True, GREEN if c.value == "GESLAAGD" else RED)


# ---------------------------------------------------------------------------------------------
# Hoofdroutine
# ---------------------------------------------------------------------------------------------
def build(out, test_results=None):
    wb = Workbook()
    wb.remove(wb.active)
    sh = {}
    for n, tab in SHEETS:
        sh[n] = wb.create_sheet(n)
        sh[n].sheet_properties.tabColor = tab
    hist = load_workbook(DATA / HIST_FILE)
    build_parameters(wb, sh["Parameters"])
    nz = copy_source(sh["SRC_ZPMW03"], DATA / SRC_Z_FILE, "ZPMW03")
    ne = copy_source(sh["SRC_Woningetalage"], DATA / SRC_E_FILE, "Woningetalage")
    build_stg(wb, sh["STG_ZPMW03"], "SRC_ZPMW03", z_fields(), z_derived(), ZCOLS)
    build_stg(wb, sh["STG_Woningetalage"], "SRC_Woningetalage", e_fields(), e_derived(), ECOLS)
    put(sh["STG_ZPMW03"], "A5", "Staging ZPMW03 — niet invullen", bold=True, color=NAVY)
    put(sh["STG_Woningetalage"], "A5", "", color=GREY)
    build_veldwoordenboek(wb, sh["Veldwoordenboek"])
    mapping = build_mapping(wb, sh["Mapping"], hist)
    zsrc = load_workbook(DATA / SRC_Z_FILE)["ZPMW03"]
    hdr = [c.value for c in zsrc[1]]
    ifrs = {row[hdr.index("WBS-element")]: row[hdr.index("IFRS J/N")] for row in zsrc.iter_rows(min_row=2, values_only=True)}
    rows = hist_rows(hist, mapping, ifrs)
    build_ledger(wb, sh["Vastlegging"], rows)
    build_aanvullend(wb, sh["Invoer_Aanvullend"])
    build_toelichting(wb, sh["Invoer_Toelichting"])
    build_koperstype(wb, sh["Invoer_Koperstype"])
    counts = {v: sum(1 for r in rows if r[5] == v) for v in VERSIES}
    build_importlog(wb, sh["Importlog"], counts, nz, ne)
    build_verw(wb, sh["VERW_Project"])
    build_calc(wb, sh["CALC_Versies"])
    build_vastleggen(wb, sh["Vastleggen"])
    build_rdfo(wb, sh["RDFO"])
    build_projecten(wb, sh["Projecten"])
    build_dashboard(wb, sh["Dashboard"])
    build_woningen(wb, sh["Woningen"])
    first = list(mapping.values())[0][2]
    build_projectdetail(wb, sh["Projectdetail"], first)
    build_controle(wb, sh["Controle"])
    build_start(wb, sh["Start"])
    build_docs(sh["Handleiding"], sh["Rapportagemapping"], sh["Productievragen"])
    build_tests(sh["Testresultaten"], test_results)
    for n in ("STG_ZPMW03", "STG_Woningetalage", "VERW_Project", "CALC_Versies"):
        sh[n].sheet_view.zoomScale = 80
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(out)
    return {"ledger_rows": len(rows), "nz": nz, "ne": ne}


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "RDFO_prototype.xlsx")
    print(build(out))
