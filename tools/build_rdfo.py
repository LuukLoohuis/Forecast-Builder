"""Bouwt het RDFO-prototype (Excel) uit de twee operationele bronnen en de fictieve historie.

Gebruik:  python tools/build_rdfo.py [uitvoerbestand.xlsx]

Lagen in de werkmap (links naar rechts in de tabvolgorde):
  rapportage  -> Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Controle
  invoer      -> Invoer (instellingen, projecten, aanvullende posten, toelichtingen, koperstype, jaarkolommen), Stoplicht
  versies     -> Vastlegging (onveranderlijk grootboek), Vastleggen (concept uit upload), Importlog
  berekening  -> VERW_Project, CALC_Versies
  staging     -> STG_ZPMW03, STG_Woningetalage  (kolommen gezocht op header, niet op positie)
  bron        -> Bron_ZPMW03, Bron_Woningetalage (ongewijzigde export, alleen plakken)
  documentatie-> Handleiding (incl. bijlagen: veldwoordenboek, rapportagemapping, productievragen, tests)
  verborgen   -> Concept, VERW_Project, CALC_Versies, STG_*, Bron_*, Lijsten (macro 'Toon rekenbladen')
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
MAP0, MAP1 = 31, 130            # projectmapping (100 regels) op tabblad Invoer
V0, V1 = 6, 105                  # verwerkingsrijen (1 per mappingregel)
L0, L1 = 6, 5005                 # grootboek vastlegging (5000 regels)
A0, A1 = 136, 335                # aanvullende invoer (Invoer, sectie 3)
T0, T1 = 341, 440                # toelichtingen (Invoer, sectie 4)
K0, K1 = 446, 545                # koperstype (Invoer, sectie 5)
YM0 = 551                        # jaarkolommen Woningetalage (Invoer, sectie 6)
VERSIES = ["Budget", "Q1", "Q2", "Q3", "Q4"]

K_OMZ, K_BOR, K_OMZA, K_BORA = "Omzet FY", "BOR FY", "Omzet actual YTD", "BOR actual YTD"
K_WFY, K_WVJ, K_WACT = "Woningverkopen FY", "Woningen verkocht t/m vorig jaar", "Woningverkopen actual YTD"
K_WQ = [f"Woningverkopen Q{q}" for q in range(1, 5)]

RUBRIEKEN = ["Taakstelling omzet", "Interne eliminaties", "IFRS-eliminaties", "AK-dekking",
             "Fees", "Overige resultaten", "Taakstelling EBIT", "Orderintake budget"]

SHEETS = [  # (naam, tabkleur, zichtbaar)
    ("Start", NAVY, True), ("Dashboard", YELLOW, True), ("RDFO", YELLOW, True), ("Projecten", YELLOW, True),
    ("Projectdetail", YELLOW, True), ("Woningen", YELLOW, True), ("Outlook", YELLOW, True),
    ("Stoplicht", YELLOW, True), ("Controle", RED, True), ("Invoer", "1F4E9A", True), ("Historie", GREEN, True),
    ("Handleiding", MID, True),
    ("Concept", GREY, False), ("VERW_Project", GREY, False), ("CALC_Versies", GREY, False),
    ("STG_ZPMW03", GREY, False), ("STG_Woningetalage", GREY, False),
    ("Bron_ZPMW03", "000000", False), ("Bron_Woningetalage", "000000", False), ("Lijsten", GREY, False),
]


def rng(sheet, c, r0, r1):
    return f"{sheet}!${c}${r0}:${c}${r1}"


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
        dict(n="PEW_BOR", h="PEW resultaat", k="num"),
        dict(n="BOB_VJ", h="Bedrijfsopbrengsten VJ", k="num"),
        dict(n="BOB_LJ", h="Bedrijfsopbrengsten LJ", k="num"),
        dict(n="Kosten_LJ", h="Kosten LJ", k="num"),
        dict(n="BOR_LJ", h="Bruto (omzet) result. LJ", k="num"),
        dict(n="BOR_VJ", h="Bruto (omzet) result. VJ", k="num"),
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
        ("M_OmzVJ", neg("BOB_VJ"), "Omzet realisatie vorig jaar = -Bedrijfsopbrengsten VJ"),
        ("M_BorVJ", neg("BOR_VJ"), "BOR realisatie vorig jaar = -Bruto (omzet) result. VJ"),
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
        ("M_PEWOmz", neg("PEW_BOB"), "PEW omzet (totale projectomzet) = -PEW opbrengst"),
        ("M_PEWBor", neg("PEW_BOR"), "PEW BOR = -PEW resultaat"),
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
        ("InMapping", f'=IF(COUNTIF(Invoer!$A${MAP0}:$A${MAP1},[WBS])>0,1,0)', "1 = WBS staat in Mapping"),
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
        ("C_Aantal", '=IFERROR(IF(ABS([Aantal]-([VJ_T]+SUM([M01]:[M12])+[VolgendJaar]+[Y_RJ2]+[Y_RJ3]+[Y_RJ4]+[Y_RJ5]+[Y_RJ6]))>P_Tol,1,0),1)', "1 = aantal <> vorig + 12 mnd + volgend + latere jaren"),
        ("C_Etalage", '=IFERROR(IF(ABS([Etalage]-([Aantal]-[Cum_T]))>P_Tol,1,0),1)', "1 = etalage <> aantal - cumulatief"),
        ("C_Cum", '=IFERROR(IF(ABS([Cum_T]-([VJ_T]+[HJ_T]))>P_Tol,1,0),1)', "1 = cumulatief <> t/m vorig jr + huidig jr"),
        ("C_HJ", '=IFERROR(IF(ABS([HJ_T]-[ActYTD])>P_Tol,1,0),1)', "1 = Huidig jr Totaal <> som maanden t/m peilmaand"),
        ("C_BP", '=IFERROR(IF(OR(ABS([VJ_B]+[VJ_P]-[VJ_T])>P_Tol,ABS([HJ_B]+[HJ_P]-[HJ_T])>P_Tol,ABS([Cum_B]+[Cum_P]-[Cum_T])>P_Tol),1,0),1)', "1 = Beleggers + Particulier <> Totaal"),
        ("C_Jaren", '=IFERROR(IF(OR(ABS([HuidigJaar]-SUM([M01]:[M12]))>P_Tol,ABS([VolgendJaar]-([VQ1]+[VQ2]+[VQ3]+[VQ4]))>P_Tol),1,0),1)', "1 = Huidig/Volgend jaar <> som maanden/kwartalen"),
        ("C_NietGeinterp", '=IFERROR(IF(ABS([Rond])+ABS([Y_NEG1])+ABS([Y_NEG2])+ABS([Y_NEG3])>0,1,0),1)', "1 = niet-geïnterpreteerde kolom bevat waarde"),
        ("InZPMW03", f'=COUNTIF(STG_ZPMW03!$B${STG0}:$B${STG1},[Stufe3])', "Aantal keer in ZPMW03"),
        ("Dubbel", f'=COUNTIF($B${STG0}:$B${STG1},[Stufe3])', "Aantal keer dat Stufe 3 voorkomt"),
        ("C_Projdef", '=IF([InZPMW03]=0,0,IF(INDEX(STG_ZPMW03!$[Z:Projectdefinitie]$%d:$[Z:Projectdefinitie]$%d,MATCH([Stufe3],STG_ZPMW03!$B$%d:$B$%d,0))=[Project]&"",0,1))' % (STG0, STG1, STG0, STG1),
         "1 = Projectdefinitie (ZPMW03) <> Project (Woningetalage) voor dezelfde WBS"),
        ("C_MapAantal", f'=IFERROR(IF(ABS(INDEX(Invoer!$D${MAP0}:$D${MAP1},MATCH([Stufe3],Invoer!$A${MAP0}:$A${MAP1},0))-[Aantal])>P_Tol,1,0),0)', "1 = aantal woningen wijkt af van Mapping"),
        ("C_Meetellen", f'=IFERROR(IF(AND(INDEX(Invoer!$G${MAP0}:$G${MAP1},MATCH([Stufe3],Invoer!$A${MAP0}:$A${MAP1},0))="Nee",ABS([Aantal])+ABS([Cum_T])+ABS([HuidigJaar])>0),1,0),0)', "1 = woningen bij project met 'Woningen meetellen = Nee'"),
        ("InMapping", f'=IF(COUNTIF(Invoer!$A${MAP0}:$A${MAP1},[Stufe3])>0,1,0)', "1 = Stufe 3 staat in Mapping"),
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
DASH = {}   # verwijzingen naar dashboard-grafiekdata (hergebruikt op Dias)
WON = {}


# ---------------------------------------------------------------------------------------------
# Invoer: één tabblad met secties (instellingen, projecten, aanvullende posten, toelichtingen,
# koperstype, jaarheaders). Keuzelijsten staan op het verborgen tabblad Lijsten.
# ---------------------------------------------------------------------------------------------
INV_SECTIONS = [  # (anker, titel, kopregel)
    ("inv_inst", "1. Instellingen rapportage", 5),
    ("inv_proj", "2. Projecten (mapping)", MAP0 - 1),
    ("inv_aanv", "3. Aanvullende RDFO-posten", A0 - 1),
    ("inv_toel", "4. Toelichtingen op verschillen", T0 - 1),
    ("inv_kt", "5. Koperstype toekomstige verkopen", K0 - 1),
    ("inv_ym", "6. Jaarkolommen Woningetalage (expert)", YM0 - 1),
]


def inv_head(ws, hdr_row, titel, uitleg, labels, helpers=0):
    r = hdr_row - 2
    for c in range(1, 16):
        ws.cell(row=r, column=c).fill = fill(NAVY)
    put(ws, (r, 1), titel, bold=True, color=WHITE, size=12, fillc=NAVY)
    ws.row_dimensions[r].height = 22
    put(ws, (r + 1, 1), uitleg, color=GREY, italic=True)
    header(ws, hdr_row, 1, labels, height=30, color=NAVY2)
    for k in range(helpers):
        ws.cell(row=hdr_row, column=len(labels) - k).fill = fill(GREY)
    put(ws, (r, 14), "▲ naar boven", color=YELLOW, size=8).hyperlink = "#'Invoer'!A1"


def build_lijsten(wb, ws):
    base_sheet(ws)
    put(ws, "A1", "Keuzelijsten (verborgen tabblad; gebruikt door validatie en formules)", bold=True)
    put(ws, "A3", "Versie", bold=True)
    put(ws, "B3", "Peilmaand", bold=True)
    for i, (v, m) in enumerate(zip(VERSIES, [12, 3, 6, 9, 12])):
        put(ws, f"A{4+i}", v)
        put(ws, f"B{4+i}", m)
    add_name(wb, "V_List", "Lijsten!$A$4:$A$8")
    add_name(wb, "V_Maand", "Lijsten!$B$4:$B$8")
    lists = [
        ("L_TypeList", ["Alle", "Opstal", "Deelneming"]),
        ("L_EenheidList", ["EUR", "EUR x 1.000", "EUR mln"]),
        ("L_JaNee", ["Ja", "Nee"]),
        ("L_StatusList", ["Concept", "Vastgesteld"]),
        ("L_AanvStatus", ["Vastgesteld", "Concept", "Fictieve aanname", "Voorbeeld"]),
        ("L_Rubrieken", RUBRIEKEN),
        ("L_Soort", ["FY", "Actual YTD"]),
        ("L_Niveau", ["Regio", "Project"]),
        ("L_Kleur", ["Groen", "Oranje", "Rood"]),
    ]
    for j, (name, vals) in enumerate(lists):
        c = col(4 + j)
        put(ws, f"{c}3", name, bold=True)
        for i, v in enumerate(vals):
            put(ws, f"{c}{4+i}", v)
        add_name(wb, name, f"Lijsten!${c}$4:${c}${3+len(vals)}")
    c = col(4 + len(lists))
    put(ws, f"{c}3", "L_Doeljaar", bold=True)
    ws[f"{c}4"] = "=P_RJ"
    ws[f"{c}5"] = "=P_RJ+1"
    add_name(wb, "L_Doeljaar", f"Lijsten!${c}$4:${c}$5")
    c = col(5 + len(lists))
    put(ws, f"{c}3", "L_Projecten", bold=True)
    put(ws, f"{c}4", "Alle")
    for i in range(MAP1 - MAP0 + 1):
        ws[f"{c}{5+i}"] = f'=IF(Invoer!$C${MAP0+i}="","",Invoer!$C${MAP0+i})'
    add_name(wb, "L_Projecten", f"Lijsten!${c}$4:${c}${5+MAP1-MAP0}")
    add_name(wb, "L_ProjectenDet", f"Lijsten!${c}$5:${c}${5+MAP1-MAP0}")


PARAMS = [
    ("P_Regio", "Regio / rapportage-eenheid", "Demo Regio", None, "Alleen voor titels."),
    ("P_RJ", "Rapportagejaar (lopend jaar in de upload)", 2026, "0", "Bepaalt welke ZPMW03-kolommen gelezen worden ('OP.BOB 2026', 'OP.BOB 2027-Q1', …)."),
    ("P_Peildatum", "Peildatum upload", PEIL, FMT_DATE, "Tot en met deze datum zijn 'Bedrijfsopbrengsten LJ' en 'Huidig jr' realisatie."),
    ("P_PeilMaand", "Peilmaand (berekend)", "=MONTH(P_Peildatum)", "0", "Maanden t/m deze maand = realisatie, daarna = prognose."),
    ("P_PeilTxt", "Peildatum als tekst (berekend)", '=RIGHT("0"&DAY(P_Peildatum),2)&"-"&RIGHT("0"&MONTH(P_Peildatum),2)&"-"&YEAR(P_Peildatum)', None, ""),
    ("P_UploadVJ", "Versiejaar van de upload", 2026, "0", "Bij welk jaar hoort de huidige prognoseronde."),
    ("P_UploadVersie", "Versie van de upload", "Q3", None, "Budget / Q1 / Q2 / Q3 / Q4. Controle vergelijkt dit met de peildatum."),
    ("P_VJ", "Versiejaar in de rapporten", 2026, "0", "De kolommen Budget, Q1 … Q4 tonen vastgelegde versies uit dit jaar."),
    ("P_Eenheid", "Eenheid RDFO en Projecten", "EUR x 1.000", None, "EUR / EUR x 1.000 / EUR mln. De bron blijft in hele euro's bewaard."),
    ("P_Factor", "Deelfactor (berekend)", '=IF(P_Eenheid="EUR mln",1000000,IF(P_Eenheid="EUR x 1.000",1000,1))', "#,##0", ""),
    ("P_Demo", "Fictieve demowaarden meenemen?", "Nee", None, "Ja = regels met status 'Fictieve aanname' (sectie 3) tellen mee. Standaard: Nee."),
    ("P_ImportID", "Import-ID huidige upload", "IMP-2026-Q3-01", None, "Wordt door de knop 'Bronnen inlezen' automatisch gevuld."),
    ("P_BronZ", "Bestandsnaam ZPMW03", SRC_Z_FILE, None, "Automatisch bij inlezen."),
    ("P_BronE", "Bestandsnaam Woningetalage", SRC_E_FILE, None, "Automatisch bij inlezen."),
    ("P_Importdatum", "Importdatum", IMPORTDATUM, FMT_DATE, "Automatisch bij inlezen."),
    ("P_RegelsZ", "Aantal regels ZPMW03 (export)", 20, "0", "Automatisch bij inlezen; Controle vergelijkt."),
    ("P_RegelsE", "Aantal regels Woningetalage (export)", 20, "0", "Automatisch bij inlezen."),
    ("P_StatusVastleg", "Status bij vastleggen", "Vastgesteld", None, "Alleen 'Vastgesteld' telt mee in de versiekolommen."),
    ("P_ToelCorrectie", "Toelichting bij correctie", "", None, "De knop 'Prognose vastleggen' vraagt hierom als de versie al bestaat."),
    ("P_Tol", "Tolerantie controles (EUR / aantal)", 1, "0.00", "Afrondingsmarge."),
]
PARAM_INPUTS = {"P_Regio", "P_RJ", "P_Peildatum", "P_UploadVJ", "P_UploadVersie", "P_VJ", "P_Eenheid", "P_Demo",
                "P_ImportID", "P_BronZ", "P_BronE", "P_Importdatum", "P_RegelsZ", "P_RegelsE", "P_StatusVastleg",
                "P_ToelCorrectie", "P_Tol"}
PROW = {}


def build_invoer(wb, ws, hist):
    base_sheet(ws, zoom=90)
    title(ws, "Invoer", "Alle invoer op één plek. Gele cellen invullen; grijze kolommen rekenen zelf.", 15)
    put(ws, "B3", "Ga naar:", bold=True, color=NAVY)
    for i, (anchor, t, hr) in enumerate(INV_SECTIONS):
        c = put(ws, (3, 3 + 2 * i), t, color="1F4E9A", size=9)
        c.hyperlink = f"#'Invoer'!A{hr - 2}"
        c.font = Font(name=FONT, size=9, color="1F4E9A", underline="single")
    widths(ws, {"A": 17, "B": 15, "C": 30, "D": 15, "E": 14, "F": 20, "G": 12, "H": 13, "I": 9, "J": 15,
                "K": 46, "L": 14, "M": 9, "N": 9, "O": 9})

    # 1. instellingen
    inv_head(ws, 7, "1. Instellingen rapportage", "Rapportagejaar, peildatum en versie bepalen alle kolomkoppen en jaartallen.",
             ["", "", "Instelling", "Waarde", "Toelichting"])
    for i, (name, lab, val, fmt, toel) in enumerate(PARAMS):
        r = 8 + i
        PROW[name] = r
        put(ws, f"C{r}", lab)
        c = put(ws, f"D{r}", val, fmt=fmt, align="right")
        if name in PARAM_INPUTS:
            input_style(c, fmt)
        else:
            helper_style(c)
        put(ws, f"E{r}", toel, color=GREY, italic=True, size=8)
        add_name(wb, name, f"Invoer!$D${r}")
    dv_list(ws, f"D{PROW['P_UploadVersie']}", "V_List")
    dv_list(ws, f"D{PROW['P_Eenheid']}", "L_EenheidList")
    dv_list(ws, f"D{PROW['P_Demo']}", "L_JaNee")
    dv_list(ws, f"D{PROW['P_StatusVastleg']}", "L_StatusList")
    assert 8 + len(PARAMS) <= MAP0 - 3

    # 2. projecten
    inv_head(ws, MAP0 - 1, "2. Projecten (mapping)", "Eén regel per WBS-element. Nieuw project? Onderaan toevoegen (niet tussenvoegen of sorteren: "
             "de stoplichtinvoer hoort bij de regel). Koppelsleutel = WBS-element = Stufe 3.",
             ["WBS-element", "Projectdefinitie", "Projectnaam", "Woningen (aandeelbasis)", "Aandeel % (info)", "Type",
              "Woningen meetellen", "Segment (demo)"])
    rows = list(hist["Projectmapping"].iter_rows(min_row=2, values_only=True))
    for r in range(MAP0, MAP1 + 1):
        row = rows[r - MAP0] if r - MAP0 < len(rows) else [None] * 8
        for j, v in enumerate(row):
            input_style(ws.cell(row=r, column=1 + j, value=v))
    dv_list(ws, f"F{MAP0}:F{MAP1}", '"Opstal,Deelneming"')
    dv_list(ws, f"G{MAP0}:G{MAP1}", '"Ja,Nee"')
    for c, n in zip("ABCDEFGH", ["WBS", "Projdef", "Naam", "Aantal", "Aandeel", "Type", "Meetellen", "Segment"]):
        add_name(wb, f"Map_{n}", f"Invoer!${c}${MAP0}:${c}${MAP1}")

    # 3. aanvullende posten
    inv_head(ws, A0 - 1, "3. Aanvullende RDFO-posten", "Posten die niet in ZPMW03 staan. Leeg = 'n.b.' in de RDFO (nooit stil nul). "
             "Bedragen in hele euro's; + verhoogt omzet/resultaat. Status: Vastgesteld telt in versiekolommen, Concept alleen in 'Upload', "
             "Fictieve aanname alleen als instelling 'demowaarden' = Ja, Voorbeeld nooit.",
             ["Doeljaar", "Versiejaar", "Versie", "Niveau", "WBS-element", "Rubriek", "Soort", "Waarde (EUR)", "Eenheid",
              "Status", "Toelichting", "Bron / eigenaar", "telt vast", "telt concept", "conflict"], helpers=3)
    demo = [
        (2026, 2026, "Q3", "Regio", None, "Taakstelling omzet", "FY", 2000000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Interne eliminaties", "FY", -350000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "AK-dekking", "FY", -1400000, "EUR", "Fictieve aanname", "DEMO-scenario: vast bedrag, GEEN generiek AK-percentage", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Fees", "FY", 300000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Regio", None, "Taakstelling EBIT", "FY", 150000, "EUR", "Fictieve aanname", "DEMO-scenario, zelfgekozen waarde — geen bron", "Demo"),
        (2026, 2026, "Q3", "Project", "G.900004.8.9090", "IFRS-eliminaties", "FY", -100000, "EUR", "Voorbeeld", "VOORBEELDREGEL (formaat) — telt nooit mee", "Voorbeeld"),
    ]
    for r in range(A0, A1 + 1):
        row = demo[r - A0] if r - A0 < len(demo) else [None] * 12
        for j, v in enumerate(row):
            input_style(ws.cell(row=r, column=1 + j, value=v), FMT_EUR if j == 7 else None)
        ws[f"M{r}"] = f'=IF(H{r}="","",IF(OR(J{r}="Vastgesteld",AND(J{r}="Fictieve aanname",P_Demo="Ja")),1,0))'
        ws[f"N{r}"] = f'=IF(H{r}="","",IF(OR(J{r}="Vastgesteld",J{r}="Concept",AND(J{r}="Fictieve aanname",P_Demo="Ja")),1,0))'
        ws[f"O{r}"] = (f'=IF(H{r}="","",IF(COUNTIFS($A${A0}:$A${A1},A{r},$B${A0}:$B${A1},B{r},$C${A0}:$C${A1},C{r},'
                       f'$F${A0}:$F${A1},F{r},$G${A0}:$G${A1},G{r},$D${A0}:$D${A1},"<>"&D{r},$H${A0}:$H${A1},"<>")>0,1,0))')
        for c in "MNO":
            helper_style(ws[f"{c}{r}"])
    for c, n in zip("ABCDEFGHIJKMNO", ["Doel", "VJ", "Versie", "Niveau", "WBS", "Rubriek", "Soort", "Waarde", "Eenheid",
                                        "Status", "Toel", "TeltVast", "TeltConcept", "Conflict"]):
        add_name(wb, f"A_{n}", f"Invoer!${c}${A0}:${c}${A1}")
    dv_list(ws, f"C{A0}:C{A1}", "V_List")
    dv_list(ws, f"D{A0}:D{A1}", "L_Niveau")
    dv_list(ws, f"F{A0}:F{A1}", "L_Rubrieken")
    dv_list(ws, f"G{A0}:G{A1}", "L_Soort")
    dv_list(ws, f"J{A0}:J{A1}", "L_AanvStatus")

    # 4. toelichtingen
    inv_head(ws, T0 - 1, "4. Toelichtingen op verschillen", "Eén tekst per project, doeljaar en versie. Verschijnt in Projecten, Projectdetail en PowerPoint.",
             ["Doeljaar", "Versiejaar", "Versie", "WBS-element", "Kengetal", "", "", "", "sleutel", "", "Toelichting", "Auteur", "Datum"], helpers=0)
    ws.cell(row=T0 - 1, column=9).fill = fill(GREY)
    for r in range(T0, T1 + 1):
        vals = [None] * 8
        for c, v in zip("ABCDEKLM", vals):
            input_style(ws.cell(row=r, column=ord(c) - 64, value=v), FMT_DATE if c == "M" else None)
        ws[f"I{r}"] = f'=IF(D{r}="","",D{r}&"|"&A{r}&"|"&B{r}&"|"&C{r})'
        helper_style(ws[f"I{r}"])
    add_name(wb, "T_Key", f"Invoer!$I${T0}:$I${T1}")
    add_name(wb, "T_Txt", f"Invoer!$K${T0}:$K${T1}")
    dv_list(ws, f"C{T0}:C{T1}", "V_List")

    # 5. koperstype
    inv_head(ws, K0 - 1, "5. Koperstype toekomstige verkopen", "Woningetalage kent Beleggers/Particulier alleen voor gerealiseerde verkopen. "
             "Zonder invoer hier blijft de prognose 'Onbekend'.",
             ["WBS-element", "Doeljaar", "Periode", "Koperstype", "Aantal", "Status", "", "", "telt", "type (mapping)", "Toelichting"], helpers=0)
    for c in (9, 10):
        ws.cell(row=K0 - 1, column=c).fill = fill(GREY)
    ex = ("G.900005.8.9090", 2026, "Q4", "Belegger", 4, "Voorbeeld")
    for r in range(K0, K1 + 1):
        vals = ex if r == K0 else [None] * 6
        for j, v in enumerate(vals):
            input_style(ws.cell(row=r, column=1 + j, value=v))
        input_style(ws.cell(row=r, column=11, value="VOORBEELDREGEL — telt nooit mee" if r == K0 else None))
        ws[f"I{r}"] = f'=IF(E{r}="","",IF(OR(F{r}="Vastgesteld",F{r}="Concept"),1,0))'
        ws[f"J{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(Map_Type,MATCH(A{r},Map_WBS,0)),"n.b."))'
        helper_style(ws[f"I{r}"])
        helper_style(ws[f"J{r}"])
    for c, n in zip("ABCDEFIJ", ["WBS", "Doel", "Periode", "Type", "Aantal", "Status", "Telt", "TypeMap"]):
        add_name(wb, f"KT_{n}", f"Invoer!${c}${K0}:${c}${K1}")
    dv_list(ws, f"C{K0}:C{K1}", '"Q1,Q2,Q3,Q4"')
    dv_list(ws, f"D{K0}:D{K1}", '"Belegger,Particulier"')
    dv_list(ws, f"F{K0}:F{K1}", '"Vastgesteld,Concept,Voorbeeld"')

    # 6. jaarheaders
    inv_head(ws, YM0 - 1, "6. Jaarkolommen Woningetalage (expert)", "Hoe de jaarkoppen in de Woningetalage geïnterpreteerd worden. Alleen wijzigen als de export verandert.",
             ["Intern veld", "Bronkop (waarde)", "Gevonden?", "Gebruikt?", "Status", "Interpretatie"])
    ym = [
        ("Y_RJ2", "=P_RJ+2", "Ja", "Bevestigd (demo)", "Verkopen in jaar RJ+2 (prototype: 2028)."),
        ("Y_RJ3", "=P_RJ+3", "Ja", "Aanname (demo)", "Verkopen in jaar RJ+3 (2029). Positieve jaarkop = verkopen in dat jaar."),
        ("Y_RJ4", "=P_RJ+4", "Ja", "Aanname (demo)", "Verkopen in jaar RJ+4 (2030)."),
        ("Y_RJ5", "=P_RJ+5", "Ja", "Aanname (demo)", "Verkopen in jaar RJ+5 (2031)."),
        ("Y_RJ6", "=P_RJ+6", "Ja", "Aanname (demo)", "Verkopen in jaar RJ+6 (2032)."),
        ("Y_NEG1", "=-(P_RJ+7)", "Nee", "Niet geïnterpreteerd", "Negatieve jaarkop: betekenis onbekend; waarde ≠ 0 geeft waarschuwing."),
        ("Y_NEG2", "=-(P_RJ+8)", "Nee", "Niet geïnterpreteerd", "Idem."),
        ("Y_NEG3", "=-(P_RJ+9)", "Nee", "Niet geïnterpreteerd", "Idem."),
        ("Rond", "Rond", "Nee", "Niet geïnterpreteerd", "Betekenis onbekend; waarde ≠ 0 geeft waarschuwing."),
    ]
    for i, (n, h, used, st, interp) in enumerate(ym):
        r = YM0 + i
        put(ws, f"A{r}", n)
        input_style(put(ws, f"B{r}", h, align="left"))
        ws[f"C{r}"] = f'=IF(ISNUMBER(MATCH(B{r},Bron_Woningetalage!$1:$1,0)),"Ja","NEE")'
        put(ws, f"D{r}", used)
        put(ws, f"E{r}", st)
        put(ws, f"F{r}", interp, color=GREY, size=8)
    add_name(wb, "YM_Intern", f"Invoer!$A${YM0}:$A${YM0+len(ym)-1}")
    add_name(wb, "YM_Header", f"Invoer!$B${YM0}:$B${YM0+len(ym)-1}")
    add_name(wb, "YM_Gevonden", f"Invoer!$C${YM0}:$C${YM0+len(ym)-1}")
    ws.conditional_formatting.add(f"C{YM0}:C{YM0+8}", FormulaRule(formula=[f'C{YM0}="NEE"'], font=Font(color=RED, bold=True)))
    ws.freeze_panes = "A5"
    return {row[0]: row for row in rows}


# ---------------------------------------------------------------------------------------------
# Historie: onveranderlijk versiegrootboek + importlog
# ---------------------------------------------------------------------------------------------
LCOLS = ["ImportID", "Bronbestand", "Importdatum", "Status", "Versiejaar", "Versie", "Revisie", "Peildatum",
         "Doeljaar", "WBS", "Projectdefinitie", "Projectnaam", "Type", "IFRS", "Kengetal", "Waarde", "Eenheid",
         "Toelichting"]
LHELP = ["Sleutel", "MaxRev", "Actief", "Dubbel"]
IL_COL = 24  # kolom X: importlog


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


def build_historie(wb, ws, rows, hist_counts, nz, ne):
    base_sheet(ws, zoom=85)
    title(ws, "Historie — vastgelegde prognoses (alleen toevoegen, nooit wijzigen)",
          "De knop 'Prognose vastleggen' voegt hier regels toe. Een correctie is een nieuwe regel met een hogere revisie; de oude regel blijft staan.", 34)
    put(ws, "B3", "Eerste vrije rij:", bold=True)
    ws["D3"] = '=SUMPRODUCT(MAX((L_WBS<>"")*ROW(L_WBS)))+1'
    ws["D3"].font = font(10, True, NAVY)
    add_name(wb, "H_VrijeRij", "Historie!$D$3")
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
        add_name(wb, f"L_{n}", f"Historie!${col(1+j)}${L0}:${col(1+j)}${L1}")
    widths(ws, {"A": 16, "B": 22, "C": 10, "D": 11, "E": 8, "F": 7, "G": 6, "H": 10, "I": 7, "J": 16, "K": 10,
                "L": 24, "M": 10, "N": 5, "O": 24, "P": 13, "Q": 7, "R": 40, "S": 30, "T": 6, "U": 6, "V": 6, "W": 3})
    ws.auto_filter.ref = f"A5:V{L0 + len(rows) - 1}"
    # importlog
    c0 = IL_COL
    put(ws, (4, c0), "Importlog", bold=True, color=NAVY, size=11)
    hdr = ["Import-ID", "Bron", "Bestandsnaam", "Importdatum", "Peildatum", "Versiejaar", "Versie", "Regels bron",
           "Regels vastgelegd (live)", "Status", "Toelichting"]
    header(ws, 5, c0, hdr, height=30, color=GREEN)
    logrows = []
    for v, peil in zip(VERSIES[:4], [dt.datetime(2025, 12, 31), dt.datetime(2026, 3, 31), dt.datetime(2026, 6, 30), PEIL]):
        logrows.append((f"HIST-2026-{v}", "Historie (fictief)", HIST_FILE, IMPORTDATUM, peil, 2026, v, hist_counts[v],
                        "Vastgesteld", "Synthetische snapshot uit testhistorie; geen echte historie."))
    logrows.append(("IMP-2026-Q3-01", "ZPMW03", SRC_Z_FILE, IMPORTDATUM, PEIL, 2026, "Q3", nz, "Ingelezen",
                    "Q3 is al vastgesteld via historie; de upload dient als aansluitcontrole."))
    logrows.append(("IMP-2026-Q3-01", "Woningetalage", SRC_E_FILE, IMPORTDATUM, PEIL, 2026, "Q3", ne, "Ingelezen", ""))
    for i in range(300):
        r = 6 + i
        row = logrows[i] if i < len(logrows) else None
        if row:
            for j, v in enumerate(row[:8]):
                c = ws.cell(row=r, column=c0 + j, value=v)
                c.font = font(8)
                if j in (3, 4):
                    c.number_format = FMT_DATE
            ws.cell(row=r, column=c0 + 9, value=row[8]).font = font(8)
            ws.cell(row=r, column=c0 + 10, value=row[9]).font = font(8)
        ws.cell(row=r, column=c0 + 8, value=f'=IF({col(c0)}{r}="","",COUNTIF(L_ImportID,{col(c0)}{r}))').font = font(8, False, GREY)
    for j, w in enumerate([16, 14, 34, 10, 10, 8, 6, 8, 9, 11, 50]):
        ws.column_dimensions[col(c0 + j)].width = w
    ws.freeze_panes = f"A{L0}"


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
    m = lambda n, mc, lab: c.append((n, lab, "=IF(Invoer!$%s{mr}=\"\",\"\",Invoer!$%s{mr})" % (mc, mc), None))  # noqa: E731
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
             ("OmzVJ", "M_OmzVJ", "Omzet realisatie RJ-1"), ("BorVJ", "M_BorVJ", "BOR realisatie RJ-1"),
             ("OmzRest", "M_OmzRest", "Omzet restant RJ"), ("BorRest", "M_BorRest", "BOR restant RJ"),
             ("OmzFY", "M_OmzFY", "Omzet FY RJ"), ("BorFY", "M_BorFY", "BOR FY RJ"),
             ("OmzFY1", "M_OmzFY1", "Omzet FY RJ+1"), ("BorFY1", "M_BorFY1", "BOR FY RJ+1"),
             ("OmzQ4", "M_OmzQ4", "Omzet Q4 RJ"), ("BorQ4", "M_BorQ4", "BOR Q4 RJ")]
    money += [(f"OmzRJ1_Q{q}", f"M_OmzRJ1_Q{q}", f"Omzet RJ+1 Q{q}") for q in range(1, 5)]
    money += [(f"BorRJ1_Q{q}", f"M_BorRJ1_Q{q}", f"BOR RJ+1 Q{q}") for q in range(1, 5)]
    money += [(f"OmzRJ{o}", f"M_OmzRJ{o}", f"Omzet RJ+{o}") for o in range(2, 9)] + [("OmzONV", "M_OmzONV", "Omzet onverdeeld")]
    money += [(f"BorRJ{o}", f"M_BorRJ{o}", f"BOR RJ+{o}") for o in range(2, 9)] + [("BorONV", "M_BorONV", "BOR onverdeeld")]
    money += [("PEWOmz", "M_PEWOmz", "PEW omzet"), ("PEWBor", "M_PEWBor", "PEW BOR")]
    for n, f, lab in money:
        c.append((n, lab, f'=IF($A{{r}}="","",IF($J{{r}}=0,"n.b.",SUMIFS({zr(f)},{zkey},$A{{r}})))', FMT_EUR))
    won = ["Aantal", "VJ_B", "VJ_P", "VJ_T", "HJ_B", "HJ_P", "HJ_T", "Cum_B", "Cum_P", "Cum_T", "Etalage"] + \
          [f"M{i:02d}" for i in range(1, 13)] + ["HQ1", "HQ2", "HQ3", "HQ4", "ActYTD", "PrognoseRest", "HuidigJaar",
                                                  "VolgendJaar", "VQ1", "VQ2", "VQ3", "VQ4", "Y_RJ2", "Y_RJ3", "Y_RJ4", "Y_RJ5", "Y_RJ6"]
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
    c.append(("OI_Datum", "Orderintake datum", '=IF($A{r}="","",IF(ISNUMBER([I_Orderintake]),IF([I_Orderintake]>0,[I_Orderintake],""),""))', FMT_DATE))
    c.append(("OI_Jaar", "Orderintake jaar", '=IF([OI_Datum]="","",YEAR([OI_Datum]))', "0"))
    c.append(("OI_Kw", "Orderintake kwartaal", '=IF([OI_Datum]="","",ROUNDUP(MONTH([OI_Datum])/3,0))', "0"))
    c.append(("OI_Key", "Sorteersleutel orderintake", '=IF([OI_Datum]="","",[OI_Datum]+ROW()/100000)', "0.00000"))
    c.append(("VerkochtPct", "Verkocht % (cumulatief / totaal)", '=IF($A{r}="","",IF(OR($G{r}="Nee",NOT(ISNUMBER([W_Aantal]))),"n.v.t.",IF([W_Aantal]>0,[W_Cum_T]/[W_Aantal],"n.v.t.")))', FMT_PCT))
    c.append(("BorPct", "BOR % FY RJ", '=IF($A{r}="","",IF(AND(ISNUMBER([BorFY]),ISNUMBER([OmzFY])),IF([OmzFY]<>0,[BorFY]/[OmzFY],"n.b."),"n.b."))', FMT_PCT))
    return c


def build_verw(wb, ws):
    import re
    base_sheet(ws, zoom=80)
    put(ws, "A1", "VERW_Project — berekeningslaag (niet invullen). Eén regel per mappingregel; bronnen per WBS geaggregeerd; bedragen in EUR met managementteken (omzet/winst positief, verlies negatief).", bold=True, color=NAVY)
    put(ws, "A2", "RJ = Invoer › Instellingen › rapportagejaar. 'n.b.' = project ontbreekt in de betreffende bron (zie Controle).", color=GREY, italic=True)
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
    ws["B4"] = ('=IF(N3>0,"LET OP: "&P_UploadVersie&" "&P_UploadVJ&" is al vastgesteld. Opnieuw plakken = CORRECTIE (revisie +1, toelichting verplicht). '
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
    add_name(wb, "C_Verschil", f"Concept!$U${C0}:$U${C1}")
    add_name(wb, "C_Kengetal", f"Concept!$O${C0}:$O${C1}")
    add_name(wb, "C_Doeljaar", f"Concept!$I${C0}:$I${C1}")
    add_name(wb, "C_Vastgelegd", f"Concept!$T${C0}:$T${C1}")
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
        "P": "Δ vs budget",
        "Q": "",
        "S": "Definitie / bron",
    }
    for i, v in enumerate(VERSIES[1:]):
        c = "DEFG"[i]
        labels[c] = f'="{v} "&P_VJ&CHAR(10)&"FY-prognose"&IF(CALC_Versies!$F${CR[v]}>0," ●","")'
    for c in "BCDEFGHIJKLMNOPQS":
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
        ("taak_omz", "Taakstelling omzet", "base", F("AV_Taakstelling omzet"), F("AA_Taakstelling omzet"), "Invoer › 3. Aanvullende RDFO-posten"),
        ("int_elim", "Interne eliminaties", "base", F("AV_Interne eliminaties"), F("AA_Interne eliminaties"), "Invoer › 3. Aanvullende RDFO-posten"),
        ("omzet_old", "Omzet old", "sub", ["omz_ops", "omz_dln", "taak_omz", "int_elim"], None, "Subtotaal: opstal + deelnemingen + taakstelling + interne eliminaties"),
        ("ifrs_elim", "IFRS-eliminaties", "base", F("AV_IFRS-eliminaties"), F("AA_IFRS-eliminaties"), "Invoer › 3. IFRS J/N wordt NIET automatisch geëlimineerd (zie memo)"),
        ("omzet_ifrs", "Omzet IFRS", "sub", ["omzet_old", "ifrs_elim"], None, "Omzet old + IFRS-eliminaties"),
        (None,),
        ("bor_ops", "Projectresultaat (BOR) opstal", "base", F("BorOps"), F("ActBorOps"), "ZPMW03: FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ. Verlies blijft negatief (geen ABS)"),
        ("bor_dln", "Projectresultaat (BOR) deelnemingen", "base", F("BorDln"), F("ActBorDln"), "Idem, Type=Deelneming"),
        ("bor_tot", "Projectresultaten / BOR", "sub", ["bor_ops", "bor_dln"], None, "Subtotaal BOR"),
        ("ak", "AK-dekking", "base", F("AV_AK-dekking"), F("AA_AK-dekking"), "Invoer › 3. Geen generiek AK-percentage toegepast"),
        ("fees", "Fees", "base", F("AV_Fees"), F("AA_Fees"), "Invoer › 3. Aanvullende RDFO-posten"),
        ("overig", "Overige resultaten", "base", F("AV_Overige resultaten"), F("AA_Overige resultaten"), "Invoer › 3. Aanvullende RDFO-posten"),
        ("taak_ebit", "Taakstelling EBIT", "base", F("AV_Taakstelling EBIT"), F("AA_Taakstelling EBIT"), "Invoer › 3. Aanvullende RDFO-posten"),
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
        put(ws, f"S{r}", desc, color=GREY, size=8)
        if kind in ("base", "sub"):
            ws[f"P{r}"] = f'=IF(AND(ISNUMBER(M{r}),M{r}<0),REPT("█",MAX(1,ROUND(-M{r}/$Q$5,0))),"")'
            ws[f"Q{r}"] = f'=IF(AND(ISNUMBER(M{r}),M{r}>0),REPT("█",MAX(1,ROUND(M{r}/$Q$5,0))),"")'
            ws[f"P{r}"].font = font(9, False, BAD)
            ws[f"Q{r}"].font = font(9, False, GOOD)
            ws[f"P{r}"].alignment = Alignment(horizontal="right", vertical="center")
            ws[f"Q{r}"].alignment = Alignment(horizontal="left", vertical="center")
            ws[f"P{r}"].border = Border(right=Side(style="thin", color="333F55"))
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
        "• 'Upload … concept' = live berekend uit Bron_ZPMW03/Bron_Woningetalage; wordt pas een versie na handmatige vastlegging (tab Vastleggen).",
        "• 'Realisatie YTD' (groen, cursief) is géén prognose: actual t/m de peildatum van de geselecteerde versie. To go = geselecteerde FY − realisatie van dezelfde peildatum.",
        "• n.b. = niet beschikbaar (geen vastgelegde of aanvullende invoer) — nooit stil als nul. n.v.t. = doeljaar ligt na het versiejaar, dus nog geen realisatie.",
        "• Oranje subtotaal = onvolledig: minstens één onderliggende post is n.b. De som bevat alleen de beschikbare posten.",
        "• Tekens: omzet en resultaat positief, verlies negatief (rood). Aanvullende invoer in managementteken (+ verhoogt omzet/resultaat).",
    ]
    for i, t in enumerate(notes):
        put(ws, f"B{r+i}", t, bold=i == 0, color=NAVY if i == 0 else GREY, size=8 if i else 9)
    ws["Q5"] = f'=MAX(0.000001,MAX(ABS(MIN(M7:M{last})),ABS(MAX(M7:M{last})))/9)'
    widths(ws, {"B": 36, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 1.5, "J": 12, "K": 12,
                "L": 11, "M": 11, "N": 11, "O": 9, "P": 11, "Q": 11, "R": 2, "S": 70})
    ws.column_dimensions["O"].hidden = True
    ws.row_dimensions[5].hidden = True
    ws.freeze_panes = "C7"
    ws.merge_cells("P6:Q6")
    landscape(ws, f"A1:Q{r+len(notes)}")


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
           "Verlies", "Toelichting (Invoer › 4)", "Order-\nintake", "Ver-\nkoop", "Vergun-\nning", "Net-\ncongestie"]
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
        ws[f"AB{r}"] = f'=IF(OR({A}="",Sel_Prev=""),"n.b.",IF(ISNUMBER({prev}),{prev}/P_Factor,{prev}))'
        ws[f"Q{r}"] = f'=IF({A}="","",IF(AND(ISNUMBER(K{r}),ISNUMBER(AB{r})),K{r}-AB{r},"n.b."))'
        ws[f"R{r}"] = f'=IF({A}="","",IF(ISNUMBER(N{r}),IF(N{r}<0,"Verlies",""),""))'
        ws[f"S{r}"] = f'=IF({A}="","",IFERROR(INDEX(T_Txt,MATCH({A}&"|"&Sel_Doeljaar&"|"&P_VJ&"|"&Sel_Versie,T_Key,0)),""))'
        match = f'AND({A}<>"",OR(Sel_Type="Alle",$C{r}=Sel_Type),OR(Sel_Project="Alle",$B{r}=Sel_Project))'
        ws[f"Z{r}"] = f'=IF(AND({match},ISNUMBER(K{r})),K{r}+ROW()/1000000000,"")'
        ws[f"AA{r}"] = f'=IF(AND({match},ISNUMBER(Q{r})),ABS(Q{r})+ROW()/1000000000,"")'
        ws[f"T{r}"] = f'=IF({A}="","",VERW_Project!${VCOLS["OI_Datum"]}{vr})'
        for c_, sc in zip("UVW", "IJK"):
            ws[f"{c_}{r}"] = f'=IF({A}="","",Stoplicht!${sc}{SL0 + (r - P0)})'
        for c in "ABCDEFGHIJKLMNOPQRSTUVW":
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
        for c in ("Z", "AA", "AB"):
            helper_style(ws[f"{c}{r}"])
        ws[f"T{r}"].number_format = FMT_DATE
        for c_ in "UVW":
            ws[f"{c_}{r}"].number_format = SL_FMT
            ws[f"{c_}{r}"].alignment = Alignment(horizontal="center")
            ws[f"{c_}{r}"].font = font(13)
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
    stoplicht_cf(ws, f"U{P0}:W{P1}")
    ws.auto_filter.ref = f"A9:W{P1}"
    widths(ws, {"A": 15, "B": 27, "C": 10, "D": 5, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 9, "K": 10, "L": 10,
                "M": 9, "N": 9, "O": 7, "P": 9, "Q": 9, "R": 7, "S": 34, "T": 10, "U": 6, "V": 7, "W": 9})
    ws.row_dimensions[8].hidden = True
    for c in ("Z", "AA", "AB"):
        ws.column_dimensions[c].hidden = True
    ws.freeze_panes = f"C{P0}"
    landscape(ws, f"A1:W{P0 + 30}")


# ---------------------------------------------------------------------------------------------
# Grafiek- en tegelhulpjes
# ---------------------------------------------------------------------------------------------
from openpyxl.chart.shapes import GraphicalProperties  # noqa: E402
from openpyxl.drawing.line import LineProperties  # noqa: E402


from openpyxl.chart import LineChart  # noqa: E402
from openpyxl.chart.marker import Marker  # noqa: E402
from openpyxl.chart.text import RichText  # noqa: E402
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties  # noqa: E402
from openpyxl.drawing.text import Font as DFont  # noqa: E402

NAVY_LT = "8A9BB8"   # vorige/concept-tint van donkerblauw
AXIS = "6B7385"


def txt(sz=800, color=AXIS, bold=False):
    cp = CharacterProperties(sz=sz, b=bold, solidFill=color, latin=DFont(typeface="Arial"))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _axes(ch, numfmt, secondary=False):
    ch.y_axis.numFmt = numfmt
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.tickLblPos = "low"
    ch.y_axis.txPr = txt()
    ch.x_axis.txPr = txt(800, "333F55")
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="BFC6D4", w=9525))
    ch.x_axis.majorTickMark = "none"
    ch.y_axis.majorTickMark = "none"
    if secondary:
        ch.y_axis.majorGridlines = None
    else:
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E6E9EF", w=6350, prstDash="dash"))


def style_chart(ch, colors, ytitle=None, numfmt="#,##0.0", w=16.5, h=7.2, labels=False, legend=True, ymin=None):
    """Rustige huisstijl: geen kader, lichte stippel-rasterlijnen, Arial 8, legenda boven."""
    ch.width, ch.height = w, h
    ch.style = 2
    ch.title = None
    ch.roundedCorners = False
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    for s_, c in zip(ch.series, colors):
        s_.graphicalProperties.solidFill = c
        s_.graphicalProperties.line.noFill = True
    _axes(ch, numfmt)
    if ymin is not None:
        ch.y_axis.scaling.min = ymin
    if ytitle:
        ch.y_axis.title = ytitle
    if legend:
        ch.legend.position = "t"
        ch.legend.txPr = txt(800, "333F55")
    else:
        ch.legend = None
    if labels:
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showVal = True
        for k in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
            setattr(ch.dataLabels, k, False)
        ch.dataLabels.numFmt = numfmt + ";-" + numfmt + ";"
        stacked = getattr(ch, "grouping", "") == "stacked"
        ch.dataLabels.txPr = txt(750, "FFFFFF" if stacked else "333F55", stacked)
        ch.dataLabels.position = "ctr" if stacked else "outEnd"
    ch.gapWidth = 55
    return ch


def bar(ws, data_ref, cats_ref, anchor, colors, kind="col", grouping="clustered", line=None, point_colors=None, **kw):
    """Staafgrafiek; optioneel met lijn (line=(ref, kleur, numfmt)) op de secundaire as."""
    ch = BarChart()
    ch.type = kind
    ch.grouping = grouping
    if grouping == "stacked":
        ch.overlap = 100
    else:
        ch.overlap = -10
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    style_chart(ch, colors, **kw)
    if point_colors:
        from openpyxl.chart.marker import DataPoint
        for idx, c in point_colors.items():
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = c
            pt.graphicalProperties.line.noFill = True
            ch.series[0].dPt.append(pt)
    if line:
        ref, color, fmt = line
        ln = LineChart()
        ln.add_data(ref, titles_from_data=True)
        ln.set_categories(cats_ref)
        sr = ln.series[0]
        sr.graphicalProperties.line.solidFill = color
        sr.graphicalProperties.line.width = 28575
        sr.marker = Marker(symbol="circle", size=7)
        sr.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill="FFFFFF", w=9525))
        sr.smooth = False
        ln.y_axis.axId = 200
        ln.y_axis.crosses = "max"
        ln.y_axis.scaling.min = 0
        _axes(ln, fmt, secondary=True)
        ln.dataLabels = DataLabelList()
        ln.dataLabels.showVal = True
        for k in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
            setattr(ln.dataLabels, k, False)
        ln.dataLabels.numFmt = fmt + ";-" + fmt + ";"
        ln.dataLabels.position = "t"
        ln.dataLabels.txPr = txt(750, "8A6A00", True)
        ch += ln
    ws.add_chart(ch, anchor)
    return ch


# ---- IBCS / Zebra-stijl: realisatie (AC) massief, prognose (FC) gearceerd, budget (PL) grijs, afwijking groen/rood
from openpyxl.drawing.fill import ColorChoice, PatternFillProperties  # noqa: E402
from openpyxl.chart.label import DataLabelList as _DLL  # noqa: E402

GOOD = "5AA02C"   # positieve afwijking
BAD = "D7372F"    # negatieve afwijking
PLGREY = "B8BEC9"  # budget / plan


def hatched(color=NAVY):
    return GraphicalProperties(pattFill=PatternFillProperties(prst="wdUpDiag", fgClr=ColorChoice(srgbClr=color),
                                                              bgClr=ColorChoice(srgbClr="FFFFFF")),
                               ln=LineProperties(solidFill=color, w=9525))


def solid(color):
    return GraphicalProperties(solidFill=color, ln=LineProperties(noFill=True))


def series_labels(sr, fmt, color="333F55", pos=None, bold=False):
    sr.dLbls = _DLL()
    sr.dLbls.showVal = True
    for k in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(sr.dLbls, k, False)
    sr.dLbls.numFmt = fmt
    sr.dLbls.txPr = txt(750, color, bold)
    if pos:
        sr.dLbls.position = pos


def ibcs_columns(ws, anchor, cats, ac, fc, pl=None, total=None, w=17, h=8, numfmt="#,##0.0", kind="col"):
    """Kolommen per periode: AC massief + FC gearceerd (gestapeld), optioneel PL-markering en totaallabel."""
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = kind, "stacked", 100
    ch.add_data(ac, titles_from_data=True)
    ch.add_data(fc, titles_from_data=True)
    ch.set_categories(cats)
    style_chart(ch, [], numfmt=numfmt, w=w, h=h, ymin=0)
    ch.series[0].graphicalProperties = solid(NAVY)
    ch.series[1].graphicalProperties = hatched(NAVY)
    ch.gapWidth = 45
    if pl is not None or total is not None:
        ln = LineChart()
        refs = [r for r in (total, pl) if r is not None]
        for r in refs:
            ln.add_data(r, titles_from_data=True)
        ln.set_categories(cats)
        k = 0
        if total is not None:
            sr = ln.series[k]
            sr.graphicalProperties.line.noFill = True
            sr.marker = Marker(symbol="none")
            series_labels(sr, numfmt + ";-" + numfmt + ";;", "14284B", "t", True)
            k += 1
        if pl is not None:
            sr = ln.series[k]
            sr.graphicalProperties.line.noFill = True
            sr.marker = Marker(symbol="dash", size=20)
            sr.marker.graphicalProperties = GraphicalProperties(solidFill="6B7385", ln=LineProperties(solidFill="6B7385", w=19050))
        ln.y_axis.axId = 100
        ln.y_axis.delete = True
        ch += ln
    ws.add_chart(ch, anchor)
    return ch


def waterfall(ws, anchor, cats, data, w=17, h=8, numfmt="#,##0.0", start_fc=False, end_fc=True, kind="col"):
    """Waterval in Zebra-stijl (gestapeld): geen waarde-as; de as begint net onder de laagste stand (zie waterfall_rows),
    zodat ook kleine verschillen zichtbaar zijn. Totalen staan in het categorielabel. data = 5 kolommen: basis, begin, plus, min, eind."""
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = kind, "stacked", 100
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    style_chart(ch, [], numfmt=numfmt, w=w, h=h, legend=False)
    base, begin, up, down, end = ch.series
    base.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    begin.graphicalProperties = hatched(NAVY) if start_fc else solid(PLGREY)
    up.graphicalProperties = solid(GOOD)
    down.graphicalProperties = solid(BAD)
    end.graphicalProperties = hatched(NAVY) if end_fc else solid(NAVY)
    # gestapelde grafiek: alleen ctr/inEnd/inBase zijn geldige labelposities
    series_labels(up, "+" + numfmt + ";-" + numfmt + ";;", "14284B", "ctr", True)
    series_labels(down, '"−"' + numfmt + ";;;", "14284B", "ctr", True)
    ch.y_axis.delete = True
    ch.y_axis.majorGridlines = None
    if kind == "bar":
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.txPr = txt(800, "333F55")
    ch.gapWidth = 30
    ws.add_chart(ch, anchor)
    return ch


def waterfall_rows(ws, r0, c0, steps, fmt=FMT_EUR1, dec=1):
    """Hulpkolommen voor een waterval. steps: lijst (label, soort, waarde-formule), soort = begin/delta/eind.
    Kolommen vanaf c0: categorielabel, waarde, stand, basis, begin, plus, min, eind; vloer (as-begin) rechts in de kop."""
    L = [col(c0 + i) for i in range(8)]
    fl = f"${col(c0 + 9)}${r0}"
    for j, h in enumerate(["Stap", "Waarde", "Stand", "Basis", "Begin", "Plus", "Min", "Eind"]):
        put(ws, (r0, c0 + j), h, color=GREY, size=8, bold=True)
    n = len(steps)
    put(ws, (r0, c0 + 8), "as vanaf", color=GREY, size=8)
    rngc = f"{L[2]}{r0+1}:{L[2]}{r0+n}"
    # as begint ruim onder de laagste stand, zodat de stappen goed zichtbaar zijn (Zebra-stijl, geen waarde-as)
    ws[fl.replace("$", "")] = f"=MAX(0,MIN({rngc})-MAX(MAX({rngc})-MIN({rngc}),ABS(MIN({rngc}))*0.02)*1.5)"
    ws[fl.replace("$", "")].number_format = fmt
    for i, (lab, kind, f) in enumerate(steps):
        r = r0 + 1 + i
        lab_f = lab[1:] if str(lab).startswith("=") else '"' + str(lab) + '"'
        ws[f"{L[0]}{r}"] = (f"={lab_f}&CHAR(10)&FIXED({L[1]}{r},{dec})" if kind != "delta" else f"={lab_f}")
        ws[f"{L[1]}{r}"] = f
        ws[f"{L[2]}{r}"] = f"={L[2]}{r-1}+{L[1]}{r}" if kind == "delta" else f"={L[1]}{r}"
        ws[f"{L[3]}{r}"] = f"=MIN({L[2]}{r-1},{L[2]}{r})-{fl}" if kind == "delta" else "=0"
        ws[f"{L[4]}{r}"] = f"=MAX(0,{L[1]}{r}-{fl})" if kind == "begin" else "=0"
        ws[f"{L[5]}{r}"] = f"=MAX({L[1]}{r},0)" if kind == "delta" else "=0"
        ws[f"{L[6]}{r}"] = f"=MAX(-{L[1]}{r},0)" if kind == "delta" else "=0"
        ws[f"{L[7]}{r}"] = f"=MAX(0,{L[1]}{r}-{fl})" if kind == "eind" else "=0"
        for c in L:
            ws[f"{c}{r}"].font = font(8, False, GREY)
        for c in L[1:]:
            ws[f"{c}{r}"].number_format = fmt
    cats = Reference(ws, min_col=c0, min_row=r0 + 1, max_row=r0 + n)
    data = Reference(ws, min_col=c0 + 3, max_col=c0 + 7, min_row=r0, max_row=r0 + n)
    return cats, data


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
    oi = 'SUMIFS(V_PEWOmz,V_OI_Jaar,P_RJ,V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)'
    tile(ws, 17, 18, r, '="Orderintake "&P_RJ&" t/m peildatum"', f"={oi}/1000000", MLN,
         '=COUNTIFS(V_OI_Jaar,P_RJ,V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)&" projecten (PEW-omzet)"',
         '=IF(SUMIFS(V_PEWOmz,V_OI_Jaar,P_RJ,V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)=0,"PEW-BOR % n.b.","PEW-BOR %: "&FIXED(SUMIFS(V_PEWBor,V_OI_Jaar,P_RJ,V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)/SUMIFS(V_PEWOmz,V_OI_Jaar,P_RJ,V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)*100,1)&"%")')
    tile_sub = r + 3
    ws.cell(row=tile_sub, column=14).value = '="budget: "&' + S("F_WonFY", "Budget") + '&" · etalage: "&SUMIFS(V_W_Etalage,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,V_WBS,Sel_WBS)'
    # stoplichtsamenvatting
    put(ws, "R4", "Rode stoplichten", bold=True, color=NAVY, size=8)
    for c, nm, lab in (("R", "SL_Verkoop", "verkoop"), ("S", "SL_Vergunning", "vergunning"), ("T", "SL_Net", "net")):
        ws[f"{c}5"] = f'="● "&COUNTIF({nm},1)&" {lab}"'
        ws[f"{c}5"].font = font(9, True, RED_L)
        ws[f"{c}5"].hyperlink = "#'Stoplicht'!A1"
    # grafiekdata (zichtbaar onder de grafieken, zodat elke grafiek naleesbaar is)
    D0 = 50
    section(ws, D0 - 2, 2, "Grafiekdata (gekoppeld aan CALC_Versies / Projecten / Outlook — niet overschrijven)", 18)
    # A. prognoseverloop als waterval: Budget -> delta per versie -> laatste versie
    put(ws, f"B{D0-1}", "A. Prognoseverloop omzet FY (EUR mln)", bold=True, color=NAVY)
    put(ws, f"K{D0}", "Versiewaarde", color=GREY, size=8, bold=True)
    put(ws, f"L{D0}", "Doorgetrokken", color=GREY, size=8, bold=True)
    for i, v in enumerate(VERSIES):
        rr = D0 + 1 + i
        o = S("F_OmzFY", v)
        put(ws, f"J{rr}", v, color=GREY, size=8)
        ws[f"K{rr}"] = f'=IF(ISNUMBER({o}),{o}/1000000,"n.b.")'
        ws[f"L{rr}"] = f"=IF(ISNUMBER(K{rr}),K{rr},{'0' if i == 0 else f'L{rr-1}'})"
        for c in "KL":
            ws[f"{c}{rr}"].number_format = FMT_EUR1
            ws[f"{c}{rr}"].font = font(8, False, GREY)
    lastv = f'INDEX($J${D0+1}:$J${D0+5},MAX(IF(ISNUMBER(K{D0+1}),1,0),IF(ISNUMBER(K{D0+2}),2,0),IF(ISNUMBER(K{D0+3}),3,0),IF(ISNUMBER(K{D0+4}),4,0),IF(ISNUMBER(K{D0+5}),5,0),1))'
    steps = [("Budget", "begin", f"=L{D0+1}")]
    steps += [(f"Δ {v}", "delta", f"=L{D0+1+i}-L{D0+i}") for i, v in enumerate(VERSIES[1:], start=1)]
    steps += [("Prognose", "eind", f"=L{D0+5}")]
    catsA, dataA = waterfall_rows(ws, D0, 2, steps)
    ws[f"B{D0+len(steps)}"] = f'="FC "&{lastv}&CHAR(10)&FIXED(C{D0+len(steps)},1)'
    # B. realisatie vs to go per project (top 10)
    T0 = D0 + 10
    put(ws, f"B{T0-1}", "B. Top 10 projecten: realisatie vs to go (EUR mln, geselecteerde versie)", bold=True, color=NAVY)
    header(ws, T0, 2, ["#", "Project", "Realisatie YTD", "To go", "FY"], height=18)
    for k in range(1, 11):
        rr = T0 + k
        put(ws, f"B{rr}", k)
        ix = f"IFERROR(MATCH(LARGE(Projecten!$Z${P0}:$Z${P1},{k}),Projecten!$Z${P0}:$Z${P1},0),0)"
        ws[f"H{rr}"] = f"={ix}"
        helper_style(ws[f"H{rr}"])
        q = f"$H{rr}"
        ws[f"C{rr}"] = f'=IF({q}=0,"",INDEX(Projecten!$B${P0}:$B${P1},{q}))'
        ws[f"D{rr}"] = f'=IF({q}=0,0,IF(ISNUMBER(INDEX(Projecten!$L${P0}:$L${P1},{q})),INDEX(Projecten!$L${P0}:$L${P1},{q})*P_Factor/1000000,0))'
        ws[f"F{rr}"] = f'=IF({q}=0,0,INDEX(Projecten!$K${P0}:$K${P1},{q})*P_Factor/1000000)'
        ws[f"E{rr}"] = f'=IF({q}=0,0,IF(ISNUMBER(INDEX(Projecten!$M${P0}:$M${P1},{q})),INDEX(Projecten!$M${P0}:$M${P1},{q})*P_Factor/1000000,0))'
        for c in "DEF":
            ws[f"{c}{rr}"].number_format = FMT_EUR1
    # C. brug vorige versie -> projectmutaties -> geselecteerde versie
    B0 = T0 + 13
    put(ws, f"B{B0-1}", "C. Brug vorige → geselecteerde versie: grootste projectmutaties (EUR mln)", bold=True, color=NAVY)
    put(ws, f"K{B0}", "#", color=GREY, size=8, bold=True)
    vor, sel = S("F_OmzFY", "Vorige"), S("F_OmzFY")
    bsteps = [('=IF(Sel_Prev="","Vorige",Sel_Prev)', "begin", f"=IF(ISNUMBER({vor}),{vor}/1000000,0)")]
    for k in range(1, 9):
        rr = B0 + 1 + k
        ws[f"K{rr}"] = f"=IFERROR(MATCH(LARGE(Projecten!$AA${P0}:$AA${P1},{k}),Projecten!$AA${P0}:$AA${P1},0),0)"
        helper_style(ws[f"K{rr}"])
        bsteps.append((f'=IF($K{rr}=0,"",LEFT(INDEX(Projecten!$B${P0}:$B${P1},$K{rr}),18))', "delta",
                       f"=IF($K{rr}=0,0,INDEX(Projecten!$Q${P0}:$Q${P1},$K{rr})*P_Factor/1000000)"))
    bsteps.append(("Overige", "delta", f"=IF(AND(ISNUMBER({vor}),ISNUMBER({sel})),({sel}-{vor})/1000000-SUM(C{B0+2}:C{B0+9}),0)"))
    bsteps.append(("=Sel_Versie", "eind", f"=IF(ISNUMBER({sel}),{sel}/1000000,0)"))
    catsC, dataC = waterfall_rows(ws, B0, 2, bsteps, fmt=FMT_EUR2, dec=2)
    put(ws, f"B{B0+len(bsteps)+1}", '=IF(Sel_Prev="","Budget heeft geen vorige versie: geen brug.","")', color=GREY, italic=True, size=8)
    # grafieken met dynamische titels in cellen
    gt = [("B12", '="Prognoseverloop omzet FY "&Sel_Doeljaar&": budget → per versie → "&B%d&" (EUR mln)"' % (D0 + len(steps))),
          ("K12", '="Realisatie (massief) vs nog te gaan (gearceerd) — top 10 projecten, "&Sel_Label&" (EUR mln)"'),
          ("B30", '="Brug "&IF(Sel_Prev="","–",Sel_Prev)&" → "&Sel_Versie&": grootste projectmutaties (Δ FY-omzet, EUR mln)"'),
          ("K30", '="Outlook omzet per jaar: realisatie massief, prognose gearceerd, budget = streep (EUR mln)"')]
    for ref, f in gt:
        ws[ref] = f
        ws[ref].font = font(10, True, NAVY)
    waterfall(ws, "B13", catsA, dataA, w=17, h=8)
    ibcs_columns(ws, "K13", Reference(ws, min_col=3, min_row=T0 + 1, max_row=T0 + 10),
                 Reference(ws, min_col=4, min_row=T0, max_row=T0 + 10), Reference(ws, min_col=5, min_row=T0, max_row=T0 + 10),
                 kind="bar", w=17, h=8)
    waterfall(ws, "B31", catsC, dataC, w=17, h=8, numfmt="#,##0.00", kind="bar")
    DASH.update(catsA=catsA, dataA=dataA, catsC=catsC, dataC=dataC, T0=T0, D0=D0, stepsA=len(steps), stepsC=len(bsteps))
    put(ws, "B47", "IBCS-notatie: massief = realisatie of vastgelegd, gearceerd = prognose, grijs = budget, groen/rood = verbetering/verslechtering. "
        "Een versie zonder vastgelegde waarde telt als 'geen wijziging' in de waterval.", color=GREY, italic=True, size=8)
    widths(ws, {c: 9.5 for c in "BCDEFGHIJKLMNOPQRST"})
    ws.column_dimensions["I"].width = 22
    ws.column_dimensions["O"].width = 22
    ws.column_dimensions["C"].width = 14
    landscape(ws, "A1:T47")


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
    ws[f"H{r+1}"] = '="Verkopen per kwartaal "&P_RJ&"–"&(P_RJ+1)&" (aantal): gerealiseerd massief, prognose gearceerd"'
    ws[f"H{r+1}"].font = font(10, True, NAVY)
    chq = bar(ws, Reference(ws, min_col=3, min_row=q0 - 1, max_col=4, max_row=qlast), Reference(ws, min_col=2, min_row=q0, max_row=qlast),
              f"H{r+2}", [NAVY, NAVY], grouping="stacked", numfmt="#,##0", w=17, h=7.5, ymin=0)
    chq.series[1].graphicalProperties = hatched(NAVY)
    WON.update(q0=q0, qlast=qlast)
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
    add_name(wb, "PPT_WonVersieTitel", f"Woningen!$H${r+1}")
    WON.update(v0=v0, vtitle=f"H{r+1}", qtitle="H11")
    ws[f"H{r+1}"].font = font(10, True, NAVY)
    bar(ws, Reference(ws, min_col=17, min_row=v0 - 1, max_row=v0 + 5), Reference(ws, min_col=3, min_row=v0, max_row=v0 + 5),
        f"H{r+2}", [NAVY], numfmt="#,##0", w=17, h=6, labels=True, legend=False, ymin=0, point_colors={5: NAVY_LT})
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
    put(ws, f"H{r}", "Segment (uit Invoer › Projecten — fictieve demo-indeling)", bold=True, color=NAVY)
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
        f"K{r+2}", [NAVY, NAVY_LT], numfmt="#,##0", w=14, h=6)
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
        f"B{r+6}", [NAVY, NAVY], grouping="stacked", numfmt="#,##0", w=20, h=6.5, ymin=0).series[1].graphicalProperties = hatched(NAVY)
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
# Stoplicht: invoer per project, kleur volgt automatisch uit de uitgangspunten rechts
# ---------------------------------------------------------------------------------------------
SL0, SL1 = 8, 8 + (MAP1 - MAP0)
SL_FMT = '"●";"●";"○"'
SL_DEMO = [  # fictieve voorbeeldinvoer: omgevingsvergunning, (vervallen: RvS), netcongestie
    ("Ja", "Geen", "Geen"), ("Ja", "Geen", "Laag"), ("Ja", "Geen", "Laag"), ("Nee", "Geen", "Laag"),
    ("In procedure", "Laag", "Geen"), ("Ja", "Geen", "Laag"), ("Ja", "Geen", "Laag"), ("Ja", "Hoog", "Midden"),
    ("Ja", "Geen", "Geen"), ("Ja", "Geen", "Geen"), ("Nee", "Midden", "Laag"), ("Ja", "Geen", "Laag"),
    ("Ja", "Geen", "Geen"), ("In procedure", "Laag", "Laag"), ("Ja", "Geen", "Laag"), ("Ja", "Geen", "Midden"),
    ("Ja", "Geen", "Laag"), ("Ja", "Laag", "Laag"), ("Ja", "Geen", "Geen"), ("Ja", "Geen", "Hoog"),
]


SL_SB = [1.0, 0.0, 1.0, 0.5, 1.0, 0.0, 0.7, 1.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.7, 0.9, 0.55, 1.0, 0.3, 0.8, None]  # fictief


def stoplicht_cf(ws, rng_):
    tl = rng_.split(":")[0].replace("$", "")
    for code, colr in ((3, GREEN_L), (2, ORANGE_L), (1, RED_L), (0, OFF_L)):
        ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'AND(ISNUMBER({tl}),{tl}={code})'], font=Font(color=colr)))


def build_stoplicht(wb, ws):
    base_sheet(ws, zoom=100)
    title(ws, "Stoplicht projecten", "Vul de gele cellen in. De kleur volgt automatisch uit de uitgangspunten rechts.", 19)
    put(ws, "A4", "Gele cellen zijn fictieve voorbeeldinvoer — overschrijf met de eigen beoordeling. Verkocht % (bron) = cumulatief verkocht / totaal uit de Woningetalage; "
        "'Verkocht % SB' (verkocht bij start bouw) gaat voor de bron als die is ingevuld.", color=GREY, italic=True, size=8)
    groups = [(1, 4, "PROJECT", NAVY), (5, 8, "INVOER (gele cellen)", "C9A000"), (9, 11, "STOPLICHT", NAVY)]
    for c1, c2, t, colr in groups:
        ws.merge_cells(start_row=6, start_column=c1, end_row=6, end_column=c2)
        cell = ws.cell(row=6, column=c1, value=t)
        cell.font = font(9, True, WHITE)
        cell.alignment = Alignment(horizontal="center")
        for c in range(c1, c2 + 1):
            ws.cell(row=6, column=c).fill = fill(colr)
    header(ws, 7, 1, ["Project", "Type", "Order-\nintake", "Verkocht %\n(bron)", "Verkocht %\nSB (handm.)", "Omgevings-\nvergunning",
                      "Net-\ncongestie", "Toelichting", "Verkoop", "Omgevings-\nvergunning", "Net-\ncongestie"], height=34)
    for c in range(5, 9):
        ws.cell(row=7, column=c).fill = fill("8A6A00")
    for r in range(SL0, SL1 + 1):
        i = r - SL0
        vr = V0 + i
        a = f"VERW_Project!$A{vr}"
        ws[f"A{r}"] = f'=IF({a}="","",VERW_Project!$C{vr})'
        ws[f"B{r}"] = f'=IF({a}="","",VERW_Project!$D{vr})'
        ws[f"C{r}"] = f'=IF({a}="","",VERW_Project!${VCOLS["OI_Datum"]}{vr})'
        ws[f"D{r}"] = f'=IF({a}="","",VERW_Project!${VCOLS["VerkochtPct"]}{vr})'
        omg, _rvs, net = SL_DEMO[i] if i < len(SL_DEMO) else (None, None, None)
        sb = SL_SB[i] if i < len(SL_SB) else None
        input_style(ws.cell(row=r, column=5, value=sb), "0%")
        input_style(ws.cell(row=r, column=6, value=omg))
        input_style(ws.cell(row=r, column=7, value=net))
        input_style(ws[f"H{r}"])
        wbs_now = wb["Invoer"].cell(row=MAP0 + i, column=1).value
        if wbs_now and i < len(SL_DEMO):
            ws[f"P{r}"] = wbs_now  # stempel van de voorbeeldinvoer
        pct = f'IF(E{r}<>"",E{r},D{r})'
        ws[f"I{r}"] = f'=IF(A{r}="","",IF(ISNUMBER({pct}),IF({pct}>=SL_Groen,3,IF({pct}>=SL_Oranje,2,1)),0))'
        for c, src, names in (("J", "F", "OV"), ("K", "G", "NET")):
            ws[f"{c}{r}"] = f'=IF(A{r}="","",IF({src}{r}="",0,IFERROR(--INDEX({names}_C,MATCH({src}{r},{names}_W,0)),0)))'
        ws[f"O{r}"] = f'=IF({a}="","",{a})'
        ws[f"N{r}"] = f'=IF(A{r}="","",IF(COUNTIF(I{r}:K{r},1)>0,1,0))'
        ws[f"Q{r}"] = f'=IF(N{r}=1,COUNTIF($N${SL0}:N{r},1),"")'
        for c in "IJK":
            cell = ws[f"{c}{r}"]
            cell.number_format = SL_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = font(16)
            cell.border = Border(bottom=thin, left=thin if c == "I" else None)
        for c in "ABCD":
            ws[f"{c}{r}"].font = font(9)
            ws[f"{c}{r}"].border = Border(bottom=thin)
        ws[f"C{r}"].number_format = "mm-yyyy"
        ws[f"D{r}"].number_format = "0%"
        ws[f"D{r}"].alignment = Alignment(horizontal="right")
        for c in "NOPQ":
            helper_style(ws[f"{c}{r}"])
        ws.row_dimensions[r].height = 17
    stoplicht_cf(ws, f"I{SL0}:K{SL1}")
    add_name(wb, "SL_Verkoop", f"Stoplicht!$I${SL0}:$I${SL1}")
    add_name(wb, "SL_Vergunning", f"Stoplicht!$J${SL0}:$J${SL1}")
    add_name(wb, "SL_Net", f"Stoplicht!$K${SL0}:$K${SL1}")
    add_name(wb, "SL_Lichten", f"Stoplicht!$I${SL0}:$K${SL1}")
    add_name(wb, "SL_RoodNr", f"Stoplicht!$Q${SL0}:$Q${SL1}")
    add_name(wb, "SL_Project", f"Stoplicht!$A${SL0}:$A${SL1}")
    add_name(wb, "SL_WBS", f"Stoplicht!$O${SL0}:$O${SL1}")
    add_name(wb, "SL_Stempel", f"Stoplicht!$P${SL0}:$P${SL1}")
    put(ws, "O7", "WBS (hulp)", color=GREY, size=7)
    put(ws, "P7", "WBS bij invoer", color=GREY, size=7)
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True)
    dv.error = "Vul een percentage in tussen 0% en 100%."
    ws.add_data_validation(dv)
    dv.add(f"E{SL0}:E{SL1}")
    dv_list(ws, f"F{SL0}:F{SL1}", "OV_W")
    dv_list(ws, f"G{SL0}:G{SL1}", "NET_W")

    # instellingen & legenda
    c0 = 19  # kolom S
    for c in range(c0, c0 + 4):
        ws.cell(row=6, column=c).fill = fill(NAVY)
    put(ws, (6, c0), "INSTELLINGEN & LEGENDA", bold=True, color=WHITE, fillc=NAVY)
    put(ws, (8, c0), "Betekenis", bold=True, color=NAVY, size=10)
    leg = [(3, "Groen — op koers"), (2, "Oranje — aandacht"), (1, "Rood — knelpunt / niet gehaald"), (0, "Leeg — niet ingevuld of n.v.t.")]
    for i, (code, t) in enumerate(leg):
        r = 9 + i
        cell = ws.cell(row=r, column=c0, value=code)
        cell.number_format = SL_FMT
        cell.alignment = Alignment(horizontal="center")
        cell.font = font(15)
        put(ws, (r, c0 + 1), t)
    stoplicht_cf(ws, f"{col(c0)}9:{col(c0)}12")
    put(ws, (14, c0), "Verkoop (verkocht %)", bold=True, color=NAVY, size=10)
    put(ws, (15, c0 + 1), "Groen vanaf")
    input_style(put(ws, (15, c0 + 2), 0.7, fmt="0%"), "0%")
    put(ws, (16, c0 + 1), "Oranje vanaf")
    input_style(put(ws, (16, c0 + 2), 0.5, fmt="0%"), "0%")
    put(ws, (17, c0 + 1), "Daaronder rood. Verkocht % SB gaat vóór de bron.", color=GREY, italic=True, size=8)
    add_name(wb, "SL_Groen", f"Stoplicht!${col(c0+2)}$15")
    add_name(wb, "SL_Oranje", f"Stoplicht!${col(c0+2)}$16")
    tables = [("OV", "Omgevingsvergunning", [("Ja", "Groen"), ("In procedure", "Oranje"), ("Nee", "Rood")]),
              ("NET", "Netcongestie", [("Geen", "Groen"), ("Laag", "Oranje"), ("Midden", "Oranje"), ("Hoog", "Rood")])]
    r = 19
    for key, t, trows in tables:
        put(ws, (r, c0), t, bold=True, color=NAVY, size=10)
        header(ws, r + 1, c0 + 1, ["Waarde", "Kleur", "Code"], height=16, color=NAVY2)
        for k in range(5):
            rr = r + 2 + k
            w_, kl = trows[k] if k < len(trows) else (None, None)
            input_style(ws.cell(row=rr, column=c0 + 1, value=w_))
            input_style(ws.cell(row=rr, column=c0 + 2, value=kl))
            kc = col(c0 + 2)
            ws.cell(row=rr, column=c0 + 3, value=f'=IF({kc}{rr}="Groen",3,IF({kc}{rr}="Oranje",2,IF({kc}{rr}="Rood",1,"")))').font = font(8, False, GREY)
        rng_k = f"{col(c0+2)}{r+2}:{col(c0+2)}{r+6}"
        dv_list(ws, rng_k, "L_Kleur")
        for kl, colr in (("Groen", "D5EFD9"), ("Oranje", "FDEBD0"), ("Rood", "F8D7DA")):
            ws.conditional_formatting.add(rng_k, FormulaRule(formula=[f'{col(c0+2)}{r+2}="{kl}"'], fill=fill(colr)))
        add_name(wb, f"{key}_W", f"Stoplicht!${col(c0+1)}${r+2}:${col(c0+1)}${r+6}")
        add_name(wb, f"{key}_C", f"Stoplicht!${col(c0+3)}${r+2}:${col(c0+3)}${r+6}")
        r += 9
    put(ws, (r, c0), "Samenvatting (aantal projecten)", bold=True, color=NAVY, size=10)
    header(ws, r + 1, c0 + 1, ["Licht", "Rood", "Oranje"], height=16, color=NAVY2)
    for k, (lab, nm) in enumerate((("Verkoop", "SL_Verkoop"), ("Omgevingsvergunning", "SL_Vergunning"), ("Netcongestie", "SL_Net"))):
        rr = r + 2 + k
        put(ws, (rr, c0 + 1), lab)
        ws.cell(row=rr, column=c0 + 2, value=f"=COUNTIF({nm},1)").font = font(9, True, RED_L)
        ws.cell(row=rr, column=c0 + 3, value=f"=COUNTIF({nm},2)").font = font(9, True, ORANGE_L)
    put(ws, (r + 6, c0 + 1), "Nieuwe waarde nodig? Vul een lege regel in een tabel hierboven.", color=GREY, size=8)
    widths(ws, {"A": 28, "B": 11, "C": 9, "D": 10, "E": 11, "F": 13, "G": 11, "H": 34, "I": 10, "J": 12, "K": 11,
                "L": 3, "M": 3, "N": 3, "O": 14, "P": 14, "Q": 4, "R": 3, "S": 7, "T": 22, "U": 10, "V": 9})
    for c in "NOPQ":
        ws.column_dimensions[c].hidden = True
    ws.freeze_panes = f"B{SL0}"
    landscape(ws, f"A1:V{SL0 + 24}")


# ---------------------------------------------------------------------------------------------
# Outlook: alle beschikbare jaren naast elkaar (omzet, BOR, woningverkopen, orderintake) + budget
# ---------------------------------------------------------------------------------------------
OI0, OI1 = 92, 92 + (V1 - V0)
OUT_YEARS = 8          # RJ-1 .. RJ+6
OUT_CH = 15            # kolom O: grafiekdata


OUT_CTX = {}


def outlook_chart(target_ws, anchor, which, w, h, fmt=None):
    """IBCS-kolommen uit de grafiekdata van Outlook (rijen): realisatie massief, prognose gearceerd, budget als streep."""
    cats, rowref = OUT_CTX["cats"], OUT_CTX["rowref"]
    base = OUT_CTX["g0"] + {"omzet": 0, "won": 5, "oi": 10}[which]
    fmt = fmt or ("#,##0" if which == "won" else "#,##0.0")
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
    ch.add_data(rowref(base), from_rows=True, titles_from_data=True)
    ch.add_data(rowref(base + 1), from_rows=True, titles_from_data=True)
    ch.set_categories(cats)
    style_chart(ch, [], numfmt=fmt, w=w, h=h, ymin=0)
    ch.series[0].graphicalProperties = solid(NAVY)
    ch.series[1].graphicalProperties = hatched(NAVY)
    ch.gapWidth = 45
    ln = LineChart()
    ln.add_data(rowref(base + 2), from_rows=True, titles_from_data=True)
    ln.add_data(rowref(base + 3), from_rows=True, titles_from_data=True)
    ln.set_categories(cats)
    tot, pl = ln.series
    tot.graphicalProperties.line.noFill = True
    tot.marker = Marker(symbol="none")
    series_labels(tot, fmt + ";-" + fmt + ";;", "14284B", "t", True)
    pl.graphicalProperties.line.noFill = True
    pl.marker = Marker(symbol="dash", size=20)
    pl.marker.graphicalProperties = GraphicalProperties(solidFill="6B7385", ln=LineProperties(solidFill="6B7385", w=19050))
    ln.y_axis.axId = 100
    ln.y_axis.delete = True
    ln.y_axis.scaling.min = 0
    ch += ln
    target_ws.add_chart(ch, anchor)
    return ch


def build_outlook(wb, ws):
    base_sheet(ws, zoom=90)
    title(ws, None, None, 13)
    ws["B1"] = '="Outlook "&P_Regio&" — "&(P_RJ-1)&" t/m "&(P_RJ+6)&"  ·  basis voor het budget"'
    ws["B1"].font = font(14, True, WHITE)
    ws["B2"] = ('="Bedragen in "&P_Eenheid&" · bron: upload peildatum "&P_PeilTxt&" · budget = vastgelegd Budget "&P_VJ'
                '&" · filter: "&Sel_FiltLabel&" · "&(P_RJ-1)&" = realisatie, "&P_RJ&" = realisatie + prognose, daarna prognose"')
    ws["B2"].font = font(9, False, GREY, True)
    yc = [col(3 + i) for i in range(OUT_YEARS)]          # C..J
    later, totc = col(3 + OUT_YEARS), col(4 + OUT_YEARS)   # K, L
    hdr = ['=(P_RJ-1)&CHAR(10)&"realisatie"', '=P_RJ&CHAR(10)&"FY"'] + [f'=(P_RJ+{o})&""' for o in range(1, OUT_YEARS - 1)] + \
          ['=(P_RJ+7)&" e.v."', '="Totaal vanaf "&P_RJ']
    section(ws, 4, 2, "Meerjarenoutlook", 11)
    header(ws, 5, 2, ["Kengetal"] + hdr, height=30)
    crit = "V_WBS,Sel_WBS"
    tf = lambda t: f'IF(OR(Sel_Type="Alle",Sel_Type="{t}"),1,0)'  # noqa: E731
    money_by_year = ["OmzVJ", "OmzFY", "OmzFY1"] + [f"OmzRJ{o}" for o in range(2, 7)]
    bor_by_year = ["BorVJ", "BorFY", "BorFY1"] + [f"BorRJ{o}" for o in range(2, 7)]
    won_by_year = [None, "W_HuidigJaar", "W_VolgendJaar"] + [f"W_Y_RJ{o}" for o in range(2, 7)]
    yr = lambda i: f"(P_RJ-1+{i})"  # noqa: E731
    rows = {}
    spec = [
        ("omz_ops", "Omzet opstal", "money"), ("omz_dln", "Omzet deelnemingen", "money"), ("omz", "Omzet totaal", "sum"),
        ("omz_ac", "  waarvan gerealiseerd", "money"), ("bor", "BOR totaal", "money"), ("borp", "BOR %", "pct"),
        ("bud", "Budget omzet (vastgelegd)", "money"), ("dbud", "Δ outlook vs budget", "money"), (None,),
        ("won", "Woningverkopen (aantal)", "cnt"), ("won_ac", "  waarvan gerealiseerd", "cnt"), ("wbud", "Budget woningverkopen", "cnt"), (None,),
        ("oi_n", "Orderintake (# projecten)", "cnt"), ("oi", "Orderintake PEW-omzet", "money"), ("oi_ac", "  waarvan gerealiseerd", "money"),
        ("oi_bor", "Orderintake PEW-BOR %", "pct"), ("oi_bud", "Budget orderintake", "money"),
    ]
    r = 6
    for sp in spec:
        if sp[0]:
            rows[sp[0]] = r
        r += 1
    for sp in spec:
        if not sp[0]:
            continue
        key, lab, kind = sp
        r = rows[key]
        put(ws, f"B{r}", lab, bold=key in ("omz", "won", "oi"), color=GREY if lab.startswith("  ") else "000000")
        for i, c in enumerate(yc):
            y = yr(i)
            f = '""'
            if key in ("omz_ops", "omz_dln"):
                t = "Opstal" if key == "omz_ops" else "Deelneming"
                f = f'{tf(t)}*SUMIFS(V_{money_by_year[i]},V_Type,"{t}",{crit})/P_Factor'
            elif key == "omz":
                f = f"{c}{rows['omz_ops']}+{c}{rows['omz_dln']}"
            elif key == "omz_ac":
                f = f"{c}{rows['omz']}" if i == 0 else (f'SUMIFS(V_OmzYTD,V_Type,Sel_TypeCrit,{crit})/P_Factor' if i == 1 else "0")
            elif key == "bor":
                f = f'SUMIFS(V_{bor_by_year[i]},V_Type,Sel_TypeCrit,{crit})/P_Factor'
            elif key == "borp":
                f = f'IF({c}{rows["omz"]}=0,"n.b.",{c}{rows["bor"]}/{c}{rows["omz"]})'
            elif key == "bud":
                cr = f'L_Actief,1,L_Versie,"Budget",L_Versiejaar,P_VJ,L_Doeljaar,{y},L_Kengetal,"{K_OMZ}",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS'
                f = f'IF(COUNTIFS({cr})=0,"n.b.",SUMIFS(L_Waarde,{cr})/P_Factor)'
            elif key == "dbud":
                f = f'IF(ISNUMBER({c}{rows["bud"]}),{c}{rows["omz"]}-{c}{rows["bud"]},"n.b.")'
            elif key == "won":
                f = '"n.b."' if i == 0 else f'SUMIFS(V_{won_by_year[i]},V_Meetellen,"Ja",V_Type,Sel_TypeCrit,{crit})'
            elif key == "won_ac":
                f = '"n.b."' if i == 0 else (f'SUMIFS(V_W_ActYTD,V_Meetellen,"Ja",V_Type,Sel_TypeCrit,{crit})' if i == 1 else "0")
            elif key == "wbud":
                cr = f'L_Actief,1,L_Versie,"Budget",L_Versiejaar,P_VJ,L_Doeljaar,{y},L_Kengetal,"{K_WFY}",L_Type,Sel_TypeCrit,L_WBS,Sel_WBS'
                f = f'IF(COUNTIFS({cr})=0,"n.b.",SUMIFS(L_Waarde,{cr}))'
            elif key == "oi_n":
                f = f"COUNTIFS(V_OI_Jaar,{y},V_Type,Sel_TypeCrit,{crit})"
            elif key == "oi":
                f = f"SUMIFS(V_PEWOmz,V_OI_Jaar,{y},V_Type,Sel_TypeCrit,{crit})/P_Factor"
            elif key == "oi_ac":
                f = f'SUMIFS(V_PEWOmz,V_OI_Jaar,{y},V_OI_Datum,"<="&P_Peildatum,V_Type,Sel_TypeCrit,{crit})/P_Factor'
            elif key == "oi_bor":
                f = f'IF({c}{rows["oi"]}=0,"",SUMIFS(V_PEWBor,V_OI_Jaar,{y},V_Type,Sel_TypeCrit,{crit})/P_Factor/{c}{rows["oi"]})'
            elif key == "oi_bud":
                cr = f'A_Doel,{y},A_Rubriek,"Orderintake budget",A_Versie,"Budget",A_TeltVast,1'
                f = f'IF(COUNTIFS({cr})=0,"n.b.",SUMIFS(A_Waarde,{cr})/P_Factor)'
            ws[f"{c}{r}"] = "=" + f
        # later en totaal
        if key in ("omz_ops", "omz_dln", "bor"):
            if key == "bor":
                ws[f"{later}{r}"] = f'=SUMIFS(V_BorRJ7,V_Type,Sel_TypeCrit,{crit})/P_Factor+SUMIFS(V_BorRJ8,V_Type,Sel_TypeCrit,{crit})/P_Factor+SUMIFS(V_BorONV,V_Type,Sel_TypeCrit,{crit})/P_Factor'
            else:
                t = "Opstal" if key == "omz_ops" else "Deelneming"
                ws[f"{later}{r}"] = "=" + "+".join(f'{tf(t)}*SUMIFS(V_{fld},V_Type,"{t}",{crit})/P_Factor' for fld in ("OmzRJ7", "OmzRJ8", "OmzONV"))
            ws[f"{totc}{r}"] = f"=SUM({yc[1]}{r}:{later}{r})"
        elif key == "omz":
            ws[f"{later}{r}"] = f"={later}{rows['omz_ops']}+{later}{rows['omz_dln']}"
            ws[f"{totc}{r}"] = f"=SUM({yc[1]}{r}:{later}{r})"
        elif key == "borp":
            ws[f"{totc}{r}"] = f'=IF({totc}{rows["omz"]}=0,"n.b.",{totc}{rows["bor"]}/{totc}{rows["omz"]})'
        elif key in ("won",):
            ws[f"{totc}{r}"] = f"=SUM({yc[1]}{r}:{yc[-1]}{r})"
        for c in yc + [later, totc]:
            cell = ws[f"{c}{r}"]
            cell.number_format = FMT_PCT if kind == "pct" else (FMT_INT if kind == "cnt" else FMT_EUR)
            cell.alignment = Alignment(horizontal="right")
            cell.font = font(9, key in ("omz", "won", "oi"), GREY if lab.startswith("  ") else "000000")
            cell.border = Border(bottom=thin, top=thick_navy if key in ("omz",) else None)
        ws[f"B{r}"].border = Border(bottom=thin, top=thick_navy if key in ("omz",) else None)
        ws[f"{yc[0]}{r}"].fill = fill("EEF3EA")
    ws.conditional_formatting.add(f"C{rows['dbud']}:L{rows['dbud']}", FormulaRule(formula=[f'AND(ISNUMBER(C{rows["dbud"]}),C{rows["dbud"]}<0)'], font=Font(color=BAD, bold=True)))
    ws.conditional_formatting.add(f"C{rows['dbud']}:L{rows['dbud']}", FormulaRule(formula=[f'AND(ISNUMBER(C{rows["dbud"]}),C{rows["dbud"]}>0)'], font=Font(color="3B7A1E", bold=True)))
    put(ws, f"B{rows['oi_bud']+1}", "Gerealiseerd = t/m peildatum. Woningen 'n.b.' bij realisatie vorig jaar: de Woningetalage geeft alleen een cumulatief t/m vorig jaar. "
        "Budget orderintake: Invoer › 3, rubriek 'Orderintake budget', versie Budget.", color=GREY, italic=True, size=8)
    add_name(wb, "PPT_OUT", f"Outlook!$B$5:${totc}${rows['oi_bud']}")

    # grafiekdata (rechts, naleesbaar)
    gc = [col(OUT_CH + 1 + i) for i in range(OUT_YEARS)]
    put(ws, (4, OUT_CH), "Grafiekdata (niet wijzigen) — 'NB' = bewust leeg in de grafiek", color=GREY, size=8, bold=True)
    put(ws, (5, OUT_CH), "Reeks", color=GREY, size=8, bold=True)
    for i, c in enumerate(gc):
        ws[f"{c}5"] = f"={yr(i)}&\"\""
        ws[f"{c}5"].font = font(8, False, GREY)
    gdata = [
        ("Realisatie", lambda c, yc_: f"={yc_}{rows['omz_ac']}*P_Factor/1000000"),
        ("Prognose", lambda c, yc_: f"=({yc_}{rows['omz']}-{yc_}{rows['omz_ac']})*P_Factor/1000000"),
        ("Totaal", lambda c, yc_: f"=IF({yc_}{rows['omz']}=0,NA(),{yc_}{rows['omz']}*P_Factor/1000000)"),
        ("Budget", lambda c, yc_: f"=IF(ISNUMBER({yc_}{rows['bud']}),{yc_}{rows['bud']}*P_Factor/1000000,NA())"),
        (None, None),
        ("Gerealiseerd", lambda c, yc_: f'=IF(ISNUMBER({yc_}{rows["won_ac"]}),{yc_}{rows["won_ac"]},0)'),
        ("Prognose", lambda c, yc_: f'=IF(ISNUMBER({yc_}{rows["won"]}),{yc_}{rows["won"]}-{yc_}{rows["won_ac"]},0)'),
        ("Totaal", lambda c, yc_: f'=IF(ISNUMBER({yc_}{rows["won"]}),IF({yc_}{rows["won"]}=0,NA(),{yc_}{rows["won"]}),NA())'),
        ("Budget", lambda c, yc_: f"=IF(ISNUMBER({yc_}{rows['wbud']}),{yc_}{rows['wbud']},NA())"),
        (None, None),
        ("Gerealiseerd", lambda c, yc_: f"={yc_}{rows['oi_ac']}*P_Factor/1000000"),
        ("Verwacht", lambda c, yc_: f"=({yc_}{rows['oi']}-{yc_}{rows['oi_ac']})*P_Factor/1000000"),
        ("Totaal", lambda c, yc_: f"=IF({yc_}{rows['oi']}=0,NA(),{yc_}{rows['oi']}*P_Factor/1000000)"),
        ("Budget", lambda c, yc_: f"=IF(ISNUMBER({yc_}{rows['oi_bud']}),{yc_}{rows['oi_bud']}*P_Factor/1000000,NA())"),
    ]
    g0 = 6
    for j, (lab, fn) in enumerate(gdata):
        rr = g0 + j
        if lab is None:
            continue
        put(ws, (rr, OUT_CH), lab, color=GREY, size=8)
        for i, c in enumerate(gc):
            ws[f"{c}{rr}"] = fn(c, yc[i])
            ws[f"{c}{rr}"].number_format = "#,##0.0"
            ws[f"{c}{rr}"].font = font(8, False, GREY)
    # jaren met gegevens: vlag per jaar en OUT_N = laatste jaar met gegevens (macro 'OutlookJaren' past grafieken/kolommen hierop aan)
    put(ws, (22, OUT_CH), "Jaar heeft gegevens", color=GREY, size=8)
    put(ws, (23, OUT_CH), "Volgnummer", color=GREY, size=8)
    for i, c in enumerate(gc):
        ws[f"{c}22"] = (f'=IF(OR(N({yc[i]}{rows["omz"]})<>0,N({yc[i]}{rows["won"]})<>0,N({yc[i]}{rows["oi"]})<>0,'
                        f'N({yc[i]}{rows["bor"]})<>0),1,0)')
        ws[f"{c}23"] = f"={c}22*{i+1}"
        for rr in (22, 23):
            ws[f"{c}{rr}"].font = font(8, False, GREY)
    put(ws, (24, OUT_CH), "Laatste jaar", color=GREY, size=8)
    ws[f"{gc[0]}24"] = f"=MAX(2,MAX({gc[0]}23:{gc[-1]}23))"
    add_name(wb, "OUT_N", f"Outlook!${gc[0]}$24")
    later_r = rows["omz"]
    ws[f"{gc[1]}24"] = f'=IF(OR(N({later}{rows["omz"]})<>0,N({later}{rows["bor"]})<>0),1,0)'
    add_name(wb, "OUT_Later", f"Outlook!${gc[1]}$24")
    nyr = OUT_CTX.get("demo_n", OUT_YEARS)
    cats = Reference(ws, min_col=OUT_CH + 1, max_col=OUT_CH + nyr, min_row=5)

    def rowref(rr):
        return Reference(ws, min_col=OUT_CH, max_col=OUT_CH + nyr, min_row=rr)

    # kolommen zonder gegevens verbergen (startsituatie; de macro herhaalt dit na elke inleesronde)
    for i, c in enumerate(yc):
        if i >= nyr:
            ws.column_dimensions[c].hidden = True
    if not OUT_CTX.get("demo_later", True):
        ws.column_dimensions[later].hidden = True

    OUT_CTX.update(ws=ws, cats=cats, rowref=rowref, g0=g0, rows=rows)
    titles = [("B26", '="Omzet per jaar (EUR mln): realisatie massief · prognose gearceerd · budget = streep"', g0),
              ("H26", '="Woningverkopen per jaar (aantal)"', g0 + 5),
              ("B44", '="Orderintake per jaar (PEW-omzet, EUR mln): gerealiseerd massief · verwacht gearceerd"', g0 + 10)]
    for ref, f, base in titles:
        ws[ref] = f
        ws[ref].font = font(10, True, NAVY)
    outlook_chart(ws, "B27", "omzet", 15.5, 7.5)
    outlook_chart(ws, "H27", "won", 14.5, 7.5)
    outlook_chart(ws, "B45", "oi", 15.5, 7.5)

    # orderintake per jaar en kwartaal
    q0 = 46
    put(ws, f"H{q0-2}", "Orderintake per jaar en kwartaal (PEW-omzet)", bold=True, color=NAVY, size=10)
    header(ws, q0 - 1, 8, ["Jaar", "Q1", "Q2", "Q3", "Q4", "Totaal", "#"], height=20)
    for i in range(OUT_YEARS):
        rr = q0 + i
        ws[f"H{rr}"] = f'=IF(MIN(V_OI_Jaar)+{i}>MAX(V_OI_Jaar),"",MIN(V_OI_Jaar)+{i})'
        for q in range(1, 5):
            c = col(8 + q)
            ws[f"{c}{rr}"] = f'=IF($H{rr}="","",SUMIFS(V_PEWOmz,V_OI_Jaar,$H{rr},V_OI_Kw,{q},V_Type,Sel_TypeCrit,{crit})/P_Factor)'
        ws[f"M{rr}"] = f'=IF($H{rr}="","",SUM(I{rr}:L{rr}))'
        ws[f"N{rr}"] = f'=IF($H{rr}="","",COUNTIFS(V_OI_Jaar,$H{rr},V_Type,Sel_TypeCrit,{crit}))'
        for c in "IJKLMN":
            ws[f"{c}{rr}"].number_format = FMT_EUR if c != "N" else FMT_INT
            ws[f"{c}{rr}"].border = Border(bottom=thin)
        ws[f"M{rr}"].font = font(9, True)
    ws.conditional_formatting.add(f"I{q0}:L{q0+OUT_YEARS-1}", FormulaRule(
        formula=[f'AND($H{q0}<>"",DATE($H{q0},(COLUMN()-COLUMN($I{q0}))*3+1,1)>P_Peildatum)'], font=Font(color="8A6A00", italic=True)))
    put(ws, f"H{q0+OUT_YEARS}", "Elk jaar met orderintake in de bron (max. 8). Cursief = kwartaal na peildatum.", color=GREY, italic=True, size=8)
    add_name(wb, "PPT_OIQ", f"Outlook!$H${q0-1}:$N${q0+OUT_YEARS-1}")

    # lijst
    section(ws, OI0 - 2, 2, "Projecten op orderintakedatum (nieuwste eerst)", 10)
    header(ws, OI0 - 1, 2, ["Jaar", "Kwartaal", "Project", "Type", "Orderintake", "PEW-omzet", "PEW-BOR", "PEW-BOR %", "Status"], height=24)
    for k in range(1, V1 - V0 + 2):
        rr = OI0 + k - 1
        if rr > OI1:
            break
        ws[f"M{rr}"] = f"=IFERROR(MATCH(LARGE(V_OI_Key,{k}),V_OI_Key,0),0)"
        helper_style(ws[f"M{rr}"])
        ix = f"$M{rr}"
        g = lambda n: f'IF({ix}=0,"",INDEX(V_{n},{ix}))'  # noqa: E731
        ws[f"B{rr}"] = "=" + g("OI_Jaar")
        ws[f"C{rr}"] = f'=IF({ix}=0,"","Q"&INDEX(V_OI_Kw,{ix}))'
        ws[f"D{rr}"] = "=" + g("Naam")
        ws[f"E{rr}"] = "=" + g("Type")
        ws[f"F{rr}"] = "=" + g("OI_Datum")
        ws[f"G{rr}"] = f'=IF({ix}=0,"",INDEX(V_PEWOmz,{ix})/P_Factor)'
        ws[f"H{rr}"] = f'=IF({ix}=0,"",INDEX(V_PEWBor,{ix})/P_Factor)'
        ws[f"I{rr}"] = f'=IF(OR({ix}=0,G{rr}=0),"",H{rr}/G{rr})'
        ws[f"J{rr}"] = f'=IF({ix}=0,"",IF(F{rr}<=P_Peildatum,"gerealiseerd","verwacht"))'
        for c in "BCDEFGHIJ":
            cell = ws[f"{c}{rr}"]
            cell.font = font(9)
            cell.border = Border(bottom=thin)
        ws[f"F{rr}"].number_format = FMT_DATE
        for c in "GH":
            ws[f"{c}{rr}"].number_format = FMT_EUR
        ws[f"I{rr}"].number_format = FMT_PCT
    ws.conditional_formatting.add(f"J{OI0}:J{OI1}", FormulaRule(formula=[f'J{OI0}="verwacht"'], font=Font(color="8A6A00", italic=True)))
    ws["M3"] = "=COUNT(V_OI_Key)"
    helper_style(ws["M3"])
    add_name(wb, "OI_N", "Outlook!$M$3")
    add_name(wb, "PPT_OI", f"Outlook!$B${OI0-1}:$J${OI1}")
    widths(ws, {"B": 28, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 11, "L": 12, "M": 11, "N": 9})
    for c in gc + [col(OUT_CH)]:
        ws.column_dimensions[c].width = 9
    ws.freeze_panes = "C6"
    landscape(ws, f"A1:N{OI0 + 20}")
    return rows


# ---------------------------------------------------------------------------------------------
# PPT: teksten voor de presentatie (kernboodschappen, aandachtspunten)
# Dias: per dia een opgemaakt blok (18 kolommen x 20 regels); de macro plakt elk blok als afbeelding.
# ---------------------------------------------------------------------------------------------
def build_ppt(wb, ws):
    base_sheet(ws)
    put(ws, "A1", "Teksten voor de knop 'Maak PowerPoint' (verborgen tabblad). De dia-inhoud zelf staat op het verborgen tabblad Dias.", bold=True)
    S = lambda m, row="Sel": f"CALC_Versies!${CC[m]}${CR[row]}"  # noqa: E731
    o, b, ob = S("F_OmzFY"), S("F_BorFY"), S("F_OmzFY", "Budget")
    rood = '"verkoop "&COUNTIF(SL_Verkoop,1)&" · vergunning "&COUNTIF(SL_Vergunning,1)&" · netcongestie "&COUNTIF(SL_Net,1)'
    kv = [
        ("PPT_Titel", '="Prognoserapportage "&P_Regio'),
        ("PPT_Sub", '=Sel_Versie&" "&P_VJ&"  ·  doeljaar "&Sel_Doeljaar&"  ·  peildatum "&P_PeilTxt'),
        ("PPT_Footer", '=P_Regio&"  ·  "&Sel_Versie&" "&P_VJ&"  ·  peildatum "&P_PeilTxt'),
        ("PPT_Msg1", f'=IF(ISNUMBER({o}),"Omzet FY "&Sel_Doeljaar&": € "&FIXED({o}/1000000,1)&" mln"&IF(ISNUMBER({ob}),", "&IF({o}>={ob},"+","")&FIXED(({o}-{ob})/1000000,1)&" mln t.o.v. budget","")&IF(AND(ISNUMBER({b}),{o}<>0),"  ·  BOR "&FIXED({b}/{o}*100,1)&"%",""),"Geen vastgelegde of concept-prognose voor de selectie.")'),
        ("PPT_Msg2", '="Bedragen in "&P_Eenheid&"  ·  massief = vastgelegd/realisatie, balkjes = afwijking t.o.v. budget"'),
        ("PPT_Msg3", f'="Rode stoplichten: "&{rood}&"  ·  "&Projecten!$R$6'),
        ("PPT_Msg4", '="Realisatie t/m peildatum: "&IF(ISNUMBER(Dashboard!$H$8),"€ "&FIXED(Dashboard!$H$8,1)&" mln = ","")&Dashboard!$H$9'),
        ("PPT_Msg5", '="Omzet "&P_RJ&"–"&(P_RJ+6)&" volgens de upload van "&P_PeilTxt&", naast het vastgelegde budget"'),
        ("PPT_Msg6", '="Orderintake "&P_RJ&" t/m peildatum: € "&FIXED(Dashboard!$Q$8,1)&" mln  ·  "&Dashboard!$Q$9'),
        ("PPT_Msg7", '="Verkocht "&P_RJ&" t/m peildatum: "&Woningen!$H$5&" woningen  ·  jaarverwachting "&Woningen!$K$5&"  ·  etalage "&Woningen!$Q$5'),
        ("PPT_Control", '=Dashboard!$P$5'),
        ("PPT_ProjN", '=COUNTIF(Map_WBS,"?*")'),
        ("PPT_Msg8", '="Prognoseverloop en brug naar de geselecteerde versie  ·  groen = hoger, rood = lager"'),
    ]
    for i, (n, f) in enumerate(kv):
        r = 3 + i
        put(ws, f"A{r}", n, color=GREY)
        ws[f"B{r}"] = f
        add_name(wb, n, f"PPT!$B${r}")
    put(ws, "D20", "Aandachtspunten", bold=True)
    bullets = [
        '="Controle: "&Dashboard!$P$5',
        '=IF(COUNT(SL_RoodNr)=0,"Geen projecten met een rood stoplicht.","Rood stoplicht: "&' + '&'.join(
            f'IFERROR(IF({k}>1,", ","")&INDEX(SL_Project,MATCH({k},SL_RoodNr,0)),"")' for k in range(1, 7)) + ')',
        '=IF(SUM(V_Verlies)=0,"Geen verliesprojecten in de upload.","Verliesprojecten (BOR < 0): "&' + '&'.join(
            f'IFERROR(IF({k}>1,", ","")&INDEX(V_Naam,MATCH({k},V_VerliesNr,0)),"")' for k in range(1, 7)) + ')',
        '=IF(COUNTIF(RDFO!$J$7:$J$%d,"n.b.")=0,"Alle RDFO-posten zijn gevuld.",COUNTIF(RDFO!$J$7:$J$%d,"n.b.")&" RDFO-posten zijn n.b. (aanvullende invoer ontbreekt).")' % (RDFO_ROWS["ebit"], RDFO_ROWS["ebit"]),
        '=IF(P_Demo="Ja","Let op: rapport bevat fictieve demowaarden (instelling demowaarden = Ja).","")',
        '=IF(Sel_Bron="Concept","Geselecteerde versie is nog niet vastgelegd (concept uit upload).","")',
    ]
    for i, f in enumerate(bullets):
        ws[f"D{21+i}"] = f
    add_name(wb, "PPT_Bullets", f"PPT!$D$21:$D${21+len(bullets)-1}")


DIA_BLOCKS = {}
DIA_CHARTS = []      # volgorde = ChartObjects-index op Dias (1-based)
CH0 = 400            # eerste regel van de grafiekblokken op Dias
CH_ROWS = 24


def build_dias(wb, ws):
    """Tabellen (zonder samengevoegde cellen), KPI-lijsten en grafieken voor de PowerPoint.
    Boven elke tabel staat een rij met kolomgewichten: 0 = kolom weglaten (bijv. lege toelichting of jaar zonder cijfers)."""
    base_sheet(ws, zoom=100)
    put(ws, "A1", "Inhoud voor de PowerPoint (verborgen). Tabellen worden in PowerPoint als bewerkbare tabel opgebouwd; grafieken als afbeelding. "
        "Rij boven elke tabel = kolomgewicht (0 = weglaten).", color=GREY, size=8)
    ws.column_dimensions["A"].width = 14
    for i in range(2, 20):
        ws.column_dimensions[col(i)].width = 11
    ws.column_dimensions["B"].width = 30

    def cell(r, c, v, fmt=None, bold=False, color="1A2233", fillc=None, align="right", italic=False, size=10):
        x = ws.cell(row=r, column=c, value=v)
        x.font = font(size, bold, color, italic)
        if fmt:
            x.number_format = fmt
        if fillc:
            x.fill = fill(fillc)
        x.alignment = Alignment(horizontal=align, vertical="center")
        return x

    def table_head(r, labels, weights, name, nrows):
        put(ws, (r - 1, 1), "gewicht", color=GREY, size=7)
        for j, (lab, wgt) in enumerate(zip(labels, weights)):
            ws.cell(row=r - 1, column=2 + j, value=wgt).font = font(7, False, GREY)
            c = cell(r, 2 + j, lab, bold=True, color=WHITE, fillc=NAVY, align="left" if j == 0 else "center", size=9)
            c.alignment = Alignment(horizontal="left" if j == 0 else "center", vertical="center", wrap_text=True)
        put(ws, (r, 1), name, color=GREY, size=7)
        add_name(wb, name, f"Dias!$B${r}:${col(1+len(labels))}${r+nrows}")
        DIA_BLOCKS[name] = r

    # ---- KPI-lijsten (label, waarde, toelichting)
    r = 4
    put(ws, (r - 1, 1), "K_DASH", color=GREY, size=7)
    for k, sc in enumerate("BEHKNQ"):
        cell(r + k, 2, f"=Dashboard!{sc}7", align="left")
        cell(r + k, 3, f"=Dashboard!{sc}8", MLN if k != 4 else '#,##0" st."', bold=True)
        cell(r + k, 4, f"=Dashboard!{sc}9", align="left", color=GREY, size=8)
    add_name(wb, "K_DASH", f"Dias!$B${r}:$D${r+5}")
    r = 12
    put(ws, (r - 1, 1), "K_WON", color=GREY, size=7)
    for k, sc in enumerate("BEHKNQ"):
        cell(r + k, 2, f"=Woningen!{sc}4", align="left")
        cell(r + k, 3, f"=Woningen!{sc}5", '#,##0" st."', bold=True)
        cell(r + k, 4, f"=Woningen!{sc}6", align="left", color=GREY, size=8)
    add_name(wb, "K_WON", f"Dias!$B${r}:$D${r+5}")

    # ---- RDFO
    t = 22
    src_cols = list("CDEFGHJKLMNPQ")
    labels = ["Rubriek"] + [f"=RDFO!{c}6" for c in src_cols]
    weights = [3.2, 1, 1, 1, 1, 1, 1.1, 1.1, 1.1, 1, 1, 1, 0.8, 0.8]
    table_head(t, labels, weights, "T_RDFO", 17)
    for j, c in enumerate(src_cols[:6]):
        put(ws, (t - 2, 3 + j), f"=RDFO!{c}5", color=GREY, size=7)
    keys = ["omz_ops", "omz_dln", "taak_omz", "int_elim", "omzet_old", "ifrs_elim", "omzet_ifrs", "bor_ops", "bor_dln",
            "bor_tot", "ak", "fees", "overig", "taak_ebit", "ebit", "bor_pct", "ebit_pct"]
    subs = {"omzet_old", "omzet_ifrs", "bor_tot", "ebit"}
    for i, k in enumerate(keys):
        r = t + 1 + i
        sr = RDFO_ROWS[k]
        bold = k in subs or k.endswith("pct")
        cell(r, 2, f"=RDFO!B{sr}", bold=bold, align="left", fillc=LIGHT if k in subs else None)
        for j, sc in enumerate(src_cols):
            c = 3 + j
            f = f'=IF(RDFO!{sc}{sr}="","",RDFO!{sc}{sr})'
            if sc in "PQ":
                cell(r, c, f, color=BAD if sc == "P" else GOOD, align="right" if sc == "P" else "left", size=9)
            else:
                fmt = FMT_PCT if k.endswith("pct") else FMT_EUR
                if k.endswith("pct") and sc in "MN":
                    fmt = '+0.0%;[Red]-0.0%;"–"'
                cell(r, c, f, fmt, bold, "3B5B33" if sc == "K" else "1A2233", LIGHT if k in subs else None, italic=sc == "K")
    ws.conditional_formatting.add(f"C{t+1}:H{t+17}", FormulaRule(formula=[f"C${t-2}=Sel_Versie"], fill=fill(YELLOW_LT)))
    ws.conditional_formatting.add(f"M{t+1}:N{t+17}", FormulaRule(formula=[f"AND(ISNUMBER(M{t+1}),M{t+1}<0)"], font=Font(color=BAD)))
    ws.conditional_formatting.add(f"M{t+1}:N{t+17}", FormulaRule(formula=[f"AND(ISNUMBER(M{t+1}),M{t+1}>0)"], font=Font(color="3B7A1E")))

    # ---- projecten en stoplicht (18 per dia)
    per = 18
    toel_w = '=IF(COUNTIF(Stoplicht!$H$8:$H$107,"?*")+COUNTIF(Projecten!$S$10:$S$109,"?*")>0,3,0)'
    plabels = ["Project", "Type", '="Omzet FY "&Sel_Versie', "Realisatie", "To go", "BOR %", "Δ vs budget", "Orderintake",
               "Verkoop", "Vergunning", "Netcongestie", "Toelichting"]
    pweights = [2.6, 1, 1, 1, 1, 0.8, 1, 1, 0.8, 0.9, 0.9, toel_w]
    for blk in range(3):
        t = 44 + blk * (per + 4)
        table_head(t, plabels, pweights, f"T_PROJ{blk+1}", per)
        for i in range(per):
            idx = blk * per + i
            r = t + 1 + i
            pr, slr = P0 + idx, SL0 + idx
            refs = [(2, f"=Projecten!B{pr}", None, "left"), (3, f"=Projecten!C{pr}", None, "left"),
                    (4, f"=Projecten!K{pr}", FMT_EUR, "right"), (5, f"=Projecten!L{pr}", FMT_EUR, "right"),
                    (6, f"=Projecten!M{pr}", FMT_EUR, "right"), (7, f"=Projecten!O{pr}", FMT_PCT, "right"),
                    (8, f"=Projecten!P{pr}", FMT_EUR, "right"), (9, f'=IF(Projecten!T{pr}="","",Projecten!T{pr})', "mm-yyyy", "center"),
                    (10, f"=Stoplicht!I{slr}", SL_FMT, "center"), (11, f"=Stoplicht!J{slr}", SL_FMT, "center"),
                    (12, f"=Stoplicht!K{slr}", SL_FMT, "center"),
                    (13, f'=IF(Stoplicht!H{slr}="",IF(Projecten!S{pr}="","",Projecten!S{pr}),Stoplicht!H{slr})', None, "left")]
            for c, f, fmt, al in refs:
                cell(r, c, f, fmt, False, GREY if c == 13 else "1A2233", None, al, italic=c in (5, 13), size=13 if fmt == SL_FMT else 9)
        stoplicht_cf(ws, f"J{t+1}:L{t+per}")
        ws.conditional_formatting.add(f"H{t+1}:H{t+per}", FormulaRule(formula=[f"AND(ISNUMBER(H{t+1}),H{t+1}<0)"], font=Font(color=BAD)))
        ws.conditional_formatting.add(f"H{t+1}:H{t+per}", FormulaRule(formula=[f"AND(ISNUMBER(H{t+1}),H{t+1}>0)"], font=Font(color="3B7A1E")))

    # ---- outlook-tabel (jaarkolommen zonder cijfers krijgen gewicht 0)
    t = 115
    ows = wb["Outlook"]
    rows = OUT_CTX["rows"]
    yc = [col(3 + i) for i in range(OUT_YEARS)] + [col(3 + OUT_YEARS), col(4 + OUT_YEARS)]
    olabels = ["Kengetal"] + [f'=SUBSTITUTE(Outlook!{c}5,CHAR(10)," ")' for c in yc]
    oweights = [3] + [f"=IF({i+1}<=OUT_N,1,0)" for i in range(OUT_YEARS)] + ["=IF(OUT_Later=1,1,0)", 1.1]
    keys_o = [k for k in ("omz_ops", "omz_dln", "omz", "omz_ac", "bor", "borp", "bud", "dbud", "won", "won_ac", "wbud",
                          "oi_n", "oi", "oi_ac", "oi_bor", "oi_bud")]
    table_head(t, olabels, oweights, "T_OUT", len(keys_o))
    for i, k in enumerate(keys_o):
        r = t + 1 + i
        sr = rows[k]
        lab = ows[f"B{sr}"].value
        sub = str(lab).startswith("  ")
        bold = k in ("omz", "won", "oi")
        cell(r, 2, f"=TRIM(Outlook!B{sr})", bold=bold, color=GREY if sub else "1A2233", align="left", italic=sub,
             fillc=LIGHT if bold else None)
        for j, c in enumerate(yc):
            fmt = ows[f"{c}{sr}"].number_format
            cell(r, 3 + j, f'=IF(Outlook!{c}{sr}="","",Outlook!{c}{sr})', fmt, bold, GREY if sub else "1A2233",
                 LIGHT if bold else ("EEF3EA" if j == 0 else None), italic=sub)
    ws.conditional_formatting.add(f"C{t+1+keys_o.index('dbud')}:L{t+1+keys_o.index('dbud')}",
                                  FormulaRule(formula=[f"AND(ISNUMBER(C{t+1+keys_o.index('dbud')}),C{t+1+keys_o.index('dbud')}<0)"], font=Font(color=BAD, bold=True)))

    # ---- orderintake: jaar x kwartaal en recentste projecten
    t = 140
    table_head(t, ["Jaar", "Q1", "Q2", "Q3", "Q4", "Totaal", "#"], [1, 1, 1, 1, 1, 1.2, 0.6], "T_OIQ", OUT_YEARS)
    for i in range(OUT_YEARS):
        r = t + 1 + i
        for j, sc in enumerate("HIJKLMN"):
            cell(r, 2 + j, f'=IF(Outlook!{sc}{46+i}="","",Outlook!{sc}{46+i})', None if j == 0 else (FMT_INT if sc == "N" else FMT_EUR),
                 sc in "HM", align="left" if j == 0 else "right")
    t = 152
    table_head(t, ["Jaar", "Kw.", "Project", "Type", "Datum", "PEW-omzet", "BOR %", "Status"], [0.7, 0.6, 2.6, 1.1, 1, 1.1, 0.8, 1.1], "T_OI", 10)
    for i in range(10):
        r = t + 1 + i
        orr = OI0 + i
        for j, (sc, fmt, al) in enumerate([("B", None, "left"), ("C", None, "left"), ("D", None, "left"), ("E", None, "left"),
                                            ("F", "dd-mm-yy", "center"), ("G", FMT_EUR, "right"), ("I", FMT_PCT, "right"), ("J", None, "left")]):
            cell(r, 2 + j, f'=IF(Outlook!{sc}{orr}="","",Outlook!{sc}{orr})', fmt, False, "1A2233", None, al, size=9)
    ws.conditional_formatting.add(f"I{t+1}:I{t+10}", FormulaRule(formula=[f'I{t+1}="verwacht"'], font=Font(color="8A6A00", italic=True)))

    # ---- grafiekblokken (elk op een eigen pagina; volgorde = index voor de macro)
    def chart_block(name, builder, w, h):
        k = len(DIA_CHARTS)
        r = CH0 + k * CH_ROWS
        put(ws, (r, 1), name, color=GREY, size=7)
        for rr in range(r, r + CH_ROWS):
            ws.row_dimensions[rr].height = 20   # 24 x 20 pt = 480 pt: ruimer dan de hoogste grafiek
        builder(f"B{r}", w, h)
        DIA_CHARTS.append(name)
        return r

    pt = lambda x: round(x * 0.03528, 2)  # noqa: E731
    T0 = DASH["T0"]
    dws = wb["Dashboard"]
    wws = wb["Woningen"]
    q0, ql, v0 = WON["q0"], WON["qlast"], WON["v0"]

    def won_q(anchor, w, h):
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
        ch.add_data(Reference(wws, min_col=3, max_col=4, min_row=q0 - 1, max_row=ql), titles_from_data=True)
        ch.set_categories(Reference(wws, min_col=2, min_row=q0, max_row=ql))
        style_chart(ch, [], numfmt="#,##0", w=w, h=h, ymin=0)
        ch.series[0].graphicalProperties = solid(NAVY)
        ch.series[1].graphicalProperties = hatched(NAVY)
        ws.add_chart(ch, anchor)

    chart_block("Prognoseverloop (waterval)", lambda a, w, h: waterfall(ws, a, DASH["catsA"], DASH["dataA"], w=w, h=h), pt(436), pt(296))
    chart_block("Outlook omzet (klein)", lambda a, w, h: outlook_chart(ws, a, "omzet", w, h), pt(436), pt(296))
    chart_block("Brug vorige → geselecteerd", lambda a, w, h: waterfall(ws, a, DASH["catsC"], DASH["dataC"], w=w, h=h, numfmt="#,##0.00", kind="bar"), pt(436), pt(376))
    chart_block("Realisatie vs to go", lambda a, w, h: ibcs_columns(ws, a, Reference(dws, min_col=3, min_row=T0 + 1, max_row=T0 + 10),
                Reference(dws, min_col=4, min_row=T0, max_row=T0 + 10), Reference(dws, min_col=5, min_row=T0, max_row=T0 + 10),
                kind="bar", w=w, h=h), pt(436), pt(376))
    chart_block("Outlook omzet", lambda a, w, h: outlook_chart(ws, a, "omzet", w, h), pt(290), pt(376))
    chart_block("Outlook woningen", lambda a, w, h: outlook_chart(ws, a, "won", w, h), pt(290), pt(376))
    chart_block("Outlook orderintake", lambda a, w, h: outlook_chart(ws, a, "oi", w, h), pt(290), pt(376))
    chart_block("Woningen per kwartaal", won_q, pt(436), pt(296))
    chart_block("Woningen per versie", lambda a, w, h: bar(ws, Reference(wws, min_col=17, min_row=v0 - 1, max_row=v0 + 5),
                Reference(wws, min_col=3, min_row=v0, max_row=v0 + 5), a, [NAVY], numfmt="#,##0", w=w, h=h, labels=True, legend=False,
                ymin=0, point_colors={5: NAVY_LT}), pt(436), pt(296))
    # vaste rijhoogtes (geen automatische aanpassing), anders verschuiven de grafieken t.o.v. de cellen
    for rr in range(1, CH0):
        if ws.row_dimensions[rr].height is None:
            ws.row_dimensions[rr].height = 15
    from openpyxl.worksheet.pagebreak import Break
    for k in range(1, len(DIA_CHARTS)):
        ws.row_breaks.append(Break(id=CH0 + k * CH_ROWS - 1))
    landscape(ws, ",".join(f"B{CH0 + k * CH_ROWS}:L{CH0 + (k + 1) * CH_ROWS - 1}" for k in range(len(DIA_CHARTS))))
    ws.sheet_properties.pageSetUpPr.fitToPage = False   # elk grafiekblok een eigen afdrukbereik = eigen pagina
    ws.page_setup.scale = 100
    ws.page_margins.top = ws.page_margins.bottom = 0.3


def build_ppt_spec(wb, ws):
    """Dia-specificatie: per dia welke elementen waar komen. De macro en de voorbeeldpresentatie lezen deze tabel."""
    r0 = 40
    put(ws, f"H{r0-2}", "Dia-specificatie (positie in punten op een dia van 960 x 540)", bold=True)
    hdr = ["Dia", "Titel", "Kernboodschap", "Soort", "Ref", "X", "Y", "B", "H", "Kop", "Actief"]
    for j, h in enumerate(hdr):
        put(ws, (r0 - 1, 8 + j), h, bold=True)
    ci = lambda n: DIA_CHARTS.index(n) + 1  # noqa: E731
    spec = [
        (2, "Kerncijfers", "=PPT_Msg1", "kpi", "K_DASH", 36, 96, 888, 74, "", 1),
        (2, "", "", "grafiek", ci("Prognoseverloop (waterval)"), 36, 200, 436, 296, "=Dashboard!B12", 1),
        (2, "", "", "grafiek", ci("Outlook omzet (klein)"), 488, 200, 436, 296, "=Dashboard!K30", 1),
        (3, "Omzet en resultaat", "=PPT_Msg2", "tabel", "T_RDFO", 36, 96, 888, 0, "", 1),
        (4, "Projecten en stoplicht", "=PPT_Msg3", "tabel", "T_PROJ1", 36, 96, 888, 0, "", 1),
        (5, "Projecten en stoplicht (vervolg)", "=PPT_Msg3", "tabel", "T_PROJ2", 36, 96, 888, 0, "", "=IF(PPT_ProjN>18,1,0)"),
        (6, "Projecten en stoplicht (vervolg)", "=PPT_Msg3", "tabel", "T_PROJ3", 36, 96, 888, 0, "", "=IF(PPT_ProjN>36,1,0)"),
        (7, "Prognoseverloop en realisatie", "=PPT_Msg8", "grafiek", ci("Brug vorige → geselecteerd"), 36, 120, 436, 376, "=Dashboard!B30", 1),
        (7, "", "", "grafiek", ci("Realisatie vs to go"), 488, 120, 436, 376, "=Dashboard!K12", 1),
        (8, "Outlook", "=PPT_Msg5", "tabel", "T_OUT", 36, 96, 888, 0, "", 1),
        (9, "Outlook in beeld", "=PPT_Msg5", "grafiek", ci("Outlook omzet"), 36, 120, 290, 376, "Omzet per jaar (EUR mln)", 1),
        (9, "", "", "grafiek", ci("Outlook woningen"), 335, 120, 290, 376, "Woningverkopen per jaar", 1),
        (9, "", "", "grafiek", ci("Outlook orderintake"), 634, 120, 290, 376, "Orderintake per jaar (EUR mln)", 1),
        (10, "Orderintake", "=PPT_Msg6", "tabel", "T_OIQ", 36, 120, 380, 0, "Per jaar en kwartaal (PEW-omzet)", 1),
        (10, "", "", "tabel", "T_OI", 440, 120, 484, 0, "Recentste orderintake", 1),
        (11, "Woningverkopen", "=PPT_Msg7", "kpi", "K_WON", 36, 96, 888, 74, "", 1),
        (11, "", "", "grafiek", ci("Woningen per kwartaal"), 36, 200, 436, 296, f"=Woningen!{WON['qtitle']}", 1),
        (11, "", "", "grafiek", ci("Woningen per versie"), 488, 200, 436, 296, f"=Woningen!{WON['vtitle']}", 1),
        (12, "Aandachtspunten", "", "opsomming", "PPT_Bullets", 36, 100, 888, 390, "", 1),
    ]
    for i, row in enumerate(spec):
        for j, v in enumerate(row):
            ws.cell(row=r0 + i, column=8 + j, value=v)
    add_name(wb, "PPT_Spec", f"PPT!$H${r0}:$R${r0+len(spec)-1}")
    # outlook-grafieken die de macro RDFO_OutlookJaren op het aantal jaren met cijfers zet: blad|index|basisrij
    g0 = OUT_CTX["g0"]
    base = {"omzet": g0, "won": g0 + 5, "oi": g0 + 10}
    lijst = [("Outlook", 1, base["omzet"]), ("Outlook", 2, base["won"]), ("Outlook", 3, base["oi"]), ("Dashboard", 4, base["omzet"]),
             ("Dias", ci("Outlook omzet (klein)"), base["omzet"]), ("Dias", ci("Outlook omzet"), base["omzet"]),
             ("Dias", ci("Outlook woningen"), base["won"]), ("Dias", ci("Outlook orderintake"), base["oi"])]
    ws["H36"] = ";".join(f"{a}|{b}|{c}" for a, b, c in lijst)
    add_name(wb, "OUT_Grafieken", "PPT!$H$36")
    ws["H35"] = len(DIA_CHARTS)
    add_name(wb, "PPT_ChartN", "PPT!$H$35")


# ---------------------------------------------------------------------------------------------
# Controle
# ---------------------------------------------------------------------------------------------


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
    frozen = "Concept!$N$3"
    checks = [
        ("Bron", "Regels ZPMW03 geplakt", f'=COUNTIF({zc("WBS")},"?*")', '="verwacht "&P_RegelsZ', 'IF({v}=P_RegelsZ,"OK","FOUT")',
         "Vergelijk met aantal regels in het exportbestand (Invoer › Instellingen; wordt bij inlezen automatisch gevuld)."),
        ("Bron", "Regels Woningetalage geplakt", f'=COUNTIF({ec("Stufe3")},"?*")', '="verwacht "&P_RegelsE', 'IF({v}=P_RegelsE,"OK","FOUT")', "Idem P_RegelsE."),
        ("Bron", "ZPMW03-regels zonder WBS-element", f'=COUNTA(Bron_ZPMW03!$A$2:$A${N_SRC+1})-COUNTIF({zc("WBS")},"?*")', "0", 'IF({v}=0,"OK","FOUT")',
         "Regels zonder sleutel kunnen niet gekoppeld worden en vallen anders stil weg."),
        ("Bron", "Woningetalage-regels zonder Stufe 3", f'=COUNTA(Bron_Woningetalage!$A$2:$A${N_SRC+1})-COUNTIF({ec("Stufe3")},"?*")', "0", 'IF({v}=0,"OK","FOUT")', "Idem."),
        ("Bron", f"Regels voorbij capaciteit ({N_SRC} per bron)", f'=COUNTA(Bron_ZPMW03!$A${N_SRC+2}:$A$20000)+COUNTA(Bron_Woningetalage!$A${N_SRC+2}:$A$20000)', "0",
         'IF({v}=0,"OK","FOUT")', "Formulebereiken STG_* uitbreiden (zie Handleiding)."),
        ("Kolommen", "ZPMW03: ontbrekende kolommen", f'=COUNTIF(STG_ZPMW03!$B$4:${zlast}$4,"ONTBREEKT")', "0", 'IF({v}=0,"OK","FOUT")',
         "Zie Handleiding › bijlage Veldwoordenboek. Aliassen Bedrijfsnaam / Naam maatschappij worden beide ondersteund."),
        ("Kolommen", "Woningetalage: ontbrekende kolommen", f'=COUNTIF(STG_Woningetalage!$B$4:${elast}$4,"ONTBREEKT")', "0", 'IF({v}=0,"OK","FOUT")',
         "Zie Handleiding › bijlage Veldwoordenboek."),
        ("Kolommen", "Woningetalage: jaarheaders uit mapping niet gevonden", '=COUNTIF(YM_Gevonden,"NEE")', "0",
         'IF({v}=0,"OK","WAARSCHUWING")', "Invoer › 6. Jaarkolommen Woningetalage bijwerken."),
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
        ("Bedragen", "OP.BOB + Kosten ≠ OP.BOR (LJ, RJ, RJ+1)", f'=COUNTIF({zc("C_BobKostBor")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Opbrengsten en kosten sluiten niet aan op het resultaat. Vaak: kosten met ander teken in de export. Welke regels: STG_ZPMW03, kolom C_BobKostBor = 1."),
        ("Bedragen", "OP.BOB RJ bevat kwartalen vóór peildatum (geen restant?)",
         f'=SUMPRODUCT((P_PeilMaand>=3)*((COUNTIF({zc("BOB_RJ_Q1")},"<0")+COUNTIF({zc("BOB_RJ_Q1")},">0"))))'
         f'+(P_PeilMaand>=6)*(COUNTIF({zc("BOB_RJ_Q2")},"<0")+COUNTIF({zc("BOB_RJ_Q2")},">0"))'
         f'+(P_PeilMaand>=9)*(COUNTIF({zc("BOB_RJ_Q3")},"<0")+COUNTIF({zc("BOB_RJ_Q3")},">0"))', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Demo-aanname: OP.BOB RJ = restant na peildatum. Afwijking → definitie in productie controleren (dubbeltelling met LJ!)."),
        ("Bedragen", "Portefeuille BOB ≠ som OP.BOB RJ..RJ+8 + onverdeeld", f'=COUNTIF({zc("C_Portefeuille")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Portefeuille sluit niet aan op de jaarverdeling. Rapportage gebruikt de jaarvelden. Regels: STG_ZPMW03, kolom C_Portefeuille = 1."),
        ("Bedragen", "PEW opbrengst ≠ BOB VJ + BOB LJ + portefeuille", f'=COUNTIF({zc("C_PEW")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', "PEW wijkt af van gerealiseerd + portefeuille (bijv. PEW nog niet bijgewerkt). Regels: STG_ZPMW03, kolom C_PEW = 1."),
        ("Woningen", "Aantal ≠ vorig jaar + 12 mnd + volgend jaar + latere jaren", f'=COUNTIF({ec("C_Aantal")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Verkopen per jaar tellen niet op tot het aantal woningen. Regels: STG_Woningetalage, kolom C_Aantal = 1."),
        ("Woningen", "Etalage ≠ aantal − cumulatief actual", f'=COUNTIF({ec("C_Etalage")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Etalage (nog te verkopen) klopt niet met aantal en verkocht. Regels: STG_Woningetalage, kolom C_Etalage = 1."),
        ("Woningen", "Cumulatief ≠ t/m vorig jaar + huidig jaar actual", f'=COUNTIF({ec("C_Cum")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Cumulatief verkocht sluit niet aan. Regels: STG_Woningetalage, kolom C_Cum = 1."),
        ("Woningen", "Huidig jr Totaal ≠ som hjr-maanden t/m peilmaand", f'=COUNTIF({ec("C_HJ")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Controleert ook de peilmaand."),
        ("Woningen", "Beleggers + Particulier ≠ Totaal", f'=COUNTIF({ec("C_BP")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Splitsing beleggers/particulier telt niet op. Regels: STG_Woningetalage, kolom C_BP = 1."),
        ("Woningen", "Huidig/Volgend jaar ≠ som maanden/kwartalen", f'=COUNTIF({ec("C_Jaren")},1)', "0", 'IF({v}=0,"OK","FOUT")', "Jaartotaal sluit niet aan op de maanden/kwartalen. Regels: STG_Woningetalage, kolom C_Jaren = 1."),
        ("Woningen", "Niet-geïnterpreteerde kolommen bevatten waarden", f'=COUNTIF({ec("C_NietGeinterp")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")',
         "Rond en negatieve jaarkoppen: eerst betekenis vastleggen in Invoer › 6."),
        ("Woningen", "Aantal woningen wijkt af van Mapping", f'=COUNTIF({ec("C_MapAantal")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Werk het aantal woningen bij in Invoer › 2 (rapportage gebruikt de Woningetalage)."),
        ("Woningen", "Woningen bij project met 'Woningen meetellen = Nee'", f'=COUNTIF({ec("C_Meetellen")},1)', "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Zet 'Woningen meetellen' op Ja in Invoer › 2, anders ontbreken deze woningen."),
        ("Weging", "Woningen niet dubbel gewogen: rapporttotaal = brontotaal", '=SUMIFS(V_W_Aantal,V_Meetellen,"Ja")', f'=SUM({ec("Aantal")})',
         'IF(ABS({v}-{n})<=P_Tol,"OK","FOUT")', '="Met aandeel nogmaals gewogen zou het "&FIXED(SUMPRODUCT(IFERROR(V_W_Aantal*V_Aandeel/100,0)),0)&" zijn — wordt bewust NIET gedaan."'),
        ("Stoplicht", "Stoplichtinvoer hoort bij een ander project (lijst verschoven?)", '=SUMPRODUCT((SL_Stempel<>"")*(SL_Stempel<>SL_WBS))', "0",
         'IF({v}=0,"OK","WAARSCHUWING")', "Projecten in Invoer › 2 zijn tussengevoegd of gesorteerd. Controleer de gele cellen op Stoplicht (stempel in verborgen kolom P)."),
        ("Stoplicht", "Projecten zonder ingevulde vergunning of netcongestie", '=COUNTIFS(SL_Project,"?*",SL_Vergunning,0)+COUNTIFS(SL_Project,"?*",SL_Net,0)', "info", '"INFO"', "Leeg rondje: nog geen beoordeling."),
        ("Verlies", "Verliesprojecten in upload (BOR FY RJ < 0)", "=SUM(V_Verlies)", "info", '"INFO"', "Zie lijst hieronder; verlies blijft negatief in alle rapporten."),
        ("Verlies", "RDFO-BOR upload = signed som projecten (geen ABS)", f'=IF(ISNUMBER(RDFO!$H${R("bor_tot")}),RDFO!$H${R("bor_tot")}*P_Factor,"n.b.")',
         f'={upFY("BorFY")}', 'IF(ISNUMBER({v}),IF(ABS({v}-{n})<=P_Tol,"OK","FOUT"),"INFO")', '="Met ABS zou het "&FIXED(SUMPRODUCT(ABS(N(+V_BorFY))),0)&" zijn."'),
        ("Versies", "Dubbele regels in Vastlegging (genegeerd)", "=SUM(L_Dubbel)", "0", 'IF({v}=0,"OK","WAARSCHUWING")', "Dubbel geplakt; telt niet dubbel. Documenteer in Importlog."),
        ("Versies", "Correctie (revisie > 0) zonder toelichting", '=COUNTIFS(L_Revisie,">0",L_Toelichting,"")', "0", 'IF({v}=0,"OK","FOUT")', "Toelichting is verplicht."),
        ("Versies", "Ongeldige status in Vastlegging", '=COUNTIF(L_WBS,"?*")-COUNTIFS(L_WBS,"?*",L_Status,"Vastgesteld")-COUNTIFS(L_WBS,"?*",L_Status,"Concept")', "0",
         'IF({v}=0,"OK","FOUT")', "Alleen Concept of Vastgesteld."),
        ("Versies", "Ongeldige versie in Vastlegging", '=COUNTIF(L_WBS,"?*")-' + "-".join(f'COUNTIFS(L_WBS,"?*",L_Versie,"{v}")' for v in VERSIES), "0",
         'IF({v}=0,"OK","FOUT")', "Alleen Budget/Q1/Q2/Q3/Q4."),
        ("Versies", "Concept-regels in Vastlegging (tellen niet mee)", '=COUNTIFS(L_WBS,"?*",L_Status,"Concept")', "info", '"INFO"', "Ter informatie; alleen Vastgesteld telt mee."),
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
        ws[f"E{r}"].number_format = 'General'
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
# Start en Handleiding
# ---------------------------------------------------------------------------------------------
# Knoppen (vormen met macro) — worden bij het verpakken als .xlsm toegevoegd. (blad, van-cel, tot-cel, tekst, macro, stijl)
BUTTONS = [
    ("Start", "B9", "B11", "1  Bronnen inlezen", "RDFO_BronnenInlezen", "geel"),
    ("Start", "C9", "C11", "2  Prognose vastleggen", "RDFO_PrognoseVastleggen", "geel"),
    ("Start", "D9", "D11", "3  Maak PowerPoint", "RDFO_MaakPowerPoint", "geel"),
    ("Start", "E9", "E11", "Toon / verberg rekenbladen", "RDFO_ToonRekenbladen", "blauw"),
    ("Dashboard", "R1", "T2", "Maak PowerPoint", "RDFO_MaakPowerPoint", "geel"),
    ("Historie", "P2", "R3", "Prognose vastleggen", "RDFO_PrognoseVastleggen", "geel"),
    ("Controle", "G2", "H3", "Bronnen opnieuw inlezen", "RDFO_BronnenInlezen", "geel"),
    ("Controle", "E2", "F3", "Nieuwe projecten toevoegen", "RDFO_ProjectenToevoegen", "blauw"),
    ("Start", "G18", "G19", "Nieuwe projecten toevoegen", "RDFO_ProjectenToevoegen", "blauw"),
    ("Start", "G22", "G23", "Bestand leegmaken", "RDFO_LeegMaken", "blauw"),
    ("Invoer", "J27", "L28", "Nieuwe projecten toevoegen", "RDFO_ProjectenToevoegen", "blauw"),
]


def build_start(wb, ws):
    base_sheet(ws, zoom=100)
    title(ws, None, None, 7)
    ws["B1"] = '="RDFO-prognoserapportage — "&P_Regio'
    ws["B1"].font = font(16, True, WHITE)
    ws.row_dimensions[1].height = 34
    ws["B2"] = '="Upload "&P_UploadVersie&" "&P_UploadVJ&" · peildatum "&P_PeilTxt&" · ingelezen "&IF(ISNUMBER(P_Importdatum),DAY(P_Importdatum)&"-"&MONTH(P_Importdatum)&"-"&YEAR(P_Importdatum),"")&" · fictieve demodata"'
    ws["B2"].font = font(9, False, GREY, True)
    widths(ws, {"A": 2, "B": 30, "C": 30, "D": 30, "E": 30, "F": 3, "G": 30})
    # status
    put(ws, "B4", "Status", bold=True, color=NAVY, size=11)
    ws["B5"] = '=IF(Controle!$D$4>0,"✖ "&Controle!$D$4&" fout(en) — oplossen",IF(Controle!$D$5>0,"▲ "&Controle!$D$5&" waarschuwing(en)","✔ geen fouten"))'
    ws["B5"].font = font(10, True, NAVY)
    ws["B5"].hyperlink = "#'Controle'!A1"
    ws.conditional_formatting.add("B5", FormulaRule(formula=["Controle!$D$4>0"], font=Font(color=RED, bold=True)))
    ws["C5"] = '="Vastgelegd: "&IF(COUNTIFS(L_Versiejaar,P_UploadVJ,L_Versie,P_UploadVersie,L_Status,"Vastgesteld")>0,P_UploadVersie&" "&P_UploadVJ&" ✔",P_UploadVersie&" "&P_UploadVJ&" nog niet")'
    ws["C5"].font = font(10, True, NAVY)
    ws["D5"] = '="Rode stoplichten: verkoop "&COUNTIF(SL_Verkoop,1)&" · vergunning "&COUNTIF(SL_Vergunning,1)&" · net "&COUNTIF(SL_Net,1)'
    ws["D5"].font = font(10, True, NAVY)
    ws["D5"].hyperlink = "#'Stoplicht'!A1"
    ws["E5"] = '="Fictieve demowaarden: "&P_Demo'
    ws["E5"].font = font(10, False, GREY)
    # knoppen
    put(ws, "B7", "Elk kwartaal", bold=True, color=NAVY, size=11)
    for c, t in zip("BCDE", ["Kies ZPMW03 en Woningetalage; peildatum en versie invullen.",
                              "Legt de jaarprognose vast in Historie (na controle).",
                              "Maakt de presentatie in huisstijl, naast dit bestand.",
                              "Laat de verborgen bron- en rekenbladen zien."]):
        put(ws, f"{c}12", t, color=GREY, size=8, wrap=True)
    for r in (9, 10, 11):
        ws.row_dimensions[r].height = 15
    ws.row_dimensions[12].height = 26
    # stappen
    put(ws, "B14", "Zo werkt het", bold=True, color=NAVY, size=11)
    for i, (t, wat, waar, let) in enumerate(DOC.STAPPEN):
        r = 15 + i
        c = ws.cell(row=r, column=2, value=f"{i+1}. {t}")
        c.font = font(10, True, NAVY)
        c.fill = fill(YELLOW_LT)
        ws.cell(row=r, column=3, value=wat).font = font(9)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        for cc in range(2, 6):
            ws.cell(row=r, column=cc).border = Border(bottom=thin)
    # tabbladen
    r = 22
    put(ws, f"B{r}", "Tabbladen", bold=True, color=NAVY, size=11)
    for i, (n, d, act) in enumerate(DOC.TABBLADEN):
        rr = r + 1 + i
        c = put(ws, f"B{rr}", n)
        c.hyperlink = f"#'{n}'!A1"
        c.font = Font(name=FONT, size=10, color="1F4E9A", underline="single")
        put(ws, f"C{rr}", d, color="333F55")
        ws.merge_cells(f"C{rr}:D{rr}")
        put(ws, f"E{rr}", act, color=GREY, italic=True, size=8)
        for cc in range(2, 6):
            ws.cell(row=rr, column=cc).border = Border(bottom=thin)
    # legenda
    put(ws, "G7", "Legenda", bold=True, color=NAVY, size=11)
    leg = [(YELLOW_LT, "Gele cel = invullen"), ("EEF3EA", "Groen cursief = realisatie (actual)"),
           (ORANGE_LT, "Oranje = subtotaal onvolledig"), (LIGHT, "Lichtgrijs = totaal / kerncijfer")]
    for i, (fc, t) in enumerate(leg):
        cell = ws[f"G{8+i}"]
        cell.value = t
        cell.fill = fill(fc)
        cell.font = font(9)
    put(ws, "G13", "n.b. = niet beschikbaar (geen nul)", color=GREY, size=8)
    put(ws, "G14", "n.v.t. = niet van toepassing", color=GREY, size=8)
    put(ws, "G17", "Beheer", bold=True, color=NAVY, size=11)
    put(ws, "G20", "Zet WBS-elementen uit ZPMW03 die nog niet in de projectlijst staan onderaan Invoer › 2.", color=GREY, size=8, wrap=True)
    put(ws, "G24", "Leeg format: wist bronnen, projectlijst, invoer, stoplicht en Historie. Instellingen blijven.", color=GREY, size=8, wrap=True)
    put(ws, "G15", "● groen  ● oranje  ● rood  ○ leeg", color=GREY, size=8)
    landscape(ws, "A1:G35")


def build_handleiding(ws, results=None):
    base_sheet(ws, zoom=100)
    title(ws, "Handleiding", "Kort en praktisch. De bijlagen onderaan zijn naslag.", 6)
    widths(ws, {"A": 2, "B": 6, "C": 30, "D": 62, "E": 34, "F": 40})
    r = 4

    def head(t):
        nonlocal r
        r += 1
        for c in range(2, 7):
            ws.cell(row=r, column=c).fill = fill(NAVY)
        put(ws, (r, 2), t, bold=True, color=WHITE, size=12, fillc=NAVY)
        ws.row_dimensions[r].height = 22
        r += 2

    def para(t, color="000000", bold=False, italic=False):
        nonlocal r
        put(ws, (r, 2), t, color=color, bold=bold, italic=italic, size=10)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 15 * max(1, len(t) // 150 + 1)
        r += 1

    def bullets(items, numbered=False):
        nonlocal r
        for i, t in enumerate(items):
            c = ws.cell(row=r, column=2, value=f"{i+1}" if numbered else "•")
            c.font = font(10, True, NAVY)
            c.alignment = Alignment(horizontal="center", vertical="top")
            put(ws, (r, 3), t, size=10)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
            ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * max(1, len(t) // 140 + 1)
            r += 1

    def table(hdrs, rows, cols=None):
        nonlocal r
        cols = cols or list(range(2, 2 + len(hdrs)))
        for h, c in zip(hdrs, cols):
            cell = ws.cell(row=r, column=c, value=h)
            cell.fill = fill(NAVY2)
            cell.font = font(9, True, WHITE)
        r += 1
        for row in rows:
            mx = 1
            for v, c in zip(row, cols):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = font(9)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = Border(bottom=thin)
                w = ws.column_dimensions[col(c)].width or 10
                mx = max(mx, len(str(v or "")) // max(1, int(w * 1.1)) + 1)
            ws.row_dimensions[r].height = 13 * mx
            r += 1

    head("Wat doet dit bestand?")
    para(DOC.INLEIDING)
    head("Elk kwartaal in 5 stappen")
    table(["", "Stap", "Wat doe je", "Waar", "Let op"], [(f"{i+1}", t, wat, waar, let) for i, (t, wat, waar, let) in enumerate(DOC.STAPPEN)],
          cols=[2, 3, 4, 5, 6])
    head("De tabbladen")
    table(["", "Tabblad", "Wat staat er", "Wat doe je er"], [("", n, d, a) for n, d, a in DOC.TABBLADEN], cols=[2, 3, 4, 5])
    head("Stoplicht invullen")
    bullets(DOC.STOPLICHT)
    head("Een vastgelegde prognose corrigeren")
    bullets(DOC.CORRECTIE, numbered=True)
    head("Spelregels")
    bullets(DOC.SPELREGELS)
    head("Problemen oplossen")
    table(["", "Probleem", "Oplossing"], [("?", p, o) for p, o in DOC.FAQ], cols=[2, 3, 4])
    head("Bijlage A — Waarom zo gebouwd")
    table(["", "Onderwerp", "Toelichting"], [("", k, v) for k, v in DOC.ONTWERPKEUZE], cols=[2, 3, 4])
    head("Bijlage B — Van bron naar rapport")
    table(["", "Rapportregel", "Berekening", "Bron / invoer", "Open punt"],
          [("", a, c, f"{b}; {d}", e) for a, b, c, d, e in DOC.MAPPING], cols=[2, 3, 4, 5, 6])
    head("Bijlage C — Vragen vóór productie")
    table(["Nr", "Thema", "Vraag", "Raakt"], [(i + 1, t, q, x) for i, (t, q, x) in enumerate(DOC.PRODUCTIEVRAGEN)], cols=[2, 3, 4, 5])
    head("Bijlage D — Veldwoordenboek (hoe kolommen gevonden worden)")
    para("Kolommen worden op naam gezocht, niet op positie. De volgorde in de export mag dus veranderen. Status ONTBREEKT = kolom niet gevonden (Controle meldt een FOUT).",
         color=GREY, italic=True)
    for label, sheet, fields, cols in (("ZPMW03", "STG_ZPMW03", z_fields(), ZCOLS), ("Woningetalage", "STG_Woningetalage", e_fields(), ECOLS)):
        rows = []
        for f in fields:
            c = cols[f["n"]]
            if "alias" in f:
                m = "Alias: " + " / ".join(f["alias"])
            elif "occ2" in f:
                m = "2e voorkomen dubbele kop (omschrijving)"
            elif f["n"] in ("ProjSrt_code", "PC_code"):
                m = "1e voorkomen dubbele kop (code)"
            elif "ym" in f:
                m = "Via Invoer › 6 (jaarkolommen)"
            else:
                m = "Exacte kop"
            rows.append((label, f["n"], f"={sheet}!{c}2", f"={sheet}!{c}4", m))
        table(["Bron", "Intern veld", "Gezochte kop", "Status", "Methode"], rows, cols=[2, 3, 4, 5, 6])
        r += 1
    ws.conditional_formatting.add(f"E1:E{r+200}", FormulaRule(formula=['E1="ONTBREEKT"'], font=Font(color=RED, bold=True)))
    head("Bijlage E — Testresultaten")
    if results:
        para("Uitgevoerd met tools/test_rdfo.py: kopieën van dit bestand worden aangepast zoals een gebruiker doet en vergeleken met een onafhankelijke berekening.",
             color=GREY, italic=True)
        table(["Nr", "Test", "Uitkomst", "Resultaat"], [(i + 1, a, b, c) for i, (a, b, c) in enumerate(results)], cols=[2, 3, 4, 5])
    else:
        para("Nog geen testrun opgenomen.", color=GREY)
    landscape(ws, f"A1:F{r}")


# ---------------------------------------------------------------------------------------------
# Hoofdroutine
# ---------------------------------------------------------------------------------------------
def code_name(sheet):
    return "sh" + "".join(ch for ch in sheet if ch.isalnum())


def demo_years(rj=2026):
    """Laatste jaar (1..8 vanaf RJ-1) met gegevens in de bronnen; bepaalt de startweergave van Outlook-kolommen en -grafieken."""
    z = load_workbook(DATA / SRC_Z_FILE, read_only=True)["ZPMW03"]
    e = load_workbook(DATA / SRC_E_FILE, read_only=True)["Woningetalage"]
    zr = list(z.iter_rows(values_only=True))
    er = list(e.iter_rows(values_only=True))
    zh, eh = list(zr[0]), list(er[0])
    zsum = lambda h: sum(abs(r[zh.index(h)] or 0) for r in zr[1:]) if h in zh else 0  # noqa: E731
    esum = lambda h: sum(abs(r[eh.index(h)] or 0) for r in er[1:]) if h in eh else 0  # noqa: E731
    oi_years = {r[zh.index("Orderintake")].year for r in zr[1:] if hasattr(r[zh.index("Orderintake")], "year")}
    last = 2
    for i in range(OUT_YEARS):
        y = rj - 1 + i
        omz = zsum("Bedrijfsopbrengsten VJ") if i == 0 else (zsum("Bedrijfsopbrengsten LJ") + zsum(f"OP.BOB {y}") if i == 1 else zsum(f"OP.BOB {y}"))
        won = esum("Huidig jaar") if i == 1 else (esum("Volgend jaar") if i == 2 else (esum(y) if i > 2 else 0))
        if omz or won or y in oi_years:
            last = i + 1
    later = sum(zsum(f"OP.BOB {rj+o}") for o in (7, 8)) + zsum("OP.BOB onverd.")
    return last, later != 0


def build(out, test_results=None):
    wb = Workbook()
    wb.remove(wb.active)
    wb.code_name = "ThisWorkbook"
    sh = {}
    for n, tab, visible in SHEETS:
        sh[n] = wb.create_sheet(n)
        sh[n].sheet_properties.tabColor = tab
        sh[n].sheet_properties.codeName = code_name(n)
        if not visible:
            sh[n].sheet_state = "hidden"
    for extra in ("PPT", "Dias"):
        sh[extra] = wb.create_sheet(extra)
        sh[extra].sheet_state = "hidden"
        sh[extra].sheet_properties.codeName = code_name(extra)
    hist = load_workbook(DATA / HIST_FILE)
    build_lijsten(wb, sh["Lijsten"])
    mapping = build_invoer(wb, sh["Invoer"], hist)
    nz = copy_source(sh["Bron_ZPMW03"], DATA / SRC_Z_FILE, "ZPMW03")
    ne = copy_source(sh["Bron_Woningetalage"], DATA / SRC_E_FILE, "Woningetalage")
    build_stg(wb, sh["STG_ZPMW03"], "Bron_ZPMW03", z_fields(), z_derived(), ZCOLS)
    build_stg(wb, sh["STG_Woningetalage"], "Bron_Woningetalage", e_fields(), e_derived(), ECOLS)
    put(sh["STG_ZPMW03"], "A5", "Staging ZPMW03 — niet invullen", bold=True, color=NAVY)
    put(sh["STG_Woningetalage"], "A5", "", color=GREY)
    zsrc = load_workbook(DATA / SRC_Z_FILE)["ZPMW03"]
    hdr = [c.value for c in zsrc[1]]
    ifrs = {row[hdr.index("WBS-element")]: row[hdr.index("IFRS J/N")] for row in zsrc.iter_rows(min_row=2, values_only=True)}
    rows = hist_rows(hist, mapping, ifrs)
    counts = {v: sum(1 for r in rows if r[5] == v) for v in VERSIES}
    build_historie(wb, sh["Historie"], rows, counts, nz, ne)
    build_verw(wb, sh["VERW_Project"])
    build_calc(wb, sh["CALC_Versies"])
    build_vastleggen(wb, sh["Concept"])
    build_rdfo(wb, sh["RDFO"])
    build_stoplicht(wb, sh["Stoplicht"])
    build_projecten(wb, sh["Projecten"])
    build_dashboard(wb, sh["Dashboard"])
    build_woningen(wb, sh["Woningen"])
    OUT_CTX["demo_n"], OUT_CTX["demo_later"] = demo_years()
    build_outlook(wb, sh["Outlook"])
    outlook_chart(sh["Dashboard"], "K31", "omzet", 17, 8)
    first = list(mapping.values())[0][2]
    build_projectdetail(wb, sh["Projectdetail"], first)
    build_controle(wb, sh["Controle"])
    build_ppt(wb, sh["PPT"])
    build_dias(wb, sh["Dias"])
    build_ppt_spec(wb, sh["PPT"])
    build_start(wb, sh["Start"])
    build_handleiding(sh["Handleiding"], test_results)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(out)
    return {"ledger_rows": len(rows), "nz": nz, "ne": ne}


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "RDFO_prototype_raw.xlsx")
    print(build(out))
