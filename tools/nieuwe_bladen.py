"""De nieuwe tabbladen van het FO-werkbestand (Start, Prognose kosten/opbrengsten, Vorige prognose,
Controle oud FO), gemaakt met openpyxl in een los werkboek; build_werkbestand.py voegt ze daarna
op XML-niveau in het werkbestand.

Vaste plekken (de macro's in vba/FOWerkbestand.bas gebruiken ze):
  Start            C6 projectnaam, C8/D8 eerste prognosekwartaal, C9/D9 oplevering, F8/F9 kwartaalindex,
                   rijen 15-18 invoertabbladen (F = laatst ingelezen, G = bestand), rij 59-61 dia 10
  Prognose kosten  zelfde rijen en kolommen als 'CF - kosten': rij 58 'Onderdelen', posten 59-70,
                   rij 75 totaal, rij 77 periode in cashflow SAP; kolom C = methode
  Prognose opbrengsten  zelfde kolommen als 'CF - opbrengsten': rij 194 totaal, rij 196 periode;
                   regels per termijnblok vanaf rij 11 (voorstel / handmatig / definitief)
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

# ---- huisstijl (thema van het FO: accent1 = geel, accent3 = donkerblauw, accent4 = grijsblauw) ----
FONT = 'Heijmans-Light'
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name='Heijmans-BoldCon', size=16, color=Color(theme=6))
F_SECTION = Font(name='Heijmans-BoldCon', size=11, color=Color(theme=0))
F_NOTE = Font(name=FONT, size=9, italic=True, color='FF6B7280')
F_HELP = Font(name=FONT, size=8, color='FF9CA3AF')
F_HEAD = Font(name=FONT, size=10, bold=True)


def fill(theme=None, tint=0.0, rgb=None):
    c = Color(rgb=rgb) if rgb else Color(theme=theme, tint=tint)
    return PatternFill('solid', fgColor=c, bgColor=c)


FILL_SECTION = fill(6)                 # donkerblauw
FILL_HEAD = fill(4)                    # Heijmans-geel (zoals de koppen in het FO)
FILL_INPUT = fill(7, 0.6)              # handmatig (zoals 'Toelichting onderlegger')
FILL_SAP = fill(6, 0.8)                # exports uit SAP
FILL_TOTAL = fill(7, 0.4)
FILL_CALC = fill(rgb='FFF3F4F6')       # berekend
THIN = Side(style='thin', color='FFD1D5DB')
B_BOTTOM = Border(bottom=THIN)
NUM = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
NUMK = '#,##0.0;-#,##0.0;"-"'
WRAP = Alignment(wrap_text=True, vertical='top')

OK_FILL = PatternFill('solid', fgColor='FFD1FAE5', bgColor='FFD1FAE5')
FOUT_FILL = PatternFill('solid', fgColor='FFFECACA', bgColor='FFFECACA')
LETOP_FILL = PatternFill('solid', fgColor='FFFEF3C7', bgColor='FFFEF3C7')
OVERRIDE_FILL = PatternFill('solid', fgColor='FFFDE68A', bgColor='FFFDE68A')

MAANDEN = '"januari","februari","maart","april","mei","juni","juli","augustus","september","oktober","november","december"'

# kolommen van de bronbladen
K_COLS = [L(c) for c in range(8, 28)]        # 'CF - kosten' H..AA
O_COLS = [L(c) for c in range(7, 27)]        # 'CF - opbrengsten' G..Z
K_POSTEN = list(range(59, 71))               # 'CF - kosten' posten (blok 'Onderdelen')
# termijnblokken op 'CF - opbrengsten' (zoals de formule in rij 194) met de naam van het blok
O_BLOKKEN = [((129, 139), 'B5'), ((143, 156), 'B15'), ((161, 171), 'B28'), ((175, 185), 'B38'),
             ((190, 191), 'B189')]
P_HUIDIG, P_OPL = 'p_huidig', 'p_oplevering'


def cell(ws, ref, value=None, font=F_BASE, fill_=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font
    if fill_ is not None:
        c.fill = fill_
    if fmt:
        c.number_format = fmt
    if align is not None:
        c.alignment = align
    if border is not None:
        c.border = border
    return c


def section(ws, row, text, c1=2, c2=12):
    for c in range(c1, c2 + 1):
        ws.cell(row, c).fill = FILL_SECTION
    cell(ws, '%s%d' % (L(c1), row), text, F_SECTION, FILL_SECTION)
    ws.row_dimensions[row].height = 20


def status_rules(ws, rng):
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s,2)="OK"' % rng.split(':')[0].replace('$', '')],
                                                   fill=OK_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("FOUT",%s))' % rng.split(':')[0].replace('$', '')],
                                                   fill=FOUT_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['OR(LEFT(%s,6)="LET OP",LEFT(%s,4)="LEEG")' % (
        rng.split(':')[0].replace('$', ''), rng.split(':')[0].replace('$', ''))], fill=LETOP_FILL, stopIfTrue=True))


# ================================================================================================
#  Start
# ================================================================================================

def maak_start(wb, demo):
    ws = wb.create_sheet('Start')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = 'FEB900'
    widths = {'A': 2, 'B': 60, 'C': 22, 'D': 22, 'E': 14, 'F': 18, 'G': 34, 'H': 22, 'I': 3, 'J': 16, 'K': 16, 'L': 3}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    cell(ws, 'B2', 'FO-werkbestand', F_TITLE)
    cell(ws, 'B3', 'Eén bestand voor het kwartaal-FO: SAP- en SF-exports erin, prognose kosten en opbrengsten '
                   '(voorstel of handmatig), cashflow, scenario\'s en de PowerPoint.', F_NOTE)
    cell(ws, 'G2', 'Legenda', F_BOLD)
    cell(ws, 'H2', 'handmatig invullen', F_BASE, FILL_INPUT)
    cell(ws, 'H3', 'export uit SAP / SF', F_BASE, FILL_SAP)
    cell(ws, 'H4', 'berekend', F_BASE, FILL_CALC)

    # 1 project en periode -----------------------------------------------------------------------
    section(ws, 5, '1 · PROJECT EN PERIODE')
    rows = [
        (6, 'Projectnaam', demo['project'], None, 'op de titeldia en in de voettekst'),
        (7, 'SAP-boekingsperiode van FO - huidig', "=IFERROR(--('FO - huidig'!E4&\"\"),\"\")", None,
         'uit de export (cel E4)'),
    ]
    for r, lab, val, _, note in rows:
        cell(ws, 'B%d' % r, lab)
        c = cell(ws, 'C%d' % r, val, F_BASE, FILL_INPUT if r == 6 else FILL_CALC)
        cell(ws, 'G%d' % r, note, F_NOTE)
    cell(ws, 'B8', 'Eerste prognosekwartaal · jaar en kwartaal')
    cell(ws, 'C8', '=IF(C7="","",INT(C7/100))', F_BASE, FILL_CALC)
    cell(ws, 'D8', '=IF(C7="","",ROUNDUP(MOD(C7,100)/3,0))', F_BASE, FILL_CALC)
    cell(ws, 'E8', '=IF(C8="","","Q"&D8&" "&C8)', F_BOLD)
    cell(ws, 'F8', '=IF(C8="",0,C8*4+D8)', F_HELP, None, ';;;')
    cell(ws, 'G8', 'uit de boekingsperiode; overschrijven mag', F_NOTE)
    cell(ws, 'B9', 'Oplevering · jaar en kwartaal')
    cell(ws, 'C9', demo['oplevering'][0], F_BASE, FILL_INPUT)
    cell(ws, 'D9', demo['oplevering'][1], F_BASE, FILL_INPUT)
    cell(ws, 'E9', '=IF(C9="","","Q"&D9&" "&C9)', F_BOLD)
    cell(ws, 'F9', '=IF(C9="",0,C9*4+D9)', F_HELP, None, ';;;')
    cell(ws, 'G9', 'voorstel kosten: restant budget gelijk verdeeld t/m dit kwartaal', F_NOTE)
    cell(ws, 'B10', 'Cijfers per')
    cell(ws, 'C10', '=IF(C7="","",EOMONTH(DATE(C8,MOD(C7,100),1),0))', F_BASE, FILL_INPUT, 'dd-mm-yyyy')
    cell(ws, 'G10', 'standaard: einde boekingsperiode; overschrijven mag', F_NOTE)
    cell(ws, 'B11', 'Tekst in PowerPoint (dia 1 en 10)')
    cell(ws, 'C11', '=IF(C10="","","Cijfers per "&DAY(C10)&" "&CHOOSE(MONTH(C10),%s)&" "&YEAR(C10))' % MAANDEN,
         F_BASE, FILL_CALC)
    for r in (8, 9):
        for col in 'CD':
            ws['%s%d' % (col, r)].alignment = Alignment(horizontal='center')

    # 2 invoer -------------------------------------------------------------------------------------
    section(ws, 13, '2 · INVOER · plakken of inlezen met de knoppen rechts')
    heads = ['Tabblad', 'Bron', 'Periode / exportdatum', 'Regels', 'Laatst ingelezen', 'Bestand', 'Status']
    for i, h in enumerate(heads):
        cell(ws, '%s14' % L(2 + i), h, F_HEAD, FILL_HEAD)
    inv = [
        (15, 'FO - vorig', 'SAP PEW, lay-out FR4 · vorige periode', "=IFERROR(--('FO - vorig'!E4&\"\"),\"\")",
         "=COUNTA('FO - vorig'!I:I)", '=IF(E15<2,"LEEG",IF(AND(ISNUMBER(D15),ISNUMBER(D16)),IF(D15<D16,"OK","LET OP: niet ouder dan huidig"),"OK"))'),
        (16, 'FO - huidig', 'SAP PEW, lay-out FR4 · huidige periode', "=IFERROR(--('FO - huidig'!E4&\"\"),\"\")",
         "=COUNTA('FO - huidig'!I:I)", '=IF(E16<2,"LEEG","OK")'),
        (17, 'FO - actuals', 'SAP regels, lay-out FR5',
         "=IFERROR(_xlfn.AGGREGATE(14,6,('FO - actuals'!P2:P6000*100+'FO - actuals'!O2:O6000)/('FO - actuals'!P2:P6000<>\"\"),1),\"\")",
         "=MAX(0,COUNTA('FO - actuals'!A:A)-1)",
         '=IF(E17<1,"LEEG",IF(AND(ISNUMBER(D16),ISNUMBER(D17)),IF(INT(D17/100)*12+MOD(D17,100)>=INT(D16/100)*12+MOD(D16,100)-1,"OK","LET OP: ouder dan FO - huidig"),"OK"))'),
        (18, 'Prijslijst dump', 'Salesforce prijslijst per bouwnummer', '=Prijslijst!C1',
         "=MAX(0,COUNTA('Prijslijst dump'!A:A)-1)", '=IF(E18<1,"LEEG","OK")'),
    ]
    for r, tab, bron, per, n, st in inv:
        cell(ws, 'B%d' % r, tab, F_BOLD, FILL_SAP)
        cell(ws, 'C%d' % r, bron, F_NOTE)
        cell(ws, 'D%d' % r, per, F_BASE, FILL_CALC, 'dd-mm-yyyy' if r == 18 else '0',
             Alignment(horizontal='left'))
        cell(ws, 'E%d' % r, n, F_BASE, FILL_CALC, '#,##0')
        cell(ws, 'F%d' % r, None, F_BASE, None, 'dd-mm-yyyy hh:mm', Alignment(horizontal='left'))
        cell(ws, 'G%d' % r, None, F_NOTE)
        cell(ws, 'H%d' % r, st, F_BOLD)
        ws.row_dimensions[r].height = 22
    status_rules(ws, 'H15:H18')
    cell(ws, 'B19', 'Plakken: kopieer de export (vanaf A1) en klik Plakken. Inlezen: kies het exportbestand. '
                    'De hulpkolommen rechts van de export lopen daarna automatisch mee.', F_NOTE)

    # 3 prognose ------------------------------------------------------------------------------------
    section(ws, 21, '3 · PROGNOSE')
    cell(ws, 'C22', 'definitief', F_HEAD, FILL_HEAD)
    cell(ws, 'D22', 'PEW', F_HEAD, FILL_HEAD)
    cell(ws, 'E22', 'verschil', F_HEAD, FILL_HEAD)
    prog = [
        (23, 'Totaal kosten (Prognose kosten ↔ PEW)', "='Prognose kosten'!AC75", "='Prognose kosten'!D55"),
        (24, 'Totaal opbrengsten (Prognose opbrengsten ↔ PEW)', "='Prognose opbrengsten'!AB194",
         "=N('2. Ontw PEW'!G27)*1000"),
    ]
    for r, lab, a, b in prog:
        cell(ws, 'B%d' % r, lab)
        cell(ws, 'C%d' % r, a, F_BASE, FILL_CALC, NUM)
        cell(ws, 'D%d' % r, b, F_BASE, FILL_CALC, NUM)
        cell(ws, 'E%d' % r, '=C%d-D%d' % (r, r), F_BASE, FILL_CALC, NUM)
    cell(ws, 'B25', 'Kostenposten op Voorstel')
    cell(ws, 'C25', "=COUNTIF('Prognose kosten'!C59:C70,\"Voorstel\")&\" van \"&COUNTIF('Prognose kosten'!B59:B70,\"?*\")",
         F_BASE, FILL_CALC, None, Alignment(horizontal='right'))
    cell(ws, 'G25', 'kies per post op tab Prognose kosten (kolom C)', F_NOTE)
    cell(ws, 'B26', 'Handmatige bedragen opbrengsten')
    cell(ws, 'C26', "=COUNT('Prognose opbrengsten'!G12:Z12,'Prognose opbrengsten'!G16:Z16,'Prognose opbrengsten'!G20:Z20,"
                    "'Prognose opbrengsten'!G24:Z24,'Prognose opbrengsten'!G28:Z28)", F_BASE, FILL_CALC)
    cell(ws, 'G26', 'gele regels op tab Prognose opbrengsten', F_NOTE)
    cell(ws, 'B27', 'Eindsaldo cashflow (1. Cashflow)')
    cell(ws, 'C27', "=LOOKUP(2,1/ISNUMBER('1. Cashflow'!H4:H60),'1. Cashflow'!H4:H60)", F_BOLD, FILL_CALC, NUM)

    # 4 controles -----------------------------------------------------------------------------------
    section(ws, 29, '4 · CONTROLES')
    cell(ws, 'B30', 'Controle', F_HEAD, FILL_HEAD)
    cell(ws, 'C30', 'Uitkomst', F_HEAD, FILL_HEAD)
    cell(ws, 'D30', 'Status', F_HEAD, FILL_HEAD)
    cell(ws, 'E30', None, F_HEAD, FILL_HEAD)
    cell(ws, 'F30', 'Waar', F_HEAD, FILL_HEAD)
    cell(ws, 'G30', None, F_HEAD, FILL_HEAD)
    ven_o = ("SUMPRODUCT(('Prognose opbrengsten'!$G$196:$Z$196>=1)*('Prognose opbrengsten'!$G$196:$Z$196<=60)"
             "*'Prognose opbrengsten'!$G$194:$Z$194)")
    ven_k = ("SUMPRODUCT(('Prognose kosten'!$H$77:$AA$77>=1)*('Prognose kosten'!$H$77:$AA$77<=60)"
             "*'Prognose kosten'!$H$75:$AA$75)")
    checks = [
        ('Invoer compleet', '=COUNTIF(H15:H18,"OK")&" van 4 OK"', '=IF(COUNTIF(H15:H18,"OK")=4,"OK","LET OP")', 'blok 2'),
        ('Aansluiting opbrengsten (SKO ↔ prijslijst ↔ PEW)',
         "=COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")=0,\"OK\",\"FOUT\")", 'Aansluiting opbrengsten, kolom J'),
        ('Opbrengsten volledig ingepland', "=COUNTIF('CF - opbrengsten'!AD1:AD300,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('CF - opbrengsten'!AD1:AD300,\"FOUT\")=0,\"OK\",\"FOUT\")", 'CF - opbrengsten, kolom AD'),
        ('Kosten volledig ingepland (handmatig)',
         "=COUNTIF('CF - kosten'!AD1:AE300,\"FOUT\")+COUNTIF('CF - kosten'!E72,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('CF - kosten'!AD1:AE300,\"FOUT\")+COUNTIF('CF - kosten'!E72,\"FOUT\")=0,\"OK\",\"FOUT\")",
         'CF - kosten, kolom AD/AE'),
        ('Definitieve kosten = PEW per post', "=COUNTIF('Prognose kosten'!AD59:AD70,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('Prognose kosten'!AD59:AD70,\"FOUT\")=0,\"OK\",\"FOUT\")", 'Prognose kosten, kolom AD'),
        ('Posten waar werkelijk > PEW (voorstel verdeelt dan niets)',
         "=COUNTIF('Prognose kosten'!AD13:AD24,\"OVERSCHRIJDING\")&\" post(en)\"",
         "=IF(COUNTIF('Prognose kosten'!AD13:AD24,\"OVERSCHRIJDING\")=0,\"OK\",\"LET OP\")", 'Prognose kosten, voorstel'),
        ('Handmatige opbrengsten in gerealiseerde kwartalen (tellen niet)',
         "=SUM('Prognose opbrengsten'!AC12,'Prognose opbrengsten'!AC16,'Prognose opbrengsten'!AC20,"
         "'Prognose opbrengsten'!AC24,'Prognose opbrengsten'!AC28)&\" cel(len)\"",
         "=IF(SUM('Prognose opbrengsten'!AC12,'Prognose opbrengsten'!AC16,'Prognose opbrengsten'!AC20,"
         "'Prognose opbrengsten'!AC24,'Prognose opbrengsten'!AC28)=0,\"OK\",\"LET OP\")", 'Prognose opbrengsten'),
        ('1. Cashflow sluit aan op PEW', "='1. Cashflow'!D26&\" / \"&'1. Cashflow'!F26",
         "=IF(AND('1. Cashflow'!D26=\"OK\",'1. Cashflow'!F26=\"OK\"),\"OK\",\"FOUT\")", '1. Cashflow, rij 26'),
        ('CF - hulp VENA = definitieve cashflow',
         "=FIXED(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s),0)&\" verschil\"" % (ven_o, ven_k),
         "=IF(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s)<1,\"OK\",\"FOUT\")" % (ven_o, ven_k),
         'CF - hulp VENA, kolom C'),
        ('Niet-gemapte CODA-posten met bedrag (FO - huidig)',
         "=SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))&\" regel(s)\"",
         "=IF(SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))=0,\"OK\",\"FOUT\")",
         'Behind the scens, kolom A-C'),
        ('Scenario actueel (tab Invoer = 1. Cashflow)',
         "=FIXED(ABS(SUM(Invoer!D8:D500)-N('1. Cashflow'!D23))+ABS(SUM(Invoer!F8:F500)-N('1. Cashflow'!F23)),0)&\" verschil\"",
         "=IF(ABS(SUM(Invoer!D8:D500)-N('1. Cashflow'!D23))+ABS(SUM(Invoer!F8:F500)-N('1. Cashflow'!F23))<1,\"OK\",\"LET OP: klik Scenario bijwerken\")",
         'knop Scenario bijwerken'),
        ('Verschil met oud FO (alleen zinvol direct na Oud FO inlezen)',
         "=COUNTIF('Controle oud FO'!M9:M60,\"VERSCHIL\")&\" periode(s)\"",
         "=IF(COUNTIF('Controle oud FO'!M9:M60,\"VERSCHIL\")=0,\"OK\",\"INFO\")", 'tab Controle oud FO'),
    ]
    r = 31
    for lab, uit, st, waar in checks:
        cell(ws, 'B%d' % r, lab)
        cell(ws, 'C%d' % r, uit, F_BASE, FILL_CALC, None, Alignment(horizontal='right'))
        cell(ws, 'D%d' % r, st, F_BOLD)
        cell(ws, 'F%d' % r, waar, F_NOTE)
        r += 1
    last = r - 1
    cell(ws, 'B%d' % r, 'Alle controles', F_BOLD, FILL_TOTAL)
    cell(ws, 'C%d' % r, None, F_BOLD, FILL_TOTAL)
    cell(ws, 'D%d' % r, '=IF(COUNTIF(D31:D%d,"FOUT")>0,COUNTIF(D31:D%d,"FOUT")&" FOUT",IF(COUNTIF(D31:D%d,"LET OP*")>0,"LET OP","OK"))'
         % (last, last, last), F_BOLD)
    status_rules(ws, 'D31:D%d' % r)
    total_row = r

    # 5 stappen -------------------------------------------------------------------------------------
    r = total_row + 2
    section(ws, r, '5 · STAPPEN PER KWARTAAL')
    stappen = [
        'Klik Nieuw kwartaal starten: FO - huidig gaat naar FO - vorig, de cashflow wordt bewaard als vorige prognose.',
        'SAP PEW-export (FR4) van de huidige periode → FO - huidig (Plakken of Inlezen).',
        'SAP actuals (FR5) → FO - actuals. Check SAP op correcties in de werkelijke opbrengsten.',
        'Salesforce prijslijst → Prijslijst dump; zet de exportdatum op tab Prijslijst (cel C1).',
        'Controleer Aansluiting opbrengsten (SKO ↔ prijslijst ↔ PEW) en vraag ontbrekende documentatie op.',
        'Opbrengsten: bouwplanning en verkooptempo op CF - opbrengsten; eventueel bedragen overschrijven op Prognose opbrengsten.',
        'Kosten: per post Handmatig (CF - kosten, zoals altijd) of Voorstel (restant gelijk t/m oplevering) op Prognose kosten.',
        'Alle controles in blok 4 groen.',
        'Klik Scenario bijwerken; bekijk Dashboard en Scenario\'s.',
        'Klik PowerPoint maken (werkt de scenario\'s eerst zelf bij).',
        'Na het FO: correcties bijwerken en CF - hulp VENA inlezen in SAP/VENA.',
    ]
    for i, s in enumerate(stappen, 1):
        cell(ws, 'B%d' % (r + i), '%d.  %s' % (i, s))
        ws.merge_cells('B%d:H%d' % (r + i, r + i))

    # 6 dia 10 --------------------------------------------------------------------------------------
    r = r + len(stappen) + 2
    section(ws, r, '6 · DIA 10 · UPDATE PEW')
    for i, h in enumerate(['Kaart', 'Waarde', 'Toelichting']):
        cell(ws, '%s%d' % (L(2 + i), r + 1), h, F_HEAD, FILL_HEAD)
    dia10 = [('Meest recente FD', 'SKO', 'Inclusief uitruilregeling'),
             ('Huidige PEW-basis', 'FD DO (Aanbieding)', '10 april 2026'),
             ('Mogelijk resultaat', '4,8%', 'Bij rentelasten richting nul')]
    for i, (a, b, c) in enumerate(dia10):
        rr = r + 2 + i
        cell(ws, 'B%d' % rr, a)
        cell(ws, 'C%d' % rr, b, F_BASE, FILL_INPUT)
        cell(ws, 'D%d' % rr, c, F_BASE, FILL_INPUT)
        ws.merge_cells('D%d:F%d' % (rr, rr))
    cell(ws, 'G%d' % (r + 4),
         "=\"PEW-marge nu: \"&IFERROR(FIXED((N('2. Ontw PEW'!G27)-N('2. Ontw PEW'!G20))/N('2. Ontw PEW'!G27)*100,1),\"-\")&\"%\"",
         F_NOTE)
    dia10_row = r + 2

    # knoppen (de macro zet ze neer): kolommen J-K
    cell(ws, 'J5', 'Knoppen', F_BOLD)
    ws.freeze_panes = 'A5'
    return {'dia10_row': dia10_row, 'checks_total': total_row}


# ================================================================================================
#  Prognose kosten
# ================================================================================================

def maak_prognose_kosten(wb, merges57):
    ws = wb.create_sheet('Prognose kosten')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = 'B4C5CC'
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 13
    ws.column_dimensions['F'].width = 13
    ws.column_dimensions['G'].width = 9
    for c in K_COLS:
        ws.column_dimensions[c].width = 12.5
    ws.column_dimensions['AB'].width = 2
    ws.column_dimensions['AC'].width = 13
    ws.column_dimensions['AD'].width = 15
    ws.column_dimensions['AE'].width = 2
    ws.column_dimensions['AF'].width = 50
    src = "'CF - kosten'!"

    cell(ws, 'B1', 'Prognose kosten · voorstel, handmatig en definitief', F_TITLE)
    cell(ws, 'B2', 'Handmatig = de verdeling op CF - kosten (zoals altijd). Voorstel = restant budget (PEW − verwerkt) '
                   'gelijk verdeeld over de kwartalen vanaf het eerste prognosekwartaal t/m oplevering. '
                   'Per post kies je in blok 3 (kolom C) welke telt. Blok 3 gaat naar 1. Cashflow, CF - hulp VENA en het scenario.',
         F_NOTE)
    ws.merge_cells('B2:AA2')
    ws.row_dimensions[2].height = 28
    ws['B2'].alignment = WRAP
    cell(ws, 'B3', 'Eerste prognosekwartaal')
    cell(ws, 'D3', '=Start!E8', F_BOLD)
    cell(ws, 'B4', 'Oplevering')
    cell(ws, 'D4', '=Start!E9', F_BOLD)
    cell(ws, 'F3', '(tab Start)', F_NOTE)

    # hulprijen 7-8: jaar en kwartaalindex per kolom
    cell(ws, 'B7', 'hulp: jaar', F_HELP)
    cell(ws, 'B8', 'hulp: kwartaalindex (jaar × 4 + kw)', F_HELP)
    for i, c in enumerate(K_COLS):
        prev = K_COLS[i - 1] if i else None
        cell(ws, '%s7' % c, '=IF(ISNUMBER(%s%s$57),%s%s$57,%s)' % (src, c, src, c, ('%s7' % prev) if prev else '""'), F_HELP)
        cell(ws, '%s8' % c, '=IF(AND(ISNUMBER(%s7),LEFT(%s%s$58&"",1)="Q"),%s7*4+VALUE(MID(%s%s$58,2,1)),0)'
             % (c, src, c, c, src, c), F_HELP)

    def koprijen(r):
        for c in K_COLS:
            cell(ws, '%s%d' % (c, r), '=IF(%s8>0,"Q"&(%s8-4*%s7)&" \'"&RIGHT(%s7,2),"actuals")' % (c, c, c, c),
                 F_HEAD, FILL_HEAD, None, Alignment(horizontal='center'))
            cell(ws, '%s%d' % (c, r + 1),
                 '=IF(%s8<%s,"verwerkt",IF(%s8<=MAX(%s,%s),"prognose","na oplevering"))' % (c, P_HUIDIG, c, P_HUIDIG, P_OPL),
                 F_HELP, None, None, Alignment(horizontal='center'))

    # blok 1: voorstel (rijen 13-24) ----------------------------------------------------------------
    section(ws, 10, '1 · VOORSTEL · restant budget gelijk verdeeld t/m oplevering', 2, 30)
    for c, h in zip('BDEFG', ['Post', 'PEW (huidig)', 'Verwerkt', 'Restant', 'Kwartalen']):
        cell(ws, '%s11' % c, h, F_HEAD, FILL_HEAD)
    cell(ws, 'C11', None, F_HEAD, FILL_HEAD)
    cell(ws, 'AC11', 'Totaal', F_HEAD, FILL_HEAD)
    cell(ws, 'AD11', 'Check', F_HEAD, FILL_HEAD)
    koprijen(11)
    cell(ws, 'E12', 't/m vorig kwartaal', F_HELP)
    for i, p in enumerate(K_POSTEN):
        r = 13 + i
        cell(ws, 'B%d' % r, '=IF(%sB%d="","",%sB%d)' % (src, p, src, p))
        cell(ws, 'D%d' % r, '=N(%sD%d)' % (src, p), F_BASE, FILL_CALC, NUM)
        cell(ws, 'E%d' % r, '=SUMIFS(%s$H%d:$AA%d,$H$8:$AA$8,"<"&%s)' % (src, p, p, P_HUIDIG), F_BASE, FILL_CALC, NUM)
        cell(ws, 'F%d' % r, '=MAX(0,D%d-E%d)' % (r, r), F_BASE, FILL_CALC, NUM)
        cell(ws, 'G%d' % r, '=COUNTIFS($H$8:$AA$8,">="&%s,$H$8:$AA$8,"<="&MAX(%s,%s))' % (P_HUIDIG, P_HUIDIG, P_OPL),
             F_BASE, FILL_CALC, '0')
        for c in K_COLS:
            cell(ws, '%s%d' % (c, r),
                 '=IF(%s$8<%s,N(%s%s%d),IF(%s$8<=MAX(%s,%s),IF($G%d=0,0,$F%d/$G%d),0))'
                 % (c, P_HUIDIG, src, c, p, c, P_HUIDIG, P_OPL, r, r, r), F_BASE, FILL_CALC, NUM)
        cell(ws, 'AC%d' % r, '=SUM(H%d:AA%d)' % (r, r), F_BOLD, FILL_CALC, NUM)
        cell(ws, 'AD%d' % r, '=IF(B%d="","",IF(D%d-E%d<-1,"OVERSCHRIJDING",IF(ABS(AC%d-D%d)<1,"OK","FOUT")))'
             % (r, r, r, r, r), F_BOLD)
    r = 25
    cell(ws, 'B%d' % r, 'Totaal voorstel', F_BOLD, FILL_TOTAL)
    for c in ['C', 'D', 'E', 'F', 'G'] + K_COLS + ['AC']:
        f = '=SUM(%s13:%s24)' % (c, c) if c not in ('C', 'G') else None
        cell(ws, '%s%d' % (c, r), f, F_BOLD, FILL_TOTAL, NUM)
    status_rules(ws, 'AD13:AD24')

    # blok 2: verschil handmatig - voorstel (rijen 28-42) ----------------------------------------------
    section(ws, 27, '2 · VERSCHIL · handmatig (CF - kosten) min voorstel', 2, 30)
    cell(ws, 'B28', 'Post', F_HEAD, FILL_HEAD)
    for c in 'CDEFG':
        cell(ws, '%s28' % c, None, F_HEAD, FILL_HEAD)
    cell(ws, 'AC28', 'Totaal', F_HEAD, FILL_HEAD)
    koprijen(28)
    for i, p in enumerate(K_POSTEN):
        r = 30 + i
        cell(ws, 'B%d' % r, '=B%d' % (13 + i))
        for c in K_COLS:
            cell(ws, '%s%d' % (c, r), '=N(%s%s%d)-%s%d' % (src, c, p, c, 13 + i), F_BASE, None, NUM)
        cell(ws, 'AC%d' % r, '=SUM(H%d:AA%d)' % (r, r), F_BOLD, None, NUM)
    r = 42
    cell(ws, 'B%d' % r, 'Totaal verschil', F_BOLD, FILL_TOTAL)
    for c in ['C', 'D', 'E', 'F', 'G'] + K_COLS + ['AC']:
        cell(ws, '%s%d' % (c, r), '=SUM(%s30:%s41)' % (c, c) if c in K_COLS + ['AC'] else None, F_BOLD, FILL_TOTAL, NUM)
    cell(ws, 'B44', 'Positief = handmatig hoger dan het voorstel in dat kwartaal. Over de hele looptijd telt elke post op '
                    'tot het PEW, dus het totaal per post is 0 als beide volledig zijn ingepland.', F_NOTE)

    # blok 3: definitief (rijen 55-77, zelfde plek als 'CF - kosten') -----------------------------------
    cell(ws, 'B55', 'Totaal kosten (PEW)', F_BOLD)
    cell(ws, 'D55', '=N(%sD55)' % src, F_BOLD, FILL_CALC, NUM)
    section(ws, 56, '3 · DEFINITIEF · kies per post in kolom C · gaat naar 1. Cashflow, CF - hulp VENA en het scenario', 2, 30)
    for rng in merges57:
        first = rng.split(':')[0]
        col = ''.join(ch for ch in first if ch.isalpha())
        cell(ws, first, '=IF(%s%s="","",%s%s)' % (src, first, src, first), F_HEAD, fill(4, 0.4), None,
             Alignment(horizontal='center'))
        ws.merge_cells(rng)
    cell(ws, 'B58', 'Onderdelen', F_HEAD, FILL_HEAD)
    cell(ws, 'C58', 'Methode', F_HEAD, FILL_HEAD)
    cell(ws, 'D58', 'PEW €', F_HEAD, FILL_HEAD)
    for c in 'EFG':
        cell(ws, '%s58' % c, None, F_HEAD, FILL_HEAD)
    for c in K_COLS:
        cell(ws, '%s58' % c, '=IF(%s%s58="","",%s%s58)' % (src, c, src, c), F_HEAD, FILL_HEAD, None,
             Alignment(horizontal='center'))
    cell(ws, 'AC58', 'Totaal', F_HEAD, FILL_HEAD)
    cell(ws, 'AD58', 'Check', F_HEAD, FILL_HEAD)
    cell(ws, 'AF58', 'Toelichting (CF - kosten)', F_HEAD, FILL_HEAD)
    dv = DataValidation(type='list', formula1='"Handmatig,Voorstel"', allow_blank=False,
                        showErrorMessage=True, errorTitle='Methode', error='Kies Handmatig of Voorstel.')
    ws.add_data_validation(dv)
    for i, p in enumerate(K_POSTEN):
        r = p
        cell(ws, 'B%d' % r, '=IF(%sB%d="","",%sB%d)' % (src, p, src, p))
        cell(ws, 'C%d' % r, 'Handmatig', F_BOLD, FILL_INPUT, None, Alignment(horizontal='center'))
        dv.add('C%d' % r)
        cell(ws, 'D%d' % r, '=N(%sD%d)' % (src, p), F_BASE, FILL_CALC, NUM)
        for c in K_COLS:
            cell(ws, '%s%d' % (c, r), '=IF($C%d="Voorstel",%s%d,N(%s%s%d))' % (r, c, 13 + i, src, c, p), F_BASE, None, NUM)
        cell(ws, 'AC%d' % r, '=ROUND(SUM(H%d:AA%d),2)' % (r, r), F_BOLD, FILL_CALC, NUM)
        cell(ws, 'AD%d' % r, '=IF(B%d="","",IF(AND(AC%d-D%d<1,AC%d-D%d>-1),"OK","FOUT"))' % (r, r, r, r, r), F_BOLD)
        cell(ws, 'AF%d' % r, '=%sAF%d&""' % (src, p), F_NOTE)
    ws.conditional_formatting.add('H59:AA70', FormulaRule(formula=['$C59="Voorstel"'], fill=OVERRIDE_FILL))
    status_rules(ws, 'AD59:AD70')
    cell(ws, 'D71', '=SUM(D59:D70)', F_BOLD, None, NUM)
    cell(ws, 'B72', 'Check:', F_BASE)
    cell(ws, 'D72', '=D71-D55', F_BASE, None, NUM)
    cell(ws, 'E72', '=IF(AND(D72<1,D72>-1),"OK","FOUT")', F_BOLD)
    cell(ws, 'B75', 'Totaal kosten', F_BOLD, FILL_TOTAL)
    for c in ['C', 'D', 'E', 'F', 'G']:
        cell(ws, '%s75' % c, None, F_BOLD, FILL_TOTAL)
    for c in K_COLS:
        cell(ws, '%s75' % c, '=SUM(%s59:%s70)' % (c, c), F_BOLD, FILL_TOTAL, NUM)
    cell(ws, 'AC75', '=SUM(H75:AA75)', F_BOLD, FILL_TOTAL, NUM)
    cell(ws, 'B76', 'Kosten in %')
    for c in K_COLS:
        cell(ws, '%s76' % c, '=IFERROR(%s75/$D$55,0)' % c, F_BASE, None, '0%')
    cell(ws, 'B77', 'Periode in cashflow SAP')
    for c in K_COLS:
        cell(ws, '%s77' % c, '=N(%s%s77)' % (src, c), F_BASE, FILL_CALC, '0', Alignment(horizontal='center'))
    ws.freeze_panes = 'H12'
    return ws


# ================================================================================================
#  Prognose opbrengsten
# ================================================================================================

def maak_prognose_opbrengsten(wb):
    ws = wb.create_sheet('Prognose opbrengsten')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = 'B4C5CC'
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 12
    for c in 'DEF':
        ws.column_dimensions[c].width = 4
    for c in O_COLS:
        ws.column_dimensions[c].width = 12.5
    ws.column_dimensions['AA'].width = 2
    ws.column_dimensions['AB'].width = 13
    ws.column_dimensions['AC'].width = 12
    src = "'CF - opbrengsten'!"

    cell(ws, 'B1', 'Prognose opbrengsten · voorstel en handmatig', F_TITLE)
    cell(ws, 'B2', 'Voorstel = de termijnen volgens CF - opbrengsten (bouwplanning × koopsommen uit Aansluiting opbrengsten). '
                   'Vul in de gele regel een bedrag in om het voorstel voor dat kwartaal te overschrijven; leeg = voorstel. '
                   'In verwerkte kwartalen telt altijd het voorstel. Rij 194 gaat naar 1. Cashflow en CF - hulp VENA.', F_NOTE)
    ws.merge_cells('B2:Z2')
    ws.row_dimensions[2].height = 28
    ws['B2'].alignment = WRAP
    cell(ws, 'B3', 'Eerste prognosekwartaal')
    cell(ws, 'C3', '=Start!E8', F_BOLD)

    cell(ws, 'B8', 'hulp: jaar', F_HELP)
    cell(ws, 'B9', 'hulp: kwartaalindex (jaar × 4 + kw)', F_HELP)
    for i, c in enumerate(O_COLS):
        prev = O_COLS[i - 1] if i else None
        cell(ws, '%s8' % c, '=IF(ISNUMBER(%s%s$127),%s%s$127,%s)' % (src, c, src, c, ('%s8' % prev) if prev else '""'), F_HELP)
        cell(ws, '%s9' % c, '=IF(AND(ISNUMBER(%s8),LEFT(%s%s$128&"",1)="Q"),%s8*4+VALUE(MID(%s%s$128,2,1)),0)'
             % (c, src, c, c, src, c), F_HELP)
        cell(ws, '%s6' % c, '=IF(%s9>0,"Q"&(%s9-4*%s8)&" \'"&RIGHT(%s8,2),IF(ISNUMBER(%s%s$128),%s%s$128&"",""))'
             % (c, c, c, c, src, c, src, c), F_HEAD, FILL_HEAD, None, Alignment(horizontal='center'))
        cell(ws, '%s7' % c, '=IF(%s9<%s,"verwerkt","prognose")' % (c, P_HUIDIG), F_HELP, None, None,
             Alignment(horizontal='center'))
    cell(ws, 'B6', 'Regel', F_HEAD, FILL_HEAD)
    cell(ws, 'C6', None, F_HEAD, FILL_HEAD)
    cell(ws, 'AB6', 'Totaal', F_HEAD, FILL_HEAD)
    cell(ws, 'AC6', 'genegeerd', F_HEAD, FILL_HEAD)

    vrows, drows = [], []
    for k, ((r1, r2), lab) in enumerate(O_BLOKKEN):
        b = 11 + 4 * k
        cell(ws, 'B%d' % b, '=%s%s&""' % (src, lab), F_BOLD)
        cell(ws, 'C%d' % b, 'voorstel', F_NOTE)
        cell(ws, 'C%d' % (b + 1), 'handmatig', F_NOTE)
        cell(ws, 'C%d' % (b + 2), 'definitief', F_BOLD)
        for c in O_COLS:
            cell(ws, '%s%d' % (c, b), '=SUM(%s%s%d:%s%d)' % (src, c, r1, c, r2), F_BASE, FILL_CALC, NUM)
            cell(ws, '%s%d' % (c, b + 1), None, F_BASE, FILL_INPUT, NUM)
            cell(ws, '%s%d' % (c, b + 2), '=IF(OR(%s$9<%s,%s%d=""),%s%d,%s%d)' % (c, P_HUIDIG, c, b + 1, c, b, c, b + 1),
                 F_BOLD, None, NUM)
        for rr in (b, b + 1, b + 2):
            cell(ws, 'AB%d' % rr, '=SUM(G%d:Z%d)' % (rr, rr), F_BOLD if rr == b + 2 else F_BASE, FILL_CALC, NUM)
        cell(ws, 'AC%d' % (b + 1), '=SUMPRODUCT(($G$9:$Z$9<%s)*(G%d:Z%d<>""))' % (P_HUIDIG, b + 1, b + 1), F_BASE, None, '0')
        ws.conditional_formatting.add('G%d:Z%d' % (b + 2, b + 2),
                                      FormulaRule(formula=['AND(G$9>=%s,G%d<>"")' % (P_HUIDIG, b + 1)], fill=OVERRIDE_FILL))
        vrows.append(b)
        drows.append(b + 2)
    cell(ws, 'B31', 'Totaal voorstel', F_BOLD, FILL_TOTAL)
    cell(ws, 'B32', 'Totaal definitief', F_BOLD, FILL_TOTAL)
    cell(ws, 'B33', 'Verschil (handmatig)', F_BOLD)
    cell(ws, 'C31', None, F_BOLD, FILL_TOTAL)
    cell(ws, 'C32', None, F_BOLD, FILL_TOTAL)
    for c in O_COLS + ['AB']:
        cell(ws, '%s31' % c, '=' + '+'.join('%s%d' % (c, r) for r in vrows), F_BOLD, FILL_TOTAL, NUM)
        cell(ws, '%s32' % c, '=' + '+'.join('%s%d' % (c, r) for r in drows), F_BOLD, FILL_TOTAL, NUM)
        cell(ws, '%s33' % c, '=%s32-%s31' % (c, c), F_BASE, None, NUM)
    cell(ws, 'B35', 'De rijen hieronder zijn verborgen: rij 194 en 196 staan op dezelfde plek als op CF - opbrengsten, '
                    'zodat 1. Cashflow en CF - hulp VENA er rechtstreeks naar verwijzen.', F_NOTE)
    for r in range(36, 193):
        ws.row_dimensions[r].hidden = True
    cell(ws, 'B194', 'Totaal opbrengsten (definitief)', F_BOLD, FILL_TOTAL)
    for c in O_COLS:
        cell(ws, '%s194' % c, '=%s32' % c, F_BOLD, FILL_TOTAL, NUM)
    cell(ws, 'AB194', '=SUM(G194:Z194)', F_BOLD, FILL_TOTAL, NUM)
    cell(ws, 'B195', 'Opbrengsten in %')
    for c in O_COLS:
        cell(ws, '%s195' % c, '=IFERROR(%s194/$AB$194,0)' % c, F_BASE, None, '0%')
    cell(ws, 'B196', 'Periode in cashflow SAP')
    for c in O_COLS:
        cell(ws, '%s196' % c, '=N(%s%s196)' % (src, c), F_BASE, FILL_CALC, '0', Alignment(horizontal='center'))
    ws.freeze_panes = 'G7'
    return ws


# ================================================================================================
#  Vorige prognose en Controle oud FO
# ================================================================================================

def maak_vorige_prognose(wb):
    ws = wb.create_sheet('Vorige prognose')
    ws.sheet_view.showGridLines = False
    for k, v in {'A': 2, 'B': 10, 'C': 8, 'D': 16, 'E': 16, 'F': 16, 'G': 16}.items():
        ws.column_dimensions[k].width = v
    cell(ws, 'B2', 'Vorige prognose', F_TITLE)
    cell(ws, 'B3', 'Wordt gevuld door de knop Nieuw kwartaal starten (tab Start): de definitieve cashflow van het '
                   'afgesloten kwartaal, als waarden. 1. Cashflow kolom J (CF vorig kwartaal) komt hieruit.', F_NOTE)
    cell(ws, 'B5', 'Bewaard op', F_BOLD)
    cell(ws, 'D5', None, F_BASE, None, 'dd-mm-yyyy hh:mm', Alignment(horizontal='left'))
    for i, h in enumerate(['Jaar', 'Q', 'Kosten', 'Omzet', 'CF', 'CF (x 1000)']):
        cell(ws, '%s7' % L(2 + i), h, F_HEAD, FILL_HEAD)
    for r in range(8, 68):
        for i in range(6):
            cell(ws, '%s%d' % (L(2 + i), r), None, F_BASE, None, NUM if i >= 2 else None)
    return ws


def maak_controle(wb, oud=None, bestand=''):
    ws = wb.create_sheet('Controle oud FO')
    ws.sheet_view.showGridLines = False
    widths = {'A': 2, 'B': 8, 'C': 6, 'D': 15, 'E': 15, 'F': 11, 'G': 15, 'H': 15, 'I': 11, 'J': 15, 'K': 15, 'L': 11, 'M': 12}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    cell(ws, 'B2', 'Controle oud FO', F_TITLE)
    cell(ws, 'B3', 'Na Oud FO inlezen: de cashflow uit het oude bestand (waarden) naast de cashflow in dit werkbestand. '
                   'Verschil 0 = de werkwijze loopt 1-op-1 door.', F_NOTE)
    cell(ws, 'B5', 'Bestand', F_BOLD)
    cell(ws, 'D5', bestand, F_BASE)
    cell(ws, 'B6', 'Ingelezen op', F_BOLD)
    cell(ws, 'D6', None, F_BASE, None, 'dd-mm-yyyy hh:mm', Alignment(horizontal='left'))
    heads = ['Jaar', 'Q', 'Kosten oud', 'Kosten nu', 'Verschil', 'Omzet oud', 'Omzet nu', 'Verschil', 'CF oud', 'CF nu',
             'Verschil', 'Status']
    for i, h in enumerate(heads):
        cell(ws, '%s8' % L(2 + i), h, F_HEAD, FILL_HEAD)
    for i in range(52):
        r, s = 9 + i, 4 + i              # rij op '1. Cashflow'
        cell(ws, 'B%d' % r, "=IF('1. Cashflow'!B%d=\"\",\"\",'1. Cashflow'!B%d)" % (s, s))
        cell(ws, 'C%d' % r, "=IF('1. Cashflow'!C%d=\"\",\"\",'1. Cashflow'!C%d)" % (s, s))
        for col_old, col_new, col_d, src in (('D', 'E', 'F', 'D'), ('G', 'H', 'I', 'F'), ('J', 'K', 'L', 'H')):
            v = None
            if oud is not None and i < len(oud):
                v = oud[i].get(src)
            cell(ws, '%s%d' % (col_old, r), v, F_BASE, FILL_SAP, NUM)
            cell(ws, '%s%d' % (col_new, r), "=IF(%s%d=\"\",\"\",N('1. Cashflow'!%s%d))" % (col_old, r, src, s), F_BASE,
                 FILL_CALC, NUM)
            cell(ws, '%s%d' % (col_d, r), '=IF(%s%d="","",%s%d-%s%d)' % (col_old, r, col_new, r, col_old, r), F_BASE, None, NUM)
        cell(ws, 'M%d' % r, '=IF(AND(D%d="",G%d="",J%d=""),"",IF(MAX(ABS(N(F%d)),ABS(N(I%d)),ABS(N(L%d)))<1,"OK","VERSCHIL"))'
             % (r, r, r, r, r, r), F_BOLD)
    status_rules(ws, 'M9:M60')
    ws.freeze_panes = 'D9'
    return ws


def bouw(path, demo, merges57, oud_cashflow, oud_bestand):
    wb = Workbook()
    wb.remove(wb.active)
    info = maak_start(wb, demo)
    maak_prognose_opbrengsten(wb)
    maak_prognose_kosten(wb, merges57)
    maak_vorige_prognose(wb)
    maak_controle(wb, oud_cashflow, oud_bestand)
    wb.save(path)
    return info
