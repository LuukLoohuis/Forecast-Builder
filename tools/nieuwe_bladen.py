"""De nieuwe tabbladen van het FO-werkbestand (Start, Cashflow, 2. Ontw PEW in een nieuw jasje, Vorige prognose,
Controle oud FO, hulpbladen), gemaakt met openpyxl in een los werkboek; build_werkbestand.py voegt ze daarna
op XML-niveau in het werkbestand.

Vaste plekken (de macro's in vba/FOWerkbestand.bas gebruiken ze):
  Start            C6 projectnaam, C8/D8 eerste prognosekwartaal, C9/D9 oplevering, F8/F9 kwartaalindex,
                   rijen 15-18 invoertabbladen (F = laatst ingelezen, G = bestand)
  Start            C12/D12 actuals t/m, F12 kwartaalindex (naam p_actuals)
  Vorige prognose  rij 8-67: jaar, Q, kosten, omzet, CF, CF x 1000
  _sjabloon        blad / cel / inhoud: wordt na Oud FO inlezen teruggezet (transportregels, methode, voorstel, ...)
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
    cell(ws, 'B12', 'Actuals t/m · jaar en kwartaal')
    cell(ws, 'C12', '=IFERROR(INT(D17/100),"")', F_BASE, FILL_CALC)
    cell(ws, 'D12', '=IFERROR(ROUNDUP(MOD(D17,100)/3,0),"")', F_BASE, FILL_CALC)
    cell(ws, 'E12', '=IF(C12="","","Q"&D12&" "&C12)', F_BOLD)
    cell(ws, 'F12', '=IF(C12="",0,C12*4+D12)', F_HELP, None, ';;;')
    cell(ws, 'G12', 'laatste boekingsperiode in FO - actuals; t/m dit kwartaal veranderen de scenario\'s niets', F_NOTE)
    for r in (8, 9, 12):
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
        (23, 'Totaal kosten (CF - kosten, rij 75 ↔ PEW)', "=SUM('CF - kosten'!H75:AA75)", "=N('CF - kosten'!D55)"),
        (24, 'Totaal opbrengsten (CF - opbrengsten, rij 194 ↔ PEW)', "=SUM('CF - opbrengsten'!G194:Z194)",
         "=N('2. Ontw PEW'!G27)*1000"),
    ]
    for r, lab, a, b in prog:
        cell(ws, 'B%d' % r, lab)
        cell(ws, 'C%d' % r, a, F_BASE, FILL_CALC, NUM)
        cell(ws, 'D%d' % r, b, F_BASE, FILL_CALC, NUM)
        cell(ws, 'E%d' % r, '=C%d-D%d' % (r, r), F_BASE, FILL_CALC, NUM)
    cell(ws, 'B25', 'Kostenposten op Voorstel')
    cell(ws, 'C25', "=COUNTIF('CF - kosten'!C59:C70,\"Voorstel\")&\" van \"&COUNTIF('CF - kosten'!B59:B70,\"?*\")",
         F_BASE, FILL_CALC, None, Alignment(horizontal='right'))
    cell(ws, 'G25', 'kies per post op CF - kosten, kolom C (blok Onderdelen)', F_NOTE)
    cell(ws, 'B26', 'Handmatige correctie opbrengsten')
    cell(ws, 'C26', "=SUM('CF - opbrengsten'!G192:Z192)", F_BASE, FILL_CALC, NUM)
    cell(ws, 'G26', 'gele regel 192 op CF - opbrengsten', F_NOTE)
    cell(ws, 'B27', 'Eindsaldo cashflow (tab Invoer)')
    cell(ws, 'C27', "=LOOKUP(2,1/ISNUMBER(Invoer!H8:H25),Invoer!H8:H25)", F_BOLD, FILL_CALC, NUM)

    # 4 controles -----------------------------------------------------------------------------------
    section(ws, 29, '4 · CONTROLES')
    cell(ws, 'B30', 'Controle', F_HEAD, FILL_HEAD)
    cell(ws, 'C30', 'Uitkomst', F_HEAD, FILL_HEAD)
    cell(ws, 'D30', 'Status', F_HEAD, FILL_HEAD)
    cell(ws, 'E30', None, F_HEAD, FILL_HEAD)
    cell(ws, 'F30', 'Waar', F_HEAD, FILL_HEAD)
    cell(ws, 'G30', None, F_HEAD, FILL_HEAD)
    ven_o = ("SUMPRODUCT(('CF - opbrengsten'!$G$196:$Z$196>=1)*('CF - opbrengsten'!$G$196:$Z$196<=60)"
             "*'CF - opbrengsten'!$G$194:$Z$194)")
    ven_k = ("SUMPRODUCT(('CF - kosten'!$H$77:$AA$77>=1)*('CF - kosten'!$H$77:$AA$77<=60)"
             "*'CF - kosten'!$H$75:$AA$75)")
    checks = [
        ('Invoer compleet', '=COUNTIF(H15:H18,"OK")&" van 4 OK"', '=IF(COUNTIF(H15:H18,"OK")=4,"OK","LET OP")', 'blok 2'),
        ('Aansluiting opbrengsten (SKO ↔ prijslijst ↔ PEW)',
         "=COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")=0,\"OK\",\"FOUT\")", 'Aansluiting opbrengsten, kolom J'),
        ('Opbrengsten volledig ingepland (bouwtermijnen, transport, verkoop)',
         "=COUNTIF('CF - opbrengsten'!AD1:AD190,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('CF - opbrengsten'!AD1:AD190,\"FOUT\")=0,\"OK\",\"FOUT\")", 'CF - opbrengsten, kolom AD'),
        ('Kosten volledig ingepland (per post de gekozen methode)',
         "=COUNTIF('CF - kosten'!AD1:AE92,\"FOUT\")+COUNTIF('CF - kosten'!E72,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('CF - kosten'!AD1:AE92,\"FOUT\")+COUNTIF('CF - kosten'!E72,\"FOUT\")=0,\"OK\",\"FOUT\")",
         'CF - kosten, kolom AD/AE'),
        ('Posten waar werkelijk > PEW (voorstel verdeelt dan niets)',
         "=COUNTIF('CF - kosten'!AD100:AD111,\"OVERSCHRIJDING\")&\" post(en)\"",
         "=IF(COUNTIF('CF - kosten'!AD100:AD111,\"OVERSCHRIJDING\")=0,\"OK\",\"LET OP\")", 'CF - kosten, voorstel'),
        ('Handmatige correctie opbrengsten', "=FIXED(SUM('CF - opbrengsten'!G192:Z192),0)",
         "=IF(SUM('CF - opbrengsten'!G192:Z192)=0,\"OK\",\"INFO\")", 'CF - opbrengsten, rij 192'),
        ('Cashflow sluit aan op PEW', "=Invoer!D30&\" / \"&Invoer!F30",
         "=IF(AND(Invoer!D30=\"OK\",Invoer!F30=\"OK\"),\"OK\",\"FOUT\")", 'Invoer, rij 30'),
        ('CF - hulp VENA = definitieve cashflow',
         "=FIXED(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s),0)&\" verschil\"" % (ven_o, ven_k),
         "=IF(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s)<1,\"OK\",\"FOUT\")" % (ven_o, ven_k),
         'CF - hulp VENA, kolom C'),
        ('Niet-gemapte CODA-posten met bedrag (FO - huidig)',
         "=SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))&\" regel(s)\"",
         "=IF(SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))=0,\"OK\",\"FOUT\")",
         'Behind the scens, kolom A-C'),
        ('Woningtypes dekken de opbrengsten (aantal × koopsom ÷ eigen opbrengsten)',
         "=FIXED(N(Dashboard!W13)*100,1)&\"%\"",
         "=IF(N(Dashboard!W13)=0,\"LET OP\",\"INFO\")", 'Dashboard W13 · Woningtypes'),
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
        'CF - opbrengsten: per woningtype de bouwtermijnen en de transportregel, onderaan het verkooptempo; correctie in rij 192.',
        'CF - kosten: de bouwtermijnen volgen de planning van de opbrengsten (overschrijven mag); per post Handmatig of Voorstel (kolom C).',
        'Alle controles in blok 4 groen; bekijk Dashboard (cashflow) en 2. Ontw PEW.',
        'Bekijk Dashboard en Scenario\'s: Invoer, Woningtypes en Scenario kosten rekenen automatisch mee met de CF-tabbladen.',
        'Klik PowerPoint maken.',
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
#  2. Ontw PEW in een nieuw jasje (zelfde cellen als het FO, dus Oud FO inlezen blijft 1-op-1)
# ================================================================================================

NUM_PEW = '#,##0;-#,##0;"-"'


def maak_pew(wb, inhoud):
    """inhoud: {cel: waarde of formule} van '2. Ontw PEW' uit het FO. Geeft de extra (opmaak)teksten terug."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import DataBarRule
    ws = wb.create_sheet('2. Ontw PEW')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '003366'
    for k, v in {'A': 2, 'B': 2, 'C': 34, 'D': 2, 'E': 13, 'F': 13, 'G': 13, 'H': 13, 'I': 13, 'J': 2, 'K': 58}.items():
        ws.column_dimensions[k].width = v
    for ref, v in inhoud.items():
        ws[ref] = v
        ws[ref].font = F_BASE
    extra = {'B2': 'Ontwikkeling PEW', 'B3': 'x € 1.000 · huidig t.o.v. de vorige SAP-export en het meest recente FD · '
                                             'groen = gunstig, rood = ongunstig',
             'C6': 'KOSTEN', 'C21': 'OPBRENGSTEN'}
    cell(ws, 'B2', extra['B2'], F_TITLE)
    cell(ws, 'B3', extra['B3'], F_NOTE)
    cell(ws, 'C4', None, F_HELP)
    for c in 'EFGHI':
        cell(ws, '%s4' % c, None, F_HEAD, FILL_SECTION)
    cell(ws, 'E4', None, F_SECTION, FILL_SECTION)
    cell(ws, 'K4', None, F_SECTION, FILL_SECTION)
    for c in 'CEFGHIK':
        cell(ws, '%s5' % c, None, F_HEAD, FILL_HEAD, None, Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws['C5'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[5].height = 30
    for ref in ('C6', 'C21'):
        cell(ws, ref, extra[ref], Font(name='Heijmans-BoldCon', size=10, color=Color(theme=6)))
    for r in list(range(7, 19)) + list(range(22, 26)):
        cell(ws, 'C%d' % r, None, F_BASE, None, None, None, B_BOTTOM)
        cell(ws, 'E%d' % r, None, F_BASE, FILL_INPUT, NUM_PEW, None, B_BOTTOM)
        for c in 'FG':
            cell(ws, '%s%d' % (c, r), None, F_BOLD if c == 'G' else F_BASE, None, NUM_PEW, None, B_BOTTOM)
        for c in 'HI':
            cell(ws, '%s%d' % (c, r), None, F_BASE, FILL_CALC, NUM_PEW, None, B_BOTTOM)
        cell(ws, 'K%d' % r, None, F_NOTE, FILL_INPUT, None, Alignment(wrap_text=True, vertical='top'), B_BOTTOM)
    for r in (20, 27):
        for c in 'CEFGHIK':
            cell(ws, '%s%d' % (c, r), None, F_BOLD if c != 'K' else F_NOTE, FILL_TOTAL, NUM_PEW if c in 'EFGHI' else None)
    for r in (29, 30):
        for c in 'CEFG':
            cell(ws, '%s%d' % (c, r), None, Font(name='Heijmans-BoldCon', size=11, color=Color(theme=6)), None,
                 NUM_PEW if r == 29 else '0.0%')
    for r in (32, 33):
        for c in 'CEFG':
            cell(ws, '%s%d' % (c, r), None, F_HELP, None, '#,##0.0;-#,##0.0;"OK"' if c != 'C' else None)
    # groen = gunstig: kosten lager (H/I > 0), opbrengsten hoger (H/I < 0)
    groen, rood = Font(color='FF15803D', bold=True), Font(color='FFB91C1C', bold=True)
    ws.conditional_formatting.add('H7:I20', CellIsRule(operator='greaterThan', formula=['0.5'], font=groen))
    ws.conditional_formatting.add('H7:I20', CellIsRule(operator='lessThan', formula=['-0.5'], font=rood))
    ws.conditional_formatting.add('H22:I27', CellIsRule(operator='lessThan', formula=['-0.5'], font=groen))
    ws.conditional_formatting.add('H22:I27', CellIsRule(operator='greaterThan', formula=['0.5'], font=rood))
    ws.conditional_formatting.add('G7:G18', DataBarRule(start_type='num', start_value=0, end_type='max',
                                                        color='FFFEB900', showValue=True))
    ch = BarChart()
    ch.type = 'bar'
    ch.grouping = 'clustered'
    ch.gapWidth = 40
    from openpyxl.chart import Series
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.data_source import StrRef
    for col in (5, 6, 7):
        s_ = Series(Reference(ws, min_col=col, min_row=7, max_row=18))
        s_.tx = SeriesLabel(strRef=StrRef("'2. Ontw PEW'!$%s$5" % L(col)))
        ch.series.append(s_)
    ch.set_categories(Reference(ws, min_col=3, min_row=7, max_row=18))
    for s, kleur in zip(ch.series, ['B4C5CC', '9CA3AF', 'FEB900']):
        s.graphicalProperties.solidFill = kleur
        s.graphicalProperties.line.noFill = True
    ch.title = 'Kosten per groep · FD, vorig en huidig (x € 1.000)'
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.number_format = '#,##0'
    ch.y_axis.majorGridlines = None
    ch.x_axis.scaling.orientation = 'maxMin'
    ch.legend.position = 'b'
    ch.height = 15
    ch.width = 18
    ws.add_chart(ch, 'M4')
    ws.freeze_panes = 'D6'
    return extra


# ================================================================================================
#  Vorige prognose en Controle oud FO
# ================================================================================================

def maak_vorige_prognose(wb, vooraf=None, bron=''):
    ws = wb.create_sheet('Vorige prognose')
    ws.sheet_view.showGridLines = False
    for k, v in {'A': 2, 'B': 10, 'C': 8, 'D': 16, 'E': 16, 'F': 16, 'G': 16}.items():
        ws.column_dimensions[k].width = v
    cell(ws, 'B2', 'Vorige prognose', F_TITLE)
    cell(ws, 'B3', 'Wordt gevuld door Nieuw kwartaal starten (de cashflow van het afgesloten kwartaal, als waarden) of door '
                   'Oud FO inlezen (kolom CF vorig kwartaal van het oude FO). Tab Invoer kolom J komt hieruit.', F_NOTE)
    cell(ws, 'B5', 'Bewaard op', F_BOLD)
    cell(ws, 'D5', bron or None, F_BASE, None, 'dd-mm-yyyy hh:mm', Alignment(horizontal='left'))
    for i, h in enumerate(['Jaar', 'Q', 'Kosten', 'Omzet', 'CF', 'CF (x 1000)']):
        cell(ws, '%s7' % L(2 + i), h, F_HEAD, FILL_HEAD)
    for i, r in enumerate(range(8, 68)):
        rij = (vooraf or [])[i] if vooraf and i < len(vooraf) else None
        vals = [None] * 6
        if rij:
            vals = [rij['B'], rij.get('C') or None, None, None, None, rij.get('J')]
        for j in range(6):
            cell(ws, '%s%d' % (L(2 + j), r), vals[j], F_BASE, None, NUM if j >= 2 else None)
    return ws


def maak_controle(wb, oud=None, bestand=''):
    ws = wb.create_sheet('Controle oud FO')
    ws.sheet_view.showGridLines = False
    widths = {'A': 2, 'B': 8, 'C': 6, 'D': 15, 'E': 15, 'F': 11, 'G': 15, 'H': 15, 'I': 11, 'J': 15, 'K': 15, 'L': 11, 'M': 12}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    cell(ws, 'B2', 'Controle oud FO', F_TITLE)
    cell(ws, 'B3', 'Na Oud FO inlezen: de cashflow uit het oude bestand (waarden) naast tab Invoer, per jaar en kwartaal. '
                   'Verschil 0 = de werkwijze loopt 1-op-1 door.', F_NOTE)
    cell(ws, 'B5', 'Bestand', F_BOLD)
    cell(ws, 'D5', bestand, F_BASE)
    cell(ws, 'B6', 'Ingelezen op', F_BOLD)
    cell(ws, 'D6', None, F_BASE, None, 'dd-mm-yyyy hh:mm', Alignment(horizontal='left'))
    heads = ['Jaar', 'Q', 'Kosten oud', 'Kosten nu', 'Verschil', 'Omzet oud', 'Omzet nu', 'Verschil', 'CF oud', 'CF nu',
             'Verschil', 'Status']
    for i, h in enumerate(heads):
        cell(ws, '%s8' % L(2 + i), h, F_HEAD, FILL_HEAD)
    cf = "Invoer!$%s$8:$%s$25"
    for i in range(52):
        r = 9 + i
        rij = oud[i] if oud and i < len(oud) else {}
        cell(ws, 'B%d' % r, rij.get('B'), F_BASE, FILL_SAP)
        cell(ws, 'C%d' % r, rij.get('C') or None, F_BASE, FILL_SAP)
        for col_old, col_new, col_d, src in (('D', 'E', 'F', 'D'), ('G', 'H', 'I', 'F'), ('J', 'K', 'L', 'H')):
            cell(ws, '%s%d' % (col_old, r), rij.get(src), F_BASE, FILL_SAP, NUM)
            cell(ws, '%s%d' % (col_new, r),
                 '=IF(%s%d="","",SUMIFS(%s,%s,$B%d,%s,$C%d&""))' % (col_old, r, cf % (src, src), cf % ('B', 'B'), r,
                                                                   cf % ('C', 'C'), r), F_BASE, FILL_CALC, NUM)
            cell(ws, '%s%d' % (col_d, r), '=IF(%s%d="","",%s%d-%s%d)' % (col_old, r, col_new, r, col_old, r), F_BASE, None, NUM)
        cell(ws, 'M%d' % r, '=IF(AND(D%d="",G%d="",J%d=""),"",IF(MAX(ABS(N(F%d)),ABS(N(I%d)),ABS(N(L%d)))<1,"OK","VERSCHIL"))'
             % (r, r, r, r, r, r), F_BOLD)
    status_rules(ws, 'M9:M60')
    ws.freeze_panes = 'D9'
    return ws


# ================================================================================================
#  hulpbladen: stijlen (alleen voor de bouw) en sjabloon (voor Oud FO inlezen)
# ================================================================================================

STIJLEN = {
    'kop': (F_HEAD, FILL_HEAD, None, Alignment(horizontal='center', wrap_text=True)),
    'kop_links': (F_HEAD, FILL_HEAD, None, None),
    'getal': (F_BASE, FILL_CALC, NUM, None),
    'getal_tot': (F_BOLD, FILL_TOTAL, NUM, None),
    'label_tot': (F_BOLD, FILL_TOTAL, None, None),
    'invoer': (F_BOLD, FILL_INPUT, None, Alignment(horizontal='center')),
    'invoer_getal': (F_BASE, FILL_INPUT, NUM, None),
    'invoer_label': (F_BOLD, FILL_INPUT, None, None),
    'notitie': (F_NOTE, None, None, None),
    'hulp': (F_HELP, None, None, Alignment(horizontal='center')),
    'hulp_links': (F_HELP, None, None, None),
    'sectie': (F_SECTION, FILL_SECTION, None, None),
    'vet': (F_BOLD, None, None, None),
    'tekst': (F_BASE, None, None, None),
    'aantal': (F_BASE, FILL_CALC, '0', Alignment(horizontal='center')),
    'pct': (F_HELP, FILL_CALC, '0%', None),
    'getal1': (F_BASE, FILL_CALC, NUMK, None),
    'verborgen': (F_HELP, None, ';;;', None),
}


def maak_stijlen(wb):
    ws = wb.create_sheet('_stijlen')
    refs = {}
    for i, (naam, (font, fl, fmt, al)) in enumerate(STIJLEN.items(), 1):
        refs[naam] = 'A%d' % i
        cell(ws, 'A%d' % i, naam, font, fl, fmt, al)
    return refs


def maak_sjabloon(wb):
    ws = wb.create_sheet('_sjabloon')
    ws.sheet_state = 'hidden'
    cell(ws, 'A1', 'Blad', F_BOLD)
    cell(ws, 'B1', 'Cel', F_BOLD)
    cell(ws, 'C1', 'Inhoud (wordt na Oud FO inlezen teruggezet)', F_BOLD)
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['C'].width = 120
    return ws


def bouw(path, demo, oud_cashflow, oud_bestand, pew_inhoud):
    wb = Workbook()
    wb.remove(wb.active)
    info = maak_start(wb, demo)
    info['pew_extra'] = maak_pew(wb, pew_inhoud)
    maak_vorige_prognose(wb, oud_cashflow, 'uit ' + oud_bestand)
    maak_controle(wb, oud_cashflow, oud_bestand)
    info['stijlen'] = maak_stijlen(wb)
    maak_sjabloon(wb)
    h = wb.create_sheet('_koppeling')
    h.sheet_state = 'hidden'
    wb.save(path)
    return info
