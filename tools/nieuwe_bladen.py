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

P_HUIDIG, P_OPL = 'p_huidig', 'p_oplevering'
KO, OO, PW = "'CF - kosten'!", "'CF - opbrengsten'!", "'2. Ontw PEW'!"


def verwijzingen(oz, inv_tot, daeb_type):
    """celverwijzingen naar de FO-tabbladen volgens de opzet van het FO"""
    ob, k, pw = oz['opbrengsten'], oz['kosten'], oz['pew']
    p1, pn = k['posten'][0], k['posten'][-1]
    off = k['voorstel_offset']
    c0 = 3 + 3 * (daeb_type - 1)
    daeb_van = ob['daeb'] or ob['daeb_sjabloon']
    return {
        'k_tot': '%s%s%d:%s%d' % (KO, k['eerste'], k['totaal'], k['laatste'], k['totaal']),
        'k_tot_abs': '%s$%s$%d:$%s$%d' % (KO, k['eerste'], k['totaal'], k['laatste'], k['totaal']),
        'k_per_abs': '%s$%s$%d:$%s$%d' % (KO, k['eerste'], k['periode'], k['laatste'], k['periode']),
        'k_pew': '%sD%d' % (KO, k['pew_totaal']),
        'k_methode': '%sC%d:C%d' % (KO, p1, pn), 'k_namen': '%sB%d:B%d' % (KO, p1, pn),
        'k_status': '%s%s1:%s%d' % (KO, k['status'], _volgende(k['status']), k['voorstel_start'] - 1),
        'k_check': '%sE%d' % (KO, k['check']),
        'k_voorstel_status': '%s%s%d:%s%d' % (KO, k['status'], p1 + off, k['status'], pn + off),
        'k_tot_rij': k['totaal'], 'k_status_kol': k['status'],
        'o_tot': '%s%s%d:%s%d' % (OO, ob['eerste'], ob['totaal'], ob['laatste'], ob['totaal']),
        'o_tot_abs': '%s$%s$%d:$%s$%d' % (OO, ob['eerste'], ob['totaal'], ob['laatste'], ob['totaal']),
        'o_per_abs': '%s$%s$%d:$%s$%d' % (OO, ob['eerste'], ob['periode'], ob['laatste'], ob['periode']),
        'o_corr': '%s%s%d:%s%d' % (OO, ob['eerste'], ob['correctie'], ob['laatste'], ob['correctie']),
        'o_status': '%s%s1:%s%d' % (OO, ob['status'], ob['status'], ob['totaal'] - 1),
        'o_tot_rij': ob['totaal'], 'o_corr_rij': ob['correctie'], 'o_status_kol': ob['status'],
        'o_verkoop': '%d-%d' % (ob['types'][0]['verkoop'], ob['types'][-1]['verkoop']),
        'pew_k': '%s%s%d' % (PW, pw['huidig'], pw['kosten_totaal']),
        'pew_o': '%s%s%d' % (PW, pw['huidig'], pw['opbrengsten_totaal']),
        'inv_h': 'Invoer!H8:H%d' % (inv_tot - 2), 'inv_ok': inv_tot + 3, 'inv_laatste': inv_tot - 2,
        'daeb_naam': 'Woningtypes!$%s$7' % L(c0), 'daeb_aantal': 'Woningtypes!$%s$8' % L(c0),
        'daeb_fees': 'Woningtypes!$%s$54:$%s$63' % (L(c0 + 1), L(c0 + 1)), 'daeb_type': daeb_type,
        'daeb_rijen': 'rij %d e.v.' % daeb_van,
        'vt_status': "'Verkoop & transport'!%s1:%s%d" % (ob['verschil'], ob['verschil'], 9 + 9 * len(ob['types'])),
        'sap_open': "'SAP-correcties'!F%d" % (7 + len(SAP_BOEKING)),
    }


def _volgende(kol):
    from openpyxl.utils import column_index_from_string as ci
    return L(ci(kol) + 1)


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

def maak_start(wb, demo, R):
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
    cell(ws, 'C8', '=IF(F12>0,INT(F12/4),IF(C7="","",INT(C7/100)))', F_BASE, FILL_CALC)
    cell(ws, 'D8', '=IF(F12>0,F12+1-4*INT(F12/4),IF(C7="","",ROUNDUP(MOD(C7,100)/3,0)))', F_BASE, FILL_CALC)
    cell(ws, 'E8', '=IF(C8="","","Q"&D8&" "&C8)', F_BOLD)
    cell(ws, 'F8', '=IF(C8="",0,C8*4+D8)', F_HELP, None, ';;;')
    cell(ws, 'G8', 'het kwartaal na de laatste actuals (blok 2: FO - actuals); overschrijven mag', F_NOTE)
    cell(ws, 'B9', 'Oplevering · jaar en kwartaal')
    cell(ws, 'C9', demo['oplevering'][0], F_BASE, FILL_INPUT)
    cell(ws, 'D9', demo['oplevering'][1], F_BASE, FILL_INPUT)
    cell(ws, 'E9', '=IF(C9="","","Q"&D9&" "&C9)', F_BOLD)
    cell(ws, 'F9', '=IF(C9="",0,C9*4+D9)', F_HELP, None, ';;;')
    cell(ws, 'G9', 'voorstel kosten: restant budget gelijk verdeeld t/m dit kwartaal', F_NOTE)
    cell(ws, 'B10', 'Cijfers per')
    cell(ws, 'C10', '=IF(C7="","",EOMONTH(DATE(INT(C7/100),MOD(C7,100),1),0))', F_BASE, FILL_INPUT, 'dd-mm-yyyy')
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
        (23, 'Totaal kosten (CF - kosten, rij %d ↔ PEW)' % R['k_tot_rij'], "=SUM(%s)" % R['k_tot'], "=N(%s)" % R['k_pew']),
        (24, 'Totaal opbrengsten (CF - opbrengsten, rij %d ↔ PEW)' % R['o_tot_rij'], "=SUM(%s)" % R['o_tot'],
         "=N(%s)*1000" % R['pew_o']),
    ]
    for r, lab, a, b in prog:
        cell(ws, 'B%d' % r, lab)
        cell(ws, 'C%d' % r, a, F_BASE, FILL_CALC, NUM)
        cell(ws, 'D%d' % r, b, F_BASE, FILL_CALC, NUM)
        cell(ws, 'E%d' % r, '=C%d-D%d' % (r, r), F_BASE, FILL_CALC, NUM)
    cell(ws, 'B25', 'Kostenposten op Voorstel of Transport')
    cell(ws, 'C25', "=(COUNTIF(%s,\"Voorstel\")+COUNTIF(%s,\"Transport\"))&\" van \"&COUNTIF(%s,\"?*\")"
         % (R['k_methode'], R['k_methode'], R['k_namen']),
         F_BASE, FILL_CALC, None, Alignment(horizontal='right'))
    cell(ws, 'G25', 'kies per post op CF - kosten, kolom C (blok Onderdelen)', F_NOTE)
    cell(ws, 'B26', 'Handmatige correctie opbrengsten')
    cell(ws, 'C26', "=SUM(%s)" % R['o_corr'], F_BASE, FILL_CALC, NUM)
    cell(ws, 'G26', 'gele regel %d op CF - opbrengsten' % R['o_corr_rij'], F_NOTE)
    cell(ws, 'B27', 'Eindsaldo cashflow (tab Invoer)')
    cell(ws, 'C27', "=LOOKUP(2,1/ISNUMBER(%s),%s)" % (R['inv_h'], R['inv_h']), F_BOLD, FILL_CALC, NUM)

    # 4 controles -----------------------------------------------------------------------------------
    section(ws, 29, '4 · CONTROLES')
    cell(ws, 'B30', 'Controle', F_HEAD, FILL_HEAD)
    cell(ws, 'C30', 'Uitkomst', F_HEAD, FILL_HEAD)
    cell(ws, 'D30', 'Status', F_HEAD, FILL_HEAD)
    cell(ws, 'E30', None, F_HEAD, FILL_HEAD)
    cell(ws, 'F30', 'Waar', F_HEAD, FILL_HEAD)
    cell(ws, 'G30', None, F_HEAD, FILL_HEAD)
    ven_o = "SUMPRODUCT((%s>=1)*(%s<=60)*%s)" % (R['o_per_abs'], R['o_per_abs'], R['o_tot_abs'])
    ven_k = "SUMPRODUCT((%s>=1)*(%s<=60)*%s)" % (R['k_per_abs'], R['k_per_abs'], R['k_tot_abs'])
    checks = [
        ('Invoer compleet', '=COUNTIF(H15:H18,"OK")&" van 4 OK"', '=IF(COUNTIF(H15:H18,"OK")=4,"OK","LET OP")', 'blok 2'),
        ('Aansluiting opbrengsten (SKO ↔ prijslijst ↔ PEW)',
         "=COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")&\" fout\"",
         "=IF(COUNTIF('Aansluiting opbrengsten'!J1:J300,\"FOUT\")=0,\"OK\",\"FOUT\")", 'Aansluiting opbrengsten, kolom J'),
        ('Opbrengsten volledig ingepland (bouwtermijnen, transport, verkoop)',
         "=COUNTIF(%s,\"FOUT\")&\" fout\"" % R['o_status'],
         "=IF(COUNTIF(%s,\"FOUT\")=0,\"OK\",\"FOUT\")" % R['o_status'], 'CF - opbrengsten, kolom %s' % R['o_status_kol']),
        ('Kosten volledig ingepland (per post de gekozen methode)',
         "=COUNTIF(%s,\"FOUT\")+COUNTIF(%s,\"FOUT\")&\" fout\"" % (R['k_status'], R['k_check']),
         "=IF(COUNTIF(%s,\"FOUT\")+COUNTIF(%s,\"FOUT\")=0,\"OK\",\"FOUT\")" % (R['k_status'], R['k_check']),
         'CF - kosten, kolom %s' % R['k_status_kol']),
        ('Posten waar werkelijk > PEW (voorstel verdeelt dan niets)',
         "=COUNTIF(%s,\"OVERSCHRIJDING\")&\" post(en)\"" % R['k_voorstel_status'],
         "=IF(COUNTIF(%s,\"OVERSCHRIJDING\")=0,\"OK\",\"LET OP\")" % R['k_voorstel_status'], 'CF - kosten, voorstel'),
        ('Handmatige correctie opbrengsten', "=FIXED(SUM(%s),0)" % R['o_corr'],
         "=IF(SUM(%s)=0,\"OK\",\"INFO\")" % R['o_corr'], 'CF - opbrengsten, rij %d' % R['o_corr_rij']),
        ('Cashflow sluit aan op PEW', "=Invoer!D%d&\" / \"&Invoer!F%d" % (R['inv_ok'], R['inv_ok']),
         "=IF(AND(Invoer!D%d=\"OK\",Invoer!F%d=\"OK\"),\"OK\",\"FOUT\")" % (R['inv_ok'], R['inv_ok']),
         'Invoer, rij %d' % R['inv_ok']),
        ('Verkoop en transport ingepland (% per woningtype)',
         "=COUNTIF(%s,\"FOUT*\")&\" fout, \"&COUNTIF(%s,\"LET OP*\")&\" let op\"" % (R['vt_status'], R['vt_status']),
         "=IF(COUNTIF(%s,\"FOUT*\")>0,\"FOUT\",IF(COUNTIF(%s,\"LET OP*\")>0,\"LET OP\",\"OK\"))" % (R['vt_status'], R['vt_status']),
         'tab Verkoop & transport'),
        ('SAP-correcties toegewezen aan een SAP-nummer',
         "=FIXED(%s,0)&\" niet toegewezen\"" % R['sap_open'],
         "=IF(ABS(N(%s))<1,\"OK\",\"LET OP\")" % R['sap_open'], 'tab SAP-correcties'),
        ('CF - hulp VENA = definitieve cashflow',
         "=FIXED(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s),0)&\" verschil\"" % (ven_o, ven_k),
         "=IF(ABS(N('CF - hulp VENA'!C3)-%s)+ABS(N('CF - hulp VENA'!C4)-%s)<1,\"OK\",\"FOUT\")" % (ven_o, ven_k),
         'CF - hulp VENA, kolom C'),
        ('Niet-gemapte CODA-posten met bedrag (FO - huidig)',
         "=SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))&\" regel(s)\"",
         "=IF(SUMPRODUCT(ISNUMBER(--'FO - huidig'!X17:X1500)*('FO - huidig'!Y17:Y1500=\"\")*ISNUMBER('FO - huidig'!R17:R1500)*('FO - huidig'!R17:R1500<>0))=0,\"OK\",\"FOUT\")",
         'Behind the scens, kolom A-C'),
        ('DAEB-fees op Woningtypes (type %d, blok 5) = DAEB-blok op CF - opbrengsten' % R['daeb_type'],
         "=IF(%s=\"\",\"geen DAEB\",FIXED(SUM(%s),0)&\" van \"&FIXED(_koppeling!$C$418,0))" % (R['daeb_naam'], R['daeb_fees']),
         "=IF(%s=\"\",\"OK\",IF(%s=\"\",\"LET OP: aantal DAEB-woningen ontbreekt\","
         "IF(ABS(SUM(%s)-_koppeling!$C$418)<1,\"OK\",\"LET OP\")))" % (R['daeb_naam'], R['daeb_aantal'], R['daeb_fees']),
         'CF - opbrengsten %s' % R['daeb_rijen']),
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
        'Controleer Aansluiting opbrengsten (SKO ↔ prijslijst ↔ PEW); boek de SAP-correcties (tab SAP-correcties) en vraag ontbrekende documentatie op.',
        'Verkoop & transport: per woningtype verkocht en getransporteerd in % per kwartaal.',
        'CF - opbrengsten: per woningtype de bouwtermijnen (planning met de ontwikkelaar); correctie in rij %d.'
        % R['o_corr_rij'],
        'CF - kosten: de bouwtermijnen volgen de planning van de opbrengsten (overschrijven mag); per post Handmatig, Voorstel of Transport (kolom C).',
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
    dia10 = demo.get('dia10') or [('Meest recente FD', None, None), ('Huidige PEW-basis', None, None),
                                   ('Mogelijk resultaat', None, None)]
    for i, (a, b, c) in enumerate(dia10):
        rr = r + 2 + i
        cell(ws, 'B%d' % rr, a)
        cell(ws, 'C%d' % rr, b, F_BASE, FILL_INPUT)
        cell(ws, 'D%d' % rr, c, F_BASE, FILL_INPUT)
        ws.merge_cells('D%d:F%d' % (rr, rr))
    cell(ws, 'G%d' % (r + 4),
         "=\"PEW-marge nu: \"&IFERROR(FIXED((N(%s)-N(%s))/N(%s)*100,1),\"-\")&\"%%\"" % (R['pew_o'], R['pew_k'], R['pew_o']),
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


def maak_pew(wb, inhoud, pw):
    """inhoud: {cel: waarde of formule} van '2. Ontw PEW' uit het FO; pw: de opzet van het blad (kolommen en rijen).
    Geeft de extra (opmaak)teksten terug."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.utils import column_index_from_string as ci
    ws = wb.create_sheet('2. Ontw PEW')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '003366'
    kols = pw['kolommen']                                   # E .. laatste delta
    iv, ih = kols.index(pw['vorig']), kols.index(pw['huidig'])
    invoer, delta = kols[:iv], kols[ih + 1:]
    toel = pw['toelichting']
    kop, r1 = pw['kop'], pw['eerste']
    rk = list(range(r1, pw['kosten_totaal'] - 1))           # kostengroepen
    ro = list(range(pw['opbrengsten'], pw['opbrengsten_totaal'] - 1))
    tk, to = pw['kosten_totaal'], pw['opbrengsten_totaal']
    res = [r for r in (pw['resultaat'], pw['resultaat_pct']) if r]
    chk = [r for r in pw['checks'] if r]
    breed = {'A': 2, 'B': 2, 'C': 34, 'D': 2, toel: 58}
    for c in kols:
        breed[c] = 13
    for c in range(ci(kols[-1]) + 1, ci(toel)):
        breed[L(c)] = 2
    for k, v in breed.items():
        ws.column_dimensions[k].width = v
    for ref, v in inhoud.items():
        ws[ref] = v
        ws[ref].font = F_BASE
    extra = {'B2': 'Ontwikkeling PEW', 'B3': 'x € 1.000 · huidig t.o.v. de vorige SAP-export en het meest recente FD · '
                                             'groen = gunstig, rood = ongunstig',
             'C%d' % (r1 - 1): 'KOSTEN', 'C%d' % (ro[0] - 1): 'OPBRENGSTEN'}
    cell(ws, 'B2', extra['B2'], F_TITLE)
    cell(ws, 'B3', extra['B3'], F_NOTE)
    cell(ws, 'C%d' % (kop - 1), None, F_HELP)
    for c in kols:
        cell(ws, '%s%d' % (c, kop - 1), None, F_HEAD, FILL_SECTION)
    cell(ws, '%s%d' % (kols[0], kop - 1), None, F_SECTION, FILL_SECTION)
    cell(ws, '%s%d' % (toel, kop - 1), None, F_SECTION, FILL_SECTION)
    for c in ['C'] + kols + [toel]:
        cell(ws, '%s%d' % (c, kop), None, F_HEAD, FILL_HEAD, None, Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws['C%d' % kop].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[kop].height = 30
    for ref in ('C%d' % (r1 - 1), 'C%d' % (ro[0] - 1)):
        cell(ws, ref, extra[ref], Font(name='Heijmans-BoldCon', size=10, color=Color(theme=6)))
    for r in rk + ro:
        cell(ws, 'C%d' % r, None, F_BASE, None, None, None, B_BOTTOM)
        for c in invoer:
            cell(ws, '%s%d' % (c, r), None, F_BASE, FILL_INPUT, NUM_PEW, None, B_BOTTOM)
        for c in (pw['vorig'], pw['huidig']):
            cell(ws, '%s%d' % (c, r), None, F_BOLD if c == pw['huidig'] else F_BASE, None, NUM_PEW, None, B_BOTTOM)
        for c in delta:
            cell(ws, '%s%d' % (c, r), None, F_BASE, FILL_CALC, NUM_PEW, None, B_BOTTOM)
        cell(ws, '%s%d' % (toel, r), None, F_NOTE, FILL_INPUT, None, Alignment(wrap_text=True, vertical='top'), B_BOTTOM)
    for r in (tk, to):
        for c in ['C'] + kols + [toel]:
            cell(ws, '%s%d' % (c, r), None, F_BOLD if c != toel else F_NOTE, FILL_TOTAL, NUM_PEW if c in kols else None)
    for i, r in enumerate(res):
        for c in ['C'] + kols[:ih + 1]:
            cell(ws, '%s%d' % (c, r), None, Font(name='Heijmans-BoldCon', size=11, color=Color(theme=6)), None,
                 NUM_PEW if i == 0 else '0.0%')
    for r in chk:
        for c in ['C'] + kols[:ih + 1]:
            cell(ws, '%s%d' % (c, r), None, F_HELP, None, '#,##0.0;-#,##0.0;"OK"' if c != 'C' else None)
    # groen = gunstig: kosten lager (delta > 0), opbrengsten hoger (delta < 0)
    groen, rood = Font(color='FF15803D', bold=True), Font(color='FFB91C1C', bold=True)
    if delta:
        d1, d2 = delta[0], delta[-1]
        ws.conditional_formatting.add('%s%d:%s%d' % (d1, r1, d2, tk), CellIsRule(operator='greaterThan', formula=['0.5'], font=groen))
        ws.conditional_formatting.add('%s%d:%s%d' % (d1, r1, d2, tk), CellIsRule(operator='lessThan', formula=['-0.5'], font=rood))
        ws.conditional_formatting.add('%s%d:%s%d' % (d1, ro[0], d2, to), CellIsRule(operator='lessThan', formula=['-0.5'], font=groen))
        ws.conditional_formatting.add('%s%d:%s%d' % (d1, ro[0], d2, to), CellIsRule(operator='greaterThan', formula=['0.5'], font=rood))
    h = pw['huidig']
    ws.conditional_formatting.add('%s%d:%s%d' % (h, rk[0], h, rk[-1]),
                                  DataBarRule(start_type='num', start_value=0, end_type='max', color='FFFEB900', showValue=True))
    ch = BarChart()
    ch.type = 'bar'
    ch.grouping = 'clustered'
    ch.gapWidth = 40
    from openpyxl.chart import Series
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.data_source import StrRef
    reeks = kols[:ih + 1]
    for col in reeks:
        s_ = Series(Reference(ws, min_col=ci(col), min_row=rk[0], max_row=rk[-1]))
        s_.tx = SeriesLabel(strRef=StrRef("'2. Ontw PEW'!$%s$%d" % (col, kop)))
        ch.series.append(s_)
    ch.set_categories(Reference(ws, min_col=3, min_row=rk[0], max_row=rk[-1]))
    kleuren = ['D6DEE2', 'B4C5CC', '9CA3AF', 'FEB900'][-len(reeks):] if len(reeks) <= 4 else None
    for i, s in enumerate(ch.series):
        kl = kleuren[i] if kleuren else ('FEB900' if i == len(ch.series) - 1 else 'B4C5CC')
        s.graphicalProperties.solidFill = kl
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
    ws.add_chart(ch, '%s%d' % (L(ci(toel) + 2), kop - 1))
    ws.freeze_panes = 'D%d' % (kop + 1)
    return extra


# ================================================================================================
#  Verkoop & transport: per woningtype verkocht en getransporteerd in % per kwartaal
# ================================================================================================

PCT = '0.0%;-0.0%;"-"'
ST_LETOP = 'LET OP'


def maak_verkoop_transport(wb, oz):
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import column_index_from_string as ci
    from opzet import VT_BLAD, VT_KEY, vt_rij, vt_totaal
    ob = oz['opbrengsten']
    cols = [L(c) for c in range(ci(ob['eerste']), ci(ob['laatste']) + 1)]
    c1, cn = cols[0], cols[-1]
    chk, stat = ob['check'], ob['verschil']
    vk = ob['verkoop_kop']
    ws = wb.create_sheet(VT_BLAD)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = 'FEB900'
    for k, v in {'A': 2, 'B': 42, 'C': 10, 'D': 3, 'E': 3, 'F': 3, chk: 11, stat: 44}.items():
        ws.column_dimensions[k].width = v
    for c in cols:
        ws.column_dimensions[c].width = 8.5
    cell(ws, 'B2', 'Verkoop & transport', F_TITLE)
    cell(ws, 'B3', 'Per woningtype: verkocht en getransporteerd (notarieel geleverd) in % van alle woningen van het type, '
                   'per kwartaal. Geel = invullen. De eerste kolom bij verkocht is de realisatie t/m nu (Salesforce).', F_NOTE)
    cell(ws, 'B4', 'Getransporteerd kan niet vóór verkocht: transport dat op verkoop wacht, schuift mee. Het transport bepaalt '
                   'de grondtermijn en de bouwtermijnen die dan al vervallen zijn (CF - opbrengsten), en de kostenposten met '
                   'methode Transport (CF - kosten).', F_NOTE)
    OO = "'CF - opbrengsten'!"
    cell(ws, 'B6', 'Kwartaal', F_HEAD, FILL_HEAD)
    cell(ws, 'C6', 'Woningen', F_HEAD, FILL_HEAD)
    cell(ws, '%s6' % chk, 'Totaal', F_HEAD, FILL_HEAD)
    cell(ws, '%s6' % stat, 'Status', F_HEAD, FILL_HEAD)
    for c in cols:
        cell(ws, '%s5' % c, '=IF(ISNUMBER(%s%s$%d),%s%s$%d,"")' % (OO, c, vk - 1, OO, c, vk - 1), F_HELP, None, None,
             Alignment(horizontal='center'))
        cell(ws, '%s6' % c, '=%s%s$%d&""' % (OO, c, vk), F_HEAD, FILL_HEAD, None, Alignment(horizontal='center'))
        cell(ws, '%s%d' % (c, VT_KEY), '=_koppeling!%s$521' % c, F_HELP, None, ';;;')
    cell(ws, 'B%d' % VT_KEY, 'hulp: kwartaalindex', F_HELP)
    n = len(ob['types'])
    for t in ob['types']:
        r = vt_rij(t['k'])
        won = '$C$%d' % r
        cell(ws, 'B%d' % r, '=%sB%d&""' % (OO, t['verkoop']), F_BOLD, FILL_TOTAL)
        cell(ws, 'C%d' % r, '=N(%s$G$%d)' % (OO, t['hb'] - 1), F_BOLD, FILL_TOTAL, '0')
        for c in ['D', 'E', 'F'] + cols + [chk, stat]:
            cell(ws, '%s%d' % (c, r), None, F_BOLD, FILL_TOTAL)
        labels = ['Verkocht · % per kwartaal', 'Verkocht · cumulatief', 'Verkocht · aantal woningen',
                  'Getransporteerd · % per kwartaal', 'Getransporteerd · cumulatief (niet vóór verkocht)',
                  'Getransporteerd · aantal woningen']
        for i, lab in enumerate(labels, 1):
            cell(ws, 'B%d' % (r + i), lab, F_BOLD if i in (1, 4) else F_BASE)
        vw, gw = t['verkoop_waarden'], t['grond_waarden']
        w = t['won'] or 0
        for j, c in enumerate(cols):
            prev = cols[j - 1] if j else None
            if j == 0:      # realisatie t/m nu: uit CF - opbrengsten (Salesforce)
                cell(ws, '%s%d' % (c, r + 1), '=IF(%s=0,0,N(%s%s%d)/%s)' % (won, OO, c, t['verkoop'], won), F_BASE, FILL_SAP, PCT)
                cell(ws, '%s%d' % (c, r + 3), '=N(%s%s%d)' % (OO, c, t['verkoop']), F_BASE, FILL_CALC, '0.#')
            else:
                cell(ws, '%s%d' % (c, r + 1), (vw[j] / w) if w and vw[j] else None, F_BASE, FILL_INPUT, PCT)
                cell(ws, '%s%d' % (c, r + 3), '=ROUND(%s%d*%s,6)' % (c, r + 1, won), F_BASE, FILL_CALC, '0.#')
            cell(ws, '%s%d' % (c, r + 2), '=SUM($%s%d:%s%d)' % (c1, r + 1, c, r + 1), F_HELP, None, PCT)
            cell(ws, '%s%d' % (c, r + 4), (gw[j] / w) if w and gw[j] else None, F_BASE, FILL_INPUT, PCT)
            cell(ws, '%s%d' % (c, r + 7), '=SUM($%s%d:%s%d)' % (c1, r + 4, c, r + 4), F_HELP, None, ';;;')
            cell(ws, '%s%d' % (c, r + 5), '=MIN(%s%d,%s%d)' % (c, r + 7, c, r + 2), F_HELP, None, PCT)
            cell(ws, '%s%d' % (c, r + 6), '=ROUND((%s%d%s)*%s,6)' % (c, r + 5, '-%s%d' % (prev, r + 5) if prev else '', won),
                 F_BASE, FILL_CALC, '0.#')
        cell(ws, '%s%d' % (chk, r + 2), '=%s%d' % (cn, r + 2), F_BOLD, None, PCT)
        cell(ws, '%s%d' % (chk, r + 5), '=%s%d' % (cn, r + 5), F_BOLD, None, PCT)
        cell(ws, '%s%d' % (chk, r + 3), '=SUM(%s%d:%s%d)' % (c1, r + 3, cn, r + 3), F_BASE, None, '0.#')
        cell(ws, '%s%d' % (chk, r + 6), '=SUM(%s%d:%s%d)' % (c1, r + 6, cn, r + 6), F_BASE, None, '0.#')
        tv, tt = '%s%d' % (chk, r + 2), '%s%d' % (chk, r + 5)
        cell(ws, '%s%d' % (stat, r + 2), '=IF(%s=0,"",IF(%s>1.00001,"FOUT: meer dan 100%% verkocht",IF(%s<0.99999,'
                                         '"%s: "&TEXT(1-%s,"0%%")&" nog niet ingepland","OK")))' % (won, tv, tv, ST_LETOP, tv), F_BOLD)
        cell(ws, '%s%d' % (stat, r + 5), '=IF(%s=0,"",IF(SUM(%s%d:%s%d)>1.00001,"FOUT: meer dan 100%% getransporteerd",'
                                         'IF(SUMPRODUCT((%s%d:%s%d-%s%d:%s%d>0.00001)*1)>0,"%s: transport wacht op verkoop (schuift mee)",'
                                         'IF(%s<0.99999,"%s: "&TEXT(1-%s,"0%%")&" nog niet getransporteerd","OK"))))'
             % (won, c1, r + 4, cn, r + 4, c1, r + 7, cn, r + 7, c1, r + 5, cn, r + 5, ST_LETOP, tt, ST_LETOP, tt), F_BOLD)
        cell(ws, 'B%d' % (r + 7), 'hulp: getransporteerd ingevuld, cumulatief', F_HELP, None, ';;;')
    # alle woningen
    T = vt_totaal(n)
    cell(ws, 'B%d' % T, 'Alle woningen (koop)', F_BOLD, FILL_HEAD)
    cell(ws, 'C%d' % T, '=' + '+'.join('C%d' % vt_rij(t['k']) for t in ob['types']), F_BOLD, FILL_HEAD, '0')
    for c in ['D', 'E', 'F'] + cols + [chk, stat]:
        cell(ws, '%s%d' % (c, T), None, F_BOLD, FILL_HEAD)
    for i, lab in enumerate(['Verkocht · aantal woningen', 'Verkocht · cumulatief % van alle woningen',
                             'Getransporteerd · aantal woningen', 'Getransporteerd · cumulatief % van alle woningen'], 1):
        cell(ws, 'B%d' % (T + i), lab, F_BOLD if i in (2, 4) else F_BASE)
    for c in cols:
        cell(ws, '%s%d' % (c, T + 1), '=' + '+'.join('%s%d' % (c, vt_rij(t['k']) + 3) for t in ob['types']), F_BASE, FILL_CALC, '0.#')
        cell(ws, '%s%d' % (c, T + 3), '=' + '+'.join('%s%d' % (c, vt_rij(t['k']) + 6) for t in ob['types']), F_BASE, FILL_CALC, '0.#')
        cell(ws, '%s%d' % (c, T + 2), '=IF($C$%d=0,0,SUM($%s%d:%s%d)/$C$%d)' % (T, c1, T + 1, c, T + 1, T), F_BOLD, None, PCT)
        cell(ws, '%s%d' % (c, T + 4), '=IF($C$%d=0,0,SUM($%s%d:%s%d)/$C$%d)' % (T, c1, T + 3, c, T + 3, T), F_BOLD, None, PCT)
    for i in (1, 3):
        cell(ws, '%s%d' % (chk, T + i), '=SUM(%s%d:%s%d)' % (c1, T + i, cn, T + i), F_BOLD, None, '0.#')
    status_rules(ws, '%s%d:%s%d' % (stat, vt_rij(1), stat, T))
    # grafiek: cumulatief verkocht en getransporteerd (alle woningen)
    ch = LineChart()
    ch.title = 'Verkocht en getransporteerd · cumulatief % van alle woningen'
    ch.height, ch.width = 7.5, 24
    for rr, kleur in ((T + 2, 'FEB900'), (T + 4, '003366')):
        ch.add_data(Reference(ws, min_col=ci(c1) - 1, max_col=ci(cn), min_row=rr), titles_from_data=True, from_rows=True)
        s_ = ch.series[-1]
        s_.graphicalProperties.line.solidFill = kleur
        s_.graphicalProperties.line.width = 28000
        s_.smooth = False
    ch.set_categories(Reference(ws, min_col=ci(c1), max_col=ci(cn), min_row=6))
    ch.y_axis.number_format = '0%'
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 1
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.legend.position = 'b'
    ws.add_chart(ch, 'B%d' % (T + 7))
    ws.freeze_panes = '%s8' % c1
    # kolom F (vóór de eerste kwartaalkolom) als reekstitel van de grafiek
    cell(ws, '%s%d' % (L(ci(c1) - 1), T + 2), 'Verkocht', F_HELP, None, ';;;')
    cell(ws, '%s%d' % (L(ci(c1) - 1), T + 4), 'Getransporteerd', F_HELP, None, ';;;')
    return T


# ================================================================================================
#  SAP-correcties: wat er na de aansluiting in SAP aangepast moet worden, per SAP-nummer
# ================================================================================================

SAP_KOPPELING = {'grond + ontwikkelkosten': (9010, None), 'aanneemsom': (9000, None), 'kao opties': (9000, 5230),
                 'kopersmeerwerk': (9200, 5250), 'kadastrale kosten': (9030, None), 'aanvullende rente': (9110, None)}
SAP_BOEKING = [(9010, None, 'Grond + ontwikkelkosten'), (9000, None, 'Aanneemsom'), (9000, 5230, 'KAO-opties'),
               (9200, 5250, 'Meer-/minderwerk kopers'), (9030, None, 'Kadastrale kosten'), (9110, None, 'Aanvullende rente')]


def maak_sap(wb, oz):
    ws = wb.create_sheet('SAP-correcties')
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = '003366'
    for k, v in {'A': 2, 'B': 26, 'C': 14, 'D': 14, 'E': 34, 'F': 16, 'G': 3, 'H': 40}.items():
        ws.column_dimensions[k].width = v
    cell(ws, 'B2', 'SAP-correcties', F_TITLE)
    cell(ws, 'B3', "Uit de kolom 'Corr' op Aansluiting opbrengsten (prijslijst/Salesforce t.o.v. het actuele budget): wat er "
                   'in SAP aangepast moet worden, per SAP-nummer en kostendrager. Geel = aanpasbaar.', F_NOTE)
    blokken = oz.get('aansluiting') or []
    regels = []
    for b in blokken:
        for r, lab in b['regels']:
            if lab not in regels:
                regels.append(lab)
    # 2 koppeling regel -> SAP-nummer (onder de boekingen, maar eerst bepalen waar)
    n_b = len(SAP_BOEKING)
    b1 = 7
    rn = b1 + n_b - 1
    k0 = rn + 5                                  # sectie koppeling
    m1 = k0 + 2
    mn = m1 + len(regels) - 1
    d0 = mn + 3                                  # sectie detail
    d1 = d0 + 2
    n_d = sum(len(b['regels']) for b in blokken)
    dn = d1 + n_d - 1
    rng = lambda c: '$%s$%d:$%s$%d' % (c, d1, c, max(d1, dn))
    # 1 te boeken
    section(ws, 5, '1 · TE BOEKEN IN SAP · per SAP-nummer en kostendrager', 2, 8)
    for c, h in zip('BCDEF', ['SAP-nummer', 'Kostendrager', '', 'Omschrijving', 'Bedrag (€)']):
        cell(ws, '%s6' % c, h or None, F_HEAD, FILL_HEAD)
    for i, (nr, kd, oms) in enumerate(SAP_BOEKING):
        r = b1 + i
        cell(ws, 'B%d' % r, nr, F_BOLD, FILL_INPUT, '0', Alignment(horizontal='left'))
        cell(ws, 'C%d' % r, kd, F_BASE, FILL_INPUT, '0', Alignment(horizontal='left'))
        cell(ws, 'E%d' % r, oms, F_BASE, FILL_INPUT)
        cell(ws, 'F%d' % r, '=SUMPRODUCT((%s&""=B%d&"")*(%s&""=C%d&"")*%s)' % (rng('D'), r, rng('E'), r, rng('F')),
             F_BOLD, FILL_CALC, NUM)
    cell(ws, 'E%d' % (rn + 1), 'Niet toegewezen (geen of ander SAP-nummer)', F_BASE)
    cell(ws, 'F%d' % (rn + 1), '=ROUND(F%d-SUM(F%d:F%d),2)' % (rn + 2, b1, rn), F_BOLD, FILL_CALC, NUM)
    cell(ws, 'H%d' % (rn + 1), '=IF(ABS(F%d)<1,"OK","LET OP: koppel de regel(s) aan een SAP-nummer (blok 2)")' % (rn + 1), F_BOLD)
    status_rules(ws, 'H%d:H%d' % (rn + 1, rn + 1))
    cell(ws, 'E%d' % (rn + 2), 'Totaal correcties', F_BOLD, FILL_TOTAL)
    cell(ws, 'F%d' % (rn + 2), '=SUM(%s)' % rng('F'), F_BOLD, FILL_TOTAL, NUM)
    # 2 koppeling
    section(ws, k0, '2 · KOPPELING · regel op Aansluiting opbrengsten → SAP-nummer en kostendrager', 2, 8)
    for c, h in zip('BCD', ['Regel', 'SAP-nummer', 'Kostendrager']):
        cell(ws, '%s%d' % (c, k0 + 1), h, F_HEAD, FILL_HEAD)
    for i, lab in enumerate(regels):
        r = m1 + i
        nr, kd = SAP_KOPPELING.get(lab.strip().lower(), (None, None))
        cell(ws, 'B%d' % r, lab, F_BASE)
        cell(ws, 'C%d' % r, nr, F_BASE, FILL_INPUT, '0', Alignment(horizontal='left'))
        cell(ws, 'D%d' % r, kd, F_BASE, FILL_INPUT, '0', Alignment(horizontal='left'))
    # 3 detail
    section(ws, d0, '3 · DETAIL · per woningtype (Aansluiting opbrengsten, kolom Corr)', 2, 8)
    for c, h in zip('BCDEF', ['Woningtype', 'Regel', 'SAP-nummer', 'Kostendrager', 'Bedrag (€)']):
        cell(ws, '%s%d' % (c, d0 + 1), h, F_HEAD, FILL_HEAD)
    AO = "'Aansluiting opbrengsten'!"
    r = d1
    ml = '$B$%d:$B$%d' % (m1, max(m1, mn))
    for b in blokken:
        for rr, lab in b['regels']:
            cell(ws, 'B%d' % r, b['naam'], F_BASE)
            cell(ws, 'C%d' % r, '=%sA%d&""' % (AO, rr), F_BASE)
            for c, mc in (('D', 'C'), ('E', 'D')):
                x = 'INDEX($%s$%d:$%s$%d,MATCH(C%d,%s,0))' % (mc, m1, mc, max(m1, mn), r, ml)
                cell(ws, '%s%d' % (c, r), '=IFERROR(IF(%s="","",%s),"")' % (x, x), F_BASE, FILL_CALC, '0',
                     Alignment(horizontal='left'))
            cell(ws, 'F%d' % r, '=N(%s%s%d)' % (AO, b['corr'], rr), F_BASE, FILL_CALC, NUM)
            r += 1
    ws.freeze_panes = 'A5'
    return {'niet_toegewezen': 'F%d' % (rn + 1), 'totaal': 'F%d' % (rn + 2)}


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


def maak_controle(wb, oud=None, bestand='', inv_laatste=25):
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
    cf = "Invoer!$%s$8:$%s$" + str(inv_laatste)
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
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 60
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['C'].width = 120
    return ws


def bouw(path, demo, oud_cashflow, oud_bestand, pew_inhoud, oz, inv_tot, daeb_type):
    wb = Workbook()
    wb.remove(wb.active)
    R = verwijzingen(oz, inv_tot, daeb_type)
    info = maak_start(wb, demo, R)
    info['pew_extra'] = maak_pew(wb, pew_inhoud, oz['pew'])
    info['vt_totaal'] = maak_verkoop_transport(wb, oz)
    info['sap'] = maak_sap(wb, oz)
    maak_vorige_prognose(wb, oud_cashflow, 'uit ' + oud_bestand)
    maak_controle(wb, oud_cashflow, oud_bestand, R['inv_laatste'])
    info['stijlen'] = maak_stijlen(wb)
    maak_sjabloon(wb)
    h = wb.create_sheet('_koppeling')
    h.sheet_state = 'hidden'
    wb.save(path)
    return info
