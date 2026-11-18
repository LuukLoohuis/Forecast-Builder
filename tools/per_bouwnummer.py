"""FO per bouwnummer: termijnverdeling (1-3) en planning (1-20) per bouwnummer.

Tabblad 'Planning bouwnummers':
  A · termijnverdelingen 1-3 (per verdeling max. 10 termijnen: naam, % van aanneemsom + KAO)
  B · planningen 1-20 (naam, start bouw, oplevering, optioneel kw na start per termijnnummer)
  C · overzicht per planning
Tabblad 'Bouwnummers': per bouwnummer de keuze verdeling / planning / verkocht in / transport in, met hulpkolommen
(plan- en factuurkwartaal en bedrag per termijnnummer). CF - opbrengsten en CF - kosten tellen die op met SUMIFS.

Kwartaalindex = jaar x 4 + kwartaal (zoals overal in het werkbestand); 0 = 't/m nu' (al verkocht / getransporteerd).
"""
from __future__ import annotations

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

SH_PL = 'Planning bouwnummers'
NV, NP, NJ = 3, 20, 10                    # verdelingen, planningen, termijnen per verdeling
# blok A: rij 5..14 = termijn 1..10; per verdeling twee kolommen (naam, %): C/D, F/G, I/J; rij 15 totaal, 16 status
A_KOP, A1 = 4, 5
A_TOT = A1 + NJ
A_NAAM = {1: 'C', 2: 'F', 3: 'I'}
A_PCT = {1: 'D', 2: 'G', 3: 'J'}
# blok B: rij 21..40 = planning 1..20
B_KOP, B1 = 20, 21
B_COL = {'nr': 'B', 'naam': 'C', 'start': 'D', 'opl': 'E', 'kw1': 6}       # kw termijn j in kolom 6 + j - 1 (F..O)
B_IDX_S, B_IDX_O = 'Q', 'R'                                                  # hulp: kwartaalindex start / oplevering
A_N = 'S'                                                                    # hulp: aantal termijnen per verdeling (rij A1..)
# blok C: rij 45..64 = overzicht planning 1..20
C_KOP, C1 = 44, 45

# Bouwnummers: kolommen W.. (de prijslijst-kolommen A..V blijven)
BN = {'verdeling': 'W', 'planning': 'X', 'verkocht': 'Y', 'transport': 'Z', 'v_idx': 'AA', 't_idx': 'AB', 's_idx': 'AC',
      'o_idx': 'AD', 'n': 'AE', 'kmw25': 'AF', 'kmw75': 'AG', 'melding': 'AH'}
BN_PLAN1 = 35                             # AI: plan-kwartaal termijn 1 .. (10 kolommen)
BN_FACT1 = BN_PLAN1 + NJ                  # factuurkwartaal (max plan, transport)
BN_EUR1 = BN_FACT1 + NJ                   # bedrag (aanneemsom + KAO) x %
BN_LAATST = BN_EUR1 + NJ - 1

GEEL = PatternFill('solid', fgColor='FFFFF2CC')
GRIJS = PatternFill('solid', fgColor='FFF3F4F6')
KOP = PatternFill('solid', fgColor='FF003366')
LICHT = PatternFill('solid', fgColor='FFE9ECEF')
F = Font(name='Heijmans-Light', size=10)
FB = Font(name='Heijmans-Light', size=10, bold=True)
FW = Font(name='Heijmans-BoldCon', size=10, color='FFFFFFFF')
FK = Font(name='Heijmans-Light', size=8, color='FF6B7280')
FT = Font(name='Heijmans-BoldCon', size=14, color='FF003366')
NUM = '#,##0;-#,##0;"-"'


def kw_index(cel):
    t = 'UPPER(TRIM(%s&""))' % cel
    return ('IFERROR(IF(LEFT(%s,1)="Q",VALUE(RIGHT(%s,4))*4+VALUE(MID(%s,2,1)),VALUE(LEFT(%s,4))*4+VALUE(RIGHT(%s,1))),"?")'
            % (t, t, t, t, t))


def kw_tekst(idx):
    return 'IF(N(%s)>0,"Q"&(MOD(%s-1,4)+1)&" "&INT((%s-1)/4),"")' % (idx, idx, idx)


def zet(ws, ref, v, font=F, fill=None, fmt=None, align=None):
    c = ws[ref]
    c.value = v
    c.font = font
    if fill is not None:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    return c


def band(ws, r, c1, c2, tekst):
    for cc in range(c1, c2 + 1):
        ws.cell(r, cc).fill = KOP
    zet(ws, '%s%d' % (L(c1), r), tekst, FW, KOP)


def maak_planning(ws, bn, bn1, bnn, status_kol='I', von_kol='M', grond_kol='K', aanneem_kol='L', nu='$U$2'):
    """tabblad Planning bouwnummers. bn = naam van het tabblad Bouwnummers, bn1..bnn = de bouwnummerrijen,
    nu = cel met de kwartaalindex van het huidige kwartaal"""
    ws.sheet_view.showGridLines = False
    for c, w in (('A', 3), ('B', 9), ('C', 30), ('D', 11), ('E', 11)):
        ws.column_dimensions[c].width = w
    for cc in range(6, 18):
        ws.column_dimensions[L(cc)].width = 11
    for c in ('F', 'I'):
        ws.column_dimensions[c].width = 30
    zet(ws, 'B1', 'Planning bouwnummers', FT)
    zet(ws, 'B2', 'Per bouwnummer kies je op Bouwnummers een termijnverdeling (A) en een planning (B).', FK)
    # A termijnverdelingen
    band(ws, A_KOP - 1, 2, 10, 'A · TERMIJNVERDELINGEN')
    zet(ws, 'B%d' % A_KOP, 'Termijn', FB, LICHT)
    for v in range(1, NV + 1):
        zet(ws, '%s%d' % (A_NAAM[v], A_KOP), 'Verdeling %d · termijn' % v, FB, LICHT)
        zet(ws, '%s%d' % (A_PCT[v], A_KOP), '%', FB, LICHT)
    for j in range(1, NJ + 1):
        r = A1 + j - 1
        zet(ws, 'B%d' % r, j, FK)
        for v in range(1, NV + 1):
            zet(ws, '%s%d' % (A_NAAM[v], r), None, F, GEEL)
            zet(ws, '%s%d' % (A_PCT[v], r), None, F, GEEL, '0.0%')
    zet(ws, 'B%d' % A_TOT, 'Totaal', FB, LICHT)
    for v in range(1, NV + 1):
        p, nm = A_PCT[v], A_NAAM[v]
        zet(ws, '%s%d' % (nm, A_TOT), '=IF(COUNTIF(%s%d:%s%d,"?*")=0,"",IF(ABS(SUM(%s%d:%s%d)-1)<0.0001,"OK","LET OP: samen "&TEXT(SUM(%s%d:%s%d),"0.0%%")))'
            % (nm, A1, nm, A1 + NJ - 1, p, A1, p, A1 + NJ - 1, p, A1, p, A1 + NJ - 1), FB, LICHT)
        zet(ws, '%s%d' % (p, A_TOT), '=SUM(%s%d:%s%d)' % (p, A1, p, A1 + NJ - 1), FB, LICHT, '0.0%')
        # hulp: aantal termijnen (rijen met een naam) per verdeling
        zet(ws, '%s%d' % (A_N, A1 + v - 1), '=COUNTIF(%s%d:%s%d,"?*")' % (nm, A1, nm, A1 + NJ - 1), FK)
    zet(ws, '%s%d' % (A_N, A_KOP), 'aantal termijnen (hulp)', FK)
    # B planningen
    band(ws, B_KOP - 1, 2, 15, 'B · PLANNINGEN · start bouw en oplevering (bv. Q1 2027); kw na start per termijn leeg = gelijk verdeeld')
    for c, t in (('B', 'Planning'), ('C', 'Naam (bv. verdieping)'), ('D', 'Start bouw'), ('E', 'Oplevering')):
        zet(ws, '%s%d' % (c, B_KOP), t, FB, LICHT)
    for j in range(1, NJ + 1):
        zet(ws, '%s%d' % (L(B_COL['kw1'] + j - 1), B_KOP), 'kw termijn %d' % j, FB, LICHT, align=Alignment(horizontal='center'))
    zet(ws, '%s%d' % (B_IDX_S, B_KOP), 'start (idx)', FK)
    zet(ws, '%s%d' % (B_IDX_O, B_KOP), 'opl. (idx)', FK)
    for p in range(1, NP + 1):
        r = B1 + p - 1
        zet(ws, 'B%d' % r, p, FB)
        zet(ws, 'C%d' % r, None, F, GEEL)
        zet(ws, 'D%d' % r, None, F, GEEL, align=Alignment(horizontal='center'))
        zet(ws, 'E%d' % r, None, F, GEEL, align=Alignment(horizontal='center'))
        for j in range(1, NJ + 1):
            zet(ws, '%s%d' % (L(B_COL['kw1'] + j - 1), r), None, F, GEEL, '0', Alignment(horizontal='center'))
        zet(ws, '%s%d' % (B_IDX_S, r), '=IF(TRIM(D%d&"")="","",%s)' % (r, kw_index('D%d' % r)), FK)
        zet(ws, '%s%d' % (B_IDX_O, r), '=IF(TRIM(E%d&"")="","",%s)' % (r, kw_index('E%d' % r)), FK)
    # C overzicht per planning
    band(ws, C_KOP - 1, 2, 13, 'C · OVERZICHT PER PLANNING')
    kop = ['Planning', 'Naam', 'Aantal', 'VON min', 'VON max', 'Grond', 'Aanneemsom', 'Verkocht t/m nu', 'Getransp. t/m nu',
           'Zonder verkoopkw.', 'Status']
    for i, t in enumerate(kop):
        zet(ws, '%s%d' % (L(2 + i), C_KOP), t, FB, LICHT)
    rng = lambda c: "%s!$%s$%d:$%s$%d" % (bn, c, bn1, c, bnn)
    pl = rng(BN['planning'])
    for p in range(1, NP + 1):
        r = C1 + p - 1
        zet(ws, 'B%d' % r, p, FB)
        zet(ws, 'C%d' % r, '=C%d&""' % (B1 + p - 1))
        zet(ws, 'D%d' % r, '=COUNTIF(%s,%d)' % (pl, p), FB, None, '0')
        zet(ws, 'E%d' % r, '=IF(D%d=0,"",_xlfn.MINIFS(%s,%s,%d))' % (r, rng(von_kol), pl, p), F, None, NUM)
        zet(ws, 'F%d' % r, '=IF(D%d=0,"",_xlfn.MAXIFS(%s,%s,%d))' % (r, rng(von_kol), pl, p), F, None, NUM)
        zet(ws, 'G%d' % r, '=SUMIF(%s,%d,%s)' % (pl, p, rng(grond_kol)), F, None, NUM)
        zet(ws, 'H%d' % r, '=SUMIF(%s,%d,%s)' % (pl, p, rng(aanneem_kol)), F, None, NUM)
        zet(ws, 'I%d' % r, '=COUNTIFS(%s,%d,%s,"<="&(%s-1))' % (pl, p, rng(BN['v_idx']), nu), F, None, '0')
        zet(ws, 'J%d' % r, '=COUNTIFS(%s,%d,%s,"<="&(%s-1))' % (pl, p, rng(BN['t_idx']), nu), F, None, '0')
        zet(ws, 'K%d' % r, '=COUNTIFS(%s,%d,%s,"")' % (pl, p, rng(BN['v_idx'])), F, None, '0;-0;"-"')
        zet(ws, 'L%d' % r, '=IF(D%d=0,"",IF(%s%d="","FOUT: geen start bouw",IF(K%d>0,"LET OP: "&K%d&" zonder verkoopkwartaal","OK")))'
            % (r, B_IDX_S, B1 + p - 1, r, r), FB)
    rood = Font(bold=True, color='FFB91C1C')
    ws.conditional_formatting.add('L%d:L%d' % (C1, C1 + NP - 1), FormulaRule(formula=['LEFT(L%d,2)<>"OK"' % C1], font=rood))
    ws.conditional_formatting.add('L%d:L%d' % (C1, C1 + NP - 1), FormulaRule(formula=['L%d="OK"' % C1], font=Font(bold=True, color='FF15803D')))
    for v in range(1, NV + 1):
        c = '%s%d' % (A_NAAM[v], A_TOT)
        ws.conditional_formatting.add(c, FormulaRule(formula=['LEFT(%s,6)="LET OP"' % c], font=rood))
    ws.freeze_panes = 'B3'


def maak_bouwnummer_kolommen(ws, bn1, bnn, nu='$U$2', status_kol='I', grond_kol='K', aanneem_kol='L', kao_kol='N',
                             pl=SH_PL):
    """invoer- en hulpkolommen op Bouwnummers (rij bn1..bnn); kop op rij bn1 - 1"""
    kr = bn1 - 1
    P = "'%s'!" % pl
    for c, t in (('verdeling', 'Verdeling (1-3)'), ('planning', 'Planning (1-20)'), ('verkocht', 'Verkocht in'),
                 ('transport', 'Transport in')):
        zet(ws, '%s%d' % (BN[c], kr), t, FW, KOP, align=Alignment(horizontal='center', wrap_text=True))
        ws.column_dimensions[BN[c]].width = 11
    hulp = [('v_idx', 'verkocht (idx)'), ('t_idx', 'transport (idx)'), ('s_idx', 'start (idx)'), ('o_idx', 'opl. (idx)'),
            ('n', 'termijnen'), ('kmw25', 'kmw 25% (idx)'), ('kmw75', 'kmw 75% (idx)'), ('melding', 'Melding')]
    for c, t in hulp:
        zet(ws, '%s%d' % (BN[c], kr), t, FK, LICHT)
    for j in range(1, NJ + 1):
        zet(ws, '%s%d' % (L(BN_PLAN1 + j - 1), kr), 'plan kw %d' % j, FK, LICHT)
        zet(ws, '%s%d' % (L(BN_FACT1 + j - 1), kr), 'factuur kw %d' % j, FK, LICHT)
        zet(ws, '%s%d' % (L(BN_EUR1 + j - 1), kr), '€ termijn %d' % j, FK, LICHT)
    dv_v = DataValidation(type='whole', operator='between', formula1='1', formula2=str(NV), allow_blank=True)
    dv_p = DataValidation(type='whole', operator='between', formula1='1', formula2=str(NP), allow_blank=True)
    ws.add_data_validation(dv_v)
    ws.add_data_validation(dv_p)
    dv_v.add('%s%d:%s%d' % (BN['verdeling'], bn1, BN['verdeling'], bnn))
    dv_p.add('%s%d:%s%d' % (BN['planning'], bn1, BN['planning'], bnn))
    for r in range(bn1, bnn + 1):
        b = lambda c: '%s%d' % (BN[c], r)
        heeft = 'A%d<>""' % r
        st = '%s%d&""' % (status_kol, r)
        for c in ('verdeling', 'planning', 'verkocht', 'transport'):
            zet(ws, b(c), None, F, GEEL, align=Alignment(horizontal='center'))
        # verkocht / transport: kwartaal, of 0 = t/m nu (al verkocht / getransporteerd volgens Salesforce)
        zet(ws, b('v_idx'), '=IF(NOT(%s),"",IF(TRIM(%s&"")<>"",%s,IF(OR(%s="Verkocht",LEFT(%s,6)="Getran"),0,"")))'
            % (heeft, b('verkocht'), kw_index(b('verkocht')), st, st), FK)
        zet(ws, b('s_idx'), '=IF(N(%s)=0,"",INDEX(%s$%s$%d:$%s$%d,%s))' % (b('planning'), P, B_IDX_S, B1, B_IDX_S, B1 + NP - 1, b('planning')), FK)
        zet(ws, b('o_idx'), '=IF(N(%s)=0,"",INDEX(%s$%s$%d:$%s$%d,%s))' % (b('planning'), P, B_IDX_O, B1, B_IDX_O, B1 + NP - 1, b('planning')), FK)
        zet(ws, b('n'), '=IF(NOT(%s),"",INDEX(%s$%s$%d:$%s$%d,MAX(1,N(%s))))' % (heeft, P, A_N, A1, A_N, A1 + NV - 1, b('verdeling')), FK)
        zet(ws, b('t_idx'), ('=IF(NOT({h}),"",IF(TRIM({t}&"")<>"",{ti},IF(LEFT({st},6)="Getran",0,IF({v}="","",'
                             'MAX(IF({v}=0,{nu},{v}),N({s}))))))').format(h=heeft, t=b('transport'), ti=kw_index(b('transport')),
                                                                   st=st, v=b('v_idx'), s=b('s_idx'), nu=nu), FK)
        zet(ws, b('kmw25'), '=IF(OR(%s="",%s=""),"",MAX(%s,%s))' % (b('t_idx'), b('s_idx'), b('t_idx'), b('s_idx')), FK)
        zet(ws, b('kmw75'), '=IF(OR(%s="",%s=""),"",MAX(%s,%s))' % (b('t_idx'), b('o_idx'), b('t_idx'), b('o_idx')), FK)
        zet(ws, b('melding'), ('=IF(NOT({h}),"",IF(N({p})=0,"FOUT: kies een planning",IF({s}="","FOUT: planning zonder start bouw",'
                               'IF(OR({v}="?",{t}="?"),"FOUT: kwartaal niet herkend (typ bv. Q2 2027)",IF({v}="","LET OP: geen verkoopkwartaal","")))))')
            .format(h=heeft, p=b('planning'), s=b('s_idx'), v=b('v_idx'), t=b('t_idx')), Font(name='Heijmans-Light', size=8, bold=True, color='FFB91C1C'))
        for j in range(1, NJ + 1):
            pc, fc, ec = (L(x + j - 1) for x in (BN_PLAN1, BN_FACT1, BN_EUR1))
            kw = 'INDEX(%s$%s$%d:$%s$%d,%s)' % (P, L(B_COL['kw1'] + j - 1), B1, L(B_COL['kw1'] + j - 1), B1 + NP - 1, b('planning'))
            gelijk = 'IF(N(%s)<=1,0,ROUND((%d-1)*(N(%s)-N(%s))/(N(%s)-1),0))' % (b('n'), j, b('o_idx'), b('s_idx'), b('n'))
            zet(ws, '%s%d' % (pc, r), '=IF(OR(%s="",%d>N(%s)),"",N(%s)+IF(%s&""<>"",N(%s),IF(%s="",0,%s)))'
                % (b('s_idx'), j, b('n'), b('s_idx'), kw, kw, b('o_idx'), gelijk), FK)
            zet(ws, '%s%d' % (fc, r), '=IF(OR(%s%d="",%s=""),"",MAX(%s%d,%s))' % (pc, r, b('t_idx'), pc, r, b('t_idx')), FK)
            pct = 'INDEX(CHOOSE(MAX(1,N(%s)),%s$%s$%d:$%s$%d,%s$%s$%d:$%s$%d,%s$%s$%d:$%s$%d),%d)' % (
                b('verdeling'), P, A_PCT[1], A1, A_PCT[1], A1 + NJ - 1, P, A_PCT[2], A1, A_PCT[2], A1 + NJ - 1,
                P, A_PCT[3], A1, A_PCT[3], A1 + NJ - 1, j)
            zet(ws, '%s%d' % (ec, r), '=IF(%s%d="",0,(N(%s%d)+N(%s%d))*N(%s))' % (pc, r, aanneem_kol, r, kao_kol, r, pct), FK, None, NUM)
    rood = Font(bold=True, color='FFB91C1C')
    ws.conditional_formatting.add('%s%d:%s%d' % (BN['melding'], bn1, BN['melding'], bnn),
                                  FormulaRule(formula=['LEN(%s%d)>0' % (BN['melding'], bn1)], font=rood))
    # hulpkolommen ingeklapt
    ws.column_dimensions.group(BN['v_idx'], BN['o_idx'], outline_level=1, hidden=False)
    ws.column_dimensions.group(L(BN_PLAN1), L(BN_LAATST), outline_level=1, hidden=True)
    ws.column_dimensions[BN['melding']].width = 34


def cf_formules(bn, bn1, bnn, q_cel, eerste_q):
    """formules voor één kwartaalkolom van 'Alle bouwnummers' op CF - opbrengsten. q_cel = cel met de kwartaalindex van de
    kolom (of None voor de kolom 't/m nu': alles vóór eerste_q). Geeft dict: grond, termijn[j], kmw25, kmw75, verkocht, transport"""
    rng = lambda c: "%s!$%s$%d:$%s$%d" % (bn, c, bn1, c, bnn)
    crit = ('"<"&%s' % eerste_q) if q_cel is None else q_cel
    sf = lambda som, krit: 'SUMIFS(%s,%s,%s)' % (rng(som), rng(krit), crit)
    uit = {'grond': '=' + sf('K', BN['t_idx']),
           'kad': '=' + sf('J', BN['t_idx']),
           'kmw25': '=0.25*' + sf('P', BN['kmw25']),
           'kmw75': '=0.75*' + sf('P', BN['kmw75']),
           'verkocht': '=COUNTIFS(%s,%s)' % (rng(BN['v_idx']), crit),
           'transport': '=COUNTIFS(%s,%s)' % (rng(BN['t_idx']), crit),
           'termijn': {}, 'kosten': {}}
    for j in range(1, NJ + 1):
        uit['termijn'][j] = '=' + sf(L(BN_EUR1 + j - 1), L(BN_FACT1 + j - 1))
        uit['kosten'][j] = sf(L(BN_EUR1 + j - 1), L(BN_PLAN1 + j - 1))
    return uit
