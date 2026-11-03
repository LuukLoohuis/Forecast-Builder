"""Het standaard-FO: één vaste, schone opzet voor elk project (het sjabloon van het werkbestand).

    python tools/standaard_fo.py uit.xlsx            leeg standaard-FO (exports leeg)

Vaste plekken (ook gebruikt door omzetten.py, de bouw en de macro Oud FO inlezen):
  CF - opbrengsten  kolom G = actuals (t/m vóór het eerste kwartaal), H..AU = 40 kwartalen vanaf Q1 van het startjaar (D3)
                    per woningtype k (1..6) een blok van 16 rijen vanaf rij WON(k): # won, termijnen (titel + 10
                    planningsrijen), transportregel, verkocht %, getransporteerd %, grondprijs
                    verkooptempo, omzetblokken (aantallen), financieel (€), overige opbrengsten, totaal, DAEB-blok
  CF - kosten       per woningtype de bouwtermijnen kosten (eigen planning, standaard = opbrengsten), daarna de posten
  Prijslijst        per woningtype (uit de prijslijst of handmatig), aansluiting met SKO/PEW per type + termijnschema,
                    per bouwnummer (btw eraf, uitgesplitst)
De oude tabbladen Aansluiting opbrengsten, Verkooptempo, Woningtypes (als invoer), 3. Restant bdgt en de verborgen
FO-bladen bestaan hier niet.
"""
from __future__ import annotations

import os
import sys

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO = os.path.join(ROOT, 'input', 'FO_Parkzoom_fictieve_demoversie.xlsx')

NT, NTERM, NQ = 6, 10, 40
C_ACT, C_Q1 = 7, 8                     # G = actuals, H = eerste kwartaal
C_QN = C_Q1 + NQ - 1                   # AU
C_CHK = C_QN + 2                       # AW
COLS = [L(c) for c in range(C_ACT, C_QN + 1)]          # G..AU
QCOLS = COLS[1:]
CHK, VER, STAT, TOEL = L(C_CHK), L(C_CHK + 1), L(C_CHK + 2), L(C_CHK + 3)
JAAR = "'CF - opbrengsten'!$D$3"

# ---- CF - opbrengsten -----------------------------------------------------------------------------
BLOK = NTERM + 6                        # # won, titel, 10 planning, transport, verkocht %, getransp. %, grondprijs


def won_rij(k):
    return 5 + BLOK * (k - 1)


def hb(k):                              # titelrij 'Termijnen k - naam'
    return won_rij(k) + 1


def plan(k, t):                         # t = 1..NTERM
    return hb(k) + t


def tr(k):
    return hb(k) + NTERM + 1


VK = won_rij(NT + 1) + 2                # koprij verkooptempo
def vk_rij(k):
    return VK + k


OMZ_N = NTERM + 4                       # grond, termijnen, kmw 25%, kmw 75%, kadastrale kosten + rente
OBLOK = OMZ_N + 3
VT_PCT = vk_rij(NT) + 2                # alle woningen: verkocht %/kw, cumulatief, getransporteerd %/kw, cumulatief
OZ1 = vk_rij(NT) + 8                    # eerste omzet-koprij (# won op de rij erboven)


def oz_kop(k):
    return OZ1 + OBLOK * (k - 1)


def oz_rij(k, i):                       # i = 0 grond, 1..NTERM termijnen, NTERM+1 kmw 25, +2 kmw 75, +3 kad/rente
    return oz_kop(k) + 1 + i


EU1 = oz_kop(NT) + OBLOK + 3            # eerste 'Termijnen 1 in €'


def eu_kop(k):
    return EU1 + OBLOK * (k - 1)


def eu_rij(k, i):
    return eu_kop(k) + 1 + i


OV_KOP = eu_kop(NT) + OBLOK + 1         # 'Overige opbrengsten in €'
OV1, OV2 = OV_KOP + 1, OV_KOP + 2
CORR = OV2 + 2
TOT = OV2 + 3
PCT = TOT + 1
PER = TOT + 2
DAEB0 = PER + 4                          # notitie; componenten D0+2, D0+3; jaar D0+4; koprij D0+5
DAEB_KOP = DAEB0 + 5
DAEB_TOT = DAEB_KOP + 19
DAEB_AANTAL, DAEB_VERKOOP = DAEB_TOT + 2, DAEB_TOT + 3

# ---- CF - kosten ------------------------------------------------------------------------------------
KBLOK = NTERM + 7                        # budget, kop, 10 termijnen, kmw 25, kmw 75, aanneemsom, meerwerk, leeg


def kb(k):
    return 3 + KBLOK * (k - 1)


def kterm(k, t):                         # t = 1..NTERM, NTERM+1 = kmw 25%, NTERM+2 = kmw 75%
    return kb(k) + 1 + t


def k_aanneem(k):
    return kb(k) + NTERM + 4


K_PEW = kb(NT) + KBLOK + 2              # 'Totaal kosten' (PEW)
K_JAAR = K_PEW + 2
K_OND = K_PEW + 3
K_POST1 = K_OND + 1
NPOST = 12
K_POSTN = K_POST1 + NPOST - 1
K_SOM = K_POSTN + 1
K_CHECK = K_POSTN + 2
K_TOT = K_POSTN + 5
K_PER = K_TOT + 2
POSTEN = ['Grondkosten', 'Aanneemsom', 'Kopersmeerwerk', 'Overige bouwkosten / meerwerken', 'Honoraria',
          'Promotiekosten', 'Verhuur & verkoopkosten', 'Aansluitkosten', 'Legeskosten', 'Rente / Subsidies / bankgarantie',
          'Commerciele / Operationele reserve', 'Direct doorbelaste projectkosten']

# ---- Prijslijst --------------------------------------------------------------------------------------
P_TYPE1 = 5                              # rij type k = P_TYPE1 + k - 1
P_REGELS = ['Grond + Ontwikkelkosten', 'Aanneemsom', 'VON verhogingen', 'KAO opties', 'Korting', 'Kopersmeerwerk',
            'Kadastrale kosten', 'Bouwrente', 'Aanvullende rente', 'Opbrengsten commercieel', 'Opbrengsten parkeren']
P_BLOK = len(P_REGELS) + 4
P_AANS1 = P_TYPE1 + NT + 4


def p_aans(k):                           # titel 'Vergelijking excl. BTW - type k'; kop +1; regels +2..
    return P_AANS1 + P_BLOK * (k - 1)


P_SAP = p_aans(NT) + P_BLOK            # ruimte voor 'te boeken in SAP' (de bouw zet het hier)
P_BN_KOP = P_SAP + 110                  # kop per bouwnummer
P_BN1, P_BNN = P_BN_KOP + 1, P_BN_KOP + 500
# per type: kolommen 'uit prijslijst', 'handmatig' (overschrijft), 'gebruikt'
P_UIT = {'aantal': 'D', 'grond': 'E', 'aanneemsom': 'F', 'kao': 'G', 'kmw': 'H', 'kad': 'I', 'bouwrente': 'J',
         'rente': 'K', 'korting': 'L', 'von': 'M'}
P_HAND = {'aantal': 'N', 'grond': 'O', 'aanneemkao': 'P', 'kmw': 'Q', 'kadrente': 'R'}
P_GEBR = {'aantal': 'S', 'grond': 'T', 'aanneemkao': 'U', 'kmw': 'V', 'kadrente': 'W'}
P_TERM_NAAM, P_TERM_PCT = 'L', 'M'      # termijnschema naast het aansluitblok
P_BN = {'type': 'C', 'status': 'I', 'kad': 'J', 'grond': 'K', 'aanneemsom': 'L', 'von': 'M', 'kao': 'N', 'kmw': 'P',
        'bouwrente': 'Q', 'rente': 'R', 'korting': 'S'}

GEEL = PatternFill('solid', fgColor='FFFFF2CC')
GRIJS = PatternFill('solid', fgColor='FFF3F4F6')
KOP = PatternFill('solid', fgColor='FFFEB900')
DONKER = PatternFill('solid', fgColor='FF003366')
F = Font(name='Heijmans-Light', size=10)
FB = Font(name='Heijmans-Light', size=10, bold=True)
FT = Font(name='Heijmans-BoldCon', size=14, color='FF003366')
FW = Font(name='Heijmans-BoldCon', size=10, color='FFFFFFFF')
FN = Font(name='Heijmans-Light', size=9, italic=True, color='FF6B7280')
NUM = '#,##0;-#,##0;"-"'
AANT = 'General;-General;"-"'          # aantallen: geen punt of komma achter een heel getal
NUM2 = '#,##0.00;-#,##0.00;"-"'


def zet(ws, ref, v, font=F, fill=None, fmt=None):
    c = ws[ref]
    c.value = v
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def kopregels(ws, r_jaar, r_q, act='actuals', jaar=JAAR):
    """jaartallen (op Q1) en Q1..Q4 boven de kolommen"""
    zet(ws, '%s%d' % (COLS[0], r_q), act, FB, KOP)
    for i, c in enumerate(QCOLS):
        # jaartal boven elk kwartaal (Q1 vet), zodat het ook na het inklappen van eerdere kwartalen zichtbaar is
        zet(ws, '%s%d' % (c, r_jaar), '=%s+%d' % (jaar, i // 4), FB if i % 4 == 0 else Font(name='Heijmans-Light', size=8, color='FF9CA3AF'))
        zet(ws, '%s%d' % (c, r_q), 'Q%d' % (i % 4 + 1), FB, KOP)


def markeringen(ws, r_mark, r1, r2):
    """rij r_mark: 'Start bouw' / 'Oplevering' boven het kwartaal (Start, rij 7 en 9); in r1..r2 de kolom gekleurd en
    een rand aan het begin van elk jaar (vakken per jaar)"""
    key = lambda c: '(%s*4+1+COLUMN(%s%d)-COLUMN($%s$%d))' % (JAAR, c, r1, QCOLS[0], r1)
    for c in QCOLS:
        zet(ws, '%s%d' % (c, r_mark), '=IF(%s=p_startbouw,"Start bouw",IF(%s=p_oplevering,"Oplevering",""))'
            % (key(c), key(c)), Font(name='Heijmans-BoldCon', size=9, color='FF003366'))
    zet(ws, 'B%d' % r_mark, '="Start bouw "&Start!$E$7&" · oplevering "&Start!$E$9', FN)
    rng = '%s%d:%s%d' % (QCOLS[0], r1, QCOLS[-1], r2)
    k0 = key(QCOLS[0]).replace('COLUMN(%s%d)' % (QCOLS[0], r1), 'COLUMN(%s%d)' % (QCOLS[0], r1))
    groen = PatternFill('solid', fgColor='FFE3F2E1', bgColor='FFE3F2E1')
    blauw = PatternFill('solid', fgColor='FFDCE8F5', bgColor='FFDCE8F5')
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s=p_startbouw' % k0], fill=groen))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s=p_oplevering' % k0], fill=blauw))
    rand = Side(style='thin', color='FF003366')
    ws.conditional_formatting.add(rng, FormulaRule(formula=['MOD(COLUMN(%s%d)-COLUMN($%s$%d),4)=0' % (QCOLS[0], r1, QCOLS[0], r1)],
                                                   border=Border(left=rand)))
    ws.conditional_formatting.add('%s%d:%s%d' % (COLS[0], r1, COLS[0], r2), FormulaRule(formula=['TRUE'], border=Border(right=rand)))


def pl(ref):
    return 'Prijslijst!' + ref


def p_gebr(k, x):
    return pl('$%s$%d' % (P_GEBR[x], P_TYPE1 + k - 1))


# ===================================================================================================
def maak_opbrengsten(ws):
    ws.title = 'CF - opbrengsten'
    zet(ws, 'B2', 'Bouwplanning, verkoop en transport', FT)
    zet(ws, 'B3', 'Eerste jaar van de kwartalen (Q1)', FB)
    zet(ws, 'D3', None, FB, GEEL, '0')
    for k in range(1, NT + 1):
        w, h = won_rij(k), hb(k)
        zet(ws, 'F%d' % w, '# won', FB)
        zet(ws, 'G%d' % w, '=%s' % p_gebr(k, 'aantal'), FB, None, '0')
        kopregels(ws, w, h)
        zet(ws, 'B%d' % h, '="Termijnen %d - "&%s' % (k, pl('$B$%d' % (P_TYPE1 + k - 1))), FB, KOP)
        for t in range(1, NTERM + 1):
            r = plan(k, t)
            zet(ws, 'B%d' % r, '=B%d' % oz_rij(k, t))
            for c in COLS:
                zet(ws, '%s%d' % (c, r), None, F, GEEL, AANT)
            zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
            zet(ws, '%s%d' % (VER, r), '=IF(B%d="","",ROUND($G$%d-%s%d,2))' % (r, w, CHK, r), F, None, AANT)
            zet(ws, '%s%d' % (STAT, r), '=IF(%s%d="","",IF(%s%d=0,"OK","FOUT"))' % (VER, r, VER, r), FB)
        zet(ws, '%s%d' % (CHK, h), 'Check', FB, KOP)
        zet(ws, '%s%d' % (VER, h), 'Verschil', FB, KOP)
    # verkooptempo (aantallen; komen in het werkbestand uit verkocht %)
    kopregels(ws, VK - 1, VK)
    for c, t in zip('BCDE', ['Verkooptempo', 'Aantal', 'Verkocht', 'Te koop']):
        zet(ws, '%s%d' % (c, VK), t, FB, KOP)
    zet(ws, '%s%d' % (CHK, VK), 'Check', FB, KOP)
    n_bn = '$%s$%d:$%s$%d' % (P_BN['type'], P_BN1, P_BN['type'], P_BNN)
    s_bn = '$%s$%d:$%s$%d' % (P_BN['status'], P_BN1, P_BN['status'], P_BNN)
    for k in range(1, NT + 1):
        r = vk_rij(k)
        sl = pl('$C$%d' % (P_TYPE1 + k - 1))
        zet(ws, 'B%d' % r, '=%s&""' % pl('$B$%d' % (P_TYPE1 + k - 1)))
        zet(ws, 'C%d' % r, '=G%d' % won_rij(k), F, None, '0')
        zet(ws, 'D%d' % r, '=IF(%s="",0,COUNTIFS(Prijslijst!%s,%s,Prijslijst!%s,"Verkocht")+COUNTIFS(Prijslijst!%s,%s,'
                           'Prijslijst!%s,"Getransporteerd"))' % (sl, n_bn, sl, s_bn, n_bn, sl, s_bn), F, GRIJS, '0')
        zet(ws, 'E%d' % r, '=C%d-D%d' % (r, r), F, None, '0')
        zet(ws, 'G%d' % r, '=D%d' % r, F, GRIJS, '0')
        for c in QCOLS:
            zet(ws, '%s%d' % (c, r), None, F, GEEL, AANT)
        zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
        zet(ws, '%s%d' % (VER, r), '=ROUND(C%d-%s%d,2)' % (r, CHK, r), F, None, AANT)
    # omzetblokken (aantallen woningen per termijn per kwartaal)
    for k in range(1, NT + 1):
        o = oz_kop(k)
        w = '$G$%d' % (o - 1)
        zet(ws, 'F%d' % (o - 1), '# won', FB)
        zet(ws, 'G%d' % (o - 1), '=G%d' % won_rij(k), FB, None, '0')
        kopregels(ws, o - 1, o)
        zet(ws, 'B%d' % o, '="Omzet HVG "&B%d' % hb(k), FB, KOP)
        for c, t in zip('CDE', ['%', '€', '€ per wn']):
            zet(ws, '%s%d' % (c, o), t, FB, KOP)
        v = vk_rij(k)
        a = P_AANS1 + P_BLOK * (k - 1)
        # grond (= transport)
        r = oz_rij(k, 0)
        zet(ws, 'B%d' % r, 'Grondtermijn (bij transport)')
        zet(ws, 'D%d' % r, '=%s' % p_gebr(k, 'grond'), F, None, NUM)
        zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
        for c in COLS:
            zet(ws, '%s%d' % (c, r), None, F, None, AANT)
        for t in range(1, NTERM + 1):
            r = oz_rij(k, t)
            p = plan(k, t)
            zet(ws, 'B%d' % r, '=%s&""' % pl('$%s$%d' % (P_TERM_NAAM, a + 1 + t)))
            zet(ws, 'C%d' % r, '=N(%s)' % pl('$%s$%d' % (P_TERM_PCT, a + 1 + t)), F, None, '0.0%')
            zet(ws, 'D%d' % r, '=C%d*%s' % (r, p_gebr(k, 'aanneemkao')), F, None, NUM)
            zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
            for j, c in enumerate(COLS):
                if j == 0:
                    f = '=IF(%s%d<SUM($%s$%d),%s%d,SUM($%s$%d))' % (c, p, c, v, c, p, c, v)
                else:
                    f = '=IF(SUM($%s%d:%s%d)<SUM($%s$%d:%s$%d),%s%d,SUM($%s$%d:%s$%d)-SUM($%s%d:%s%d))' % (
                        COLS[0], p, c, p, COLS[0], v, c, v, c, p, COLS[0], v, c, v, COLS[0], r, COLS[j - 1], r)
                zet(ws, '%s%d' % (c, r), f, F, None, AANT)
        t1, tn = oz_rij(k, 1), oz_rij(k, NTERM)
        for i, (lab, pct) in enumerate((('Kopersmeerwerk 25% (bij start bouw)', 0.25), ('Kopersmeerwerk 75% (bij oplevering)', 0.75)), 1):
            r = oz_rij(k, NTERM + i)
            zet(ws, 'B%d' % r, lab)
            zet(ws, 'C%d' % r, pct, F, None, '0%')
            zet(ws, 'D%d' % r, '=C%d*%s' % (r, p_gebr(k, 'kmw')), F, None, NUM)
            zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
            if i == 2:      # laatste termijn met een % (oplevering)
                zet(ws, 'F%d' % r, '=IFERROR(LOOKUP(2,1/($C$%d:$C$%d>0),ROW($C$%d:$C$%d)-ROW($C$%d)+1),1)'
                    % (t1, tn, t1, tn, t1), F, None, '0')
            for c in COLS:
                f = '=%s%d' % (c, t1) if i == 1 else '=INDEX(%s$%d:%s$%d,$F%d)' % (c, t1, c, tn, r)
                zet(ws, '%s%d' % (c, r), f, F, None, AANT)
        r = oz_rij(k, NTERM + 3)
        zet(ws, 'B%d' % r, 'Kadastrale kosten en aanvullende rente (aantal woningen, handmatig)')
        zet(ws, 'D%d' % r, '=%s' % p_gebr(k, 'kadrente'), F, None, NUM)
        zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
        for c in COLS:
            zet(ws, '%s%d' % (c, r), None, F, GEEL, AANT)
        for i in range(OMZ_N):
            r = oz_rij(k, i)
            zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
    # financieel (€)
    zet(ws, 'B%d' % (EU1 - 3), 'Financieel (klap open)', FB)
    for k in range(1, NT + 1):
        e = eu_kop(k)
        kopregels(ws, e - 1, e, act='actuals')
        zet(ws, 'B%d' % e, 'Termijnen %d in €' % k, FB, KOP)
        for i in range(OMZ_N):
            r, o = eu_rij(k, i), oz_rij(k, i)
            zet(ws, 'B%d' % r, '=B%d' % o)
            for c in COLS:
                zet(ws, '%s%d' % (c, r), '=%s%d*$E%d' % (c, o, o), F, None, NUM)
    kopregels(ws, OV_KOP - 1, OV_KOP)
    zet(ws, 'B%d' % OV_KOP, 'Overige opbrengsten in €', FB, KOP)
    for r, lab in ((OV1, 'Opbrengsten commercieel'), (OV2, 'Opbrengsten parkeren')):
        zet(ws, 'B%d' % r, lab)
        for c in COLS:
            zet(ws, '%s%d' % (c, r), None, F, GEEL, NUM)
    zet(ws, 'B%d' % TOT, 'Totaal opbrengsten', FB)
    for c in COLS + [CHK]:
        delen = ['%s%d:%s%d' % (c, eu_rij(k, 0), c, eu_rij(k, OMZ_N - 1)) for k in range(1, NT + 1)]
        zet(ws, '%s%d' % (c, TOT), '=SUM(%s,%s%d:%s%d)+%s%d' % (','.join(delen), c, OV1, c, OV2, c, DAEB_TOT), FB, None, NUM)
    zet(ws, 'B%d' % PCT, 'Opbrengsten in %')
    for c in COLS:
        zet(ws, '%s%d' % (c, PCT), '=IFERROR(%s%d/$%s$%d,0)' % (c, TOT, CHK, TOT), F, None, '0%')
    zet(ws, 'B%d' % PER, 'Periode in cashflow SAP')
    zet(ws, 'D%d' % PER, 'Start:')
    zet(ws, 'E%d' % PER, None, F, GEEL, 'dd-mm-yyyy')
    for i, c in enumerate(QCOLS):
        # maandnummer (laatste maand van het kwartaal) vanaf de startmaand in E
        zet(ws, '%s%d' % (c, PER), '=IF(N($E$%d)=0,0,(%s+%d-YEAR($E$%d))*12+%d-MONTH($E$%d)+1)'
            % (PER, JAAR, i // 4, PER, (i % 4 + 1) * 3, PER), F, None, '0')
    zet(ws, '%s%d' % (COLS[0], PER), 0)
    # DAEB-blok (leeg; zelfde opzet als in het werkbestand)
    zet(ws, 'C%d' % DAEB0, 'DAEB-woningen (corporatie): fee-componenten (naam + totaal), bedragen per kwartaal, aantal en '
                           'verkoop. Leeg = geen DAEB.', FN)
    for i, r in enumerate((DAEB_KOP - 3, DAEB_KOP - 2), 1):
        zet(ws, 'B%d' % r, 'Fee-component %d · naam (C) en totaal € (E)' % i)
        zet(ws, 'C%d' % r, None, F, GEEL)
        zet(ws, 'E%d' % r, None, F, GEEL, NUM)
    kopregels(ws, DAEB_KOP - 1, DAEB_KOP, act='actuals')
    zet(ws, 'B%d' % DAEB_KOP, 'Opbrengsten DAEB', FB, KOP)
    for r in list(range(DAEB_KOP + 1, DAEB_KOP + 9)) + list(range(DAEB_KOP + 10, DAEB_KOP + 18)):
        for c in ['B', 'C', 'D'] + COLS:
            zet(ws, '%s%d' % (c, r), None, F, GEEL, NUM if c not in ('B', 'C') else ('0%' if c == 'C' else None))
        zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, NUM)
    zet(ws, 'B%d' % DAEB_TOT, 'Totaal opbrengsten DAEB', FB)
    for c in COLS + [CHK]:
        zet(ws, '%s%d' % (c, DAEB_TOT), '=SUM(%s%d:%s%d)' % (c, DAEB_KOP + 1, c, DAEB_TOT - 2), FB, None, NUM)
    zet(ws, 'B%d' % DAEB_AANTAL, 'Aantal DAEB-woningen')
    zet(ws, 'D%d' % DAEB_AANTAL, None, F, GEEL, '0')
    zet(ws, 'B%d' % DAEB_VERKOOP, 'Verkoop en levering DAEB (aantal woningen per kwartaal; transport = verkoop)')
    for c in COLS:
        zet(ws, '%s%d' % (c, DAEB_VERKOOP), None, F, GEEL, '0')
    # verkocht en getransporteerd van alle woningen (alle types samen), per kwartaal en cumulatief
    tw = '+'.join('$C$%d' % vk_rij(k) for k in range(1, NT + 1))
    labels = ['Alle woningen · verkocht % per kwartaal', 'Alle woningen · verkocht % cumulatief',
              'Alle woningen · getransporteerd % per kwartaal', 'Alle woningen · getransporteerd % cumulatief']
    for i, lab in enumerate(labels):
        r = VT_PCT + i
        zet(ws, 'B%d' % r, lab, FB if i % 2 else F)
        for j, c in enumerate(COLS):
            if i == 0:
                f = '=IFERROR((%s)/(%s),0)' % ('+'.join('%s%d' % (c, vk_rij(k)) for k in range(1, NT + 1)), tw)
            elif i == 2:
                f = '=IFERROR((%s)/(%s),0)' % ('+'.join('%s%d' % (c, tr(k)) for k in range(1, NT + 1)), tw)
            else:
                f = '=SUM($%s%d:%s%d)' % (COLS[0], r - 1, c, r - 1)
            zet(ws, '%s%d' % (c, r), f, FB if i % 2 else F, GRIJS, '0.0%;-0.0%;"-"')
    markeringen(ws, 4, 4, PER + 40)
    # inklappen: invoerblok per woningtype, en de rekenblokken (aantallen per termijn en €)
    for k in range(1, NT + 1):
        ws.row_dimensions.group(plan(k, 1), tr(k) + 3, outline_level=1)
    ws.row_dimensions.group(OZ1 - 1, OV_KOP - 2, outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.column_dimensions['B'].width = 46
    for c in COLS:
        ws.column_dimensions[c].width = 9
    ws.freeze_panes = 'H6'


# ===================================================================================================
def maak_kosten(ws):
    ws.title = 'CF - kosten'
    OO = "'CF - opbrengsten'!"
    zet(ws, 'B1', 'Kosten', FT)
    for k in range(1, NT + 1):
        b = kb(k)
        w = '$E$%d' % b
        zet(ws, 'B%d' % b, 'Budget aanneemsom (kosten, excl. KAO) en KAO/meerwerk · type %d' % k, FB)
        zet(ws, 'D%d' % b, None, FB, GEEL, NUM)
        zet(ws, 'E%d' % b, '=%sG%d' % (OO, won_rij(k)), FB, None, '0')
        zet(ws, 'F%d' % b, '=%s' % p_uit_kao(k), FB, None, NUM)
        kopregels(ws, b, b + 1)
        zet(ws, 'B%d' % (b + 1), "=%sB%d" % (OO, hb(k)), FB, KOP)
        for c, t in zip('CDEF', ['%', '€', '€ per wn', 'MMW € per wn']):
            zet(ws, '%s%d' % (c, b + 1), t, FB, KOP)
        zet(ws, '%s%d' % (CHK, b + 1), 'Check', FB, KOP)
        a = P_AANS1 + P_BLOK * (k - 1)
        for t in range(1, NTERM + 3):
            r = kterm(k, t)
            if t <= NTERM:
                zet(ws, 'B%d' % r, "=%sB%d" % (OO, plan(k, t)))
                zet(ws, 'C%d' % r, '=N(%s)' % pl('$%s$%d' % (P_TERM_PCT, a + 1 + t)), F, GEEL, '0.0%')
                zet(ws, 'D%d' % r, '=C%d*$D$%d' % (r, b), F, None, NUM)
                zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
                zet(ws, 'F%d' % r, '=IFERROR($F$%d*C%d/%s,0)' % (b, r, w), F, None, NUM2)
                for c in COLS:
                    zet(ws, '%s%d' % (c, r), "=%s%s%d" % (OO, c, plan(k, t)), F, GEEL, AANT)
            else:
                i = t - NTERM
                zet(ws, 'B%d' % r, 'Kopersmeerwerk %d%%' % (25 if i == 1 else 75))
                zet(ws, 'C%d' % r, 0.25 if i == 1 else 0.75, F, None, '0%')
                zet(ws, 'D%d' % r, "=%sD%d" % (OO, oz_rij(k, NTERM + i)), F, None, NUM)
                zet(ws, 'F%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
                t1, tn = kterm(k, 1), kterm(k, NTERM)
                for c in COLS:
                    f = '=%s%d' % (c, t1) if i == 1 else "=INDEX(%s$%d:%s$%d,%s$F$%d)" % (c, t1, c, tn, OO, oz_rij(k, NTERM + 2))
                    zet(ws, '%s%d' % (c, r), f, F, None, AANT)
            zet(ws, '%s%d' % (CHK, r), '=ROUND(SUM(%s%d:%s%d),2)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
        ra = k_aanneem(k)
        t1, tn, m1, m2 = kterm(k, 1), kterm(k, NTERM), kterm(k, NTERM + 1), kterm(k, NTERM + 2)
        zet(ws, 'B%d' % ra, 'Aanneemsom', FB)
        zet(ws, 'B%d' % (ra + 1), 'Meerwerk', FB)
        for c in COLS:
            zet(ws, '%s%d' % (c, ra), '=SUMPRODUCT($E$%d:$E$%d,%s%d:%s%d)' % (t1, tn, c, t1, c, tn), F, None, NUM)
            zet(ws, '%s%d' % (c, ra + 1), '=SUMPRODUCT($F$%d:$F$%d,%s%d:%s%d)+SUMPRODUCT($F$%d:$F$%d,%s%d:%s%d)'
                % (m1, m2, c, m1, c, m2, t1, tn, c, t1, c, tn), F, None, NUM)
    zet(ws, 'B%d' % K_PEW, 'Totaal kosten', FB)
    zet(ws, 'D%d' % K_PEW, "=N('2. Ontw PEW'!G20)*1000", FB, None, NUM)
    kopregels(ws, K_JAAR, K_OND, act='Actuals')
    zet(ws, '%s%d' % (COLS[0], K_JAAR), 'Actuals', FB)
    zet(ws, 'B%d' % K_OND, 'Onderdelen', FB, KOP)
    zet(ws, 'D%d' % K_OND, '€', FB, KOP)
    zet(ws, '%s%d' % (CHK, K_OND), 'Check', FB, KOP)
    zet(ws, '%s%d' % (TOEL, K_OND), 'Toelichting', FB, KOP)
    for i, p in enumerate(POSTEN):
        r = K_POST1 + i
        zet(ws, 'B%d' % r, "='2. Ontw PEW'!C%d" % (7 + i))
        zet(ws, 'D%d' % r, "=N('2. Ontw PEW'!G%d)*1000" % (7 + i), F, None, NUM)
        zet(ws, '%s%d' % (COLS[0], r), "=SUMIFS('FO - huidig'!$M:$M,'FO - huidig'!$Y:$Y,$B%d)" % r, F, GRIJS, NUM)
        for c in QCOLS:
            if p == 'Aanneemsom':
                f = '=' + '+'.join('%s%d' % (c, k_aanneem(k)) for k in range(1, NT + 1))
            elif p == 'Kopersmeerwerk':
                f = '=' + '+'.join('%s%d' % (c, k_aanneem(k) + 1) for k in range(1, NT + 1))
            else:
                f = None
            zet(ws, '%s%d' % (c, r), f, F, GEEL, NUM)
        zet(ws, '%s%d' % (CHK, r), '=ROUND(SUM(%s%d:%s%d),2)' % (COLS[0], r, COLS[-1], r), F, None, NUM)
        zet(ws, '%s%d' % (VER, r), '=IF(AND(%s%d-D%d<1,%s%d-D%d>-1),"OK","FOUT")' % (CHK, r, r, CHK, r, r), FB)
        zet(ws, '%s%d' % (TOEL, r), None, FN, GEEL)
    zet(ws, 'D%d' % K_SOM, '=SUM(D%d:D%d)' % (K_POST1, K_POSTN), FB, None, NUM)
    zet(ws, 'B%d' % K_CHECK, 'Check:', FB)
    zet(ws, 'D%d' % K_CHECK, '=D%d-D%d' % (K_SOM, K_PEW), F, None, NUM)
    zet(ws, 'E%d' % K_CHECK, '=IF(AND(D%d<1,D%d>-1),"OK","FOUT")' % (K_CHECK, K_CHECK), FB)
    zet(ws, 'B%d' % K_TOT, 'Totaal kosten', FB)
    for c in COLS:
        zet(ws, '%s%d' % (c, K_TOT), '=SUM(%s%d:%s%d)' % (c, K_POST1, c, K_POSTN), FB, None, NUM)
    zet(ws, '%s%d' % (CHK, K_TOT), '=ROUND(SUM(%s%d:%s%d),2)' % (COLS[0], K_TOT, COLS[-1], K_TOT), FB, None, NUM)
    zet(ws, 'B%d' % (K_TOT + 1), 'Kosten in %')
    for c in COLS:
        zet(ws, '%s%d' % (c, K_TOT + 1), '=IFERROR(%s%d/$D$%d,0)' % (c, K_TOT, K_PEW), F, None, '0%')
    zet(ws, 'B%d' % K_PER, 'Periode in cashflow SAP')
    for c in COLS:
        zet(ws, '%s%d' % (c, K_PER), "='CF - opbrengsten'!%s%d" % (c, PER), F, None, '0')
    markeringen(ws, 2, 2, K_PER + 40)
    for k in range(1, NT + 1):
        ws.row_dimensions.group(kterm(k, 1), k_aanneem(k) + 1, outline_level=1)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.column_dimensions['B'].width = 46
    for c in COLS:
        ws.column_dimensions[c].width = 9
    ws.freeze_panes = 'H3'


def p_uit_kao(k):
    return pl('$%s$%d' % (P_UIT['kao'], P_TYPE1 + k - 1))


# ===================================================================================================
def maak_prijslijst(ws):
    """per bouwnummer (btw eraf), per woningtype, aansluiting met SKO/PEW en het termijnschema"""
    ws.delete_rows(1, ws.max_row + 1)
    ws.title = 'Prijslijst'
    zet(ws, 'A1', 'Datum export:', FB)
    zet(ws, 'C1', None, FB, GEEL, 'dd-mm-yyyy')
    zet(ws, 'E1', 'Prijslijst en aansluiting · per woningtype uit Salesforce (of handmatig), aansluiting met SKO en PEW, '
                  'termijnschema en per bouwnummer', FN)
    kop = ['Nr', 'Woningtype (naam)', 'Type in prijslijst', 'Aantal', 'Grond (facturatie)', 'Aanneemsom', 'KAO-opties',
           'Kopersmeerwerk', 'Kadastrale kosten', 'Bouwrente', 'Aanvullende rente', 'Korting', 'VON excl. btw',
           'Handm. aantal', 'Handm. grond', 'Handm. aanneemsom + KAO', 'Handm. kopersmeerwerk', 'Handm. kad. + rente',
           'Aantal', 'Grond', 'Aanneemsom + KAO', 'Kopersmeerwerk', 'Kad. kosten + rente']
    zet(ws, 'D3', 'UIT DE PRIJSLIJST (excl. btw)', FW, DONKER)
    zet(ws, 'N3', 'HANDMATIG (overschrijft, leeg = prijslijst)', FW, DONKER)
    zet(ws, 'S3', 'GEBRUIKT IN DE CASHFLOW', FW, DONKER)
    for i, t in enumerate(kop):
        zet(ws, '%s4' % L(1 + i), t, FB, KOP)
        ws['%s4' % L(1 + i)].alignment = Alignment(wrap_text=True, vertical='center')
    rng = lambda c: '$%s$%d:$%s$%d' % (c, P_BN1, c, P_BNN)
    for k in range(1, NT + 1):
        r = P_TYPE1 + k - 1
        zet(ws, 'A%d' % r, k, FB)
        zet(ws, 'B%d' % r, None, FB, GEEL)
        zet(ws, 'C%d' % r, None, F, GEEL)
        zet(ws, 'D%d' % r, '=IF($C%d="",0,COUNTIF(%s,$C%d))' % (r, rng(P_BN['type']), r), F, GRIJS, '0')
        for x in ('grond', 'aanneemsom', 'kao', 'kmw', 'kad', 'bouwrente', 'rente', 'korting', 'von'):
            zet(ws, '%s%d' % (P_UIT[x], r), '=IF($C%d="",0,SUMIF(%s,$C%d,%s))' % (r, rng(P_BN['type']), r, rng(P_BN[x])),
                F, GRIJS, NUM)
        for x in P_HAND.values():
            zet(ws, '%s%d' % (x, r), None, F, GEEL, NUM)
        g = lambda h, u: '=IF(%s%d<>"",%s%d,%s)' % (h, r, h, r, u)
        zet(ws, '%s%d' % (P_GEBR['aantal'], r), g('N', 'D%d' % r), FB, None, '0')
        zet(ws, '%s%d' % (P_GEBR['grond'], r), g('O', 'E%d' % r), FB, None, NUM)
        zet(ws, '%s%d' % (P_GEBR['aanneemkao'], r), g('P', 'F%d+G%d' % (r, r)), FB, None, NUM)
        zet(ws, '%s%d' % (P_GEBR['kmw'], r), g('Q', 'H%d' % r), FB, None, NUM)
        zet(ws, '%s%d' % (P_GEBR['kadrente'], r), g('R', 'I%d+J%d+K%d+L%d' % (r, r, r, r)), FB, None, NUM)
    # aansluiting per type + termijnschema
    uit = {'Grond + Ontwikkelkosten': 'E', 'Aanneemsom': 'F', 'KAO opties': 'G', 'Korting': 'L', 'Kopersmeerwerk': 'H',
           'Kadastrale kosten': 'I', 'Bouwrente': 'J', 'Aanvullende rente': 'K'}
    for k in range(1, NT + 1):
        a = p_aans(k)
        t = P_TYPE1 + k - 1
        zet(ws, 'A%d' % a, 'Vergelijking excl. BTW - type %d' % k, FB)
        zet(ws, 'C%d' % a, '=B%d&""' % t, FB)
        for c, h in zip('BDFGHIJ', ['SKO', 'Prijslijst', 'Actueel budget', 'Corr', 'PEW', 'Verschil', 'Status']):
            zet(ws, '%s%d' % (c, a + 1), h, FB, KOP)
        zet(ws, '%s%d' % (P_TERM_NAAM, a + 1), 'Termijnschema · termijn', FB, KOP)
        zet(ws, '%s%d' % (P_TERM_PCT, a + 1), '% van aanneemsom + KAO', FB, KOP)
        for i, lab in enumerate(P_REGELS):
            r = a + 2 + i
            zet(ws, 'A%d' % r, lab)
            zet(ws, 'B%d' % r, None, F, GEEL, NUM)
            zet(ws, 'D%d' % r, '=%s%d' % (uit[lab], t) if lab in uit else None, F, GRIJS, NUM)
            zet(ws, 'F%d' % r, '=B%d' % r, F, None, NUM)
            zet(ws, 'G%d' % r, '=N(D%d)-N(F%d)' % (r, r), FB, None, NUM)
            zet(ws, 'H%d' % r, '=F%d+G%d' % (r, r), F, None, NUM)
            zet(ws, 'I%d' % r, '=N(D%d)-H%d' % (r, r), F, None, NUM)
            zet(ws, 'J%d' % r, '=IF(ABS(I%d)<1,"OK","FOUT")' % r, FB)
        for i in range(1, NTERM + 1):
            r = a + 1 + i
            zet(ws, '%s%d' % (P_TERM_NAAM, r), None, F, GEEL)
            zet(ws, '%s%d' % (P_TERM_PCT, r), None, F, GEEL, '0.0%')
        r = a + 2 + len(P_REGELS)
        zet(ws, '%s%d' % (P_TERM_NAAM, r), 'Totaal', FB)
        zet(ws, '%s%d' % (P_TERM_PCT, r), '=SUM(%s%d:%s%d)' % (P_TERM_PCT, a + 2, P_TERM_PCT, a + 1 + NTERM), FB, None, '0.0%')
    # per bouwnummer (zoals de oude Prijslijst)
    kb_ = ['Bouwnummer', 'Termijnen / Blok', 'Type', 'SF: Grondkosten start verkoop (incl. Kad)', 'SF: Aanneemsom start verkoop',
           'SF: VON start verkoop', 'SF: Totaal woonwensen in VON', 'SF: Totaal woonwensen in OB', 'SF: Status',
           'Kadastrale kosten', 'Grondkosten start verkoop (facturatie)', 'Aanneemsom start verkoop', 'VON start verkoop',
           'Totaal woonwensen in VON', 'Aanneemsom incl. woonwensen VON', 'Totaal woonwensen in OB', 'Bouwrente',
           'Aanvullende rente', 'Korting', 'Totaal opbrengsten']
    zet(ws, 'A%d' % (P_BN_KOP - 1), 'PER BOUWNUMMER · bouwnummer en type invullen, de rest komt uit de Salesforce-dump',
        FW, DONKER)
    for i, t in enumerate(kb_):
        zet(ws, '%s%d' % (L(1 + i), P_BN_KOP), t, FB, KOP)
    for r in range(P_BN1, P_BNN + 1):
        zet(ws, 'A%d' % r, None, F, GEEL)
        zet(ws, 'C%d' % r, None, F, GEEL)
        for j, c in enumerate('DEFGHI'):
            zet(ws, '%s%d' % (c, r), "=IF($A%d=\"\",\"\",_xlfn.IFNA(_xlfn.XLOOKUP($A%d,'Prijslijst dump'!$A:$A,'Prijslijst dump'!%s:%s),\"\"))"
                % (r, r, L(2 + j), L(2 + j)), F, GRIJS)
        zet(ws, 'K%d' % r, '=IFERROR((D%d-J%d)/1.21,"")' % (r, r), F, None, NUM)
        zet(ws, 'L%d' % r, '=IFERROR(E%d/1.21,"")' % r, F, None, NUM)
        zet(ws, 'M%d' % r, '=IFERROR(J%d+K%d+L%d,"")' % (r, r, r), F, None, NUM)
        zet(ws, 'N%d' % r, '=IF(G%d="","",IFERROR(G%d/1.21,-LEFT(G%d,SEARCH("-",G%d)-1)))' % (r, r, r, r), F, None, NUM)
        zet(ws, 'O%d' % r, '=IFERROR(L%d+N%d,"")' % (r, r), F, None, NUM)
        zet(ws, 'P%d' % r, '=IF(H%d="","",IFERROR(H%d/1.21,-LEFT(H%d,SEARCH("-",H%d)-1)/1.21))' % (r, r, r, r), F, None, NUM)
        for c in 'JQRS':
            zet(ws, '%s%d' % (c, r), None, F, GEEL, NUM)
        zet(ws, 'T%d' % r, '=IF($A%d="","",N(J%d)+N(K%d)+N(O%d)+N(P%d)+N(Q%d)+N(R%d))' % (r, r, r, r, r, r, r), F, None, NUM)
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 24
    ws.freeze_panes = 'A5'


# ===================================================================================================
def maak_vena(ws):
    for row in ws.iter_rows():
        for c in row:
            c.value = None
    zet(ws, 'B1', 'Mnd')
    zet(ws, 'C1', 'Check')
    for i in range(60):
        zet(ws, '%s1' % L(4 + i), i + 1)
    O = "'CF - opbrengsten'!"
    K = "'CF - kosten'!"
    po = '%s$%s$%d:$%s$%d' % (O, COLS[0], PER, COLS[-1], PER)
    to = '%s$%s$%d:$%s$%d' % (O, COLS[0], TOT, COLS[-1], TOT)
    pk = '%s$%s$%d:$%s$%d' % (K, COLS[0], K_PER, COLS[-1], K_PER)
    tk = '%s$%s$%d:$%s$%d' % (K, COLS[0], K_TOT, COLS[-1], K_TOT)
    zet(ws, 'B3', 'Opbrengsten')
    zet(ws, 'B4', 'Kosten')
    zet(ws, 'C3', '=SUM(D3:BK3)')
    zet(ws, 'C4', '=SUM(D4:BK4)')
    for i in range(60):
        c = L(4 + i)
        for r, p, t in ((3, po, to), (4, pk, tk)):
            zet(ws, '%s%d' % (c, r), '=IF(SUMIF(%s,%s1,%s)=0,"",(SUMIF(%s,%s1,%s)))' % (p, c, t, p, c, t))


def maak_cashflow_plek(ws, jaar=2026, rijen=None):
    """placeholder '1. Cashflow' (vervalt in het werkbestand): periodes voor Invoer, waarden voor de controle"""
    for row in ws.iter_rows():
        for c in row:
            c.value = None
    for c, t in zip('BCDEFGHIJ', ['Jaar', 'Q', 'Kosten', '% Kosten', 'Omzet', '% Omzet', 'CF', 'CF ( x 1000)', 'CF vorig kwartaal']):
        ws['%s3' % c] = t
    per = [(jaar - 5 + i, None) for i in range(5)] + [(jaar + i // 4, 'Q%d' % (i % 4 + 1)) for i in range(NQ)]
    rijen = rijen or {}
    for i, (j, q) in enumerate(per):
        r = 4 + i
        ws['B%d' % r] = j
        ws['C%d' % r] = q
        for c, v in (rijen.get((j, q)) or {}).items():
            ws['%s%d' % (c, r)] = v
    t = 4 + len(per) + 1
    ws['C%d' % t] = 'Totaal:'
    ws['D%d' % (t + 3)] = 'OK'
    ws['F%d' % (t + 3)] = 'OK'


def maak(uit, bron=DEMO, leeg=True, jaar=2026, cashflow=None, verborgen_types=None, verleden_tot=None):
    wb = openpyxl.load_workbook(bron)
    weg = [n for n in wb.sheetnames if n not in ('FO - vorig', 'FO - huidig', 'FO - actuals', 'Prijslijst dump', 'Prijslijst',
                                                 'Behind the scens', '2. Ontw PEW', 'CF - hulp VENA', '1. Cashflow',
                                                 'CF - opbrengsten', 'CF - kosten')]
    for n in weg:
        del wb[n]
    for n in ('CF - opbrengsten', 'CF - kosten'):
        del wb[n]
    maak_opbrengsten(wb.create_sheet('CF - opbrengsten'))
    maak_kosten(wb.create_sheet('CF - kosten'))
    maak_prijslijst(wb['Prijslijst'])
    for n in ('CF - hulp VENA', '1. Cashflow'):
        i = wb.sheetnames.index(n)
        wb.remove(wb[n])
        wb.create_sheet(n, i)
    maak_vena(wb['CF - hulp VENA'])
    maak_cashflow_plek(wb['1. Cashflow'], jaar, cashflow)
    pew = wb['2. Ontw PEW']
    if leeg:
        for r in range(7, 28):
            for c in ('E', 'K'):
                if not (isinstance(pew['%s%d' % (c, r)].value, str) and str(pew['%s%d' % (c, r)].value).startswith('=SUM')):
                    pew['%s%d' % (c, r)].value = None
        leegmaken(wb)
    wb['CF - opbrengsten']['D3'] = jaar
    # FO - actuals: hulpkolom S voor de top 3 kosten in het reviewkwartaal (per kostensoortomschrijving, kolom E)
    act = wb['FO - actuals']
    act['S1'] = 'Top 3 (hulp)'
    for r in range(2, 4251):
        act['S%d' % r] = ('=IF(AND($Q{r}="Kosten",$P{r}&""=Start!$C$12&"",$R{r}="Q"&Start!$D$12),'
                          'IF(COUNTIFS($E$2:$E{r},$E{r},$Q$2:$Q{r},"Kosten",$P$2:$P{r},$P{r},$R$2:$R{r},$R{r})=1,'
                          'SUMIFS($K:$K,$E:$E,$E{r},$Q:$Q,"Kosten",$P:$P,$P{r},$R:$R,$R{r}),""),"")').format(r=r)
    if verborgen_types or verleden_tot:
        inklappen(wb, verborgen_types or [], jaar, verleden_tot)
    wb.save(uit)
    return uit


def inklappen(wb, types, jaar, verleden_tot):
    """woningtypes zonder woningen ingeklapt; kwartalen t/m de laatste actuals (kwartaalindex) als ingeklapte groep"""
    o, k = wb['CF - opbrengsten'], wb['CF - kosten']
    for t in types:
        for r in range(plan(t, 1), tr(t) + 4):
            o.row_dimensions[r].hidden = True
        for r in range(kterm(t, 1), k_aanneem(t) + 2):
            k.row_dimensions[r].hidden = True
    if verleden_tot:
        n = verleden_tot - (jaar * 4 + 1) + 1
        if n > 0:
            for ws in (o, k):
                ws.column_dimensions.group(QCOLS[0], QCOLS[min(n, NQ) - 1], outline_level=1, hidden=True)


def leegmaken(wb):
    """exports leeg (kop en hulpformules blijven)"""
    for n, r1, kols in (('FO - vorig', 17, 21), ('FO - huidig', 17, 21), ('FO - actuals', 2, 16), ('Prijslijst dump', 2, 9)):
        ws = wb[n]
        for row in ws.iter_rows(min_row=r1, max_col=kols):
            for c in row:
                c.value = None
        if n.startswith('FO - h') or n.startswith('FO - v'):
            ws['E4'] = None


if __name__ == '__main__':
    print(maak(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'output', 'standaard_fo.xlsx')))
