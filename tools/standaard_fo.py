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
import re
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
BLOK = NTERM + 11                       # # won, titel, 10 planning, transport, verkocht %, getransp. %, grondprijs, 5x €


def won_rij(k):
    return 5 + BLOK * (k - 1)


def hb(k):                              # titelrij 'Termijnen k - naam'
    return won_rij(k) + 1


def plan(k, t):                         # t = 1..NTERM
    return hb(k) + t


def tr(k):
    return hb(k) + NTERM + 1


VK = won_rij(NT + 1) + 2                # koprij tabel verkoop en transport (per type: verkocht, getransporteerd)
def vk_rij(k):                          # verkocht (aantal woningen per kwartaal)
    return VK + 2 * k - 1


OMZ_N = NTERM + 4                       # grond, termijnen, kmw 25%, kmw 75%, kadastrale kosten + rente
OBLOK = OMZ_N + 3
def tt_rij(k):                          # getransporteerd (aantal woningen per kwartaal), direct onder verkocht
    return VK + 2 * k


VK_TOT = VK + 2 * NT + 1               # totaal verkocht (alle types)
TT_TOT = VK + 2 * NT + 2               # totaal getransporteerd
VT_PCT = TT_TOT + 2                # alle woningen: verkocht %/kw, cumulatief, getransporteerd %/kw, cumulatief
KANS = VT_PCT + 5                      # kans grondtermijn: kop, per type (netto effect), daarna de totaalregels:
KT = KANS + NT + 1                     # KT grond huidig, +1 grond met kans, +2 extra grond, +3 minder bouwtermijnen,
                                       # +4 netto per kwartaal, +5 netto cumulatief, +6/+7 CF cum. huidig/met kans (verborgen)
OZ1 = KANS + NT + 12                   # eerste omzet-koprij (# won op de rij erboven)


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
OVR1, OVR2 = OV_KOP + 1, OV_KOP + 2     # commercieel / parkeren in de cashflow (naar rato bouwtermijnen of handmatig)
OV1, OV2 = OV_KOP + 3, OV_KOP + 4       # handmatige bedragen (Oud FO inlezen zet hier de overige opbrengsten)
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
K_OVZ = K_PER + 47                      # CF-overzicht onderaan CF - kosten (onder het voorstelblok)
K_AH = K_PER + 23                       # hulprijen onder het voorstelblok: geplande aanneemsom per kwartaal
K_AF = K_PER + 24                       # en het eerste kwartaal vanaf het huidige met een geplande termijn
POSTEN = ['Grondkosten', 'Aanneemsom', 'Kopersmeerwerk', 'Overige bouwkosten / meerwerken', 'Honoraria',
          'Promotiekosten', 'Verhuur & verkoopkosten', 'Aansluitkosten', 'Legeskosten', 'Rente / Subsidies / bankgarantie',
          'Commerciele / Operationele reserve', 'Direct doorbelaste projectkosten']

# ---- Prijslijst --------------------------------------------------------------------------------------
P_TYPE1 = 13                             # rij type k = P_TYPE1 + k - 1 (rij 3-9: kaarten bovenaan)
P_REGELS = ['Grond + Ontwikkelkosten', 'Aanneemsom', 'VON verhogingen', 'KAO opties', 'Korting', 'Kopersmeerwerk',
            'Kadastrale kosten', 'Bouwrente', 'Aanvullende rente', 'Opbrengsten commercieel', 'Opbrengsten parkeren']
P_REGELS_HAND = ['VON verhogingen', 'Opbrengsten commercieel', 'Opbrengsten parkeren']   # prijslijst: invullen
P_BLOK = len(P_REGELS) + 4
P_AANS1 = P_TYPE1 + NT + 24              # na de totaalregel en de grafieken


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
P_BN = {'wt': 'U', 'type': 'C', 'status': 'I', 'kad': 'J', 'grond': 'K', 'aanneemsom': 'L', 'von': 'M', 'kao': 'N', 'kmw': 'P',
        'bouwrente': 'Q', 'rente': 'R', 'korting': 'S'}

GEEL = PatternFill('solid', fgColor='FFFFF2CC')
GRIJS = PatternFill('solid', fgColor='FFF3F4F6')
KOPLICHT = PatternFill('solid', fgColor='FFE9ECEF')   # kwartaalkoppen: rustig grijs
KOP = PatternFill('solid', fgColor='FF003366')       # kopregels: donkerblauw met witte tekst (contrast met geel = invoer)
TOTAAL = PatternFill('solid', fgColor='FFDCE8F5')    # totaalregels
DONKER = PatternFill('solid', fgColor='FF003366')
F = Font(name='Heijmans-Light', size=10)
FB = Font(name='Heijmans-Light', size=10, bold=True)
FT = Font(name='Heijmans-BoldCon', size=14, color='FF003366')
FW = Font(name='Heijmans-BoldCon', size=10, color='FFFFFFFF')
FN = Font(name='Heijmans-Light', size=9, italic=True, color='FF4B5563')
NUM = '#,##0;-#,##0;"-"'
AANT = 'General;-General;"-"'          # aantallen: geen punt of komma achter een heel getal
NUM2 = '#,##0.00;-#,##0.00;"-"'


def ci_(c):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(c)


def zet(ws, ref, v, font=F, fill=None, fmt=None):
    c = ws[ref]
    c.value = v
    if fill is KOP:
        font = Font(name=font.name, size=font.size, bold=True, italic=font.italic, color='FFFFFFFF')
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def navigatie(ws, link, legenda=None):
    """'← Start' (link naar tab Start) en een legenda met kleurvakjes: geel = invullen, wit = berekend,
    grijs = uit een export, blauw = totaal"""
    zet(ws, link, '=HYPERLINK("#Start!A1","← Start")', Font(name='Heijmans-BoldCon', size=10, color='FF003366', underline='single'))
    if legenda:
        from openpyxl.utils import column_index_from_string as cis
        c0, r0 = re.match(r'([A-Z]+)(\d+)', legenda).groups()
        rand = Border(left=Side(style='thin', color='FFD1D5DB'), right=Side(style='thin', color='FFD1D5DB'),
                      top=Side(style='thin', color='FFD1D5DB'), bottom=Side(style='thin', color='FFD1D5DB'))
        for i, (t, f) in enumerate((('invullen', GEEL), ('berekend', None), ('uit export', GRIJS), ('totaal', TOTAAL))):
            cel = zet(ws, '%s%s' % (L(cis(c0) + i), r0), t, Font(name='Heijmans-Light', size=8, color='FF1F2937'), f)
            cel.alignment = Alignment(horizontal='center')
            cel.border = rand


def statuskleuren(ws, rng):
    """FOUT rood, LET OP oranje, OK groen (statuskolommen)"""
    from openpyxl.formatting.rule import FormulaRule
    c = rng.split(':')[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s,4)="FOUT"' % c], font=Font(bold=True, color='FF991B1B'),
                                  fill=PatternFill('solid', fgColor='FFFECACA', bgColor='FFFECACA')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s,6)="LET OP"' % c], font=Font(bold=True, color='FF92400E'),
                                  fill=PatternFill('solid', fgColor='FFFEF3C7', bgColor='FFFEF3C7')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s="OK"' % c], font=Font(bold=True, color='FF15803D')))


def totaalrijen(ws, rijen):
    """totaalregels: lichtblauw vlak, vet en donkerblauw, met een lijn erboven"""
    lijn = Side(style='thin', color='FF003366')
    for r in rijen:
        for c in ['B', 'C', 'D', 'E', 'F'] + COLS:
            cel = ws['%s%d' % (c, r)]
            cel.fill = TOTAAL
            cel.font = Font(name='Heijmans-Light', size=10, bold=True, color='FF003366')
            cel.border = Border(top=lijn)


def cf_overzicht(ws, R):
    """CF-overzicht onderaan CF - opbrengsten en CF - kosten (beide hetzelfde): per kwartaal de definitieve cashflow zoals
    in Invoer en VENA (opbrengsten, kosten, saldo, cumulatief), met links totaal · actuals · prognose. Kolom G = alles
    vóór het eerste jaar. Status per kolom: actuals (afgesloten, SAP) / nu (huidig kwartaal: SAP + prognose) / prognose.
    De namen cf_idx, cf_omzet, cf_kosten, cf_omzet_tot en cf_kosten_tot zet de bouw (Invoer)."""
    grijs = PatternFill('solid', fgColor='FFF3F4F6')
    geel = PatternFill('solid', fgColor='FFFFF2CC')
    for c in ['B', 'C', 'D', 'E', 'F'] + COLS:
        zet(ws, '%s%d' % (c, R), None, FW, DONKER)
    zet(ws, 'B%d' % R, 'CASHFLOW · ZO WORDT DE CF  (definitief, zoals in de cashflow en VENA)', FW, DONKER)
    zet(ws, 'B%d' % (R + 1), 'grijs = afgesloten kwartalen (SAP-actuals) · geel = huidig kwartaal (geboekt in SAP + prognose '
                             'voor de rest van het kwartaal) · wit = prognose', FN)
    kopregels(ws, R + 2, R + 3, act='eerder')
    for c, t in (('B', '€'), ('D', 'Totaal'), ('E', 'Actuals'), ('F', 'Prognose')):
        zet(ws, '%s%d' % (c, R + 3), t, FB, KOPLICHT)
    rs, ro, rk, rsal, rc, rchk = R + 4, R + 5, R + 6, R + 7, R + 8, R + 10
    zet(ws, 'B%d' % rs, 'status', FN)
    zet(ws, '%s%d' % (COLS[0], rs), 'actuals', FN)
    for i, c in enumerate(QCOLS, 1):
        q = '%s*4+%d' % (JAAR, i)
        zet(ws, '%s%d' % (c, rs), '=IF(%s<p_huidig,"actuals",IF(%s=p_huidig,"nu","prognose"))' % (q, q), FN)
        zet(ws, '%s%d' % (c, ro), '=SUMIF(cf_idx,%s,cf_omzet)' % q, F, None, NUM)
        zet(ws, '%s%d' % (c, rk), '=SUMIF(cf_idx,%s,cf_kosten)' % q, F, None, NUM)
    eerste, laatste = QCOLS[0], QCOLS[-1]
    zet(ws, '%s%d' % (COLS[0], ro), '=N(cf_omzet_tot)-SUM(%s%d:%s%d)' % (eerste, ro, laatste, ro), F, None, NUM)
    zet(ws, '%s%d' % (COLS[0], rk), '=N(cf_kosten_tot)-SUM(%s%d:%s%d)' % (eerste, rk, laatste, rk), F, None, NUM)
    for r, lab, font in ((ro, 'Opbrengsten', FB), (rk, 'Kosten', FB), (rsal, 'Saldo (opbrengsten − kosten)', FB),
                         (rc, 'Cumulatief saldo', FB)):
        zet(ws, 'B%d' % r, lab, font)
    for c in COLS:
        zet(ws, '%s%d' % (c, rsal), '=%s%d-%s%d' % (c, ro, c, rk), FB, None, NUM)
        prev = COLS[COLS.index(c) - 1] if c != COLS[0] else None
        zet(ws, '%s%d' % (c, rc), ('=%s%d+%s%d' % (prev, rc, c, rsal)) if prev else '=%s%d' % (c, rsal), FB, None, NUM)
    vol = '$%s$%d:$%s$%d' % (COLS[0], rs, laatste, rs)
    for r in (ro, rk, rsal):
        rr = '$%s$%d:$%s$%d' % (COLS[0], r, laatste, r)
        zet(ws, 'D%d' % r, '=SUM(%s)' % rr, FB, TOTAAL, NUM)
        zet(ws, 'E%d' % r, '=SUMIF(%s,"actuals",%s)' % (vol, rr), FB, TOTAAL, NUM)
        zet(ws, 'F%d' % r, '=D%d-E%d' % (r, r), FB, TOTAAL, NUM)
    zet(ws, 'C%d' % rc, 'eindsaldo', FN)
    zet(ws, 'D%d' % rc, '=%s%d' % (laatste, rc), FB, TOTAAL, NUM)
    dun = Side(style='thin', color='FF003366')
    for c in ['B', 'C', 'D', 'E', 'F'] + COLS:
        ws['%s%d' % (c, rsal)].border = Border(top=dun)
    zet(ws, 'B%d' % rchk, 'Controle: totaal = PEW (opbrengsten · kosten)', FN)
    zet(ws, 'D%d' % rchk, "=IF(AND(ABS(D%d-N('2. Ontw PEW'!G27)*1000)<10,ABS(D%d-N('2. Ontw PEW'!G20)*1000)<10),\"OK\","
                          "\"LET OP: opbrengsten \"&TEXT(D%d-N('2. Ontw PEW'!G27)*1000,\"#,##0\")&\" · kosten \"&"
                          "TEXT(D%d-N('2. Ontw PEW'!G20)*1000,\"#,##0\")&\" t.o.v. PEW\")" % (ro, rk, ro, rk), FB)
    rng = '%s%d:%s%d' % (COLS[0], rs, laatste, rc)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s$%d="actuals"' % (COLS[0], rs)], fill=grijs))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s$%d="nu"' % (COLS[0], rs)], fill=geel))
    ws.conditional_formatting.add('%s%d:%s%d' % (COLS[0], rc, laatste, rc),
                                  FormulaRule(formula=['%s%d<-0.5' % (COLS[0], rc)], font=Font(color='FFB91C1C', bold=True)))
    ws.conditional_formatting.add('D%d' % rchk, FormulaRule(formula=['LEFT(D%d,2)="OK"' % rchk], font=Font(color='FF15803D', bold=True)))
    ws.conditional_formatting.add('D%d' % rchk, FormulaRule(formula=['LEFT(D%d,6)="LET OP"' % rchk], font=Font(color='FFB45309', bold=True)))
    zet(ws, 'I1', '=HYPERLINK("#B%d","↓ CF-overzicht (wat wordt de cashflow)")' % R, Font(name='Heijmans-BoldCon', size=10, color='FF003366', underline='single'))


def kopregels(ws, r_jaar, r_q, act='actuals', jaar=JAAR):
    """jaartallen (op Q1) en Q1..Q4 boven de kolommen"""
    zet(ws, '%s%d' % (COLS[0], r_q), act, FB, KOPLICHT)
    for i, c in enumerate(QCOLS):
        # jaartal boven elk kwartaal (Q1 vet), zodat het ook na het inklappen van eerdere kwartalen zichtbaar is
        zet(ws, '%s%d' % (c, r_jaar), '=%s+%d' % (jaar, i // 4), FB if i % 4 == 0 else Font(name='Heijmans-Light', size=8, color='FF6B7280'))
        zet(ws, '%s%d' % (c, r_q), 'Q%d' % (i % 4 + 1), FB, KOPLICHT)


def markeringen(ws, r_mark, r1, r2):
    """rij r_mark: 'Start bouw' / 'Oplevering' boven het kwartaal (Start, rij 7 en 9); in r1..r2 de kolom gekleurd en
    een rand aan het begin van elk jaar (vakken per jaar)"""
    key = lambda c: '(%s*4+1+COLUMN(%s%d)-COLUMN($%s$%d))' % (JAAR, c, r1, QCOLS[0], r1)
    # in de cel zelf COLUMN() zonder verwijzing (geen verwijzing naar de eigen cel)
    kc = '(%s*4+1+COLUMN()-COLUMN($%s$%d))' % (JAAR, QCOLS[0], r1)
    for c in QCOLS:
        zet(ws, '%s%d' % (c, r_mark), '=IF(%s=p_startbouw,"Start bouw",IF(%s=p_oplevering,"Oplevering",""))'
            % (kc, kc), Font(name='Heijmans-BoldCon', size=9, color='FF003366'))
    zet(ws, 'B%d' % r_mark, '="Start bouw "&Start!$E$7&" · oplevering "&Start!$E$9', FN)
    rng = '%s%d:%s%d' % (QCOLS[0], r1, QCOLS[-1], r2)
    k0 = key(QCOLS[0]).replace('COLUMN(%s%d)' % (QCOLS[0], r1), 'COLUMN(%s%d)' % (QCOLS[0], r1))
    # kleur alleen in de markeringsrij; in de tabel een gekleurde lijn links en rechts van het kwartaal, zodat koppen
    # (blauw) en invoer (geel) hun eigen kleur houden
    rm = '%s%d:%s%d' % (QCOLS[0], r_mark, QCOLS[-1], r_mark)
    ws.conditional_formatting.add(rm, FormulaRule(formula=['%s=p_startbouw' % k0.replace('%d)' % r1, '%d)' % r_mark)],
                                                  fill=PatternFill('solid', fgColor='FF15803D', bgColor='FF15803D'),
                                                  font=Font(bold=True, color='FFFFFFFF')))
    ws.conditional_formatting.add(rm, FormulaRule(formula=['%s=p_oplevering' % k0.replace('%d)' % r1, '%d)' % r_mark)],
                                                  fill=PatternFill('solid', fgColor='FF2563EB', bgColor='FF2563EB'),
                                                  font=Font(bold=True, color='FFFFFFFF')))
    lg, lb = Side(style='thin', color='FF15803D'), Side(style='thin', color='FF2563EB')
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s=p_startbouw' % k0], border=Border(left=lg, right=lg)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['%s=p_oplevering' % k0], border=Border(left=lb, right=lb)))
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
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 8
    for c_, w_ in (('C', 8), ('D', 13), ('E', 11), ('F', 8)):
        ws.column_dimensions[c_].width = w_
    zet(ws, 'B2', 'Bouwplanning, verkoop en transport', FT)
    zet(ws, 'B3', 'Eerste jaar van de kwartalen (Q1)', FB)
    zet(ws, 'D3', None, FB, GEEL, '0')
    for k in range(1, NT + 1):
        w, h = won_rij(k), hb(k)
        zet(ws, 'F%d' % w, '# won', FB)
        zet(ws, 'G%d' % w, '=%s' % p_gebr(k, 'aantal'), FB, None, '0')
        # koopsom per woning en het aandeel grond (van Prijslijst)
        zet(ws, 'B%d' % w, 'Koopsom per woning · grond %', FN)
        zet(ws, 'D%d' % w, '=IFERROR((%s+%s)/%s,0)' % (p_gebr(k, 'grond'), p_gebr(k, 'aanneemkao'), p_gebr(k, 'aantal')), FB, None, NUM)
        zet(ws, 'E%d' % w, '=%s' % pl('$X$%d' % (P_TYPE1 + k - 1)), FB, None, '0.0%')
        kopregels(ws, w, h)
        zet(ws, 'B%d' % h, '="Termijnen %d - "&%s' % (k, pl('$B$%d' % (P_TYPE1 + k - 1))), FB, KOP)
        for t in range(1, NTERM + 1):
            r = plan(k, t)
            zet(ws, 'B%d' % r, '=B%d' % oz_rij(k, t))
            # % van de termijn, bedrag (alle woningen) en per woning: van het omzetblok
            zet(ws, 'C%d' % r, '=IF(B%d="","",C%d)' % (r, oz_rij(k, t)), FN, None, '0.0%')
            zet(ws, 'D%d' % r, '=IF(B%d="","",D%d)' % (r, oz_rij(k, t)), FN, None, NUM)
            zet(ws, 'E%d' % r, '=IF(B%d="","",E%d)' % (r, oz_rij(k, t)), FN, None, NUM)
            for c in COLS:
                zet(ws, '%s%d' % (c, r), None, F, None, AANT)
            zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
            zet(ws, '%s%d' % (VER, r), '=IF(B%d="","",ROUND($G$%d-%s%d,2))' % (r, w, CHK, r), F, None, AANT)
            zet(ws, '%s%d' % (STAT, r), '=IF(%s%d="","",IF(%s%d=0,"OK","FOUT"))' % (VER, r, VER, r), FB)
        # geel (invoer) alleen op de termijnen die er zijn: lege termijnregels blijven wit
        ws.conditional_formatting.add('%s%d:%s%d' % (COLS[0], plan(k, 1), COLS[-1], plan(k, NTERM)),
                                      FormulaRule(formula=['$B%d<>""' % plan(k, 1)], fill=PatternFill('solid', fgColor='FFFFF2CC', bgColor='FFFFF2CC'), stopIfTrue=False))
        for c, t_ in zip('CDE', ['%', '€ totaal', '€ per won.']):
            zet(ws, '%s%d' % (c, h), t_, FB, KOP)
        zet(ws, '%s%d' % (CHK, h), 'Check', FB, KOPLICHT)
        zet(ws, '%s%d' % (VER, h), 'Verschil', FB, KOPLICHT)
        # opbrengsten van dit type in € per kwartaal (schuiven mee met het transport)
        e0 = tr(k) + 4
        for i, (lab, f) in enumerate((
                ('Grondtermijn · koopsom grond (€, bij transport) · % van de koopsom', lambda c: '=%s%d' % (c, eu_rij(k, 0))),
                ('Bouwtermijnen (€) · % van de koopsom', lambda c: '=SUM(%s%d:%s%d)' % (c, eu_rij(k, 1), c, eu_rij(k, NTERM))),
                ('Kopersmeerwerk 25% (€, bij start bouw)', lambda c: '=%s%d' % (c, eu_rij(k, NTERM + 1))),
                ('Kopersmeerwerk 75% (€, bij oplevering)', lambda c: '=%s%d' % (c, eu_rij(k, NTERM + 2))),
                ('Totaal opbrengsten type %d (€)' % k, lambda c: '=SUM(%s%d:%s%d)' % (c, eu_rij(k, 0), c, eu_rij(k, OMZ_N - 1))))):
            r = e0 + i
            tot_ = i == 4
            zet(ws, 'B%d' % r, lab, FB if tot_ else F)
            if i in (0, 1):         # aandeel van de koopsom (grond + bouwtermijnen) en € per woning
                zet(ws, 'C%d' % r, '=IFERROR(D%d/($D$%d+$D$%d),0)' % (r, e0, e0 + 1), FN, None, '0.0%')
                zet(ws, 'E%d' % r, '=IFERROR(D%d/$G$%d,0)' % (r, w), FN, None, NUM)
            if tot_:                # totaal per woning (koopsom + meerwerk)
                zet(ws, 'E%d' % r, '=IFERROR(D%d/$G$%d,0)' % (r, w), FB, None, NUM)
            if i in (2, 3):         # % van het kopersmeerwerk en € per woning, zoals bij de termijnen
                zet(ws, 'C%d' % r, '=C%d' % oz_rij(k, NTERM + i - 1), FN, None, '0%')
                zet(ws, 'E%d' % r, '=E%d' % oz_rij(k, NTERM + i - 1), FN, None, NUM)
            zet(ws, 'D%d' % r, '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), FB if tot_ else F, None, NUM)
            for c in COLS:
                zet(ws, '%s%d' % (c, r), f(c), FB if tot_ else F, None, NUM)
    # verkoop en transport: één tabel, per type een regel verkocht en een regel getransporteerd (aantallen woningen per
    # kwartaal). Kolom G (t/m nu) komt uit de Salesforce-status in de prijslijst; de kwartalen erna zijn invoer.
    from openpyxl.formatting.rule import DataBarRule
    kopregels(ws, VK - 1, VK, act='t/m nu')
    for c, t in zip('BCDEF', ['VERKOOP EN TRANSPORT · aantal woningen', 'Aantal', 't/m nu (SF)', 'Nog', 'Voortgang']):
        zet(ws, '%s%d' % (c, VK), t, FB, KOP)
    zet(ws, '%s%d' % (CHK, VK), 'Check', FB, KOP)
    s_bn = '$%s$%d:$%s$%d' % (P_BN['status'], P_BN1, P_BN['status'], P_BNN)
    w_bn = 'Prijslijst!$%s$%d:$%s$%d' % (P_BN['wt'], P_BN1, P_BN['wt'], P_BNN)
    tel = lambda k, st_: 'COUNTIFS(%s,%d,Prijslijst!%s,"%s")' % (w_bn, k, s_bn, st_)
    FI = Font(name='Heijmans-Light', size=10, color='FF4B5563')
    for k in range(1, NT + 1):
        for r, lab, sf in ((vk_rij(k), '=%s&" · verkocht"' % pl('$B$%d' % (P_TYPE1 + k - 1)),
                            '%s+%s' % (tel(k, 'Verkocht'), tel(k, 'Getransporteer*'))),
                           (tt_rij(k), '    getransporteerd', tel(k, 'Getransporteer*'))):
            zet(ws, 'B%d' % r, lab, FB if r == vk_rij(k) else FI)
            zet(ws, 'C%d' % r, '=G%d' % won_rij(k), F, None, '0')
            zet(ws, 'D%d' % r, '=' + sf, F, GRIJS, '0')
            zet(ws, 'E%d' % r, '=C%d-D%d' % (r, r), F, None, '0')
            zet(ws, 'F%d' % r, '=IF(N(C%d)=0,"",D%d/C%d)' % (r, r, r), F, None, '0%')
            zet(ws, 'G%d' % r, '=D%d' % r, F, GRIJS, AANT)
            for c in QCOLS:
                zet(ws, '%s%d' % (c, r), None, F, GEEL, AANT)
            zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, AANT)
            zet(ws, '%s%d' % (VER, r), '=ROUND(C%d-%s%d,2)' % (r, CHK, r), F, None, AANT)
        ws['B%d' % tt_rij(k)].border = Border(bottom=Side(style='thin', color='FFD1D5DB'))
    for r, lab, rijen in ((VK_TOT, 'Totaal · verkocht', [vk_rij(k) for k in range(1, NT + 1)]),
                          (TT_TOT, 'Totaal · getransporteerd', [tt_rij(k) for k in range(1, NT + 1)])):
        zet(ws, 'B%d' % r, lab, FB, TOTAAL)
        for c in ['C', 'D', 'E'] + COLS + [CHK]:
            zet(ws, '%s%d' % (c, r), '=' + '+'.join('%s%d' % (c, x) for x in rijen), FB, TOTAAL, '0' if c in 'CDE' else AANT)
        zet(ws, 'F%d' % r, '=IF(N(C%d)=0,"",D%d/C%d)' % (r, r, r), FB, TOTAAL, '0%')
    # bron zichtbaar: de statussen in de Salesforce-export (Prijslijst dump), per bouwnummer gekoppeld op Prijslijst
    st_ = lambda x: 'COUNTIF(Prijslijst!%s,"%s")' % (s_bn, x)
    zet(ws, 'B%d' % (TT_TOT + 1), '=HYPERLINK("#\'Prijslijst dump\'!A1","t/m nu uit Salesforce (Prijslijst dump): "&%s&" verkocht · "&'
                                  '%s&" getransporteerd · "&%s&" in optie · "&%s&" te koop  →")'
        % (st_('Verkocht'), st_('Getransporteer*'), st_('In optie'), st_('Te koop')),
        Font(name='Heijmans-Light', size=9, italic=True, color='FF003366', underline='single'))
    ws.conditional_formatting.add('F%d:F%d' % (VK + 1, TT_TOT), DataBarRule(start_type='num', start_value=0, end_type='num',
                                                                             end_value=1, color='FF003366', showValue=True))
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
    from openpyxl.worksheet.datavalidation import DataValidation
    zet(ws, 'C%d' % OV_KOP, 'verdeling', FB, KOP)
    zet(ws, 'D%d' % OV_KOP, 'PEW (aansl.)', FB, KOP)
    zet(ws, 'E%d' % OV_KOP, 'te verdelen', FB, KOP)
    zet(ws, '%s%d' % (CHK, OV_KOP), 'Verdeeld', FB, KOPLICHT)
    zet(ws, '%s%d' % (VER, OV_KOP), 'Verschil', FB, KOPLICHT)
    dv = DataValidation(type='list', formula1='"bouwtermijnen,handmatig"', allow_blank=False)
    ws.add_data_validation(dv)
    # verdeelsleutel: aandeel van de gefactureerde bouwtermijnen (alle woningtypes) in het kwartaal
    bt_tot = '+'.join('$D$%d' % (tr(k) + 5) for k in range(1, NT + 1))
    for r, h, lab in ((OVR1, OV1, 'Opbrengsten commercieel'), (OVR2, OV2, 'Opbrengsten parkeren')):
        std = '%s · handmatig (€)' % lab
        zet(ws, 'B%d' % r, '=IF($B%d="%s","%s",$B%d)' % (h, std, lab, h), FB)
        zet(ws, 'C%d' % r, 'bouwtermijnen', F, GEEL)
        dv.add('C%d' % r)
        # bedrag volgens de aansluiting op Prijslijst (kolom PEW, alle woningtypes)
        i_ = P_REGELS.index(lab)
        zet(ws, 'D%d' % r, '=IF($B%d<>"%s","",%s)' % (h, std, '+'.join(pl('$H$%d' % (p_aans(k) + 2 + i_)) for k in range(1, NT + 1))), FB, None, NUM)
        zet(ws, 'E%d' % r, '=IF(D%d="","",ROUND(D%d-%s%d,2))' % (r, r, CHK, r), F, None, NUM)
        rato = 'AND($C%d="bouwtermijnen",$B%d="%s")' % (r, h, std)
        for c in COLS:
            bt_c = '+'.join('%s%d' % (c, tr(k) + 5) for k in range(1, NT + 1))
            zet(ws, '%s%d' % (c, r), '=IF(%s,IFERROR(N($D%d)*(%s)/(%s),0),N(%s%d))' % (rato, r, bt_c, bt_tot, c, h), F, None, NUM)
        zet(ws, '%s%d' % (CHK, r), '=SUM(%s%d:%s%d)' % (COLS[0], r, COLS[-1], r), F, None, NUM)
        zet(ws, '%s%d' % (VER, r), '=E%d' % r, F, None, NUM)
        zet(ws, '%s%d' % (STAT, r), ('=IF(AND($C{r}="bouwtermijnen",$B{h}="{std}",SUM({c0}{h}:{cn}{h})<>0),'
                                     '"LET OP: handmatige bedragen (rij {h}) tellen niet mee bij bouwtermijnen",'
                                     'IF(D{r}="","",IF(AND(D{r}=0,{chk}{r}=0),"",IF(ABS(E{r})<1,"OK",'
                                     '"LET OP: "&TEXT(E{r},"#,##0")&" nog te verdelen"))))').format(
            r=r, h=h, std=std, c0=COLS[0], cn=COLS[-1], chk=CHK), FB)
        # handmatige bedragen (geel); bij 'bouwtermijnen' niet gebruikt
        zet(ws, 'B%d' % h, std, FN)
        for c in COLS:
            zet(ws, '%s%d' % (c, h), None, F, GEEL, NUM)
        zet(ws, '%s%d' % (CHK, h), '=SUM(%s%d:%s%d)' % (COLS[0], h, COLS[-1], h), F, None, NUM)
    zet(ws, 'B%d' % TOT, 'Totaal opbrengsten', FB)
    for c in COLS + [CHK]:
        delen = ['%s%d:%s%d' % (c, eu_rij(k, 0), c, eu_rij(k, OMZ_N - 1)) for k in range(1, NT + 1)]
        zet(ws, '%s%d' % (c, TOT), '=SUM(%s,%s%d:%s%d)+%s%d' % (','.join(delen), c, OVR1, c, OVR2, c, DAEB_TOT), FB, None, NUM)
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
            zet(ws, '%s%d' % (c, r), f, FB if i % 2 else F, None, '0.0%;-0.0%;"-"')
    # kans: grondtermijn naar een ander % van de koopsom (Prijslijst kolom Y); effect op de cashflow (alleen prognose)
    zet(ws, 'B%d' % KANS, 'KANS · grondtermijn naar het % op Prijslijst (kolom Y) · effect op de opbrengsten', FW, DONKER)
    for c in ['C', 'D', 'E', 'F'] + COLS:
        zet(ws, '%s%d' % (c, KANS), None, FW, DONKER)
    zet(ws, 'C%d' % KANS, 'al getransp.', FW, DONKER)
    zet(ws, 'D%d' % KANS, 'totaal prognose', FW, DONKER)
    prog = lambda i: '(%s*4+%d)>=p_huidig' % (JAAR, i)
    for k in range(1, NT + 1):
        r = KANS + k
        z_ = 'N(%s)' % pl('$Z$%d' % (P_TYPE1 + k - 1))
        zet(ws, 'B%d' % r, '="Type %d "&%s&" · "&IF(%s="","geen kans",TEXT(%s,"0%%")&" grond · netto effect")'
            % (k, pl('$B$%d' % (P_TYPE1 + k - 1)), pl('$Y$%d' % (P_TYPE1 + k - 1)), pl('$Y$%d' % (P_TYPE1 + k - 1))))
        # al getransporteerd (t/m de actuals): voor die woningen ligt de verdeling grond / bouwtermijnen vast
        t_ = tr(k)
        zet(ws, 'C%d' % r, '=N(%s%d)+SUMPRODUCT(%s%d:%s%d,--((%s*4+COLUMN(%s%d:%s%d)-COLUMN(%s%d)+1)<p_huidig))'
            % (COLS[0], t_, QCOLS[0], t_, QCOLS[-1], t_, JAAR, QCOLS[0], t_, QCOLS[-1], t_, QCOLS[0], t_), FN, None, '0')
        # per kwartaal: + extra grond bij transport (woningen die nog getransporteerd worden), − het deel van hun
        # bouwtermijnen dat met het grond-% verschuift (de eerste woningen, al getransporteerd, krijgen de termijnen eerst)
        for i, c in enumerate(COLS):
            if i == 0:
                f = 0
            else:
                p_ = COLS[i - 1]
                bt = '+'.join('$C$%d*(MAX(0,SUM($%s%d:%s%d)-$C%d)-MAX(0,SUM($%s%d:%s%d)-$C%d))'
                              % (oz_rij(k, t), COLS[0], oz_rij(k, t), c, oz_rij(k, t), r, COLS[0], oz_rij(k, t), p_, oz_rij(k, t), r)
                              for t in range(1, NTERM + 1))
                f = '=IF(%s=0,0,IF(%s,%s*(%s%d-(%s)),0))' % (z_, prog(i), z_, c, t_, bt)
            zet(ws, '%s%d' % (c, r), f, F, None, NUM)
        zet(ws, 'D%d' % r, '=SUM(%s%d:%s%d)' % (QCOLS[0], r, QCOLS[-1], r), F, None, NUM)
    rt = KT
    labels = ['Grondtermijn · huidig (€)', 'Grondtermijn · met kans (€)', 'Extra opbrengst grondtermijn (€, eerder ontvangen)',
              'Minder bouwtermijnen (€, later)', 'Netto effect kans per kwartaal (€)', 'Netto effect cumulatief (€)',
              'Cashflow cumulatief · huidig (€, hulp voor het Dashboard)', 'Cashflow cumulatief · met kans (€, hulp)']
    for j, lab in enumerate(labels):
        zet(ws, 'B%d' % (rt + j), lab, FB if j in (1, 2, 4) else F)
    for i, c in enumerate(COLS):
        if i == 0:
            for j in range(6):
                zet(ws, '%s%d' % (c, rt + j), 0, F, None, NUM)
            continue
        grond = '+'.join('%s%d' % (c, tr(k) + 4) for k in range(1, NT + 1))
        extra = '+'.join('N(%s)*%s%d' % (pl('$Z$%d' % (P_TYPE1 + k - 1)), c, tr(k)) for k in range(1, NT + 1))
        zet(ws, '%s%d' % (c, rt), '=IF(%s,%s,0)' % (prog(i), grond), F, None, NUM)
        zet(ws, '%s%d' % (c, rt + 2), '=IF(%s,%s,0)' % (prog(i), extra), FB, None, NUM)
        zet(ws, '%s%d' % (c, rt + 1), '=%s%d+%s%d' % (c, rt, c, rt + 2), FB, None, NUM)
        zet(ws, '%s%d' % (c, rt + 4), '=SUM(%s%d:%s%d)' % (c, KANS + 1, c, KANS + NT), FB, None, NUM)
        zet(ws, '%s%d' % (c, rt + 3), '=%s%d-%s%d' % (c, rt + 4, c, rt + 2), F, None, NUM)
        zet(ws, '%s%d' % (c, rt + 5), '=SUM($%s%d:%s%d)' % (QCOLS[0], rt + 4, c, rt + 4), F, None, NUM)
        zet(ws, '%s%d' % (c, rt + 6), '=IFERROR(INDEX(Invoer!$H$8:$H$67,MATCH(%s*4+%d,Invoer!$A$8:$A$67,0)),"")' % (JAAR, i), F, None, NUM)
        zet(ws, '%s%d' % (c, rt + 7), '=IF(%s%d="","",%s%d+%s%d)' % (c, rt + 6, c, rt + 6, c, rt + 5), F, None, NUM)
    for j in range(5):
        zet(ws, 'D%d' % (rt + j), '=SUM(%s%d:%s%d)' % (QCOLS[0], rt + j, QCOLS[-1], rt + j), FB if j in (1, 2, 4) else F, None, NUM)
    zet(ws, 'C%d' % (rt + 5), 'hoogste:', FN)
    zet(ws, 'D%d' % (rt + 5), '=MAX(0,MAX(%s%d:%s%d))' % (QCOLS[0], rt + 5, QCOLS[-1], rt + 5), F, None, NUM)
    for r in (rt + 6, rt + 7):
        zet(ws, 'C%d' % r, 'laagste:', FN)
        zet(ws, 'D%d' % r, '=MIN(%s%d:%s%d)' % (QCOLS[0], r, QCOLS[-1], r), F, None, NUM)
        ws.row_dimensions[r].hidden = True
    totaalrijen(ws, [rt + 2, rt + 4])
    markeringen(ws, 4, 4, PER + 40)
    # inklappen: invoerblok per woningtype, en de rekenblokken (aantallen per termijn en €)
    for k in range(1, NT + 1):
        ws.row_dimensions.group(plan(k, 1), hb(k) + BLOK - 2, outline_level=1)
    ws.row_dimensions.group(OZ1 - 1, OV_KOP - 2, outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.column_dimensions['B'].width = 52
    for c in COLS:
        ws.column_dimensions[c].width = 11
    totaalrijen(ws, [tr(k) + 8 for k in range(1, NT + 1)] + [TOT])
    cf_overzicht(ws, DAEB_VERKOOP + 4)
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12
    navigatie(ws, 'A1', 'D1')
    statuskleuren(ws, '%s1:%s%d' % (CHK, TOEL, DAEB_VERKOOP + 5))
    ws.freeze_panes = 'H6'


# ===================================================================================================
def maak_kosten(ws):
    ws.title = 'CF - kosten'
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 8
    OO = "'CF - opbrengsten'!"
    zet(ws, 'B1', None)
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
                zet(ws, 'C%d' % r, '=N(%s)' % pl('$%s$%d' % (P_TERM_PCT, a + 1 + t)), F, None, '0.0%')
                zet(ws, 'D%d' % r, '=C%d*$D$%d' % (r, b), F, None, NUM)
                zet(ws, 'E%d' % r, '=IFERROR(D%d/%s,0)' % (r, w), F, None, NUM2)
                zet(ws, 'F%d' % r, '=IFERROR($F$%d*C%d/%s,0)' % (b, r, w), F, None, NUM2)
                for c in COLS:
                    zet(ws, '%s%d' % (c, r), "=%s%s%d" % (OO, c, plan(k, t)), F, None, AANT)
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
        # geel (overschrijfbaar) alleen op de termijnen die er zijn
        geel_cf = PatternFill('solid', fgColor='FFFFF2CC', bgColor='FFFFF2CC')
        for rng_ in ('C%d:C%d' % (kterm(k, 1), kterm(k, NTERM)), '%s%d:%s%d' % (COLS[0], kterm(k, 1), COLS[-1], kterm(k, NTERM))):
            ws.conditional_formatting.add(rng_, FormulaRule(formula=['$B%d<>""' % kterm(k, 1)], fill=geel_cf))
        # lege termijnregels (geen termijn op Prijslijst) onzichtbaar: geen 0,0% en streepjes
        ws.conditional_formatting.add('C%d:%s%d' % (kterm(k, 1), CHK, kterm(k, NTERM)),
                                      FormulaRule(formula=['$B%d=""' % kterm(k, 1)], font=Font(color='FFFFFFFF')))
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
    zet(ws, 'B%d' % (K_PEW + 1), 'Methode per post (kolom C) = hoe het bedrag dat nog niet is uitgegeven (PEW − al geboekt in SAP) over '
                                 'de kwartalen gaat · HANDMATIG: je vult zelf per kwartaal de bedragen in · VOORSTEL: het restant gelijk '
                                 'verdeeld vanaf het huidige kwartaal t/m oplevering (vul je zelf een kwartaal in, dan wordt de rest '
                                 'verdeeld) · TRANSPORT: het restant naar rato van het aantal woningen dat per kwartaal wordt '
                                 'getransporteerd. De uitkomst staat in het voorstelblok onderaan; de totaalregel is de definitieve cashflow.', FN)
    zet(ws, 'D%d' % K_OND, '€', FB, KOP)
    zet(ws, '%s%d' % (CHK, K_OND), 'Check', FB, KOP)
    zet(ws, '%s%d' % (TOEL, K_OND), 'Toelichting', FB, KOP)
    for i, p in enumerate(POSTEN):
        r = K_POST1 + i
        zet(ws, 'B%d' % r, "='2. Ontw PEW'!C%d" % (7 + i))
        zet(ws, 'D%d' % r, "=N('2. Ontw PEW'!G%d)*1000" % (7 + i), F, None, NUM)
        zet(ws, '%s%d' % (COLS[0], r), "=SUMIFS('FO - huidig'!$M:$M,'FO - huidig'!$Y:$Y,$B%d)" % r, F, None, NUM)
        for i_, c in enumerate(QCOLS, 1):
            if p == 'Aanneemsom':
                # wat al betaald is telt niet nog eens mee met de geplande termijnen: het verschil tussen de planning t/m nu
                # (kolom G) en het werkelijk betaalde gaat af in het eerste kwartaal vanaf het huidige met een geplande
                # termijn (hulprijen K_AH / K_AF); zonder zo'n kwartaal in het huidige kwartaal
                vlag = 'OR(%s%d=1,AND(SUM($%s$%d:$%s$%d)=0,%s*4+%d=p_huidig))' % (c, K_AF, QCOLS[0], K_AF, QCOLS[-1], K_AF, JAAR, i_)
                f = '=' + '+'.join('%s%d' % (c, k_aanneem(k)) for k in range(1, NT + 1)) + \
                    '+IF(%s,%s-$%s%d,0)' % (vlag, '+'.join('$%s%d' % (COLS[0], k_aanneem(k)) for k in range(1, NT + 1)), COLS[0], r)
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
    # hulprijen: geplande aanneemsom per kwartaal en het eerste kwartaal vanaf het huidige met een termijn
    zet(ws, 'B%d' % K_AH, 'hulp · aanneemsom volgens planning (alle types)', FN)
    zet(ws, 'B%d' % K_AF, 'hulp · eerste kwartaal vanaf het huidige met een geplande termijn (1)', FN)
    for i_, c in enumerate(QCOLS, 1):
        zet(ws, '%s%d' % (c, K_AH), '=' + '+'.join('%s%d' % (c, k_aanneem(k)) for k in range(1, NT + 1)), FN, None, NUM)
        eerder = 'SUM($%s%d:%s%d)=0' % (QCOLS[0], K_AF, QCOLS[i_ - 2], K_AF) if i_ > 1 else 'TRUE'
        zet(ws, '%s%d' % (c, K_AF), '=IF(AND(%s*4+%d>=p_huidig,%s%d<>0,%s),1,0)' % (JAAR, i_, c, K_AH, eerder), FN, None, '0')
    ws.row_dimensions.group(K_AH, K_AF, outline_level=1, hidden=True)
    zet(ws, 'B%d' % K_PER, 'Periode in cashflow SAP')
    for c in COLS:
        zet(ws, '%s%d' % (c, K_PER), "='CF - opbrengsten'!%s%d" % (c, PER), F, None, '0')
    markeringen(ws, 2, 2, K_PER + 40)
    for k in range(1, NT + 1):
        ws.row_dimensions.group(kterm(k, 1), k_aanneem(k) + 1, outline_level=1)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.column_dimensions['B'].width = 52
    for c in COLS:
        ws.column_dimensions[c].width = 11
    for c_, w_ in (('C', 12), ('D', 13), ('E', 12), ('F', 12)):
        ws.column_dimensions[c_].width = w_
    totaalrijen(ws, [K_PEW, K_TOT] + [k_aanneem(k) for k in range(1, NT + 1)])
    cf_overzicht(ws, K_OVZ)
    navigatie(ws, 'A1', 'D1')
    statuskleuren(ws, '%s1:%s%d' % (CHK, TOEL, K_PER + 80))
    ws.freeze_panes = 'H3'


def p_uit_kao(k):
    return pl('$%s$%d' % (P_UIT['kao'], P_TYPE1 + k - 1))


# ===================================================================================================
def maak_prijslijst(ws):
    """dashboard: kaarten, woningtypes (prijslijst / handmatig / gebruikt / kans), grafieken, aansluiting per type met het
    termijnschema; daaronder (door de bouw) 'te boeken in SAP' en de prijslijst per bouwnummer"""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import CellIsRule, DataBarRule
    ws.delete_rows(1, ws.max_row + 1)
    ws.title = 'Prijslijst'
    ws.column_dimensions.clear()             # oude (overlappende) kolomdefinities en groepen van het sjabloon weg
    ws.sheet_view.showGridLines = False
    WIT = PatternFill('solid', fgColor='FFFFFFFF')
    KAART = PatternFill('solid', fgColor='FFF5F7FA')
    LICHT = PatternFill('solid', fgColor='FFE9ECEF')
    BAND = {'uit': PatternFill('solid', fgColor='FF6B7280'), 'hand': PatternFill('solid', fgColor='FFD97706'),
            'gebr': PatternFill('solid', fgColor='FF003366'), 'kans': PatternFill('solid', fgColor='FF15803D')}
    FK = Font(name='Heijmans-Light', size=9, color='FF4B5563')
    FV = Font(name='Heijmans-BoldCon', size=18, color='FF003366')
    dun = Side(style='thin', color='FFD1D5DB')
    zet(ws, 'A1', 'Datum export:', FK)
    zet(ws, 'C1', None, FB, GEEL, 'dd-mm-yyyy')
    zet(ws, 'B2', 'Prijslijst en aansluiting', FT)
    zet(ws, 'F2', 'per woningtype uit Salesforce (of handmatig) · aansluiting met SKO en PEW · termijnschema · kans '
                  'grondtermijn', FN)
    t1, tn = P_TYPE1, P_TYPE1 + NT - 1
    rij_tot = tn + 1
    regels = lambda col: '+'.join('SUM(%s%d:%s%d)' % (col, p_aans(k) + 2, col, p_aans(k) + 1 + len(P_REGELS))
                                  for k in range(1, NT + 1))
    rk = KT
    OO = "'CF - opbrengsten'!"
    kaarten = [
        ('B', 'WONINGEN', '=SUM(S%d:S%d)' % (t1, tn), '0', '=COUNTIF(B%d:B%d,"?*")&" woningtypes"' % (t1, tn), None),
        ('F', 'OPBRENGSTEN PRIJSLIJST (excl. btw)', '=SUM(T%d:W%d)' % (t1, tn), NUM, 'grond + aanneemsom + KAO + meerwerk + kad.', None),
        ('J', 'SKO (actueel budget)', '=' + regels('B'), NUM, 'verschil met prijslijst:', '=SUM(T%d:W%d)-(%s)' % (t1, tn, regels('B'))),
        ('N', 'TE BOEKEN IN SAP (Corr)', '=' + regels('G'), NUM, 'per SAP-nummer: zie blok 3', None),
        ('R', 'AANSLUITING', '=IF((%s)=0,"OK",(%s)&" type(s) met FOUT")' % (('+'.join('N(J%d="FOUT")' % p_aans(k) for k in range(1, NT + 1)),) * 2), None, 'SKO ↔ prijslijst ↔ PEW, per regel', None),
        ('V', 'KANS GRONDTERMIJN · extra eerder ontvangen', '=%sD%d' % (OO, rk + 5), NUM,
         '="max. financiering huidig "&TEXT(%sD%d,"#,##0")&" · met kans:"' % (OO, rk + 6), '=%sD%d' % (OO, rk + 7)),
    ]
    for c, kop_, waarde, fmt, sub, subw in kaarten:
        c2 = L(ci_(c) + 2)
        for r in range(4, 9):
            for cc in range(ci_(c), ci_(c) + 3):
                cel = ws.cell(r, cc)
                cel.fill = KAART
                cel.border = Border(top=Side(style='medium', color='FFFEB900') if r == 4 else None,
                                    left=dun if cc == ci_(c) else None, right=dun if cc == ci_(c) + 2 else None,
                                    bottom=dun if r == 8 else None)
        zet(ws, '%s4' % c, kop_, FK, KAART)
        zet(ws, '%s5' % c, waarde, FV, KAART, fmt)
        ws.merge_cells('%s5:%s6' % (c, c2))
        ws['%s5' % c].alignment = Alignment(horizontal='left', vertical='center')
        zet(ws, '%s7' % c, sub, FK, KAART)
        if subw:
            zet(ws, '%s8' % c, subw, FB, KAART, NUM)
    ws.conditional_formatting.add('R5', CellIsRule(operator='equal', formula=['"OK"'], font=Font(color='FF15803D', bold=True, size=18)))
    ws.conditional_formatting.add('R5', CellIsRule(operator='notEqual', formula=['"OK"'], font=Font(color='FFB91C1C', bold=True, size=18)))
    # 1 woningtypes
    zet(ws, 'A%d' % (t1 - 3), '1 · WONINGTYPES', FW, DONKER)
    for cc in range(2, 27):
        ws.cell(t1 - 3, cc).fill = DONKER
    for (c1_, c2_), tekst, band in ((('D', 'M'), 'UIT DE PRIJSLIJST (excl. btw)', 'uit'),
                                    (('N', 'R'), 'HANDMATIG · overschrijft, leeg = prijslijst', 'hand'),
                                    (('S', 'X'), 'GEBRUIKT IN DE CASHFLOW', 'gebr'), (('Y', 'Z'), 'KANS', 'kans')):
        for cc in range(ci_(c1_), ci_(c2_) + 1):
            ws.cell(t1 - 2, cc).fill = BAND[band]
        zet(ws, '%s%d' % (c1_, t1 - 2), tekst, FW, BAND[band])
    kop = ['Nr', 'Woningtype (naam)', 'Type of blok in prijslijst (bv. A, of 4A;4B, * = alle)', 'Aantal', 'Grond (facturatie)', 'Aanneemsom', 'KAO-opties',
           'Kopersmeerwerk', 'Kadastrale kosten', 'Bouwrente', 'Aanvullende rente', 'Korting', 'VON excl. btw',
           'Aantal', 'Grond', 'Aanneemsom + KAO', 'Kopersmeerwerk', 'Kad. kosten + rente',
           'Aantal', 'Grond', 'Aanneemsom + KAO', 'Kopersmeerwerk', 'Kad. kosten + rente', 'Grond % · grond / (grond + aanneemsom), excl. KAO',
           'Kans: grond % (leeg = geen kans)', 'Kans: grond +/- per woning']
    for i, t in enumerate(kop):
        zet(ws, '%s%d' % (L(1 + i), t1 - 1), t, FB, LICHT)
        ws['%s%d' % (L(1 + i), t1 - 1)].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[t1 - 1].height = 32
    rng = lambda c: '$%s$%d:$%s$%d' % (c, P_BN1, c, P_BNN)
    for k in range(1, NT + 1):
        r = P_TYPE1 + k - 1
        zet(ws, 'A%d' % r, k, FB)
        zet(ws, 'B%d' % r, None, FB, GEEL)
        zet(ws, 'C%d' % r, None, F, GEEL)
        zet(ws, 'D%d' % r, '=COUNTIF(%s,$A%d)' % (rng(P_BN['wt']), r), F, GRIJS, '0')
        for x in ('grond', 'aanneemsom', 'kao', 'kmw', 'kad', 'bouwrente', 'rente', 'korting', 'von'):
            zet(ws, '%s%d' % (P_UIT[x], r), '=SUMIF(%s,$A%d,%s)' % (rng(P_BN['wt']), r, rng(P_BN[x])),
                F, GRIJS, NUM)
        for x in P_HAND.values():
            zet(ws, '%s%d' % (x, r), None, F, GEEL, NUM)
        # bron van de bedragen in de cashflow: 1 handmatig, 2 de prijslijst, 3 zonder prijslijst voor dit type (aantal
        # uit de prijslijst = 0, bijvoorbeeld vroeg in het project) de SKO uit het aansluitblok van het type
        a = p_aans(k)
        sko = lambda *namen: '+'.join('N($B$%d)' % (a + 2 + P_REGELS.index(n_)) for n_ in namen)
        g = lambda h, u, s_=None: ('=IF(%s%d<>"",%s%d,IF(N($D%d)=0,%s,%s))' % (h, r, h, r, r, s_, u) if s_
                                   else '=IF(%s%d<>"",%s%d,%s)' % (h, r, h, r, u))
        kao = 'IF(N($D%d)=0,%s,G%d)' % (r, sko('KAO opties'), r)
        zet(ws, '%s%d' % (P_GEBR['aantal'], r), g('N', 'D%d' % r), FB, WIT, '0')
        zet(ws, '%s%d' % (P_GEBR['grond'], r), g('O', 'E%d' % r, sko('Grond + Ontwikkelkosten')), FB, WIT, NUM)
        zet(ws, '%s%d' % (P_GEBR['aanneemkao'], r), g('P', 'F%d+G%d' % (r, r), sko('Aanneemsom', 'KAO opties')), FB, WIT, NUM)
        zet(ws, '%s%d' % (P_GEBR['kmw'], r), g('Q', 'H%d' % r, sko('Kopersmeerwerk')), FB, WIT, NUM)
        zet(ws, '%s%d' % (P_GEBR['kadrente'], r), g('R', 'I%d+J%d+K%d+L%d' % (r, r, r, r),
                                                    sko('Kadastrale kosten', 'Bouwrente', 'Aanvullende rente', 'Korting')), FB, WIT, NUM)
        zet(ws, 'X%d' % r, '=IFERROR(T%d/(T%d+U%d-%s),0)' % (r, r, r, kao), FB, WIT, '0.0%')     # excl. KAO
        zet(ws, 'Y%d' % r, None, FB, GEEL, '0%')
        zet(ws, 'Z%d' % r, '=IF(Y%d="",0,IFERROR((Y%d-X%d)*(T%d+U%d-%s)/S%d,0))' % (r, r, r, r, r, kao, r), F, WIT, NUM)
        # hulp voor de grafiek: SKO en prijslijst per type (uit het aansluitblok)
        zet(ws, 'AB%d' % r, '=SUM(B%d:B%d)' % (a + 2, a + 1 + len(P_REGELS)), FK, None, NUM)
        zet(ws, 'AC%d' % r, '=SUM(D%d:D%d)' % (a + 2, a + 1 + len(P_REGELS)), FK, None, NUM)
        for cc in range(1, 27):
            ws.cell(r, cc).border = Border(bottom=dun)
    zet(ws, 'AB%d' % (t1 - 1), 'SKO', FK)
    zet(ws, 'AC%d' % (t1 - 1), 'Prijslijst', FK)
    zet(ws, 'AB%d' % (t1 - 2), 'hulp voor de grafiek', FK)
    zet(ws, 'B%d' % rij_tot, 'Totaal', FB, LICHT)
    for c in [L(x) for x in range(4, 24)] + ['Z']:
        zet(ws, '%s%d' % (c, rij_tot), '=SUM(%s%d:%s%d)' % (c, t1, c, tn) if c not in ('X', 'Z') else None, FB, LICHT,
            '0' if c in ('D', 'N', 'S') else NUM)
    zet(ws, 'X%d' % rij_tot, '=IFERROR(T%d/(T%d+U%d-G%d),0)' % (rij_tot, rij_tot, rij_tot, rij_tot), FB, LICHT, '0.0%')
    for c in ('A', 'C', 'Y', 'Z'):
        ws['%s%d' % (c, rij_tot)].fill = LICHT
    ws.conditional_formatting.add('X%d:X%d' % (t1, tn), DataBarRule(start_type='num', start_value=0, end_type='num',
                                                                     end_value=1, color='FFFEB900', showValue=True))
    ws.conditional_formatting.add('Y%d:Y%d' % (t1, tn), CellIsRule(operator='greaterThan', formula=['0'],
                                                                   fill=PatternFill('solid', fgColor='FFD1FAE5', bgColor='FFD1FAE5')))
    # grafieken
    g1 = BarChart()
    g1.type, g1.grouping, g1.gapWidth = 'col', 'clustered', 60
    g1.title = 'Prijslijst en SKO per type'
    g1.add_data(Reference(ws, min_col=ci_('AC'), min_row=t1 - 1, max_row=tn), titles_from_data=True)
    g1.add_data(Reference(ws, min_col=ci_('AB'), min_row=t1 - 1, max_row=tn), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=2, min_row=t1, max_row=tn))
    for s_, kl in zip(g1.series, ('003366', 'FEB900')):
        s_.graphicalProperties.solidFill = kl
        s_.graphicalProperties.line.noFill = True
    g1.y_axis.number_format = '#,##0'
    g1.y_axis.majorGridlines = None
    g1.x_axis.delete = g1.y_axis.delete = False
    g1.legend.position = 'b'
    g1.height, g1.width = 7, 16
    ws.add_chart(g1, 'B%d' % (rij_tot + 2))
    g2 = BarChart()
    g2.type, g2.grouping, g2.gapWidth = 'col', 'clustered', 60
    g2.title = 'Grond % · huidig en kans'
    g2.add_data(Reference(ws, min_col=ci_('X'), min_row=t1 - 1, max_row=tn), titles_from_data=True)
    g2.add_data(Reference(ws, min_col=ci_('Y'), min_row=t1 - 1, max_row=tn), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=2, min_row=t1, max_row=tn))
    for s_, kl in zip(g2.series, ('9CA3AF', '15803D')):
        s_.graphicalProperties.solidFill = kl
        s_.graphicalProperties.line.noFill = True
    g2.y_axis.number_format = '0%'
    g2.y_axis.scaling.min = 0
    g2.y_axis.majorGridlines = None
    g2.x_axis.delete = g2.y_axis.delete = False
    g2.legend.position = 'b'
    g2.height, g2.width = 7, 13
    ws.add_chart(g2, 'H%d' % (rij_tot + 2))
    # 2 aansluiting per type + termijnschema
    zet(ws, 'A%d' % (P_AANS1 - 2), '2 · AANSLUITING PER WONINGTYPE · SKO ↔ prijslijst ↔ PEW, met het termijnschema', FW, DONKER)
    for cc in range(2, 27):
        ws.cell(P_AANS1 - 2, cc).fill = DONKER
    uit = {'Grond + Ontwikkelkosten': 'E', 'Aanneemsom': 'F', 'KAO opties': 'G', 'Korting': 'L', 'Kopersmeerwerk': 'H',
           'Kadastrale kosten': 'I', 'Bouwrente': 'J', 'Aanvullende rente': 'K'}
    for k in range(1, NT + 1):
        a = p_aans(k)
        t = P_TYPE1 + k - 1
        ra, rz = a + 2, a + 1 + len(P_REGELS)
        for cc in range(1, 15):
            ws.cell(a, cc).fill = BAND['gebr']
        zet(ws, 'A%d' % a, 'Vergelijking excl. BTW - type %d' % k, FW, BAND['gebr'])
        zet(ws, 'C%d' % a, '=B%d&""' % t, Font(name='Heijmans-BoldCon', size=11, color='FFFFFFFF'), BAND['gebr'])
        zet(ws, 'J%d' % a, '=IF(COUNTIF(J%d:J%d,"FOUT")=0,"OK","FOUT")' % (ra, rz), FW, BAND['gebr'])
        zet(ws, 'E%d' % a, '=IF(AND(N($D%d)=0,SUM(B%d:B%d)<>0),IF(N(S%d)=0,"LET OP: vul het aantal woningen in (kolom N, rij %d)",'
                           '"geen prijslijst: de cashflow rekent met de SKO"),"")' % (t, ra, rz, t, t),
            Font(name='Heijmans-Light', size=9, italic=True, color='FFFFFFFF'), BAND['gebr'])
        zet(ws, 'N%d' % a, '=IF(ABS(%s%d-1)<0.0001,"OK",IF(%s%d=0,"","LET OP: niet 100%%"))'
            % (P_TERM_PCT, rz + 1, P_TERM_PCT, rz + 1), FW, BAND['gebr'])
        for c, h in zip('BDFGHIJ', ['SKO', 'Prijslijst', 'Actueel budget', 'Corr', 'PEW', 'Verschil', 'Status']):
            zet(ws, '%s%d' % (c, a + 1), h, FB, LICHT)
        ws['A%d' % (a + 1)].fill = LICHT
        zet(ws, '%s%d' % (P_TERM_NAAM, a + 1), 'Termijnschema · termijn', FB, LICHT)
        zet(ws, '%s%d' % (P_TERM_PCT, a + 1), '% (aanneemsom + KAO)', FB, LICHT)
        for i, lab in enumerate(P_REGELS):
            r = a + 2 + i
            zet(ws, 'A%d' % r, lab)
            zet(ws, 'B%d' % r, None, F, GEEL, NUM)
            # type zonder prijslijst (aantal uit de prijslijst = 0: vroeg in het project, of parkeren / commercieel met de
            # hand): de kolom Prijslijst toont dan het bedrag waarmee de cashflow rekent (handmatig, anders de SKO), zodat
            # er geen tegenboeking ontstaat en de PEW gelijk is aan de SKO
            if lab == 'Grond + Ontwikkelkosten':
                fd = '=IF(N($D%d)=0,T%d,E%d)' % (t, t, t)
            elif lab == 'Aanneemsom':
                fd = '=IF(N($D%d)=0,U%d-N(B%d),F%d)' % (t, t, a + 2 + P_REGELS.index('KAO opties'), t)
            elif lab == 'Kopersmeerwerk':
                fd = '=IF(N($D%d)=0,V%d,H%d)' % (t, t, t)
            elif lab in uit:
                fd = '=IF(N($D%d)=0,N(B%d),%s%d)' % (t, r, uit[lab], t)
            else:
                fd = None
            zet(ws, 'D%d' % r, fd, F, GRIJS if lab in uit else GEEL, NUM)
            zet(ws, 'F%d' % r, '=B%d' % r, F, None, NUM)
            zet(ws, 'G%d' % r, '=N(D%d)-N(F%d)' % (r, r), FB, None, NUM)
            zet(ws, 'H%d' % r, '=F%d+G%d' % (r, r), F, None, NUM)
            zet(ws, 'I%d' % r, '=N(D%d)-H%d' % (r, r), F, None, NUM)
            zet(ws, 'J%d' % r, '=IF(ABS(I%d)<1,"OK","FOUT")' % r, FB)
            for cc in range(1, 11):
                ws.cell(r, cc).border = Border(bottom=dun)
        zet(ws, 'C%d' % (rz + 1), 'Totaal', FB, LICHT)
        ws['A%d' % (rz + 1)].fill = LICHT
        for c in 'BDFGHI':
            zet(ws, '%s%d' % (c, rz + 1), '=SUM(%s%d:%s%d)' % (c, ra, c, rz), FB, LICHT, NUM)
        for c in 'EJ':
            ws['%s%d' % (c, rz + 1)].fill = LICHT
        ws.conditional_formatting.add('G%d:G%d' % (ra, rz), CellIsRule(operator='notBetween', formula=['-0.5', '0.5'],
                                      font=Font(color='FF92400E', bold=True), fill=PatternFill('solid', fgColor='FFFEF3C7', bgColor='FFFEF3C7')))
        for i in range(1, NTERM + 1):
            r = a + 1 + i
            zet(ws, '%s%d' % (P_TERM_NAAM, r), None, F, GEEL)
            zet(ws, '%s%d' % (P_TERM_PCT, r), None, F, GEEL, '0.0%')
        zet(ws, '%s%d' % (P_TERM_NAAM, rz + 1), 'Totaal', FB, LICHT)
        zet(ws, '%s%d' % (P_TERM_PCT, rz + 1), '=SUM(%s%d:%s%d)' % (P_TERM_PCT, a + 2, P_TERM_PCT, a + 1 + NTERM), FB, LICHT, '0.0%')
    rood = Font(color='FFB91C1C', bold=True)
    groen = Font(color='FF15803D', bold=True)
    for k in range(1, NT + 1):
        a = p_aans(k)
        ws.conditional_formatting.add('J%d:J%d' % (a + 2, a + 1 + len(P_REGELS)), CellIsRule(operator='equal', formula=['"FOUT"'], font=rood))
        ws.conditional_formatting.add('J%d:J%d' % (a + 2, a + 1 + len(P_REGELS)), CellIsRule(operator='equal', formula=['"OK"'], font=groen))
    for c, w in {'A': 26, 'B': 24, 'C': 16}.items():
        ws.column_dimensions[c].width = w
    for c in range(4, 27):
        ws.column_dimensions[L(c)].width = 13
    ws.column_dimensions[P_TERM_NAAM].width = 44
    for c in range(4, 27):
        if L(c) != P_TERM_NAAM:
            ws.column_dimensions[L(c)].width = 14
    navigatie(ws, 'E1', 'H1')
    # per bouwnummer (zoals de oude Prijslijst)
    kb_ = ['Bouwnummer', 'Termijnen / Blok', 'Type', 'SF: Grondkosten start verkoop (incl. Kad)', 'SF: Aanneemsom start verkoop',
           'SF: VON start verkoop', 'SF: Totaal woonwensen in VON', 'SF: Totaal woonwensen in OB', 'SF: Status',
           'Kadastrale kosten', 'Grondkosten start verkoop (facturatie)', 'Aanneemsom start verkoop', 'VON start verkoop',
           'Totaal woonwensen in VON', 'Aanneemsom incl. woonwensen VON', 'Totaal woonwensen in OB', 'Bouwrente',
           'Aanvullende rente', 'Korting', 'Totaal opbrengsten', 'Woningtype (nr) · automatisch', 'Rij in de dump (hulp)']
    zet(ws, 'A%d' % (P_BN_KOP - 1), 'PER BOUWNUMMER · bouwnummer en type invullen, de rest komt uit de Salesforce-dump',
        FW, DONKER)
    for i, t in enumerate(kb_):
        zet(ws, '%s%d' % (L(1 + i), P_BN_KOP), t, FB, KOP)
    for r in range(P_BN1, P_BNN + 1):
        zet(ws, 'A%d' % r, None, F, GEEL)
        zet(ws, 'C%d' % r, None, F, GEEL)
        # één zoekactie per bouwnummer (kolom V: rij in de dump), daarna INDEX per kolom
        zet(ws, 'V%d' % r, "=IF($A%d=\"\",\"\",IFERROR(MATCH($A%d,'Prijslijst dump'!$A:$A,0),\"\"))" % (r, r), FN, None, '0')
        for j, c in enumerate('DEFGHI'):
            zet(ws, '%s%d' % (c, r), "=IF($V%d=\"\",\"\",INDEX('Prijslijst dump'!%s:%s,$V%d))" % (r, L(2 + j), L(2 + j), r),
                F, GRIJS)
        zet(ws, 'K%d' % r, '=IFERROR((D%d-J%d)/1.21,"")' % (r, r), F, None, NUM)
        zet(ws, 'L%d' % r, '=IFERROR(E%d/1.21,"")' % r, F, None, NUM)
        zet(ws, 'M%d' % r, '=IFERROR(J%d+K%d+L%d,"")' % (r, r, r), F, None, NUM)
        zet(ws, 'N%d' % r, '=IF(G%d="","",IFERROR(G%d/1.21,-LEFT(G%d,SEARCH("-",G%d)-1)))' % (r, r, r, r), F, None, NUM)
        zet(ws, 'O%d' % r, '=IFERROR(L%d+N%d,"")' % (r, r), F, None, NUM)
        zet(ws, 'P%d' % r, '=IF(H%d="","",IFERROR(H%d/1.21,-LEFT(H%d,SEARCH("-",H%d)-1)/1.21))' % (r, r, r, r), F, None, NUM)
        for c in 'JQRS':
            zet(ws, '%s%d' % (c, r), None, F, GEEL, NUM)
        zet(ws, 'T%d' % r, '=IF($A%d="","",N(J%d)+N(K%d)+N(O%d)+N(P%d)+N(Q%d)+N(R%d))' % (r, r, r, r, r, r, r), F, None, NUM)
        # bij welk woningtype dit bouwnummer hoort (precies één): eerst op de kolom Type, dan op Termijnen / blok, dan '*';
        # bij meerdere passende types het eerste
        sl_ = 'SUBSTITUTE($C$%d:$C$%d,"; ",";")' % (P_TYPE1, P_TYPE1 + NT - 1)
        m_ = lambda c: 'MATCH(1,INDEX(($%s%d<>"")*ISNUMBER(SEARCH(";"&$%s%d&";",";"&%s&";")),0),0)' % (c, r, c, r, sl_)
        ex_ = lambda c: 'IF($%s%d="",NA(),MATCH($%s%d&"",INDEX($C$%d:$C$%d&"",0),0))' % (c, r, c, r, P_TYPE1, P_TYPE1 + NT - 1)
        zet(ws, 'U%d' % r, '=IF($A%d="","",IFERROR(%s,IFERROR(%s,IFERROR(%s,IFERROR(%s,IFERROR(MATCH("~*",$C$%d:$C$%d,0),""))))))'
            % (r, ex_('C'), m_('C'), ex_('B'), m_('B'), P_TYPE1, P_TYPE1 + NT - 1), FB, None, '0')
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 24
    ws.freeze_panes = 'A3'


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
    # per maand: het kwartaal (bij de laatste maand van het kwartaal) en de status t.o.v. het huidige kwartaal (tab Start)
    from openpyxl.formatting.rule import FormulaRule
    st = "%s$E$%d" % (O, PER)
    zet(ws, 'A2', 'Kwartaal', FB)
    zet(ws, 'A5', 'Status', FB)
    zet(ws, 'B2', '="Startmaand: "&IF(N(%s)=0,"(vul in op CF - opbrengsten rij %d)",TEXT(%s,"mm-yyyy"))' % (st, PER, st), FN)
    for i in range(60):
        c = L(4 + i)
        d = 'EDATE(%s,%s$1-1)' % (st, c)
        idx = '(YEAR(%s)*4+ROUNDUP(MONTH(%s)/3,0))' % (d, d)
        zet(ws, '%s2' % c, '=IF(N(%s)=0,"",IF(MOD(MONTH(%s),3)=0,YEAR(%s)&" Q"&ROUNDUP(MONTH(%s)/3,0),""))' % (st, d, d, d), FB)
        zet(ws, '%s5' % c, '=IF(N(%s)=0,"",IF(%s<p_huidig,"actuals",IF(%s=p_huidig,"nu","prognose")))' % (st, idx, idx), FN)
    zet(ws, 'B6', '=IF(OR(N(%s)=0,N(p_huidig)=0),"","Huidig kwartaal: Q"&(p_huidig-4*INT((p_huidig-1)/4))&" "&INT((p_huidig-1)/4)'
                  '&" = maand "&((INT((p_huidig-1)/4)-YEAR(%s))*12+(p_huidig-4*INT((p_huidig-1)/4))*3-MONTH(%s)+1)&'
                  '" (kolom met het groene vak)")' % (st, st, st), FB)
    groen = PatternFill('solid', fgColor='FFD1FAE5', bgColor='FFD1FAE5')
    grijs = Font(color='FF9CA3AF')
    ws.conditional_formatting.add('D1:BK5', FormulaRule(formula=['D$5="nu"'], fill=groen))
    ws.conditional_formatting.add('D1:BK5', FormulaRule(formula=['D$5="actuals"'], font=grijs))
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 14
    for i in range(60):
        ws.column_dimensions[L(4 + i)].width = 11
    ws.freeze_panes = 'D1'
    navigatie(ws, 'A1')


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
    # FO - vorig/huidig kolom Y (groepering): zoekt de CODA-post als tekst én als getal op in Behind the scens
    # (na kopiëren in Excel wordt '5100' vaak het getal 5100)
    for n in ('FO - vorig', 'FO - huidig'):
        ws_ = wb[n]
        for r in range(17, ws_.max_row + 1):
            v = ws_['Y%d' % r].value
            if isinstance(v, str) and 'VLOOKUP' in v.upper() and 'Behind the scens' in v:
                ws_['Y%d' % r] = ('=IF(X{r}="","",IFERROR(VLOOKUP(X{r}&"",\'Behind the scens\'!$A:$C,3,0),'
                                  'IFERROR(VLOOKUP(--X{r},\'Behind the scens\'!$A:$C,3,0),"")))').format(r=r)
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
    """woningtypes zonder woningen ingeklapt; kwartalen vóór de laatste actuals (kwartaalindex) als ingeklapte groep"""
    o, k = wb['CF - opbrengsten'], wb['CF - kosten']
    for t in types:
        for r in range(plan(t, 1), hb(t) + BLOK - 1):
            o.row_dimensions[r].hidden = True
        p = wb['Prijslijst']
        for r in [P_TYPE1 + t - 1, KANS + t] + list(range(p_aans(t), p_aans(t) + P_BLOK - 1)):
            (o if r == KANS + t else p).row_dimensions[r].hidden = True
        for r in (vk_rij(t), tt_rij(t)):
            o.row_dimensions[r].hidden = True
        for r in range(kterm(t, 1), k_aanneem(t) + 2):
            k.row_dimensions[r].hidden = True
    if verleden_tot:
        n = verleden_tot - (jaar * 4 + 1)          # het kwartaal van de laatste actuals blijft zichtbaar
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
