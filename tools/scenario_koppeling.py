"""Scenariotabbladen (Invoer, Woningtypes, Scenario kosten, Dashboard) met formules koppelen aan de FO-tabbladen.

Wat de macro 'Ophalen uit FO' eerst als waarden schreef, staat nu als formules in het bestand en rekent dus altijd mee:
  Invoer           periodes, kosten, omzet, cashflow, vorige prognose (rij 8-25) en verkocht/transport per woningtype
  Woningtypes      per woningtype: naam, aantal, koopsom, start bouw, grondtermijn, bouwtermijnen (% en bouwkwartaal),
                   extra opbrengsten per woning en bouwkosten per termijn (blok 7)
  Scenario kosten  per kostenpost: naam, 'schuift mee met' (standaard op de naam, overschrijven mag), gerealiseerd en
                   de definitieve prognose per periode (volgens de methode op CF - kosten)
  Dashboard        actuals t/m, FO-totalen
Het verborgen blad _koppeling zet de termijnen om naar losse regels (zelfde regels als de macro FOTypen): een termijn die
over meer kwartalen valt, wordt per kwartaal een regel '(1/2)', '(2/2)'; passen ze niet in het typeblok, dan een regel
per termijn op het gewogen gemiddelde kwartaal.
"""
from __future__ import annotations

import re

from xlsxpkg import Sheet, num2col

OK = "'CF - opbrengsten'!"
KK = "'CF - kosten'!"
KEYO = OK + "$G$199:$Z$199"
KEYK = KK + "$H$96:$AA$96"
MATO = OK + "$G$1:$Z$300"
MATK = KK + "$H$1:$AA$300"
ACT = 'p_actuals'
INV_R1, N_KW = 8, 16
INV_RN = INV_R1 + 1 + N_KW                  # 2 jaarrijen + 16 kwartalen: rij 8..25
SLOTS = {'T': 10, 'X': 6, 'B': 10}
KOOP_X0 = {'T': 16, 'X': 33, 'B': 82}       # eerste rij op Woningtypes
HULP = '_koppeling!'


def structuur(o: Sheet, k: Sheet):
    """woningtypes uit de opzet van CF - opbrengsten en CF - kosten"""
    get = o.get
    kop = [r for r in range(2, 48) if (get('F%d' % (r - 1)) or '').strip() == '# won']
    omz = [r for r in range(48, 190) if (get('F%d' % (r - 1)) or '').strip() == '# won']
    assert len(kop) == len(omz) == 4, (kop, omz)
    vk = [r for r in range(2, 120) if (get('B%d' % r) or '').strip().lower() == 'verkooptempo'][0]
    types = []
    for i, (hb, oh) in enumerate(zip(kop, omz)):
        plan, tr, d = [], None, None
        r = hb + 1
        while get('B%d' % r):
            b = get('B%d' % r)
            if b.startswith('Transport'):
                tr = r
            else:
                m = re.match(r'=B(\d+)$', b)
                assert m, (r, b)
                if d is None:
                    d = int(m.group(1)) - r
                assert int(m.group(1)) - r == d, 'planningsrij %d past niet bij omzetrij %s' % (r, b)
                plan.append(r)
            r += 1
        orows = []
        r = oh + 1
        while get('B%d' % r):
            orows.append(r)
            r += 1
        koop = orows[0]
        assert get('G%d' % koop) == '=G%d' % tr, (koop, tr)
        # termijnblok op CF - kosten met dezelfde titel
        kr = [r for r in range(2, 60) if k.get('B%d' % r) == "='CF - opbrengsten'!B%d" % hb]
        krows = []
        if kr:
            r = kr[0] + 1
            while k.get('B%d' % r) and not k.get('B%d' % r).lower().startswith(('aanneemsom', 'meerwerk')):
                krows.append(r)
                r += 1
        types.append({'k': i + 1, 'hb': hb, 'plan': plan, 'd': d, 'tr': tr, 'koop': koop, 'omzet': orows,
                      'verkoop': vk + 1 + i, 'kosten': krows})
    return types


def rij(mat, r):
    return 'INDEX(%s,%s,0)' % (mat, r)


def vul_hulp(h: Sheet, types, log):
    """blad _koppeling: per type een blok van 100 rijen; geeft per type de slotrijen terug"""
    slots = {}
    for t in types:
        k, R0 = t['k'], 1 + (t['k'] - 1) * 100
        aant = "N(%sG%d)" % (OK, t['hb'] - 1)
        h.set('A%d' % R0, 'Type %d' % k)
        h.set('B%d' % R0, 'start (kwartaalindex)')
        plan2d = "%s$G$%d:$Z$%d" % (OK, t['plan'][0], t['plan'][-1])
        start = 'C%d' % R0
        h.set(start, '=IFERROR(_xlfn.AGGREGATE(15,6,%s/((%s<>0)*(%s>0)),1),0)' % (KEYO, plan2d, KEYO))
        h.set('D%d' % R0, 'aantal')
        h.set('E%d' % R0, '=%s' % aant)
        st = '$C$%d' % R0
        slots[k] = {}
        base = R0 + 2
        for lijst, items in (('T', t['plan']), ('X', [o for o in t['omzet'] if o != t['koop']]), ('B', t['kosten'])):
            n_sl = SLOTS[lijst]
            h.set('A%d' % base, {'T': 'bouwtermijnen', 'X': "extra's per woning", 'B': 'bouwkosten per termijn'}[lijst])
            for c, kop in zip('ABCDEFGHIJKLM', ['bronrij', 'naam', 'waarde', 'telt', 'transport', 'n gesplitst', 'n enkel',
                                                 'n', 'begin', 'som', 'gewogen kw', 'gerealiseerd', 'n later']):
                h.set('%s%d' % (c, base + 1), kop)
            i1 = base + 2
            iN = i1 + len(items) - 1
            for j, src in enumerate(items):
                r = i1 + j
                if lijst == 'T':
                    o = src + t['d']
                    h.set('A%d' % r, src)
                    h.set('B%d' % r, '=%sB%d&""' % (OK, src))
                    h.set('C%d' % r, '=N(%sC%d)' % (OK, o))
                    h.set('D%d' % r, '=IF(C%d<>0,1,0)' % r)
                    h.set('E%d' % r, 0)
                elif lijst == 'X':
                    p = src - t['d']
                    if p in t['plan']:
                        h.set('A%d' % r, '=IF(SUMIFS(%s$G$%d:$Z$%d,%s,">0")<>0,%d,%d)' % (OK, p, p, KEYO, p, src))
                    else:
                        h.set('A%d' % r, src)
                    h.set('B%d' % r, '=%sB%d&""' % (OK, src))
                    h.set('C%d' % r, '=N(%sE%d)' % (OK, src))
                    h.set('D%d' % r, '=IF(AND(N(%sC%d)=0,ABS(C%d)>=0.005),1,0)' % (OK, src, r))
                    h.set('E%d' % r, '=IF(OR(N(%sJ%d)<>0,J%d=0),1,0)' % (OK, src, r))
                else:
                    h.set('A%d' % r, src)
                    h.set('B%d' % r, '=%sB%d&""' % (KK, src))
                    h.set('C%d' % r, '=N(%sE%d)' % (KK, src))
                    h.set('D%d' % r, '=IF(C%d<>0,1,0)' % r)
                    h.set('E%d' % r, 0)
                mat, key = (MATK, KEYK) if lijst == 'B' else (MATO, KEYO)
                R = rij(mat, 'A%d' % r)
                if lijst == 'B':
                    # gerealiseerd: kolom actuals (K) + kwartalen t/m actuals; later: kwartalen erna
                    h.set('L%d' % r, '=N(%sK%d)+SUMIFS(%s,%s,"<="&%s,%s,">0")' % (KK, src, R, key, ACT, key))
                    h.set('M%d' % r, '=SUMPRODUCT((%s<>0)*(%s>%s))' % (R, key, ACT))
                    h.set('J%d' % r, '=L%d+SUMIFS(%s,%s,">"&%s)' % (r, R, key, ACT))
                    h.set('K%d' % r, '=IF(J%d=0,"",(L%d*%s+SUMPRODUCT(%s*(%s>%s)*%s))/J%d)' % (r, r, ACT, R, key, ACT, key, r))
                    h.set('F%d' % r, '=IF(D%d=0,0,IF(AND(M%d=0,L%d=0),1,(L%d<>0)+M%d))' % (r, r, r, r, r))
                else:
                    h.set('J%d' % r, '=SUMIFS(%s,%s,">0")' % (R, key))
                    h.set('K%d' % r, '=IF(J%d=0,"",SUMPRODUCT(%s*(%s>0)*%s)/J%d)' % (r, R, key, key, r))
                    h.set('M%d' % r, '=SUMPRODUCT((%s<>0)*(%s>0))' % (R, key))
                    h.set('F%d' % r, '=IF(D%d=0,0,IF(OR(E%d=1,J%d=0),1,M%d))' % (r, r, r, r))
                h.set('G%d' % r, '=D%d' % r)
            mode = 'C%d' % base
            h.set('B%d' % base, 'enkel (te veel regels)?')
            h.set(mode, '=IF(SUM(F%d:F%d)>%d,1,0)' % (i1, iN, n_sl))
            for j in range(len(items)):
                r = i1 + j
                h.set('H%d' % r, '=IF($%s=1,G%d,F%d)' % (mode.replace('C', 'C$'), r, r))
                h.set('I%d' % r, '=1' if j == 0 else '=I%d+H%d' % (r - 1, r - 1))
            tot = 'E%d' % base
            h.set('D%d' % base, 'regels')
            h.set(tot, '=SUM(H%d:H%d)' % (i1, iN))
            # slots
            s1 = iN + 2
            for c, kop in zip('ABCDEFG', ['slot', 'item', 'deel', 'naam', 'waarde', 'kw', 'kolom']):
                h.set('%s%d' % (c, s1), kop)
            rng = lambda c: '$%s$%d:$%s$%d' % (c, i1, c, iN)
            mat, key = (MATK, KEYK) if lijst == 'B' else (MATO, KEYO)
            keyfirst = 'H' if lijst == 'B' else 'G'
            m = '$' + mode.replace('C', 'C$')
            for j in range(1, n_sl + 1):
                r = s1 + j
                it = 'INDEX(%s,$B%d)' % ('%s', r)
                I = lambda c: it % rng(c)
                h.set('A%d' % r, j)
                h.set('B%d' % r, '=IF(A%d>%s,"",MATCH(A%d,%s,1))' % (r, '$' + tot.replace('E', 'E$'), r, rng('I')))
                h.set('C%d' % r, '=IF(B%d="","",A%d-%s+1)' % (r, r, I('I')))
                R = rij(mat, I('A'))
                if lijst == 'B':
                    lat = 'C%d-(%s<>0)' % (r, I('L'))      # deelnummer onder de latere kwartalen
                    h.set('G%d' % r, '=IF(B%d="","",IF(OR(%s=1,%s<1),"",IFERROR(_xlfn.AGGREGATE(15,6,(COLUMN(%s)-COLUMN(%s$%s$1)+1)/((%s<>0)*(%s>%s)),%s),"")))'
                          % (r, m, lat, R, KK, keyfirst, R, key, ACT, lat))
                    h.set('D%d' % r, '=IF(B%d="","",%s&IF(%s=1,"",IF(G%d="",IF(AND(%s<>0,%s>0)," (gerealiseerd)",""),'
                                     'IF(%s>1," ("&(%s)&"/"&%s&")",""))))'
                          % (r, I('B'), m, r, I('L'), I('M'), I('M'), lat, I('M')))
                    h.set('E%d' % r, '=IF(B%d="","",ROUND(%s*IF(%s=1,IF(%s=0,%s,%s),IF(G%d="",IF(%s=0,%s,%s),INDEX(%s,G%d))),2))'
                          % (r, I('C'), m, I('J'), '$E$%d' % R0, I('J'), r, I('L'), '$E$%d' % R0, I('L'), R, r))
                    h.set('F%d' % r, '=IF(OR(B%d="",%s=0),"",IF(%s=1,IF(%s="","",ROUND(%s,0)-%s+1),IF(G%d="",IF(%s=0,"",%s-%s+1),INDEX(%s,G%d)-%s+1)))'
                          % (r, st, m, I('K'), I('K'), st, r, I('L'), ACT, st, key, r, st))
                else:
                    split = 'AND(%s=0,%s=0,%s>0)' % (m, I('E'), I('J'))
                    h.set('G%d' % r, '=IF(B%d="","",IF(%s,IFERROR(_xlfn.AGGREGATE(15,6,(COLUMN(%s)-COLUMN(%s$%s$1)+1)/((%s<>0)*(%s>0)),C%d),""),""))'
                          % (r, split, R, OK, keyfirst, R, key, r))
                    h.set('D%d' % r, '=IF(B%d="","",%s&IF(AND(G%d<>"",%s>1)," ("&C%d&"/"&%s&")",""))'
                          % (r, I('B'), r, I('M'), r, I('M')))
                    dec = 6 if lijst == 'T' else 2
                    h.set('E%d' % r, '=IF(B%d="","",IF(G%d="",%s,ROUND(%s*INDEX(%s,G%d)/%s,%d)))'
                          % (r, r, I('C'), I('C'), R, r, I('J'), dec))
                    h.set('F%d' % r, '=IF(OR(B%d="",%s=0),"",IF(G%d<>"",INDEX(%s,G%d)-%s+1,IF(OR(%s=1,%s=""),"",ROUND(%s,0)-%s+1)))'
                          % (r, st, r, key, r, st, I('E'), I('K'), I('K'), st))
            slots[k][lijst] = s1 + 1
            base = s1 + n_sl + 2
        log('  _koppeling type %d: %d planningsrijen, %d omzetrijen, %d kostenrijen'
            % (k, len(t['plan']), len(t['omzet']), len(t['kosten'])))
    return slots


def woningtypes(w: Sheet, types, slots, st):
    for t in types:
        k = t['k']
        c0 = 3 + 3 * (k - 1)
        cn, cp, ck = num2col(c0), num2col(c0 + 1), num2col(c0 + 2)
        R0 = 1 + (k - 1) * 100
        lab = '%sB%d' % (OK, t['hb'])
        aant = '%sG%d' % (OK, t['hb'] - 1)
        w.set('%s7' % cn, '=IF(N(%s)=0,"",IF(ISNUMBER(FIND(" - ",%s)),TRIM(MID(%s,FIND(" - ",%s)+3,200)),"Type %d"))'
              % (aant, lab, lab, lab, k))
        w.set('%s8' % cn, '=IF(N(%s)=0,"",%s)' % (aant, aant))
        o1, oN = t['omzet'][1], t['omzet'][-1]
        koop = 'N(%sE%d)' % (OK, t['koop'])
        w.set('%s9' % cn, '=IF(%s8="","",ROUND(%s+SUMPRODUCT((%s$C$%d:$C$%d<>"")*(%s$C$%d:$C$%d<>0),%s$E$%d:$E$%d),2))'
              % (cn, koop, OK, o1, oN, OK, o1, oN, OK, o1, oN))
        sk = '%s$C$%d' % (HULP, R0)
        w.set('%s10' % cn, '=IF(OR(%s8="",%s=0),"",INT((%s-1)/4))' % (cn, sk, sk))
        w.set('%s11' % cn, '=IF(OR(%s8="",%s=0),"",%s-4*INT((%s-1)/4))' % (cn, sk, sk, sk))
        w.set('%s15' % cp, '=IF(OR(%s8="",N(%s9)=0),"",%s/%s9)' % (cn, cn, koop, cn))
        for lijst, r1 in KOOP_X0.items():
            s = slots[k][lijst]
            for j in range(SLOTS[lijst]):
                hr = s + j
                vis = 'OR(%s$8="",%sB%d="")' % (cn, HULP, hr)
                w.set('%s%d' % (cn, r1 + j), '=IF(%s,"",%sD%d)' % (vis, HULP, hr))
                w.set('%s%d' % (cp, r1 + j), '=IF(%s,"",%sE%d)' % (vis, HULP, hr))
                w.set('%s%d' % (ck, r1 + j), '=IF(%s,"",%sF%d)' % (vis, HULP, hr))


def invoer(inv: Sheet, types, st):
    """kolommen A..J (periodes en cashflow) en verkocht/transport per type; rij 27-30 totalen en controle"""
    a = "'FO - actuals'!"

    def act(soort, jaar, q=None):
        f = 'SUMIFS(%s$K:$K,%s$Q:$Q,"%s",%s$P:$P,%s' % (a, a, soort, a, jaar)
        if q:
            f += ',%s$R:$R,%s' % (a, q)
        return f + ')'
    r1 = INV_R1
    inv.set('A%d' % r1, 0, st['verborgen'])
    inv.set('A%d' % (r1 + 1), 0, st['verborgen'])
    q1 = r1 + 2
    inv.set('B%d' % r1, '=IFERROR(INT(($A$%d-1)/4)-2,"")' % q1, st['vet'])
    inv.set('B%d' % (r1 + 1), '=IFERROR(INT(($A$%d-1)/4)-1,"")' % q1, st['vet'])
    inv.set('C%d' % r1, None)
    inv.set('C%d' % (r1 + 1), None)
    inv.set('D%d' % r1, '=SUMPRODUCT(SUMIFS(%s$K:$K,%s$Q:$Q,"Kosten",%s$P:$P,$B%d-{0,1,2,3,4}))' % (a, a, a, r1), st['getal'])
    inv.set('F%d' % r1, '=-SUMPRODUCT(SUMIFS(%s$K:$K,%s$Q:$Q,"Omzet",%s$P:$P,$B%d-{0,1,2,3,4}))' % (a, a, a, r1), st['getal'])
    inv.set('D%d' % (r1 + 1), '=' + act('Kosten', '$B%d' % (r1 + 1)), st['getal'])
    inv.set('F%d' % (r1 + 1), '=-' + act('Omzet', '$B%d' % (r1 + 1)), st['getal'])
    for r in range(q1, INV_RN + 1):
        inv.set('A%d' % r, '=IFERROR(_xlfn.AGGREGATE(15,6,%s/(%s>0),ROW()-%d),"")' % (KEYK, KEYK, q1 - 1), st['verborgen'])
        inv.set('B%d' % r, '=IF($A%d="","",INT(($A%d-1)/4))' % (r, r), st['vet'])
        inv.set('C%d' % r, '=IF($A%d="","","Q"&($A%d-4*$B%d))' % (r, r, r), st['vet'])
        inv.set('D%d' % r, '=IF($A%d="","",%s+IF($A%d>=p_huidig,IFERROR(INDEX(%s$H$75:$AA$75,MATCH($A%d,%s,0)),0),0))'
                % (r, act('Kosten', '$B%d' % r, '$C%d' % r), r, KK, r, KEYK), st['getal'])
        inv.set('F%d' % r, '=IF($A%d="","",-%s+IF($A%d>=p_huidig,IFERROR(INDEX(%s$G$194:$Z$194,MATCH($A%d,%s,0)),0),0))'
                % (r, act('Omzet', '$B%d' % r, '$C%d' % r), r, OK, r, KEYO), st['getal'])
    tot = INV_RN + 2
    for r in range(r1, INV_RN + 1):
        inv.set('E%d' % r, '=IF(D%d="","",IFERROR(D%d/$D$%d,0))' % (r, r, tot), st['pct'])
        inv.set('G%d' % r, '=IF(F%d="","",IFERROR(F%d/$F$%d,0))' % (r, r, tot), st['pct'])
        inv.set('H%d' % r, '=-D%d+F%d' % (r, r) if r == r1 else '=IF(D%d="","",N(H%d)-D%d+F%d)' % (r, r - 1, r, r), st['getal'])
        inv.set('I%d' % r, '=IF(H%d="","",H%d/1000)' % (r, r), st['getal1'])
        inv.set('J%d' % r, "=IF(OR($B%d=\"\",COUNTIFS('Vorige prognose'!$B:$B,$B%d,'Vorige prognose'!$C:$C,$C%d&\"\")=0),\"\","
                           "SUMIFS('Vorige prognose'!$G:$G,'Vorige prognose'!$B:$B,$B%d,'Vorige prognose'!$C:$C,$C%d&\"\"))"
                % (r, r, r, r, r), st['getal1'])
        for t in types:
            cv, ct = num2col(12 + 2 * (t['k'] - 1)), num2col(13 + 2 * (t['k'] - 1))
            naam = 'Woningtypes!$%s$7' % num2col(3 + 3 * (t['k'] - 1))
            for col, bron in ((cv, t['verkoop']), (ct, t['koop'])):
                rng = '%s$G$%d:$Z$%d' % (OK, bron, bron)
                inv.set('%s%d' % (col, r), '=IF(OR($B%d="",%s=""),"",IF($A%d>0,IFERROR(INDEX(%s,MATCH($A%d,%s,0)),0),0)'
                                           '+IF($A%d=%s,SUMIFS(%s,%s,0),0))'
                        % (r, naam, r, rng, r, KEYO, r, ACT, rng, KEYO), st['aantal'])
    inv.set('C%d' % tot, 'Totaal', st['label_tot'])
    inv.set('D%d' % tot, '=SUM(D%d:D%d)' % (r1, INV_RN), st['getal_tot'])
    inv.set('F%d' % tot, '=SUM(F%d:F%d)' % (r1, INV_RN), st['getal_tot'])
    inv.set('C%d' % (tot + 1), 'PEW', st['tekst'])
    inv.set('D%d' % (tot + 1), "=N('2. Ontw PEW'!G20)*1000", st['getal'])
    inv.set('F%d' % (tot + 1), "=N('2. Ontw PEW'!G27)*1000", st['getal'])
    inv.set('C%d' % (tot + 2), 'Check', st['tekst'])
    inv.set('D%d' % (tot + 2), '=ROUND(D%d-D%d,2)' % (tot, tot + 1), st['getal'])
    inv.set('F%d' % (tot + 2), '=ROUND(F%d-F%d,2)' % (tot, tot + 1), st['getal'])
    inv.set('D%d' % (tot + 3), '=IF(AND(D%d<1000,D%d>-1000),"OK","FOUT")' % (tot + 2, tot + 2), st['vet'])
    inv.set('F%d' % (tot + 3), '=IF(AND(F%d<1000,F%d>-1000),"OK","FOUT")' % (tot + 2, tot + 2), st['vet'])
    return tot


def scenario_kosten(sk: Sheet, st):
    for i, p in enumerate(range(59, 71)):
        r = 8 + i
        sk.set('B%d' % r, '=IF(%sB%d="","",%sB%d&"")' % (KK, p, KK, p))
        sk.set('C%d' % r, '=IF(B%d="","",IF(OR(ISNUMBER(SEARCH({"promotie","verkoop","verhuur","makelaar","marketing"},B%d))),'
                          '"verkoop",IF(OR(ISNUMBER(SEARCH({"aanneemsom","bouwkosten","meerwerk","aansluit","reserve",'
                          '"bouwplaats","nutsvoorziening"},B%d))),"start bouw","vast")))' % (r, r, r))
        hv = '(%sC%d<>"Voorstel")' % (KK, p), '(%sC%d="Voorstel")' % (KK, p)
        gr = lambda q: 'N(%sK%d)+SUM(%sH%d:J%d)+SUMIFS(%sL%d:AA%d,%sL$96:AA$96,"<="&%s)' % (KK, q, KK, q, q, KK, q, q, KK, ACT)
        sk.set('E%d' % r, '=IF(B%d="","",%s*(%s)+%s*(%s))' % (r, hv[0], gr(p), hv[1], gr(p + 41)))
        for pi in range(60):
            c = num2col(9 + pi)
            key = 'INDEX(Invoer!$A$8:$A$67,%s$6)' % c
            val = lambda q: 'IFERROR(INDEX(%s$H$%d:$AA$%d,MATCH(%s,%s,0)),0)' % (KK, q, q, key, KEYK)
            sk.set('%s%d' % (c, r), '=IF(OR($B%d="",N(%s)<=%s),"",%s*%s+%s*%s)' % (r, key, ACT, hv[0], val(p), hv[1], val(p + 41)))
