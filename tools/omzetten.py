"""Een oud FO omzetten naar het standaard-FO (dezelfde logica als de macro 'Oud FO inlezen').

    python tools/omzetten.py oud_FO.xlsx standaard_gevuld.xlsx

Wat er overgaat (waarden zoals Excel ze in het oude FO heeft opgeslagen):
  exports            FO - vorig, FO - huidig, FO - actuals, Prijslijst dump, Behind the scens (alles)
  prijslijst         per bouwnummer: bouwnummer, blok, type en de handmatige kolommen (kadastrale kosten, rente, korting)
  per woningtype     naam, type in de prijslijst, aantal en bedragen (handmatig als de prijslijst ze niet verklaart),
                     termijnschema (naam, %), planning bouwtermijnen, verkocht, transport (= grondtermijn), kadastrale
                     kosten/rente (aantallen), kosten: budget aanneemsom, KAO, % en planning (als die afwijkt)
  overige            overige opbrengsten en per type 'overige'/'verschillen' in €, DAEB-blok, kostenposten (prognose)
  PEW                meest recente FD (kolom E) en toelichtingen
  1. Cashflow        voor de vorige prognose en de controle
Kolommen gaan op kwartaal (jaar x 4 + Q); wat vóór het eerste kwartaal van het standaard-FO valt, komt in 'actuals'.
"""
from __future__ import annotations

import os
import re
import sys

import openpyxl
from openpyxl.utils import column_index_from_string as ci

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import opzet as opzet_fo  # noqa: E402
import standaard_fo as S  # noqa: E402


def _t(v):
    return v.strip() if isinstance(v, str) else ('' if v is None else str(v))


def _n(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def sleutels(ws, jr, qr, c1, c2):
    """kolom -> kwartaalindex (jaar*4+q); 0 voor kolommen zonder kwartaal"""
    out, jaar = {}, None
    for c in range(c1, c2 + 1):
        if isinstance(ws.cell(jr, c).value, (int, float)):
            jaar = int(ws.cell(jr, c).value)
        q = _t(ws.cell(qr, c).value)
        out[c] = jaar * 4 + int(q[1]) if jaar and re.match(r'Q[1-4]$', q) else 0
    return out


class Doel:
    """schrijft in het standaard-FO; kwartaalkolommen op sleutel"""

    def __init__(self, wb, jaar):
        self.wb, self.k1 = wb, jaar * 4 + 1
        self.log = []

    def kol(self, key):
        """standaardkolom voor een kwartaalindex (None = valt erbuiten), 0 = actuals"""
        if key == 0 or key < self.k1:
            return S.C_ACT
        i = key - self.k1
        return S.C_Q1 + i if i < S.NQ else None

    def rij(self, blad, r_doel, waarden, keys, alleen_q=False, nul_weg=True):
        """waarden {oude kolom: waarde} -> rij r_doel; niet-kwartalen opgeteld in actuals"""
        ws = self.wb[blad]
        som = {}
        for c, v in waarden.items():
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                continue
            k = self.kol(keys.get(c, 0))
            if k is None:
                if v:
                    self.log.append('%s rij %d: %s na het laatste kwartaal valt weg' % (blad, r_doel, v))
                continue
            if alleen_q and k == S.C_ACT:
                continue
            som[k] = som.get(k, 0) + v
        for k, v in som.items():
            ws.cell(r_doel, k).value = v if (v or not nul_weg) else None


def omzetten(oud, uit, project=None):
    oz = opzet_fo.lees(oud)
    wf = openpyxl.load_workbook(oud)
    wv = openpyxl.load_workbook(oud, data_only=True)
    ob, ko = oz['opbrengsten'], oz['kosten']
    # startjaar: eerste kwartaal in 1. Cashflow, anders het eerste kwartaal van CF - opbrengsten
    qj = [j for j, q, _ in oz['cashflow']['periodes'] if q]
    o_v = wv['CF - opbrengsten']
    okeys = sleutels(o_v, ob['jaarrij'], ob['euro_kop'], ci(ob['eerste']), ci(ob['laatste']))
    jaar = min(qj) if qj else min(k for k in okeys.values() if k) // 4
    cfw = wv['1. Cashflow']
    cf = {}
    for j, q, r in oz['cashflow']['periodes']:
        cf[(j, q or None)] = {c: cfw['%s%d' % (c, r)].value for c in 'DFHJ'}
    # laatste actuals (kwartaalindex) en woningtypes zonder woningen: ingeklapt
    av = wv['FO - actuals'] if 'FO - actuals' in wv.sheetnames else None
    act_max = per_max = 0
    if av is not None:
        for r in range(2, av.max_row + 1):
            p_, o_ = av.cell(r, 16).value, av.cell(r, 15).value
            try:
                k_ = int(float(p_)) * 4 + (int(float(o_)) - 1) // 3 + 1
            except (TypeError, ValueError):
                continue
            if 1 <= int(float(o_)) <= 12:
                act_max = max(act_max, k_)
                per_max = max(per_max, int(float(p_)) * 100 + int(float(o_)))
    # afgesloten t/m: loopt het kwartaal van de laatste boekingsperiode nog (niet de laatste maand van het kwartaal), dan
    # is dat het huidige kwartaal (Start rij 8) en is het kwartaal ervoor het laatste afgesloten kwartaal
    if per_max % 100 % 3:
        act_max -= 1
    leeg_t = [t['k'] for t in ob['types'] if not t['won']] + list(range(len(ob['types']) + 1, S.NT + 1))
    S.maak(uit, leeg=True, jaar=jaar, cashflow=cf, verborgen_types=leeg_t, verleden_tot=act_max or None)
    wb = openpyxl.load_workbook(uit)
    d = Doel(wb, jaar)
    log = d.log
    if len(ob['types']) > S.NT:
        raise ValueError('%d woningtypes; het standaard-FO heeft er %d' % (len(ob['types']), S.NT))

    # ---- exports en koppeltabel: alles overnemen --------------------------------------------------------
    breedte = {'FO - vorig': 21, 'FO - huidig': 21, 'FO - actuals': 16, 'Prijslijst dump': 9, 'Behind the scens': 0}
    for n in breedte:
        if n not in wf.sheetnames:
            log.append('%s staat niet in het oude FO' % n)
            continue
        a, av, b = wf[n], wv[n], wb[n]
        # de export zelf als waarden (zoals uit SAP/SF); de hulpkolommen rechts ervan blijven die van het sjabloon
        # (Behind the scens: alles van het oude FO)
        w = breedte[n] or 10 ** 6
        for row in b.iter_rows():
            for c in row:
                if c.column <= w and not type(c).__name__ == 'MergedCell':
                    c.value = None
        for row in a.iter_rows():
            for c in row:
                if c.column <= w and c.value is not None and type(b.cell(c.row, c.column)).__name__ != 'MergedCell':
                    b.cell(c.row, c.column).value = av.cell(c.row, c.column).value if breedte[n] else c.value

    # ---- prijslijst per bouwnummer ------------------------------------------------------------------------
    pf, pv = wf['Prijslijst'], wv['Prijslijst']
    kop = next((r for r in range(1, 20) if _t(pv.cell(r, 1).value).lower() == 'bouwnummer'), None)
    hd = {_t(pv.cell(kop, c).value).lower(): c for c in range(1, pv.max_column + 1)} if kop else {}
    P = wb['Prijslijst']
    if _t(pv['A1'].value).lower().startswith('datum'):
        P['C1'] = pv['C1'].value
    bn = []
    if kop:
        for r in range(kop + 1, pv.max_row + 1):
            if pv.cell(r, 1).value in (None, ''):
                continue
            bn.append(r)
        for i, r in enumerate(bn[:S.P_BNN - S.P_BN1 + 1]):
            rr = S.P_BN1 + i
            P['A%d' % rr] = pv.cell(r, 1).value
            for naam, col in (('termijnen / blok', 'B'), ('type', 'C'), ('kadastrale kosten', 'J'), ('bouwrente', 'Q'),
                              ('aanvullende rente', 'R'), ('korting', 'S')):
                if naam in hd:
                    v = pf.cell(r, hd[naam]).value
                    if isinstance(v, str) and v.startswith('='):
                        v = pv.cell(r, hd[naam]).value      # formule (zoals de kadastrale kosten): de waarde
                    P['%s%d' % (col, rr)] = v
        log.append('prijslijst: %d bouwnummers' % len(bn))

    def type_van(r):
        """zoals kolom U in het sjabloon: elk bouwnummer bij precies één type; exact op Type, dan Type in een lijst
        (met ;), dan exact op Termijnen / blok, dan blok in een lijst, dan '*'; bij gelijkspel het eerste type"""
        sls = [str(x or '').strip().lower() for x in (oz.get('prijslijst') or {}).get('sleutels') or []]
        for kol in ('type', 'termijnen / blok'):
            w = _t(pv.cell(r, hd[kol]).value).lower() if kol in hd else ''
            if not w:
                continue
            for i, x in enumerate(sls):
                if x == w:
                    return i + 1
            for i, x in enumerate(sls):
                if w in [y.strip() for y in x.split(';')]:
                    return i + 1
        return sls.index('*') + 1 if '*' in sls else None

    def hoort_bij(r, sleutel):
        sls = [str(x or '').strip().lower() for x in (oz.get('prijslijst') or {}).get('sleutels') or []]
        k_ = sls.index(str(sleutel).strip().lower()) + 1 if str(sleutel).strip().lower() in sls else None
        return k_ is not None and type_van(r) == k_

    def uit_prijslijst(sleutel, kolnaam):
        if not sleutel or 'type' not in hd or kolnaam not in hd:
            return None
        return sum(_n(pv.cell(r, hd[kolnaam]).value) for r in bn if hoort_bij(r, sleutel))

    # ---- per woningtype -----------------------------------------------------------------------------------
    O, K = wb['CF - opbrengsten'], wb['CF - kosten']
    of = wf['CF - opbrengsten']
    sl = (oz.get('prijslijst') or {}).get('sleutels') or []
    ov2 = {}
    via_type = set()                    # 'parke' / 'commerc': deze omzet loopt via een woningtype, niet via Overige
    # €-blokken per type ('Termijnen n in €') in het oude FO: een type met woningen maar zonder bedragen daar (bijvoorbeeld
    # parkeren of commercieel) heeft zijn omzet onder 'Overige opbrengsten in €'; dan krijgt het type prijs 0 (de kosten
    # blijven) en komt de omzet via de regels parkeren / commercieel
    oc_ = range(ci(ob['eerste']), ci(ob['laatste']) + 1)
    eu_rijen = [r for r in range(ob['euro_kop'], ob['totaal'])
                if _t(o_v.cell(r, 2).value).lower().startswith('termijnen')
                and '€' in _t(o_v.cell(r, 2).value) + _t(o_v.cell(r, 3).value)]

    def eu_som(i):
        if i >= len(eu_rijen):
            return None
        r, som = eu_rijen[i] + 1, 0
        while _t(o_v.cell(r, 2).value) and r < eu_rijen[i] + 30:
            som += sum(abs(_n(o_v.cell(r, c).value)) for c in oc_)
            r += 1
        return som
    for t in ob['types']:
        k = t['k']
        tr_ = S.P_TYPE1 + k - 1
        if not t['won']:
            log.append('type %d: geen woningen, niet overgenomen' % k)
            continue
        titel = _t(o_v.cell(t['hb'], 2).value)
        naam = titel.split(' - ', 1)[1].strip() if ' - ' in titel else (_t(o_v.cell(t['hb'], 3).value) or
                                                                       re.sub(r'(?i)^termijnen\s*', '', titel) or 'Type %d' % k)
        P['B%d' % tr_] = naam
        sleutel = sl[k - 1] if len(sl) >= k else None
        P['C%d' % tr_] = sleutel
        won = t['won']
        # omzetrijen indelen: termijnen (horen bij een planningsrij), kopersmeerwerk, kadastrale kosten/rente, overige
        termijnen, kmw25, kmw75, kad, overig = [], None, None, None, []
        for r in t['omzet'][1:]:
            lab = _t(o_v.cell(r, 2).value).lower()
            if 'meerwerk' in lab and '25' in lab:
                kmw25 = r
            elif 'meerwerk' in lab and '75' in lab:
                kmw75 = r
            elif (r - t['d']) in t['plan']:
                termijnen.append(r)
            elif 'kadast' in lab or 'rente' in lab:
                kad = r
            else:
                overig.append(r)
        if len(termijnen) > S.NTERM:
            log.append('type %d: %d termijnen, het standaard-FO heeft er %d' % (k, len(termijnen), S.NTERM))
            termijnen = termijnen[:S.NTERM]
        # bedragen (totaal per type) zoals het oude FO rekende
        oud = {'aantal': won, 'grond': _n(o_v.cell(t['koop'], 4).value),
               'kmw': _n(o_v.cell(kmw25, 4).value) / 0.25 if kmw25 else (_n(o_v.cell(kmw75, 4).value) / 0.75 if kmw75 else 0),
               'kadrente': _n(o_v.cell(kad, 4).value) if kad else 0}
        basis = [(_n(o_v.cell(r, 4).value) / _n(o_v.cell(r, 3).value)) for r in termijnen if _n(o_v.cell(r, 3).value)]
        oud['aanneemkao'] = basis[0] if basis else 0
        pl_ = {'aantal': (sum(1 for r in bn if hoort_bij(r, sleutel)) if sleutel and 'type' in hd else None),
               'grond': uit_prijslijst(sleutel, 'grondkosten start verkoop (facturatie)'),
               'aanneemkao': (uit_prijslijst(sleutel, 'aanneemsom start verkoop') or 0) + (uit_prijslijst(sleutel, 'totaal woonwensen in von') or 0) if sleutel else None,
               'kmw': uit_prijslijst(sleutel, 'totaal woonwensen in ob'),
               'kadrente': sum(uit_prijslijst(sleutel, x) or 0 for x in ('kadastrale kosten', 'bouwrente', 'aanvullende rente', 'korting')) if sleutel else None}
        for x, col in S.P_HAND.items():
            if pl_[x] is None or abs(pl_[x] - oud[x]) > 1:
                P['%s%d' % (col, tr_)] = oud[x]
                if pl_[x] is not None:
                    log.append('type %d (%s): %s uit het oude FO (%.0f) wijkt af van de prijslijst (%.0f): handmatig gezet'
                               % (k, naam, x, oud[x], pl_[x]))
        if t['won'] and eu_som(k - 1) == 0:
            # parkeren / commercieel met een eigen termijnschema: als echt type via de termijnen (de regel onder
            # 'Overige opbrengsten in €' van het oude FO vervalt dan, anders telt het dubbel)
            sleutel_ov = 'parke' if 'parke' in naam.lower() else 'commerc' if 'commerc' in naam.lower() else ''
            if sleutel_ov and oud['aanneemkao'] > 0:
                via_type.add(sleutel_ov)
                log.append('type %d (%s): omzet via het termijnschema van het type (%.0f); de regel onder Overige opbrengsten '
                           'van het oude FO is niet overgenomen, de kwartalen kunnen afwijken van het oude FO'
                           % (k, naam, oud['aanneemkao']))
            else:
                for x in ('grond', 'aanneemkao', 'kmw', 'kadrente'):
                    P['%s%d' % (S.P_HAND[x], tr_)] = 0
                log.append('type %d (%s): geen bedragen in het €-blok van het oude FO; de omzet staat onder Overige opbrengsten '
                           '(prijs 0, de kosten blijven)' % (k, naam))
        # termijnschema
        a = S.p_aans(k)
        for i, r in enumerate(termijnen, 1):
            P['%s%d' % (S.P_TERM_NAAM, a + 1 + i)] = _t(o_v.cell(r, 2).value)
            P['%s%d' % (S.P_TERM_PCT, a + 1 + i)] = _n(o_v.cell(r, 3).value)
        # planning, verkocht, transport, kadastrale kosten
        oc = range(ci(ob['eerste']), ci(ob['laatste']) + 1)
        val = lambda ws, r: {c: ws.cell(r, c).value for c in oc}
        for i, r in enumerate(termijnen, 1):
            d.rij('CF - opbrengsten', S.plan(k, i), val(o_v, r - t['d']), okeys)
        d.rij('CF - opbrengsten', S.vk_rij(k), val(o_v, t['verkoop']), okeys, alleen_q=True)
        if not sleutel:
            O['D%d' % S.vk_rij(k)] = _n(o_v.cell(t['verkoop'], 4).value)
        d.rij('CF - opbrengsten', S.oz_rij(k, 0), val(o_v, t['koop']), okeys)
        if kad:
            d.rij('CF - opbrengsten', S.oz_rij(k, S.NTERM + 3), val(o_v, kad), okeys)
        # overige omzetrijen (aantal x € per woning) naar de projectregel 'overige / verschillen'
        for r in overig:
            e = _n(o_v.cell(r, 5).value)
            for c in oc:
                k2 = d.kol(okeys.get(c, 0))
                if e and k2 and _n(o_v.cell(r, c).value):
                    ov2[k2] = ov2.get(k2, 0) + _n(o_v.cell(r, c).value) * e
        log.append('type %d (%s): %d termijnen, %d won%s' % (k, naam, len(termijnen), won,
                                                             ', overige: ' + ', '.join(_t(o_v.cell(r, 2).value) for r in overig) if overig else ''))
        # kosten: budget aanneemsom, KAO, % en planning (als die afwijkt van de opbrengsten)
        kv = wv['CF - kosten']
        krows = t.get('kosten') or []
        if krows:
            kop_r = krows[0] - 1
            b = S.kb(k)
            K['D%d' % b] = _n(kv.cell(kop_r - 1, 4).value)
            K['F%d' % b] = _n(kv.cell(kop_r - 1, 6).value)
            kkeys = sleutels(kv, kop_r - 1, kop_r, ci(ko['eerste']), ci(ko['laatste']))
            kc = range(ci(ko['eerste']), ci(ko['laatste']) + 1)
            kterm = [r for r in krows if 'meerwerk' not in _t(kv.cell(r, 2).value).lower()][:S.NTERM]
            for i, r in enumerate(kterm, 1):
                K['C%d' % S.kterm(k, i)] = _n(kv.cell(r, 3).value)
                kw = {c: kv.cell(r, c).value for c in kc}
                pw = {c2: o_v.cell(termijnen[i - 1] - t['d'], c2).value for c2 in oc} if i <= len(termijnen) else {}
                som_k = sum(_n(v) for v in kw.values())
                som_p = sum(_n(v) for v in pw.values())
                # eigen planning (HWB) alleen als die afwijkt van de planning van de opbrengsten
                if abs(som_k - som_p) > 1e-9 or any(abs(_n(kw.get(c, 0)) - _n(pw.get(c2, 0))) > 1e-9
                                                    for c, c2 in zip(kc, oc) if kkeys.get(c) == okeys.get(c2)):
                    for c in range(S.C_ACT, S.C_QN + 1):
                        K.cell(S.kterm(k, i), c).value = None
                    d.rij('CF - kosten', S.kterm(k, i), kw, kkeys)
                    log.append('type %d: kostenplanning termijn %d wijkt af van de opbrengsten: overgenomen' % (k, i))
    # ---- aansluiting: SKO-bedragen per regel (kolom B van de vergelijkingsblokken) -----------------------------
    av_ = wv[oz['aansluiting_blad']]
    for i, blk in enumerate((oz.get('aansluiting') or [])[:S.NT]):
        a0 = S.p_aans(i + 1)
        for r, lab in blk['regels']:
            if lab in S.P_REGELS and isinstance(av_.cell(r, 2).value, (int, float)):
                P['B%d' % (a0 + 2 + S.P_REGELS.index(lab))] = av_.cell(r, 2).value
            # regels zonder koppeling met de prijslijst (VON verhogingen, commercieel, parkeren): het bedrag dat in
            # het oude FO in de kolom Prijslijst stond
            if lab in S.P_REGELS_HAND and isinstance(av_.cell(r, 4).value, (int, float)) and av_.cell(r, 4).value:
                P['D%d' % (a0 + 2 + S.P_REGELS.index(lab))] = av_.cell(r, 4).value
    # ---- overige opbrengsten (project) ------------------------------------------------------------------
    # parkeren en commercieel naar hun eigen regel; de rest opgeteld als 'Overige opbrengsten (uit het oude FO)'
    ov1, ovp, ovc = {}, {}, {}
    ov = next((r for r in range(ob['euro_kop'], ob['totaal']) if _t(o_v.cell(r, 2).value).lower().startswith('overige opbrengsten in')), None)
    oc = range(ci(ob['eerste']), ci(ob['laatste']) + 1)
    if ov:
        for r in range(ov + 1, ov + 4):
            lab = _t(of.cell(r, 2).value).lower()
            if not lab:
                break
            if any(x in lab for x in via_type):
                continue                # loopt via het woningtype
            doel = ovp if 'parkeer' in lab or 'parkeren' in lab else ovc if 'commerc' in lab else ov1
            for c in oc:
                k2 = d.kol(okeys.get(c, 0))
                if k2 and _n(o_v.cell(r, c).value):
                    doel[k2] = doel.get(k2, 0) + _n(o_v.cell(r, c).value)
    if ovp and ov2:                     # de regel parkeren is al bezet door de verschillen per type
        for k2, v in ovp.items():
            ov1[k2] = ov1.get(k2, 0) + v
        ovp = {}
    if ovc and ov1:
        for k2, v in ovc.items():
            ov1[k2] = ov1.get(k2, 0) + v
        ovc = {}
    regels = ((S.OV1, ov1, 'Overige opbrengsten (uit het oude FO)'), (S.OV1, ovc, 'Opbrengsten commercieel (uit het oude FO)'),
              (S.OV2, ov2, 'Overige / verschillen per type (uit het oude FO)'), (S.OV2, ovp, 'Opbrengsten parkeren (uit het oude FO)'))
    for rr, som, lab in regels:
        if som:
            O['B%d' % rr] = lab
            O['C%d' % (rr - 2)] = 'handmatig'          # de regel in de cashflow neemt de handmatige bedragen
            for k2, v in som.items():
                O.cell(rr, k2).value = v
    # ---- DAEB --------------------------------------------------------------------------------------------
    if ob['daeb']:
        daeb_omzetten(o_v, ob, okeys, O, d, wv)
    # ---- kostenposten (prognose; Aanneemsom en Kopersmeerwerk volgen de termijnen) ----------------------
    kv = wv['CF - kosten']
    kkeys = sleutels(kv, ko['onderdelen'] - 1, ko['onderdelen'], ci(ko['eerste']), ci(ko['laatste']))
    kc = range(ci(ko['eerste']), ci(ko['laatste']) + 1)
    verschoven = 0
    for i, r in enumerate(ko['posten'][:S.NPOST]):
        naam = _t(kv.cell(r, 2).value)
        rr = S.K_POST1 + i
        toel = kv['%s%d' % (ko['toelichting'], r)].value
        if toel:
            K['%s%d' % (S.TOEL, rr)] = toel
        if naam.lower() in ('aanneemsom', 'kopersmeerwerk'):
            continue
        w = {c: kv.cell(r, c).value for c in kc if c != (ci(ko['werkelijk']) if ko['werkelijk'] else -1)}
        d.rij('CF - kosten', rr, w, kkeys, alleen_q=True)
        # prognose in een kwartaal dat volgens de SAP-export al is afgesloten (de cashflow neemt daar de actuals):
        # naar het eerste prognosekwartaal, anders telt die prognose nergens mee
        if act_max >= d.k1 and d.kol(act_max + 1):
            for q in range(d.k1, act_max + 1):
                v = K.cell(rr, d.kol(q)).value
                if isinstance(v, (int, float)) and v:
                    K.cell(rr, d.kol(q)).value = None
                    naar = K.cell(rr, d.kol(act_max + 1))
                    naar.value = (naar.value or 0) + v
                    verschoven += v
    if verschoven:
        log.append('CF - kosten: %.0f gepland in kwartalen die volgens de SAP-export al zijn afgesloten (t/m Q%d %d) '
                   'verschoven naar het eerste prognosekwartaal' % (verschoven, (act_max - 1) % 4 + 1, (act_max - 1) // 4))
    # ---- PEW: meest recente FD en toelichting ----------------------------------------------------------------
    pw_ = oz['pew']
    a, b = wf['2. Ontw PEW'], wb['2. Ontw PEW']
    for r in range(pw_['eerste'], pw_['opbrengsten_totaal']):
        if r in (pw_['kosten_totaal'],):
            continue
        b['E%d' % r] = a['%s%d' % (pw_['fd'], r)].value
        b['K%d' % r] = a['%s%d' % (pw_['toelichting'], r)].value
    extra = [c for c in pw_['kolommen'] if c not in (pw_['fd'], pw_['vorig'], pw_['huidig'])]
    if len(extra) > 2:
        log.append('2. Ontw PEW: extra kolom(men) %s van het oude FO niet overgenomen' % extra[:-2])
    # startdatum VENA
    O['E%d' % S.PER] = o_v['E%d' % ob['periode']].value
    # aantal woningen per type als waarde voor de bouw (de cel zelf is een formule zonder berekende waarde)
    h = wb.create_sheet('_omzetting')
    h.sheet_state = 'hidden'
    # 2B. Ontw PEW (gemengd project): A = concept SKO DAEB, B = concept SKO niet-DAEB en de toelichting, per post
    # (kolom D:G van _omzetting; de bouw zet ze op het nieuwe tabblad 2B. Ontw PEW)
    if '2B. Ontw PEW' in wv.sheetnames:
        bv = wv['2B. Ontw PEW']
        ka = kb_ = kt = kr = None
        for r in range(1, 12):
            for c in range(3, 20):
                t = _t(bv.cell(r, c).value).lower()
                # alleen kolomkoppen ('Concept SKO ...'), niet een toelichting die toevallig 'niet-DAEB' noemt
                if 'sko' in t and ('niet-daeb' in t or 'niet daeb' in t) and not kb_:
                    kb_, kr = c, r
                elif 'sko' in t and 'daeb' in t and 'niet' not in t and not ka:
                    ka, kr = c, r
                elif t == 'toelichting' and not kt:
                    kt = c
        n2 = 0
        if ka and kb_:
            for r in range(kr + 1, kr + 30):
                lab = _t(bv.cell(r, 3).value)
                if not lab or lab.lower().startswith(('totaal', 'resultaat', 'check')):
                    continue
                va, vb = bv.cell(r, ka).value, bv.cell(r, kb_).value
                n2 += 1
                h.cell(n2, 4).value = lab
                h.cell(n2, 5).value = va if isinstance(va, (int, float)) else None
                h.cell(n2, 6).value = vb if isinstance(vb, (int, float)) else None
                h.cell(n2, 7).value = bv.cell(r, kt).value if kt else None
            log.append('2B. Ontw PEW: A (DAEB) en B (niet-DAEB) overgenomen voor %d posten' % n2)
    for t in ob['types']:
        h.cell(t['k'], 1).value = t['k']
        h.cell(t['k'], 2).value = t['won']
    wb.save(uit)
    mp = ob.get('mijlpalen') or {}
    return {'log': log, 'jaar': jaar, 'oplevering': mp.get('oplevering'), 'startbouw': mp.get('startbouw'),
            'types': len(ob['types'])}


def daeb_omzetten(o_v, ob, okeys, O, d, wv):
    """DAEB-blok van het oude FO: componenten, termijnregels (per groep), aantal en verkoop"""
    r0 = ob['daeb']
    oc = range(ci(ob['eerste']), ci(ob['laatste']) + 1)
    comp = [r for r in range(r0 - 3, r0) if _t(o_v.cell(r, 3).value) and _n(o_v.cell(r, 5).value)]
    for i, r in enumerate(comp[:2]):
        O['C%d' % (S.DAEB_KOP - 3 + i)] = _t(o_v.cell(r, 3).value)
        O['E%d' % (S.DAEB_KOP - 3 + i)] = _n(o_v.cell(r, 5).value)
    groepen, g, r = [], [], r0 + 1
    act = ci(ob.get('actuals') or 'J')
    while r < r0 + 40:
        lab = _t(o_v.cell(r, 2).value)
        if lab.lower().startswith('totaal'):
            break
        if not lab:
            if g:
                groepen.append(g)
                g = []
            if not _t(o_v.cell(r + 1, 2).value):
                break
        else:
            g.append(r)
        r += 1
    if g:
        groepen.append(g)
    for gi, grp in enumerate(groepen[:2]):
        for i, r in enumerate(grp[:8]):
            rr = S.DAEB_KOP + 1 + gi * 9 + i
            O['B%d' % rr] = _t(o_v.cell(r, 2).value)
            O['C%d' % rr] = o_v.cell(r, 3).value if isinstance(o_v.cell(r, 3).value, (int, float)) else None
            w = {c: o_v.cell(r, c).value for c in oc}
            if act not in oc:
                w[act] = o_v.cell(r, act).value
            d.rij('CF - opbrengsten', rr, w, okeys)
    # aantal en verkoop: regels in het blok, anders Verkooptempo (de DAEB-rij of de rij na de koopwoningtypes)
    aantal, vk = None, None
    for r in range(r0, r0 + 60):
        lab = _t(o_v.cell(r, 2).value).lower()
        if lab.startswith('aantal daeb'):
            aantal = o_v.cell(r, 4).value
        if lab.startswith('verkoop en levering daeb'):
            vk = ({c: o_v.cell(r, c).value for c in oc}, okeys)
    if (aantal is None or vk is None) and 'Verkooptempo' in wv.sheetnames:
        v = wv['Verkooptempo']
        kop = next((r for r in range(1, 30) if _t(v.cell(r, 1).value).lower() == 'verkooptempo'), None)
        if kop:
            rij = next((r for r in range(kop + 1, kop + 20) if 'daeb' in _t(v.cell(r, 1).value).lower()), None)
            if rij is None:
                r = kop + 1 + len(ob['types'])
                rij = r if _t(v.cell(r, 1).value) and not _t(v.cell(r, 1).value).lower().startswith('totaal') else None
            if rij:
                acol = next((c for c in range(1, 40) if _t(v.cell(kop, c).value).lower() == 'aantal'), 3)
                if aantal is None:
                    aantal = v.cell(rij, acol).value
                if vk is None:
                    vk = ({c: v.cell(rij, c).value for c in range(2, 40) if c != acol}, sleutels(v, kop - 1, kop, 2, 40))
    O['D%d' % S.DAEB_AANTAL] = aantal
    if vk:
        w, keys = vk
        w = {c: x for c, x in w.items() if keys.get(c)}
        d.rij('CF - opbrengsten', S.DAEB_VERKOOP, w, keys)


if __name__ == '__main__':
    r = omzetten(sys.argv[1], sys.argv[2])
    print('\n'.join('  ' + x for x in r['log']))
    print('startjaar %s, oplevering %s' % (r['jaar'], r['oplevering']))
