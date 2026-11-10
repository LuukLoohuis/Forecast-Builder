"""De opzet van een FO aflezen: welke rijen en kolommen waar staan. Elk FO wijkt wat af (aantal woningtypes, rijen,
kwartalen t/m 2029 of 2030, extra kolommen op 2. Ontw PEW); het werkbestand wordt op deze opzet gebouwd.

    opzet = lees(fo_pad)   -> dict (ook als JSON naast het werkbestand bewaard, voor controleer.py)
"""
from __future__ import annotations

import re

import openpyxl
from openpyxl.utils import column_index_from_string as ci
from openpyxl.utils import get_column_letter as L


def _t(v):
    return v.strip() if isinstance(v, str) else ('' if v is None else str(v))


def _getal(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


class Fout(Exception):
    pass


def lees(pad):
    wf = openpyxl.load_workbook(pad)
    wv = openpyxl.load_workbook(pad, data_only=True)
    o = {'bladen': [(ws.title, ws.sheet_state) for ws in wf.worksheets]}
    o['opbrengsten'] = _opbrengsten(wf['CF - opbrengsten'], wv['CF - opbrengsten'])
    o['kosten'] = _kosten(wf['CF - kosten'], wv['CF - kosten'], o['opbrengsten'])
    o['pew'] = _pew(wv['2. Ontw PEW'])
    o['cashflow'] = _cashflow(wv['1. Cashflow'])
    o['verkooptempo'] = _verkooptempo(wv['Verkooptempo']) if 'Verkooptempo' in wv.sheetnames else None
    o['hulpkolommen'] = _hulp(wf)
    ab = 'Aansluiting opbrengsten' if 'Aansluiting opbrengsten' in wf.sheetnames else 'Prijslijst'
    o['aansluiting_blad'] = ab
    o['standaard'] = ab == 'Prijslijst'
    o['aansluiting'] = _aansluiting(wv[ab])
    o['aansluiting_max'] = (max(r for b in o['aansluiting'] for r, _ in b['regels']) + 1) if o['standaard'] \
        else wf[ab].max_row
    o['prijslijst'] = _prijslijst(wf)
    return o


def _kopkolommen(ws, r):
    """kolommen met Q1..Q4 in rij r"""
    return [c for c in range(1, ws.max_column + 1) if _t(ws.cell(r, c).value) in ('Q1', 'Q2', 'Q3', 'Q4')]


def _opbrengsten(f, v):
    get = lambda r, c: v.cell(r, c).value
    form = lambda r, c: f.cell(r, c).value
    maxr = f.max_row
    vk = next(r for r in range(2, maxr) if _t(get(r, 2)).lower() == 'verkooptempo')
    kop = [r for r in range(2, vk) if _t(get(r - 1, 6)).lower() == '# won']
    omz = [r for r in range(vk, maxr) if _t(get(r - 1, 6)).lower() == '# won']
    if kop and len(omz) < len(kop):
        # terugval: een omzetblok zonder '# won' in kolom F (wel het aantal in G erboven), herkend aan de kop '%' / '€',
        # na het verkooptempo en vóór het financiële blok ('Financieel' / 'Termijnen 1 in €')
        eind = next((r for r in range(vk, maxr) if _t(get(r, 2)).lower().startswith('financieel')
                     or (_t(get(r, 2)).lower().startswith('termijnen') and '€' in _t(get(r, 2)))), maxr)
        extra = [r for r in range(vk + 1, eind) if r not in omz and _t(get(r, 2)) and _t(get(r, 3)).startswith('%')
                 and _t(get(r, 4)) == '€' and isinstance(get(r - 1, 7), (int, float))]
        omz = sorted(omz + extra)
    if not kop or len(kop) != len(omz):
        raise Fout('CF - opbrengsten: %d bouwplanningsblokken en %d omzetblokken' % (len(kop), len(omz)))
    qk = _kopkolommen(v, kop[0])
    eerste, laatste = 7, max(qk)
    check = next((c for c in range(laatste + 1, laatste + 6) if _t(get(kop[0], c)).lower() == 'check'), laatste + 2)
    types = []
    for i, (hb, oh) in enumerate(zip(kop, omz)):
        plan, d, tr = [], None, None
        r = hb + 1
        while _t(form(r, 2)):
            m = re.match(r'=B(\d+)$', _t(form(r, 2)))
            if not m:
                break
            if d is None:
                d = int(m.group(1)) - r
            plan.append(r)
            r += 1
        tr = r
        if any(_t(form(tr, c)) for c in range(2, check + 3)):
            raise Fout('CF - opbrengsten: geen lege rij onder de bouwplanning van type %d (rij %d)' % (i + 1, tr))
        orows = []
        r = oh + 1
        while _t(form(r, 2)):
            orows.append(r)
            r += 1
        won = get(hb - 1, 7)
        if won is None and '_omzetting' in v.parent.sheetnames:       # standaard-FO na omzetten: aantal als hint
            won = v.parent['_omzetting'].cell(i + 1, 2).value
        types.append({'k': i + 1, 'hb': hb, 'plan': plan, 'd': d, 'tr': tr, 'omzet_kop': oh, 'omzet': orows,
                      'koop': orows[0], 'verkoop': vk + 1 + i, 'won': won if isinstance(won, (int, float)) else 0,
                      # waarden die Excel in het FO heeft opgeslagen: verkocht en grondtermijn (= transport) per kolom
                      'verkoop_waarden': [_getal(get(vk + 1 + i, c)) for c in range(7, max(qk) + 1)],
                      'grond_waarden': [_getal(get(orows[0], c)) for c in range(7, max(qk) + 1)]})
    euro = next(r for r in range(omz[-1], maxr) if _t(get(r, 2)).lower().startswith('termijnen') and '€' in _t(get(r, 2)))
    periode = next(r for r in range(euro, maxr + 1) if _t(get(r, 2)).lower().startswith('periode in cashflow'))
    totaal = max(r for r in range(euro, periode) if _t(get(r, 2)).lower().startswith('totaal opbrengsten'))
    # DAEB-blok (titel met 'daeb', geen typeblok)
    daeb = next((r for r in range(2, maxr + 1) if 'daeb' in _t(get(r, 2)).lower()
                 and not _t(get(r, 2)).lower().startswith(('omzet hvg', 'termijnen '))), None)
    # correctieregel: een lege rij vlak boven het totaal
    leeg = lambda r: not any(_t(form(r, c)) for c in range(2, check + 3))
    corr = next((r for r in range(totaal - 2, totaal - 8, -1) if leeg(r)), None)
    if corr is None:
        raise Fout('CF - opbrengsten: geen lege rij boven het totaal (rij %d) voor de correctieregel' % totaal)
    kop_tekst, mijlpaal = {}, {}
    jaar, jaren = None, {}
    for c in range(eerste, laatste + 1):
        if isinstance(get(kop[0] - 1, c), (int, float)):
            jaar = int(get(kop[0] - 1, c))
        jaren[c] = jaar
    for r in range(max(1, kop[0] - 3), kop[0] - 1):
        for c in range(eerste, laatste + 1):
            t = _t(get(r, c)).lower().replace(' ', '')
            if t in ('startbouw', 'oplevering') and t not in kop_tekst:
                kop_tekst[t] = L(c)
                q = _t(get(kop[0], c))
                if jaren[c] and re.match(r'Q[1-4]$', q):
                    mijlpaal[t] = (jaren[c], int(q[1]))
    act = next((L(c) for c in range(eerste, laatste + 1) if _t(get(kop[0], c)).lower().startswith('actual')), 'J')
    # standaard-FO: tabel 'Transport' (aantal woningen per kwartaal) onder het verkooptempo
    tt = next((r for r in range(vk + 1, omz[0]) if _t(get(r, 2)).lower() == 'transport'), None)
    for i, t in enumerate(types):
        t['transport_tabel'] = tt + 1 + i if tt else None
    return {'eerste': L(eerste), 'laatste': L(laatste), 'check': L(check), 'verschil': L(check + 1), 'status': L(check + 2),
            'actuals': act,
            'types': types, 'verkoop_kop': vk, 'euro_kop': euro, 'totaal': totaal, 'periode': periode,
            'correctie': corr, 'daeb': daeb, 'daeb_sjabloon': None if daeb else periode + 4, 'max': maxr,
            'daeb_totaal': None if daeb else periode + 28, 'daeb_eind': None if daeb else periode + 31,
            'kop_markering': kop_tekst, 'mijlpalen': mijlpaal, 'jaarrij': euro - 1}


def _kosten(f, v, opb):
    get = lambda r, c: v.cell(r, c).value
    form = lambda r, c: f.cell(r, c).value
    maxr = f.max_row
    ond = next(r for r in range(2, maxr) if _t(get(r, 2)).lower() == 'onderdelen')
    qk = _kopkolommen(v, ond)
    eerste_q, laatste = min(qk), max(qk)
    posten = []
    r = ond + 1
    while _t(form(r, 2)) and not _t(get(r, 2)).lower().startswith('totaal'):
        posten.append(r)
        r += 1
    totaal = next(r for r in range(posten[-1] + 1, maxr) if _t(get(r, 2)).lower().startswith('totaal kosten'))
    periode = next(r for r in range(totaal, maxr + 1) if _t(get(r, 2)).lower().startswith('periode in cashflow'))
    pew = max(r for r in range(2, ond) if _t(get(r, 2)).lower().startswith('totaal kosten'))
    check = next(r for r in range(posten[-1] + 1, totaal) if _t(get(r, 2)).lower().startswith('check'))
    som = next(r for r in range(posten[-1] + 1, check) if _t(form(r, 4)).upper().startswith('=SUM('))
    status = next(c for c in range(laatste + 1, laatste + 8) if '"OK"' in _t(form(posten[0], c)))
    toel = next((c for c in range(status, status + 6) if _t(get(ond, c)).lower().startswith('toelichting')), status + 2)
    werkelijk = next((c for c in range(3, eerste_q) if "'3. Restant bdgt'!" in _t(form(posten[0], c))
                      or "'FO - huidig'!$M" in _t(form(posten[0], c))), None)
    eerste = 7 if _t(get(ond - 1, 7)).lower().startswith('actual') or _t(get(ond, 7)).lower().startswith('actual') else 8
    # termijnblokken per woningtype (titel = de bloktitel op CF - opbrengsten)
    for t in opb['types']:
        kr = [r for r in range(2, ond) if _t(form(r, 2)) == "='CF - opbrengsten'!B%d" % t['hb']]
        rows = []
        if kr:
            r = kr[0] + 1
            while _t(form(r, 2)) and not _t(get(r, 2)).lower().startswith(('aanneemsom', 'meerwerk')):
                rows.append(r)
                r += 1
        t['kosten'] = rows
    start = maxr + 1
    return {'eerste': L(eerste), 'eerste_q': L(eerste_q), 'laatste': L(laatste), 'onderdelen': ond, 'posten': posten,
            'totaal': totaal, 'pct': totaal + 1, 'periode': periode, 'pew_totaal': pew, 'check': check, 'som': som,
            'status': L(status), 'totaalkol': L(status - 1), 'toelichting': L(toel),
            'werkelijk': L(werkelijk) if werkelijk else None, 'voorstel_start': start,
            'voorstel_offset': start + 7 - posten[0], 'voorstel_eind': posten[-1] + start + 8 - posten[0],
            'sleutelrij': start + 3, 'max': maxr}


def _pew(v):
    get = lambda r, c: v.cell(r, c).value
    kop = next(r for r in range(1, 12) if any(_t(get(r, c)).lower() == 'huidig' for c in range(3, 16)))
    kol = {_t(get(kop, c)).lower(): L(c) for c in range(3, 16) if _t(get(kop, c))}
    rij = {}
    for r in range(kop + 1, 45):
        t = _t(get(r, 3)).lower()
        if t in ('totaal kosten', 'totaal opbrengsten', 'opbrengsten', 'resultaat', 'resultaat %', 'check kosten',
                 'check opbrengsten') and t not in rij:
            rij[t] = r
    toel = next((L(c) for c in range(3, 20) if _t(get(kop - 1, c)).lower() == 'toelichting'), 'K')
    getallen = [c for c in kol.values()]
    return {'kop': kop, 'fd': kol.get('meest recente fd', 'E'), 'vorig': kol['vorig'], 'huidig': kol['huidig'],
            'kolommen': sorted(getallen, key=ci), 'toelichting': toel, 'eerste': kop + 2,
            'kosten_totaal': rij['totaal kosten'], 'opbrengsten': rij['opbrengsten'],
            'opbrengsten_totaal': rij['totaal opbrengsten'], 'resultaat': rij.get('resultaat'),
            'resultaat_pct': rij.get('resultaat %'), 'checks': [rij.get('check kosten'), rij.get('check opbrengsten')]}


def _cashflow(v):
    """periodes van 1. Cashflow: (jaar, 'Qn' of '') + de rijen, voor tab Invoer en de controle"""
    kop = next(r for r in range(1, 30) if _t(v.cell(r, 2).value).lower() == 'jaar')
    per = []
    r = kop + 1
    while True:
        b = v.cell(r, 2).value
        try:
            jaar = int(float(b))
        except (TypeError, ValueError):
            break
        q = _t(v.cell(r, 3).value)
        per.append((jaar, q if re.match(r'Q[1-4]$', q) else '', r))
        r += 1
    tot = next(rr for rr in range(r, r + 8) if _t(v.cell(rr, 3).value).lower().startswith('totaal'))
    ok = next((rr for rr in range(tot + 1, tot + 6) if _t(v.cell(rr, 4).value) in ('OK', 'FOUT')), tot + 3)
    return {'kop': kop, 'periodes': per, 'totaal': tot, 'ok': ok}


def _verkooptempo(v):
    kop = next((r for r in range(1, 30) if _t(v.cell(r, 1).value).lower() == 'verkooptempo'), None)
    return {'kop': kop}


def _hulp(wf):
    """hulpkolommen rechts van de exports: per blad de kolommen en de eerste rij met een formule"""
    out = {}
    for blad, kols in (('FO - vorig', 'XY'), ('FO - huidig', 'XY'), ('FO - actuals', 'QRS')):
        ws = wf[blad]
        out[blad] = {}
        for k in kols:
            rijen = [c.row for c in ws[k] if isinstance(c.value, str) and c.value.startswith('=')]
            if rijen:
                out[blad][k] = (min(rijen), max(rijen))
    return out


def _aansluiting(v):
    """blokken 'Vergelijking excl. BTW' op Aansluiting opbrengsten: per blok de regels met de kolom 'Corr'"""
    blokken = []
    for r in range(1, v.max_row + 1):
        t = _t(v.cell(r, 1).value)
        if not t.lower().startswith('vergelijking'):
            continue
        kop = r + 1
        corr = next((c for c in range(2, 15) if _t(v.cell(kop, c).value).lower() == 'corr'), None)
        if corr is None:
            continue
        naam = t.split(' - ', 1)[1].strip() if ' - ' in t else 'Blok %d' % (len(blokken) + 1)
        regels = []
        rr = kop + 1
        while _t(v.cell(rr, 1).value):
            regels.append((rr, _t(v.cell(rr, 1).value)))
            rr += 1
        blokken.append({'titel': r, 'naam': naam, 'corr': L(corr), 'regels': regels})
    return blokken


def _prijslijst(wf):
    """kolommen Type en Status op Prijslijst en per woningtype de sleutel (uit SUMIF(Prijslijst!C..,"sleutel",...) in de
    SKO-regels op Aansluiting opbrengsten), voor 'verkocht t/m nu' uit Salesforce"""
    if 'Prijslijst' not in wf.sheetnames:
        return None
    p = wf['Prijslijst']
    kop = {_t(p.cell(4, c).value).lower(): L(c) for c in range(1, 40) if _t(p.cell(4, c).value)}
    typ, stat = kop.get('type'), kop.get('sf: status')
    if not typ or not stat:
        return None
    a = wf['Aansluiting opbrengsten']
    sleutels = []
    for r in range(1, 40):
        f = _t(a.cell(r, 3).value)
        m = re.search(r'SUMIF\(Prijslijst!\$?[A-Z]{1,2}\$?\d+:\$?[A-Z]{1,2}\$?\d+,"([^"]+)"', f)   # kolom Type of Termijnen / blok
        if m:
            sleutels.append(m.group(1))
    return {'type': typ, 'status': stat, 'sleutels': sleutels}

