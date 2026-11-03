"""Aanpassingen in de FO-tabbladen CF - opbrengsten en CF - kosten (op XML-niveau, in het samengevoegde bestand).
Alle rijen en kolommen komen uit de opzet van het FO (opzet.py), dus ook een FO met 3 woningtypes of kwartalen t/m 2030.

CF - opbrengsten
  * per woningtype een transportregel in de lege rij onder de bouwplanning (standaard de formules van de oude
    grondtermijn, dus zonder aanpassing dezelfde uitkomst); de grondtermijn volgt het transport;
  * de bouwtermijnen van het omzetblok tellen vanaf transport (begrensd door het transport i.p.v. door de verkoop);
  * een correctieregel per kwartaal in een lege rij boven het totaal, opgeteld in het totaal;
  * zonder eigen DAEB-blok: een leeg blok 'Opbrengsten DAEB' onder de periode-rij (in het sjabloon als
    'CF - opbrengsten|DAEB': Oud FO inlezen zet het alleen terug als het oude FO geen eigen DAEB-blok heeft).
CF - kosten
  * werkelijke kosten per post rechtstreeks uit FO - huidig (tab 3. Restant bdgt vervalt);
  * kolom C bij de posten: methode Handmatig of Voorstel; de totaalregel = definitief volgens die keuze;
  * voorstelblok onder de laatste gebruikte rij: restant budget gelijk verdeeld t/m oplevering.
Elke gezette cel komt in de sjabloonlijst, zodat Oud FO inlezen hem na het overnemen terugzet. De totaalregel van de
opbrengsten staat er niet in: een oud FO houdt zijn eigen totaalformule, de macro zet de correctieregel erachter.
"""
from __future__ import annotations

import re

from opzet import VT_BLAD, VT_KEY, vt_rij, vt_totaal
from xlsxpkg import Sheet, col2num, num2col

VT = "'%s'!" % VT_BLAD

P_H, P_O = 'p_huidig', 'p_oplevering'
MIJLPALEN = ['na akkoord SO', 'na akkoord VO', 'na akkoord DO', 'na omgevingsvergunning', 'na tekenen TKO',
             'bij start bouw', 'bij gevelsluiting', 'bij oplevering']


def kolommen(a, b):
    return [num2col(c) for c in range(col2num(a), col2num(b) + 1)]


class Zetter:
    """zet cellen en onthoudt ze voor het sjabloon"""

    def __init__(self, sh: Sheet, blad: str, sjabloon: list):
        self.sh, self.blad, self.sjabloon = sh, blad, sjabloon

    def __call__(self, ref, inhoud, style=None):
        self.sh.set(ref, inhoud, style)
        self.sjabloon.append((self.blad, ref, '' if inhoud is None else str(inhoud)))


def ref_rij(text, oud, nieuw):
    """verwijzingen naar rij `oud` op hetzelfde blad (A1-stijl, met of zonder $) vervangen door rij `nieuw`"""
    return re.sub(r"(\"[^\"]*\"|'[^']*'![$A-Z0-9:]+)|(?<![A-Za-z0-9_.!$])(\$?[A-Z]{1,3}\$?)%d(?![\d(])" % oud,
                  lambda m: m.group(0) if m.group(1) else '%s%d' % (m.group(2), nieuw), text)


def opbrengsten(sh: Sheet, oz: dict, st: dict, sjabloon: list, log):
    ob = oz['opbrengsten']
    cols = kolommen(ob['eerste'], ob['laatste'])
    chk, verschil, stat = ob['check'], ob['verschil'], ob['status']
    z = Zetter(sh, 'CF - opbrengsten', sjabloon)
    verkoop = [t['verkoop'] for t in ob['types']]
    for t in ob['types']:
        tr, koop, p1 = t['tr'], t['koop'], t['plan'][0]
        for c in ['B'] + cols:
            assert not sh.get('%s%d' % (c, tr)), 'transportrij %d is niet leeg' % tr
        won = '$G$%d' % (t['hb'] - 1)
        rv = vt_rij(t['k'])
        z('B%d' % tr, 'Transport (aantal woningen, uit Verkoop & transport) · grondtermijn', st['tekst'])
        for j, c in enumerate(cols):
            # transport en verkocht komen uit het tabblad Verkoop & transport (in % ingevuld); de eerste kolom van de
            # verkooprij blijft de realisatie uit Salesforce
            z('%s%d' % (c, tr), '=%s%s%d' % (VT, c, rv + 6), st['aantal'])
            z('%s%d' % (c, koop), '=%s%d' % (c, tr))
            if j:
                z('%s%d' % (c, t['verkoop']), '=%s%s%d' % (VT, c, rv + 3), st['aantal'])
        z('%s%d' % (chk, tr), '=SUM(%s%d:%s%d)' % (cols[0], tr, cols[-1], tr), sh.style_of('%s%d' % (chk, p1)))
        z('%s%d' % (verschil, tr), '=ROUND(%s-%s%d,2)' % (won, chk, tr), sh.style_of('%s%d' % (verschil, p1)))
        z('%s%d' % (stat, tr), '=IF(%s%d=0,"OK","FOUT")' % (verschil, tr), sh.style_of('%s%d' % (stat, p1)))
        # bouwtermijnen van dit omzetblok: begrensd door het transport van dit type (ook als een termijnrij per
        # vergissing naar de verkooprij van een ander type verwees)
        n = 0
        for r in t['omzet'][1:]:
            for c in cols + ['C', 'D', 'E', 'F']:
                f = sh.get('%s%d' % (c, r))
                if isinstance(f, str) and f.startswith('='):
                    nf = f
                    for v in verkoop:
                        nf = ref_rij(nf, v, tr)
                    if nf != f:
                        z('%s%d' % (c, r), nf)
                        n += 1
        log('  type %d: transport rij %d, grondtermijn rij %d, %d termijncellen vanaf transport' % (t['k'], tr, koop, n))

    corr, tot = ob['correctie'], ob['totaal']
    z('B%d' % corr, 'Handmatige correctie (€, per kwartaal)', st['invoer_label'])
    for c in cols:
        z('%s%d' % (c, corr), None, st['invoer_getal'])
    z('%s%d' % (chk, corr), '=SUM(%s%d:%s%d)' % (cols[0], corr, cols[-1], corr), st['getal'])
    extra = None
    if ob['daeb_sjabloon']:
        daeb_blok(sh, oz, st, sjabloon, log)
        extra = ob['daeb_totaal']
    for c in cols:
        f = sh.get('%s%d' % (c, tot))
        assert isinstance(f, str) and f.startswith('='), (c, tot, f)
        sh.set('%s%d' % (c, tot), f + '+%s%d' % (c, corr) + ('+%s%d' % (c, extra) if extra else ''))
    z('C2', 'Verkocht en transport vul je in %% in op tab Verkoop & transport (verkooptempo rij %d-%d en de transportregels '
            'komen daaruit). Hier: bouwtermijnen per woningtype en de correctie (rij %d). Termijnen tellen vanaf transport.'
      % (verkoop[0], verkoop[-1], corr), st['notitie'])


def daeb_blok(sh: Sheet, oz: dict, st: dict, sjabloon: list, log):
    """leeg DAEB-blok in de opzet die de macro (LeesDaebBlok) ook kende: componenten (naam in C, totaal in E) in de rijen
    boven de koprij, 'Startbouw' mag boven een kwartaal; termijnen per component als groep, gescheiden door een lege rij,
    tot 'Totaal ...'; daaronder aantal en verkoop van de DAEB-woningen"""
    ob = oz['opbrengsten']
    cols = kolommen(ob['eerste'], ob['laatste'])
    chk = ob['check']
    b0 = ob['daeb_sjabloon']
    r0 = b0 + 5
    tot = r0 + 19
    ob.update({'daeb_kop': r0, 'daeb_totaal': tot, 'daeb_eind': tot + 3})
    for r in range(b0, tot + 5):
        for c in ['B', 'C', 'D', 'E', 'F'] + cols + [chk]:
            assert not sh.get('%s%d' % (c, r)), 'rij %d van CF - opbrengsten is niet leeg' % r
    z = Zetter(sh, 'CF - opbrengsten|DAEB', sjabloon)   # alleen terugzetten als een oud FO geen DAEB-blok heeft
    z('C%d' % b0, 'DAEB-woningen (corporatie): vul de fee-componenten (naam + totaal), de bedragen per kwartaal, het aantal '
                  'en de verkoop in. Leeg = geen DAEB. Zet eventueel "Startbouw" boven het kwartaal van start bouw (rij %d).'
      % (r0 - 3), st['notitie'])
    z('B%d' % (r0 - 3), 'Fee-component 1 · naam (C) en totaal € (E)', st['tekst'])
    z('B%d' % (r0 - 2), 'Fee-component 2 · naam (C) en totaal € (E)', st['tekst'])
    for r in (r0 - 3, r0 - 2):
        z('C%d' % r, None, st['invoer_label'])
        z('E%d' % r, None, st['invoer_getal'])
    jr, kr = ob['jaarrij'], ob['euro_kop']
    for c in cols:
        z('%s%d' % (c, r0 - 1), '=IF(ISNUMBER(%s$%d),%s$%d,"")' % (c, jr, c, jr), st['hulp'])
        z('%s%d' % (c, r0), '=%s$%d&""' % (c, kr), st['kop'])
    z('B%d' % r0, 'Opbrengsten DAEB', st['kop_links'])
    z('C%d' % r0, '%', st['kop'])
    z('D%d' % r0, '€', st['kop'])
    z('J%d' % r0, 'actuals', st['kop'])
    for r1 in (r0 + 1, r0 + 10):
        for i, m in enumerate(MIJLPALEN):
            r = r1 + i
            z('B%d' % r, m, st['invoer_label'])
            for c in ['C', 'D'] + cols:
                z('%s%d' % (c, r), None, st['invoer_getal'])
            z('%s%d' % (chk, r), '=SUM(%s%d:%s%d)' % (cols[0], r, cols[-1], r), st['getal'])
    z('B%d' % tot, 'Totaal opbrengsten DAEB', st['label_tot'])
    for c in cols + [chk]:
        z('%s%d' % (c, tot), '=SUM(%s%d:%s%d)' % (c, r0 + 1, c, tot - 2), st['getal_tot'])
    z('B%d' % (tot + 2), 'Aantal DAEB-woningen', st['tekst'])
    z('D%d' % (tot + 2), None, st['invoer_getal'])
    z('B%d' % (tot + 3), 'Verkoop en levering DAEB (aantal woningen per kwartaal; transport = verkoop)', st['tekst'])
    for c in cols:
        z('%s%d' % (c, tot + 3), None, st['invoer_getal'])
    z('%s%d' % (chk, tot + 3), '=SUM(%s%d:%s%d)' % (cols[0], tot + 3, cols[-1], tot + 3), st['getal'])
    log('  CF - opbrengsten: leeg DAEB-blok rij %d-%d (koprij %d, totaal rij %d in het totaal)' % (b0, tot + 3, r0, tot))


def kosten(sh: Sheet, oz: dict, st: dict, sjabloon: list, styles, log):
    k = oz['kosten']
    cols = kolommen(k['eerste'], k['laatste'])
    posten, off = k['posten'], k['voorstel_offset']
    p1, pn = posten[0], posten[-1]
    v1, vn = p1 + off, pn + off
    tk, st_k, toel = k['totaalkol'], k['status'], k['toelichting']
    z = Zetter(sh, 'CF - kosten', sjabloon)
    # werkelijk per post rechtstreeks uit FO - huidig (was '3. Restant bdgt'!F*1000)
    if k['werkelijk']:
        for p in posten:
            f = sh.get('%s%d' % (k['werkelijk'], p)) or ''
            if "'3. Restant bdgt'!" in f:
                z('%s%d' % (k['werkelijk'], p), "=SUMIFS('FO - huidig'!$M:$M,'FO - huidig'!$Y:$Y,$B%d)" % p)
    z('C%d' % k['onderdelen'], 'Methode', st['kop'])
    for p in posten:
        z('C%d' % p, 'Handmatig', st['invoer'])
        z('%s%d' % (st_k, p), '=IF(OR($C%d="Voorstel",$C%d="Transport"),IF(ABS(%s%d-D%d)<1,"OK","FOUT"),'
                              'IF(AND(%s%d-D%d<1,%s%d-D%d>-1),"OK","FOUT"))' % (p, p, tk, p + off, p, tk, p, p, tk, p, p))
    sh.add_list_validation('C%d:C%d' % (p1, pn), ['Handmatig', 'Voorstel', 'Transport'])
    m = '$C$%d:$C$%d' % (p1, pn)
    for c in cols:
        hand, vs = '%s%d:%s%d' % (c, p1, c, pn), '%s%d:%s%d' % (c, v1, c, vn)
        z('%s%d' % (c, k['totaal']), '=SUM(%s)-SUMIFS(%s,%s,"Voorstel")-SUMIFS(%s,%s,"Transport")'
                                     '+SUMIFS(%s,%s,"Voorstel")+SUMIFS(%s,%s,"Transport")' % (hand, hand, m, hand, m, vs, m, vs, m))
    z('%s%d' % (toel, k['onderdelen'] - 1), 'Kolom C: Handmatig = deze regels · Voorstel = restant gelijk verdeeld t/m '
                                            'oplevering · Transport = restant naar rato van het transport (bv. grond). '
                                            'Blok onderaan rij %d-%d; rij %d = definitief.' % (v1, vn, k['totaal']), st['notitie'])
    z('%s4' % toel, 'Bouwtermijnen volgen de geplande bouwtermijnen van CF - opbrengsten; overschrijven mag.', st['notitie'])

    # voorstelblok
    s0 = k['voorstel_start']
    hj, hk, ht, hkop, hst = s0 + 2, s0 + 3, s0 + 4, s0 + 5, s0 + 6
    for c in ['B'] + kolommen('C', st_k):
        z('%s%d' % (c, s0), 'VOORSTEL · restant budget (PEW − verwerkt) vanaf het eerste prognosekwartaal: gelijk verdeeld t/m '
                            'oplevering (tab Start), of bij "Transport" naar rato van het transport · kies per post in kolom C' if c == 'B' else None, st['sectie'])
    z('B%d' % hj, 'hulp: jaar', st['hulp_links'])
    z('B%d' % hk, 'hulp: kwartaalindex (jaar × 4 + kw)', st['hulp_links'])
    z('B%d' % ht, 'hulp: transport (aantal woningen, alle types)', st['hulp_links'])
    T = vt_totaal(len(oz['opbrengsten']['types']))
    ob = oz['opbrengsten']
    vt_rng = lambda r: "%s$%s$%d:$%s$%d" % (VT, ob['eerste'], r, ob['laatste'], r)
    o = k['onderdelen']
    for i, c in enumerate(cols):
        prev = '%s%d' % (cols[i - 1], hj) if i else '""'
        z('%s%d' % (c, hj), '=IF(ISNUMBER(%s$%d),%s$%d,%s)' % (c, o - 1, c, o - 1, prev), st['hulp'])
        z('%s%d' % (c, hk), '=IF(AND(ISNUMBER(%s%d),LEFT(%s$%d&"",1)="Q"),%s%d*4+VALUE(MID(%s$%d,2,1)),0)'
          % (c, hj, c, o, c, hj, c, o), st['hulp'])
        z('%s%d' % (c, hkop), '=IF(%s$%d>0,"Q"&(%s$%d-4*%s$%d)&" \'"&RIGHT(%s$%d,2),"actuals")' % (c, hk, c, hk, c, hj, c, hj),
          st['kop'])
        z('%s%d' % (c, hst), '=IF(%s$%d<%s,"verwerkt",IF(%s$%d<=MAX(%s,%s),"prognose","na oplevering"))'
          % (c, hk, P_H, c, hk, P_H, P_O), st['hulp'])
        z('%s%d' % (c, ht), '=IF(%s$%d=0,0,SUMIFS(%s,%s,%s$%d))' % (c, hk, vt_rng(T + 3), vt_rng(VT_KEY), c, hk), st['hulp'])
    for c, h in zip(['B', 'C', 'D', 'E', 'F', 'G', tk, st_k],
                    ['Post', '', 'PEW (huidig)', 'Verwerkt', 'Restant', 'Kw. / won.', 'Totaal', 'Check']):
        z('%s%d' % (c, hkop), h or None, st['kop_links'] if c == 'B' else st['kop'])
    rk = '$%s$%d:$%s$%d' % (cols[0], hk, cols[-1], hk)
    for p in posten:
        r = p + off
        z('B%d' % r, '=IF(B%d="","",B%d)' % (p, p), st['tekst'])
        z('D%d' % r, '=N(D%d)' % p, st['getal'])
        z('E%d' % r, '=SUMIFS($%s%d:$%s%d,%s,"<"&%s)' % (cols[0], p, cols[-1], p, rk, P_H), st['getal'])
        z('F%d' % r, '=MAX(0,D%d-E%d)' % (r, r), st['getal'])
        rt = '$%s$%d:$%s$%d' % (cols[0], ht, cols[-1], ht)
        z('G%d' % r, '=IF($C%d="Transport",SUMIFS(%s,%s,">="&%s),COUNTIFS(%s,">="&%s,%s,"<="&MAX(%s,%s)))'
          % (p, rt, rk, P_H, rk, P_H, rk, P_H, P_O), st['aantal'])
        for c in cols:
            z('%s%d' % (c, r), '=IF(%s$%d<%s,N(%s%d),IF($C%d="Transport",IF($G%d=0,0,$F%d*%s$%d/$G%d),'
                               'IF(%s$%d<=MAX(%s,%s),IF($G%d=0,0,$F%d/$G%d),0)))'
              % (c, hk, P_H, c, p, p, r, r, c, ht, r, c, hk, P_H, P_O, r, r, r), st['getal'])
        z('%s%d' % (tk, r), '=SUM(%s%d:%s%d)' % (cols[0], r, cols[-1], r), st['getal_tot'])
        z('%s%d' % (st_k, r), '=IF(B%d="","",IF(D%d-E%d<-1,"OVERSCHRIJDING",IF(ABS(%s%d-D%d)<1,"OK","FOUT")))'
          % (r, r, r, tk, r, r), st['vet'])
    r = vn + 1
    z('B%d' % r, 'Totaal voorstel', st['label_tot'])
    for c in ['C', 'D', 'E', 'F', 'G'] + cols + [tk]:
        z('%s%d' % (c, r), '=SUM(%s%d:%s%d)' % (c, v1, c, vn) if c not in ('C', 'G') else None, st['getal_tot'])
    k.update({'voorstel_eerste': v1, 'voorstel_laatste': vn, 'voorstel_eind': r, 'sleutelrij': hk})
    # opmaak: gekozen methode zichtbaar maken
    grijs = styles.add_dxf('<font><strike val="1"/><color rgb="FF9CA3AF"/></font>')
    geel = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFDE68A"/></patternFill></fill>')
    groen = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFD1FAE5"/></patternFill></fill>')
    rood = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFECACA"/></patternFill></fill>')
    sh.add_cf('%s%d:%s%d' % (cols[0], p1, cols[-1], pn), 'OR($C%d="Voorstel",$C%d="Transport")' % (p1, p1), grijs)
    sh.add_cf('%s%d:%s%d' % (cols[0], v1, cols[-1], vn), 'OR($C%d="Voorstel",$C%d="Transport")' % (p1, p1), geel)
    sh.add_cf('%s%d:%s%d' % (st_k, v1, st_k, vn), '%s%d="OK"' % (st_k, v1), groen)
    sh.add_cf('%s%d:%s%d' % (st_k, v1, st_k, vn), 'OR(%s%d="FOUT",%s%d="OVERSCHRIJDING")' % (st_k, v1, st_k, v1), rood)
    log('  CF - kosten: methode per post (rij %d-%d), voorstelblok rij %d-%d, rij %d = definitief, werkelijk uit FO - huidig'
        % (p1, pn, s0, r, k['totaal']))
