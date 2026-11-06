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

from xlsxpkg import Sheet, col2num, num2col

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
        ra, rb, rg = tr + 1, tr + 2, tr + 3            # ingevoegde rijen: verkocht %, getransporteerd %, grondprijs
        for r in (ra, rb, rg):
            for c in ['B'] + cols:
                assert not sh.get('%s%d' % (c, r)), 'rij %d is niet leeg' % r
        c1, cn = cols[0], cols[-1]
        keyo_ = '_koppeling!%s$521'
        z('B%d' % ra, 'Verkocht % · van alle woningen (uit het verkooptempo; 1e kolom: t/m nu, SF)', st['tekst'])
        z('B%d' % rb, 'Getransporteerd % · van alle woningen (uit de tabel Transport)', st['tekst'])
        z('B%d' % rg, 'Grondprijs +/- per woning (€, na SWK)', st['tekst'])
        z('D%d' % rg, None, st['invoer_getal'])
        z('B%d' % tr, 'Transport · aantal woningen (= grondtermijn)', st['tekst'])
        for j, c in enumerate(cols):
            prev = cols[j - 1] if j else None
            if j == 0:
                z('%s%d' % (c, ra), '=IF(N(%s)=0,0,N(%s%d)/%s)' % (won, c, t['verkoop'], won), st['pct_calc'])
            else:
                # verkocht vóór het kwartaal van de prijslijstexport (Prijslijst C1; leeg = het huidige kwartaal) zit al in
                # kolom 1 (t/m nu, Salesforce) en telt niet nog eens mee
                sf = ('IF(N(Prijslijst!$C$1)=0,%s,YEAR(Prijslijst!$C$1)*4+ROUNDUP(MONTH(Prijslijst!$C$1)/3,0))' % P_H)
                z('%s%d' % (c, ra), '=IF(OR(N(%s)=0,%s<%s),0,N(%s%d)/%s)' % (won, keyo_ % c, sf, c, t['verkoop'], won),
                  st['pct_calc'])
                v = t['verkoop_waarden'][j]
                z('%s%d' % (c, t['verkoop']), v if v else None, st['invoer_getal'])
            g = t['grond_waarden'][j]
            # transport in aantallen in de tabel Transport; het % rekent daaruit
            z('%s%d' % (c, t['transport_tabel']), g if g else None, st['invoer_getal'])
            z('%s%d' % (c, rb), '=IF(N(%s)=0,0,N(%s%d)/%s)' % (won, c, t['transport_tabel'], won), st['pct_calc'])
            eff = lambda k: 'MIN(SUM($%s%d:%s%d),SUM($%s%d:%s%d))' % (c1, rb, k, rb, c1, ra, k, ra)
            z('%s%d' % (c, tr), '=ROUND((%s%s)*%s,6)' % (eff(c), '-' + eff(prev) if prev else '', won), st['aantal'])
            z('%s%d' % (c, koop), '=%s%d' % (c, tr))
        # koopsom grond per woning + aanpassing
        e = sh.get('E%d' % koop)
        if isinstance(e, str) and e.startswith('=') and 'D%d' % rg not in e:
            z('E%d' % koop, e + '+N($D$%d)' % rg)
        z('%s%d' % (chk, tr), '=SUM(%s%d:%s%d)' % (c1, tr, cn, tr), sh.style_of('%s%d' % (chk, p1)))
        z('%s%d' % (verschil, tr), '=ROUND(%s-%s%d,2)' % (won, chk, tr), sh.style_of('%s%d' % (verschil, p1)))
        z('%s%d' % (stat, tr), '=IF(%s%d=0,"OK","FOUT")' % (verschil, tr), sh.style_of('%s%d' % (stat, p1)))
        for r in (ra, rb):
            z('%s%d' % (chk, r), '=SUM(%s%d:%s%d)' % (c1, r, cn, r), st['pct_calc'])
        gew = '(COLUMN($%s$1)-COLUMN($%s$1:$%s$1)+1)' % (cn, c1, cn)
        z('%s%d' % (stat, ra), '=IF(N(%s)=0,"",IF(%s%d>1.00001,"FOUT: meer verkocht dan woningen (t/m nu + prognose)",IF(%s%d<0.99999,'
                               '"LET OP: "&TEXT(1-%s%d,"0%%")&" nog niet ingepland","OK")))' % (won, chk, ra, chk, ra, chk, ra), st['vet'])
        z('%s%d' % (stat, rb), '=IF(N(%s)=0,"",IF(%s%d>1.00001,"FOUT: meer dan 100%%",IF(SUMPRODUCT(%s%d:%s%d*%s)-'
                               'SUMPRODUCT(%s%d:%s%d*%s)/%s>0.00001,"LET OP: transport wacht op verkoop (schuift mee)",'
                               'IF(%s%d<0.99999,"LET OP: "&TEXT(1-%s%d,"0%%")&" nog niet getransporteerd","OK"))))'
          % (won, chk, rb, c1, rb, cn, rb, gew, c1, tr, cn, tr, gew, won, chk, rb, chk, rb), st['vet'])
        t.update({'rij_verkocht': ra, 'rij_transport_pct': rb, 'rij_grond': rg})
        # verkocht t/m nu (kolom D van het verkooptempo) rechtstreeks uit de Salesforce-status in de prijslijst
        pl = oz.get('prijslijst')
        if pl and len(pl['sleutels']) >= t['k']:
            sl = pl['sleutels'][t['k'] - 1].replace('"', '""')
            tel = lambda s_: 'COUNTIFS(Prijslijst!$%s$5:$%s$2000,"%s",Prijslijst!$%s$5:$%s$2000,"%s")' % (
                pl['type'], pl['type'], sl, pl['status'], pl['status'], s_)
            z('D%d' % t['verkoop'], '=%s+%s' % (tel('Verkocht'), tel('Getransporteerd')), st['aantal'])
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
    # het standaard-FO heeft een eigen DAEB-blok (al in het totaal); hier komt alleen de correctieregel erbij
    for c in cols:
        f = sh.get('%s%d' % (c, tot))
        assert isinstance(f, str) and f.startswith('='), (c, tot, f)
        sh.set('%s%d' % (c, tot), f + '+%s%d' % (c, corr))
    z('C2', 'Geel = invullen. Per woningtype de bouwtermijnen (planning). Verkoop (rij %d-%d) en transport (tabel Transport) '
            'in aantallen per kwartaal; kolom 1 = verkocht t/m de exportdatum van de prijslijst (Salesforce); kwartalen vóór dat '
            'kwartaal tellen daarom niet nog eens mee. '
            'Verkocht %% en getransporteerd %% rekenen daaruit. Correctie: rij %d.' % (verkoop[0], verkoop[-1], corr), st['notitie'])


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
            # (opzet.py vond de kolom aan de verwijzing naar '3. Restant bdgt'; die is bij het overnemen al waarde geworden)
            z('%s%d' % (k['werkelijk'], p), "=SUMIFS('FO - huidig'!$M:$M,'FO - huidig'!$Y:$Y,$B%d)" % p)
    z('C%d' % k['onderdelen'], 'Methode', st['kop'])
    for p in posten:
        z('C%d' % p, 'Handmatig', st['invoer'])
        # FOUT pas bij meer dan €10 verschil (afrondingen uit het oude FO geven OK)
        z('%s%d' % (st_k, p), '=IF(OR($C%d="Voorstel",$C%d="Transport"),IF(ABS(%s%d-D%d)<10,"OK","FOUT"),'
                              'IF(ABS(%s%d-D%d)<10,"OK","FOUT"))' % (p, p, tk, p + off, p, tk, p, p))
    sh.add_list_validation('C%d:C%d' % (p1, pn), ['Handmatig', 'Voorstel', 'Transport'])
    if oz.get('standaard'):
        z('E%d' % k['onderdelen'], 'Verdeeld', st['kop'])
        z('F%d' % k['onderdelen'], 'Nog te verdelen', st['kop'])
        for p in posten:
            z('E%d' % p, '=IF(OR($C%d="Voorstel",$C%d="Transport"),%s%d,SUM(%s%d:%s%d))' % (p, p, tk, p + off, cols[0], p, cols[-1], p),
              st['getal'])
            z('F%d' % p, '=IF(B%d="","",IF(ABS(N(D%d)-E%d)<10,0,ROUND(N(D%d)-E%d,0)))' % (p, p, p, p, p), st['getal_tot'])
    m = '$C$%d:$C$%d' % (p1, pn)
    for c in cols:
        hand, vs = '%s%d:%s%d' % (c, p1, c, pn), '%s%d:%s%d' % (c, v1, c, vn)
        z('%s%d' % (c, k['totaal']), '=SUM(%s)-SUMIFS(%s,%s,"Voorstel")-SUMIFS(%s,%s,"Transport")'
                                     '+SUMIFS(%s,%s,"Voorstel")+SUMIFS(%s,%s,"Transport")' % (hand, hand, m, hand, m, vs, m, vs, m))
    # de uitleg van de methodes staat één keer, boven de kostenposten (standaard_fo, rij K_PEW + 1)
    z('%s4' % toel, 'Bouwtermijnen volgen de geplande bouwtermijnen van CF - opbrengsten; overschrijven mag.', st['notitie'])

    # voorstelblok
    s0 = k['voorstel_start']
    hj, hk, ht, hkop, hst = s0 + 2, s0 + 3, s0 + 4, s0 + 5, s0 + 6
    for c in ['B'] + kolommen('C', st_k):
        z('%s%d' % (c, s0), 'VOORSTEL · restant budget (PEW − verwerkt) vanaf het eerste prognosekwartaal: gelijk verdeeld t/m '
                            'oplevering (tab Start), of bij "Transport" naar rato van het transport · kies per post in kolom C' if c == 'B' else None, st['sectie'])
    # kolom voor het aantal kwartalen (of woningen) waarover verdeeld wordt: G, tenzij G een kwartaalkolom is (in het
    # standaard-FO is G de kolom Actuals; dan C)
    N_ = 'C' if 'G' in cols else 'G'
    # verdelen t/m: de oplevering (tab Start); leeg of al voorbij = het laatste kwartaal met opbrengsten
    ob = oz['opbrengsten']
    eind = '$D$%d' % (s0 + 1)
    keyo = '_koppeling!$%s$521:$%s$521' % (ob['eerste'], ob['laatste'])
    tot_o = "'CF - opbrengsten'!$%s$%d:$%s$%d" % (ob['eerste'], ob['totaal'], ob['laatste'], ob['totaal'])
    z('B%d' % (s0 + 1), 'verdelen t/m (oplevering op Start; leeg of voorbij = laatste kwartaal met opbrengsten)', st['hulp_links'])
    z('D%d' % (s0 + 1), '=IF(%s>=%s,%s,MAX(%s,IFERROR(LOOKUP(2,1/((%s<>0)*(%s>0)),%s),%s)))'
      % (P_O, P_H, P_O, P_H, tot_o, keyo, keyo, P_H), st['aantal'])
    z('E%d' % (s0 + 1), '=IF(N(D%d)=0,"","Q"&(D%d-4*INT((D%d-1)/4))&" "&INT((D%d-1)/4))' % (s0 + 1, s0 + 1, s0 + 1, s0 + 1), st['vet'])
    z('B%d' % hj, 'hulp: jaar', st['hulp_links'])
    z('B%d' % hk, 'hulp: kwartaalindex (jaar × 4 + kw)', st['hulp_links'])
    z('B%d' % ht, 'hulp: transport (aantal woningen, alle types)', st['hulp_links'])
    tr_rng = lambda t: "'CF - opbrengsten'!$%s$%d:$%s$%d" % (ob['eerste'], t['tr'], ob['laatste'], t['tr'])
    o = k['onderdelen']
    for i, c in enumerate(cols):
        prev = '%s%d' % (cols[i - 1], hj) if i else '""'
        z('%s%d' % (c, hj), '=IF(ISNUMBER(%s$%d),%s$%d,%s)' % (c, o - 1, c, o - 1, prev), st['hulp'])
        z('%s%d' % (c, hk), '=IF(AND(ISNUMBER(%s%d),LEFT(%s$%d&"",1)="Q"),%s%d*4+VALUE(MID(%s$%d,2,1)),0)'
          % (c, hj, c, o, c, hj, c, o), st['hulp'])
        z('%s%d' % (c, hkop), '=IF(%s$%d>0,"Q"&(%s$%d-4*%s$%d)&" \'"&RIGHT(%s$%d,2),"actuals")' % (c, hk, c, hk, c, hj, c, hj),
          st['kop'])
        z('%s%d' % (c, hst), '=IF(%s$%d<%s,"verwerkt",IF(%s$%d<=%s,"prognose","na oplevering"))'
          % (c, hk, P_H, c, hk, eind), st['hulp'])
        z('%s%d' % (c, ht), '=IF(%s$%d=0,0,%s)' % (c, hk, '+'.join('SUMIFS(%s,%s,%s$%d)' % (tr_rng(t), keyo, c, hk)
                                                                     for t in ob['types'])), st['hulp'])
    for c, h in zip(['B', 'C', 'D', 'E', 'F', N_, tk, st_k],
                    ['Post', '', 'PEW (huidig)', 'Verwerkt', 'Restant', 'Kw. / won.', 'Totaal', 'Check']):
        z('%s%d' % (c, hkop), h or None, st['kop_links'] if c == 'B' else st['kop'])
    rk = '$%s$%d:$%s$%d' % (cols[0], hk, cols[-1], hk)
    for p in posten:
        r = p + off
        z('B%d' % r, '=IF(B%d="","",B%d)' % (p, p), st['tekst'])
        z('D%d' % r, '=N(D%d)' % p, st['getal'])
        z('E%d' % r, '=SUMIFS($%s%d:$%s%d,%s,"<"&%s)' % (cols[0], p, cols[-1], p, rk, P_H), st['getal'])
        rt = '$%s$%d:$%s$%d' % (cols[0], ht, cols[-1], ht)
        rp = '$%s$%d:$%s$%d' % (cols[0], p, cols[-1], p)
        # restant na de zelf ingevulde (overschreven) prognosekwartalen
        z('F%d' % r, '=MAX(0,D%d-E%d-SUMIFS(%s,%s,">="&%s))' % (r, r, rp, rk, P_H), st['getal'])
        # een kwartaal telt als leeg (wordt verdeeld) als er niets of 0 staat
        z('%s%d' % (N_, r), '=IF($C%d="Transport",SUMPRODUCT(%s*(%s>=%s)*(((%s&"")="")+((%s&"")="0"))),'
                            'COUNTIFS(%s,">="&%s,%s,"<="&%s,%s,"")+COUNTIFS(%s,">="&%s,%s,"<="&%s,%s,0))'
          % (p, rt, rk, P_H, rp, rp, rk, P_H, rk, eind, rp, rk, P_H, rk, eind, rp), st['aantal'])
        for c in cols:
            z('%s%d' % (c, r), '=IF(%s$%d<%s,N(%s%d),IF(N(%s%d)<>0,%s%d,IF($C%d="Transport",IF($%s%d=0,0,$F%d*%s$%d/$%s%d),'
                               'IF(%s$%d<=%s,IF($%s%d=0,0,$F%d/$%s%d),0))))'
              % (c, hk, P_H, c, p, c, p, c, p, p, N_, r, r, c, ht, N_, r, c, hk, eind, N_, r, r, N_, r), st['getal'])
        z('%s%d' % (tk, r), '=SUM(%s%d:%s%d)' % (cols[0], r, cols[-1], r), st['getal_tot'])
        z('%s%d' % (st_k, r), '=IF(B%d="","",IF(D%d-E%d<-1,"OVERSCHRIJDING",IF(ABS(%s%d-D%d)<10,"OK","FOUT")))'
          % (r, r, r, tk, r, r), st['vet'])
    r = vn + 1
    z('B%d' % r, 'Totaal voorstel', st['label_tot'])
    for c in ['C', 'D', 'E', 'F'] + [x for x in ['G'] if x not in cols] + cols + [tk]:
        z('%s%d' % (c, r), '=SUM(%s%d:%s%d)' % (c, v1, c, vn) if c not in ('C', N_) else None, st['getal_tot'])
    k.update({'voorstel_eerste': v1, 'voorstel_laatste': vn, 'voorstel_eind': r, 'sleutelrij': hk})
    # opmaak: gekozen methode zichtbaar maken
    grijs = styles.add_dxf('<font><strike val="1"/><color rgb="FF9CA3AF"/></font>')
    geel = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFDE68A"/></patternFill></fill>')
    groen = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFD1FAE5"/></patternFill></fill>')
    rood = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFECACA"/></patternFill></fill>')
    sh.add_cf('%s%d:%s%d' % (cols[0], p1, cols[-1], pn), 'AND(OR($C%d="Voorstel",$C%d="Transport"),N(%s%d)=0)'
              % (p1, p1, cols[0], p1), grijs)
    sh.add_cf('%s%d:%s%d' % (cols[0], p1, cols[-1], pn), 'AND(OR($C%d="Voorstel",$C%d="Transport"),N(%s%d)<>0)'
              % (p1, p1, cols[0], p1), geel)
    sh.add_cf('%s%d:%s%d' % (cols[0], v1, cols[-1], vn), 'OR($C%d="Voorstel",$C%d="Transport")' % (p1, p1), geel)
    sh.add_cf('%s%d:%s%d' % (st_k, v1, st_k, vn), '%s%d="OK"' % (st_k, v1), groen)
    sh.add_cf('%s%d:%s%d' % (st_k, v1, st_k, vn), 'OR(%s%d="FOUT",%s%d="OVERSCHRIJDING")' % (st_k, v1, st_k, v1), rood)
    log('  CF - kosten: methode per post (rij %d-%d), voorstelblok rij %d-%d, rij %d = definitief, werkelijk uit FO - huidig'
        % (p1, pn, s0, r, k['totaal']))


# ================================================================================================
#  Aansluiting opbrengsten: te boeken in SAP, per SAP-nummer en kostendrager (uit de kolom Corr)
# ================================================================================================

SAP_KOPPELING = {'grond + ontwikkelkosten': (9010, None), 'aanneemsom': (9000, None), 'kao opties': (9000, 5230),
                 'kopersmeerwerk': (9200, 5250), 'kadastrale kosten': (9030, None), 'aanvullende rente': (9110, None)}
SAP_REGEL = {'Grond + ontwikkelkosten': 'Grond + Ontwikkelkosten', 'Aanneemsom': 'Aanneemsom', 'KAO-opties': 'KAO opties',
             'Meer-/minderwerk kopers': 'Kopersmeerwerk', 'Kadastrale kosten': 'Kadastrale kosten',
             'Aanvullende rente': 'Aanvullende rente'}       # boekingsregel -> regel in de koppeling
SAP_BOEKING = [('Grond + ontwikkelkosten', 9010, None), ('Aanneemsom', 9000, None), ('KAO-opties', 9000, 5230),
               ('Meer-/minderwerk kopers', 9200, 5250), ('Kadastrale kosten', 9030, None), ('Aanvullende rente', 9110, None)]


def sap_opzet(oz):
    """rijen van het blok 'te boeken in SAP' onder de bestaande inhoud van Aansluiting opbrengsten"""
    blokken = oz.get('aansluiting') or []
    regels = []
    for b in blokken:
        for _, lab in b['regels']:
            if lab not in regels:
                regels.append(lab)
    s0 = oz['aansluiting_max'] + 3
    b1 = s0 + 2
    bn = b1 + len(SAP_BOEKING) - 1
    k0 = bn + 4
    m1, mn = k0 + 2, k0 + 1 + len(regels)
    d0 = mn + 2
    d1 = d0 + 2
    dn = d1 + sum(len(b['regels']) for b in blokken) - 1
    return {'s0': s0, 'b1': b1, 'bn': bn, 'open': bn + 1, 'totaal': bn + 2, 'k0': k0, 'm1': m1, 'mn': mn, 'd0': d0,
            'd1': d1, 'dn': dn, 'regels': regels, 'blokken': blokken}


def aansluiting(sh: Sheet, oz: dict, st: dict, sjabloon: list, log):
    L_ = sap_opzet(oz)
    z = Zetter(sh, oz.get('aansluiting_blad', 'Aansluiting opbrengsten'), sjabloon)
    d1, dn = L_['d1'], max(L_['d1'], L_['dn'])
    rng = lambda c: '$%s$%d:$%s$%d' % (c, d1, c, dn)

    def sectie(r, tekst):
        for c in 'ABCDEFGHI':
            z('%s%d' % (c, r), tekst if c == 'A' else None, st['sectie'])
    sectie(L_['s0'], 'TE BOEKEN IN SAP · per SAP-nummer en kostendrager (uit de kolom Corr hierboven)')
    for c, h in zip('ABCD', ['Omschrijving', 'SAP-nummer', 'Kostendrager', 'Bedrag (€)']):
        z('%s%d' % (c, L_['s0'] + 1), h, st['kop_links'] if c == 'A' else st['kop'])
    # SAP-nummer en kostendrager staan op één plek: de koppeling hieronder; de boekingsregels nemen ze daaruit over
    ml_ = lambda c: '$%s$%d:$%s$%d' % (c, L_['m1'], c, max(L_['m1'], L_['mn']))
    for i, (oms, nr, kd) in enumerate(SAP_BOEKING):
        r = L_['b1'] + i
        z('A%d' % r, oms, st['tekst'])
        regel = SAP_REGEL.get(oms)
        if regel and regel in L_['regels']:
            for c in 'BC':
                x = 'INDEX(%s,MATCH("%s",%s,0))' % (ml_(c), regel, ml_('A'))
                z('%s%d' % (c, r), '=IFERROR(IF(%s="","",%s),"")' % (x, x), st['aantal'])
        else:
            z('B%d' % r, nr, st['invoer'])
            z('C%d' % r, kd, st['invoer'])
        z('D%d' % r, '=SUMPRODUCT((%s&""=B%d&"")*(%s&""=C%d&"")*%s)' % (rng('B'), r, rng('C'), r, rng('D')), st['getal_tot'])
    r = L_['open']
    z('A%d' % r, 'Niet toegewezen (geen of een ander SAP-nummer)', st['tekst'])
    z('D%d' % r, '=ROUND(D%d-SUM(D%d:D%d),2)' % (r + 1, L_['b1'], L_['bn']), st['getal'])
    z('E%d' % r, '=IF(ABS(D%d)<1,"OK","LET OP: koppel de regel aan een SAP-nummer (hieronder)")' % r, st['vet'])
    z('A%d' % (r + 1), 'Totaal correcties', st['label_tot'])
    z('D%d' % (r + 1), '=SUM(%s)' % rng('D'), st['getal_tot'])
    sectie(L_['k0'], 'KOPPELING · regel → SAP-nummer en kostendrager · hier aanpassen (de boekingen hierboven volgen)')
    for c, h in zip('ABC', ['Regel', 'SAP-nummer', 'Kostendrager']):
        z('%s%d' % (c, L_['k0'] + 1), h, st['kop_links'] if c == 'A' else st['kop'])
    for i, lab in enumerate(L_['regels']):
        r = L_['m1'] + i
        nr, kd = SAP_KOPPELING.get(lab.strip().lower(), (None, None))
        z('A%d' % r, lab, st['tekst'])
        z('B%d' % r, nr, st['invoer'])
        z('C%d' % r, kd, st['invoer'])
    sectie(L_['d0'], 'DETAIL · per woningtype (kolom Corr van de vergelijkingsblokken)')
    for c, h in zip('ABCDE', ['Regel', 'SAP-nummer', 'Kostendrager', 'Bedrag (€)', 'Woningtype']):
        z('%s%d' % (c, L_['d0'] + 1), h, st['kop_links'] if c == 'A' else st['kop'])
    ml = '$A$%d:$A$%d' % (L_['m1'], max(L_['m1'], L_['mn']))
    r = d1
    for b in L_['blokken']:
        for rr, lab in b['regels']:
            z('A%d' % r, '=A%d&""' % rr, st['tekst'])
            for c, mc in (('B', 'B'), ('C', 'C')):
                x = 'INDEX($%s$%d:$%s$%d,MATCH(A%d,%s,0))' % (mc, L_['m1'], mc, max(L_['m1'], L_['mn']), r, ml)
                z('%s%d' % (c, r), '=IFERROR(IF(%s="","",%s),"")' % (x, x), st['aantal'])
            z('D%d' % r, '=N(%s%d)' % (b['corr'], rr), st['getal'])
            z('E%d' % r, b['naam'], st['tekst'])
            r += 1
    if oz.get('standaard'):
        import standaard_fo as S_
        km = lambda c: '$%s$%d:$%s$%d' % (c, L_['m1'], c, max(L_['m1'], L_['mn']))
        # per regel naast Corr: waar de correctie geboekt wordt (zoals de kolom a/b/c in het oude tabblad)
        for b in L_['blokken']:
            z('K%d' % (b['titel'] + 1), 'Boeken op SAP', st['kop'])
            for rr, lab in b['regels']:
                x = lambda c: 'INDEX(%s,MATCH($A%d,%s,0))' % (km(c), rr, km('A'))
                z('K%d' % rr, '=IF(ABS(N($G%d))<0.5,"",IFERROR(%s&IF(%s&""="",""," · "&%s),"koppel aan SAP-nr"))'
                  % (rr, x('B'), x('C'), x('C')), st['vet'])
        # op het dashboard: te boeken in SAP per SAP-nummer (totaal, vet) met de kostendragers eronder
        r0 = S_.P_TYPE1 + S_.NT + 2
        for c in 'OPQRST':
            z('%s%d' % (c, r0), 'TE BOEKEN IN SAP · per SAP-nummer, met de kostendragers eronder' if c == 'O' else None, st['sectie'])
        for c, h in zip('OPQRST', ['SAP-nr', 'Omschrijving', None, None, 'Bedrag (€)', 'Status']):
            z('%s%d' % (c, r0 + 1), h, st['kop_links'] if c in 'OP' else st['kop'])
        groepen = []
        for i, (oms, nr, kd) in enumerate(SAP_BOEKING):
            g = next((g for g in groepen if g[0] == nr), None)
            if g:
                g[1].append(i)
            else:
                groepen.append((nr, [i]))
        bb = lambda c: '$%s$%d:$%s$%d' % (c, L_['b1'], c, L_['bn'])
        r = r0 + 2
        for nr, leden in groepen:
            b0 = L_['b1'] + leden[0]
            z('O%d' % r, '=B%d' % b0, st['label_tot'])
            z('P%d' % r, ' + '.join(SAP_BOEKING[i][0] for i in leden) + (' · totaal' if len(leden) > 1 else ''), st['label_tot'])
            for c in 'QR':
                z('%s%d' % (c, r), None, st['label_tot'])
            z('S%d' % r, '=SUMIF(%s,O%d,%s)' % (bb('B'), r, bb('D')), st['getal_tot'])
            z('T%d' % r, None, st['label_tot'])
            r += 1
            if len(leden) > 1 or any(SAP_BOEKING[i][2] for i in leden):
                for n, i in enumerate(leden):
                    rb = L_['b1'] + i
                    teken = '└ ' if n == len(leden) - 1 else '├ '
                    z('P%d' % r, '="   %s"&IF(C%d="","zonder kostendrager","kostendrager "&C%d)&" · %s"' % (teken, rb, rb, SAP_BOEKING[i][0]),
                      st['tekst'])
                    z('S%d' % r, '=D%d' % rb, st['getal'])
                    r += 1
        z('P%d' % r, 'Niet toegewezen (geen of een ander SAP-nummer)', st['tekst'])
        z('S%d' % r, '=D%d' % L_['open'], st['getal'])
        z('T%d' % r, '=E%d' % L_['open'], st['vet'])
        r += 1
        for c in 'OPQRST':
            z('%s%d' % (c, r), 'TOTAAL CORRECTIES' if c == 'P' else None, st['label_tot'])
        z('S%d' % r, '=D%d' % L_['totaal'], st['getal_tot'])
        corr = '+'.join('SUM(G%d:G%d)' % (bk['regels'][0][0], bk['regels'][-1][0]) for bk in L_['blokken'] if bk['regels'])
        z('T%d' % r, '=IF(ABS(S%d-(%s))<1,"OK = kolom Corr","FOUT: sluit niet aan op Corr")' % (r, corr or '0'), st['vet'])
        z('O%d' % (r + 2), 'Per regel: kolom K in de blokken hieronder. SAP-nummers en kostendragers aanpassen: rij %d.' % L_['k0'],
          st['notitie'])
    log('  Aansluiting opbrengsten: te boeken in SAP vanaf rij %d (%d regels per type, %d blokken)'
        % (L_['s0'], len(L_['regels']), len(L_['blokken'])))

