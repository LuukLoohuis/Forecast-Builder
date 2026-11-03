"""Aanpassingen in de FO-tabbladen CF - opbrengsten en CF - kosten (op XML-niveau, in het samengevoegde bestand).

CF - opbrengsten
  * per woningtype een transportregel in de lege rij onder de bouwtermijnen (standaard dezelfde formule als de oude
    grondtermijn, dus zonder aanpassing dezelfde uitkomst); de grondtermijn volgt het transport;
  * de bouwtermijnen tellen vanaf transport (cumulatief begrensd door het transport i.p.v. door de verkoop);
  * rij 192: handmatige correctie per kwartaal, opgeteld in het totaal (rij 194);
  * rij 200-227: leeg blok 'Opbrengsten DAEB' (fee-componenten, termijnen per kwartaal, aantal en verkoop); het totaal
    (rij 224) telt mee in rij 194. In het sjabloon als 'CF - opbrengsten|DAEB': Oud FO inlezen zet het alleen terug als
    het oude FO geen eigen DAEB-blok heeft; de formules op _koppeling zoeken het blok op aan de titel.
CF - kosten
  * werkelijke kosten (kolom K, posten) direct uit FO - huidig (tab 3. Restant bdgt vervalt);
  * kolom C bij de posten: methode Handmatig of Voorstel; rij 75 (totaal) = definitief volgens die keuze;
  * voorstelblok rij 93-112 (post p -> rij p + 41): restant budget gelijk verdeeld t/m oplevering;
  * hulprijen 95-96: jaar en kwartaalindex per kolom.
Rij 194 (totaal opbrengsten) staat niet in het sjabloon: een oud FO houdt zijn totaalformule, de macro zet er '+G192' achter.
Elke gezette cel komt in de sjabloonlijst, zodat Oud FO inlezen hem na het overnemen terugzet.
"""
from __future__ import annotations

import re

from xlsxpkg import M, Sheet, num2col, split_ref

O_COLS = [num2col(c) for c in range(7, 27)]       # CF - opbrengsten G..Z
K_COLS = [num2col(c) for c in range(8, 28)]       # CF - kosten H..AA
K_POSTEN = list(range(59, 71))
VOORSTEL = 41                                     # rij voorstel = rij post + 41
VERKOOP_RIJEN = [49, 50, 51, 52]                  # verkooptempo per woningtype (CF - opbrengsten)
P_H, P_O = 'p_huidig', 'p_oplevering'
DAEB_KOP, DAEB_TOTAAL = 205, 224                 # koprij en totaalrij van het DAEB-blok op CF - opbrengsten
MIJLPALEN = ['na akkoord SO', 'na akkoord VO', 'na akkoord DO', 'na omgevingsvergunning', 'na tekenen TKO',
             'bij start bouw', 'bij gevelsluiting', 'bij oplevering']


class Zetter:
    """zet cellen en onthoudt ze voor het sjabloon"""

    def __init__(self, sh: Sheet, blad: str, sjabloon: list):
        self.sh, self.blad, self.sjabloon = sh, blad, sjabloon

    def __call__(self, ref, inhoud, style=None):
        self.sh.set(ref, inhoud, style)
        self.sjabloon.append((self.blad, ref, '' if inhoud is None else str(inhoud)))


def ref_rij(text, oud, nieuw):
    """verwijzingen naar rij `oud` (A1-stijl, met of zonder $) vervangen door rij `nieuw`"""
    return re.sub(r"(\"[^\"]*\"|'[^']*')|(?<![A-Za-z0-9_.!$])(\$?[A-Z]{1,3}\$?)%d(?![\d(])" % oud,
                  lambda m: m.group(0) if m.group(1) else '%s%d' % (m.group(2), nieuw), text)


def opbrengsten(sh: Sheet, st: dict, sjabloon: list, log):
    z = Zetter(sh, 'CF - opbrengsten', sjabloon)
    for v in VERKOOP_RIJEN:
        # omzetrijen die het verkooptempo van dit type gebruiken; de eerste is de grondtermijn (koopsom)
        pat = re.compile(r'(?<![A-Za-z0-9_.!$])\$?[A-Z]{1,3}\$?%d(?![\d(])' % v)
        rijen = set()
        for cel in sh.t.iter(M + 'c'):
            f = cel.find(M + 'f')
            if f is not None and f.text and pat.search(f.text):
                r = split_ref(cel.get('r'))[1]
                if r > 55:
                    rijen.add(r)
        rijen = sorted(rijen)
        if not rijen:
            log('  type met verkooprij %d heeft geen omzetregels: overgeslagen' % v)
            continue
        grond = rijen[0]
        g = sh.get('G%d' % grond) or ''
        plan = int(re.search(r'G(\d+)\s*<', g).group(1))
        tr = plan
        while sh.get('B%d' % tr):
            tr += 1
        for c in O_COLS:
            assert not sh.get('%s%d' % (c, tr)), 'transportrij %d is niet leeg' % tr
        # # won van het type (uit de check van de eerste planningsrij)
        won = re.search(r'(\$G\$\d+)', sh.get('AC%d' % plan) or '').group(1)
        z('B%d' % tr, 'Transport (notarieel) · grondtermijn', st['invoer_label'])
        for c in O_COLS:
            oud = sh.get('%s%d' % (c, grond))
            z('%s%d' % (c, tr), oud, sh.style_of('%s%d' % (c, plan)))
            z('%s%d' % (c, grond), '=%s%d' % (c, tr))
        z('AB%d' % tr, '=SUM(G%d:Z%d)' % (tr, tr), sh.style_of('AB%d' % plan))
        z('AC%d' % tr, '=ROUND(%s-AB%d,2)' % (won, tr), sh.style_of('AC%d' % plan))
        z('AD%d' % tr, '=IF(AC%d=0,"OK","FOUT")' % tr, sh.style_of('AD%d' % plan))
        n = 0
        for r in rijen[1:]:
            for c in O_COLS + ['C', 'D', 'E', 'F']:
                f = sh.get('%s%d' % (c, r))
                if isinstance(f, str) and f.startswith('='):
                    nf = ref_rij(f, v, tr)
                    if nf != f:
                        z('%s%d' % (c, r), nf)
                        n += 1
        log('  type (verkoop rij %d): transport rij %d, grondtermijn rij %d, %d termijncellen vanaf transport'
            % (v, tr, grond, n))

    # correctieregel; het totaal (rij 194) krijgt correctie en DAEB erbij. Rij 194 staat niet in het sjabloon: na Oud FO
    # inlezen houdt het oude FO zijn eigen totaalformule en zet de macro er alleen '+G192' achter
    z('B192', 'Handmatige correctie (€, per kwartaal)', st['invoer_label'])
    for c in O_COLS:
        z('%s192' % c, None, st['invoer_getal'])
    z('AB192', '=SUM(G192:Z192)', st['getal'])
    daeb_blok(sh, st, sjabloon, log)
    for c in O_COLS:
        f = sh.get('%s194' % c)
        assert f.startswith('=') and '192' not in f
        sh.set('%s194' % c, f + '+%s192+%s%d' % (c, c, DAEB_TOTAAL))
    z('C2', 'geel = handmatig: bouwtermijnen en transport per woningtype, verkooptempo (rij 49-52), correctie (rij 192). '
            'Termijnen tellen vanaf transport.', st['notitie'])


def daeb_blok(sh: Sheet, st: dict, sjabloon: list, log):
    """leeg DAEB-blok in de opzet die de macro (LeesDaebBlok) ook kende: componenten (naam in C, totaal in E) in de drie
    rijen boven de koprij, 'Startbouw' mag boven een kwartaal; termijnen per component als groep, gescheiden door een lege
    rij, tot 'Totaal ...'; daaronder aantal en verkoop van de DAEB-woningen"""
    r0 = DAEB_KOP
    for r in range(200, 229):
        for c in ['B', 'C', 'D', 'E', 'F'] + O_COLS + ['AB']:
            assert not sh.get('%s%d' % (c, r)), 'rij %d van CF - opbrengsten is niet leeg' % r
    z = Zetter(sh, 'CF - opbrengsten|DAEB', sjabloon)   # alleen terugzetten als een oud FO geen DAEB-blok heeft
    z('C200', 'DAEB-woningen (corporatie): vul de fee-componenten (naam + totaal), de bedragen per kwartaal, het aantal '
                   'en de verkoop in. Leeg = geen DAEB. Zet eventueel "Startbouw" boven het kwartaal van start bouw (rij 202).',
           st['notitie'])
    z('B%d' % (r0 - 3), 'Fee-component 1 · naam (C) en totaal € (E)', st['tekst'])
    z('B%d' % (r0 - 2), 'Fee-component 2 · naam (C) en totaal € (E)', st['tekst'])
    for r in (r0 - 3, r0 - 2):
        z('C%d' % r, None, st['invoer_label'])
        z('E%d' % r, None, st['invoer_getal'])
    for c in O_COLS:
        z('%s%d' % (c, r0 - 1), '=IF(ISNUMBER(%s$127),%s$127,"")' % (c, c), st['hulp'])
        z('%s%d' % (c, r0), '=%s$128&""' % c, st['kop'])
    z('B%d' % r0, 'Opbrengsten DAEB', st['kop_links'])
    z('C%d' % r0, '%', st['kop'])
    z('D%d' % r0, '€', st['kop'])
    z('J%d' % r0, 'actuals', st['kop'])
    for g, r1 in ((1, r0 + 1), (2, r0 + 10)):
        for i, m in enumerate(MIJLPALEN):
            r = r1 + i
            z('B%d' % r, m, st['invoer_label'])
            for c in ['C', 'D'] + O_COLS:
                z('%s%d' % (c, r), None, st['invoer_getal'])
            z('AB%d' % r, '=SUM(G%d:Z%d)' % (r, r), st['getal'])
    z('B%d' % DAEB_TOTAAL, 'Totaal opbrengsten DAEB', st['label_tot'])
    for c in O_COLS + ['AB']:
        z('%s%d' % (c, DAEB_TOTAAL), '=SUM(%s%d:%s%d)' % (c, r0 + 1, c, DAEB_TOTAAL - 2), st['getal_tot'])
    z('B226', 'Aantal DAEB-woningen', st['tekst'])
    z('D226', None, st['invoer_getal'])
    z('B227', 'Verkoop en levering DAEB (aantal woningen per kwartaal; transport = verkoop)', st['tekst'])
    for c in O_COLS:
        z('%s227' % c, None, st['invoer_getal'])
    z('AB227', '=SUM(G227:Z227)', st['getal'])
    log('  CF - opbrengsten: leeg DAEB-blok rij 200-227 (koprij %d, totaal rij %d in rij 194)' % (r0, DAEB_TOTAAL))


def kosten(sh: Sheet, st: dict, sjabloon: list, styles, log):
    z = Zetter(sh, 'CF - kosten', sjabloon)
    # werkelijk per post rechtstreeks uit FO - huidig (was '3. Restant bdgt'!F*1000)
    for p in K_POSTEN:
        f = sh.get('K%d' % p) or ''
        assert "'3. Restant bdgt'!F" in f, f
        z('K%d' % p, "=SUMIFS('FO - huidig'!$M:$M,'FO - huidig'!$Y:$Y,$B%d)" % p)
    # methode
    z('C58', 'Methode', st['kop'])
    for p in K_POSTEN:
        z('C%d' % p, 'Handmatig', st['invoer'])
        z('AD%d' % p, '=IF($C%d="Voorstel",IF(ABS(AC%d-D%d)<1,"OK","FOUT"),IF(AND(AC%d-D%d<1,AC%d-D%d>-1),"OK","FOUT"))'
          % (p, p + VOORSTEL, p, p, p, p, p))
    sh.add_list_validation('C59:C70', ['Handmatig', 'Voorstel'])
    # totaal = definitief
    for c in K_COLS:
        z('%s75' % c, '=SUMIFS(%s59:%s70,$C$59:$C$70,"<>Voorstel")+SUMIFS(%s%d:%s%d,$C$59:$C$70,"Voorstel")'
          % (c, c, c, 59 + VOORSTEL, c, 70 + VOORSTEL))
    z('AF57', 'Kolom C: Handmatig = deze regels · Voorstel = blok onderaan (rij 100-111). Rij 75 = definitief.',
      st['notitie'])
    z('AF4', 'Bouwtermijnen volgen de geplande bouwtermijnen van CF - opbrengsten; overschrijven mag.', st['notitie'])

    # voorstelblok
    for c in ['B'] + [num2col(i) for i in range(3, 31)]:
        z('%s93' % c, 'VOORSTEL · restant budget (PEW − verwerkt) gelijk verdeeld vanaf het eerste prognosekwartaal t/m '
                      'oplevering (tab Start) · kies per post "Voorstel" in kolom C' if c == 'B' else None, st['sectie'])
    z('B95', 'hulp: jaar', st['hulp_links'])
    z('B96', 'hulp: kwartaalindex (jaar × 4 + kw)', st['hulp_links'])
    for i, c in enumerate(K_COLS):
        prev = '%s95' % K_COLS[i - 1] if i else '""'
        z('%s95' % c, '=IF(ISNUMBER(%s$57),%s$57,%s)' % (c, c, prev), st['hulp'])
        z('%s96' % c, '=IF(AND(ISNUMBER(%s95),LEFT(%s$58&"",1)="Q"),%s95*4+VALUE(MID(%s$58,2,1)),0)' % (c, c, c, c), st['hulp'])
        z('%s98' % c, '=IF(%s$96>0,"Q"&(%s$96-4*%s$95)&" \'"&RIGHT(%s$95,2),"actuals")' % (c, c, c, c), st['kop'])
        z('%s99' % c, '=IF(%s$96<%s,"verwerkt",IF(%s$96<=MAX(%s,%s),"prognose","na oplevering"))' % (c, P_H, c, P_H, P_O),
          st['hulp'])
    for c, h in zip(['B', 'C', 'D', 'E', 'F', 'G', 'AC', 'AD'],
                    ['Post', '', 'PEW (huidig)', 'Verwerkt', 'Restant', 'Kw.', 'Totaal', 'Check']):
        z('%s98' % c, h or None, st['kop_links'] if c == 'B' else st['kop'])
    for p in K_POSTEN:
        r = p + VOORSTEL
        z('B%d' % r, '=IF(B%d="","",B%d)' % (p, p), st['tekst'])
        z('D%d' % r, '=N(D%d)' % p, st['getal'])
        z('E%d' % r, '=SUMIFS($H%d:$AA%d,$H$96:$AA$96,"<"&%s)' % (p, p, P_H), st['getal'])
        z('F%d' % r, '=MAX(0,D%d-E%d)' % (r, r), st['getal'])
        z('G%d' % r, '=COUNTIFS($H$96:$AA$96,">="&%s,$H$96:$AA$96,"<="&MAX(%s,%s))' % (P_H, P_H, P_O), st['aantal'])
        for c in K_COLS:
            z('%s%d' % (c, r), '=IF(%s$96<%s,N(%s%d),IF(%s$96<=MAX(%s,%s),IF($G%d=0,0,$F%d/$G%d),0))'
              % (c, P_H, c, p, c, P_H, P_O, r, r, r), st['getal'])
        z('AC%d' % r, '=SUM(H%d:AA%d)' % (r, r), st['getal_tot'])
        z('AD%d' % r, '=IF(B%d="","",IF(D%d-E%d<-1,"OVERSCHRIJDING",IF(ABS(AC%d-D%d)<1,"OK","FOUT")))' % (r, r, r, r, r),
          st['vet'])
    r = 70 + VOORSTEL + 1
    z('B%d' % r, 'Totaal voorstel', st['label_tot'])
    for c in ['C', 'D', 'E', 'F', 'G'] + K_COLS + ['AC']:
        z('%s%d' % (c, r), '=SUM(%s%d:%s%d)' % (c, 59 + VOORSTEL, c, 70 + VOORSTEL) if c not in ('C', 'G') else None,
          st['getal_tot'])
    # opmaak: gekozen methode zichtbaar maken
    grijs = styles.add_dxf('<font><strike val="1"/><color rgb="FF9CA3AF"/></font>')
    geel = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFDE68A"/></patternFill></fill>')
    groen = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFD1FAE5"/></patternFill></fill>')
    rood = styles.add_dxf('<fill><patternFill patternType="solid"><bgColor rgb="FFFECACA"/></patternFill></fill>')
    sh.add_cf('H59:AA70', '$C59="Voorstel"', grijs)
    sh.add_cf('H100:AA111', '$C59="Voorstel"', geel)
    sh.add_cf('AD100:AD111', 'AD100="OK"', groen)
    sh.add_cf('AD100:AD111', 'OR(AD100="FOUT",AD100="OVERSCHRIJDING")', rood)
    log('  CF - kosten: methode per post, voorstelblok rij 93-112, rij 75 = definitief, werkelijk uit FO - huidig')
