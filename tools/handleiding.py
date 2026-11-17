"""Handleiding: de uitleg gaat van de tabbladen naar één tabblad 'Handleiding'.

Op de tabbladen blijven korte koppen, korte labels en de cijfers staan (zoals in het RDFO). Lange uitlegteksten gaan naar
het hoofdstuk van hun tabblad in de Handleiding; op hun plek komt een klein '?' dat ernaar doorlinkt. De rode meldingen in
kolom A, de controles op Start en de teksten die de macro's of de code lezen blijven staan.

Regels (vaste tekst, geen formule):
- een label uit KORT (of KORT_RE) krijgt de korte versie; de lange gaat naar de Handleiding;
- een kop als '2 · INVOER · plakken of inlezen ...' houdt het eerste deel ('2 · INVOER'); de rest gaat naar de Handleiding
  en rechts van de kop komt een '?';
- elke andere tekst van meer dan GRENS tekens, of een cursieve notitie van meer dan GRENS_NOTITIE tekens, gaat naar de
  Handleiding en wordt een '?' in dezelfde cel.
"""
from __future__ import annotations

import copy
import re

from lxml import etree

from xlsxpkg import M, Sheet, col2num, num2col, split_ref


def _get(sh, ref):
    try:
        return sh.get(ref)
    except AttributeError:          # lege inlineStr-cel
        return None

HL = 'Handleiding'
GRENS = 60
GRENS_NOTITIE = 45

# tabbladen in de volgorde van de Handleiding (de naam tijdens de bouw; Prijslijst heet daarna Aansluiting opbrengsten)
BLADEN = ['Start', 'Prijslijst', 'CF - opbrengsten', 'CF - kosten', 'Dashboard', "Scenario's", '2. Ontw PEW', 'CF - hulp VENA']
TITEL = {'Prijslijst': 'Aansluiting opbrengsten'}

# niet aankomen: de controles en de prognose op Start (rij 21 t/m 45), en de teksten van de macro's
BEHOUD_RIJEN = {'Start': (21, 45)}

# hele blokken die naar de Handleiding gaan (rijen verbergen): de uitleg 'Zo werkt het' van het scenariomodel
BLOKKEN = {'Dashboard': [(136, 161, 'Scenariomodel · zo werkt het')]}

KORT = {
    'Verkocht · % van de woningen van dit type (1e kolom = al verkocht volgens Salesforce; volgt uit het verkooptempo)':
        'Verkocht %',
    'Getransporteerd · % van de woningen van dit type (volgt uit de tabel Transport)': 'Getransporteerd %',
    'Grondtermijn · koopsom grond (€, bij transport) · % van de koopsom': 'Grondtermijn (€)',
    'Bouwtermijnen (€) · % van de koopsom': 'Bouwtermijnen (€)',
    'Kadastrale kosten en aanvullende rente (aantal woningen, handmatig)': 'Kadastrale kosten en rente (aantal)',
    'Cashflow cumulatief · huidig (€, hulp voor het Dashboard)': 'Cashflow cumulatief (hulp)',
    'Extra opbrengst grondtermijn (€, eerder ontvangen)': 'Extra grondtermijn (€)',
    'KANS · grondtermijn naar het % op Prijslijst (kolom Y) · effect op de opbrengsten': 'KANS GRONDTERMIJN',
    'Verkoop en levering DAEB (aantal woningen per kwartaal; transport = verkoop)': 'Verkoop DAEB (aantal)',
    'Alle woningen · getransporteerd % per kwartaal': 'Alle woningen · getransp. %/kw',
    'Alle woningen · getransporteerd % cumulatief': 'Alle woningen · getransp. % cum.',
    'Grond % · grond / (grond + aanneemsom), excl. KAO': 'Grond %',
    'Type of blok in prijslijst (bv. A, of 4A;4B, * = alle)': 'Type / blok in prijslijst',
    'Kosten: budget aanneemsom (excl. KAO) voor CF - kosten': 'Budget aanneemsom (kosten)',
    'Niet toegewezen (geen of een ander SAP-nummer)': 'Niet toegewezen',
    'hulp · eerste kwartaal vanaf het huidige met een geplande termijn (1)': 'hulp · eerste geplande kwartaal',
    'Totaal kosten (CF - kosten, rij 127 ↔ PEW)': 'Totaal kosten ↔ PEW',
    'Extra grondtermijn eerder ontvangen (hoogste stand)': 'Extra grond eerder (hoogste)',
    'Extra opbrengst grondtermijn (totaal prognose)': 'Extra grond (totaal)',
    'Kosten in dat kwartaal · vorige prognose ↔ gerealiseerd': 'Kosten · vorig ↔ gerealiseerd',
    'Omzet in dat kwartaal · vorige prognose ↔ gerealiseerd': 'Omzet · vorig ↔ gerealiseerd',
    'Afgesloten periodes: cashflow = realisatie (SAP)': 'Afgesloten = realisatie (SAP)',
    'Bouwtermijnen en fees vallen binnen de periodes': 'Termijnen binnen de periodes',
    'DAEB: elke fee-termijn heeft een geldig kwartaal en bedrag': 'DAEB-termijnen geldig',
    'Invoer sluit aan op het FO (kosten en opbrengsten)': 'Invoer = FO',
    'Grondtermijn (% koopsom) · bouwtermijnen naar rato lager': 'Grondtermijn (% koopsom)',
}
KORT_RE = [
    (re.compile(r'^Budget aanneemsom \(kosten, excl\. KAO\) en KAO/meerwerk · type (\d+)$'), r'Budget aanneemsom · type \1'),
    (re.compile(r'^Totaal opbrengsten \(CF - opbrengsten, rij \d+ ↔ PEW\)$'), 'Totaal opbrengsten ↔ PEW'),
]

# vaste uitleg bovenaan de Handleiding (de hoofdstukken per tabblad vult de bouw met de verplaatste teksten)
ALGEMEEN = [
    ('Kleuren en meldingen', [
        ('Geel', 'invullen (handmatig of met een knop).'),
        ('Wit', 'berekend: niet overschrijven, tenzij het erbij staat.'),
        ('Grijs', 'uit een export (SAP of Salesforce).'),
        ('Lichtblauw', 'totaal.'),
        ('Rode melding in kolom A', 'op CF - opbrengsten en CF - kosten: op die regel moet je iets doen (bv. "4 woningen nog '
                                    'te verdelen"). B1 telt de meldingen.'),
        ('Controles', 'op Start, blok 4: alles op OK voordat je de PowerPoint maakt.'),
        ('?', 'achter een kop of naast een veld: de uitleg in deze handleiding (klik erop).'),
    ]),
    ('Planning en verdeling (kort)', [
        ('Planning per woningtype', 'CF - opbrengsten, rij Planning boven elk type: start bouw en oplevering als tekst (bv. '
                                    'Q1 2027) en de spreiding (kwartalen per termijn).'),
        ('Kw na start', 'kolom F per termijn. Leeg: gelijk verdeeld tussen start bouw en oplevering; zonder oplevering de '
                        'kwartalen van type 1.'),
        ('Standaard schema', 'een type zonder eigen termijnschema (Aansluiting opbrengsten, kolom L leeg) volgt type 1.'),
        ('Naar rato', 'Aansluiting opbrengsten, blok van het type, Verdeling = naar rato met de types (bv. 1;2): opbrengsten '
                      'en kosten volgen die types (zoals parkeren in het oude FO).'),
        ('Nog te verdelen', 'verkooptempo: aantal woningen min t/m nu en de kwartalen; anders dan 0 = rode melding.'),
    ]),
    ('Knoppen', [
        ('Oud FO inlezen', 'neemt een FO in de oude opzet over (planning, termijnen, kosten, PEW); daarna de controles op Start.'),
        ('Nieuw kwartaal starten', 'de huidige cashflow en planning worden de vorige prognose; FO - huidig gaat naar FO - vorig.'),
        ('PowerPoint maken', 'vult het sjabloon met de cijfers, tabellen en grafieken en bewaart een nieuwe presentatie.'),
        ('Verborgen tabbladen', 'toont of verbergt de naslag: Bouwnummers, CF vorig en huidig, Woningtypes, Scenario kosten en '
                                'de exports.'),
    ]),
]


# ------------------------------------------------------------------------------------------------------------------
def _font_italic(styles, s):
    try:
        xf = styles.t.find(M + 'cellXfs')[int(s or 0)]
        font = styles.t.find(M + 'fonts')[int(xf.get('fontId', '0'))]
        return font.find(M + 'i') is not None and font.find(M + 'i').get('val', '1') != '0'
    except (IndexError, ValueError, TypeError):
        return False


def _invoer(styles, s):
    """gele invulcel (bv. de toelichting per post): dat is invoer van de gebruiker, die blijft staan"""
    try:
        xf = styles.t.find(M + 'cellXfs')[int(s or 0)]
        fill = styles.t.find(M + 'fills')[int(xf.get('fillId', '0'))]
        fg = fill.find('.//' + M + 'fgColor')
        return fg is not None and (fg.get('rgb') or '').upper().endswith('FFF2CC')
    except (IndexError, ValueError, TypeError):
        return False


def _donker(styles, s):
    """vulling van de cel donker (koppen): dan een wit vraagteken"""
    try:
        xf = styles.t.find(M + 'cellXfs')[int(s or 0)]
        fill = styles.t.find(M + 'fills')[int(xf.get('fillId', '0'))]
        fg = fill.find('.//' + M + 'fgColor')
        if fg is None:
            return False
        if fg.get('theme') in ('1', '3', '6'):
            return True
        rgb = fg.get('rgb')
        if not rgb:
            return False
        r, g, b = (int(rgb[-6:][i:i + 2], 16) for i in (0, 2, 4))
        return 0.299 * r + 0.587 * g + 0.114 * b < 120
    except (IndexError, ValueError, TypeError):
        return False


def vraag_stijl(styles, s, cache):
    """stijl voor het '?': die van de cel (vulling, rand) met een kleine onderstreepte letter (navy, of wit op donker)"""
    key = s or '0'
    if key in cache:
        return cache[key]
    wit = _donker(styles, s)
    font = etree.fromstring('<font xmlns="%s"><b/><u/><sz val="9"/><color rgb="%s"/><name val="Heijmans-Light"/></font>'
                            % (M[1:-1], 'FFFFFFFF' if wit else 'FF003366'))
    fid = styles._add('fonts', font)
    try:
        xf = copy.deepcopy(styles.t.find(M + 'cellXfs')[int(key)])
    except (IndexError, ValueError):
        xf = etree.Element(M + 'xf', numFmtId='0', fontId='0', fillId='0', borderId='0', xfId='0')
    xf.set('fontId', str(fid))
    xf.set('applyFont', '1')
    for a in xf.findall(M + 'alignment'):
        xf.remove(a)
    etree.SubElement(xf, M + 'alignment', horizontal='center', vertical='center')
    xf.set('applyAlignment', '1')
    cache[key] = styles._add('cellXfs', xf)
    return cache[key]


def _breedtes(sh):
    w = {}
    cols = sh.t.find(M + 'cols')
    if cols is not None:
        for c in cols:
            for i in range(int(c.get('min')), int(c.get('max')) + 1):
                w[i] = float(c.get('width', '8.43'))
    return lambda i: w.get(i, 8.43)


def _is_label(v):
    return (isinstance(v, str) and v and not v.startswith('=') and len(v) <= GRENS and re.search(r'[A-Za-z]', v)
            and not re.match(r'^[\d.,% -]+$', v))


def _label_links(sh, r, c):
    """het korte label bij een uitlegtekst (titel in de Handleiding): links op dezelfde rij (hooguit 6 kolommen), anders
    de kop erboven in dezelfde kolom (hooguit 3 rijen)"""
    for cc in range(c - 1, max(0, c - 7), -1):
        v = _get(sh, '%s%d' % (num2col(cc), r))
        if _is_label(v):
            return v.strip()
    for rr in range(r - 1, max(0, r - 4), -1):
        v = _get(sh, '%s%d' % (num2col(c), rr))
        if _is_label(v):
            return v.strip()
    return ''


def _kop(t):
    """'2 · INVOER · plakken of inlezen ...' -> ('2 · INVOER', 'plakken of inlezen ...'): de delen in hoofdletters blijven
    (ook 'ZO WORDT DE CF' met de uitleg tussen haakjes erachter); anders None"""
    delen = [d.strip() for d in t.split(' · ')]
    if len(delen) < 2 and '(' not in t:
        return None
    kop, rest = [], []
    for i, d in enumerate(delen):
        m = re.match(r'^([^()]*?)\s*\((.*)\)\s*$', d)
        kern, uitleg = (m.group(1), m.group(2)) if m else (d, '')
        woorden = re.sub(r'^\d+\s*', '', kern)
        if not rest and (re.search(r'[A-Z]', woorden) and woorden.upper() == woorden or (i == 0 and re.match(r'^\d+$', kern))):
            kop.append(kern)
            if uitleg:
                rest.append(uitleg)
        else:
            rest.append(d)
    if not kop and len(delen) >= 2 and len(delen[0]) <= 35 and delen[0][:1].isupper():
        return delen[0], ' · '.join(delen[1:])          # kop van een grafiek: 'Cashflow per kwartaal · balken: ...'
    if not kop or not rest or len(' · '.join(kop)) > 48:
        return None
    return ' · '.join(kop), ' · '.join(rest)


def verwerk(dst, part_of, sst, styles, st, note):
    """teksten verplaatsen en de Handleiding vullen; geeft het aantal verplaatste teksten"""
    hoofdstukken = []            # (titel, blad, [(label, tekst, [(blad, ref)])])
    links = []                   # (blad, ref, hoofdstuk-index, regel-index): '?' na het vullen naar de juiste rij
    cache = {}
    n = 0
    # Start: het stappenplan (blok 5) gaat helemaal naar de Handleiding
    stappen = []
    if 'Start' in part_of:
        sh = Sheet(dst.x(part_of['Start']), sst)
        for r in range(1, 120):
            v = _get(sh, 'B%d' % r)
            if isinstance(v, str) and v.startswith('5 · STAPPEN'):
                for rr in range(r + 1, r + 20):
                    w = _get(sh, 'B%d' % rr)
                    if not (isinstance(w, str) and re.match(r'^\d+\.\s', w)):
                        break
                    stappen.append(w)
                    sh.set_value('B%d' % rr, None)
                    sh.row(rr).set('hidden', '1')
                sh.set_value('B%d' % r, '5 · STAPPEN PER KWARTAAL', sh.style_of('B%d' % r))
                links.append(('Start', 'E%d' % r, -1, sh.style_of('B%d' % r)))
                break
        dst.put(part_of['Start'], sh.t)
    for blad in BLADEN:
        if blad not in part_of:
            continue
        sh = Sheet(dst.x(part_of[blad]), sst)
        breedte = _breedtes(sh)
        regels = []
        h_i = len(hoofdstukken)
        behoud = BEHOUD_RIJEN.get(blad)
        # hele blokken
        for r1, r2, titel in BLOKKEN.get(blad, []):
            for r in range(r1, r2 + 1):
                rij = sh.row(r, create=False)
                if rij is None:
                    continue
                teksten = [_get(sh, c.get('r')) for c in rij]
                teksten = [t for t in teksten if isinstance(t, str) and t and not t.startswith('=')]
                if r == r1 and len(teksten) == 1:
                    regels.append((titel, '', []))
                elif teksten:
                    regels.append((teksten[0] if len(teksten) > 1 else '', teksten[-1], []))
                for c in rij:
                    sh.set_value(c.get('r'), None)
                rij.set('hidden', '1')
                n += 1
        for rij in list(sh.data):
            r = int(rij.get('r'))
            if behoud and behoud[0] <= r <= behoud[1]:
                continue
            if any(r1 <= r <= r2 for r1, r2, _ in BLOKKEN.get(blad, [])):
                continue
            for c in list(rij):
                ref = c.get('r')
                t = _get(sh, ref)
                if not isinstance(t, str) or not t or t.startswith('='):
                    continue
                s = c.get('s')
                if _invoer(styles, s):
                    continue
                kort = KORT.get(t)
                if kort is None:
                    for rx, vervang in KORT_RE:
                        if rx.match(t):
                            kort = rx.sub(vervang, t)
                            break
                if kort is not None:
                    sh.set_value(ref, kort, s)
                    regels.append((kort, t, []))
                    n += 1
                    continue
                lang = len(t) > GRENS or (len(t) > GRENS_NOTITIE and _font_italic(styles, s))
                if not lang:
                    continue
                col, _ = split_ref(ref)
                ci = col2num(col)
                k = _kop(t)
                if k:
                    kop, rest = k
                    sh.set_value(ref, kop, s)
                    # '?' rechts van de kop: de eerste kolom voorbij de tekst
                    nodig, cc = len(kop) * 1.15 + 1, ci
                    while nodig > 0 and cc < ci + 12:
                        nodig -= breedte(cc)
                        cc += 1
                    doel = '%s%d' % (num2col(cc), r)
                    vrij = _get(sh, doel) in (None, '')
                    regels.append((kop, rest, [(blad, doel, sh.style_of(doel) or s)] if vrij else []))
                else:
                    titel = _label_links(sh, r, ci)
                    sh.set_value(ref, None, s)
                    regels.append((titel, t, [(blad, ref, s)]))
                n += 1
        dst.put(part_of[blad], sh.t)
        uniek, waar = [], {}
        for label, tekst, plekken in regels:
            sleutel = tekst or label
            if sleutel in waar:
                uniek[waar[sleutel]][2].extend(plekken)
            else:
                waar[sleutel] = len(uniek)
                uniek.append((label, tekst, list(plekken)))
        hoofdstukken.append((TITEL.get(blad, blad), blad, uniek))
    # de Handleiding vullen
    hl = Sheet(dst.x(part_of[HL]), sst)
    rij = 6
    anker = {}
    inhoud = []

    def zet(ref, v, stijl):
        hl.set(ref, v, st[stijl])

    zet('B2', 'Handleiding', 'titel')
    zet('B3', '=HYPERLINK("#Start!A1","← Start")', 'link')
    zet('B4', 'Uitleg bij het FO-werkbestand. Op de tabbladen zelf staan alleen korte koppen en labels; een ? brengt je hier.',
        'notitie')
    alle = [('Werkwijze per kwartaal', None, [('', re.sub(r'^\d+\.\s+', '', s), []) for s in stappen])]
    alle += [(t, None, [(a, b, []) for a, b in regels]) for t, regels in ALGEMEEN]
    alle += hoofdstukken
    rij = 6 + len(alle) + 2
    for i, (titel, blad, regels) in enumerate(alle):
        if not regels:
            continue
        anker[i] = rij
        inhoud.append((titel, rij))
        zet('B%d' % rij, titel, 'sectie')
        zet('C%d' % rij, None, 'sectie')
        if blad:
            zet('D%d' % rij, '=HYPERLINK("#\'%s\'!A1","→ naar %s")' % (blad.replace("'", "''"), TITEL.get(blad, blad)), 'link')
        rij += 1
        for j, (label, tekst, plekken) in enumerate(regels):
            if label and not tekst:
                zet('B%d' % rij, label, 'vet')
            else:
                zet('B%d' % rij, label, 'vet')
                zet('C%d' % rij, tekst, 'tekst')
            for (b, ref, s) in plekken:
                links.append((b, ref, rij, s))
            rij += 1
        rij += 1
    for k, (titel, r) in enumerate(inhoud):
        zet('B%d' % (6 + k), '=HYPERLINK("#Handleiding!B%d","%s")' % (r, titel.replace('"', '""')), 'link')
    dst.put(part_of[HL], hl.t)
    # de '?'-links op de tabbladen
    per_blad = {}
    for b, ref, doelrij, s in links:
        per_blad.setdefault(b, []).append((ref, anker.get(0, 6) if doelrij == -1 else doelrij, s))
    for b, items in per_blad.items():
        sh = Sheet(dst.x(part_of[b]), sst)
        for ref, doelrij, s in items:
            sh.set_formula(ref, '=HYPERLINK("#Handleiding!B%d","?")' % doelrij, vraag_stijl(styles, s, cache))
        dst.put(part_of[b], sh.t)
    note('Handleiding: %d teksten verplaatst of ingekort, %d hoofdstukken, %d ?-links' % (n, len(inhoud), len(links)))
    return n
