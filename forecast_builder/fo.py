"""Invoer uit een FO-werkboek (Financieel Overzicht van Heijmans Vastgoed) halen.

Het FO heeft per project dezelfde opbouw (tabbladen '1. Cashflow', 'CF - opbrengsten', 'FO - actuals'); alles wordt op
labels gezocht, niet op vaste celadressen, zodat kleine verschillen tussen projecten geen kwaad kunnen:
  * '1. Cashflow': de kwartaalcashflow (koprij met 'Jaar' in kolom B; daaronder jaar, Q, kosten, % kosten, omzet, % omzet,
    CF, CF x 1000, CF vorig kwartaal) -> tab Invoer B:J. Rijen zonder Q (jaartotalen) blijven jaarrijen.
  * 'FO - actuals': laatste boekjaar/kwartaal met boekingen -> Dashboard 'Actuals t/m'.
  * 'CF - opbrengsten': per woningtype een blok 'Termijnen k - naam' (bouwplanning: per termijn het aantal woningen per
    kwartaal), een blok 'Omzet HVG Termijnen' (koopsom en termijnen in %, € en € per woning; kopersmeerwerk, kadastrale
    kosten, overige) en het verkooptempo (verkocht per kwartaal; 'actuals' = al verkocht). Daaruit: naam, aantal, koopsom
    per woning (grond + termijnen, excl. btw), grondtermijn %, start bouw, bouwtermijnen (naam, %, bouwkwartaal; een termijn
    die over meer kwartalen is verdeeld wordt gesplitst), extra's per woning en verkocht/transport per kwartaal.
  * DAEB-fees: een rij 'AK fee' of 'Bijkomende kosten' met alleen een totaal in € (geen %, geen kwartalen) is een
    fee-component; elke rij daarna met een % of € én een kwartaal of 'actuals' is een termijn van die component, wat het
    label ook is ('na akkoord SO', 'bij start bouw', ...). 'Onvoorzien' wordt overgeslagen. Zo'n type wordt DAEB: alleen
    de fees, geen koopsom, bouwtermijnen of extra's.
Meldingen (ontbrekend tabblad, samengevoegde termijnen, overgeslagen rijen) staan in p.fo['waarschuwingen'].
De macro 'Ophalen uit FO' in het werkboek doet hetzelfde (zie vba/CashflowNaarPowerPoint_v10.bas, UitFOOphalen).
"""
import os
import re
from datetime import datetime

from openpyxl import load_workbook

from . import layout as LY
from .data import PARAM_STANDAARD, ProjectData, TypeData

KW = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}


def _txt(v):
    return str(v).strip() if v is not None else ""


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _jaar(v):
    """Jaartal uit een getal of tekst ('2026'); None als het geen jaar is."""
    try:
        j = int(float(str(v).strip()))
    except (TypeError, ValueError):
        return None
    return j if 1990 <= j <= 2100 else None


def _kw(v):
    t = _txt(v).upper().replace(" ", "")
    if t in KW:
        return KW[t]
    if t in ("1", "2", "3", "4"):
        return int(t)
    return None


def _zoek_blad(wb, *namen):
    """Werkblad op naam (hoofdletterongevoelig, spaties genegeerd); None als het ontbreekt."""
    norm = lambda s: re.sub(r"\s+", "", s.lower())  # noqa: E731
    for ws in wb.worksheets:
        if norm(ws.title) in [norm(n) for n in namen]:
            return ws
    return None


def _kolomkaart(ws, rij_q):
    """{kolom: (jaar, kw)} uit een koprij met Q1..Q4 en de jaartallen in de rij erboven (één jaartal per vier kwartalen);
    ook de kolom met 'actuals' (of None)."""
    kaart, actuals = {}, None
    jaar = None
    for c in range(1, ws.max_column + 1):
        j = _jaar(ws.cell(rij_q - 1, c).value)
        if j is not None:
            jaar = j
        v = ws.cell(rij_q, c).value
        if _txt(v).lower().startswith("actual"):
            actuals = c
        q = _kw(v)
        if q is not None and jaar is not None:
            kaart[c] = (jaar, q)
    return kaart, actuals


def _tellingen(ws, r, kaart, c_act=None):
    """[(jaar, kw, aantal)] van de kwartaalcellen in rij r (aantal <> 0); de kolom 'actuals' als (None, None, aantal) vooraan."""
    uit = []
    if c_act:
        v = _num(ws.cell(r, c_act).value)
        if v:
            uit.append((None, None, v))
    for c, (j, q) in sorted(kaart.items()):
        v = _num(ws.cell(r, c).value)
        if v:
            uit.append((j, q, v))
    return uit


def _typenaam(ws, r, k):
    """Naam van het type: uit de bloktitel 'Termijnen k - naam', anders de cel naast de titel (bv. '<300K VON'), anders de
    tekst na 'Termijnen ' als dat geen nummer is, anders 'Type k'."""
    t = _txt(ws.cell(r, 2).value)
    if " - " in t:
        return t.split(" - ", 1)[1].strip() or f"Type {k}"
    naast = ws.cell(r, 3).value
    if _txt(naast) and _num(naast) is None:
        return _txt(naast)
    rest = t[10:].strip() if t.lower().startswith("termijnen ") else ""
    if rest and not rest.replace(".", "").isdigit():
        return rest
    return f"Type {k}"


def voeg_vorige_prognose_toe(project, pad_vorig):
    """Vorige prognose uit een ouder FO: per type de start bouw (blok 4) en per fee-termijn het kwartaal (blok 6). Typen op naam, anders op
    positie; fee-termijnen op mijlpaal, anders op positie. Meldingen in project.fo['waarschuwingen']."""
    vorig = lees_fo(pad_vorig)
    meldingen = project.fo.setdefault("waarschuwingen", [])
    op_naam = {t.naam.strip().lower(): t for t in vorig.types}
    for k, t in enumerate(project.types):
        tv = op_naam.get(t.naam.strip().lower()) or (vorig.types[k] if k < len(vorig.types) else None)
        if tv is None:
            meldingen.append(f"{t.naam}: niet gevonden in het vorige FO ({os.path.basename(pad_vorig)}); vorige prognose leeg gelaten.")
            continue
        if tv.naam.strip().lower() != t.naam.strip().lower():
            meldingen.append(f"{t.naam}: op naam niet gevonden in het vorige FO; het type op dezelfde positie ({tv.naam}) is gebruikt.")
        t.vorig_start_jaar, t.vorig_start_kw = tv.start_jaar, tv.start_kw
        if any(lab for lab, _, _ in t.fee_termijnen):
            kw_vorig = {str(lab).strip().lower(): kw for lab, _, kw in tv.fee_termijnen if lab}
            t.fee_vorig = []
            for i, (lab, _, _) in enumerate(t.fee_termijnen):
                kw = kw_vorig.get(str(lab).strip().lower()) if lab else None
                if kw is None and lab and i < len(tv.fee_termijnen) and tv.fee_termijnen[i][0]:
                    kw = tv.fee_termijnen[i][2]
                t.fee_vorig.append(kw)
    project.fo["vorig_bestand"] = os.path.basename(pad_vorig)
    return project


def _is_fee_label(ll):
    return bool(re.search(r"\bfees?\b", ll)) or ll.startswith("ak-fee") or ll.startswith("bijkomende kosten") or ll.startswith("onvoorzien")


def _binnen_limiet(rijen, n_max, naam, wat, waarschuwingen):
    """[(label, waarde, bouwkwartaal)] terugbrengen tot n_max rijen: gesplitste delen (label 'x (i/n)') van dezelfde termijn worden
    samengevoegd bij het grootste deel (kleinste deel eerst), met een melding. Zonder splitsingen om samen te voegen: foutmelding."""
    rijen = list(rijen)
    samengevoegd = []
    while len(rijen) > n_max:
        basis = lambda lab: re.sub(r" \(\d+/\d+\)$", "", lab)
        groepen = {}
        for i, (lab, w, kw) in enumerate(rijen):
            if re.search(r" \(\d+/\d+\)$", lab):
                groepen.setdefault(basis(lab), []).append(i)
        groepen = {b: ix for b, ix in groepen.items() if len(ix) > 1}
        if not groepen:
            raise ValueError(f"{naam}: {len(rijen)} {wat} in het FO, het werkboek heeft ruimte voor {n_max}.")
        b, ix = max(groepen.items(), key=lambda kv: len(kv[1]))
        klein = min(ix, key=lambda i: abs(rijen[i][1] or 0))
        groot = max((i for i in ix if i != klein), key=lambda i: abs(rijen[i][1] or 0))
        lab_g, w_g, kw_g = rijen[groot]
        rijen[groot] = (lab_g, round((w_g or 0) + (rijen[klein][1] or 0), 6), kw_g)
        del rijen[klein]
        samengevoegd.append(b)
        n = len([i for i, (lab, _, _) in enumerate(rijen) if basis(lab) == b and re.search(r" \(\d+/\d+\)$", lab)])
        k = 0
        for i, (lab, w, kw) in enumerate(rijen):
            if basis(lab) == b and re.search(r" \(\d+/\d+\)$", lab):
                k += 1
                rijen[i] = (b if n == 1 else f"{b} ({k}/{n})", w, kw)
    if samengevoegd:
        waarschuwingen.append(f"{naam}: meer dan {n_max} {wat} na het splitsen over kwartalen; kleinste delen samengevoegd bij het "
                              f"grootste deel van {', '.join(sorted(set(samengevoegd)))}.")
    return rijen


def _idx(j, q):
    return j * 4 + q


def lees_fo(pad):
    wb = load_workbook(pad, data_only=True)
    p = ProjectData()
    p.params = dict(PARAM_STANDAARD)
    p.fo = {"bestand": os.path.basename(pad), "datum": datetime.now().strftime("%d-%m-%Y %H:%M"), "waarschuwingen": []}
    waarschuwingen = p.fo["waarschuwingen"]

    # ---- 1. Cashflow -> periodes ----------------------------------------------------------------------------------
    ws = _zoek_blad(wb, "1. Cashflow", "Cashflow")
    if ws is None:
        raise ValueError("Tabblad '1. Cashflow' niet gevonden in het FO.")
    # koprij 'Jaar' ergens in A1:F30; de kolommen erachter: Q, kosten, % kosten, omzet, % omzet, CF, CF x 1000, CF vorig
    kop = next(((r, c) for r in range(1, 31) for c in range(1, 7) if _txt(ws.cell(r, c).value).lower() == "jaar"), None)
    if kop is None:
        raise ValueError(f"Koprij 'Jaar' niet gevonden op '{ws.title}' (gezocht in A1:F30).")
    r, cj = kop[0] + 1, kop[1]
    while _jaar(ws.cell(r, cj).value) is not None:
        q = _kw(ws.cell(r, cj + 1).value)
        v = [_num(ws.cell(r, cj + k).value) for k in range(2, 9)]
        p.periodes.append({"jaar": _jaar(ws.cell(r, cj).value), "kw": f"Q{q}" if q else None,
                           "kosten": v[0], "pct_kosten": v[1], "omzet": v[2], "pct_omzet": v[3], "cf": v[4], "cf1000": v[5], "cf_vorig": v[6]})
        r += 1
    for rr in range(r, r + 6):          # totalen onder de tabel ('Totaal:' in de kolom van Q)
        if _txt(ws.cell(rr, cj + 1).value).lower().startswith("totaal"):
            p.fo["kosten"] = _num(ws.cell(rr, cj + 2).value)
            p.fo["opbrengsten"] = _num(ws.cell(rr, cj + 4).value)
    n_per = len(p.periodes)
    periode_van = {}
    for i, per in enumerate(p.periodes):
        if per["kw"]:
            periode_van[_idx(per["jaar"], int(per["kw"][1]))] = i

    # ---- FO - actuals -> actuals t/m -------------------------------------------------------------------------------
    wa = _zoek_blad(wb, "FO - actuals", "actuals")
    if wa is not None:
        koppen = {_txt(wa.cell(1, c).value).lower(): c for c in range(1, wa.max_column + 1)}
        cj, cq, cw = koppen.get("boekjaar"), koppen.get("q"), koppen.get("waarde/trval")
        laatste = None
        if cj and cq:
            for r in range(2, wa.max_row + 1):
                j, q = _jaar(wa.cell(r, cj).value), _kw(wa.cell(r, cq).value)
                if j and q and (cw is None or wa.cell(r, cw).value is not None):
                    if laatste is None or _idx(j, q) > _idx(*laatste):
                        laatste = (j, q)
        if laatste:
            p.params["actuals_jaar"], p.params["actuals_kw"] = laatste
            p.fo["actuals"] = f"Q{laatste[1]} {laatste[0]}"
        else:
            waarschuwingen.append("'FO - actuals' heeft geen rij met boekjaar en kwartaal: 'Actuals t/m' is niet uit het FO gehaald.")
    else:
        waarschuwingen.append("tabblad 'FO - actuals' ontbreekt: 'Actuals t/m' is niet uit het FO gehaald.")
    idx_act = _idx(p.params["actuals_jaar"], p.params["actuals_kw"])

    # ---- CF - opbrengsten -> woningtypes ----------------------------------------------------------------------------
    wo = _zoek_blad(wb, "CF - opbrengsten", "CF-opbrengsten", "Opbrengsten")
    if wo is None:
        raise ValueError("Tabblad 'CF - opbrengsten' niet gevonden in het FO.")
    blokken, omzet, verkoop = [], [], None
    for r in range(1, wo.max_row + 1):
        b = _txt(wo.cell(r, 2).value)
        lb = b.lower()
        if lb.startswith("termijnen ") and " in €" not in lb and _txt(wo.cell(r - 1, 6).value).lower().startswith("# won"):
            blokken.append(r)
        elif lb.startswith("omzet hvg") or lb.startswith("omzet termijnen"):
            omzet.append(r)
        elif lb == "verkooptempo" and verkoop is None:
            verkoop = r
    p.types = []
    for k, r0 in enumerate(blokken):
        naam = _typenaam(wo, r0, k + 1)
        aantal = _num(wo.cell(r0 - 1, 7).value) or 0
        kaart, c_act = _kolomkaart(wo, r0)
        # bouwplanning: per termijn het aantal woningen per kwartaal
        planning = []
        r = r0 + 1
        while _txt(wo.cell(r, 2).value):
            planning.append((_txt(wo.cell(r, 2).value), _tellingen(wo, r, kaart), _num(wo.cell(r, c_act).value) if c_act else None))
            r += 1
        start = None
        for tn, tel, _ in planning:
            if tel:
                start = (tel[0][0], tel[0][1])
                break
        if start and start[0] is None:
            start = None
        t = TypeData(naam, int(aantal) if aantal == int(aantal) else aantal, None, start[0] if start else None, start[1] if start else None, None, [])
        if k < len(omzet):
            ro = omzet[k]
            kaart_o, c_act_o = _kolomkaart(wo, ro)
            koop_pw, koop_tot, transport = 0.0, 0.0, []
            termijnen, extras, fees = [], [], []
            r = ro + 1
            comp = None
            while _txt(wo.cell(r, 2).value):
                lab = _txt(wo.cell(r, 2).value)
                ll = lab.lower()
                pct, eur, pw = _num(wo.cell(r, 3).value), _num(wo.cell(r, 4).value), _num(wo.cell(r, 5).value)
                tel = _tellingen(wo, r, kaart_o)
                act = _num(wo.cell(r, c_act_o).value) if c_act_o else None
                comp_rij = bool(eur) and not pct and not tel and not act      # alleen een totaal in €: fee-component
                if ll == "koopsom":
                    koop_pw, koop_tot = pw or 0.0, eur or 0.0
                    transport = _tellingen(wo, r, kaart_o, c_act_o)
                elif (comp_rij and _is_fee_label(ll)) or (comp is not None and comp_rij):
                    # DAEB: componentrij (totaal in €, geen % en geen kwartalen); 'onvoorzien' doet niet mee
                    comp = (lab, eur)
                    if ll.startswith("onvoorzien"):
                        comp = (None, None)
                    else:
                        fees.append((comp, []))
                elif comp is not None and (pct or eur) and (tel or act):
                    # termijn van de laatste component (label vrij: 'na akkoord SO', 'bij start bouw', ...)
                    if comp[0] is not None:
                        kw = "actuals" if (act and not tel) else f"{tel[0][0]} Q{tel[0][1]}"
                        fees[-1][1].append((lab, pct if pct else eur, kw))
                elif comp is not None and _is_fee_label(ll):
                    waarschuwingen.append(f"{naam}: rij '{lab}' in het fee-blok heeft geen bedrag met kwartaal en is overgeslagen.")
                elif pct is not None and pct != 0:
                    termijnen.append((lab, pct, pw or 0.0, tel))
                elif pw and abs(pw) >= 0.005:                 # extra per woning (minder dan een cent telt niet)
                    extras.append((lab, pw, tel, act))
                r += 1
            # koopsom per woning = grond (rij Koopsom) + termijnen per woning
            koopsom = koop_pw + sum(pw for _, _, pw, _ in termijnen)
            t.koopsom = round(koopsom, 2) if koopsom else None
            t.grond_pct = round(koop_pw / koopsom, 6) if koopsom else None
            idx_start = _idx(*start) if start else None
            for lab, pct, pw, tel in termijnen:
                # kwartalen uit de bouwplanning (zelfde naam), anders uit het €-blok; verdeeld over meer kwartalen -> gesplitst
                tel_p = next((tp for tn, tp, _ in planning if tn.lower() == lab.lower() and tp), None) or tel
                if tel_p and idx_start:
                    som = sum(v for _, _, v in tel_p)
                    delen = [(j, q, v / som) for j, q, v in tel_p] if som else []
                    for i, (j, q, deel) in enumerate(delen):
                        suffix = f" ({i + 1}/{len(delen)})" if len(delen) > 1 else ""
                        t.termijnen.append((lab + suffix, round(pct * deel, 6), _idx(j, q) - idx_start + 1))
                else:
                    t.termijnen.append((lab, pct, None))
            for lab, pw, tel, act in extras:
                tel_p = next((tp for tn, tp, _ in planning if tn.lower() == lab.lower() and tp), None) or tel
                if tel_p and idx_start and not act:
                    som = sum(v for _, _, v in tel_p)
                    delen = [(j, q, v / som) for j, q, v in tel_p] if som else []
                    for i, (j, q, deel) in enumerate(delen):
                        suffix = f" ({i + 1}/{len(delen)})" if len(delen) > 1 else ""
                        t.extras.append((lab + suffix, round(pw * deel, 2), _idx(j, q) - idx_start + 1))
                else:
                    t.extras.append((lab, round(pw, 2), None))        # bij transport
            t.termijnen = _binnen_limiet(t.termijnen, LY.N_TERMIJNEN, naam, "bouwtermijnen", waarschuwingen)
            t.extras = _binnen_limiet(t.extras, LY.N_EXTRA, naam, "extra's", waarschuwingen)
            if fees:
                t.soort = "DAEB"
                t.koopsom, t.grond_pct, t.termijnen, t.extras = None, None, [], []
                if len(fees) > LY.N_FEE_COMP:
                    waarschuwingen.append(f"{naam}: {len(fees)} fee-componenten in het FO, alleen de eerste {LY.N_FEE_COMP} zijn overgenomen "
                                          f"({', '.join(c[0] for c, _ in fees[LY.N_FEE_COMP:])} niet).")
                t.fee_comp = [(c[0], c[1]) for c, _ in fees][:LY.N_FEE_COMP]
                rijen = []
                for i, (c, termijnen_c) in enumerate(fees[:LY.N_FEE_COMP]):
                    if len(termijnen_c) > LY.N_FEE_PER_COMP:
                        waarschuwingen.append(f"{naam}: {c[0]} heeft {len(termijnen_c)} termijnen in het FO, alleen de eerste "
                                              f"{LY.N_FEE_PER_COMP} zijn overgenomen.")
                    rijen_c = [(lab, w, kw) for lab, w, kw in termijnen_c][:LY.N_FEE_PER_COMP]
                    rijen_c += [("", None, None)] * (LY.N_FEE_PER_COMP - len(rijen_c))
                    rijen += rijen_c
                t.fee_termijnen = rijen
            t._transport = transport
        t._blok = k                                   # positie van het blok in het FO (ook lege blokken tellen mee in het verkooptempo)
        p.types.append(t)
    p.types = [t for t in p.types if (t.aantal or 0) > 0]
    # verkoop per kwartaal per type (blok Verkooptempo: rijen in dezelfde volgorde als de typen) en transport
    verkocht = {}
    if verkoop is not None:
        kaart_v, c_act_v = _kolomkaart(wo, verkoop)
        r = verkoop + 1
        i = 0
        while _txt(wo.cell(r, 2).value):
            verkocht[i] = _tellingen(wo, r, kaart_v, c_act_v)   # kwartaalcellen tellen op tot het aantal; 'actuals' = al verkocht
            r += 1
            i += 1
    typen_namen = [t.naam for t in p.types]
    nT = len(p.types)
    p.verkocht = [[None] * nT for _ in range(n_per)]
    p.transport = [[None] * nT for _ in range(n_per)]
    pos_act = periode_van.get(idx_act)
    if pos_act is None and periode_van:
        pos_act = min(periode_van.values())

    def zet(matrix, k, i, v):
        if v:
            matrix[i][k] = (matrix[i][k] or 0) + int(round(v))

    for k, t in enumerate(p.types):
        for jj, q, v in verkocht.get(getattr(t, "_blok", k), []):
            i = pos_act if jj is None else periode_van.get(_idx(jj, q))
            if i is not None:
                zet(p.verkocht, k, i, v)
        for jj, q, v in getattr(t, "_transport", []):
            i = pos_act if jj is None else periode_van.get(_idx(jj, q))
            if i is not None:
                zet(p.transport, k, i, v)
        for attr in ("_transport", "_blok"):
            if hasattr(t, attr):
                delattr(t, attr)
    p.fo["typen"] = typen_namen
    return p
