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
  * DAEB-fees (AK fee, bijkomende kosten) worden herkend aan de labels als het FO ze in het blok van een type heeft.
De macro 'Ophalen uit FO' in het werkboek doet hetzelfde (zie vba/CashflowNaarPowerPoint_v10.bas, UitFOOphalen).
"""
import os
import re

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


def _idx(j, q):
    return j * 4 + q


def lees_fo(pad):
    wb = load_workbook(pad, data_only=True)
    p = ProjectData()
    p.params = dict(PARAM_STANDAARD)
    p.fo = {"bestand": os.path.basename(pad)}

    # ---- 1. Cashflow -> periodes ----------------------------------------------------------------------------------
    ws = _zoek_blad(wb, "1. Cashflow", "Cashflow")
    if ws is None:
        raise ValueError("Tabblad '1. Cashflow' niet gevonden in het FO.")
    kop = None
    for r in range(1, 15):
        if _txt(ws.cell(r, 2).value).lower() == "jaar":
            kop = r
            break
    if kop is None:
        raise ValueError("Koprij 'Jaar' niet gevonden op '1. Cashflow'.")
    r = kop + 1
    while _jaar(ws.cell(r, 2).value) is not None:
        q = _kw(ws.cell(r, 3).value)
        p.periodes.append({"jaar": _jaar(ws.cell(r, 2).value), "kw": f"Q{q}" if q else None,
                           "kosten": _num(ws.cell(r, 4).value), "pct_kosten": _num(ws.cell(r, 5).value),
                           "omzet": _num(ws.cell(r, 6).value), "pct_omzet": _num(ws.cell(r, 7).value),
                           "cf": _num(ws.cell(r, 8).value), "cf1000": _num(ws.cell(r, 9).value), "cf_vorig": _num(ws.cell(r, 10).value)})
        r += 1
    for rr in range(r, r + 6):          # totalen onder de tabel
        if _txt(ws.cell(rr, 3).value).lower().startswith("totaal"):
            p.fo["kosten"] = _num(ws.cell(rr, 4).value)
            p.fo["opbrengsten"] = _num(ws.cell(rr, 6).value)
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
        naam = b_naam = _txt(wo.cell(r0, 2).value)
        if " - " in b_naam:
            naam = b_naam.split(" - ", 1)[1].strip()
        else:
            naam = f"Type {k + 1}"
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
                if ll == "koopsom":
                    koop_pw, koop_tot = pw or 0.0, eur or 0.0
                    transport = _tellingen(wo, r, kaart_o, c_act_o)
                elif ll.startswith("ak fee") or ll.startswith("ak-fee") or ll.startswith("bijkomende kosten") or ll.startswith("onvoorzien"):
                    # DAEB: componentrij (totaal in €, geen kwartalen) of termijnrij (% of € in een kwartaal / actuals)
                    if not tel and not act and eur and pct is None:
                        comp = (lab, eur)
                        fees.append((comp, []))
                    elif fees:
                        kw = "actuals" if (act and not tel) else (f"{tel[0][0]} Q{tel[0][1]}" if tel else None)
                        fees[-1][1].append((lab, pct if pct is not None else eur, kw))
                elif pct is not None and pct != 0:
                    termijnen.append((lab, pct, pw or 0.0, tel))
                elif pw:
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
            if fees:
                t.soort = "DAEB"
                t.koopsom, t.grond_pct, t.termijnen, t.extras = None, None, [], []
                t.fee_comp = [(c[0], c[1]) for c, _ in fees][:LY.N_FEE_COMP]
                rijen = []
                for i, (_, termijnen_c) in enumerate(fees[:LY.N_FEE_COMP]):
                    rijen_c = [(lab, w, kw) for lab, w, kw in termijnen_c][:LY.N_FEE_PER_COMP]
                    rijen_c += [("", None, None)] * (LY.N_FEE_PER_COMP - len(rijen_c))
                    rijen += rijen_c
                t.fee_termijnen = rijen
            t._transport = transport
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
        for jj, q, v in verkocht.get(k, []):
            i = pos_act if jj is None else periode_van.get(_idx(jj, q))
            if i is not None:
                zet(p.verkocht, k, i, v)
        for jj, q, v in getattr(t, "_transport", []):
            i = pos_act if jj is None else periode_van.get(_idx(jj, q))
            if i is not None:
                zet(p.transport, k, i, v)
        if hasattr(t, "_transport"):
            del t._transport
    p.fo["typen"] = typen_namen
    return p
