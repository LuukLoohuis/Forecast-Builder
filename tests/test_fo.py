"""FO-lezer (forecast_builder/fo.py): een klein nagebouwd FO-werkboek met dezelfde labels als het echte FO (zonder LibreOffice)."""
import os
import sys
import tempfile

from openpyxl import Workbook

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))

from forecast_builder import fo  # noqa: E402

KW = ["Q1", "Q2", "Q3", "Q4"]


def _kopregels(ws, r, jaar0=2026, n_jaar=3, extra=None):
    """rij r-1: jaartallen (elke vier kolommen), rij r: 'actuals' in J en Q1..Q4 vanaf K; geeft {(jaar, kw): kolom}."""
    ws.cell(r, 10).value = "actuals"
    kaart = {}
    c = 11
    for j in range(n_jaar):
        ws.cell(r - 1, c).value = jaar0 + j
        for q in range(4):
            ws.cell(r, c).value = KW[q]
            kaart[(jaar0 + j, q + 1)] = c
            c += 1
    return kaart


def maak_fo(pad):
    wb = Workbook()
    ws = wb.active
    ws.title = "1. Cashflow"
    ws["B3"], ws["C3"] = "Jaar", "Q"
    rijen = [(2025, None, 100.0, 0.0), (2026, "Q1", 50.0, 0.0), (2026, "Q2", 60.0, 0.0), (2026, "Q3", 70.0, 0.0),
             (2026, "Q4", 500.0, 900.0), (2027, "Q1", 400.0, 800.0), (2027, "Q2", 300.0, 700.0), (2027, "Q3", 200.0, 600.0)]
    for i, (j, q, k, o) in enumerate(rijen):
        r = 4 + i
        ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 6).value = j, q, k, o
        ws.cell(r, 8).value = o - k
    ws.cell(13, 3).value = "Totaal:"
    ws.cell(13, 4).value = sum(k for _, _, k, _ in rijen)
    ws.cell(13, 6).value = sum(o for _, _, _, o in rijen)
    wa = wb.create_sheet("FO - actuals")
    for c, kop in enumerate(["Waarde/TrVal", "Boekjaar", "KOSTEN/OMZET", "Q"], start=11):
        wa.cell(1, c).value = kop
    for i, (j, q) in enumerate([(2026, "Q1"), (2026, "Q2"), (2026, "Q3"), (2025, "Q4")], start=2):
        wa.cell(i, 11).value, wa.cell(i, 12).value, wa.cell(i, 14).value = 10, str(j), q
    wo = wb.create_sheet("CF - opbrengsten")
    # type 1 (koop): bouwplanning
    wo["F4"], wo["G4"] = "# won", 10
    wo["B5"] = "Termijnen 1 - Rijwoning"
    k1 = _kopregels(wo, 5)
    wo["B6"] = "Start bouw"; wo.cell(6, k1[(2027, 1)]).value = 10
    wo["B7"] = "Casco"; wo.cell(7, k1[(2027, 2)]).value = 10
    wo["B8"] = "Oplevering"; wo.cell(8, k1[(2027, 3)]).value = 4; wo.cell(8, k1[(2027, 4)]).value = 6
    wo["B9"] = "Kopersmeerwerk"; wo.cell(9, k1[(2027, 2)]).value = 10
    # type 2 (geen woningen) en type 3 (DAEB)
    wo["F12"], wo["G12"] = "# won", 0
    wo["B13"] = "Termijnen 2 - Leeg"
    _kopregels(wo, 13)
    wo["B14"] = "Start bouw"
    wo["F17"], wo["G17"] = "# won", 20
    wo["B18"] = "Termijnen 3 - DAEB"
    k3 = _kopregels(wo, 18)
    wo["B19"] = "Start bouw"; wo.cell(19, k3[(2027, 2)]).value = 20
    # verkooptempo
    wo["B22"] = "Verkooptempo"
    kv = _kopregels(wo, 22)
    wo["B23"] = "Rijwoning"; wo.cell(23, 10).value = 3; wo.cell(23, kv[(2026, 4)]).value = 7
    wo["B24"] = "Leeg"
    wo["B25"] = "DAEB"; wo.cell(25, kv[(2027, 1)]).value = 20
    # omzetblok type 1
    wo["B28"] = "Omzet HVG Termijnen"
    ko = _kopregels(wo, 28)
    wo["B29"], wo["D29"], wo["E29"] = "Koopsom", 1000000, 100000
    wo.cell(29, ko[(2026, 4)]).value = 7; wo.cell(29, ko[(2027, 1)]).value = 3
    wo["B30"], wo["C30"], wo["D30"], wo["E30"] = "Start bouw", 0.30, 600000, 60000
    wo["B31"], wo["C31"], wo["D31"], wo["E31"] = "Casco", 0.50, 1000000, 100000
    wo["B32"], wo["C32"], wo["D32"], wo["E32"] = "Oplevering", 0.20, 400000, 40000
    wo["B33"], wo["D33"], wo["E33"] = "Kopersmeerwerk", 50000, 5000
    wo["B34"], wo["D34"], wo["E34"] = "Kadastrale kosten", 10000, 1000
    wo.cell(34, 10).value = 10
    # omzetblok type 2 (leeg) en type 3 (DAEB met fees)
    wo["B37"] = "Omzet HVG Termijnen"
    _kopregels(wo, 37)
    wo["B38"] = "Koopsom"
    wo["B41"] = "Omzet HVG Termijnen"
    kd = _kopregels(wo, 41)
    # zoals het fee-overzicht van de gebruiker: componentrij met alleen een totaal (hier met een 0 in de %-kolom), daaronder
    # termijnrijen zonder prefix; 'Onvoorzien' (met termijn) doet niet mee; een fee-rij zonder bedrag geeft een melding
    wo["B42"], wo["C42"], wo["D42"] = "AK fee", 0, 1000000
    wo["B43"], wo["C43"] = "na akkoord SO", 0.30; wo.cell(43, 10).value = 300000
    wo["B44"], wo["C44"] = "na omgevingsvergunning", 0.70; wo.cell(44, kd[(2026, 4)]).value = 700000
    wo["B45"], wo["D45"] = "Bijkomende kosten (excl. AK)", 500000
    wo["B46"], wo["D46"] = "na tekenen TKO", 500000; wo.cell(46, kd[(2027, 2)]).value = 500000
    wo["B47"], wo["D47"] = "Onvoorzien", 50000
    wo["B48"], wo["C48"] = "bij start bouw", 1.0; wo.cell(48, kd[(2027, 2)]).value = 50000
    wo["B49"] = "AK fee (toelichting)"
    # CF - kosten: per type een termijnblok (wordt niet gelezen) en het blok 'Onderdelen': 'Actuals' samengevoegd boven H:J, de
    # jaartallen boven de kwartalen (K = 2026 Q1 .. R = 2027 Q4); rijen tot de eerste lege
    wk = wb.create_sheet("CF - kosten")
    # termijnblokken per type (zelfde titel als op 'CF - opbrengsten'): bedrag per woning in E, woningen per kwartaal; rijen tot 'Aanneemsom'
    wk["F3"], wk["G3"] = "# won", 10
    wk["B4"] = "Termijnen 1 - Rijwoning"
    kk = _kopregels(wk, 4)
    wk["B5"], wk["C5"], wk["E5"] = "Start bouw", 0.2, 20000
    wk.cell(5, kk[(2027, 1)]).value = 10
    wk["B6"], wk["C6"], wk["E6"] = "Casco", 0.5, 50000
    wk.cell(6, kk[(2027, 2)]).value, wk.cell(6, kk[(2027, 3)]).value = 6, 4
    wk["B7"], wk["C7"], wk["E7"] = "Oplevering", 0.3, 30000
    wk.cell(7, kk[(2027, 4)]).value = 10
    wk["B8"], wk["C8"] = "Kopersmeerwerk 25%", 0.25                   # geen bedrag per woning: telt niet mee
    wk.cell(8, kk[(2027, 2)]).value = 10
    wk["B9"], wk["B10"] = "Aanneemsom", "Meerwerk"
    wk["F12"], wk["G12"] = "# won", 20
    wk["B13"] = "Termijnen 3 - DAEB"
    kd = _kopregels(wk, 13)
    wk["B14"], wk["E14"] = "Bouw DAEB", 100000
    wk.cell(14, 10).value = 5                                          # al gebouwd (kolom actuals)
    wk.cell(14, kd[(2027, 2)]).value = 15
    wk["B15"] = "Aanneemsom"
    wk["B49"], wk["D49"] = "Totaal kosten", 9999
    wk["H51"] = "Actuals"
    wk.merge_cells("H51:J51")
    wk["K51"], wk["O51"] = 2026, 2027
    wk["B52"], wk["D52"] = "Onderdelen", "€"
    for i, c in enumerate("KLMNOPQR"):
        wk[f"{c}52"] = KW[i % 4]
    wk["B53"], wk["D53"], wk["J53"] = "Grondkosten", 1000, 1000
    wk["B54"], wk["M54"], wk["N54"], wk["O54"], wk["R54"] = "Aanneemsom", 50, 200, 300, 100   # Q3 '26 = actuals, Q4 '27 na de laatste periode
    wk["B55"], wk["I55"], wk["N55"], wk["O55"] = "Promotiekosten", 5, 10, 10
    wk["B56"], wk["H56"], wk["I56"], wk["P56"] = "Honoraria", 20, 30, 40
    wk["E56"] = 0.05                                                  # percentage in een tussenkolom telt niet mee
    wk["B58"], wk["D58"] = "Check:", 0
    wb.save(pad)


def test_lees_fo():
    with tempfile.TemporaryDirectory() as d:
        pad = os.path.join(d, "fo.xlsx")
        maak_fo(pad)
        p = fo.lees_fo(pad)
    assert len(p.periodes) == 8 and p.periodes[0]["kw"] is None and p.periodes[1]["kw"] == "Q1"
    assert p.fo["opbrengsten"] == 3000.0 and p.fo["kosten"] == 1680.0
    assert (p.params["actuals_jaar"], p.params["actuals_kw"]) == (2026, 3)
    assert [t.naam for t in p.types] == ["Rijwoning", "DAEB"]          # blok zonder woningen overgeslagen
    t = p.types[0]
    assert t.aantal == 10 and t.koopsom == 300000 and abs(t.grond_pct - 1 / 3) < 1e-6 and (t.start_jaar, t.start_kw) == (2027, 1)
    assert t.termijnen == [("Start bouw", 0.3, 1), ("Casco", 0.5, 2), ("Oplevering (1/2)", 0.08, 3), ("Oplevering (2/2)", 0.12, 4)]
    assert t.extras == [("Kopersmeerwerk", 5000.0, 2), ("Kadastrale kosten", 1000.0, None)]
    daeb = p.types[1]
    assert daeb.soort == "DAEB" and daeb.koopsom is None and daeb.termijnen == []
    assert daeb.fee_comp == [("AK fee", 1000000), ("Bijkomende kosten (excl. AK)", 500000)]
    assert daeb.fee_termijnen[0] == ("na akkoord SO", 300000, "actuals")            # bedragen uit de cellen (€), niet het %
    assert daeb.fee_termijnen[1] == ("na omgevingsvergunning", 700000, "2026 Q4")
    assert daeb.fee_termijnen[2] == ("", None, None)
    assert daeb.fee_termijnen[5] == ("na tekenen TKO", 500000, "2027 Q2") and daeb.fee_termijnen[6] == ("", None, None)
    assert daeb.extras == [] and len(daeb.fee_termijnen) == 10
    assert p.fo["bestand"] == "fo.xlsx" and p.fo["datum"] and p.fo["actuals"] == "Q3 2026"
    assert [w for w in p.fo["waarschuwingen"] if "AK fee (toelichting)" in w and "overgeslagen" in w]
    assert not [w for w in p.fo["waarschuwingen"] if "FO - actuals" in w]
    # verkocht: actuals (3) in het actuals-kwartaal Q3 '26 (periode-index 3), 7 in Q4 '26; DAEB 20 in Q1 '27 (tweede kolom)
    assert p.verkocht[3] == [3, None] and p.verkocht[4] == [7, None] and p.verkocht[5] == [None, 20]
    # transport van de rij Koopsom: 7 in Q4 '26, 3 in Q1 '27; DAEB heeft geen transport in het FO: transport = verkoop
    assert p.transport[4] == [7, None] and p.transport[5] == [3, 20]
    # CF - kosten (blok Onderdelen): gerealiseerd = kolommen onder 'Actuals' + kwartalen t/m actuals; na de laatste periode in de laatste
    kp = {k["naam"]: k for k in p.kosten_posten}
    assert list(kp) == ["Grondkosten", "Aanneemsom", "Promotiekosten", "Honoraria"]
    assert [kp[n]["mee"] for n in kp] == ["vast", "start bouw", "verkoop", "vast"]
    assert [kp[n]["gerealiseerd"] for n in kp] == [1000, 50, 5, 50]
    assert kp["Grondkosten"]["bedragen"] == [None] * 8
    assert kp["Aanneemsom"]["bedragen"] == [None, None, None, None, 200, 300, None, 100]
    assert kp["Promotiekosten"]["bedragen"] == [None, None, None, None, 10, 10, None, None]
    assert kp["Honoraria"]["bedragen"] == [None, None, None, None, None, None, 40, None]
    assert [w for w in p.fo["waarschuwingen"] if "CF - kosten" in w and "Aanneemsom" in w and "laatste periode" in w]
    # bouwkosten per termijn (blok 7): bedrag = per woning x woningen per kwartaal, bouwkwartaal t.o.v. start bouw van het type
    assert p.types[0].bouwkosten == [("Start bouw", 200000, 1), ("Casco (1/2)", 300000, 2), ("Casco (2/2)", 200000, 3), ("Oplevering", 300000, 4)]
    assert p.types[1].bouwkosten == [("Bouw DAEB (gerealiseerd)", 500000, -2), ("Bouw DAEB", 1500000, 1)]   # actuals in het actuals-kwartaal


def test_binnen_limiet():
    rijen = [("A (1/3)", 0.1, 1), ("A (2/3)", 0.05, 2), ("A (3/3)", 0.15, 3)] + [(f"T{i}", 0.07, i + 4) for i in range(9)]
    meldingen = []
    uit = fo._binnen_limiet(rijen, 10, "Rijwoning", "bouwtermijnen", meldingen)
    assert len(uit) == 10 and uit[0] == ("A", 0.3, 3) and uit[1:] == rijen[3:]       # kleinste delen bij het grootste deel
    assert abs(sum(w for _, w, _ in uit) - sum(w for _, w, _ in rijen)) < 1e-9 and len(meldingen) == 1 and "samengevoegd" in meldingen[0]
    assert fo._binnen_limiet(rijen[3:], 10, "x", "bouwtermijnen", []) == rijen[3:]
    try:
        fo._binnen_limiet([(f"T{i}", 0.05, i + 1) for i in range(11)], 10, "Rijwoning", "bouwtermijnen", [])
        assert False, "ValueError verwacht"
    except ValueError as e:
        assert "11 bouwtermijnen" in str(e)
    # afkappen=True (bouwkosten, zoals de macro): eerst samenvoegen, dan de laatste rijen weg, met beide meldingen
    meldingen = []
    uit = fo._binnen_limiet(rijen + [("T9", 0.01, 13)], 10, "Rijwoning", "bouwkosten-termijnen", meldingen, afkappen=True)
    assert uit == [("A", 0.3, 3)] + rijen[3:12]
    assert len(meldingen) == 2 and "niet overgenomen" in meldingen[0] and "samengevoegd" in meldingen[1]


def maak_fo_bouwkosten(pad):
    """maak_fo met lastige termijnblokken op 'CF - kosten': Rijwoning en DAEB hebben dezelfde bloktitel ('Termijnen <300K', naam in
    kolom C); Rijwoning heeft 14 regels (Casco over twee kwartalen plus T1..T10), DAEB een gerealiseerd deel en elf latere delen."""
    from openpyxl import load_workbook
    maak_fo(pad)
    wb = load_workbook(pad)
    wo, wk = wb["CF - opbrengsten"], wb["CF - kosten"]
    for r, naam in ((5, "Rijwoning"), (18, "DAEB")):
        wo.cell(r, 2).value, wo.cell(r, 3).value = "Termijnen <300K", naam
    for r in range(3, 49):
        for c in range(2, 27):
            wk.cell(r, c).value = None
    wk["F3"], wk["G3"] = "# won", 10
    wk["B4"], wk["C4"] = "Termijnen <300K", "Rijwoning"
    kk = _kopregels(wk, 4)
    wk["B5"], wk["E5"] = "Start bouw", 20000
    wk.cell(5, kk[(2027, 1)]).value = 10
    wk["B6"], wk["E6"] = "Casco", 50000
    wk.cell(6, kk[(2027, 2)]).value, wk.cell(6, kk[(2027, 3)]).value = 6, 4
    wk["B7"], wk["E7"] = "Oplevering", 30000
    wk.cell(7, kk[(2027, 4)]).value = 10
    for i in range(1, 11):
        wk.cell(7 + i, 2).value, wk.cell(7 + i, 5).value = f"T{i}", 1000 * i
        wk.cell(7 + i, kk[(2028, 1)]).value = 10
    wk["B18"] = "Aanneemsom"
    wk["F21"], wk["G21"] = "# won", 20
    wk["B22"], wk["C22"] = "Termijnen <300K", "DAEB"
    kd = _kopregels(wk, 22)
    wk["B23"], wk["E23"] = "Bouw DAEB", 100000
    wk.cell(23, 10).value = 5                                          # kolom actuals
    wk.cell(23, kd[(2026, 3)]).value = 1                               # t/m het actuals-kwartaal: ook gerealiseerd
    for j, q in [(2026, 4)] + [(2027, q) for q in range(1, 5)] + [(2028, q) for q in range(1, 5)]:
        wk.cell(23, kd[(j, q)]).value = 1
    wk["B24"], wk["E24"] = "Afbouw DAEB", 10000
    wk.cell(24, kd[(2027, 2)]).value, wk.cell(24, kd[(2027, 3)]).value = 10, 10
    wk["B25"] = "Aanneemsom"
    wb.save(pad)


def test_bouwkosten_randgevallen():
    """Dubbele bloktitels: het n-de type met die titel krijgt het n-de termijnblok. Meer dan tien regels: eerst samenvoegen, dan
    afkappen (zoals de macro). Het gerealiseerde deel ('(gerealiseerd)') doet niet mee in het samenvoegen, zodat er geen prognose
    in de realisatie belandt."""
    with tempfile.TemporaryDirectory() as d:
        pad = os.path.join(d, "fo.xlsx")
        maak_fo_bouwkosten(pad)
        p = fo.lees_fo(pad)
    rij, daeb = p.types[0], p.types[1]
    assert (rij.naam, daeb.naam) == ("Rijwoning", "DAEB")
    assert rij.bouwkosten == ([("Start bouw", 200000, 1), ("Casco", 500000, 2), ("Oplevering", 300000, 4)]
                              + [(f"T{i}", 10000 * i, 5) for i in range(1, 8)])
    w = p.fo["waarschuwingen"]
    assert [x for x in w if "Rijwoning: 13 bouwkosten-termijnen" in x] and [x for x in w if "Rijwoning" in x and "Casco" in x]
    assert not [x for x in w if "geen termijnblok" in x]
    assert len(daeb.bouwkosten) == 10 and daeb.bouwkosten[0] == ("Bouw DAEB (gerealiseerd)", 600000, -2)
    assert sum(b for _, b, _ in daeb.bouwkosten) == 600000 + 9 * 100000 + 200000
    assert all(kw >= 0 for _, _, kw in daeb.bouwkosten[1:])                    # na actuals (Q3 '26): kw -1 en later


def maak_fo_daeb(pad, vorig=False):
    """maak_fo, maar DAEB als los blok 'Opbrengsten DAEB' onderaan 'CF - opbrengsten' (zoals in het echte FO): boven de koprij per
    component de naam (C) en het totaal (E) en 'Startbouw' boven een kwartaal; daaronder de termijnen met het component-voorvoegsel,
    in twee groepen. Geen termijnblok en geen omzetblok voor DAEB; aantal en verkoop op tabblad 'Verkooptempo' (rij 'Corporatie').
    vorig=True: een ouder FO (start bouw en twee termijnen een kwartaal eerder)."""
    from openpyxl import load_workbook
    maak_fo(pad)
    wb = load_workbook(pad)
    wo = wb["CF - opbrengsten"]
    for r in list(range(17, 20)) + [25] + list(range(40, 50)):        # DAEB-type uit maak_fo weg (termijnblok, verkooptempo, omzetblok)
        for c in range(2, 30):
            wo.cell(r, c).value = None
    kq = _kopregels(wo, 56)
    wo["C54"], wo["E54"] = "AK", 950000
    wo["C55"], wo["E55"] = "Bijkomende kosten ex. AK", 500000
    wo.cell(54, kq[(2027, 1) if vorig else (2027, 2)]).value = "Startbouw"
    wo["B56"] = "Opbrengsten DAEB"
    wo["B57"], wo["C57"], wo["J57"] = "AK fee na akkoord SO", 0.3, 250000             # alleen actuals (ander bedrag dan 30%)
    wo["B58"], wo["C58"], wo["J58"] = "AK fee na akkoord VO", 0.3, 150000             # deels ontvangen, de rest in Q4 '26
    wo.cell(58, kq[(2026, 4)]).value = 150000
    wo["B59"], wo["C59"] = "AK fee na akkoord DO", 0.4                                 # over twee kwartalen
    for q in ((2026, 4), (2027, 1)) if vorig else ((2027, 1), (2027, 2)):
        wo.cell(59, kq[q]).value = 200000
    wo["B61"], wo["C61"] = "Bijkomende kosten na tekenen TKO", 0.5
    wo.cell(61, kq[(2027, 1) if vorig else (2027, 2)]).value = 250000
    wo["B62"], wo["C62"] = "Gevelsluiting", 0.5                                         # zonder voorvoegsel: component van de groep
    wo.cell(62, kq[(2027, 3)]).value = 250000
    wo["B66"] = "Totaal opbrengsten"
    wv = wb.create_sheet("Verkooptempo")
    kv = _kopregels(wv, 3)
    wv["A3"], wv["B3"] = "Verkooptempo", "Aantal"
    wv["A4"], wv["B4"] = "Rijwoning", 10
    wv.cell(4, kv[(2026, 4)]).value = 10
    wv["A5"], wv["B5"] = "Corporatie", 20
    wv.cell(5, kv[(2026, 4)]).value = 20
    wb.save(pad)


def maak_fo_daeb_lastig(pad, vorig=False):
    """maak_fo_daeb met de randgevallen uit de review: componenten 'Fee AK' / 'Fee bijkomende kosten' / 'Onvoorzien' (drie rijen
    boven de koprij, 'Startbouw' in de bovenste), een termijn met ontvangen deel en later deel, een negatieve correctie, een deel van
    40 cent, een scheidingsteken na het voorvoegsel, subtotalen, een groep 'Onvoorzien'; op 'Verkooptempo' een titel in A1, een
    koopnaam die net afwijkt, een lege rij en een totaalrij. vorig=True: een ouder FO (DO in één kwartaal, TKO een kwartaal eerder)."""
    from openpyxl import load_workbook
    maak_fo_daeb(pad)
    wb = load_workbook(pad)
    wo = wb["CF - opbrengsten"]
    kq = _kopregels(wo, 56)
    for r in list(range(53, 56)) + list(range(57, 75)):
        for c in range(2, 30):
            wo.cell(r, c).value = None
    _kopregels(wo, 56)
    wo["C53"], wo["E53"] = "Fee AK", 1000000
    wo["C54"], wo["E54"] = "Fee bijkomende kosten", 500000
    wo["C55"], wo["E55"] = "Onvoorzien", 50000
    wo.cell(53, kq[(2027, 2)]).value = "Startbouw"
    wo["B57"], wo["J57"] = "Fee AK na akkoord SO", 250000
    wo.cell(57, kq[(2026, 4)]).value = 10000
    wo["B58"], wo["J58"] = "Fee AK - na akkoord VO", 150000
    wo.cell(58, kq[(2026, 4)]).value = 150000
    wo.cell(58, kq[(2027, 1)]).value = -10000                             # correctie
    wo["B59"] = "Fee AK na akkoord DO"
    for q in ((2027, 1),) if vorig else ((2027, 1), (2027, 2)):
        wo.cell(59, kq[q]).value = 400000 if vorig else 200000
    wo["B60"], wo["J60"] = "Fee AK na omgevingsvergunning", 0.4           # 40 cent: afronding
    wo.cell(60, kq[(2027, 3)]).value = 50000
    wo["B61"], wo["J61"] = "Subtotaal AK", 400005
    wo["B63"] = "Fee bijkomende kosten na tekenen TKO"
    wo.cell(63, kq[(2027, 1) if vorig else (2027, 2)]).value = 250000
    wo["B64"] = "Gevelsluiting"
    wo.cell(64, kq[(2027, 3)]).value = 250000
    wo["B65"], wo["J65"] = "Totaal fee bijkomende kosten", 500000
    wo["B67"] = "Onvoorzien bij start bouw"
    wo.cell(67, kq[(2027, 2)]).value = 50000
    wo["B71"] = "Totaal opbrengsten"
    wv = wb["Verkooptempo"]
    wv["A1"] = "Verkooptempo"
    wv["A4"] = "Rijwoningen"
    wv.move_range("A5:Z5", rows=1)
    wv["A7"], wv["B7"] = "Totaal", 30
    wb.save(pad)


def test_daeb_los_blok_randgevallen():
    """Zie maak_fo_daeb_lastig: sleutelwoord zonder 'fee', geen bedrag kwijt boven vijf delen (eerst al het ontvangene in één rij, dan
    hetzelfde kwartaal samen), correcties bij het vorige deel, subtotalen en 'Onvoorzien' overgeslagen, 'Verkooptempo' met titel,
    afwijkende koopnaam, lege rij en totaal; vorige prognose per component op mijlpaal."""
    with tempfile.TemporaryDirectory() as d:
        pad, pad_vorig = os.path.join(d, "fo.xlsx"), os.path.join(d, "fo_vorig.xlsx")
        maak_fo_daeb_lastig(pad)
        maak_fo_daeb_lastig(pad_vorig, vorig=True)
        p = fo.lees_fo(pad)
        fo.voeg_vorige_prognose_toe(p, pad_vorig)
    t = p.types[1]
    assert (t.naam, t.aantal, t.start_jaar, t.start_kw) == ("DAEB", 20, 2027, 2)
    assert t.fee_comp == [("Fee AK", 1000000), ("Fee bijkomende kosten", 500000)]
    assert t.fee_termijnen == [("ontvangen t/m actuals", 400000, "actuals"), ("na akkoord SO + na akkoord VO", 150000, "2026 Q4"),
                               ("na akkoord DO (1/2)", 200000, "2027 Q1"), ("na akkoord DO (2/2)", 200000, "2027 Q2"),
                               ("na omgevingsvergunning", 50000.4, "2027 Q3"),
                               ("na tekenen TKO", 250000, "2027 Q2"), ("Gevelsluiting", 250000, "2027 Q3"),
                               ("", None, None), ("", None, None), ("", None, None)]
    w = p.fo["waarschuwingen"]
    assert [x for x in w if "'Corporatie' is gebruikt" in x] and [x for x in w if "correctie of afronding" in x]
    assert t.fee_vorig[:7] == ["actuals", "2026 Q4", "2027 Q1", "2027 Q1", "2027 Q3", "2027 Q1", "2027 Q3"]


def test_daeb_los_blok():
    """Los DAEB-blok: een eigen DAEB-type met de componenten van boven de koprij, de termijnen per component (op voorvoegsel, anders
    per groep; voorvoegsel weg uit de mijlpaal), bedragen uit de cellen per kwartaal gesplitst ('actuals' = ontvangen), start bouw
    van 'Startbouw', aantal en verkoop van tabblad 'Verkooptempo' en transport = verkoop. Vorige prognose op mijlpaal."""
    with tempfile.TemporaryDirectory() as d:
        pad, pad_vorig = os.path.join(d, "fo.xlsx"), os.path.join(d, "fo_vorig.xlsx")
        maak_fo_daeb(pad)
        maak_fo_daeb(pad_vorig, vorig=True)
        p = fo.lees_fo(pad)
        fo.voeg_vorige_prognose_toe(p, pad_vorig)
    assert [t.naam for t in p.types] == ["Rijwoning", "DAEB"]
    t = p.types[1]
    assert (t.soort, t.aantal, t.start_jaar, t.start_kw, t.koopsom, t.termijnen) == ("DAEB", 20, 2027, 2, None, [])
    assert t.fee_comp == [("AK", 950000), ("Bijkomende kosten ex. AK", 500000)]
    assert t.fee_termijnen == [("na akkoord SO", 250000, "actuals"), ("na akkoord VO (ontvangen)", 150000, "actuals"),
                               ("na akkoord VO", 150000, "2026 Q4"), ("na akkoord DO (1/2)", 200000, "2027 Q1"),
                               ("na akkoord DO (2/2)", 200000, "2027 Q2"),
                               ("na tekenen TKO", 250000, "2027 Q2"), ("Gevelsluiting", 250000, "2027 Q3"),
                               ("", None, None), ("", None, None), ("", None, None)]
    assert p.verkocht[4] == [7, 20] and p.transport[4] == [7, 20]            # Q4 '26: DAEB verkocht en (= verkoop) getransporteerd
    assert sum(r[1] or 0 for r in p.verkocht) == 20 and sum(r[1] or 0 for r in p.transport) == 20
    assert [w for w in p.fo["waarschuwingen"] if "20 woningen en de verkoop uit 'Verkooptempo' (rij 'Corporatie')" in w]
    assert not [w for w in p.fo["waarschuwingen"] if "termijnblok" in w and "DAEB" in w]
    # vorige prognose: start bouw en fee-kwartalen uit het oudere FO (op mijlpaal)
    assert (t.vorig_start_jaar, t.vorig_start_kw) == (2027, 1)
    assert t.fee_vorig[:7] == ["actuals", "actuals", "2026 Q4", "2026 Q4", "2027 Q1", "2027 Q1", "2027 Q3"]


def maak_fo_vorig(pad_vorig):
    """Ouder FO (vorige prognose): Rijwoning startte een kwartaal eerder (2026 Q4), 'na omgevingsvergunning' stond een kwartaal
    eerder (2026 Q3) en 'na tekenen TKO' twee kwartalen eerder (2026 Q4)."""
    from openpyxl import load_workbook
    maak_fo(pad_vorig)
    wb = load_workbook(pad_vorig)
    wo = wb["CF - opbrengsten"]
    k1 = _kopregels(wo, 5)
    wo.cell(6, k1[(2027, 1)]).value = None
    wo.cell(6, k1[(2026, 4)]).value = 10
    kd = _kopregels(wo, 41)
    wo.cell(44, kd[(2026, 4)]).value = None
    wo.cell(44, kd[(2026, 3)]).value = 700000
    wo.cell(46, kd[(2027, 2)]).value = None
    wo.cell(46, kd[(2026, 4)]).value = 500000
    wb.save(pad_vorig)


def test_vorige_prognose_uit_ouder_fo():
    """--fo-vorig: start bouw en fee-kwartalen uit een ouder FO naar blok 4 en 6 (typen op naam, fee-termijnen op mijlpaal)."""
    with tempfile.TemporaryDirectory() as d:
        pad, pad_vorig = os.path.join(d, "fo.xlsx"), os.path.join(d, "fo_vorig.xlsx")
        maak_fo(pad)
        maak_fo_vorig(pad_vorig)
        p = fo.lees_fo(pad)
        fo.voeg_vorige_prognose_toe(p, pad_vorig)
    assert (p.types[0].vorig_start_jaar, p.types[0].vorig_start_kw) == (2026, 4)
    assert (p.types[1].vorig_start_jaar, p.types[1].vorig_start_kw) == (2027, 2)
    assert p.types[1].fee_vorig[:2] == ["actuals", "2026 Q3"] and p.types[1].fee_vorig[5] == "2026 Q4" and p.types[1].fee_vorig[2] is None
    assert p.fo["vorig_bestand"] == "fo_vorig.xlsx"
    assert [w for w in p.fo["waarschuwingen"] if "toelichting" not in w and "CF - kosten" not in w] == []


def test_kosten_randgevallen():
    """'CF - kosten': een blok 'Onderdelen' in rij 1 (zonder jaarrij erboven) geeft een melding en geen crash; zonder 'FO - actuals'
    bepaalt de terugval-actuals (build.py --fo --from: die van het werkboek) de grens gerealiseerd/prognose."""
    with tempfile.TemporaryDirectory() as d:
        pad = os.path.join(d, "fo.xlsx")
        maak_fo(pad)
        from openpyxl import load_workbook
        wb = load_workbook(pad)
        del wb["FO - actuals"]
        wk = wb["CF - kosten"]
        wb.save(pad)
        p = fo.lees_fo(pad, actuals=(2026, 4))
        assert (p.params["actuals_jaar"], p.params["actuals_kw"]) == (2026, 4)
        kp = {k["naam"]: k for k in p.kosten_posten}
        assert kp["Aanneemsom"]["gerealiseerd"] == 250 and kp["Aanneemsom"]["bedragen"][4] is None   # Q4 '26 nu gerealiseerd
        wk.delete_rows(1, 51)                                          # 'Onderdelen' komt in rij 1
        wb.save(pad)
        p = fo.lees_fo(pad)
    assert p.kosten_posten == [] and [w for w in p.fo["waarschuwingen"] if "Onderdelen" in w]
