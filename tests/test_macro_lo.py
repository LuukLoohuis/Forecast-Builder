"""Einde-tot-einde: de macro 'Ophalen uit FO' draait in LibreOffice op het nagebouwde FO en zet hetzelfde neer als fo.lees_fo.
Wordt overgeslagen zonder pyuno en soffice."""
import os
import sys
import tempfile

import pytest

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))
sys.path.insert(0, HIER)

from forecast_builder import fo, data as D  # noqa: E402
import macro_lo  # noqa: E402
from test_fo import maak_fo, maak_fo_vorig  # noqa: E402


@pytest.mark.skipif(not macro_lo.beschikbaar(), reason="pyuno of soffice ontbreekt")
def test_macro_uit_fo_in_libreoffice():
    with tempfile.TemporaryDirectory() as d:
        fo_pad = os.path.join(d, "fo.xlsx")
        maak_fo(fo_pad)
        uit = os.path.join(d, "uit.xlsx")
        melding = macro_lo.voer_uit(fo_pad, uit, werkmap=d)
        assert melding.startswith("KLAAR:"), melding
        assert "8 periodes" in melding and "2 woningtype(n)" in melding
        p_macro = D.lees_werkboek(uit)
        verschillen = macro_lo.vergelijk(fo.lees_fo(fo_pad), p_macro)
    assert verschillen == [], "\n".join(verschillen)
    # de planning van vóór het ophalen (voorbeeld: Rijwoning start 2027 Q1) is de vorige prognose geworden (blok 4)
    assert (p_macro.types[0].vorig_start_jaar, p_macro.types[0].vorig_start_kw) == (2027, 1)


@pytest.mark.skipif(not macro_lo.beschikbaar(), reason="pyuno of soffice ontbreekt")
def test_vorige_prognose_uit_fo_in_libreoffice():
    """Knop 'Vorige prognose uit FO': na het ophalen uit het FO de vorige prognose uit een ouder FO (start bouw naar blok 4,
    fee-kwartalen op mijlpaal naar blok 6); zelfde uitkomst als fo.voeg_vorige_prognose_toe."""
    with tempfile.TemporaryDirectory() as d:
        fo_pad, pad_vorig = os.path.join(d, "fo.xlsx"), os.path.join(d, "fo_vorig.xlsx")
        maak_fo(fo_pad)
        maak_fo_vorig(pad_vorig)
        uit = os.path.join(d, "uit.xlsx")
        meldingen = macro_lo.voer_uit(fo_pad, uit, werkmap=d, stappen=[("UitFOOphalen", fo_pad), ("VorigePrognoseUitFO", pad_vorig)])
        assert all(m.startswith("KLAAR:") for m in meldingen), meldingen
        p_macro = D.lees_werkboek(uit)
        p = fo.lees_fo(fo_pad)
        fo.voeg_vorige_prognose_toe(p, pad_vorig)
    for t, tm in zip(p.types, p_macro.types):
        assert (t.vorig_start_jaar, t.vorig_start_kw) == (tm.vorig_start_jaar, tm.vorig_start_kw), t.naam
        assert [kw for kw in t.fee_vorig] == [kw for kw in tm.fee_vorig] or (not any(t.fee_vorig) and not any(tm.fee_vorig)), (t.naam, t.fee_vorig, tm.fee_vorig)
