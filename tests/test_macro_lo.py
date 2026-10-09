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
from test_fo import maak_fo  # noqa: E402


@pytest.mark.skipif(not macro_lo.beschikbaar(), reason="pyuno of soffice ontbreekt")
def test_macro_uit_fo_in_libreoffice():
    with tempfile.TemporaryDirectory() as d:
        fo_pad = os.path.join(d, "fo.xlsx")
        maak_fo(fo_pad)
        uit = os.path.join(d, "uit.xlsx")
        melding = macro_lo.voer_uit(fo_pad, uit, werkmap=d)
        assert melding.startswith("KLAAR:"), melding
        assert "8 periodes" in melding and "2 woningtype(n)" in melding
        verschillen = macro_lo.vergelijk(fo.lees_fo(fo_pad), D.lees_werkboek(uit))
    assert verschillen == [], "\n".join(verschillen)
