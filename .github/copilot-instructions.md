# Instructies voor Copilot

Lees eerst `docs/Overdracht_Copilot.md`: daarin staan de opbouw van het RDFO-bestand, de vaste rekenafspraken, de open punten en wat er in de scenarioanalyse (bandbreedte) gebouwd is.

Kort:
- `RDFO_prototype.xlsm` en `RDFO_template.xlsm` worden gegenereerd. Wijzig `tools/build_rdfo.py` en `macros/*.bas`, niet het .xlsm-bestand zelf.
- Bouwen en testen: `python tools/release.py <pad-naar-recalc.py>`.
- Nooit `ABS` op bedragen, ontbrekend = `n.b.` (nooit 0), bedragen en woningen zijn al op aandeelbasis.
- Communiceer in het Nederlands, kort en concreet.
