# Forecast-Builder

RDFO-prognoserapportage in Excel (met macro's).

- `RDFO_template.xlsm` — het bestand om te gebruiken (afgeslankt).
- `bron/RDFO_template_bron.xlsm` — hetzelfde bestand vóór het afslanken; hierop werken de scripts.
- `tools/` — scripts die het bestand stap voor stap hebben opgebouwd. Elk script draait één keer, in deze volgorde:
  1. `huisstijl.py` — Heijmans-kleuren
  2. `bouw_ak_vergelijk.py` — over- of onderdekking op AK-dekking en het tabblad Vergelijk
  3. `bouw_powerpoint.py` — PowerPoint-macro en twee nieuwe dia's
  4. `bouw_tabbladen_grafieken.py` — verdieping-tabbladen achter één knop, rustiger grafieken
  5. `afslanken.py` — kleiner bestand (laatste stap, van `bron/` naar `RDFO_template.xlsm`)

Een volgende wijziging: werk op `bron/RDFO_template_bron.xlsm` en draai daarna `afslanken.py` opnieuw.
