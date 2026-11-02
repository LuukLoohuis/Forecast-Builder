# Forecast-Builder · FO-werkbestand

Eén Excel-bestand voor het kwartaal-FO: de FO-onderlegger en het cashflow-scenariomodel samen, met een
prognose op basis van een voorstel of met de hand, en een PowerPoint die uit hetzelfde bestand komt.

| Bestand | Wat |
|---|---|
| `output/FO_Parkzoom_werkbestand.xlsm` | het werkbestand (fictieve Parkzoom-data) |
| `output/Kwartaal_Template_cashflow_v21.pptx` | PowerPoint-sjabloon: v20 met namen voor de vormen die nu ook gevuld worden (lay-out ongewijzigd) |
| `docs/werkinstructie.md` | kwartaal-FO in stappen (één A4) |
| `docs/prompt-forecast-1-bestand.md` | de uitgangspunten (prompt) |
| `input/` | de bronbestanden: demo-FO, scenariomodel, sjabloon v20 |
| `vba/` | broncode van alle macro's |
| `tools/` | bouw- en controlescripts |

## Opbouw van het werkbestand

| Groep | Tabbladen |
|---|---|
| Start | `Start`: project en periode, status van de invoer, controles, stappen, teksten voor dia 10 en de knoppen |
| Invoer (plakken of inlezen) | `FO - vorig`, `FO - huidig` (SAP PEW, FR4), `FO - actuals` (SAP FR5), `Prijslijst dump` (Salesforce) |
| FO zoals je gewend bent | `Prijslijst`, `Aansluiting opbrengsten`, `Verkooptempo`, `CF - opbrengsten`, `CF - kosten` |
| Prognose (nieuw) | `Prognose opbrengsten`, `Prognose kosten`: voorstel, handmatig en definitief |
| Uitvoer | `1. Cashflow`, `2. Ontw PEW`, `3. Restant bdgt`, `CF - hulp VENA` |
| Scenario's en PowerPoint | `Dashboard`, `Scenario's`, `Invoer`, `Woningtypes`, `Scenario kosten` (was `CF kosten`), `Model`, `PowerPoint` |
| Overig | `Notities FO gesprek`, `Vorige prognose`, `Controle oud FO`, `Behind the scens` (+ verborgen bladen uit het FO) |

**Prognose kosten.** Per post kies je in kolom C:
- **Handmatig**: de verdeling op `CF - kosten`, zoals je die altijd maakt.
- **Voorstel**: het restantbudget (PEW min wat al verwerkt is), gelijk verdeeld vanaf het eerste prognosekwartaal t/m het opleveringskwartaal op `Start`.

Blok 2 laat per kwartaal het verschil tussen handmatig en voorstel zien.

**Prognose opbrengsten.** Het voorstel bestaat uit de termijnen volgens `CF - opbrengsten`: bouwplanning × koopsommen uit `Aansluiting opbrengsten`. Een bedrag in de gele regel overschrijft het voorstel voor dat kwartaal. Voor kwartalen die al verwerkt zijn telt altijd het voorstel.

**Definitief.** De definitieve totalen staan op dezelfde rij en kolom als op `CF - kosten` (rij 75/77) en `CF - opbrengsten` (rij 194/196).
- `1. Cashflow` en `CF - hulp VENA` verwijzen daardoor alleen naar een andere tabbladnaam; hun formules zijn verder gelijk aan het oude FO.
- "Oud FO inlezen" kan oude FO's daardoor 1-op-1 overnemen.

## Knoppen (tab Start)

| Knop | Doet |
|---|---|
| Plakken / Inlezen... | per invoertabblad: leegmaken, export erin (klembord of bestand), hulpkolommen doortrekken, status bijwerken |
| Oud FO inlezen... | neemt alle tabbladen van een FO in de oude opzet over (formules en waarden), daarna `Controle oud FO` |
| Nieuw kwartaal starten | cashflow bewaren als vorige prognose (en in kolom J van `1. Cashflow`), `FO - huidig` → `FO - vorig`, huidig en actuals leeg |
| Scenario bijwerken | vult `Invoer`, `Woningtypes` en `Scenario kosten` uit de FO-tabbladen van dit bestand (geen bestand kiezen meer) |
| PowerPoint maken | eerst Scenario bijwerken, dan Naar PowerPoint met sjabloon v21 (zet het sjabloon in dezelfde map) |

## Wat is gecontroleerd, en wat nog niet

- **Getest met LibreOffice (headless, volledig doorgerekend):**
  - `1. Cashflow`, `CF - hulp VENA`, `2. Ontw PEW`, `3. Restant bdgt` en `Aansluiting opbrengsten` zijn per cel gelijk aan de waarden die Excel in het demo-FO heeft opgeslagen.
  - Alle controles op `Start` zijn groen. De enige uitzondering is "Scenario actueel": die vraagt om één klik op Scenario bijwerken.
  - Met alle posten op Voorstel en een handmatig bedrag bij de opbrengsten blijven de kosten gelijk aan het PEW. De VENA-export volgt de definitieve cashflow, en de controles melden het handmatige bedrag.
- **Macro's:**
  - Het VBA-project is opnieuw opgebouwd uit `vba/`, zonder gecompileerde code. Excel compileert het bij het openen.
  - LibreOffice leest alle 33 modules correct uit het bestand.
  - De macro's zijn **niet in Excel uitgevoerd**: daar is in deze omgeving geen Excel voor. Test bij het eerste gebruik de knoppen op een kopie.
  - Geeft Excel een melding over het VBA-project? Importeer dan `vba/FOWerkbestand.bas` en `vba/CashflowNaarPowerPoint.bas` via Alt+F11 → Bestand → Importeren, en zet de code van `vba/ThisWorkbook.cls` in ThisWorkbook.
- Het scenariodeel (`Invoer`, `Woningtypes`, `Scenario kosten`) wordt pas gevuld met de eerste klik op Scenario bijwerken.

## Opnieuw bouwen

```bash
python3 tools/build_werkbestand.py
```

```bash
python3 tools/controleer.py
```

Nodig: Python 3 met `lxml`, `openpyxl`, `python-pptx`, `olefile` en `oletools`. Voor de controle ook LibreOffice.
