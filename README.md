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
| Opbrengsten | `Prijslijst`, `Aansluiting opbrengsten`, `CF - opbrengsten` (bouwtermijnen, transport, verkooptempo, correctie) |
| Kosten | `CF - kosten` (bouwtermijnen + per post Handmatig of Voorstel) |
| Uitvoer | `Cashflow` (tabel + grafiek), `2. Ontw PEW` (nieuw jasje + grafiek), `CF - hulp VENA` |
| Scenario's en PowerPoint | `Dashboard`, `Scenario's`, `Invoer`, `Woningtypes`, `Scenario kosten` (was `CF kosten`), `Model`, `PowerPoint` |
| Overig | `Notities FO gesprek`, `Vorige prognose`, `Controle oud FO`, `Verkooptempo` (telling uit Salesforce), `Behind the scens` (+ verborgen hulpbladen) |

Vervallen ten opzichte van het oude FO: `1. Cashflow` (nu `Cashflow`), `3. Restant bdgt` (de werkelijke kosten komen
rechtstreeks uit `FO - huidig`).

**CF - opbrengsten**
- Per woningtype staan de bouwtermijnen (handmatig) met daaronder een **transportregel**. De standaardformule rekent zoals de oude grondtermijn, dus zonder aanpassing verandert er niets. Overschrijf hem met het aantal woningen per kwartaal.
- De **grondtermijn** volgt het transport.
- De **bouwtermijnen tellen vanaf transport**: termijnen die bij transport al vervallen zijn, komen in het transportkwartaal in één keer binnen.
- Het verkooptempo staat in rij 49-52.
- Rij 192 is een gele **correctieregel** per kwartaal (in €). Die telt mee in het totaal (rij 194).

**CF - kosten**
- De bouwtermijnen volgen de *geplande* bouwtermijnen van CF - opbrengsten. Overschrijven mag.
- Bij de posten (blok Onderdelen) kies je in kolom C:
  - **Handmatig**: de regels zoals je ze altijd invult;
  - **Voorstel**: het blok onderaan (rij 100-111), met het restantbudget (PEW min wat al verwerkt is) gelijk verdeeld vanaf het eerste prognosekwartaal t/m het opleveringskwartaal op `Start`.
- De totaalregel (rij 75) is het **definitieve** resultaat. `Cashflow`, `CF - hulp VENA` en het scenario lezen die regel.

**Cashflow.** Een tabel per periode met dezelfde opbouw als het oude `1. Cashflow`:
- verwerkte kwartalen: werkelijk uit `FO - actuals`;
- vanaf het eerste prognosekwartaal: de definitieve CF-regels;
- kolom J: de vorige prognose.

Erbij staat een grafiek met kosten, opbrengsten, de cumulatieve cashflow en de vorige prognose.

## Knoppen (tab Start)

| Knop | Doet |
|---|---|
| Plakken / Inlezen... | per invoertabblad: leegmaken, export erin (klembord of bestand), hulpkolommen doortrekken, status bijwerken |
| Oud FO inlezen... | neemt alle tabbladen van een FO in de oude opzet over. Zet daarna de nieuwe onderdelen terug (transportregels, methode, voorstel, correctie; lijst op het verborgen blad `_sjabloon`), vult `Vorige prognose` en `Controle oud FO` |
| Nieuw kwartaal starten | cashflow bewaren als vorige prognose, `FO - huidig` → `FO - vorig`, huidig en actuals leeg |
| Scenario bijwerken | vult `Invoer`, `Woningtypes` en `Scenario kosten` uit dit bestand. Kosten per post volgens de gekozen methode; transport uit de grondtermijn |
| PowerPoint maken | eerst Scenario bijwerken, dan Naar PowerPoint met sjabloon v21 (zet het sjabloon in dezelfde map) |

## Wat is gecontroleerd, en wat nog niet

- **Getest met LibreOffice (headless, volledig doorgerekend):**
  - `Cashflow` is per cel gelijk aan het oude `1. Cashflow`, inclusief totalen en controles.
  - `CF - opbrengsten` (rij 57-194), `CF - kosten` (rij 59-77), `CF - hulp VENA`, `2. Ontw PEW` en `Aansluiting opbrengsten` zijn gelijk aan de waarden die Excel in het demo-FO heeft opgeslagen.
  - Alle controles op `Start` zijn groen. De enige uitzondering is "Scenario actueel": die vraagt om één klik op Scenario bijwerken.
  - Getest met woningtype A pas in Q1 2028 getransporteerd: de grondtermijn en alle eerder vervallen termijnen komen in Q1 2028 binnen.
  - Getest met een correctie van €250.000 en de aanneemsom op Voorstel: de correctie telt mee, het voorstel wordt gelijk verdeeld, de VENA-export volgt en de controles melden de afwijking.
- **Macro's:**
  - Het VBA-project is opnieuw opgebouwd uit `vba/`, zonder gecompileerde code. Excel compileert het bij het openen.
  - LibreOffice leest alle 31 modules correct uit het bestand.
  - De macro's zijn **niet in Excel uitgevoerd**: daar is in deze omgeving geen Excel voor. Test bij het eerste gebruik de knoppen op een kopie.
  - Geeft Excel een melding over het VBA-project? Importeer dan `vba/FOWerkbestand.bas` en `vba/CashflowNaarPowerPoint.bas` via Alt+F11 → Bestand → Importeren, en zet de code van `vba/ThisWorkbook.cls` in ThisWorkbook.
- Het scenariodeel (`Invoer`, `Woningtypes`, `Scenario kosten`) wordt pas gevuld met de eerste klik op Scenario bijwerken.
- "Oud FO inlezen" gaat uit van de standaardopzet van het FO: transportregels in de lege rijen 13, 26, 36 en 46 van `CF - opbrengsten`, posten in rij 59-70 van `CF - kosten`. Wijkt een oud FO daarvan af, dan meldt de knop dat. `Controle oud FO` laat het effect zien.

## Opnieuw bouwen

```bash
python3 tools/build_werkbestand.py
```

```bash
python3 tools/controleer.py
```

Nodig: Python 3 met `lxml`, `openpyxl`, `python-pptx`, `olefile` en `oletools`. Voor de controle ook LibreOffice.
