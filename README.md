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
| Resultaat | `Dashboard` (cashflow, KPI's, scenario's in grafieken), `Scenario's`, `2. Ontw PEW`, `CF - hulp VENA`, `PowerPoint` |
| Rekenbladen (automatisch, niet invullen) | `Invoer`, `Woningtypes`, `Scenario kosten`, `Model` |
| Overig | `Notities FO gesprek`, `Vorige prognose`, `Controle oud FO`, `Verkooptempo` (telling uit Salesforce), `Behind the scens` (+ verborgen hulpbladen) |

Vervallen ten opzichte van het oude FO: `1. Cashflow` (de cashflowtabel staat nu op `Invoer`, de grafieken op het
`Dashboard`) en `3. Restant bdgt` (de werkelijke kosten komen rechtstreeks uit `FO - huidig`).

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
- De totaalregel (rij 75) is het **definitieve** resultaat.

**DAEB-fees.** Onderaan `CF - opbrengsten` (rij 200-227) staat een blok "Opbrengsten DAEB":
- in de drie rijen boven de koprij per fee-component de naam (kolom C) en het totaal (kolom E);
- daaronder per component een groep termijnen (mijlpaal, bedragen per kwartaal of in actuals), gescheiden door een lege rij;
- het aantal DAEB-woningen en de verkoop/levering per kwartaal; "Startbouw" boven een kwartaal mag.

Het totaal telt mee in rij 194. Woningtype 5 op `Woningtypes` wordt dan automatisch het DAEB-type:
- naam, aantal en start bouw;
- soort DAEB;
- in blok 5 per component maximaal vijf fee-regels: per kwartaal "(1/2)", "(2/2)", een bedrag in actuals als "actuals". Bij meer delen één regel per termijn, en de rest samen in "overige termijnen".

Verkoop en transport staan op `Invoer` (transport = verkoop). De formules zoeken het blok op aan de titel met "DAEB" en volgen dezelfde regels als de oude macro. Dat werkt dus ook voor het DAEB-blok van een oud FO. Staat het aantal daar niet, dan komt het van `Verkooptempo`: uit de rij met "DAEB" in de naam, of anders uit de eerste rij na de koopwoningtypes (bijvoorbeeld de naam van de corporatie), net als de oude macro. Getest op het DAEB-blok van een echt FO, met AK fees en bijkomende kosten.

**Scenariomodel zonder knop.** `Invoer`, `Woningtypes` en `Scenario kosten` zijn formules op de FO-tabbladen, net zoals `Aansluiting opbrengsten` op de prijslijst. Het `Dashboard` en `Scenario's` rekenen dus altijd mee.

| Tabblad | Wat erin staat |
|---|---|
| `Invoer` | de cashflow per periode, met dezelfde opbouw als het oude `1. Cashflow`: verwerkte kwartalen uit `FO - actuals`, prognose uit rij 75 en rij 194, kolom J = vorige prognose. Daarnaast verkocht en transport per woningtype per kwartaal |
| `Woningtypes` | per woningtype naam, aantal, koopsom, start bouw, grondtermijn, bouwtermijnen (% en bouwkwartaal), extra's per woning en bouwkosten per termijn |
| `Scenario kosten` | per kostenpost het gerealiseerde bedrag en de definitieve prognose per periode, volgens de methode op `CF - kosten`. "Schuift mee met" wordt op de naam bepaald; overschrijven mag |

Het verborgen hulpblad `_koppeling` zet de termijnen om naar regels per bouwkwartaal, met dezelfde regels als de oude macro "Ophalen uit FO". Een termijn die over twee kwartalen valt, wordt "(1/2)" en "(2/2)".

## Knoppen (tab Start)

| Knop | Doet |
|---|---|
| Plakken / Inlezen... | per invoertabblad: leegmaken, export erin (klembord of bestand), hulpkolommen doortrekken, status bijwerken |
| Oud FO inlezen... | neemt alle tabbladen van een FO in de oude opzet over. Zet daarna de nieuwe onderdelen terug (transportregels, methode, voorstel, correctie; lijst op het verborgen blad `_sjabloon`), vult `Vorige prognose` en `Controle oud FO` |
| Nieuw kwartaal starten | cashflow bewaren als vorige prognose, `FO - huidig` → `FO - vorig`, huidig en actuals leeg |
| PowerPoint maken | doorrekenen en Naar PowerPoint met sjabloon v21 (zet het sjabloon in dezelfde map) |

## Wat is gecontroleerd, en wat nog niet

- **Getest met LibreOffice (headless, volledig doorgerekend, inclusief het scenariomodel):**
  - De cashflow op `Invoer` is per cel gelijk aan het oude `1. Cashflow`, inclusief totalen en controles.
  - `CF - opbrengsten`, `CF - kosten`, `CF - hulp VENA`, `2. Ontw PEW` en `Aansluiting opbrengsten` zijn gelijk aan de waarden die Excel in het demo-FO heeft opgeslagen.
  - Alle controles op `Start` zijn groen.
  - "Model dekt de opbrengsten" komt uit op 100,0%: aantal × koopsom op `Woningtypes` sluit precies aan op de opbrengsten. `Scenario kosten` sluit aan op `Invoer`.
  - Scenario met de start bouw drie kwartalen later (Q1 '28): de maximale financieringsbehoefte gaat van €1,6 mln naar €6,8 mln.
  - Scenario met de verkoop pas vanaf Q2 '27: €4,9 mln. Een kostenpost op Voorstel komt direct terug in `Scenario kosten`.
  - Getest met woningtype A pas in Q1 2028 getransporteerd en met een correctie van €250.000; de controles melden de afwijking van het PEW.
  - DAEB getest met twee fee-componenten (€1,0 mln), actuals, een termijn over twee kwartalen, een Startbouw-markering en 30 woningen: woningtype 5 wordt DAEB met de juiste fee-regels, het `Model` telt €1,0 mln aan fees. Met 14 delen werkt de terugval naar één regel per termijn.
- **Macro's:**
  - Het VBA-project is opnieuw opgebouwd uit `vba/`, zonder gecompileerde code. Excel compileert het bij het openen.
  - LibreOffice leest alle 31 modules correct uit het bestand.
  - De macro's zijn **niet in Excel uitgevoerd**: daar is in deze omgeving geen Excel voor. Test bij het eerste gebruik de knoppen op een kopie.
  - Geeft Excel een melding over het VBA-project? Importeer dan `vba/FOWerkbestand.bas` en `vba/CashflowNaarPowerPoint.bas` via Alt+F11 → Bestand → Importeren, en zet de code van `vba/ThisWorkbook.cls` in ThisWorkbook.
- Het scenariodeel rekent met de vier woningtypeblokken van de standaardopzet, plus het DAEB-blok als woningtype 5.
- **Let op:** het werkbestand is gebouwd op de opzet van het demo-FO (4 woningtypes, kwartalen 2026-2029, vaste rijen). Een FO met een andere opzet (bijvoorbeeld 3 woningtypes, andere rijen of kwartalen tot 2030) past daar niet 1-op-1 in.
- "Oud FO inlezen" gaat uit van de standaardopzet van het FO: transportregels in de lege rijen 13, 26, 36 en 46 van `CF - opbrengsten`, posten in rij 59-70 van `CF - kosten`. Wijkt een oud FO daarvan af, dan meldt de knop dat. `Controle oud FO` laat het effect zien.

## Opnieuw bouwen

```bash
python3 tools/build_werkbestand.py
```

```bash
python3 tools/controleer.py
```

Nodig: Python 3 met `lxml`, `openpyxl`, `python-pptx`, `olefile` en `oletools`. Voor de controle ook LibreOffice.
