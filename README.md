# Forecast-Builder · FO-werkbestand

Eén Excel-bestand voor het kwartaal-FO: de FO-onderlegger en het cashflow-scenariomodel samen, met een
prognose op basis van een voorstel of met de hand, en een PowerPoint die uit hetzelfde bestand komt.

| Bestand | Wat |
|---|---|
| `output/FO_Parkzoom_werkbestand.xlsm` | het werkbestand (fictieve Parkzoom-data) |
| `output/FO_Parkzoom_werkbestand.opzet.json` | de opzet van het FO waarop het is gebouwd (rijen, kolommen; voor de controle) |
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
| Opbrengsten | `Prijslijst` (btw eraf, uitgesplitst), `Aansluiting opbrengsten` (SKO ↔ prijslijst ↔ PEW, termijnschema, onderaan "te boeken in SAP"), `CF - opbrengsten` (per woningtype: bouwplanning, verkocht % en getransporteerd %, grondprijs; correctie) |
| Kosten | `CF - kosten` (bouwtermijnen met eigen planning + per post Handmatig, Voorstel of Transport) |
| Resultaat | `Dashboard` (cashflow, KPI's, scenario's in grafieken), `Scenario's`, `2. Ontw PEW` (en eventueel `2B. Ontw PEW`), `CF - hulp VENA`, `PowerPoint` |
| Overig | `Vorige prognose`, `Controle oud FO`, `Behind the scens` |
| Verborgen rekenbladen (alleen formules) | `Invoer`, `Woningtypes`, `Scenario kosten`, `Model` (het scenariomodel), `_koppeling`, `_sjabloon` |

Vervallen ten opzichte van het oude FO: `1. Cashflow` (de cashflowtabel staat nu op `Invoer`, de grafieken op het
`Dashboard`), `3. Restant bdgt` (de werkelijke kosten komen rechtstreeks uit `FO - huidig`), `Notities FO gesprek` en
`Verkooptempo` (verkocht t/m nu komt uit de Salesforce-status in de prijslijst, de planning staat op `CF - opbrengsten`).
`Woningtypes`, `Invoer` en `Scenario kosten` zijn verborgen: ze rekenen alleen door wat op de FO-tabbladen staat.

**CF - opbrengsten.** Per woningtype, onder de bouwplanning:
- de **transportregel** (aantal woningen; rekent uit de % hieronder) = de grondtermijn;
- **Verkocht %**: per kwartaal, van alle woningen van het type. De eerste kolom is verkocht t/m nu: het aantal met status "Verkocht" of "Getransporteerd" in de prijslijst (Salesforce);
- **Getransporteerd %**: per kwartaal, van alle woningen. Transport kan niet vóór verkoop: schuift de verkoop op, dan schuift het transport mee (status "LET OP: transport wacht op verkoop");
- **Grondprijs**: een aanpassing per woning (na goedkeuring SWK) telt op bij de koopsom grond;
- de **bouwtermijnen** vul je zelf in (planning met de ontwikkelaar) en tellen vanaf transport.

Het verkooptempo onderaan rekent uit de %. Per type staat in kolom AB het totaal en in kolom AD de status (meer dan 100% = FOUT). De gele correctieregel boven het totaal (demo: rij 204) telt mee in het totaal.

**CF - kosten**
- De bouwtermijnen (aanneemsom, planning HWB) volgen standaard de *geplande* bouwtermijnen van CF - opbrengsten. Overschrijven mag, dus een eigen planning kan.
- Bij de posten (blok Onderdelen) kies je in kolom C:
  - **Handmatig**: de regels zoals je ze altijd invult;
  - **Voorstel**: het blok onderaan (rij 100-111), met het restantbudget (PEW min wat al verwerkt is) gelijk verdeeld vanaf het eerste prognosekwartaal t/m het opleveringskwartaal op `Start`;
  - **Transport**: het restantbudget naar rato van het transport van alle woningtypes (bijvoorbeeld de grond).
- Bij Voorstel en Transport gaat wat je in de postregel invult voor; lege kwartalen krijgen het voorstel, en het restant wordt over die kwartalen verdeeld.
- De totaalregel (rij 75) is het **definitieve** resultaat.

**Aansluiting opbrengsten · te boeken in SAP.** Onderaan het tabblad staat per SAP-nummer en kostendrager wat er na de aansluiting in SAP aangepast moet worden (uit de kolom "Corr"):
- standaard: 9010 grond + ontwikkelkosten, 9000 aanneemsom, 9000/5230 KAO-opties, 9200/5250 meer-/minderwerk kopers, 9030 kadastrale kosten, 9110 aanvullende rente;
- daaronder de koppeling regel → SAP-nummer (aanpasbaar) en het detail per woningtype;
- staat er een bedrag op een regel zonder SAP-nummer, dan meldt `Start` dat.

**DAEB-fees.** Onderaan `CF - opbrengsten` (demo: rij 212-239) staat een blok "Opbrengsten DAEB":
- in de drie rijen boven de koprij per fee-component de naam (kolom C) en het totaal (kolom E);
- daaronder per component een groep termijnen (mijlpaal, bedragen per kwartaal of in actuals), gescheiden door een lege rij;
- het aantal DAEB-woningen en de verkoop/levering per kwartaal; "Startbouw" boven een kwartaal mag.

Het totaal telt mee in de totaalregel van CF - opbrengsten. Woningtype 5 op `Woningtypes` wordt dan automatisch het DAEB-type:
- naam, aantal en start bouw;
- soort DAEB;
- in blok 5 per component maximaal vijf fee-regels: per kwartaal "(1/2)", "(2/2)", een bedrag in actuals als "actuals". Bij meer delen één regel per termijn, en de rest samen in "overige termijnen".

Verkoop en transport staan op `Invoer` (transport = verkoop, uit de regel "Verkoop en levering DAEB"). De formules zoeken het blok op aan de titel met "DAEB" en volgen dezelfde regels als de oude macro. Dat werkt dus ook voor het DAEB-blok van een oud FO. Stond het aantal en de verkoop van de corporatie in het FO op `Verkooptempo`, dan zet de bouw ze als regels "Aantal DAEB-woningen" en "Verkoop en levering DAEB" onder het blok. Getest op het DAEB-blok van een echt FO, met AK fees en bijkomende kosten.

**Scenariomodel zonder knop.** `Invoer`, `Woningtypes` en `Scenario kosten` zijn formules op de FO-tabbladen, net zoals `Aansluiting opbrengsten` op de prijslijst. Het `Dashboard` en `Scenario's` rekenen dus altijd mee.

| Tabblad | Wat erin staat |
|---|---|
| `Invoer` | de cashflow per periode, met dezelfde opbouw als het oude `1. Cashflow`: verwerkte kwartalen uit `FO - actuals`, prognose uit de totaalregels van CF - kosten en CF - opbrengsten, kolom J = vorige prognose. Daarnaast verkocht en transport per woningtype per kwartaal |
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
- **Getest op een echt FO (3 woningtypes, DAEB met AK fees en bijkomende kosten, kwartalen t/m 2030, 1. Cashflow vanaf 2019):**
  - het werkbestand is uit dat FO gebouwd zonder iets aan te passen;
  - de cashflow op `Invoer` (36 periodes, met totalen en controles), `CF - opbrengsten`, `CF - kosten`, `CF - hulp VENA`, `2. Ontw PEW` en `Aansluiting opbrengsten` zijn per cel gelijk aan de waarden die Excel in dat FO heeft opgeslagen;
  - de controles die op FOUT staan, staan in het originele FO ook op FOUT (zie hieronder bij "Bevindingen bij een echt FO").
  - het echte FO zelf staat niet in deze repository.
- Het scenariodeel kan tot 4 woningtypes aan, plus het DAEB-blok als het volgende type (bij 4 typen: type 5), en tot 60 periodes.
- "Oud FO inlezen" gaat uit van de opzet van het FO waarop het werkbestand is gebouwd (de rijen staan op het verborgen blad `_sjabloon`, kolom E:F). Lees dus een oud FO van hetzelfde project in. Wijkt het daarvan af, dan meldt de knop dat; `Controle oud FO` laat het effect zien.

## Werkbestand voor een ander project

```bash
python3 tools/build_werkbestand.py --fo pad/naar/FO_Project_Q4.xlsx --out werkbestand/FO_Project_werkbestand.xlsm
```

```bash
python3 tools/controleer.py werkbestand/FO_Project_werkbestand.xlsm --fo pad/naar/FO_Project_Q4.xlsx
```

Het script leest eerst de opzet van het FO (`tools/opzet.py`):
- het aantal woningtypes, de bouwplanning en de lege rij eronder (daar komt de transportregel), en de omzetblokken;
- de kwartaalkolommen, de check- en statuskolommen;
- de totaalrij, de correctieregel (een lege rij boven het totaal) en een eventueel eigen DAEB-blok;
- op `CF - kosten`: de posten, het totaal en de vrije rijen voor het voorstelblok;
- de kolommen van `2. Ontw PEW` (ook met een extra kolom zoals "FD corr DAEB") en de periodes van `1. Cashflow`.

Daarna bouwt het alles op die opzet:
- de projectnaam komt uit de bestandsnaam (of `--project`);
- het opleveringskwartaal komt uit de markering "Oplevering" boven de kwartalen (of `--oplevering 2030-2`);
- extra tabbladen van het FO (zoals `2B. Ontw PEW`) gaan mee;
- verwijzingen naar een ander bestand (`'[2]FO - huidig'!`) wijzen daarna naar het tabblad in het werkbestand;
- namen die naar een ander bestand of naar `#REF!` wijzen, vallen weg.

Het eerste prognosekwartaal is het kwartaal na de laatste actuals in `FO - actuals` (overschrijven mag op `Start`).

**Bevindingen bij een echt FO.** Het script meldt wat het tegenkomt. Bij het geteste FO:
- de grondtermijn van één woningtype verwees naar de lege rij onder de bouwplanning (uitkomst 0). In de transportregel is dat de eerste planningsrij geworden; de uitkomst bleef gelijk;
- 2B. Ontw PEW gaf in Excel `#VALUE!` in de controlekolom (die verwees naar een ander bestand); nu rekent die kolom met `FO - huidig` van het werkbestand;
- de kosten in `1. Cashflow` sloten niet aan op het PEW (FOUT, ook in het FO zelf): het PEW-bedrag was met de hand ingevuld en er stonden actuals van vóór de eerste kolom van `CF - kosten`. Het werkbestand vergelijkt met `2. Ontw PEW` en laat het verschil dus zien.

## Opnieuw bouwen

```bash
python3 tools/build_werkbestand.py
```

```bash
python3 tools/controleer.py
```

Nodig: Python 3 met `lxml`, `openpyxl`, `python-pptx`, `olefile` en `oletools`. Voor de controle ook LibreOffice.
