# Forecast-Builder

Bouwt het werkboek **Cashflow_scenario** (Dashboard · Invoer · Woningtypes · Model · PowerPoint) met Python/openpyxl.
Het werkboek rekent scenario's (downside/upside) uit bovenop je eigen projectcashflow en vult met één knop
het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v9.pptx`).

## Wat zit erin

| Map / bestand | Inhoud |
|---|---|
| `build.py` | opdrachtregel: bouwt `out/Cashflow_scenario.xlsx` (en `.xlsm` als er een VBA-bron is) |
| `forecast_builder/layout.py` | vaste indeling: kolommen, rijen, hulpcellen, capaciteit (`N_TYPES = 10`) |
| `forecast_builder/sheets.py` | alle tabbladen en formules |
| `forecast_builder/data.py` | projectgegevens: inlezen uit een bestaand werkboek (oude of nieuwe indeling) of het ingebouwde voorbeeld |
| `forecast_builder/charts.py` | grafiek-XML voor het Dashboard |
| `forecast_builder/package.py` | nabewerking: grafieken en VBA in het bestand zetten |
| `vba/CashflowNaarPowerPoint_v10.bas` | VBA-module: Naar PowerPoint, Sjabloon controleren, Type invoegen, Type verwijderen |
| `forecast_builder/vbabuild.py` | schrijft de VBA-module in het .xlsm (vbaProject.bin: MS-CFB + MS-OVBA), zodat importeren niet nodig is |
| `tools/pptx_rente_card.py` | zet de vijfde KPI-kaart (rente) op dia 2 van het PowerPoint-sjabloon (v9 → v10) |

## Bouwen

```bash
python build.py                                      # voorbeeldproject -> out/Cashflow_scenario.xlsx + .xlsm
python build.py --from Cashflow_scenario.xlsm        # invoer en knoppen uit je huidige werkboek (oude of nieuwe indeling)
python build.py --from oud.xlsm --out map/naam       # eigen uitvoernaam (zonder extensie)
python build.py --geen-vba                           # alleen .xlsx
```

Het `.xlsm` bevat de macro al: `vba/CashflowNaarPowerPoint_v10.bas` wordt bij het bouwen in `xl/vbaProject.bin`
geschreven (`forecast_builder/vbabuild.py`, op basis van `vba/vbaProject_template.bin` met de documentmodules en
projectinstellingen). Importeren in de VBA-editor is dus niet nodig. Bij het openen zet `KnoppenControleren`
de knoppen neer (Dashboard, PowerPoint, Woningtypes). Zet `Kwartaal_Template_cashflow_v10.pptx` naast het werkboek
(of vul het pad in op tab PowerPoint).

`--from` begrijpt zowel de oude indeling (vijf types naast elkaar op tab Invoer) als de nieuwe (typeblokken), dus een
bijgewerkt werkboek genereer je later opnieuw zonder gegevens over te typen.

Tests: `python3 -m pytest tests/ -q` (de modeltests gebruiken LibreOffice om formules door te rekenen).

## Indeling van het werkboek

* **Dashboard** – kaarten, drie grafieken en in kolom V alle knoppen: actuals t/m, jaarrente, rente t/m (start bouw of
  hele looptijd), rente ook in de basis (ja/nee), verkooptempo-model aan/uit, per scenario: verschuiving verkoop/transport,
  uitstel start bouw, opbrengsten %, kosten %; norm verkocht vóór start bouw; controles; overzicht woningtypes.
* **Invoer** – je eigen cashflow (B:J, plakken als waarden) en per woningtype twee kolommen naast elkaar:
  verkocht | transport (vanaf kolom L). Rijen (kwartalen) mag je verwijderen of invoegen en lege rijen tellen niet
  mee: het Model leest de i-de gevulde rij (kolom Bronrij, een matrixformule), dus grafieken en KPI's volgen vanzelf.
* **Woningtypes** – per type één blok van drie kolommen (vanaf kolom C): naam, aantal, koopsom, start bouw en een eigen
  lijst termijnen (naam, % van de koopsom, bouwkwartaal). Elk type kan dus andere termijnen hebben.
* **Model** – rekenblad. Kolommen A:L uit de invoer (B = bronrij op Invoer), M:BF model, scenario's, rente en
  grafiekkolommen, BX:BY hulpcellen (BY8 = aantal periodes, de VBA leest die cel), vanaf CA de typeblokken.
* **PowerPoint** – teksten, tabellen en grafiekblokken voor de dia's; kolom F = vormnaam op de dia.

### Type toevoegen of verwijderen

Het model leest de typeblokken op positie (`INDEX` op `Woningtypes!$A:$ZZ` en `Invoer!$A:$ZZ`). Daardoor schuift alles
mee als je kolommen invoegt of verwijdert:

* toevoegen: vul het eerstvolgende lege blok, **of** voeg op de gewenste plek drie kolommen in op Woningtypes en twee op
  Invoer (knop *Type invoegen* doet dat in één keer en neemt de opmaak over);
* verwijderen: drie kolommen weg op Woningtypes en twee op Invoer (knop *Type verwijderen*).

Capaciteit: `N_TYPES` in `layout.py` (standaard 10). Na aanpassen opnieuw bouwen en in het `.bas` `MAX_TYPES` gelijk zetten.

### Rente

Rente = negatieve stand van het vorige kwartaal × jaarrente / 4, alleen in prognosekwartalen, t/m het kwartaal vóór start
bouw van het project (vroegste type) plus het uitstel van dat scenario, of over de hele looptijd.

* *Rente ook in de basis = nee* (standaard): de basis blijft je eigen cashflow; downside en upside krijgen alleen het
  verschil in rente ten opzichte van de basis. Met alle knoppen op nul zijn de scenario's dan gelijk aan de basis.
* *ja*: basis en scenario's krijgen rente; kaarten, kosten en dia's rekenen daarmee.

## Controle

`verify`-stappen die bij het bouwen zijn gedaan: LibreOffice-herberekening zonder foutcellen (behalve de bedoelde `#N/B`
in grafiekkolommen), het oude werkboek wordt met rente 0% cel voor cel gereproduceerd, de rente is nagerekend met een
losse simulatie, en met alle knoppen op nul vallen de scenario's samen met de basis.
