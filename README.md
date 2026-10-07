# Forecast-Builder

Bouwt het werkboek **Cashflow_scenario** (Dashboard · Invoer · Woningtypes · Model · PowerPoint) met Python/openpyxl.
Het werkboek rekent scenario's (downside/upside) uit bovenop je eigen projectcashflow en vult met één knop
het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v11.pptx`).

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
| `forecast_builder/cache.py` | rekent door met LibreOffice en zet cachewaarden in cellen en grafieken |
| `tools/pptx_rente_card.py` | zet de vijfde KPI-kaart (rente) op dia 2 van het PowerPoint-sjabloon (v9 → v10) |
| `tools/pptx_scenario_line.py` | zet de scenariolijn als zevende reeks (kolom H, amber) in de cashflowgrafiek op dia 3 (v10 → v11) |

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
de knoppen neer (Dashboard, PowerPoint, Woningtypes). Zet `Kwartaal_Template_cashflow_v11.pptx` naast het werkboek
(of vul het pad in op tab PowerPoint).

`--from` begrijpt zowel de oude indeling (vijf types naast elkaar op tab Invoer) als de nieuwe (typeblokken), dus een
bijgewerkt werkboek genereer je later opnieuw zonder gegevens over te typen.

Als LibreOffice (`soffice`) aanwezig is, rekent de build het werkboek door en zet de uitkomsten als cachewaarden in
cellen en grafieken (`forecast_builder/cache.py`), zodat Excel ook in de beveiligde weergave meteen cijfers toont.
Zonder LibreOffice werkt het bestand ook; Excel rekent dan bij het openen (`--geen-cache` slaat de stap over).

Tests: `python3 -m pytest tests/ -q` (de modeltests gebruiken LibreOffice om formules door te rekenen).

## Indeling van het werkboek

* **Dashboard** – kaarten, drie grafieken en in kolom V alle knoppen: actuals t/m, jaarrente, rente t/m (start bouw of
  hele looptijd), rente ook in de basis (ja/nee), verkooptempo-model aan/uit, per scenario (downside, upside én de
  scenariolijn): verschuiving verkoop/transport, uitstel start bouw, opbrengsten %, kosten %, eigen jaarrente (leeg =
  algemeen) en koopsom % (VON-prijs via het verkooptempo-model); scenariolijn aan/uit en naam;
  norm verkocht vóór start bouw; tien controles (oplopende kwartalen zonder gaten, maximaal 60 periodes, rijen zonder jaar,
  jaar als getal, koprij, verkocht/transport = aantal, termijnen); overzicht types.
  De grafieken (`forecast_builder/chartxml.py`): *scenario's* (basis met de band tussen downside en upside), *cashflow per
  kwartaal* (opbrengsten, kosten, cumulatieve cashflow, vorige prognose en de gele scenariolijn vanaf het punt 'nu') en
  *verkoop en transport per woningtype* (twee panelen boven elkaar: verkocht en getransporteerd per kwartaal, gestapeld
  per type in een vaste kleur per typeblok, totaal per kwartaal boven de stapel, cel-legenda boven het eerste paneel,
  het aantal in elk blokje, mijlpalen 'uitverkocht' en 'laatste transport'). Een grijs vlak markeert de gerealiseerde
  kwartalen.
* **Invoer** – je eigen cashflow (B:J, plakken als waarden) en per woningtype twee kolommen naast elkaar:
  verkocht | transport (vanaf kolom L). Rijen (kwartalen) mag je verwijderen of invoegen en lege rijen tellen niet
  mee: het Model leest de i-de gevulde rij (kolom Bronrij, een matrixformule), dus grafieken en KPI's volgen vanzelf.
* **Woningtypes** – per type één blok van drie kolommen (vanaf kolom C): naam, aantal, koopsom, start bouw, grondtermijn
  (% van de koopsom, bij transport) en een eigen lijst bouwtermijnen (naam, %, bouwkwartaal). Bouwtermijnen mogen als
  % van de aanneemsom (samen 100%) of als % van de koopsom (samen 100% − grondtermijn) worden ingevuld: het model
  schaalt ze naar 100% − grondtermijn (Model, blok `tup` rij 5), zodat grondtermijn + bouwtermijnen altijd de koopsom
  vormen. Elk type kan andere termijnen hebben. Blok 3 per type: extra opbrengsten per woning in euro's (kopersmeerwerk,
  kadastrale kosten/rentes, overige) met een bouwkwartaal, of leeg = bij notarieel transport; het model telt ze mee in de
  opbrengst per woning (Model, blok `vx`), buiten de koopsomknop om.
* **Model** – rekenblad. Kolommen A:L uit de invoer (B = bronrij op Invoer), M:BU model, scenario's, scenariolijn, rente en
  grafiekkolommen, BX:BY hulpcellen (BY8 = aantal periodes, de VBA leest die cel), vanaf CA de typeblokken (verkocht cum,
  transport cum, transport cum up/down/scenariolijn, termijnen, en per type verkocht en getransporteerd per kwartaal voor
  de grafiek, benoemde bereiken `g_v<k>`/`g_t<k>`).
* **PowerPoint** – teksten, tabellen en grafiekblokken voor de dia's; kolom F = vormnaam op de dia. Het cashflowblok heeft
  acht kolommen; de achtste (N) is de scenariolijn, die in sjabloon v11 als reeks op kolom H van het gegevensblad van de
  grafiek op dia 3 staat. Staat de lijn uit, dan luidt de kop '… (uit)' en haalt de macro die reeks uit de grafiek.

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
* Elk scenario kan een eigen jaarrente krijgen (rij *Jaarrente per scenario*); leeg betekent de algemene jaarrente.

### Koopsom per scenario

De knop *Koopsom (VON-prijs)* zet per scenario de koopsom van alle woningen zoveel procent hoger of lager. Grondtermijn
en bouwtermijnen schalen mee, op de momenten van het verkooptempo-model (`model_up` … `model_scn` × (1 + %)); de knop
werkt dus alleen met *Woningtypes en termijnen gebruiken = ja* (een controle op het Dashboard waarschuwt anders). De knop
*Opbrengsten* werkt daarentegen op alle prognose-opbrengsten, ook buiten het model.

### Scenariolijn

De gele lijn in de cashflowgrafiek is een derde scenario met eigen knoppen (kolom *Scenario* naast downside en upside):
verschuiving, uitstel start bouw, opbrengsten %, kosten %. Hij staat los van de band in de scenariografiek, begint bij
het punt 'nu' en is met *Scenariolijn tonen = nee* uit te zetten; de naam staat in de legenda en bij het eindpunt. Het
Model rekent hem in de kolommen `model_scn` … `stand_scn` en het typeblok `tscn`; eindsaldo, dieptepunt, break-even en
rente staan in de hulpcellen en op tab PowerPoint (blok *extra knoppen en scenariolijn*).

In het PowerPoint-sjabloon (v11, gemaakt met `tools/pptx_scenario_line.py`) staat de lijn ook in de cashflowgrafiek op
dia 3: reeks op kolom H van het gegevensblad, legenda-naam = kop H1, dus de naam uit het werkboek.

## Controle

`verify`-stappen die bij het bouwen zijn gedaan: LibreOffice-herberekening zonder foutcellen (behalve de bedoelde `#N/B`
in grafiekkolommen), het oude werkboek wordt met rente 0% cel voor cel gereproduceerd, de rente is nagerekend met een
losse simulatie, en met alle knoppen op nul vallen de scenario's samen met de basis.
