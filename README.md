# Forecast-Builder

Bouwt het werkboek **Cashflow_scenario** (Dashboard · Invoer · Woningtypes · Model · PowerPoint) met Python/openpyxl.
Het werkboek rekent scenario's (downside/upside) uit bovenop je eigen projectcashflow en vult met één knop
het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v14.pptx`).

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
| `tools/pptx_v12_bouwtermijnen.py` | dia 5: bouwtermijnen-tijdlijn (BT_GRAFIEK) naast de tabel; nieuwe dia 6 met de twee verkoop/transport-panelen (v11 → v12) |
| `tools/pptx_v13_banen.py` | dia 5: bouwtermijnen per woningtype als banen (koprijen jaar/kwartaal, per termijn een baan per type), grafiek hoger (v12 → v13) |
| `tools/pptx_v14_daeb.py` | dia 8 'Opbrengsten per woningtype' (PT_GRAFIEK; de macro maakt per type een kopie) en herkleurde dia 5/6: vorige prognose lichtrood (v13 → v14) |

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
de knoppen neer (Dashboard, PowerPoint, Woningtypes). Zet `Kwartaal_Template_cashflow_v14.pptx` naast het werkboek
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
  het aantal in elk blokje, mijlpalen 'uitverkocht' en 'laatste transport') en *bouwtermijnen per woningtype* (tijdlijn
  als planningstabel: bovenaan een koprij met de jaartallen en een koprij met de kwartaalnummers 1-4, daaronder per
  bouwtermijn een baan per woningtype dat hem heeft (dezelfde naam bij meerdere typen = één termijn) met een blokje in de
  typekleur in het kwartaal waarin de termijn vervalt, een lichtrood blokje op de plek volgens de vorige prognose,
  grijs over de gerealiseerde kwartalen en een lege rij tussen de termijnen; chip-cellen met de typekleuren erboven; de
  tijd-as (kwartaalindex jaar × 4 + kwartaal) loopt van 1 januari van het eerste jaar van de periodes, vast bij het
  bouwen, tot het einde van het laatste jaar) en *opbrengsten per woningtype* (één paneel voor het type uit de blauwe
  keuzecel: gestapelde kolommen per kwartaal met grondtermijn, bouwtermijnen en extra's (koop) of de fees per component
  (DAEB), cumulatieve lijn, chips per onderdeel; Model-selectieblok `pt_*`, benoemde bereiken `g_pt_*`). Een grijs vlak
  markeert de gerealiseerde kwartalen.
* **Invoer** – je eigen cashflow (B:J, plakken als waarden) en per woningtype twee kolommen naast elkaar:
  verkocht | transport (vanaf kolom L). Rijen (kwartalen) mag je verwijderen of invoegen en lege rijen tellen niet
  mee: het Model leest de i-de gevulde rij (kolom Bronrij, een matrixformule), dus grafieken en KPI's volgen vanzelf.
* **Woningtypes** – per type één blok van drie kolommen (vanaf kolom C): naam, aantal, koopsom, start bouw, grondtermijn
  (% van de koopsom, bij transport) en een eigen lijst bouwtermijnen (naam, %, bouwkwartaal). Bouwtermijnen mogen als
  % van de aanneemsom (samen 100%) of als % van de koopsom (samen 100% − grondtermijn) worden ingevuld: het model
  schaalt ze naar 100% − grondtermijn (Model, blok `tup` rij 5), zodat grondtermijn + bouwtermijnen altijd de koopsom
  vormen. Elk type kan andere termijnen hebben. Blok 3 per type: extra opbrengsten per woning in euro's (kopersmeerwerk,
  kadastrale kosten/rentes, overige) met een bouwkwartaal, of leeg = bij notarieel transport; het model telt ze mee in de
  opbrengst per woning (Model, blok `vx`), buiten de koopsomknop om. Blok 4: start bouw volgens de vorige prognose (jaar,
  kwartaal) voor de lichtrode blokjes in de bouwtermijnengrafiek. Blok 5: soort (niet-DAEB/DAEB) en fees: bij DAEB (sociale
  huur aan een corporatie, gescheiden koop-/aannemingsovereenkomst) telt het type geen koopsom, grondtermijn, bouwtermijnen of
  extra's, maar fees voor het hele type: twee componenten (AK fee, bijkomende kosten) met een totaalbedrag en elk vijf termijnen
  (mijlpaal, bedrag in € of percentage ≤ 1 van het totaal, kwartaal als '2026 Q4', bouwkwartaal of 'actuals'). Percentages worden
  niet opgeschaald; de kolom 'gepland' en twee Dashboard-controles melden componenten die niet volledig zijn ingepland en
  termijnen zonder geldig kwartaal. De fees staan als banen in de bouwtermijnengrafiek (gegroepeerd per component) en in de
  grafiek per woningtype; de knoppen en de koopsomknop raken ze niet (ze staan in alle vier de modelkolommen gelijk, dus
  scenario's, rente en standen veranderen niet). Een project mag DAEB en niet-DAEB mengen; DAEB-woningen tellen niet mee in
  'verkocht vóór start bouw'.
* **Model** – rekenblad. Kolommen A:L uit de invoer (B = bronrij op Invoer), M:BU model, scenario's, scenariolijn, rente en
  grafiekkolommen, BX:BY hulpcellen (BY8 = aantal periodes, de VBA leest die cel), vanaf CA de typeblokken (verkocht cum,
  transport cum, transport cum up/down/scenariolijn, termijnen, en per type verkocht en getransporteerd per kwartaal voor
  de grafiek, benoemde bereiken `g_v<k>`/`g_t<k>`).
* **PowerPoint** – teksten, tabellen en grafiekblokken voor de dia's; kolom F = vormnaam op de dia. Blokken (kop in rij 7):
  cashflow (G, dia 3, acht kolommen; de achtste (N) is de scenariolijn, in het sjabloon een reeks op kolom H van het
  gegevensblad; staat de lijn uit, dan luidt de kop '… (uit)' en haalt de macro die reeks uit de grafiek), scenario's (P,
  dia 4), bouwtermijnen (AA, dia 5: 108 kolommen = termijn + 107 reeksen (stapel 1: onzichtbaar, grijs, onzichtbaar,
  vorige prognose per type, grijs, onzichtbaar, 16 jaarblokjes; stapel 2: onzichtbaar, huidige planning per type,
  onzichtbaar, 64 kwartaalblokjes), tot 132 rijen, Model!BY96 = aantal rijen, waarden in kwartaalindex; koppen '(uit)' =
  reeks niet in de presentatie), verkoop en transport per type (EF, dia 6: 26 kolommen, twee grafieken lezen hetzelfde
  blok), van verkoop naar omzet (FG, dia 7), opbrengsten per woningtype (FM, dia 8: 8 kolommen, het type uit de keuzecel;
  de macro zet per type de keuzecel, rekent door en vult een kopie van dia 8); tabellen vanaf FV. De bouwtermijnentabel
  staat in het Model (kolommen vanaf GW: raster van alle type × termijn-combinaties (bij DAEB de fee-termijnen) met
  groepsnaam, bedrag, kwartaalindex, rang en positie, daarna de compacte lijst met de koprijen, de banen en de
  stapelsegmenten voor de grafiek), gevolgd door het selectieblok van de grafiek per type.

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

In het PowerPoint-sjabloon (sinds v11, `tools/pptx_scenario_line.py`) staat de lijn ook in de cashflowgrafiek op
dia 3: reeks op kolom H van het gegevensblad, legenda-naam = kop H1, dus de naam uit het werkboek.

### Dia's en sjabloon v14

Dia 1 titel · 2 kerncijfers · 3 cashflow per kwartaal (met scenariolijn) · 4 scenario's · 5 bouwtermijnen per woningtype
(tijdlijn) met de tabel woningtypes · 6 verkoop en transport per kwartaal (twee panelen) · 7 van verkoop naar omzet ·
8 opbrengsten per woningtype (de macro maakt per type met aantal > 0 een kopie van deze dia en haalt dia 8 zelf weg).
De macro *Naar PowerPoint* vult alle teksten, tabellen en grafieken: per grafiek schrijft hij het blok van tab PowerPoint
in het gegevensblad van de grafiek op de dia en zet de reeksen op de juiste kolommen; reeksen met een kop die op '(uit)'
eindigt haalt hij weg (scenariolijn uit; typen zonder termijnen en jaren buiten het bereik in de bouwtermijnengrafiek);
voor de bouwtermijnengrafiek zet hij ook de tijd-as (Model!BY97 t/m BY99, kwartaalindex). Sjabloon v14 is gemaakt met
`tools/pptx_v12_bouwtermijnen.py` (v11 → v12), `tools/pptx_v13_banen.py` (v12 → v13) en `tools/pptx_v14_daeb.py` (v13 → v14).

## Controle

`verify`-stappen die bij het bouwen zijn gedaan: LibreOffice-herberekening zonder foutcellen (behalve de bedoelde `#N/B`
in grafiekkolommen), het oude werkboek wordt met rente 0% cel voor cel gereproduceerd, de rente is nagerekend met een
losse simulatie, en met alle knoppen op nul vallen de scenario's samen met de basis.
