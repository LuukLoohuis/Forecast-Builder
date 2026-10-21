# Forecast-Builder

Bouwt het werkboek **Cashflow_scenario** (Dashboard · Scenario's · Invoer · Woningtypes · CF kosten · Model · PowerPoint) met Python/openpyxl.
Het werkboek rekent scenario's (downside/upside en verkoopscenario's) uit bovenop je eigen projectcashflow en vult met één knop
het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v20.pptx`; v19 werkt ook).

## Wat zit erin

| Map / bestand | Inhoud |
|---|---|
| `build.py` | opdrachtregel: bouwt `out/Cashflow_scenario.xlsx` (en `.xlsm` als er een VBA-bron is) |
| `forecast_builder/layout.py` | vaste indeling: kolommen, rijen, hulpcellen, capaciteit (`N_TYPES = 10`) |
| `forecast_builder/sheets.py` | alle tabbladen en formules |
| `forecast_builder/scenarios.py` | tab Scenario's en de Modelkolommen van de verkoopscenario's |
| `forecast_builder/kosten.py` | tab CF kosten (kosten per post en per kwartaal) en de Modelkolommen voor het meeschuiven van kosten in de scenario's |
| `forecast_builder/data.py` | projectgegevens: inlezen uit een bestaand werkboek (oude of nieuwe indeling) of het ingebouwde voorbeeld |
| `forecast_builder/charts.py` | grafiek-XML voor het Dashboard |
| `forecast_builder/package.py` | nabewerking: grafieken en VBA in het bestand zetten |
| `vba/CashflowNaarPowerPoint_v10.bas` | VBA-module: Naar PowerPoint, Sjabloon controleren, Type invoegen, Type verwijderen, Ophalen uit FO, Vorige prognose uit FO |
| `forecast_builder/vbabuild.py` | schrijft de VBA-module in het .xlsm (vbaProject.bin: MS-CFB + MS-OVBA), zodat importeren niet nodig is |
| `forecast_builder/cache.py` | rekent door met LibreOffice en zet cachewaarden in cellen en grafieken |
| `tools/pptx_rente_card.py` | zet de vijfde KPI-kaart (rente) op dia 2 van het PowerPoint-sjabloon (v9 → v10) |
| `tools/pptx_scenario_line.py` | zet de scenariolijn als zevende reeks (kolom H, amber) in de cashflowgrafiek op dia 3 (v10 → v11) |
| `tools/pptx_v12_bouwtermijnen.py` | dia 5: bouwtermijnen-tijdlijn (BT_GRAFIEK) naast de tabel; nieuwe dia 6 met de twee verkoop/transport-panelen (v11 → v12) |
| `tools/pptx_v13_banen.py` | dia 5: bouwtermijnen per woningtype als banen (koprijen jaar/kwartaal, per termijn een baan per type), grafiek hoger (v12 → v13) |
| `forecast_builder/fo.py` | lezer voor een FO-werkboek (`build.py --fo`); zelfde logica als de macro `UitFOOphalen` |
| `tools/kringverwijzing.py` | statische controle op kringverwijzingen zoals Excel die meldt (formule hangt af van elk cel in de genoemde bereiken, ook bij INDEX); ook als test op het gebouwde werkboek |
| `tools/pptx_v15_pertype.py` | sjabloon v14 → v15: lichtblauwe 'Vorige prognose' (dia 3), vernieuwde grafiek per woningtype (dia 8) |
| `tools/pptx_v16_verkooppct.py` | sjabloon v15 → v16: dia 6 met het cumulatieve percentage als lijn met labels, blok van 30 kolommen, tijdvak vanaf het actuals-jaar |
| `tools/pptx_v14_daeb.py` | dia 8 'Opbrengsten per woningtype' (PT_GRAFIEK; de macro maakt per type een kopie) en herkleurde dia 5/6: vorige prognose lichtrood (v13 → v14) |
| `tools/pptx_v18_scenarios.py` | sjabloon v17 → v18: nieuwe dia 4 'Verkoopscenario's' (kaarten, legenda, VS_GRAFIEK) en drie scenariolijnen in de cashflowgrafiek op dia 3; `--simuleer` doet na wat de macro doet |
| `tools/pptx_v20_daeb.py` | sjabloon v19 → v20: dia 7 met de DAEB-markering (gouden ruit, kolom AE/AF) in VP1/VP2 en drie kaartjes VP_STAT1..3, de grafieken smaller; BT_GRAFIEK toont elk rijlabel; `--simuleer` doet dia 6 en 7 na |
| `tools/pptx_v19_kort.py` | sjabloon v18 → v19: minder tekst (ondertitels van dia 2, 3, 4 en 7, KPI4_SUB, de regel onderaan de scenariokaarten en de voetnoot van dia 4 weg) en de scenariolijnen in contrastrijkere kleuren met een eigen lijnstijl; `--simuleer` doet na wat de macro doet |
| `tools/pptx_v17_kerncijfers.py` | sjabloon v16 → v17: dia 2 'kerncijfers' in het nieuwe ontwerp (tegels, tijdlijnband met de kleine cashflowgrafiek KPI_GRAFIEK, kaarten met balken, grootste lasten); `--simuleer` doet na wat de macro doet |

## Bouwen

```bash
python build.py                                      # voorbeeldproject -> out/Cashflow_scenario.xlsx + .xlsm
python build.py --from Cashflow_scenario.xlsm        # invoer en knoppen uit je huidige werkboek (oude of nieuwe indeling)
python build.py --from oud.xlsm --out map/naam       # eigen uitvoernaam (zonder extensie)
python build.py --geen-vba                           # alleen .xlsx
python build.py --leeg --out out/Cashflow_scenario_leeg   # leeg sjabloon: geen periodes en geen woningtypes, wel de knoppen
```

Het `.xlsm` bevat de macro al: `vba/CashflowNaarPowerPoint_v10.bas` wordt bij het bouwen in `xl/vbaProject.bin`
geschreven (`forecast_builder/vbabuild.py`, op basis van `vba/vbaProject_template.bin` met de documentmodules en
projectinstellingen). Importeren in de VBA-editor is dus niet nodig. Bij het openen zet `KnoppenControleren`
de knoppen neer (Dashboard, PowerPoint, Woningtypes). Zet `Kwartaal_Template_cashflow_v20.pptx` naast het werkboek
(of vul het pad in op tab PowerPoint). Sjabloon v20 maak je van v19 met `python3 tools/pptx_v20_daeb.py v19.pptx v20.pptx`, v19 van v18 met `python3 tools/pptx_v19_kort.py v18.pptx v19.pptx`,
v18 van v17 met `python3 tools/pptx_v18_scenarios.py v17.pptx v18.pptx` (v17 van v16 met `tools/pptx_v17_kerncijfers.py`).
Met sjabloon v18 haalt de macro (v21) de vormen die geen tekst meer krijgen zelf van de dia's.

`--from` begrijpt zowel de oude indeling (vijf types naast elkaar op tab Invoer) als de nieuwe (typeblokken), dus een
bijgewerkt werkboek genereer je later opnieuw zonder gegevens over te typen.

Als LibreOffice (`soffice`) aanwezig is, rekent de build het werkboek door en zet de uitkomsten als cachewaarden in
cellen en grafieken (`forecast_builder/cache.py`), zodat Excel ook in de beveiligde weergave meteen cijfers toont.
Zonder LibreOffice werkt het bestand ook; Excel rekent dan bij het openen (`--geen-cache` slaat de stap over).

Tests: `python3 -m pytest tests/ -q` (de modeltests gebruiken LibreOffice om formules door te rekenen; `tests/test_macro_lo.py` voert de macro 'Ophalen uit FO' uit in LibreOffice (VBA-compatibiliteit, via pyuno) op een nagebouwd FO en vergelijkt het resultaat met de Python-lezer; `tests/test_vba.py` controleert de VBA-module statisch, onder meer op functies die zichzelf aanroepen).

## Indeling van het werkboek

* **Dashboard** – kaarten, drie grafieken en in kolom V alle knoppen: actuals t/m, jaarrente, rente t/m (start bouw of
  hele looptijd), rente ook in de basis (ja/nee), verkooptempo-model aan/uit, per scenario (downside, upside én de
  scenariolijn): verschuiving verkoop/transport, uitstel start bouw, opbrengsten %, kosten %, eigen jaarrente (leeg =
  algemeen) en koopsom % (VON-prijs via het verkooptempo-model); scenariolijn aan/uit en naam;
  norm verkocht bij start bouw; tien controles (oplopende kwartalen zonder gaten, maximaal 60 periodes, rijen zonder jaar,
  jaar als getal, koprij, verkocht/transport = aantal, termijnen); overzicht types. De grafiek *opbrengsten per woningtype*
  heeft een blauwe keuzecel (B120): een type, of 'Totaal (zonder DAEB)' = alle koopwoningen samen; ernaast (W120) wat de
  macro naar PowerPoint zet: 'totaal (zonder DAEB)' (standaard), 'per type' of 'totaal en per type'.
  De grafieken (`forecast_builder/chartxml.py`): *scenario's* (basis met de band tussen downside en upside), *cashflow per
  kwartaal* (opbrengsten, kosten, cumulatieve cashflow, vorige prognose en de gele scenariolijn vanaf het punt 'nu') en
  *verkoop en transport per woningtype* (twee panelen boven elkaar: verkocht en getransporteerd per kwartaal, gestapeld
  per type in een vaste kleur per typeblok, het aantal in elk blokje, het cumulatieve percentage van alle woningen als
  donkerblauwe lijn op een rechteras 0-100 % met het percentage in vet boven de lijn, chips boven het eerste paneel met
  per type het percentage verkocht t/m de actuals, mijlpalen 'uitverkocht' en 'laatste transport'; alleen de kwartalen
  vanaf 1 januari van het actuals-jaar (Model `vp_start`/`vp_n`, benoemde bereiken `g_vk_*`)) en *bouwtermijnen per woningtype* (tijdlijn
  als planningstabel: bovenaan een koprij met de jaartallen en een koprij met de kwartaalnummers 1-4, daaronder per
  bouwtermijn een baan per woningtype dat hem heeft (dezelfde naam bij meerdere typen = één termijn) met een blokje in de
  typekleur in het kwartaal waarin de termijn vervalt, een lichtrood blokje op de plek volgens de vorige prognose,
  grijs over de gerealiseerde kwartalen en een lege rij tussen de termijnen; chip-cellen met de typekleuren erboven; de
  tijd-as (kwartaalindex jaar × 4 + kwartaal) loopt van 1 januari van het eerste jaar van de periodes, vast bij het
  bouwen, tot het einde van het laatste jaar) en *opbrengsten per woningtype* (één paneel voor het type uit de blauwe
  keuzecel: gestapelde kolommen per kwartaal met grondtermijn, bouwtermijnen en extra's (koop) of de fees per component
  (DAEB), cumulatieve lijn, chips per onderdeel; Model-selectieblok `pt_*`, benoemde bereiken `g_pt_*`). Een grijs vlak
  markeert de gerealiseerde kwartalen.
* **Scenario's** – verkoopscenario's naast de Basis (je planning). Per scenario (tot drie) vul je alleen in wat anders is:
  *start bouw* (hele project), *verkoop vanaf* (het verkooptempo schuift mee) en *transport vanaf* (leeg = net zo ver na
  start bouw als in de planning; nooit vóór de verkoop), de *grondtermijn* als % van de koopsom (bv. 50%; leeg = zoals per type
  gepland, de Basis-kolom toont het geplande % gewogen naar koopsom), plus een naam (leeg = automatisch, bv. 'Verkoop Q2 ’28'). Kwartalen
  kies je uit de keuzelijst of typ je ('Q2 2028', '2028 Q2'); een lege kolom staat uit. Het tabblad toont kaarten per scenario
  (max. financiering, dieptepunt, break-even, verkocht bij start bouw tegen de norm, verschil met de Basis), de uitkomst per
  scenario, de grafiek met de cumulatieve cashflow per scenario (een lijn alleen waar hij van de Basis afwijkt, dieptepunt als
  stip) en de opbrengsten per termijn en per kwartaal voor een gekozen scenario en woningtype (zoals 'CF - opbrengsten' in het FO),
  met daaronder de bouwkosten (wat de aannemer krijgt, Woningtypes blok 7; bij alle types een rij per type, bij één type een rij
  per termijn; verschoven met de start bouw van het scenario, gerealiseerde delen blijven in hun kwartaal) en 'bouwtermijnen
  koper − aannemer' per kwartaal en cumulatief vanaf het begin van het project (DAEB-typen tellen daarin niet mee: die hebben
  geen bouwtermijnen van kopers): rood = de bouw is tot dan voorgefinancierd.
  Rekenregels: verkoop per type schuift met het verschil tussen *verkoop vanaf* en de eerste geplande verkoop; transport volgt
  de geplande transportcurve, verschoven, en nooit meer dan er verkocht is; bouwtermijnen en extra's schuiven met de start bouw
  en worden bij transport ingehaald (grond + alle vervallen termijnen); een andere grondtermijn geldt voor de niet-DAEB-woningen
  die na de actuals in transport gaan: de koopsom blijft gelijk, de bouwtermijnen gaan naar rato omlaag (samen weer 100%), dus
  er komt meer bij transport binnen en minder in de bouwkwartalen (per woning koopsom × (grond nieuw − grond gepland) ×
  (1 − vervallen bouwtermijnen / (1 − grond gepland)); werkt met 'Woningtypes en termijnen gebruiken' = ja); kosten: met posten op tab CF kosten schuiven alleen de
  posten op 'start bouw' mee met de start bouw en die op 'verkoop' met de verkoop (zonder posten: alle kosten vanaf de geplande
  start bouw); rente rekent zoals bij de scenariolijn. Een scenario zonder invoer is precies de Basis. **Naar PowerPoint**
  (ja/nee, standaard ja) bepaalt of de ingevulde scenario's op dia 4 komen; zonder ingevuld scenario komt er niets in PowerPoint
  (de macro haalt dia 4 weg). **Lijnen in de cashflowgrafiek** (standaard nee; ja = alle ingevulde scenario's, '1' of '1,3' = alleen
  die) zet de scenario's daarnaast als lijn in de cashflowgrafiek op het Dashboard en op dia 3 (met Naar PowerPoint = ja). Kleuren
  (`layout.VS_KLEUREN`): Basis navy, scenario 1 violet, 2 roze (lange streep), 3 teal (stippel), zodat ze naast de groene en
  rode balken, de lichtblauwe vorige prognose en de gele scenariolijn te onderscheiden zijn. Model: kolommen vanaf LK (per scenario een blok,
  `layout.vs_col`; `gcf` = lijn voor de cashflowgrafiek), hulpcellen `vs_*` en `vs<s>_*` (BY145 e.v.), benoemde bereiken
  `g_vs<s>` en `g_vsw_*`.
* **CF kosten** – kosten per post en per kwartaal, zoals het FO-tabblad 'CF - kosten' (blok 'Onderdelen'): per post de naam,
  *schuift mee met* (start bouw = bouwkosten die met de bouw meeschuiven: aanneemsom, meerwerk, aansluitkosten, reserve na
  oplevering; verkoop = promotie- en verkoopkosten; vast = grond, honoraria, leges, rente, AK), het gerealiseerde bedrag t/m het
  actuals-kwartaal (kolom E) en de prognose per periode (kolom I = periode 1, zoals de rijen op tab Invoer). 'Ophalen uit FO'
  vult de tab en zet de keuze op de naam van de post (een eigen keuze voor een post met dezelfde naam blijft staan). Een
  controleregel vergelijkt elk prognosekwartaal met de kosten op tab Invoer: de Basis blijft altijd je eigen cashflow, de posten
  bepalen alleen wat er in een scenario verschuift (Model: cumulatieve prognose per soort in de kolommen `cksb`/`ckv`; scenario =
  basis − eigen curve + verschoven curve; wat vóór het eerste prognosekwartaal zou vallen komt daarin, wat na de laatste periode
  zou vallen in de laatste). Blok 2 telt per post de kosten t/m het dieptepunt van de Basis; de drie grootste gaan naar dia 2 als
  het Dashboard geen lasten heeft. Blok 3 splitst de aanneemsom uit per type en per kwartaal (uit tab Woningtypes blok 7; wat
  na de laatste periode valt in de laatste, zoals de posten) en vergelijkt die met de post(en) met 'aanneemsom' in de naam; zonder
  zo'n post is de uitsplitsing ter informatie, en bouwkosten zonder bouwkwartaal (of van een type zonder start bouw) geven een
  melding.
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
  huur, gescheiden koop-/aannemingsovereenkomst) telt het type geen koopsom, grondtermijn, bouwtermijnen of
  extra's, maar fees voor het hele type: twee componenten (AK fee, bijkomende kosten) met een totaalbedrag en elk vijf termijnen
  (mijlpaal, bedrag in € of percentage ≤ 1 van het totaal, en het kwartaal waarin het binnenkomt: een kalenderkwartaal als '2026 Q4' of 'Q4 2026', een getal ten opzichte van start bouw van dat type (1 = kwartaal van start bouw, 2 = het kwartaal erna, 0 = het kwartaal ervoor, -1 = twee ervoor; schuift mee met start bouw) of 'actuals' = al ontvangen (telt in het actuals-kwartaal); de cel toont deze uitleg als invoerbericht zodra ze geselecteerd is). Blok 6: per fee-termijn het kwartaal volgens de vorige prognose (zelfde schrijfwijze; leeg = zoals nu, een bouwkwartaal schuift mee met de start bouw van blok 4): de lichtrode blokjes van de fees in de bouwtermijnengrafiek. De tijd-as van die grafiek begint twee jaar vóór de eerste termijn (huidig of vorig), niet vóór de eerste periode (Model BY97; cache.py zet de Excel-grafiek erop, de macro de PowerPoint-grafiek). Percentages worden
  niet opgeschaald; de kolom 'gepland' en twee Dashboard-controles melden componenten die niet volledig zijn ingepland en
  termijnen zonder geldig kwartaal. De fees staan als banen in de bouwtermijnengrafiek (gegroepeerd per component) en in de
  grafiek per woningtype; de knoppen en de koopsomknop raken ze niet (ze staan in alle vier de modelkolommen gelijk, dus
  scenario's, rente en standen veranderen niet). Een project mag DAEB en niet-DAEB mengen; DAEB-woningen tellen niet mee in
  'verkocht bij start bouw'. Blok 7: bouwkosten per termijn (wat de aannemer krijgt): per termijn de naam, het bedrag in € voor
  alle woningen van het type en het bouwkwartaal t.o.v. start bouw, zoals de bouwtermijnen (0 of lager = vóór start bouw, bv.
  al gerealiseerd); 'Ophalen uit FO' vult het uit de termijnblokken op 'CF - kosten'. Alleen voor de uitsplitsing: de kosten in
  de cashflow komen uit tab Invoer en tab CF kosten.
* **Model** – rekenblad. Kolommen A:L uit de invoer (B = bronrij op Invoer), M:BU model, scenario's, scenariolijn, rente en
  grafiekkolommen, BX:BY hulpcellen (BY8 = aantal periodes, de VBA leest die cel), vanaf CA de typeblokken (verkocht cum,
  transport cum, transport cum up/down/scenariolijn, termijnen, en per type verkocht en getransporteerd per kwartaal voor
  de grafiek, benoemde bereiken `g_v<k>`/`g_t<k>`).
* **PowerPoint** – teksten, tabellen en grafiekblokken voor de dia's; kolom F = vormnaam op de dia. Blokken (kop in rij 7):
  cashflow (GT, dia 3, elf kolommen; de achtste is de scenariolijn, in het sjabloon een reeks op kolom H van het
  gegevensblad, de laatste drie zijn de verkoopscenario's (kolom I..K, alleen met 'Lijnen in de cashflowgrafiek' = ja); staat
  een lijn uit, dan luidt de kop '… (uit)' en haalt de macro die reeks uit de grafiek), verkoopscenario's (HF, dia 4: negen kolommen = jaarlabel, Basis, scenario 1..3,
  dieptepunt Basis en scenario 1..3; tijdvak vanaf het actuals-jaar), scenario's (P, dia 5), bouwtermijnen (AA, dia 6: 108 kolommen = termijn + 107 reeksen (stapel 1: onzichtbaar, grijs, onzichtbaar,
  vorige prognose per type, grijs, onzichtbaar, 16 jaarblokjes; stapel 2: onzichtbaar, huidige planning per type,
  onzichtbaar, 64 kwartaalblokjes), tot 132 rijen, Model!BY96 = aantal rijen, waarden in kwartaalindex; koppen '(uit)' =
  reeks niet in de presentatie), verkoop en transport per type (EF, dia 7: 30 kolommen = kwartaal, verkocht en
  getransporteerd per type, totalen, mijlpalen, realisatie, cumulatief % en label voor verkocht en voor transport; alleen
  de periodes vanaf het actuals-jaar (Model!BY131/BY132); twee grafieken lezen hetzelfde blok), van verkoop naar omzet (FK,
  dia 8), opbrengsten per woningtype (FQ, dia 9: 10 kolommen = kwartaal, zes reeksen, totaal per kwartaal (label boven de stapel), eindpunt (label met het totaal), realisatie; alleen de periodes vanaf twee jaar vóór de eerste opbrengst (Model!BY127/BY128); het type uit de keuzecel;
  de macro zet de keuzecel, rekent door en vult een kopie van dia 9: volgens Dashboard W120 één dia met 'Totaal (zonder DAEB)'
  (keuzecel 0, standaard), een dia per type, of beide); tabellen vanaf GB. De bouwtermijnentabel
  staat in het Model (kolommen vanaf FX: raster van alle type × termijn-combinaties (bij DAEB de fee-termijnen) met
  groepsnaam, bedrag, kwartaalindex, rang en positie, daarna de compacte lijst met de koprijen, de banen en de
  stapelsegmenten voor de grafiek), gevolgd door het selectieblok van de grafiek per type.

### Koppeling met het FO (Financieel Overzicht)

Een FO-werkboek van Heijmans Vastgoed heeft per project dezelfde tabbladen. De knop **Ophalen uit FO** (tab Invoer, .xlsm;
macro `UitFOOphalen`) en `python3 build.py --fo FO.xlsx` (module `forecast_builder/fo.py`) lezen hetzelfde, op labels:

| FO-tabblad | Wat | Naar |
|---|---|---|
| `1. Cashflow` | koprij 'Jaar' (ergens in A1:F30); daarachter Q, kosten, % kosten, omzet, % omzet, CF, CF x 1000, CF vorig; 'Totaal:' onder de tabel | tab Invoer B:J; Dashboard FO-cellen (totalen) |
| `FO - actuals` | laatste boekjaar/kwartaal met een boeking (kolommen Boekjaar, Q, Waarde/TrVal) | Dashboard 'Actuals t/m' |
| `CF - opbrengsten` | per type: blok 'Termijnen k - naam' (of 'Termijnen <300K' met de naam in de cel ernaast) met '# won' erboven (bouwplanning: aantal woningen per termijn per kwartaal), blok 'Omzet HVG Termijnen' (Koopsom = grond + ontwikkeling per woning; termijnen in %, € en € per woning; kopersmeerwerk, kadastrale kosten, overige) en blok 'Verkooptempo' (verkocht per kwartaal, kolom 'actuals' = al verkocht) | tab Woningtypes: naam, aantal, koopsom per woning (grond + termijnen, excl. btw, incl. KAO-opties), grondtermijn %, start bouw, bouwtermijnen (naam, %, bouwkwartaal; een termijn die het FO over meer kwartalen verdeelt wordt gesplitst in '(1/2)', '(2/2)'; komen er dan meer dan tien regels, dan gaan de kleinste delen op in het grootste deel van dezelfde termijn, in de macro en in Python), extra's per woning (kwartaal uit de planning, of bij transport); tab Invoer: verkocht en transport per kwartaal (transport = de tellingen op de rij Koopsom) |
| `CF - kosten`, termijnblokken | per type een blok met dezelfde titel als op 'CF - opbrengsten' ('Termijnen ...', '# won' in de rij erboven): per termijn het bedrag per woning (kolom E) en het aantal woningen per kwartaal, tot de rij 'Aanneemsom' | tab Woningtypes blok 7: termijn, bedrag (per woning × woningen) en bouwkwartaal; wat t/m actuals ligt is één deel in het actuals-kwartaal ('… (gerealiseerd)'), elk later kwartaal een eigen deel ('(1/3)' …); meer dan tien: kleinste latere delen samengevoegd (nooit bij het gerealiseerde deel), daarna vallen de laatste weg; dubbele bloktitels: het n-de type met die titel krijgt het n-de blok |
| `CF - kosten` | blok 'Onderdelen' (label in kolom B; kwartalen Q1..Q4 met de jaartallen erboven): per post de naam, het gerealiseerde bedrag (kolommen onder 'Actuals', ook samengevoegd; anders de kolom vóór het eerste kwartaal; kwartalen t/m actuals tellen mee) en de prognose per kwartaal, tot de eerste rij zonder label of 'Totaal' | tab CF kosten (maximaal 20 posten; bedragen na de laatste periode van '1. Cashflow' in de laatste periode) |
| idem, DAEB | in het omzetblok van een type: een rij met een fee-label ('AK fee', 'Bijkomende kosten'; 'Onvoorzien' wordt overgeslagen) en alleen een totaal in € (geen %, geen kwartalen) is een fee-component; elke rij daarna met een % of € én een kwartaal of 'actuals' is een termijn van die component, wat het label ook is ('na akkoord SO', 'bij start bouw', …) | blok 5 (soort DAEB, componenten, termijnen; koopsom, bouwtermijnen en extra's leeg); transport = verkoop als het FO geen transport voor DAEB heeft |
| idem, los DAEB-blok | onderaan, een titel met 'DAEB' ('Opbrengsten DAEB') op een koprij met 'actuals' en Q1..Q4: boven de koprij per component de naam (kolom C, bv. 'AK', 'Bijkomende kosten ex. AK') en het totaal (kolom E); daaronder per termijn een rij ('AK fee na akkoord SO', % in C, het bedrag in 'actuals' of in een kwartaal), groepen gescheiden door een lege rij, tot 'Totaal'; 'Startbouw' boven een kwartaal | een eigen DAEB-type (laatste type): blok 5 met per termijn de mijlpaal (zonder het voorvoegsel van de component), het bedrag uit het FO en het kwartaal ('2027 Q2', of 'actuals' = al ontvangen); een termijn bij de component met hetzelfde sleutelwoord (eerste woord dat geen 'fee' is), anders die van zijn groep; 'Onvoorzien' en subtotalen tellen niet mee; een correctie van 1 euro of minder gaat op in het vorige deel; meer dan vijf delen per component: eerst al het ontvangene in één rij, dan hetzelfde kwartaal samen, dan gesplitste delen, de rest bij de laatste rij (geen bedrag kwijt); start bouw van 'Startbouw'; bestaat er al een DAEB-type met fees uit een omzetblok, dan alleen een melding |
| `Verkooptempo` (tabblad) | de rij die bij geen woningtype hoort (de corporatie; 'Totaal' en lege rijen tellen niet; bij meer zulke rijen die met 'DAEB' in de naam of de eerste na de koopwoningen): aantal (kolom 'Aantal') en verkoop per kwartaal | aantal woningen en verkocht van het losse DAEB-type; transport = verkoop |


Vorige prognose: vóór het leegmaken vraagt de knop of de planning die in het werkboek staat (start bouw per type, kwartalen van de fee-termijnen) bewaard moet worden als vorige prognose (blok 4 en 6; typen op positie met een melding als de naam verschilt, fee-termijnen op mijlpaal). De knop **Vorige prognose uit FO** (tab Woningtypes, macro `VorigePrognoseUitFO`) haalt beide uit een ouder FO (typen op naam, anders positie; fee-termijnen op mijlpaal binnen dezelfde component); in Python: `build.py --fo FO.xlsx --fo-vorig FO_vorig.xlsx` (`fo.voeg_vorige_prognose_toe`). Het FO zelf bevat de vorige planning niet ('FO - vorig' en 'FO - huidig' zijn SAP-budgetexports; alleen 'CF vorig kwartaal' op '1. Cashflow' gaat naar Invoer J).

Tab CF kosten heeft een eigen controle: het gerealiseerde bedrag (kolom E) en elk prognosekwartaal tegen tab Invoer, en bedragen in kolommen na de laatste periode (die tellen niet mee). Ligt het dieptepunt van de Basis al in de realisatie, dan kan het FO de kosten niet per post tot dat kwartaal splitsen (het geeft één gerealiseerd bedrag per post): de tab en de kaart op dia 2 heten dan 'gerealiseerd t/m <actuals>' en 'GEREALISEERDE LASTEN'.

De Dashboard-controle 'Invoer sluit aan op het FO' vergelijkt de totalen van tab Invoer met de totalen uit '1. Cashflow';
de dekking (W13) laat zien hoeveel van de opbrengsten het verkooptempo-model verklaart. De knoppen op het Dashboard blijven bij
ophalen staan; alleen 'Actuals t/m' wordt gezet (en alleen als het FO een tabblad 'FO - actuals' heeft). Meldingen (tabblad
ontbreekt, meer termijnen dan er ruimte is, overgeslagen rijen) komen in de afsluitende melding van de macro en in
`p.fo['waarschuwingen']` (build.py drukt ze af); de Python-lezer voegt gesplitste delen samen als een type na het splitsen
meer dan 10 bouwtermijnen zou krijgen, de macro laat de rest weg met een melding. Staat het FO al open, dan gebruikt de macro
dat werkboek (en sluit het niet); de macro werkt ook op Excel voor Mac (geen Scripting.Dictionary, bestandskeuze zonder filter).

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

### Dia's en sjabloon v20

**Projectnaam**: Dashboard W10 (geel). De macro zet hem op de titeldia (COVER_PROJECT, met het kwartaal in COVER_KWARTAAL) en in
de voettekst van alle dia's (VOET_TEKST in de dia-indeling, 'PROJECT  |  Q3 2026'); ook in de bestandsnaam van de presentatie. Is
hij leeg, dan vult Ophalen uit FO hem uit de bestandsnaam van het FO ('20261001_FO_Project_2_Q4.xlsx' -> 'Project 2').
**Dia 6 / 7**: de bouwtermijnen vullen dia 6 (drie koprijen: jaartallen, lege rij, kwartalen; een extra lege rij vóór de kopregel
DAEB-fees); de tabel woningtypes staat niet meer in het sjabloon (de macro meldt hem dan ook niet als ontbrekend).

**Veel DAEB** (meer dan de helft van de woningen, Model `vp_daeb_apart`): de verkoop- en transportgrafiek (Dashboard en dia 7)
tonen balken en cumulatief % alleen voor de koopwoningen; DAEB staat als gouden ruit in het eerste DAEB-kwartaal met het label
'DAEB · 187 woningen in één keer' (of 'vanaf Q1 ’27'; kolommen `dvm`/`dtm`, PowerPoint-blok kolom AE/AF, anders '(uit)'). Titel en
kaartjes van dia 7 (VP_STAT1..3) gaan dan over koop, DAEB en het totaal; anders verkocht, getransporteerd en nog te verkopen.
**Bouwtermijnen met DAEB-fees** (dia 6 en Dashboard): titel 'Bouwtermijnen Q3 ’28 → Q2 ’30 · DAEB-fees €7,4 mln vanaf Q3 ’26'; twee
groepen met een kopregel en een lege rij ('KOOP · BOUWTERMIJNEN · 3 TYPES', 'DAEB · FEES · €7,4 MLN'), eerst koop; het
termijnlabel bij de middelste typestrook; elke fee als dikke gouden balk (reeks `dfee`, twee rijen) met rechts het bedrag (reeksen
`fgap` + `fl1..fl10`, reeksnaam = bedrag) en een lege rij eronder; legenda (■ per type, ◆ DAEB-fee) als ondertitel. De macro zet
de lettergrootte van de rijlabels naar het aantal rijen.
**Dia 7** zoals de preview: legenda VP_LEGENDA, grafiektitels 'Verkocht per kwartaal · koop (36)', % rechts van het punt, een
gestippelde gouden lijn onder de DAEB-ruit en een gekleurde accentstreep per kaartje.


Dia 1 titel · 2 kerncijfers (tegels, tijdlijnband met kleine cashflowgrafiek, kaarten met balken) · 3 cashflow per kwartaal (met scenariolijn en verkoopscenario's) · 4 verkoopscenario's · 5 scenario's · 6 bouwtermijnen per woningtype
(tijdlijn) met de tabel woningtypes · 7 verkoop en transport per kwartaal (twee panelen) · 8 van verkoop naar omzet ·
9 opbrengsten per woningtype (volgens Dashboard W120 'Naar PowerPoint': standaard één dia met het totaal van alle
koopwoningen zonder DAEB, of per type met aantal > 0 een kopie, of eerst het totaal en dan per type; dia 9 zelf gaat weg).
De macro *Naar PowerPoint* vult alle teksten, tabellen en grafieken: per grafiek schrijft hij het blok van tab PowerPoint
in het gegevensblad van de grafiek op de dia en zet de reeksen op de juiste kolommen; reeksen met een kop die op '(uit)'
eindigt haalt hij weg (scenariolijn uit; typen zonder termijnen en jaren buiten het bereik in de bouwtermijnengrafiek);
voor de bouwtermijnengrafiek zet hij ook de tijd-as (Model!BY97 t/m BY99, kwartaalindex). Sjabloon v16 is gemaakt met
`tools/pptx_v12_bouwtermijnen.py` (v11 → v12), `tools/pptx_v13_banen.py` (v12 → v13), `tools/pptx_v14_daeb.py` (v13 → v14) en `tools/pptx_v15_pertype.py` (v14 → v15: 'Vorige prognose' op dia 3 lichtblauw; dia 8 met totaal per kwartaal boven de stapel, eindpunt met het totaal en een tijdvak vanaf twee jaar vóór de eerste opbrengst) en `tools/pptx_v16_verkooppct.py` (v15 → v16: dia 6 met het cumulatieve percentage verkocht/getransporteerd als lijn met labels, blok van 30 kolommen, alleen de kwartalen vanaf het actuals-jaar).

Dia 2 (kerncijfers, sjabloon v17 via `tools/pptx_v17_kerncijfers.py`) bestaat uit drie tegels (max. financieringsbehoefte,
break-even, projectresultaat, elk met het verschil met de vorige prognose en 'was …'), een tijdlijnband met de kleine
cashflowgrafiek `KPI_GRAFIEK` (cum. cashflow, vorige prognose als lichtblauwe stippellijn, stippen met label voor nu, dieptepunt,
break-even en eindsaldo, grijs vlak voor de gerealiseerde kwartalen, de stand per actuals-kwartaal als nu-blok bij de stip) en
vier kaarten: verkocht bij start bouw (balk met normstreep), verkocht / getransporteerd (twee balken op het aantal woningen),
marge en de grootste lasten. *Verkocht bij start bouw* telt per type de woningen die t/m het kwartaal van start bouw verkocht
zijn (dat kwartaal telt mee; DAEB telt niet mee), tegen de norm op het Dashboard. De margekaart toont bij een project met DAEB
twee marges: *niet-DAEB* = projectresultaat / (opbrengsten − fees DAEB), want DAEB heeft geen marge, en *totaal* =
projectresultaat / opbrengsten; zonder DAEB één marge (hulpcellen `opbr_daeb`, `marge`, `marge_nd`, `marge_split`, `marge_tekst`).
De *grootste lasten* vul je zelf in op het Dashboard (sectie 'GROOTSTE LASTEN (dia 2)': kop van de kaart en tot drie posten met
naam en bedrag in €, bv. grond, honoraria, aansluitkosten; het Dashboard toont het aandeel in de kosten). Laat je ze leeg, dan
komen de drie grootste posten van tab CF kosten op de kaart, met hun kosten t/m het dieptepunt van de Basis (kop 'LASTEN T/M
DIEPTEPUNT'; hulpcellen `kl_bron`, `kl<i>_naam`, `kl<i>_bedrag`; lange namen worden ingekort). De kaart toont ze met een balk ten
opzichte van de hoogste post; zonder lasten en zonder posten haalt de macro de kaart weg en verdeelt hij de drie overige kaarten
onder de tegels (`SchikDia2`). Tab PowerPoint heeft voor dia 2 naast de teksten balkregels (vormnaam met `_BALK` of `_NORM`,
kolom D = aandeel 0..1; `VulBalken` zet de breedte op aandeel × breedte van het spoor met dezelfde naam op `_SPOOR`) en het blok
van de grafiek (kolom GK: jaarlabel, cum. cashflow, vorige prognose, nu, dieptepunt, break-even, eindsaldo, €0; de laatste zes als
één-punts-reeksen met de reeksnaam als label, '(uit)' = reeks niet in de presentatie). Na het vullen zet `SchikGrafiekDia2` de
waarde-as op de hulpcellen `kc_min`/`kc_max`, het grijze vlak t/m het actuals-kwartaal en het nu-blok ernaast (links van de
stip, of rechts als het links niet past); de x van een kwartaal volgt uit het vaste tekenvlak (`layout.KC_PLOT`). Met een ouder
sjabloon (v15/v16) werkt de knop nog: de bandbreedtekaart (KPI5) en de voetnoot gaan weg en de overige kaarten verdelen de
breedte; de grafiek en de balken ontbreken dan (gemeld als 'niet gevonden').

Dia 4 'Verkoopscenario's' (sjabloon v18 via `tools/pptx_v18_scenarios.py`) komt alleen in de presentatie als er op tab
Scenario's een scenario is ingevuld en *Naar PowerPoint* op ja staat (Model `vs_pp_aan`); anders haalt de macro de dia weg en
staan de scenariolijnen in de cashflowgrafiek op '(uit)' (die hebben daarnaast hun eigen schakelaar, standaard uit). De dia heeft een kopje, een titel over het scenario met de diepste
financiering, een kaart per scenario (Basis en ingevulde scenario's: max. financiering, dieptepunt, break-even, verkocht bij
start bouw tegen de norm, verschil met de Basis), een legenda en de grafiek `VS_GRAFIEK` (blok HF). `SchikDiaScenarios` haalt
lege kaarten en legenda-regels weg en verdeelt de breedte; `SchikGrafiekVS` zet de waarde-as op de hulpcellen
`vs_min`/`vs_max`/`vs_stap` en het grijze vlak t/m het actuals-kwartaal (vast tekenvlak `layout.VS_PLOT`). Instellingen van
tab PowerPoint staan sinds v18 op C130/D130 (sjabloon) en D131 (naam).

## Controle

`verify`-stappen die bij het bouwen zijn gedaan: LibreOffice-herberekening zonder foutcellen (behalve de bedoelde `#N/B`
in grafiekkolommen), het oude werkboek wordt met rente 0% cel voor cel gereproduceerd, de rente is nagerekend met een
losse simulatie, en met alle knoppen op nul vallen de scenario's samen met de basis.
