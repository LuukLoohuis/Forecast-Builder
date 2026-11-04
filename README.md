# Forecast-Builder · FO-werkbestand

Eén Excel-bestand voor het kwartaal-FO. Er is één vast, leeg **sjabloon** voor elk project. Een lopend project zet je
over met de knop **Oud FO inlezen**; die zet het oude FO om naar de vaste opzet. Daarna werk je alleen in het werkbestand,
en de PowerPoint komt uit hetzelfde bestand.

| Bestand | Wat |
|---|---|
| `output/FO_werkbestand_sjabloon.xlsm` | het lege sjabloon: begin hier voor elk project |
| `output/FO_Parkzoom_werkbestand.xlsm` | voorbeeld: het fictieve demo-FO omgezet |
| `output/Kwartaal_Template_cashflow_v21.pptx` | PowerPoint-sjabloon (lay-out van v20, met namen voor de vormen die gevuld worden) |
| `docs/werkinstructie.md` | kwartaal-FO in stappen |
| `input/` | demo-FO, scenariomodel, PowerPoint v20 |
| `vba/` | broncode van de macro's |
| `tools/` | bouw- en controlescripts |

## Tabbladen (12 zichtbaar)

| Tabblad | Wat je doet / wat erop staat |
|---|---|
| `Start` | project, oplevering, knoppen, status van de invoer, alle controles, teksten voor dia 10 |
| `FO - vorig`, `FO - huidig`, `FO - actuals`, `Prijslijst dump` | exports plakken of inlezen (SAP FO-overzicht van vorige en huidige periode, SAP actuals, Salesforce) |
| `Prijslijst` | dashboard (zie hieronder): kaarten, woningtypes, grafieken, te boeken in SAP, aansluiting per type met het termijnschema, en onderaan de prijslijst per bouwnummer |
| `CF - opbrengsten` | per woningtype: planning bouwtermijnen, verkoop en transport in aantallen (verkooptempo en tabel Transport), grondprijsaanpassing; kadastrale kosten en rente (handmatig); overige opbrengsten, correctie, DAEB |
| `CF - kosten` | bouwtermijnen kosten per type (eigen planning HWB, standaard gelijk aan de opbrengsten) en per kostenpost Handmatig, Voorstel of Transport |
| `Dashboard`, `Scenario's` | cashflow, kerncijfers, grafieken en scenario's: rekenen automatisch mee |
| `2. Ontw PEW` | PEW huidig/vorig uit SAP, meest recente FD (invullen), toelichting |
| `CF - hulp VENA` | de definitieve cashflow per maand, in het VENA-format; per maand het kwartaal en de status (actuals / nu / prognose), het huidige kwartaal groen, en in B6 bij welke maand het huidige kwartaal hoort |

Verborgen: de rekenbladen van het scenariomodel (`Invoer`, `Woningtypes`, `Scenario kosten`, `Model`), `PowerPoint`
(welke tekst naar welke vorm), `Vorige prognose`, `Controle oud FO` (verschijnt na Oud FO inlezen) en `Behind the scens`
(koppeling CODA-post → groep).

Vervallen ten opzichte van het oude FO: `Aansluiting opbrengsten` (nu op `Prijslijst`), `Verkooptempo`, `Woningtypes` als
invoer, `1. Cashflow`, `3. Restant bdgt`, `Notities FO gesprek`, `Toelichting onderlegger`, `1. Ontw businessc FD` en
`5. K&R businessc`.

## Prijslijst als dashboard

- **Kaarten bovenaan:**
  - aantal woningen;
  - opbrengsten volgens de prijslijst;
  - SKO en het verschil met de prijslijst;
  - te boeken in SAP;
  - aansluiting (OK of het aantal types met FOUT);
  - kans grondtermijn: eerder ontvangen, en de max. financiering huidig en met kans.
- **1 · Woningtypes:** één regel per type, in vier kleurgroepen: uit de prijslijst (grijs), handmatig (oranje, overschrijft), gebruikt in de cashflow (blauw) en kans (groen). Met een totaalregel en het grond % als balkje.
- **Type of blok in prijslijst** (kolom C): hiermee horen bouwnummers uit de prijslijst per bouwnummer bij een woningtype. Je kunt invullen:
  - een waarde uit de kolom Type (bv. `A`) of uit de kolom Termijnen / blok (bv. `4A`);
  - meerdere waarden met `;` ertussen (`4A;4B`);
  - `*` voor alle bouwnummers.

  Oud FO inlezen neemt de waarde over uit de SUMIF-formules van het oude tabblad Aansluiting opbrengsten.
- **Regels zonder koppeling met de prijslijst** (VON verhogingen, opbrengsten commercieel en parkeren): de kolom Prijslijst is daar geel; vul het bedrag in. Oud FO inlezen neemt het bedrag uit het oude FO over.
- **Grafieken:** prijslijst en SKO per type, en het grond % huidig en met kans.
- **Te boeken in SAP:** naast de grafieken staan de correcties per SAP-nummer en kostendrager (uit de kolom Corr). In de aansluitblokken toont kolom K per regel op welk SAP-nummer (en kostendrager) de correctie geboekt wordt.
- **2 · Aansluiting per woningtype:** SKO, prijslijst, actueel budget, Corr, PEW, verschil en status. De status per type staat in de kopregel, het termijnschema ernaast.

## Weergave

- **Jaarvakken:** in CF - opbrengsten en CF - kosten staat aan het begin van elk jaar een lijn, en boven elk kwartaal het jaartal (Q1 vet).
- **Start bouw en oplevering:** vul je in op `Start` (rij 7 en 9). Bovenaan CF - opbrengsten (rij 4) en CF - kosten (rij 2) staat "Start bouw" of "Oplevering" boven het kwartaal, en die kolom is gekleurd (groen en blauw). Oud FO inlezen neemt ze over van de markeringen in het oude FO.
- **Inklappen:**
  - per woningtype is het invoerblok een groep; woningtypes zonder woningen zijn ingeklapt;
  - de rekenblokken (aantallen per termijn en €) zijn standaard ingeklapt;
  - de kwartalen vóór het kwartaal van de laatste actuals zijn een ingeklapte kolomgroep; met + klap je ze open. Het kwartaal van de laatste actuals (vaak het lopende kwartaal) blijft zichtbaar.

  De macro werkt dit bij na Oud FO inlezen, Nieuw kwartaal en het inlezen van actuals.
- **Verkoop:** in het verkooptempo vul je per woningtype het aantal verkochte woningen per kwartaal in; "verkocht %" in het woningtypeblok rekent daaruit.
  - Kolom 1 is verkocht t/m de exportdatum van de prijslijst (Salesforce, Prijslijst C1).
  - Kwartalen vóór het kwartaal van die export tellen niet nog eens mee; die verkopen zitten al in kolom 1.
  - Verkoop en transport mogen per woningtype verschillen: transport kan later dan verkoop.
- **Transport:** onder het verkooptempo staat de tabel Transport. Daar vul je per woningtype het aantal woningen in dat per kwartaal getransporteerd wordt; "getransporteerd %" in het woningtypeblok rekent daaruit.
  - Getransporteerd kan niet vóór verkocht; wat op verkoop wacht, schuift mee.
  - De omzet schuift mee: de grondtermijn en de bouwtermijnen die bij transport al vervallen zijn, komen in het transportkwartaal binnen.
  - De kosten van de bouwtermijnen volgen de bouwplanning (HWB), niet het transport. Een kostenpost met methode Transport (bijvoorbeeld de grond) schuift wel mee.
- **Alle woningen:** onder de transporttabel per kwartaal het % verkocht en getransporteerd van alle woningen, en cumulatief.
- **Bedragen per woningtype:** onder de transportregel staan in € per kwartaal: de grondtermijn (koopsom grond), de bouwtermijnen, kopersmeerwerk 25% (met de eerste termijn, start bouw), kopersmeerwerk 75% (met de laatste termijn, oplevering) en het totaal. Ze schuiven mee met het transport; de kopersmeerwerkregels zijn vooraf ingevuld. Bij elke termijn staan het %, het bedrag (alle woningen) en het bedrag per woning.
- **Kleuren:** kopregels donkerblauw, geel = invoer (alleen bij termijnen die er zijn), totaalregels lichtblauw en vet.
- **Grafiekperiode:** op het Dashboard (V87–Z91) vul je "Grafiek vanaf" en "Grafiek t/m" in (jaar · kwartaal).
  - Leeg = automatisch: de grafiek loopt mee met CF - opbrengsten en CF - kosten, vanaf het eerste kwartaal met opbrengsten of kosten (minstens €500), maar niet vóór Q1 van het eerste jaar (cel D3), tot één kwartaal na het laatste.
  - Haal je in de CF-tabbladen een kwartaal leeg, dan schuift de grafiek mee.
  - De cashflowgrafiek en de andere grafieken over de hele looptijd volgen deze periode.
- **Vorige en huidige prognose:** op het Dashboard (V93 en verder) per periode van de grafiek: de cashflow vorige en huidig per periode, cumulatief, en het verschil.
- **Tabbladkleuren:** Start donkerblauw, exports grijs, invoer en rekenwerk geel (Prijslijst, CF - opbrengsten, CF - kosten, 2. Ontw PEW), uitkomsten groen (Dashboard, Scenario's, CF - hulp VENA).
- **Fouten vooraan:** kolom A op CF - opbrengsten en CF - kosten toont FOUT of LET OP voor de rij; cel C1 telt ze.
- **Kans grondtermijn:** vul op `Prijslijst` in kolom Y een ander grond-% van de koopsom in (bijvoorbeeld 50%).
  - De kans geldt alleen voor woningen die nog getransporteerd worden. Bij al getransporteerde woningen ligt de verdeling grond/bouwtermijnen vast (CF - opbrengsten, kolom C van de kansregels).
  - Per type staat het netto effect per kwartaal: meer grondtermijn bij transport, minder bouwtermijnen daarna.
  - De totaalregels tonen:
    - de grondtermijn huidig en met kans;
    - de extra opbrengst grondtermijn (eerder ontvangen);
    - de minder bouwtermijnen (later);
    - het netto effect per kwartaal en cumulatief (de hoogste stand).
  - Over de hele looptijd is het netto 0: de kans verschuift opbrengsten naar voren.
  - Het Dashboard (V82–Z86) toont de max. financiering huidig en met kans, en de extra grondtermijn.
  - De kans zit niet in de cashflow zelf; het is een doorkijk.
- **Parkeren en commercieel:** onder "Overige opbrengsten in €" op CF - opbrengsten:
  - kolom D toont het bedrag uit de aansluiting (PEW, alle woningtypes);
  - in kolom C kies je de verdeling:
    - **bouwtermijnen** (standaard): naar rato van de gefactureerde bouwtermijnen van alle woningtypes. Wordt in een kwartaal 10% van de bouwtermijnen gefactureerd, dan komt 10% van het bedrag binnen;
    - **handmatig**: de bedragen in € die je in de gele regel eronder invult;
  - kolom E en de statuskolom tonen wat nog te verdelen is (LET OP in kolom A);
  - Oud FO inlezen zet de overige opbrengsten van het oude FO in de handmatige regel; die gaan dan altijd als handmatig mee.
- Op `Prijslijst` zijn ongebruikte woningtypes ingeklapt (na Oud FO inlezen).
- **Koopsom en grond %:** grond % = grond / (grond + aanneemsom), zonder KAO; de kans in kolom Y rekent met dezelfde basis. Per woningtype op `Prijslijst` (kolom X) en bovenaan elk woningtypeblok op CF - opbrengsten.
- **Aantallen** staan zonder punt of komma erachter.

## PowerPoint

- Op de dia "Update PEW" komt in het lege vak (`PEW_TABEL` in sjabloon v21) het overzicht van `2. Ontw PEW` als afbeelding: kosten, opbrengsten en resultaat, met de toelichting.
- Daaronder staat de **top 3 kosten** van het reviewkwartaal.

## Top 3 kosten

`Start`, blok 7: de drie kostensoortomschrijvingen (FO - actuals, kolom E) met de hoogste actuals in het kwartaal van de
laatste actuals, met het bedrag en het aandeel in de kosten van dat kwartaal. Hulpkolom S op FO - actuals telt per
omschrijving op.

## Vaste opzet

- 6 woningtypes met elk 10 termijnen, en 40 kwartalen vanaf Q1 van het startjaar (cel D3 op `CF - opbrengsten`).
- Kolom G is "actuals": alles van vóór het eerste kwartaal.

**Opbrengsten per woningtype**
- De koopsom komt van `Prijslijst`. De grondtermijn = aantal getransporteerd × grond per woning (+ de grondprijsaanpassing).
- Bouwtermijnen = aantal woningen dat de termijn haalt (planning), begrensd door het transport, × % × (aanneemsom + KAO) per woning.
- Kopersmeerwerk: 25% met de eerste termijn (start bouw), 75% met de laatste termijn (oplevering).
- Kadastrale kosten en aanvullende rente: aantal woningen per kwartaal, met de hand.
- Verkocht t/m nu = status "Verkocht" of "Getransporteerd" in de prijslijst.
- Getransporteerd kan niet vóór verkocht: schuift de verkoop op, dan schuift het transport mee.

**Kosten**
- Per type budget aanneemsom (excl. KAO), KAO, % per termijn en de planning (standaard die van de opbrengsten).
- Aanneemsom en Kopersmeerwerk volgen de termijnen.
- Bij Voorstel of Transport gaat wat je in de postregel invult voor; de rest van het budget wordt verdeeld over de kwartalen die leeg zijn of op 0 staan.
- Voorstel verdeelt t/m de oplevering (tab Start). Is die leeg of al voorbij, dan t/m het laatste kwartaal met opbrengsten (zie de regel "verdelen t/m" in het voorstelblok).
- Naast elke post: "Verdeeld" (het totaal in de cashflow volgens de gekozen methode) en "Nog te verdelen" (PEW − verdeeld).

## Oud FO inlezen

De macro (en `tools/omzetten.py`, dezelfde regels in Python) neemt over:
- de exports als waarden, en `Behind the scens` (CODA-posten als tekst blijven tekst, zodat de groepering in kolom Y van FO - vorig/huidig werkt);
- de prijslijst per bouwnummer (bouwnummer, blok, type, handmatige kolommen);
- per woningtype:
  - naam en type;
  - bedragen: handmatig als de prijslijst ze niet verklaart;
  - termijnschema, planning, verkocht en transport (omgezet naar %), kadastrale kosten/rente;
  - kosten: budget, KAO, % en de planning als die afwijkt;
- de SKO-bedragen van de aansluiting;
- overige opbrengsten en "verschillen" per type;
- het DAEB-blok (aantal en verkoop zo nodig van `Verkooptempo`);
- de kostenposten (prognose);
- de FD-kolom en de toelichting van `2. Ontw PEW`;
- de vorige prognose uit `1. Cashflow`.

`Controle oud FO` vergelijkt daarna de cashflow per periode met het oude FO.

**Getest** (LibreOffice, volledig doorgerekend):
- **Demo-FO:** opbrengsten per kwartaal gelijk. Bij de kosten verschillen 4 kwartalen €150–€600 (totaal €9). Het oude FO rekende de kosten van kopersmeerwerk bij één type met de KAO-opties; het sjabloon rekent bij alle types met het kopersmeerwerk.
- **Een echt FO** (3 types, DAEB met AK fees en bijkomende kosten, 2019 t/m 2030): alle 36 periodes gelijk. Dat FO staat niet in deze repository.

Wat niet meegaat:
- een extra kolom op `2. Ontw PEW`, zoals "FD corr DAEB" (de omzetting meldt het);
- eigen afwijkende formules in het oude FO: de waarden gaan mee, de afwijkende rekenregel niet.

## Bouwen en controleren

```bash
python3 tools/bouw.py
```

```bash
python3 tools/controleer.py
```

Een ander oud FO omzetten (zonder Excel):

```bash
python3 tools/bouw.py --oud pad/naar/FO.xlsx --project "Naam" --out werkbestand.xlsm
```

Nodig: Python 3 met `lxml`, `openpyxl`, `python-pptx`, `olefile` en `oletools`. Voor de controle ook LibreOffice.

De macro's zijn niet in Excel uitgevoerd; daar is in deze omgeving geen Excel voor. "Oud FO inlezen" is wel met de
VBA-uitvoering van LibreOffice gedraaid op het lege sjabloon, voor het demo-FO en het echte FO:
- het demo-FO geeft dezelfde uitkomst als de Python-omzetting;
- het echte FO is per periode gelijk aan het oude FO.

Test de knoppen bij het eerste gebruik op een kopie.
