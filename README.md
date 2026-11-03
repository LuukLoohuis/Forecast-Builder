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
| `Prijslijst` | per woningtype: naam, type in de prijslijst, aantal en bedragen (uit Salesforce, btw eraf; handmatig overschrijven mag). Aansluiting met SKO en PEW per type, met daarnaast het termijnschema (termijn + %). Onderaan "te boeken in SAP" per SAP-nummer en kostendrager, en de prijslijst per bouwnummer |
| `CF - opbrengsten` | per woningtype: planning bouwtermijnen, verkocht %, getransporteerd %, grondprijsaanpassing; kadastrale kosten en rente (handmatig); overige opbrengsten, correctie, DAEB |
| `CF - kosten` | bouwtermijnen kosten per type (eigen planning HWB, standaard gelijk aan de opbrengsten) en per kostenpost Handmatig, Voorstel of Transport |
| `Dashboard`, `Scenario's` | cashflow, kerncijfers, grafieken en scenario's: rekenen automatisch mee |
| `2. Ontw PEW` | PEW huidig/vorig uit SAP, meest recente FD (invullen), toelichting |
| `CF - hulp VENA` | de definitieve cashflow per maand, in het VENA-format |

Verborgen: de rekenbladen van het scenariomodel (`Invoer`, `Woningtypes`, `Scenario kosten`, `Model`), `PowerPoint`
(welke tekst naar welke vorm), `Vorige prognose`, `Controle oud FO` (verschijnt na Oud FO inlezen) en `Behind the scens`
(koppeling CODA-post → groep).

Vervallen ten opzichte van het oude FO: `Aansluiting opbrengsten` (nu op `Prijslijst`), `Verkooptempo`, `Woningtypes` als
invoer, `1. Cashflow`, `3. Restant bdgt`, `Notities FO gesprek`, `Toelichting onderlegger`, `1. Ontw businessc FD` en
`5. K&R businessc`.

## Weergave

- **Jaarvakken:** in CF - opbrengsten en CF - kosten staat aan het begin van elk jaar een lijn, en boven elk kwartaal het jaartal (Q1 vet).
- **Start bouw en oplevering:** vul je in op `Start` (rij 7 en 9). Bovenaan CF - opbrengsten (rij 4) en CF - kosten (rij 2) staat "Start bouw" of "Oplevering" boven het kwartaal, en die kolom is gekleurd (groen en blauw). Oud FO inlezen neemt ze over van de markeringen in het oude FO.
- **Inklappen:**
  - per woningtype is het invoerblok een groep; woningtypes zonder woningen zijn ingeklapt;
  - de rekenblokken (aantallen per termijn en €) zijn standaard ingeklapt;
  - de kwartalen t/m de laatste actuals zijn een ingeklapte kolomgroep; met + klap je ze open.

  De macro werkt dit bij na Oud FO inlezen, Nieuw kwartaal en het inlezen van actuals.
- **Alle woningen:** onder het verkooptempo per kwartaal het % verkocht en getransporteerd van alle woningen, en cumulatief.
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
- Bij Voorstel of Transport gaat wat je in de postregel invult voor; de rest van het budget wordt verdeeld.

## Oud FO inlezen

De macro (en `tools/omzetten.py`, dezelfde regels in Python) neemt over:
- de exports als waarden, en `Behind the scens`;
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
