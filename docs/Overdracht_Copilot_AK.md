# Overdracht aan Copilot: AK-dekking valt te laag uit

## 1. Het probleem in de woorden van de gebruiker
> "AK-dekking valt laag uit. Dit strookt niet met hoe ik het heb."

In het RDFO ziet de regel **AK-dekking** er per prognose zo uit (jaarbedrag 2026, EUR x 1.000):

| Budget 2026 | Prognose Q1 2026 | Prognose Q2 2026 | Prognose Q3 2026 |
|---|---|---|---|
| n.b. | 2.409 | 1.358 | 740 |

- Elke prognose is een **jaarprognose voor heel 2026**, geen kwartaalbedrag. Voor 2027 telt één jaarbedrag per project.
- De jaarprognose wordt bij elke prognose **lager**. Volgens de gebruiker klopt dat niet: in zijn eigen berekening (het oude RDFO) ligt de AK-dekking hoger en zakt die niet zo.
- De **opbouw** van het tabblad AK-dekking is goed en blijft:
  - totaal per kwartaal;
  - rekenvoorbeeld per project;
  - tabel per project.

**Opdracht:**
1. Zoek uit waarom de uitkomst afwijkt van de berekening van de gebruiker.
2. Laat het bestand rekenen zoals de gebruiker het doet.
3. Maak het verschil per project zichtbaar.
4. Voeg de tabel "AK per project per prognose" toe (hoofdstuk 7c).

## 1b. Bevindingen met het eigen bestand van de gebruiker (Claude, november 2026)
Doorgerekend is het geanonimiseerde bestand `RDFO_Noordwest_Q3_Budget_geanonimiseerd_hersteld.xlsx`, met de tabbladen ZPMW03, AK Percentages en AK realisatie. Script: `tools/ak_aansluiting.py <bestand>`. De cijfers zijn geanonimiseerd en daarom indicatief; het patroon is wel duidelijk.

**Hoe het oude bestand rekent** (tabblad "AK dekking 2027", rij 8 en verder):
- per project en per kwartaal: `Kosten <jaar>-Qq / 1,05 × AK %`;
- AK % = `VLOOKUP(WBS-element, 'AK Percentages'!D:H, 5)`;
- jaar = som van de vier kwartalen;
- RDFO-regel AK Dekking = `SUMIFS(..., Bedrijfsnummer, "1001")`, dus **alleen bedrijf 1001**;
- de kolom "Actual" = −totaal tabblad AK realisatie (geboekt);
- regel "Totaal dbl AK" = Actual + resterende kwartalen;
- tabblad "BOB, BOR & AK 2027-2030" (outlook): AK = Kosten jaar × AK %, **zonder /1,05**. Dat is inconsistent met de rest; navragen.

**In de Q3-export zijn "Kosten 2026-Q1..Q3" leeg (0).** Alleen Q4 (= het restant) is gevuld. Gerealiseerde kosten staan in "Kosten LJ".

**De rekenwijze van het nieuwe bestand op deze data:**

| Variant | AK-dekking 2026 |
|---|---|
| Nieuw bestand: geboekt (839k) + zonder project (498k) + kosten Q4 / 1,05 × AK % | **626.256** (≈ de 740 die de gebruiker ziet) |
| Idem, maar negatieve kosten Q4 tellen als 0 | **2.089.852** |
| Idem, maar zonder één project (Zuiderveld, Spoorzicht) | 2.084.979 |
| Kostenmethode (Kosten LJ + Kosten 2026) / 1,05 × AK %, per project minimaal 0 | 1.876.987 |
| Q3-prognose van de gebruiker (RDFO, "Q3 Prognose EJ") | **2.686.000** |

**Conclusies:**
1. **Hoofdoorzaak: negatieve kosten.**
   - Eén project (`G.800044.8.1015`, Zuiderveld, Spoorzicht) heeft **Kosten 2026-Q4 = −25,5 mln** (een correctie of afboeking). Dat geeft −1,46 mln AK-dekking.
   - In totaal staat er −25,6 mln aan negatieve kosten Q4 tegenover +16,0 mln positief.
   - **Beslissing nodig van de gebruiker:** krijgt een negatieve kostenpost een negatieve AK (huidige werking), of telt die als 0?
2. **Het grootboek loopt achter.** De dump AK realisatie bevat boekingen t/m periode 7 (juli), terwijl de peildatum 30-09 is. Augustus en september (±0,4 mln, op basis van het gemiddelde per maand) ontbreken. Bij de methode "geboekt + rest" valt dat weg.
3. **AK geboekt zonder bekend project: 498k** (37% van het geboekte totaal). De WBS in de tekst matcht niet met een project in ZPMW03. Dit telt wel mee op regioniveau, maar niet per project.
4. **50 van de 80 WBS-elementen hebben geen AK %** in de dump, waaronder `G.800045.8.1015` (Oosterwaard, Buitenplaats). Op deze data is het effect klein (de kosten zijn per saldo negatief), maar het moet wel kloppen.
5. **Bedrijfsnummer:** het oude bestand telt alleen bedrijf 1001. Het nieuwe telt elk project met AK = Ja (ook deelnemingen, bijv. bedrijf 1089). Zet AK-dekking = Nee voor deelnemingen, of filter op bedrijfsnummer.

**Voorgestelde herstelvolgorde:**
1. Laat de gebruiker beslissen over negatieve kosten. Advies: AK per project per kwartaal minimaal 0, met een Controle-WAARSCHUWING die de projecten met negatieve kosten noemt.
2. Voer de instelling `P_AKMethode` in (hoofdstuk 4). Advies voor de Q-prognoses: "Kosten (zoals oude RDFO)", met de geboekte AK alleen als realisatie/controle. Dan heeft de achterstand in het grootboek geen invloed.
3. Maak het WBS-matchen in het grootboek ruimer (begin van het WBS-element / projectdefinitie) en toon de lijst "zonder project".
4. Voeg het filter op bedrijfsnummer toe, of zet de keuze per project op de Projectlijst.

## 2. Hoe de gebruiker het berekent (te bevestigen met het oude RDFO-bestand)
- **Per project:** AK-dekking = **kosten / 1,05 × AK %**.
  - AK % = het percentage van het project uit SAP (dump "AK-percentages"), of een eigen % op de Projectlijst.
  - 1,05 = instelling `P_AKFactor`.
- **Jaar 2026** = AK over de kosten van **heel 2026**.
- **2027** = kosten 2027 / 1,05 × AK %.
- De **geboekte AK** (grootboekdump "AK realisatie") is in het oude bestand de realisatie.
- **Nog te verifiëren in het oude bestand:**
  - welke kostenkolom de basis is (Kosten LJ + Kosten RJ? een kolom "Kosten 2026"? totale kosten volgens de PEW?);
  - of de jaarprognose de geboekte AK gebruikt of alleen de kostenberekening.

  Leg één project uit het oude bestand naast het nieuwe en reken het met de hand na.

## 3. Wat het bestand nu doet
Het bestand wordt gegenereerd door `tools/build_rdfo.py`. Wijzig nooit het .xlsm zelf.

**Tabblad AK-dekking, functie `build_ak`.** Per project (rij vanaf 21) en per kwartaal van 2026 (kolom F–I):
- **t/m het peilkwartaal** (`AK_PeilQ`):
  - zijn er grootboekboekingen in 2026 (`AKR_NJaar > 0`)? Dan geldt de **geboekte AK** van dat WBS-element in dat kwartaal: `SUMIFS(AKR_Bedrag, AKR_WBS, wbs, AKR_Jaar, 2026, AKR_Kw, q)`;
  - anders: **Kosten LJ / 1,05 × AK %**, verdeeld over de kwartalen.
- **na het peilkwartaal:** kosten van dat kwartaal (ZPMW03 "Kosten 2026-Qq"; als die ontbreken: Kosten RJ gelijk verdeeld) / 1,05 × AK %.
- **Jaar (kolom J)** = Q1 + Q2 + Q3 + Q4. Dit wordt per project vastgelegd als kengetal "AK-dekking FY" (doeljaar 2026).
- **2027 (kolom S)** = Kosten 2027 / 1,05 × AK %. Wordt vastgelegd als "AK-dekking FY" doeljaar 2027.

**Regio:**
- AK die geboekt is zonder (bekend) WBS (`AK_NietToegewezen`) telt mee in het RDFO-totaal.
- Sinds de laatste aanpassing wordt die ook vastgelegd, met WBS "REGIO". Eerder vastgelegde prognoses hebben deze regel niet.

**Bronvelden** (`z_fields` / `z_derived` in `build_rdfo.py`):

| Veld | Bron |
|---|---|
| `M_KostYTD` | ZPMW03 "Kosten LJ" (aanname: gerealiseerd t/m peildatum) |
| `M_KostRest` | "Kosten RJ" (aanname: restant na peildatum) |
| `M_KostQ1..4` | "Kosten 2026-Q1..Q4" |
| `M_KostFY1` | "Kosten 2027" |

**Grootboek** (`build_stg_akr`):
- bedrag = **min** "Bedrag in eigen valuta" (AK-dekking is credit);
- jaar = "Boekjaar";
- kwartaal = "Boekingsperiode" / 3;
- WBS = kolom "WBS-element", anders uit de tekst (zoekt "G.").

## 4. Belangrijkste verdenking: methodeverschil "geboekt" tegenover "kosten"
Het patroon past bij een methodeverschil: hoe later de prognose, hoe lager de uitkomst (2.409 → 1.358 → 740).
- **Het nieuwe bestand** neemt voor de afgelopen kwartalen de **geboekte AK** uit het grootboek. Het rekent alleen de rest van het jaar met kosten × AK %.
- **Wordt AK in het grootboek later of maar deels geboekt?** Bijvoorbeeld pas bij de jaarafsluiting, per halfjaar, of niet op alle projecten. Dan is "geboekt t/m de peildatum" veel lager dan kosten × AK % over dezelfde periode. Elke volgende prognose vervangt meer kwartalen door dat lage geboekte bedrag, en daardoor daalt het jaarbedrag.
- **De gebruiker rekent waarschijnlijk** het hele jaar met kosten × AK %. Dan zakt het bedrag niet.

**Te bouwen:** een instelling `P_AKMethode` op Invoer › 1 met twee keuzes:
1. **"Kosten (zoals oude RDFO)"** — standaard. Jaar 2026 = (Kosten LJ + Kosten RJ) / 1,05 × AK %. De kwartaalverdeling blijft alleen ter informatie.
2. **"Geboekt + rest"** — de huidige werking.

De geboekte AK blijft bij beide methodes zichtbaar als **realisatie** (eigen kolom). Per project komt een kolom **"geboekt t/m peildatum − berekend t/m peildatum"**. Zo zie je meteen waar het grootboek achterloopt.

## 5. Overige verdachten (allemaal nalopen, met het eigen bestand van de gebruiker)
1. **Betekenis van de kostenkolommen.** Is "Kosten LJ" echt gerealiseerd 2026 t/m de peildatum, en is "Kosten RJ" het restant (niet het hele jaar)? Zo niet, dan valt de basis te laag of dubbel uit. Vergelijk per project Kosten LJ + Kosten RJ met de jaarkosten in het oude bestand.
2. **AK % ontbreekt of is te laag.**
   - De SAP-dump koppelt op "WBS-element". Staat het % in de dump op projectdefinitieniveau of een ander WBS-niveau, dan matcht het niet en krijgt het project 0.
   - Controle "Projecten zonder AK %" moet 0 zijn.
   - Wordt het % als 6 gelezen in plaats van 6%, of andersom?
3. **Projecten tellen niet mee.** Op de Projectlijst moeten Prognose = Ja en AK-dekking = Ja staan. Nieuwe projecten komen er bij inlezen vanzelf bij (alles op Ja). Projecten die nooit zijn toegevoegd tellen niet mee.
4. **WBS-match in het grootboek.** Boekingen op een sub-WBS (bijv. `G.900001.8.9090.01`) of zonder WBS gaan naar "zonder project". Ze tellen dan wel in de regio, maar niet per project.
5. **Teken.** Is "Bedrag in eigen valuta" in de export al positief, dan wordt de geboekte AK negatief. Controle "Geboekte AK per saldo negatief" bestaat. Voeg een controle toe voor negatieve kosten in ZPMW03.
6. **Overgezette prognoses** (`RDFO_OudePrognosesOverzetten`).
   - Q1/Q2 uit het oude bestand zijn ×1.000 omgerekend uit de kolom "AK dekking 2026" → "Totaal".
   - Controleer of de totaalkolom gelezen is (en niet Q1), en of de bedragen al in hele euro's stonden.
7. **Aanvullende invoer** (Invoer › 3, rubriek AK-dekking) gaat altijd voor de berekening. Een ingevuld bedrag daar overschrijft de hele berekening.

## 6. Diagnose in het bestand van de gebruiker (eerst doen, voordat je iets bouwt)
1. **Vraag de gebruiker om:** de ZPMW03-export, de AK-percentages-dump, de grootboekdump AK realisatie en het oude RDFO-bestand (met de AK-berekening).
2. **Bouw een aansluittabel per project** (tijdelijk, in Python of in een kladblad), met deze kolommen:

   | Kolom | Bron |
   |---|---|
   | Kosten LJ, Kosten RJ, Kosten 2027 | ZPMW03 |
   | AK % gebruikt | eigen % of SAP-dump |
   | A = (Kosten LJ + Kosten RJ) / 1,05 × AK % | kostenmethode |
   | B = geboekte AK 2026 | grootboek |
   | C = AK-dekking jaar | nieuw bestand, tabblad AK-dekking, kolom J |
   | D = AK-dekking 2026 | oude RDFO |
   | Verschillen | A−D, C−D, B−A |
3. **Interpreteer de uitkomst:**
   - A ≈ D maar C < D → methodeverschil (hoofdstuk 4);
   - A < D → kostenbasis of AK % wijkt af (punt 5.1, 5.2);
   - grote "zonder project" → WBS-match (punt 5.4).
4. Leg de bevinding vast in dit document voordat je iets bouwt.

## 7. Te bouwen (na de diagnose)
**a. Instelling `P_AKMethode`** (hoofdstuk 4). Standaard "Kosten (zoals oude RDFO)". Pas `build_ak` kolom F–J aan, en daarmee VERW `AK0`/`AKYTD` en het vastleggen.

**b. Aansluitkolommen op AK-dekking** (ingeklapt), per project:
- Kosten LJ, Kosten RJ;
- AK volgens kosten;
- AK geboekt;
- verschil.

Het rekenvoorbeeld moet de gekozen methode uitschrijven.

**c. Tabel "AK per project per prognose"**, direct onder het totaal per kwartaal:

| Project | WBS | Budget 2026 | Q1 2026 | Q2 2026 | Q3 2026 | Q4 2026 | Huidig 2026 | Δ vs laatste prognose | Q1–Q4 2027 | Huidig 2027 |
|---|---|---|---|---|---|---|---|---|---|---|

- Vastgelegde waarden: `SUMIFS(L_Waarde, L_WBS, wbs, L_Actief, 1, L_Versie, v, L_Versiejaar, P_VJ, L_Doeljaar, jaar, L_Kengetal, "AK-dekking FY")`.
- Niet vastgelegd → **"n.b."**, nooit 0.
- Totaalregel per kolom, inclusief de REGIO-regel. Het totaal moet aansluiten op de RDFO-regel AK-dekking per versie.
- Δ rood als het negatief is.

**d. Controle:** WAARSCHUWING als de geboekte AK t/m de peildatum meer dan 25% afwijkt van kosten × AK % over dezelfde periode, met de projecten in de lijst eronder.

## 8. Waar de code staat

| Bestand / functie | Inhoud |
|---|---|
| `tools/build_rdfo.py` › `build_ak` | Tabblad AK-dekking: samenvatting rij 5–9, rekenvoorbeeld rij 11–16, tabel vanaf rij 21; verborgen hulpkolommen Z–AC |
| `build_stg_akr`, `build_bron_akr` | Grootboekdump → STG_AKR; demo-dump |
| `z_fields`, `z_derived` | ZPMW03-velden (`M_KostYTD`, `M_KostRest`, `M_KostQ1..4`, `M_KostFY1`) |
| `verw_columns` | VERW_Project: `AKYTD` (← AK-dekking kolom Q), `AK0` (← J), `AK1` (← S) |
| `build_calc` | RDFO-regel AK-dekking per versie (Historie) en upload (`SUM(V_AK0) + AK_NietToegewezen`) |
| `KT`, `build_vastleggen` | Wat er per project wordt vastgelegd, plus de REGIO-regel |
| `PARAMS` | Instellingen op Invoer › 1 (hier `P_AKMethode` toevoegen) |
| `macros/modRDFO.bas` | Inlezen (stap 3 AK-percentages, stap 4 grootboek), vastleggen, oude prognoses overzetten |
| `tools/test_rdfo.py` | 32 tests; AK: "AK-dekking per kwartaal" en "AK zonder grootboek" |
| `tools/release.py` | Alles bouwen en testen: `python3 tools/release.py <recalc.py> <werkmap>` |

**Belangrijke namen:**

| Naam | Betekenis |
|---|---|
| `P_AKFactor` | 1,05 |
| `AK_PeilQ` | peilkwartaal |
| `AKR_Bedrag`, `AKR_Jaar`, `AKR_WBS`, `AKR_Kw`, `AKR_Bekend` | grootboekregels |
| `AKR_NJaar` | aantal grootboekregels in het rapportagejaar |
| `AK_NietToegewezen` | geboekt zonder project |
| `AK_Totaal` | totaal AK-dekking regio |
| `V_AK0`, `V_AKYTD` | per project |
| `L_*` | Historie |

## 9. Acceptatie
1. Met `P_AKMethode` = "Kosten": jaar 2026 per project = (Kosten LJ + Kosten RJ) / 1,05 × AK %, onafhankelijk van het grootboek. Dit geldt ook als het grootboek leeg is of achterloopt.
2. Met de dataset van de gebruiker sluit het jaarbedrag per project aan op het oude RDFO (verschil < € 1.000 per project), of het verschil is per project verklaard.
3. De tabel per prognose toont Budget/Q1–Q4 per project. Het totaal is gelijk aan de RDFO-regel AK-dekking per versie; "n.b." waar niet vastgelegd.
4. Alle bestaande tests blijven groen, aangevuld met tests voor 1 en 3. De herberekening geeft 0 formulefouten (behalve de bekende `#N/A` in grafiekhulpkolommen).

## 10. Randvoorwaarden
- Nederlands, kort, voor een controller zonder VBA-kennis. Huisstijl geel/donkerblauw.
- Historie is onveranderlijk; al vastgelegde prognoses worden niet herrekend.
- Niets stil op 0: ontbrekend = "n.b." plus een melding op Controle.
- Na afloop:
  - FAQ "Hoe is de AK-dekking berekend?" bijwerken, plus de mappingregel "RDFO › AK-dekking" in `tools/docs_content.py`;
  - README en `docs/Overdracht_Copilot.md` bijwerken;
  - `release.py` draaien;
  - alleen `RDFO_template.xlsm` opleveren.
- De macro's zijn alleen buiten Excel gecontroleerd. Test Bronnen inlezen en Prognose vastleggen in Excel zelf.

## 11. Vragen aan de gebruiker (eerst stellen)
0. **Negatieve kosten** (bijv. −25,5 mln op Zuiderveld, Spoorzicht): AK-dekking dan negatief, of 0? En is −25,5 mln een echte correctie, of een fout in de export?
1. Hoe reken je AK-dekking 2026 in je eigen bestand: over welke kostenkolom, en tel je de geboekte AK mee of alleen kosten × AK %?
2. Wanneer wordt AK in het grootboek geboekt: maandelijks, per kwartaal of bij de jaarafsluiting?
3. Kun je de ZPMW03-export, de AK-percentages-dump, de grootboekdump en je oude RDFO met de AK-berekening aanleveren (eventueel geanonimiseerd)?
4. Welk bedrag verwacht je ongeveer voor AK-dekking 2026 in Q3?
