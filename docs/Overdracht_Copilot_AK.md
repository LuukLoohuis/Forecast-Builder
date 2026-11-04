# Overdracht aan Copilot: AK-dekking in het RDFO repareren en per prognose zichtbaar maken

## 1. Opdracht in één alinea
De AK-dekking (opslag algemene kosten die een project draagt) komt in het RDFO **veel te laag** uit. De regel AK-dekking in de RDFO toont per prognose (jaarbedrag 2026, EUR x 1.000):

| Budget | Q1 2026 | Q2 2026 | Q3 2026 |
|---|---|---|---|
| n.b. | 2.409 | 1.358 | 740 |

**De opbouw van het tabblad is goed** en moet blijven:
- totaal per kwartaal;
- rekenvoorbeeld per project;
- tabel per project.

**De uitkomst klopt niet.** Zoek de oorzaak, herstel de berekening en voeg op het tabblad AK-dekking een tabel toe met per project de **vastgelegde jaarprognose AK-dekking per versie**: Budget, Q1, Q2, Q3, Q4 2026, de huidige prognose en het verschil. Doe hetzelfde voor 2027 als één jaarbedrag per project.

## 2. Wat de gebruiker bedoelt (functionele eisen)
1. **Elke prognose is een jaarprognose.** "Prognose Q3 2026" = de verwachte AK-dekking over **heel 2026**, gemaakt op peildatum 30-09-2026. Het is geen kwartaalbedrag.
2. **2027 is één jaarbedrag per project** (geen kwartalen).
3. **Berekening per project, zoals in het oude RDFO:**
   - prognose = kosten / 1,05 × AK % van het project (factor `P_AKFactor`, standaard 1,05);
   - AK % = eigen % op de Projectlijst (kolom I), anders het % uit de SAP-dump AK-percentages;
   - realisatie = geboekte AK uit de grootboekdump "AK realisatie".
4. **Jaar 2026 = geboekt t/m de peildatum + prognose voor de rest van het jaar.**
5. **Per project terug te zien in het tabblad AK-dekking:** wat de vastgelegde AK-jaarprognose was in Budget, Q1, Q2, Q3 (en Q4), plus de huidige berekening en het verschil met de vorige prognose. De totaalregel per versie moet aansluiten op de regel AK-dekking in de RDFO.
6. **Simpel en naloopbaar:** geen ingewikkelde formules. Per project moet je kunnen zien hoe het bedrag tot stand komt (het rekenvoorbeeld bestaat al).
7. **Projectkeuzes:** alleen projecten met Prognose = Ja en AK-dekking = Ja op de Projectlijst tellen mee.

## 3. Waar de code staat
Het Excel-bestand wordt **gegenereerd**. Pas nooit het .xlsm zelf aan, maar de generator.

| Bestand | Wat |
|---|---|
| `tools/build_rdfo.py` | Generator (openpyxl). Relevante functies: `build_ak` (tabblad AK-dekking), `build_stg_akr` (grootboekdump → STG_AKR), `build_bron_akr` (demo-dump), `z_fields`/`z_derived` (ZPMW03-velden, o.a. `M_KostQ1..4`, `M_KostYTD`, `M_KostRest`), `verw_columns` (VERW_Project: `AKYTD`, `AK0`..`AK4`), `build_calc` (CALC_Versies: RDFO-regel AK per versie), `KT` (wat er bij vastleggen per project wordt opgeslagen: `AK0` → kengetal "AK-dekking FY" doeljaar RJ, `AK1` → RJ+1). |
| `macros/modRDFO.bas` | VBA: `RDFO_BronnenInlezen` (stap 3 AK-percentages → `Bron_AK`, stap 4 grootboekdump → `Bron_AKR`), `RDFO_PrognoseVastleggen`, `RDFO_OudePrognosesOverzetten` (leest "ak dekking 20##" uit het oude bestand, x 1.000). |
| `tools/test_rdfo.py` | 31 tests; AK-test: "AK-dekking per kwartaal: Q1+Q2+Q3+Q4 = jaar = RDFO". |
| `tools/release.py` | Bouwen + tests + herberekenen (LibreOffice) + template. Draaien: `python3 tools/release.py <pad>/recalc.py <werkmap>`. |

**Belangrijke namen in de werkmap:**

| Naam | Betekenis |
|---|---|
| `P_AKFactor` | 1,05 |
| `AK_PeilQ` | kwartaal van de peildatum |
| `AKR_Bedrag`, `AKR_Jaar`, `AKR_WBS`, `AKR_Kw`, `AKR_Bekend` | grootboekregels |
| `AK_NietToegewezen` | geboekt zonder bekend project |
| `AKR_N` | aantal regels met bedrag |
| `AKR_Rijen` | regels in de dump |
| `AK_Totaal` | totaal regio, lopend jaar |
| `V_AK0`, `V_AKYTD` | per project uit het tabblad AK-dekking |
| `L_*` | Historie (vastgelegde regels) |

## 4. Huidige berekening (stand van de code)
**Tabblad AK-dekking, per project (rij vanaf `AK0R` = 21):**
- **Kolom F–I (Q1–Q4 van RJ):**
  - als het kwartaal ≤ `AK_PeilQ` en er een grootboekdump is: `SUMIFS(AKR_Bedrag, AKR_WBS, WBS, AKR_Jaar, RJ, AKR_Kw, q)`. Voor het peilkwartaal geldt `>= q`, zodat latere perioden ook meetellen;
  - anders: `Kosten RJ-Qq / P_AKFactor × AK %`, waarbij de kosten uit de ZPMW03-kolom **"Kosten 2026-Qq"** (`M_KostQq`) komen.
- **Kolom J:** jaar = F+G+H+I. Dit gaat naar VERW `AK0` en wordt per project vastgelegd als "AK-dekking FY" (doeljaar RJ).
- **Kolom S/U/W/Y:** RJ+1..RJ+4 = `Kosten RJ+n / P_AKFactor × AK %`. RJ+1 wordt vastgelegd als "AK-dekking FY" doeljaar RJ+1.

**RDFO-regel AK-dekking (`build_calc`):**
- **Upload:** `SUM(V_AK0) + AK_NietToegewezen`.
- **Vastgelegde versies (Budget/Q1–Q4):** `SUMIFS(L_Waarde, …, L_Kengetal, "AK-dekking FY")`, alleen de som van de **projectregels**.
- Aanvullende invoer (Invoer › 3, rubriek AK-dekking) gaat altijd voor.

**Grootboek (`build_stg_akr`):**
- bedrag = **−** "Bedrag in eigen valuta": AK-dekking is credit, wordt positief;
- jaar = "Boekjaar";
- WBS = kolom "WBS-element", anders uit de tekst (zoekt "G.");
- kwartaal = ROUNDUP(Boekingsperiode/3), met periode 13–16 als Q4.

## 5a. Al hersteld (door Claude, test 'AK zonder grootboek' groen)
- **Oorzaak 1:** t/m de peildatum geldt zonder geboekte AK in het rapportagejaar: `Kosten LJ / P_AKFactor × AK %`, verdeeld naar de kwartaalkosten (anders gelijk). Kolommen M–P op AK-dekking tonen de gebruikte kostenbasis. De verborgen hulpkolommen Z–AC bevatten de ruwe kosten. `AKR_NJaar` = aantal grootboekregels in het rapportagejaar.
- **Oorzaak 2:** na de peildatum geldt zonder kwartaalkosten: `Kosten RJ` (restant), gelijk verdeeld over de resterende kwartalen.
- **Oorzaak 3:** bij vastleggen komt een extra Concept-regel met WBS "REGIO", kengetal "AK-dekking FY", waarde `AK_NietToegewezen`. Daardoor sluiten vastgelegde versies aan op de upload. Let op: al vastgelegde versies hebben deze regel niet.
- **Nog open:** oorzaak 4 t/m 7 en de nieuwe tabel per prognose (hoofdstuk 6).

## 5. Vermoedelijke oorzaken van "veel te laag" (eerst uitzoeken, in deze volgorde)
1. **Terugval naar kwartaalkosten geeft nul voor afgelopen kwartalen.**
   - Zonder grootboekdump (of als `Boekjaar`/`Bedrag` niet gevonden wordt, of er geen regels van 2026 zijn) rekent het bestand de afgelopen kwartalen met "Kosten 2026-Q1..Q3".
   - In een ZPMW03-export op peildatum Q3 zijn dat vaak **restant/plan-kolommen met 0** voor voorbije kwartalen. De gerealiseerde kosten staan in **"Kosten LJ"**.
   - Gevolg: de AK over de afgelopen kwartalen valt weg, en hoe later in het jaar, hoe lager de uitkomst. Dat past bij 2.409 → 1.358 → 740.
   - In een eerdere versie werd hier `Kosten LJ / 1,05 × AK %` gebruikt.
   - **Herstel:** zonder geboekte AK geldt voor t/m de peildatum `Kosten LJ / P_AKFactor × AK %`, verdeeld naar rato van de kwartaalkosten of volledig in het peilkwartaal.
2. **Ontbrekende kwartaalverdeling van de restkosten.**
   - Heeft de export geen kolommen "Kosten 2026-Qq" (alleen "Kosten 2026"), dan is de prognose voor de rest van het jaar 0.
   - **Herstel:** rest van het jaar = `Kosten RJ` (= `M_KostRest`, restant na de peildatum) / factor × AK %. Gebruik de kwartaalkolommen alleen voor de verdeling over de kwartalen, en val terug op gelijk verdelen over de resterende kwartalen.
3. **Vastgelegde versies missen "AK geboekt zonder project".**
   - De upload telt `AK_NietToegewezen` mee, maar bij vastleggen worden alleen projectregels opgeslagen. Vastgelegde Q1/Q2/Q3 zijn daardoor lager dan de upload van hetzelfde moment.
   - **Herstel:** bij vastleggen een regel op regioniveau opslaan (WBS leeg of "REGIO", kengetal "AK-dekking FY"), die CALC meetelt. Alternatief: boekingen op sub-WBS'en toewijzen aan het project (zie punt 4).
4. **WBS in het grootboek matcht niet.**
   - Boekingen op een sub-WBS (bijv. `G.900001.8.9090.01`) of met een afwijkende notatie matchen niet met de Projectlijst. Dan komen ze bij "zonder project", en dus niet in de vastgelegde projectbedragen.
   - **Herstel:** koppelen op het begin van het WBS-element (projectdeel), of op de projectdefinitie.
5. **Teken.**
   - Het bestand draait het teken om (credit → positief). Is de export al positief, dan wordt AK negatief en trekt die de som omlaag. Hetzelfde geldt voor kosten: zijn kosten in ZPMW03 negatief, dan wordt de AK-prognose negatief.
   - Er bestaat een controle "Geboekte AK per saldo negatief". Voeg er een toe voor negatieve kosten.
6. **AK % ontbreekt.** Projecten zonder eigen % en zonder regel in de SAP-dump (match op WBS-element) krijgen geen AK. Controle "Projecten zonder AK %" bestaat (WAARSCHUWING).
7. **Andere methode in overgezette versies.**
   - Q1/Q2 kunnen met `RDFO_OudePrognosesOverzetten` uit het oude bestand komen ("AK dekking 2026"-totaalkolom, x 1.000). Dat is de oude methode.
   - Controleer of de waarden in duizenden goed zijn omgerekend (x 1.000), en of de totaalkolom gelezen is (en niet één kwartaal).

**Werkwijze:** reproduceer met de demo (`RDFO_prototype.xlsm`). Haal daarvoor de grootboekdump weg en zet de kwartaalkosten van Q1–Q3 op 0: de AK zakt dan precies zoals bij de gebruiker. Leg per oorzaak in een test vast dat het herstel werkt.

## 6. Nieuwe tabel: AK-dekking per project per prognose
Plaats deze op het tabblad AK-dekking, direct onder het totaal per kwartaal. Het rekenvoorbeeld en de kwartaaltabel blijven, eventueel ingeklapt.

| Project | WBS | Budget 2026 | Q1 2026 | Q2 2026 | Q3 2026 | Q4 2026 | Huidig 2026 | Δ huidig vs laatste vastgelegde | Q1–Q4 2027 (vastgelegd) | Huidig 2027 |
|---|---|---|---|---|---|---|---|---|---|---|

- **Vastgelegde kolommen:**
  - `SUMIFS(L_Waarde, L_WBS, wbs, L_Actief, 1, L_Versie, v, L_Versiejaar, P_VJ, L_Doeljaar, RJ (of RJ+1), L_Kengetal, "AK-dekking FY")`;
  - niet vastgelegd → **"n.b."** (nooit 0);
  - Budget heeft versiejaar = RJ (zie de bestaande budgetkolom in `build_ak`).
- **Huidig:** kolom J (2026) en kolom S (2027) van de kwartaaltabel.
- **Totaalregel per kolom**, plus een regel "geboekt zonder project" (zie oorzaak 3). Het totaal moet gelijk zijn aan de RDFO-regel AK-dekking per versie.
- **Opmaak:** Δ rood als het negatief is. Alleen projecten met AK = Ja tonen.

## 7. Acceptatie (tests toevoegen in `tools/test_rdfo.py`)
1. **Zonder grootboekdump:** jaar 2026 per project = Kosten LJ / 1,05 × AK % + Kosten RJ / 1,05 × AK %, dus nooit lager doordat de kwartaalkolommen 0 zijn.
2. **Export zonder kolommen "Kosten 2026-Qq":** jaarbedrag gelijk aan punt 1.
3. **Met grootboekdump:** jaar = geboekt (incl. zonder project) + rest/1,05 × AK %. De vastgelegde versie geeft hetzelfde totaal als de upload van dat moment.
4. **Tabel per prognose:** totaal per versie = RDFO-regel AK-dekking per versie. Een niet-vastgelegde versie toont "n.b.".
5. Alle bestaande 31 tests blijven groen, en de herberekening geeft 0 formulefouten (behalve de bekende `#N/A` in grafiekhulpkolommen).

## 8. Randvoorwaarden
- Nederlands, korte teksten, voor een controller zonder VBA-kennis. Huisstijl geel/donkerblauw.
- Historie is onveranderlijk. Nieuwe kengetallen of regelsoorten mogen wel worden toegevoegd.
- Niets stil op 0: ontbrekend = "n.b.", met een melding op Controle.
- Na afloop:
  - `docs_content.py` bijwerken: FAQ "Hoe is de AK-dekking berekend?" en de mapping-regel "RDFO › AK-dekking";
  - README en `docs/Overdracht_Copilot.md` bijwerken;
  - `release.py` draaien;
  - alleen `RDFO_template.xlsm` opleveren.
- De macro's zijn alleen buiten Excel gecontroleerd (syntax/oletools). Test de knoppen Bronnen inlezen en Prognose vastleggen in Excel zelf.
