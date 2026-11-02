# Prompt: FO + Cashflow-scenario samenvoegen tot één werkbestand

> Versie 2 – gebaseerd op `FO_Parkzoom_fictieve_demoversie.xlsx`, `Cashflow_scenario_leeg_3.xlsm`
> (macro v22) en `Kwartaal_Template_cashflow_v20.pptx`. Open punten staan onderaan.

---

## Rol

Je bent een ervaren projectcontroller vastgoedontwikkeling en Excel/VBA-modelbouwer. Je kent
SAP (PEW / Project Managers Workbench, lay-outs FR4 en FR5), Salesforce-prijslijsten,
koopsomtermijnen (grond, bouwtermijnen, kopersmeerwerk, KAO-opties) en cashflowprognoses per kwartaal.

## Aanleiding

Per kwartaal maken we een FO (financieel overleg) per project. Nu werken we met **twee Excel-bestanden**:

1. **FO-bestand** (`FO_<project>.xlsx`) – de onderlegger:
   - `FO - vorig` en `FO - huidig`: SAP-export PEW, lay-out FR4 (Plan budget, Verplichtingen,
     Tot. werkelijk, Actueel budget, Corr. kosten/opbrengst, PEW resultaat, Toelichting).
     Nu kopiëren we elke periode handmatig "huidig" naar "vorig".
   - `FO - actuals`: SAP-regels lay-out FR5 (kostensoort, kostensoortgroep, waarde, periode, boekjaar).
   - `Prijslijst dump`: Salesforce-export per bouwnummer (grond, aanneemsom, VON, woonwensen, status).
   - `Prijslijst`: SF-gegevens omgerekend naar excl. BTW per bouwnummer.
   - `Aansluiting opbrengsten`: **de brug tussen SF en SAP** – per woningtype SKO ↔ prijslijst ↔
     actueel budget ↔ PEW, met verschillen en het termijnschema (% per termijn).
   - `CF - opbrengsten`: bouwplanning (# woningen per termijn per kwartaal, handmatig), verkooptempo,
     € per termijn per kwartaal, checks.
   - `CF - kosten`: aanneemsom/meerwerk per termijn (volgt bouwplanning) + overige posten per kwartaal.
   - `Verkooptempo`, `Behind the scens` (mapping CODA-postnr → groepering, kostensoortgroep → kosten/omzet,
     periode → kwartaal), `CF - hulp VENA` (cashflow terug naar SAP/VENA).
   - Uitvoer: `1. Cashflow`, `2. Ontw PEW`, `3. Restant bdgt` (+ verborgen `1. Ontw businessc FD`,
     `5. K&R businessc`), `Notities FO gesprek` (to-do's vooraf/achteraf).
2. **Cashflow-scenariobestand** (`Cashflow_scenario.xlsm`) – rekent scenario's door en vult de PowerPoint:
   - `Dashboard`, `Scenario's` (start bouw, verkoop vanaf, transport vanaf, grondtermijn; downside/upside),
     `Invoer`, `Woningtypes`, `CF kosten`, `Model`, `PowerPoint` (welke waarde naar welke vormnaam).
   - Macro's: `UitFOOphalen` (leest het FO-bestand in), `VorigePrognoseUitFO`, `PlanningBewarenAlsVorig`,
     `NaarPowerPoint` (opent sjabloon v20, vult vormen op naam, bewaart nieuwe presentatie).
3. **PowerPoint** `Kwartaal_Template_cashflow_v20.pptx` – 10 dia's: cover, kerncijfers, cashflow per
   kwartaal, verkoopscenario's, scenario's, bouwtermijnen per type, verkoop en transport, van verkoop
   naar omzet, opbrengsten per woningtype, Financieel / Update PEW.

## Doel

**Eén werkbestand (`FO_<project>_<JJJJ>Q<k>.xlsm`)** waarin het FO-bestand en het scenariobestand
samengaan. De lay-out van de PowerPoint blijft gelijk; de cijfers moeten wel volledig en
consistent uit dit ene bestand komen.

## Eisen

### 1. Invoer: plakken óf inlezen, op elk moment opnieuw
- Vaste invoertabbladen, 1-op-1 de export zoals die uit het bronsysteem komt:
  - `SAP vorig` (FR4) en `SAP huidig` (FR4) – de twee tabbladen waar het om draait
  - `SAP actuals` (FR5)
  - `SF prijslijst`
- Per tabblad twee manieren: **zelf plakken** (vanaf cel A1) of een **knop "Inlezen…"** die een
  exportbestand kiest en de inhoud overschrijft.
- **Tussentijds opnieuw plakken moet altijd kunnen**: alle berekeningen zoeken op sleutel
  (WBS / CODA-postnr / bouwnummer / kostensoort), nooit op vaste rij. Rijen erbij of eraf mag.
- Na plakken/inlezen: datum en bestandsnaam van de export op `Start` vastleggen, en een controle
  dat de lay-out klopt (verwachte kolomkoppen aanwezig, periode in de kop = verwachte periode).
- "Vorig" wordt niet meer gemaakt door te kopiëren: bij een nieuw kwartaal plak/lees je de oude
  export in `SAP vorig`. Optioneel knop **"Huidig → Vorig"** die dat in één klik doet.

### 2. Aansluiting opbrengsten blijft de spil
- `Aansluiting opbrengsten` koppelt SF (prijslijst) aan SAP (actueel budget/PEW) per woningtype,
  zoals nu: SKO ↔ prijslijst ↔ actueel budget + correctie ↔ PEW ↔ verschil, met OK/FOUT-checks.
- Woningtypes, aantallen, koopsommen en het termijnschema uit dit tabblad zijn de **enige bron**
  voor `CF - opbrengsten`, `CF - kosten` en de scenario's (geen dubbele invoer meer op `Woningtypes`
  van het scenariobestand: die leest hier rechtstreeks uit).

### 3. CF-opbrengsten en CF-kosten: voorstel + handmatig
Per regel en per prognosekwartaal drie lagen, gerealiseerde kwartalen komen altijd uit `SAP actuals`:

| Laag | Inhoud | Opmaak |
|------|--------|--------|
| **Voorstel** | berekend | grijs, formule |
| **Handmatig** | eigen getal, overschrijft het voorstel | geel invoerveld |
| **Definitief** | `ALS(Handmatig<>"";Handmatig;Voorstel)` | gaat naar cashflow en PowerPoint |

Plus kolom **Toelichting** (verplicht bij handmatig) en een teller "aantal handmatige cellen".

Voorstelregels:
- **Opbrengsten**: bouwplanning (# woningen per termijn per kwartaal) × € per woning per termijn
  uit `Aansluiting opbrengsten`, begrensd door het verkooptempo (zoals de huidige formules in
  `CF - opbrengsten`). Verkocht/te koop volgt uit `SF prijslijst` (status).
- **Bouwkosten (aanneemsom, meerwerk)**: zelfde termijnschema en bouwplanning (zoals nu in `CF - kosten`).
- **Overige kosten**: restant = actueel budget − werkelijk (uit `SAP huidig`), lineair verdeeld over
  de resterende kwartalen t/m oplevering, of "volgt start bouw" / "volgt verkoop"
  (keuze per post, zoals kolom *Schuift mee met* in het scenariobestand).
- Het **bouwplanning-blok** (# woningen per termijn per kwartaal) blijft handmatige invoer: dat is
  de kern van de prognose.

### 4. Scenario-analyse
- De scenariomotor uit `Cashflow_scenario.xlsm` (`Dashboard`, `Scenario's`, `Model`) blijft
  zoals hij werkt, maar leest de basis direct uit de definitieve CF van dit bestand (geen
  `UitFOOphalen` meer nodig voor het eigen project).
- Basis = definitieve CF. Scenario's = afwijkingen op start bouw, verkoop vanaf, transport vanaf,
  grondtermijn, rente; downside/upside-band.
- Uitkomsten per scenario: max. financieringsbehoefte, dieptepunt, break-even, eindsaldo, rente.

### 5. Vorige prognose bewaren
- Knop **"Kwartaal afsluiten"**: bewaart de definitieve CF (opbrengsten, kosten, cumulatief),
  de planning (start bouw per type, fee-kwartalen) en de KPI's als waarden in `Vorige prognose`.
- Die vorige prognose voedt "verschil met vorige prognose" op dia 2, de lichtblauwe lijn op dia 3
  en lichtrood op dia 6 (nu: `VorigePrognoseUitFO` / `PlanningBewarenAlsVorig`).

### 6. Oud FO inlezen (overstap)
- Knop **"Oud FO inlezen…"**: kiest een FO-bestand in de huidige opzet en neemt alle **handmatige
  invoer** over naar de nieuwe opzet, zodat de werkwijze van vorige kwartalen doorloopt:
  - SAP-tabbladen `FO - vorig`, `FO - huidig`, `FO - actuals`, `Prijslijst dump` → nieuwe invoertabbladen
  - `Aansluiting opbrengsten`: aantallen, grond/ontwikkelkosten per type, termijnpercentages, correcties
  - `CF - opbrengsten`: bouwplanning per termijn per kwartaal, verkooptempo, overige opbrengsten, DAEB-blok
  - `CF - kosten`: percentages, onderdelen per kwartaal → als **handmatige** waarden (zodat niets verandert
    t.o.v. het oude FO; het voorstel staat ernaast ter vergelijking)
  - `2. Ontw PEW` kolom "Meest recente FD", `1. Ontw businessc FD`, `5. K&R businessc`,
    `Notities FO gesprek`, toelichtingen in `FO - huidig` kolom T
  - de cashflow uit `1. Cashflow` als "vorige prognose"
- Bouw voort op de bestaande `UitFOOphalen` / `FOVorig` / `FOKosten` / `FODaeb`-logica; die
  leest dezelfde tabbladen al.
- Na inlezen: een **controlerapport** – totaal opbrengsten, totaal kosten en eindsaldo cashflow
  oud FO vs. nieuw bestand per kwartaal; verschil moet 0 zijn of verklaard.

### 7. PowerPoint: zelfde lay-out, beter matchen
- Sjabloon v20 blijft ongewijzigd. Alles via tabblad `PowerPoint` (vormnaam → waarde) en `NaarPowerPoint`.
- **Eén periodecel** (kwartaal + "cijfers per"-datum) op `Start` stuurt alle dia's. In het huidige
  sjabloon staan nu drie verschillende standen (dia 1 "Q3 2026", dia 2 "Stand per Q4 '26",
  dia 3 "Stand Q3 '26"); dat mag niet meer kunnen.
- **Dia 10 "Financieel – Update PEW"** heeft nu geen vormnamen en wordt dus handmatig gevuld. Geef de
  vormen namen (alleen naam, geen lay-outwijziging) en vul ze vanuit `2. Ontw PEW` / `3. Restant bdgt`.
- Controle voor export: elk getal op de dia's is terug te vinden op tabblad `PowerPoint`, en het
  eindsaldo/max. financiering op de dia's = de waarden op `1. Cashflow` en `Dashboard`.

### 8. Startscherm en checks
- `Start`: project, kwartaal, "cijfers per"-datum, status per invoertabblad (geplakt op, periode, OK/FOUT),
  knoppen (Inlezen…, Huidig → Vorig, Oud FO inlezen…, Kwartaal afsluiten, Naar PowerPoint),
  en de to-do-lijst uit `Notities FO gesprek`.
- Alle bestaande checks blijven (Aansluiting OK/FOUT, CF-checks per termijn) en komen samen op `Start`,
  plus: niet-gemapte CODA-postnrs/kostensoortgroepen, SF-bouwnummers zonder type, som CF-kosten vs. PEW.

### Technisch
- Excel 365, Nederlandse instellingen; `.xlsm`; macro's late binding (Windows en Mac), zoals nu.
- Kleuren zoals nu: geel = handmatig, overig formule (zie `Toelichting onderlegger`).
- Geen getallen in formules behalve vaste constanten op `Behind the scens` (o.a. BTW 21%).
- Sleutel-lookups met `X.ZOEKEN` / `SOMMEN.ALS`, geen vaste rijverwijzingen naar de dumps.
- Behoud tabbladnamen die de macro gebruikt, of pas de constanten in de macro mee aan.

## Werkwijze voor jou (AI)

1. Lees beide werkbestanden en de PowerPoint en maak een **kaart**: per tabblad wat invoer, wat formule
   en wat uitvoer is, en welke verbinding er tussen FO en scenariobestand loopt (`UitFOOphalen`).
2. Stel eerst de open vragen (zie hieronder) – **niet gokken**.
3. Lever een **tabbladenoverzicht** van het nieuwe bestand (naam, doel, bron, gebruikt door).
4. Bouw daarna het bestand met de fictieve Parkzoom-data, lees het demo-FO in via "Oud FO inlezen"
   en laat met het controlerapport zien dat de uitkomst gelijk is aan het oude FO.
5. Lever een werkinstructie van één A4: "kwartaal-FO in 8 stappen", aansluitend op `Notities FO gesprek`.

## Open vragen

1. Blijft `SAP actuals` (FR5) een apart tabblad, of wil je echt alleen `SAP vorig` en `SAP huidig`?
2. "Uploaden": bedoel je een knop die een bestand kiest (macro), of ook Power Query (inlezen uit een map)?
3. Overige kosten: welk voorstel is het meest realistisch – restant budget lineair t/m oplevering,
   of volgens de verdeling van de vorige prognose?
4. Moet `CF - hulp VENA` (cashflow terug naar SAP) ook uit de definitieve CF komen?
5. Is er nog een Python-versie (`forecast_builder/fo.py`, genoemd in de macro) die we moeten
   meenemen of uitfaseren?
