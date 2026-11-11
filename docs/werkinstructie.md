# Werkinstructie · kwartaal-FO met het FO-werkbestand

Alles gebeurt in één bestand. Zet het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v22.pptx`) in dezelfde map.

Kleuren:
- **geel** (met goudrand): zelf invullen;
- **wit**: berekend;
- **lichtgrijs**: uit een export (SAP / Salesforce);
- **lichtblauw**: totaal;
- **donkerblauw**: kop.

Op `Start` (kolom M) staan alle tabbladen in werkvolgorde met een link; elk tabblad heeft linksboven "← Start".

## Eenmalig: een project overzetten
1. Kopieer `FO_werkbestand_sjabloon.xlsm`, geef het de projectnaam en open het (macro's toestaan).
2. `Start` → **Oud FO inlezen...** → kies het FO van vorig kwartaal. De macro zet het om:
   - exports en prijslijst;
   - woningtypes en termijnschema;
   - planning;
   - verkoop (verkooptempo) en transport, in aantallen;
   - kosten, DAEB en de FD-kolom van het PEW.
3. Bekijk `Controle oud FO`: per periode de cashflow van het oude FO naast het werkbestand. Lees ook de meldingen van de macro, bijvoorbeeld bedragen die handmatig zijn gezet omdat de prijslijst ze niet verklaarde.
4. Controleer op `Start` de projectnaam, start bouw, het **huidige kwartaal** en oplevering (rij 6 t/m 9). In het huidige kwartaal telt wat al in SAP is geboekt plus wat je nog als prognose invult; de kwartalen ervoor zijn afgesloten (alleen SAP). Op `CF - opbrengsten` staat in D3 het eerste jaar van de kwartalen.

## Elk kwartaal (gelijk aan blok 5 op Start)
1. **Nieuw kwartaal starten:** de cashflow en de planning worden de vorige prognose, FO - huidig gaat naar FO - vorig. Je invoer op Prijslijst, CF - opbrengsten en CF - kosten blijft staan.
2. **Exports inlezen** (knoppen op Start):
   - SAP PEW (FR4) → *FO - huidig*;
   - SAP actuals (FR5) → *FO - actuals*;
   - Salesforce → *Prijslijst dump*.

   De SAP-export voor FO - vorig en FO - huidig mag je ongewijzigd inlezen of plakken: staat hij in de SAP-indeling (een rij lager, een kolom naar links, bedragen als tekst), dan zet de macro hem zelf om en meldt dat.

   Zet de exportdatum op Prijslijst (cel C1).
3. **Prijslijst:**
   - per woningtype het type of blok (kolom C) en eventueel handmatige bedragen;
   - controleer de aansluiting (blok 2, kolom J);
   - boek "te boeken in SAP" (naast de grafieken: per SAP-nummer met de kostendragers);
   - eventueel een kans grondtermijn (kolom Y).
4. **CF - opbrengsten**, per woningtype:
   - planning bouwtermijnen (met de ontwikkelaar);
   - **verkoop** (verkooptempo) en **transport** (tabel Transport) in **aantallen woningen per kwartaal**. Kolom 1 van het verkooptempo is verkocht t/m de exportdatum (Salesforce); kwartalen daarvóór tellen niet nog eens mee;
   - kadastrale kosten en rente (aantal woningen per kwartaal);
   - parkeren/commercieel: verdeling naar rato van de bouwtermijnen of handmatig;
   - eventueel de correctieregel.
5. **CF - kosten:**
   - per type de bouwtermijnen kosten (standaard gelijk aan de opbrengsten);
   - per post *Handmatig*, *Voorstel* (wat nog over is, gelijk verdeeld t/m oplevering) of *Transport* (naar rato van het transport), tot "Nog te verdelen" 0 is.
6. **2. Ontw PEW:** meest recente FD en toelichting. Gemengd project (DAEB): vul op **2B. Ontw PEW** kolom A (concept SKO DAEB) en B (niet-DAEB) in; de controle onderaan moet OK zijn.
7. **Start blok 4:** alle controles op OK.
8. Bekijk **Dashboard**, **Scenario's** en **CF vorig en huidig** (vorige ↔ huidige prognose ↔ realisatie). Met "Grafiek vanaf / t/m" op het Dashboard kies je de periode van de grafieken; die geldt ook voor de PowerPoint.
9. Vul op Start blok 6 de meest recente FD en de PEW-basis in (die komen in de ondertitel van de dia Update PEW) en klik **PowerPoint maken**.
10. **Na het FO:** correcties boeken in SAP en `CF - hulp VENA` inlezen in VENA.
11. Bewaar het werkbestand; volgend kwartaal begin je weer bij stap 1.

## Als er tussentijds iets verandert
Plak of lees de nieuwe export opnieuw in; alles rekent mee, ook het Dashboard en de scenario's.
