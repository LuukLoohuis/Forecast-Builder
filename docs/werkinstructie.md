# Werkinstructie · kwartaal-FO met het FO-werkbestand

Alles gebeurt in één bestand. Zet het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v21.pptx`) in dezelfde map.
Gele cellen vul je zelf in, lichtblauwe komen uit SAP of Salesforce, grijze worden berekend.

## Eenmalig: een project overzetten
1. Kopieer `FO_werkbestand_sjabloon.xlsm`, geef het de projectnaam en open het (macro's toestaan).
2. `Start` → **Oud FO inlezen...** → kies het FO van vorig kwartaal. De macro zet het om:
   - exports en prijslijst;
   - woningtypes en termijnschema;
   - planning, verkocht/transport (in %);
   - kosten, DAEB en de FD-kolom van het PEW.
3. Bekijk `Controle oud FO`: per periode de cashflow van het oude FO naast het werkbestand. Lees ook de meldingen van de
   macro (bijvoorbeeld bedragen die handmatig zijn gezet omdat de prijslijst ze niet verklaarde).
4. Vul op `Start` de projectnaam en het opleveringskwartaal in. Op `CF - opbrengsten` staat in D3 het eerste jaar van de kwartalen.

## Elk kwartaal

**Voor het FO-gesprek**
1. **Nieuw kwartaal starten** (Start):
   - de cashflow wordt de vorige prognose (stippellijn in de grafiek, tabel vorige/huidig op het Dashboard);
   - de planning (start bouw per type) wordt de vorige planning;
   - FO - huidig wordt FO - vorig;
   - FO - huidig en FO - actuals worden leeggemaakt.

   Je invoer op Prijslijst, CF - opbrengsten en CF - kosten blijft staan.
2. **Exports:** Plakken of Inlezen op `Start`:
   - SAP FO-overzicht (FR4) → *FO - huidig*;
   - SAP actuals (FR5) → *FO - actuals*;
   - Salesforce → *Prijslijst dump*.
3. `Prijslijst`:
   - nieuwe bouwnummers met hun type;
   - exportdatum in C1;
   - controleer per woningtype de aansluiting (kolom J "OK");
   - boek "te boeken in SAP" (per SAP-nummer en kostendrager) in SAP.
4. `CF - opbrengsten`, per woningtype:
   - planning bouwtermijnen (met de ontwikkelaar);
   - **verkocht %** en **getransporteerd %** per kwartaal (van alle woningen van het type);
   - kadastrale kosten en rente (aantal woningen per kwartaal);
   - na goedkeuring SWK eventueel een grondprijsaanpassing.

   Eventueel ook de correctieregel en het DAEB-blok onderaan.
5. `CF - kosten`:
   - per type de planning van de bouwtermijnen kosten (HWB; standaard gelijk aan de opbrengsten);
   - per post *Handmatig*, *Voorstel* (restantbudget gelijk verdeeld t/m oplevering) of *Transport* (naar rato van het transport). Bij Voorstel en Transport gaat een zelf ingevuld kwartaal voor.
6. `2. Ontw PEW`: meest recente FD en toelichting.
7. `Start` blok 4: alle controles groen. Bekijk `Dashboard` en `Scenario's`.
8. **PowerPoint maken.**

**Na het FO-gesprek**
1. Correcties verwerken en opnieuw **PowerPoint maken**.
2. `CF - hulp VENA` inlezen in SAP/VENA (zelfde format als altijd).

## Als er tussentijds iets verandert
Plak of lees de nieuwe export opnieuw in; alles rekent mee, ook het Dashboard en de scenario's.
