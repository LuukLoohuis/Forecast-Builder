# Werkinstructie · kwartaal-FO met het FO-werkbestand

Alles gebeurt in één bestand. Zet het PowerPoint-sjabloon (`Kwartaal_Template_cashflow_v21.pptx`) in dezelfde map.
Gele cellen vul je zelf in, lichtblauwe cellen komen uit SAP of SF, grijze cellen worden berekend.

## Eenmalig: overstappen vanuit een oud FO
1. Open het werkbestand en sta macro's toe.
2. Klik op `Start` op **Oud FO inlezen...** en kies het FO van vorig kwartaal.
3. Controleer tab `Controle oud FO`: de cashflow moet per periode gelijk zijn ("OK"). De transportregels, de methode per
   kostenpost en de correctieregel staan daarna klaar met standaardwaarden.
4. Vul op `Start` het opleveringskwartaal in (blok 1) en sla op onder een nieuwe naam.

## Elk kwartaal

**Voor het FO-gesprek**
1. **Nieuw kwartaal starten** (Start): FO - huidig wordt FO - vorig, de cashflow wordt de vorige prognose.
   Bewaar zo nodig ook de planning: knop *Nu bewaren als vorige prognose* op `Woningtypes`.
2. SAP PEW-export (FR4) van de huidige periode → rij *FO - huidig*: **Plakken** (eerst de export kopiëren vanaf A1)
   of **Inlezen...** (bestand kiezen).
3. SAP actuals (FR5) → rij *FO - actuals*. Check SAP op correcties in de werkelijke opbrengsten.
4. Salesforce prijslijst → rij *Prijslijst dump*. Zet de exportdatum op tab `Prijslijst` (cel C1).
5. `Aansluiting opbrengsten`: SKO ↔ prijslijst ↔ PEW sluit aan (kolom J "OK"). Vraag ontbrekende documentatie op
   (SKO gunning, termijnschema's, afrekenstaat MMW, specificatie overige opbrengsten).
6. **Opbrengsten** op `CF - opbrengsten`, per woningtype:
   - bouwtermijnen per kwartaal;
   - daaronder de **transportregel** (aantal woningen notarieel geleverd per kwartaal). De grondtermijn volgt het transport, en bouwtermijnen tellen vanaf transport;
   - verkooptempo in rij 49-52;
   - wil je een bedrag bijsturen? Gebruik de gele correctieregel (rij 192).
7. **Kosten** op `CF - kosten`:
   - de bouwtermijnen volgen de geplande bouwtermijnen van de opbrengsten (overschrijven mag);
   - bij de posten kies je in kolom C *Handmatig* (je eigen verdeling, zoals altijd) of *Voorstel* (restantbudget gelijk verdeeld t/m oplevering, blok onderaan).

   Rij 75 is het definitieve totaal.
8. Blok 4 op `Start`: alle controles groen. Bekijk `Dashboard` (cashflow en kerncijfers) en `2. Ontw PEW`.
9. **Scenario's**: vul op `Scenario's` een andere start bouw, verkoop, transport of grondtermijn in. `Invoer`, `Woningtypes`
   en `Scenario kosten` rekenen automatisch mee met de CF-tabbladen; daar hoef je niets te doen.
10. **PowerPoint maken**: alle dia's, inclusief de datum op dia 1 en de
   kaarten van dia 10 (blok 6 op `Start`), uit dit bestand.

**Tijdens het gesprek:** aantekeningen op `Notities FO gesprek`.

**Na het FO-gesprek**
1. Correcties verwerken (prognose, toelichtingen) en opnieuw **PowerPoint maken**.
2. `CF - hulp VENA` inlezen in SAP/VENA (zelfde format als altijd; loopt mee met de definitieve cashflow).
3. Datumvelden Projectbuilder en de verdeling van woningaantallen in de tijd bijwerken.

## Als er tussentijds iets verandert
Plak of lees de nieuwe export opnieuw in; alles rekent mee, ook het Dashboard en de scenario's. Klik daarna op
**PowerPoint maken** als de presentatie bijgewerkt moet worden.
