# Prompt: RDFO-bestand vereenvoudigen (zonder functionaliteit te verliezen)

## Rol en doel
Je bent een ervaren financieel controller en Excel/VBA-bouwer. Beoordeel het complete RDFO-bestand
(`RDFO_template.xlsm`, gegenereerd door `tools/build_rdfo.py` en `macros/modRDFO.bas`) op eenvoud.
Doel: het bestand moet voor een regiocontroller **makkelijker in gebruik en beter te overzien** zijn.
Daarvoor zoek je **dubbelingen** (dezelfde invoer, berekening, tabel, knop of controle op meer dan één plek)
en **opties die eruit kunnen**. De functionaliteit en de betrouwbaarheid mogen niet minder worden.

## Wat het bestand moet blijven kunnen (niet onderhandelbaar)
1. ZPMW03, Woningetalage, AK-percentages en de AK-grootboekdump inlezen met één knop. Nieuwe projecten worden daarbij automatisch toegevoegd.
2. Per project kiezen wat meetelt: prognose, woningen, AK-dekking en PowerPoint.
3. De RDFO per versie tonen (Budget, Q1–Q4, upload). Daarin: omzet, BOR, BOR % (incl. fees, overige resultaten en taakstelling EBIT), AK-dekking (per kwartaal naloopbaar) en EBIT.
4. Prognoses vastleggen in de Historie. Vastgelegde versies blijven onveranderlijk, met correcties via revisie.
5. Oude prognoses (Budget, Q1–Q3) eenmalig overzetten.
6. Outlook 4 jaar vooruit, bandbreedte (5 scenario's), kansen en risico's, reserves, gebieden, DAEB, funnel/target, woningen per project per kwartaal, orderintake en stoplicht.
7. Controles die fouten zichtbaar maken. Niets mag stil wegvallen: n.b. is nooit 0.
8. Een PowerPoint maken met één knop.
9. Een leeg template en de knop "Bestand leegmaken".
10. Alle tests in `tools/test_rdfo.py` blijven slagen (nu 31/31).

## Werkwijze
1. **Inventariseer** alle tabbladen, de zichtbare en de verborgen, met per tabblad het doel, de invoer en de uitvoer. Inventariseer ook alle knoppen en macro's, alle invoersecties (Invoer › 1–9, Projectlijst, Kansen en risico's) en de controles op het tabblad Controle.
2. **Zoek dubbelingen** in vijf soorten:
   - dezelfde invoer op twee plekken;
   - dezelfde berekening op twee plekken;
   - dezelfde informatie in twee rapporten zonder duidelijk eigen doel;
   - dezelfde knop op meerdere tabbladen;
   - controles die hetzelfde meten, of die alleen technisch zijn en geen actie van de gebruiker vragen.
3. **Zoek overbodige opties**: functies die (bijna) nooit gebruikt worden, alleen voor de demo bestaan, of door een andere functie al gedekt zijn.
4. **Weeg elk voorstel** af op:
   - winst voor de gebruiker (minder tabbladen, kolommen, klikken of meldingen);
   - risico voor de juistheid van de cijfers;
   - impact op Historie, vastgelegde versies en bestaande bestanden;
   - hoeveel werk het is.
5. **Lever eerst alleen een voorstel op, nog niets bouwen.** Gebruik een tabel met de kolommen: nr · onderdeel · wat is er dubbel/overbodig · voorstel (behouden / samenvoegen / verplaatsen / verwijderen / verbergen) · wat de gebruiker merkt · risico · werk (S/M/L). Sluit af met een top 5 "grootste winst, kleinste risico" en een lijst met vragen voor de gebruiker.
6. Na akkoord: per voorstel bouwen, tests draaien, handleiding/README/overdracht bijwerken en alleen de template opleveren.

## Kandidaten die je in elk geval moet beoordelen
1. **Projectgegevens staan op twee plekken.** De Projectlijst bevat WBS, naam, type, gebied, de 4 keuzes, AK % en reserves. Invoer › 2 bevat projectdefinitie, aantal woningen, aandeel, strapo, productsegment-% en belegger-%. Kan alles op de Projectlijst (details inklapbaar), zodat Invoer › 2 vervalt?
2. **Aantal woningen wordt dubbel ingevoerd.** Het staat in Invoer › 2 én in de Woningetalage, met een aparte controle "wijkt af van Mapping". Kan alleen de Woningetalage leidend zijn?
3. **De stoplichtinvoer staat op het tabblad Bandbreedte.** Daarom is een regelstempel nodig, met de module shStoplicht en de controle "lijst verschoven". Kan het stoplicht als 3 kolommen op de Projectlijst? Dan vervallen de stempel, de module en die controle.
4. **Dubbele knoppen:**
   - "Nieuwe projecten toevoegen" staat 3x (Start, Controle, Projectlijst), terwijl het al automatisch bij het inlezen gebeurt;
   - "Bronnen inlezen" staat 2x;
   - "Maak PowerPoint" staat 2x;
   - "Prognose vastleggen" staat 2x.
5. **Overlappende rapporten.** Dashboard, RDFO, Projecten en Outlook tonen deels dezelfde kerncijfers. Woningen, orderintake en bandbreedte staan zowel op een eigen tabblad als in de Outlook. Welk tabblad is leidend en wat kan weg?
6. **Projectdetail (verborgen).** Weghalen of behouden?
7. **Controle (58 controles, veel INFO):**
   - welke vragen echt een actie van de gebruiker?
   - kunnen technische aansluitcontroles (Rapport, Aansluiting, Parameters) naar de testsuite of een verborgen blad?
   - verouderde verwijzingen herstellen, zoals "Invoer › 10" voor kansen die nu een eigen tabblad hebben.
8. **Verborgen rekenbladen.**
   - Concept, CALC_Versies, VERW_Project en SCEN: zitten er dubbele berekeningen in?
   - PPT en Dias zijn twee bladen voor de PowerPoint: kunnen die samen?
   - STG_AKR heeft 3.000 formuleregels: kan dat kleiner of sneller?
   - De 443 namen: welke worden niet (meer) gebruikt?
9. **Invoersecties die weinig gebruikt worden:**
   - 5 Koperstype;
   - 6 Jaarkolommen Woningetalage (expert);
   - 8 Funnel en target;
   - 9 DAEB.

   Nodig, of beter verplaatsen naar de Outlook of een "expert"-sectie?
10. **Demo-opties in het productiebestand** (P_Demo, fictieve aannames, status "Voorbeeld"): kunnen die uit de template?
11. **Handleiding** staat zowel als tabblad als in docs/Handleiding.md: houd één bron aan.
12. **Bestandsgrootte en snelheid**: zoek volatiele of zware formules (SUMPRODUCT over hele bereiken, INDEX/MATCH over 3.000 regels) die eenvoudiger kunnen.

## Randvoorwaarden
- Taal: Nederlands en kort, voor een controller zonder VBA-kennis.
- Huisstijl Heijmans: geel en donkerblauw.
- Niets mag stil wegvallen: wat verdwijnt, moet aantoonbaar elders gedekt zijn.
- Bestaande gevulde bestanden moeten blijven werken, of er komt een eenmalige overzetmacro.
- Geen nieuwe functionaliteit toevoegen in deze ronde, alleen vereenvoudigen.
