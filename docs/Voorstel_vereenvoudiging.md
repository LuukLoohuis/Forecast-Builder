# Voorstel vereenvoudiging RDFO

Dit voorstel volgt `docs/Prompt_vereenvoudiging.md`.

**Status (uitgevoerd):**
- **Uitgevoerd:**
  - nr 1–5;
  - nr 8 (443 → 335 namen; ongebruikte namen worden voortaan bij elke build automatisch opgeruimd);
  - nr 10 (demo-opties verborgen in de template);
  - nr 13 (STG_AKR rekent alleen de regels die er zijn).
- **Op verzoek gewijzigd:** Projectdetail blijft en is weer zichtbaar (nr 6 vervalt).
- **Bewust niet gedaan:** nr 7 (PPT en Dias samenvoegen). Weinig winst voor de gebruiker en risico voor de PowerPoint-macro, die buiten Excel niet te testen is.
- **Geen actie nodig:**
  - nr 9: op de Outlook staan alleen 2 regels bandbreedte (laagste/hoogste), geen dubbele tabel;
  - nr 11, 12 en 14: blijven zoals ze zijn.
- **Extra:** vier AK-controles (grootboekdump: ontbrekende kolommen, geen boekingen in het jaar, negatief saldo, geen boekingsperiode); projecten zonder AK % is nu een WAARSCHUWING.

## Inventarisatie (huidige stand)
- **Tabbladen:** 14 zichtbaar, 16 verborgen.
  - Zichtbaar: Start, Dashboard, RDFO, Projecten, Woningen, Outlook, Bandbreedte, Kansen en risico's, Reserves, Projectlijst, AK-dekking, Controle, Invoer, Handleiding.
  - Verborgen: Projectdetail, Historie, Concept, VERW_Project, CALC_Versies, SCEN, PPT, Dias, Lijsten, plus 4 Bron-bladen en 3 STG-bladen.
- **Invoer op 4 plekken:**
  - Invoer (secties 1–9);
  - Projectlijst;
  - Kansen en risico's;
  - de stoplicht- en scenario-invoer, die op het tabblad Bandbreedte staat.
- **13 knoppen**, waarvan 6 dubbel.
- **58 controles:** 9 INFO, de rest OK/FOUT/WAARSCHUWING. Ongeveer 25 daarvan zijn technische aansluitcontroles die nooit een actie van de gebruiker vragen.
- **443 namen:** ongeveer 80 worden nergens meer gebruikt. Dat zijn restanten van eerdere rondes, zoals scenariohulpnamen, kostenkolommen en oude mapping-namen.

## Voorstellen

| Nr | Onderdeel | Wat is dubbel / overbodig | Voorstel | Wat de gebruiker merkt | Risico | Werk |
|---|---|---|---|---|---|---|
| 1 | Projectgegevens | Twee plekken: de Projectlijst (WBS, naam, type, gebied, keuzes, AK %, reserves) en Invoer › 2 (projectdefinitie, woningen, aandeel, strapo, segment-%, belegger-%) | **Samenvoegen** op de Projectlijst. De details staan in een inklapbare groep. Invoer › 2 wordt een verborgen doorverwijzing. | Eén plek per project; Invoer is een sectie korter | Laag: dezelfde regelvolgorde, de formules lezen via dezelfde namen | M |
| 2 | Aantal woningen | Wordt ingevoerd in Invoer › 2 én komt uit de Woningetalage, met een eigen controle "wijkt af van Mapping" | **Verwijderen** als invoer: de Woningetalage is leidend. Het aantal blijft zichtbaar als grijze kolom. | Eén veld minder, één waarschuwing minder | Laag: de rapportage gebruikt nu al de Woningetalage | S |
| 3 | Invoer per project op het tabblad Bandbreedte | Daar staan 6 invoerkolommen per project: verkocht % (handmatig), vergunning, netcongestie, toelichting, kans % en 0-100 Ja/Nee. Daarom zijn nodig: een stempel per regel, een eigen VBA-module en de controle "lijst verschoven" | **Verplaatsen** naar de Projectlijst, als groep "Stoplicht en scenario" | Alles per project op één plek. Het tabblad Bandbreedte wordt alleen-lezen (lichten, scenario's en grafiek). | Laag: de stempel, de module en de controle vervallen, want de regel is altijd juist | M |
| 4 | Knoppen | "Nieuwe projecten toevoegen" 3x (gebeurt al automatisch), "Bronnen inlezen" 2x, "Maak PowerPoint" 2x, "Prognose vastleggen" 2x | **Opruimen**: de hoofdknoppen alleen op Start. "Nieuwe projecten" alleen op de Projectlijst. Controle houdt "Bronnen opnieuw inlezen" en "Toon alle controles". Dashboard en Historie krijgen geen knoppen. | 13 → 8 knoppen | Geen | S |
| 5 | Controle | 58 controles. Ongeveer 25 zijn technische aansluitingen: Rapport, Aansluiting, Parameters, Weging, "RDFO = som projecten". Daarnaast INFO-regels zonder actie en een verouderde verwijzing ("Invoer › 10") | **Splitsen** in ±30 gebruikerscontroles (wat moet je doen) en een verborgen blad "Techniek" voor de aansluitcontroles. Een FOUT op het verborgen blad komt als één regel bovenaan: "Technische controle mislukt — meld dit". Verwijzingen worden hersteld. | Korter en alles is actiegericht | Laag: de technische controles blijven bestaan en blijven getest | M |
| 6 | Projectdetail | Verborgen tabblad; de informatie staat ook op Projecten en in het rekenvoorbeeld op AK-dekking | **Verwijderen** | Niets (was al verborgen) | Laag | S |
| 7 | PPT en Dias | Twee verborgen bladen voor dezelfde PowerPoint (teksten en spec op PPT, tabellen en grafieken op Dias) | **Samenvoegen** tot één blad "PowerPoint" | Bij "Toon rekenbladen" een blad minder | Laag: alleen verwijzingen | M |
| 8 | Ongebruikte namen (±80) | Restanten van eerdere rondes | **Verwijderen**, met een test die controleert dat elke naam gebruikt wordt | Niets; het bestand wordt iets kleiner en overzichtelijker voor beheer | Geen | S |
| 9 | Overlap in rapporten | Dashboard, RDFO en Projecten tonen deels dezelfde kerncijfers. De Outlook toont ook bandbreedte, woningen en orderintake, die een eigen tabblad of dia hebben. | **Behouden**, maar elk tabblad krijgt een duidelijke rol: Dashboard = in één oogopslag, RDFO = de officiële tabel, Projecten = per project, Outlook = 4 jaar vooruit. Op de Outlook wordt de dubbele bandbreedte-tabel vervangen door één regel met een verwijzing naar Bandbreedte. | Minder herhaling op de Outlook | Laag | S |
| 10 | Demo-opties (P_Demo, status "Fictieve aanname"/"Voorbeeld") | Alleen nodig voor het voorbeeldbestand | **Uit de template halen**; ze blijven alleen in het prototype en de tests | Eén instelling en twee statussen minder | Laag | S |
| 11 | Handleiding | Staat als tabblad en in docs/Handleiding.md | **Behouden**: beide komen al uit één bron (docs_content.py) | — | — | — |
| 12 | Invoersecties 5, 6, 8, 9 (koperstype, jaarkolommen, funnel/target, DAEB) | Weinig gebruikt, maar wel door jou gevraagd | **Behouden**; sectie 6 krijgt de kop "expert — normaal niet aanpassen" en wordt standaard ingeklapt | Invoer oogt korter | Geen | S |
| 13 | Zware formules | STG_AKR met 3.000 regels (index en match per regel); SUMPRODUCT over hele bereiken op Controle | **Versnellen**: STG_AKR beperken tot het werkelijke aantal regels (de macro schrijft het aantal weg) | Sneller herberekenen | Laag | M |
| 14 | Verborgen rekenbladen Concept, CALC_Versies, VERW_Project, SCEN | Elk heeft een eigen rol (vastleggen, versies, per project, scenario's). Geen echte dubbeling gevonden. | **Behouden** | — | — | — |

## Top 5: grootste winst, kleinste risico
1. **Nr 4 – Knoppen opruimen** (13 → 8): direct rustiger, geen risico.
2. **Nr 1 + 2 – Alle projectgegevens op de Projectlijst**, aantal woningen alleen uit de Woningetalage: één plek per project.
3. **Nr 3 – Stoplicht- en scenario-invoer op de Projectlijst**: een VBA-module, een stempel en een controle minder. Het tabblad Bandbreedte wordt alleen-lezen.
4. **Nr 5 – Controle splitsen** in gebruikerscontroles en techniek: ±30 regels die allemaal een actie vragen.
5. **Nr 6 + 7 + 8 – Opruimen achter de schermen**: Projectdetail weg, PPT en Dias samen, ±80 ongebruikte namen weg.

Resultaat: 13 zichtbare tabbladen (14 nu). De invoer staat op 3 plekken in plaats van 4. Er blijven 8 knoppen over en ±30 controles. De functionaliteit verandert niet.

## Vragen
1. Akkoord met de top 5? En wil je nr 9 (Outlook), nr 10 (demo-opties) en nr 13 (snelheid) ook meenemen?
2. De stoplicht- en scenario-invoer (6 kolommen) naar de Projectlijst: akkoord? De kolommen staan dan in een eigen inklapbare groep, zodat de lijst rustig blijft.
3. Gebruik je Projectdetail ooit? Zo niet, dan gaat het eruit.
4. Bestaan er al gevulde bestanden die moeten blijven werken? Dan maak ik een eenmalige overzetmacro voor de stoplicht- en projectgegevens.
