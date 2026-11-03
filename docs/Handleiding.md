# Handleiding RDFO-prototype

Met dit bestand maak je elk kwartaal de RDFO, het dashboard, het stoplicht, de bandbreedte en de PowerPoint. Je hebt twee exports uit SAP nodig: ZPMW03 en Woningetalage. Je leest ze in met één knop, kijkt of er fouten zijn, vult aan wat niet in SAP staat en zet de prognose vast. Een vastgezette prognose verandert daarna niet meer.

## Elk kwartaal in 5 stappen

Begin je met een leeg bestand? Gebruik `RDFO_template.xlsm` of de knop *Bestand leegmaken* op Start.

**1. Inlezen** — Knop op Start

Klik op 'Bronnen inlezen'. Kies de ZPMW03-export, daarna de Woningetalage-export en tot slot (als die er is) de AK-percentages uit SAP. Vul de peildatum in (bijv. 30-09-2026) en de versie (bijv. Q3).

*Let op:* Het bestand opent daarna het tabblad Controle. Eerder vastgezette prognoses blijven zoals ze waren.

**2. Controleren** — Tabblad Controle

Kijk bovenaan het tabblad Controle: staat er een FOUT? Los die eerst op. Staan er nieuwe projecten? Klik op 'Nieuwe projecten toevoegen' en kies in Invoer › 2 of het Opstal of Deelneming is.

*Let op:* Achter elke controle staat wat je moet doen. Onderaan staan de lijsten met nieuwe, verdwenen en verliesgevende projecten.

**3. Aanvullen** — Tabbladen Bandbreedte en Invoer

Vul de gele cellen in: op het tabblad Bandbreedte het stoplicht (vergunning, netcongestie) en eventueel kans %; op Invoer wat niet in SAP staat (taakstellingen, toelichtingen, kansen en risico's, DAEB, reserves).

*Let op:* Niet ingevuld? Dan staat er 'n.b.' (niet bekend). Het bestand vult nooit stiekem een nul in.

**4. Vastzetten** — Knop op Start

Klik op 'Prognose vastleggen'. Je ziet de omzet en het aantal woningen. Klik op OK. Daarna kun je meteen opslaan.

*Let op:* Is deze versie al vastgezet? Dan vraagt het bestand waarom je hem wilt aanpassen. De oude cijfers blijven bewaard.

**5. Presenteren** — Dashboard of Start

Kies op het Dashboard het jaar en de versie. Klik op 'Maak PowerPoint'.

*Let op:* De PowerPoint komt in dezelfde map als dit bestand. Alle tabellen zijn gewone PowerPoint-tabellen: je kunt ze nog aanpassen.

## De tabbladen

| Tabblad | Wat staat er | Wat doe je er |
|---|---|---|
| Start | Knoppen en status. | Hier begin je. |
| Dashboard | Kerncijfers en grafieken. Kies hier jaar, versie, opstal/deelneming of één project. | Gele cellen kiezen. |
| RDFO | De RDFO: budget, Q1–Q4, upload, realisatie, to go en verschillen. Onderaan: per gebied. | Alleen lezen. |
| Outlook | De 4 jaren na het lopende jaar: omzet (ook DAEB, funnel, target), BOR, AK-dekking, kansen en risico's, reserves, bandbreedte, woningen, orderintake. | Alleen lezen. Basis voor het budget. |
| Projecten | Per project: alle versies, realisatie, to go, BOR, marge en stoplicht. | Filteren met de pijltjes. |
| Projectdetail | Eén project in detail. | Project kiezen bovenaan. |
| Woningen | Verkopen per maand en kwartaal, per project per kwartaal (dit en volgend jaar), per versie, per productsegment en koperstype. | Alleen lezen. |
| Bandbreedte | Bovenaan de vijf scenario's per jaar met grafiek. Daaronder per project het stoplicht en de scenario's. | Gele cellen invullen. |
| Controle | Alle controles. Rood = eerst oplossen. | Na elke keer inlezen bekijken. |
| Invoer | Instellingen, projectlijst en alles wat niet in SAP staat. | Gele cellen invullen. |
| Handleiding | Deze uitleg. |  |

## Stoplicht invullen

- Per project vul je in: omgevingsvergunning (Ja / In procedure / Nee) en netcongestie (Geen / Laag / Midden / Hoog).
- Verkoop kleurt vanzelf op het percentage verkochte woningen. Wil je het percentage bij start bouw gebruiken? Vul dan 'Verkocht % SB' in.
- Rechts op het tabblad staan de uitgangspunten: vanaf welk percentage het groen of oranje is, en welke keuze welke kleur krijgt. Pas je dat aan, dan kleuren alle lichten meteen mee.
- Een leeg rondje betekent: nog niet ingevuld of niet van toepassing.

## Bandbreedte: vijf scenario's

- Het bestand rekent per project vijf scenario's uit, voor het lopende jaar en de twee jaren daarna. Bovenaan het tabblad Bandbreedte zie je het totaal per jaar; de grafiek toont per jaar een band van het laagste tot het hoogste scenario.
- Optimistisch: de verkoop haalt aan het eind van het jaar minstens het ingestelde percentage (standaard 90%). De omzet groeit dan mee.
- Basis: de omzet uit de prognose (upload).
- Afslag: maar een deel van de geplande verkopen van dat jaar lukt (standaard 50%). De omzet daalt dan mee.
- Gewogen: omzet x kans %. De kans komt uit ZPMW03; je kunt hem per project zelf invullen (kolom 'Kans %').
- 0-100: een project telt helemaal mee of helemaal niet. Standaard telt het mee vanaf 50% kans; per project kun je Ja of Nee kiezen.
- Wat al verkocht en gerealiseerd is, verandert in geen enkel scenario. De instellingen staan op het tabblad onder de tabel.

## Af en toe invullen (tabblad Invoer)

| Onderdeel | Wat vul je in |
|---|---|
| 2. Projecten | Eén regel per project. Kies Type, Gebied, Strapo (Ja/Nee) en hoe de woningen verdeeld zijn over de productsegmenten en kopers. AK: het percentage komt uit de SAP-dump (kolom V); zet 'AK meenemen' op Nee als het niet mee moet, of vul in kolom K een eigen percentage in. |
| 7. Gebieden | De lijst met gebieden. Die kun je per project kiezen. |
| 8. Funnel en target | Per jaar de omzet uit de acquisitiefunnel en de target. Ook de deler voor de AK-dekking (standaard 1,05). |
| 9. DAEB-omzet | Per project de DAEB-omzet en welk deel in welk jaar valt (samen 100%). |
| 10. Kansen en risico's | Per kans of risico een regel: kies het project, vul omschrijving, jaar en kwartaal, soort en bedrag in. Bijvoorbeeld: 2026 Q4, project X, hogere koopsom, € 1.500.000 (bedrag altijd positief; een risico telt als min). |
| 11. Reserves | Per project de reserves en hoe ver het project klaar is. Het bestand rekent het effect per jaar uit. |

## Een vastgelegde prognose corrigeren

1. Pas de bron of de aanvullende invoer aan zodat de cijfers kloppen.
2. Klik op 'Prognose vastleggen'. De knop ziet dat de versie al bestaat en vraagt om een toelichting.
3. De correctie krijgt een hogere revisie. Alleen de nieuwste revisie telt; de oude regel blijft bewaard. In de RDFO staat ● achter de versie.

## Spelregels

- Wijzig nooit iets in de Historie. Alleen de knop 'Prognose vastleggen' voegt regels toe.
- Voeg nieuwe projecten altijd onderaan toe in Invoer › 2 (of met de knop). Niet sorteren of tussenvoegen: het stoplicht hoort bij de regel.
- Bedragen in de bron zijn hele euro's; de rapportage toont EUR x 1.000 of EUR mln (instelbaar).
- Omzet en resultaat zijn positief, verlies is negatief (rood). Er wordt nooit een absolute waarde genomen.
- Woningaantallen en bedragen zijn al op aandeelbasis; het aandeel wordt niet nog eens toegepast.

## Problemen oplossen

| Probleem | Oplossing |
|---|---|
| Ik wil met een leeg bestand beginnen | Gebruik RDFO_template.xlsm (al leeg), of klik op Start op 'Bestand leegmaken' en typ LEEG. Daarna: Bronnen inlezen › Nieuwe projecten toevoegen › Type en gebied kiezen. |
| Na inlezen is alles leeg of staat 'Nieuwe projecten: … FOUT' | De WBS-elementen uit je export staan nog niet in de projectlijst; alleen projecten in Invoer › 2 tellen mee. Klik op 'Nieuwe projecten toevoegen' en controleer per regel Type en Woningen meetellen. |
| Controle geeft waarschuwingen over woningen of bedragen (aantal, etalage, cumulatief, portefeuille, PEW) | Dat zijn verschillen in de SAP-export zelf; jij hoeft ze niet op te lossen om vast te zetten. Onderaan Controle staat per project welke controle afwijkt. Meld die projecten bij de verkoopadministratie of projectcontrol. |
| Controle meldt 'ontbrekende kolommen' | Het bestand zoekt kolommen op naam (zie Handleiding › Veldwoordenboek). Na inlezen noemt de melding welke kolomnamen niet gevonden zijn. Hernoem de kolom in de export of pas de naam aan in de export-layout in SAP. |
| De knoppen doen niets / gele balk 'Macro's zijn uitgeschakeld' | Klik op 'Inhoud inschakelen'. Komt het bestand uit e-mail of internet: sluit het, klik met rechts op het bestand › Eigenschappen › vink 'Blokkering opheffen' aan. |
| Mijn organisatie staat geen macro's toe | Alles rekent ook zonder macro's. Rechtermuisknop op een tabbladnaam › Zichtbaar maken › Bron_ZPMW03 / Bron_Woningetalage: plak de export in A1. Vastleggen: kopieer op het verborgen tabblad Concept A8:R(laatste regel) en plak als waarden op de eerste vrije rij van Historie. |
| 'Maak PowerPoint' geeft een foutmelding | PowerPoint moet op deze computer geïnstalleerd zijn. Sluit open dialoogvensters in PowerPoint en probeer opnieuw. |
| Er staat 'n.b.' in de RDFO | Er is (nog) geen vastgelegde waarde of aanvullende invoer. Dat is bewust geen nul. |
| Een oranje subtotaal in de RDFO | Minstens één onderliggende post is n.b.; het subtotaal is dus onvolledig. |
| Nieuw rapportagejaar | Zet in Invoer › 1 het rapportagejaar, het versiejaar van de upload en het versiejaar in de rapporten op het nieuwe jaar. Kolomkoppen schuiven mee. |
| Ik wil de rekenbladen of de Historie zien | Klik op Start op 'Toon / verberg rekenbladen'. Ook Historie (alle vastgezette prognoses) staat daar. |
| Waar is het tabblad Stoplicht? | Het stoplicht staat nu op het tabblad Bandbreedte, samen met de scenario's per project. |
| Excel meldt bij openen dat inhoud is hersteld en de knoppen werken niet | Importeer de macro's handmatig: Alt+F11 › Bestand › Bestand importeren › macros/modRDFO.bas. Klik daarna met rechts op elke knop › Macro toewijzen. |

## Waarom zo gebouwd

- **Macro's** — Drie knoppen nemen het handwerk over: bronnen inlezen (bestandskiezer), prognose vastleggen en PowerPoint maken. Alle berekeningen blijven gewone Excel-formules, zodat het bestand ook zonder macro's rekent en controleerbaar blijft.
- **Waarom geen Power Query** — Met de knop 'Bronnen inlezen' is het inlezen al één klik. Power Query zou een tweede mechanisme toevoegen zonder extra voordeel; de staginglaag zoekt kolommen al op naam.
- **Grootboek** — Elke vastgelegde waarde is één regel met versie, revisie, peildatum, doeljaar, kengetal en waarde. Nieuwe versies vragen geen nieuwe kolommen.
