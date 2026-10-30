# Handleiding RDFO-prototype

Dit bestand maakt de RDFO-rapportage, het dashboard, de projectoverzichten, het stoplicht en de PowerPoint uit twee SAP-exports: ZPMW03 en Woningetalage. Eén keer per kwartaal lees je de exports in, controleer je, vul je aan en leg je de prognose vast. Vastgelegde prognoses veranderen daarna nooit meer.

## Elk kwartaal in 5 stappen

**1. Bronnen inlezen** — Knop op Start

Klik op Start op de knop 'Bronnen inlezen'. Kies eerst de ZPMW03-export, daarna de Woningetalage-export. Vul de peildatum (bijv. 30-09-2026) en de versie (bijv. Q3) in.

*Let op:* Het bestand rekent alles opnieuw en opent het tabblad Controle. Eerdere vastgelegde prognoses blijven ongewijzigd.

**2. Controleren** — Tabblad Controle

Bekijk bovenaan het tabblad Controle het aantal FOUT en WAARSCHUWING. Los elke FOUT op. Meest voorkomend: een nieuw project → voeg het toe onderaan Invoer › 2. Projecten.

*Let op:* Onderaan Controle staan de lijsten met nieuwe, vervallen en verliesprojecten. Niets verdwijnt stil.

**3. Aanvullen** — Tabbladen Invoer en Stoplicht

Vul aan wat niet in SAP staat: taakstellingen, eliminaties, AK-dekking en fees (Invoer › 3), toelichtingen (Invoer › 4) en de gele cellen op het tabblad Stoplicht.

*Let op:* Wat je niet invult, staat in de rapportage als 'n.b.' — nooit stil als nul.

**4. Vastleggen** — Knop op Start of Historie

Klik op 'Prognose vastleggen' en bevestig. De jaarprognose van deze versie (bijv. Q3) wordt als vaste waarden toegevoegd aan het tabblad Historie.

*Let op:* Bestaat de versie al? Dan vraagt de knop om een toelichting en maakt een correctie (nieuwe revisie). Twee keer klikken geeft geen dubbeltelling.

**5. Rapporteren** — Dashboard en knop 'Maak PowerPoint'

Kies op het Dashboard het doeljaar en de versie. Klik op 'Maak PowerPoint' voor de presentatie.

*Let op:* De PowerPoint wordt opgeslagen naast dit Excelbestand en direct geopend.

## De tabbladen

| Tabblad | Wat staat er | Wat doe je er |
|---|---|---|
| Start | Knoppen, status en snelkoppelingen. | Hier begin je. |
| Dashboard | Kerncijfers, grafieken en de filters doeljaar, versie, opstal/deelneming en project. | Filters kiezen (gele cellen). |
| RDFO | De volledige RDFO: Budget, Q1–Q4 vastgelegd, upload, realisatie, to go, verschillen. | Alleen lezen. |
| Projecten | Per project: prognoses per versie, realisatie, to go, BOR, marge, orderintake, stoplicht. | Filteren met de pijltjes. |
| Projectdetail | Eén project in detail: gegevens, verloop, planning, woningverkopen. | Project kiezen bovenaan. |
| Woningen | Verkopen per maand en kwartaal, jaarverwachting per versie, koperstype, segment, projectmatrix. | Alleen lezen. |
| Orderintake | Orderintake per kwartaal en per project (PEW-omzet en PEW-BOR). | Alleen lezen. |
| Stoplicht | Stoplicht per project: verkoop, omgevingsvergunning, RvS, netcongestie, totaal. | Gele cellen invullen; uitgangspunten rechts. |
| Controle | Alle controles met status en uitzonderingslijsten. | Na elke inleesronde bekijken. |
| Invoer | Instellingen, projectenlijst, aanvullende posten, toelichtingen, koperstype. | Gele cellen invullen. |
| Historie | Alle vastgelegde prognoses en de importlog. | Nooit wijzigen: alleen de knop voegt toe. |
| Handleiding | Deze uitleg en de bijlagen. |  |

## Stoplicht invullen

- Per project vul je in: Omgevingsvergunning (Ja / In procedure / Nee), RvS (Geen / Laag / Midden / Hoog) en Netcongestie (Geen / Laag / Midden / Hoog).
- Verkoop kleurt automatisch op basis van het verkochte percentage uit de Woningetalage. Wijkt de beoordeling af? Vul dan 'Verkocht % handmatig' in.
- De uitgangspunten staan rechts op het tabblad: vanaf welk percentage verkoop groen of oranje is, en welke waarde welke kleur krijgt. Pas daar aan; alle lichten kleuren direct mee.
- Totaal = de slechtste kleur van de ingevulde lichten. Een leeg rondje betekent: niet ingevuld of niet van toepassing.
- Het stoplicht komt terug in Projecten, op het Dashboard en in de PowerPoint.

## Een vastgelegde prognose corrigeren

1. Pas de bron of de aanvullende invoer aan zodat de cijfers kloppen.
2. Klik op 'Prognose vastleggen'. De knop ziet dat de versie al bestaat en vraagt om een toelichting.
3. De correctie krijgt een hogere revisie. Alleen de nieuwste revisie telt; de oude regel blijft bewaard. In de RDFO staat ● achter de versie.

## Spelregels

- Wijzig nooit iets op het tabblad Historie. Alleen de knop 'Prognose vastleggen' voegt regels toe.
- Voeg nieuwe projecten altijd onderaan toe in Invoer › 2. Niet sorteren of tussenvoegen: het stoplicht hoort bij de regel.
- Bedragen in de bron zijn hele euro's; de rapportage toont EUR x 1.000 of EUR mln (instelbaar).
- Omzet en resultaat zijn positief, verlies is negatief (rood). Er wordt nooit een absolute waarde genomen.
- Woningaantallen en bedragen zijn al op aandeelbasis; het aandeel wordt niet nog eens toegepast.

## Problemen oplossen

| Probleem | Oplossing |
|---|---|
| De knoppen doen niets / gele balk 'Macro's zijn uitgeschakeld' | Klik op 'Inhoud inschakelen'. Komt het bestand uit e-mail of internet: sluit het, klik met rechts op het bestand › Eigenschappen › vink 'Blokkering opheffen' aan. |
| Mijn organisatie staat geen macro's toe | Alles rekent ook zonder macro's. Rechtermuisknop op een tabbladnaam › Zichtbaar maken › Bron_ZPMW03 / Bron_Woningetalage: plak de export in A1. Vastleggen: kopieer op het verborgen tabblad Concept A8:R(laatste regel) en plak als waarden op de eerste vrije rij van Historie. |
| 'Maak PowerPoint' geeft een foutmelding | PowerPoint moet op deze computer geïnstalleerd zijn. Sluit open dialoogvensters in PowerPoint en probeer opnieuw. |
| Er staat 'n.b.' in de RDFO | Er is (nog) geen vastgelegde waarde of aanvullende invoer. Dat is bewust geen nul. |
| Een oranje subtotaal in de RDFO | Minstens één onderliggende post is n.b.; het subtotaal is dus onvolledig. |
| Nieuw rapportagejaar | Zet in Invoer › 1 het rapportagejaar, het versiejaar van de upload en het versiejaar in de rapporten op het nieuwe jaar. Kolomkoppen schuiven mee. |
| Ik wil de rekenbladen zien | Klik op Start op 'Toon / verberg rekenbladen'. |
| Excel meldt bij openen dat inhoud is hersteld en de knoppen werken niet | Importeer de macro's handmatig: Alt+F11 › Bestand › Bestand importeren › macros/modRDFO.bas. Klik daarna met rechts op elke knop › Macro toewijzen. |

## Waarom zo gebouwd

- **Macro's** — Drie knoppen nemen het handwerk over: bronnen inlezen (bestandskiezer), prognose vastleggen en PowerPoint maken. Alle berekeningen blijven gewone Excel-formules, zodat het bestand ook zonder macro's rekent en controleerbaar blijft.
- **Waarom geen Power Query** — Met de knop 'Bronnen inlezen' is het inlezen al één klik. Power Query zou een tweede mechanisme toevoegen zonder extra voordeel; de staginglaag zoekt kolommen al op naam.
- **Grootboek** — Elke vastgelegde waarde is één regel met versie, revisie, peildatum, doeljaar, kengetal en waarde. Nieuwe versies vragen geen nieuwe kolommen.
