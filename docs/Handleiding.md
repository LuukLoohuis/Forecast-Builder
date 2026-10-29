# Handleiding RDFO-prototype

Bronvervanging, verversen, kwartaalvastlegging en correcties. Dezelfde tekst staat in de werkmap (tab *Handleiding*).

## Ontwerpkeuze

- **Keuze** — Excel met gestructureerde lagen en formule-staging op kolomnaam; géén Power Query, géén macro's.
- **Waarom geen Power Query (nu)** — De bronnen worden per ronde handmatig vervangen en zijn klein (tientallen tot honderden regels). Een plak-tabblad met staging op kolomnaam geeft hetzelfde voordeel als Power Query (kolomvolgorde mag wijzigen, aliassen, dubbele headers op positie) maar werkt in elke Excelversie, rekent direct mee, is cel-voor-cel controleerbaar en kon in dit prototype automatisch getest worden. Power Query-queries kunnen in deze bouwomgeving niet worden uitgevoerd of getest; een ongeteste query opleveren past niet bij 'beloof niets dat niet werkt'.
- **Wanneer wél Power Query** — Als bronnen uit een vaste map/SharePoint komen of groter worden dan ~500 regels. De staginglaag (STG_*) is dan 1-op-1 te vervangen door een query die dezelfde interne veldnamen (Veldwoordenboek) oplevert; rapportage en versiebeheer blijven ongewijzigd.
- **Waarom geen macro's** — Vastleggen is bewust een handmatige plak-als-waarden-stap met automatische revisie, dubbel-detectie en controles. Er is dus geen upload- of snapshotknop; dat wordt nergens gesuggereerd.
- **Waarom een grootboek (long format)** — Elke vastgelegde waarde is één regel: Import-ID, bronbestand, importdatum, status, versiejaar, versie, revisie, peildatum, doeljaar, WBS, historische projectnaam/type/IFRS, kengetal, waarde, eenheid, toelichting. Nieuwe kengetallen of versies vragen geen nieuwe kolommen of bladen; rapportage leest alleen actieve vastgestelde regels.

## A. Nieuwe bronupload (elke rapportageronde)

- 1. Exporteer ZPMW03 en Woningetalage zoals gebruikelijk. Wijzig niets in de export.
- 2. Tab SRC_ZPMW03: selecteer alles (Ctrl+A) → Delete. Plak de volledige export inclusief kopregel in cel A1.
- 3. Tab SRC_Woningetalage: idem. Dubbele headers (Proj.srt, Profitctr.) en jaarheaders (2028, -2033 …) gewoon laten staan.
- 4. Parameters: vul P_RJ (rapportagejaar), P_Peildatum, P_UploadVJ, P_UploadVersie, P_ImportID (nieuw, uniek), bestandsnamen, importdatum en het verwachte aantal regels per bron.
- 5. Excel rekent automatisch. Open Controle: los alle FOUT-regels op (bijv. nieuw project → regel toevoegen in Mapping). Lees de WAARSCHUWINGEN en de uitzonderingslijsten onderaan.
- 6. Vastgelegde versies (Budget/Q1/Q2/…) veranderen hierdoor NIET; alleen de kolom 'Upload … concept' en de uploadgedreven bladen (planning, woningmatrix, meerjarendoorkijk) wijzigen.

## B. Kwartaalprognose vastleggen (HANDMATIG — geen knop)

- 1. Voorwaarde: Controle toont 0 fouten en de upload is inhoudelijk akkoord (concept = de te vast te stellen jaarprognose).
- 2. Parameters › P_StatusVastleg = 'Vastgesteld' (of 'Concept' voor een voorlopige vastlegging die nog niet meetelt).
- 3. Tab Vastleggen: lees regel 4 (al vastgesteld? dan wordt het een correctie). Noteer 'Te plakken regels' en de rijen.
- 4. Selecteer Vastleggen!A8:R<laatste rij> en kopieer (Ctrl+C).
- 5. Tab Vastlegging: ga naar kolom A op de 'Eerste vrije rij' (cel D3) en kies Plakken speciaal › Waarden (Ctrl+Alt+V, W).
- 6. Controleer: 'Actieve regels' neemt toe, 'Genegeerde dubbele regels' blijft 0, Controle blijft zonder fouten. In de RDFO is de versiekolom nu gevuld.
- 7. Voeg een regel toe aan Importlog met hetzelfde Import-ID. Sla de werkmap op met de versie in de bestandsnaam.
- Nooit: regels in Vastlegging wijzigen of verwijderen, of formules in Vastlegging kolom A–R plakken.

## C. Correctie op een vastgestelde versie

- 1. Pas bron of aanvullende invoer aan zodat Vastleggen de juiste waarden toont.
- 2. Vul Parameters › P_ToelCorrectie met de reden (verplicht). Vastleggen toont automatisch revisie = hoogste revisie + 1.
- 3. Kopieer en plak zoals bij B. Alleen de hoogste vastgestelde revisie per sleutel telt; de oude regel blijft bewaard (audittrail).
- 4. Wil je slechts één project corrigeren: plak alleen de betreffende regels. RDFO markeert versies met een correctie met ●.

## D. Dubbele import

- Wordt hetzelfde blok twee keer geplakt (zelfde sleutel én revisie), dan telt de eerste regel en wordt de tweede als dubbel gemarkeerd (Vastlegging kolom V, Controle). Er ontstaat geen dubbeltelling. Verwijder de dubbele regels bij voorkeur niet maar documenteer in Importlog; verwijderen mag alleen als ze aantoonbaar identiek zijn.

## E. Aanvullende invoer en toelichtingen

- Invoer_Aanvullend: taakstellingen, eliminaties, AK-dekking, fees, overige resultaten per doeljaar/versie/rubriek, regio- of projectniveau, in EUR en managementteken. Status Vastgesteld telt in versiekolommen; Concept alleen in de upload-kolom; Fictieve aanname alleen als P_Demo = Ja; Voorbeeld nooit. Ontbreekt een post, dan toont RDFO n.b.
- Invoer_Toelichting: één tekst per project/doeljaar/versie; zichtbaar in Projecten en Projectdetail.
- Invoer_Koperstype: optionele verdeling Belegger/Particulier voor toekomstige verkopen. Zonder invoer: Onbekend.

## F. Rapporteren en naar PowerPoint

- Selecties (doeljaar, versie, type, project) staan bovenaan het Dashboard en sturen alle rapporten aan.
- RDFO, Dashboard, Projecten en Woningen hebben een afdrukbereik (A4 liggend). Kopieer een bereik als afbeelding (Start › Kopiëren › Als afbeelding) of plak grafieken als gekoppelde grafiek in PowerPoint.
- Eenheid wijzigen via Parameters › P_Eenheid (EUR, EUR x 1.000, EUR mln). Bronprecisie blijft in EUR.

## G. Nieuw jaar

- Zet P_RJ, P_UploadVJ en P_VJ op het nieuwe jaar. Kolomkoppen ('OP.BOB <jaar>') en titels schuiven automatisch mee. Controleer in Veldwoordenboek of alle velden 'OK' zijn en of de jaarheader-mapping (tabel A) nog klopt.
