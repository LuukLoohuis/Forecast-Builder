"""Teksten voor de documentatiebladen (en de markdown in docs/). Eén bron, twee uitvoervormen."""

ONTWERPKEUZE = [
    ("Keuze", "Excel met gestructureerde lagen en formule-staging op kolomnaam; géén Power Query, géén macro's."),
    ("Waarom geen Power Query (nu)",
     "De bronnen worden per ronde handmatig vervangen en zijn klein (tientallen tot honderden regels). Een plak-tabblad "
     "met staging op kolomnaam geeft hetzelfde voordeel als Power Query (kolomvolgorde mag wijzigen, aliassen, dubbele "
     "headers op positie) maar werkt in elke Excelversie, rekent direct mee, is cel-voor-cel controleerbaar en kon in dit "
     "prototype automatisch getest worden. Power Query-queries kunnen in deze bouwomgeving niet worden uitgevoerd of "
     "getest; een ongeteste query opleveren past niet bij 'beloof niets dat niet werkt'."),
    ("Wanneer wél Power Query",
     "Als bronnen uit een vaste map/SharePoint komen of groter worden dan ~500 regels. De staginglaag (STG_*) is dan 1-op-1 "
     "te vervangen door een query die dezelfde interne veldnamen (Veldwoordenboek) oplevert; rapportage en versiebeheer "
     "blijven ongewijzigd."),
    ("Waarom geen macro's",
     "Vastleggen is bewust een handmatige plak-als-waarden-stap met automatische revisie, dubbel-detectie en controles. "
     "Er is dus geen upload- of snapshotknop; dat wordt nergens gesuggereerd."),
    ("Waarom een grootboek (long format)",
     "Elke vastgelegde waarde is één regel: Import-ID, bronbestand, importdatum, status, versiejaar, versie, revisie, "
     "peildatum, doeljaar, WBS, historische projectnaam/type/IFRS, kengetal, waarde, eenheid, toelichting. Nieuwe "
     "kengetallen of versies vragen geen nieuwe kolommen of bladen; rapportage leest alleen actieve vastgestelde regels."),
]

HANDLEIDING = [
    ("A. Nieuwe bronupload (elke rapportageronde)", [
        "1. Exporteer ZPMW03 en Woningetalage zoals gebruikelijk. Wijzig niets in de export.",
        "2. Tab SRC_ZPMW03: selecteer alles (Ctrl+A) → Delete. Plak de volledige export inclusief kopregel in cel A1.",
        "3. Tab SRC_Woningetalage: idem. Dubbele headers (Proj.srt, Profitctr.) en jaarheaders (2028, -2033 …) gewoon laten staan.",
        "4. Parameters: vul P_RJ (rapportagejaar), P_Peildatum, P_UploadVJ, P_UploadVersie, P_ImportID (nieuw, uniek), "
        "bestandsnamen, importdatum en het verwachte aantal regels per bron.",
        "5. Excel rekent automatisch. Open Controle: los alle FOUT-regels op (bijv. nieuw project → regel toevoegen in Mapping). "
        "Lees de WAARSCHUWINGEN en de uitzonderingslijsten onderaan.",
        "6. Vastgelegde versies (Budget/Q1/Q2/…) veranderen hierdoor NIET; alleen de kolom 'Upload … concept' en de uploadgedreven "
        "bladen (planning, woningmatrix, meerjarendoorkijk) wijzigen.",
    ]),
    ("B. Kwartaalprognose vastleggen (HANDMATIG — geen knop)", [
        "1. Voorwaarde: Controle toont 0 fouten en de upload is inhoudelijk akkoord (concept = de te vast te stellen jaarprognose).",
        "2. Parameters › P_StatusVastleg = 'Vastgesteld' (of 'Concept' voor een voorlopige vastlegging die nog niet meetelt).",
        "3. Tab Vastleggen: lees regel 4 (al vastgesteld? dan wordt het een correctie). Noteer 'Te plakken regels' en de rijen.",
        "4. Selecteer Vastleggen!A8:R<laatste rij> en kopieer (Ctrl+C).",
        "5. Tab Vastlegging: ga naar kolom A op de 'Eerste vrije rij' (cel D3) en kies Plakken speciaal › Waarden (Ctrl+Alt+V, W).",
        "6. Controleer: 'Actieve regels' neemt toe, 'Genegeerde dubbele regels' blijft 0, Controle blijft zonder fouten. In de RDFO "
        "is de versiekolom nu gevuld.",
        "7. Voeg een regel toe aan Importlog met hetzelfde Import-ID. Sla de werkmap op met de versie in de bestandsnaam.",
        "Nooit: regels in Vastlegging wijzigen of verwijderen, of formules in Vastlegging kolom A–R plakken.",
    ]),
    ("C. Correctie op een vastgestelde versie", [
        "1. Pas bron of aanvullende invoer aan zodat Vastleggen de juiste waarden toont.",
        "2. Vul Parameters › P_ToelCorrectie met de reden (verplicht). Vastleggen toont automatisch revisie = hoogste revisie + 1.",
        "3. Kopieer en plak zoals bij B. Alleen de hoogste vastgestelde revisie per sleutel telt; de oude regel blijft bewaard (audittrail).",
        "4. Wil je slechts één project corrigeren: plak alleen de betreffende regels. RDFO markeert versies met een correctie met ●.",
    ]),
    ("D. Dubbele import", [
        "Wordt hetzelfde blok twee keer geplakt (zelfde sleutel én revisie), dan telt de eerste regel en wordt de tweede "
        "als dubbel gemarkeerd (Vastlegging kolom V, Controle). Er ontstaat geen dubbeltelling. Verwijder de dubbele regels "
        "bij voorkeur niet maar documenteer in Importlog; verwijderen mag alleen als ze aantoonbaar identiek zijn.",
    ]),
    ("E. Aanvullende invoer en toelichtingen", [
        "Invoer_Aanvullend: taakstellingen, eliminaties, AK-dekking, fees, overige resultaten per doeljaar/versie/rubriek, "
        "regio- of projectniveau, in EUR en managementteken. Status Vastgesteld telt in versiekolommen; Concept alleen in de "
        "upload-kolom; Fictieve aanname alleen als P_Demo = Ja; Voorbeeld nooit. Ontbreekt een post, dan toont RDFO n.b.",
        "Invoer_Toelichting: één tekst per project/doeljaar/versie; zichtbaar in Projecten en Projectdetail.",
        "Invoer_Koperstype: optionele verdeling Belegger/Particulier voor toekomstige verkopen. Zonder invoer: Onbekend.",
    ]),
    ("F. Rapporteren en naar PowerPoint", [
        "Selecties (doeljaar, versie, type, project) staan bovenaan het Dashboard en sturen alle rapporten aan.",
        "RDFO, Dashboard, Projecten en Woningen hebben een afdrukbereik (A4 liggend). Kopieer een bereik als afbeelding "
        "(Start › Kopiëren › Als afbeelding) of plak grafieken als gekoppelde grafiek in PowerPoint.",
        "Eenheid wijzigen via Parameters › P_Eenheid (EUR, EUR x 1.000, EUR mln). Bronprecisie blijft in EUR.",
    ]),
    ("G. Nieuw jaar", [
        "Zet P_RJ, P_UploadVJ en P_VJ op het nieuwe jaar. Kolomkoppen ('OP.BOB <jaar>') en titels schuiven automatisch mee. "
        "Controleer in Veldwoordenboek of alle velden 'OK' zijn en of de jaarheader-mapping (tabel A) nog klopt.",
    ]),
]

# rapportageregel, bron, berekening, aanvullende invoer, openstaande definitie
MAPPING = [
    ("RDFO › Omzet opstal", "ZPMW03 + Mapping (Type=Opstal)",
     "FY RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ; FY RJ+1 = −OP.BOB RJ+1; per WBS gesommeerd (VERW_Project). Versiekolommen: Vastlegging 'Omzet FY'.",
     "—", "Is OP.BOB RJ in productie een restant of volledig jaar? (Controle waarschuwt als Q1–Q3 ≠ 0.)"),
    ("RDFO › Omzet deelnemingen", "ZPMW03 + Mapping (Type=Deelneming)", "Idem; bedragen al op aandeelbasis, aandeel niet toegepast.",
     "—", "Is Type ‘Deelneming’ in productie af te leiden uit een SAP-veld of blijft het mapping?"),
    ("RDFO › Taakstelling omzet", "—", "Som Invoer_Aanvullend (rubriek, soort FY/Actual YTD).", "Verplicht", "Bron en eigenaar van taakstellingen."),
    ("RDFO › Interne eliminaties", "—", "Som Invoer_Aanvullend.", "Verplicht", "Welke projecten/bedrijven worden geëlimineerd en op welk niveau?"),
    ("RDFO › Omzet old", "—", "Opstal + deelnemingen + taakstelling + interne eliminaties.", "—", "Bevestigen dat dit de oorspronkelijke opbouw is."),
    ("RDFO › IFRS-eliminaties", "—", "Som Invoer_Aanvullend. IFRS J/N alleen als memo, geen automatische eliminatie.", "Verplicht",
     "Regel voor IFRS-eliminatie (volledig/proportioneel/equity) per project."),
    ("RDFO › Omzet IFRS", "—", "Omzet old + IFRS-eliminaties.", "—", ""),
    ("RDFO › Projectresultaten / BOR", "ZPMW03", "FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ; RJ+1 = −OP.BOR RJ+1. Geen ABS: verlies blijft negatief.",
     "—", "Behandeling verliesvoorzieningen (direct volledig of via OP.BOR)?"),
    ("RDFO › AK-dekking", "—", "Som Invoer_Aanvullend (FY en Actual YTD apart).", "Verplicht",
     "Bestaande AK-methode (percentage per project? vast bedrag?) — niet generiek toegepast."),
    ("RDFO › Fees", "—", "Som Invoer_Aanvullend.", "Verplicht", "Bron van fees."),
    ("RDFO › Overige resultaten", "—", "Som Invoer_Aanvullend.", "Verplicht", ""),
    ("RDFO › Taakstelling EBIT", "—", "Som Invoer_Aanvullend.", "Verplicht", ""),
    ("RDFO › EBIT-bijdrage", "—", "BOR + AK-dekking + fees + overige resultaten + taakstelling EBIT.", "—", "Bevestigen opbouw."),
    ("RDFO › BOR %", "—", "Totale BOR / totale projectomzet (opstal + deelnemingen); deling door nul afgevangen.", "—", ""),
    ("RDFO › EBIT %", "—", "EBIT-bijdrage / Omzet old; deling door nul afgevangen.", "—", "Noemer bevestigen: Omzet old of Omzet IFRS?"),
    ("RDFO › Budget / Q1–Q4 FY", "Vastlegging", "Actieve vastgestelde regels (hoogste revisie, eerste voorkomen) voor versiejaar P_VJ.",
     "Budget/Q1/Q2: fictieve historie", "Bevatten oudere productiebladen volledige jaarbedragen of delta’s t.o.v. budget?"),
    ("RDFO › Upload concept", "SRC_ZPMW03 / SRC_Woningetalage", "Live uit de huidige bron; geen versie tot handmatig vastgelegd.", "—", ""),
    ("RDFO › Realisatie YTD", "Vastlegging ('Omzet actual YTD') of upload", "Altijd van dezelfde versie/peildatum als de geselecteerde FY-prognose.",
     "—", "Historische actuals Budget/Q1/Q2 niet beschikbaar → n.b."),
    ("RDFO › To go", "—", "Geselecteerde FY − realisatie YTD van dezelfde peildatum. Doeljaar > versiejaar: to go = FY.", "—", ""),
    ("RDFO › Δ vs budget / vorige", "—", "Geselecteerde versie − Budget; geselecteerde versie − direct voorgaande versie.", "—", ""),
    ("Dashboard › KPI-tegels", "CALC_Versies (gefilterd)", "Zelfde logica als RDFO, met filter type/project.", "—", ""),
    ("Dashboard › Jaarprognoseverloop", "CALC_Versies", "FY-omzet/BOR per vastgelegde versie; ontbrekende versie = geen staaf + label '(n.b.)'.", "—", ""),
    ("Dashboard › Realisatie vs to go", "Projecten", "Top 10 projecten op geselecteerde FY; gestapeld realisatie + to go.", "—", ""),
    ("Dashboard › Grootste mutaties", "Projecten", "Top 8 op |Δ FY-omzet| geselecteerde vs vorige versie.", "—", ""),
    ("Dashboard › Meerjarendoorkijk", "ZPMW03 (upload)", "−OP.BOB/−OP.BOR per jaar RJ..RJ+8 + onverdeeld; niet vastgelegd.", "—", ""),
    ("Projecten › alle kolommen", "Vastlegging + VERW_Project + Invoer_Toelichting", "Per WBS, zelfde definities als RDFO. Marge = BOR/omzet.", "Toelichting", ""),
    ("Projectdetail › planning", "ZPMW03", "Realisatie t/m peildatum, Q4 rest, RJ+1 per kwartaal, RJ+2..RJ+8, onverdeeld.", "—", ""),
    ("Woningen › verkopen per maand/kwartaal", "Woningetalage", "hjr 01..12 (≤ peilmaand = actual), Volgend jr Q1–Q4 = RJ+1.", "—",
     "Bevestigen dat hjr-maanden na peildatum altijd prognose zijn."),
    ("Woningen › jaarverwachting per versie", "Vastlegging ('Woningverkopen FY')", "Per versie; upload = 'Huidig jaar' / 'Volgend jaar'.", "—", ""),
    ("Woningen › totaal / cumulatief / etalage", "Woningetalage", "Aantal woningen, Cumulatief Totaal, Etalage Totaal; al op aandeelbasis, NIET opnieuw gewogen.",
     "—", "Is ‘Aantal woningen’ in productie bruto of aandeelbasis?"),
    ("Woningen › bruto (afgeleid)", "Woningetalage + Mapping", "Aandeelbasis ÷ aandeel% — alleen informatief.", "—", ""),
    ("Woningen › koperstype", "Woningetalage (actual) + Invoer_Koperstype (prognose)", "Actual Beleggers/Particulier; prognose zonder invoer = Onbekend.", "Optioneel", ""),
    ("Woningen › segment", "Mapping ('Segment demo')", "Fictieve demo-indeling; vervangen door echte segmentbron.", "Mapping", "Bron voor segment in productie."),
]

PRODUCTIEVRAGEN = [
    ("Definitie", "Bevat OP.BOB <jaar> in productie het restant na de peildatum of het volledige jaar? Idem OP.BOR.", "FY-formule; dubbeltelling of onderschatting."),
    ("Definitie", "Zijn ZPMW03-bedragen en Woningetalage-aantallen in productie al op aandeelbasis (zoals in de demo)?", "Woningen/omzet van deelnemingen."),
    ("Definitie", "Tekenconventie in productie: opbrengsten en winst negatief, kosten positief — geldt dat voor alle bedrijven?", "Tekenomzetting."),
    ("Definitie", "Welke rubrieken vormen exact Omzet old, Omzet IFRS en EBIT-bijdrage in de huidige RDFO? Wat is de noemer van EBIT%?", "RDFO-opbouw."),
    ("Definitie", "Hoe wordt AK-dekking nu bepaald (methode, percentages, realisatie)? Wie levert dit aan?", "Aanvullende invoer."),
    ("Definitie", "Welke regel bepaalt IFRS-eliminaties? Is IFRS J/N voldoende of is er een consolidatiemethode per project?", "Omzet IFRS."),
    ("Definitie", "Bevatten de oudere prognosebladen (Q1/Q2 en ouder) volledige jaarbedragen of verschillen t.o.v. budget?", "Historische migratie."),
    ("Definitie", "Betekenis van Woningetalage-kolommen 'Rond', 2029–2032 en negatieve jaarheaders (-2033 …).", "Jaarheader-mapping."),
    ("Data", "Kan één Projectdefinitie in productie meerdere WBS-elementen hebben met elk eigen woningen? Op welk niveau staat Stufe 3?", "Koppeling; nu 1 WBS = 1 regel."),
    ("Data", "Is 'Aantal woningen' bruto of aandeelbasis, en is 'Aandeel Heijm.' altijd gelijk aan het participatiepercentage in SAP?", "Weging."),
    ("Data", "Waar komen Type (opstal/deelneming) en segment/koperstype voor de prognose vandaan in productie?", "Mapping-onderhoud."),
    ("Data", "Hoe vaak verandert de kolomvolgorde/naamgeving van de exports (bijv. Bedrijfsnaam vs Naam maatschappij)?", "Aliassen."),
    ("Proces", "Wie stelt een versie vast (rol) en wanneer is een versie definitief? Is er een vier-ogenprincipe?", "Status Concept/Vastgesteld."),
    ("Proces", "Hoe worden historische versies uit de bestaande RDFO gemigreerd (welke bladen, welke jaren)?", "Eerste vulling grootboek."),
    ("Proces", "Moet het budget van het volgende jaar (bijv. Budget 2027) een eigen versiejaar krijgen? (Prototype: versiejaar = jaar van de cyclus.)", "Versiejaarconventie."),
    ("Techniek", "Mag de werkmap Power Query gebruiken (bronnen uit vaste map/SharePoint)? Dan vervangt een query de STG-tabbladen.", "Importautomatisering."),
    ("Techniek", "Verwacht volume per bron en aantal rondes: capaciteit is nu 500 bronregels, 100 projecten en 5.000 grootboekregels.", "Uitbreiding formulebereiken."),
]
