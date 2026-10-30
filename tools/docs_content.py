"""Teksten voor de handleiding (werkmap én docs/). Eén bron, twee uitvoervormen."""

INLEIDING = (
    "Dit bestand maakt de RDFO-rapportage, het dashboard, de projectoverzichten, het stoplicht en de PowerPoint "
    "uit twee SAP-exports: ZPMW03 en Woningetalage. Eén keer per kwartaal lees je de exports in, controleer je, "
    "vul je aan en leg je de prognose vast. Vastgelegde prognoses veranderen daarna nooit meer."
)

# (titel, wat je doet, waar / welke knop, waar je op let)
STAPPEN = [
    ("Bronnen inlezen",
     "Klik op Start op de knop 'Bronnen inlezen'. Kies eerst de ZPMW03-export, daarna de Woningetalage-export. "
     "Vul de peildatum (bijv. 30-09-2026) en de versie (bijv. Q3) in.",
     "Knop op Start",
     "Het bestand rekent alles opnieuw en opent het tabblad Controle. Eerdere vastgelegde prognoses blijven ongewijzigd."),
    ("Controleren",
     "Bekijk bovenaan het tabblad Controle het aantal FOUT en WAARSCHUWING. Los elke FOUT op. "
     "Meest voorkomend: een nieuw project → voeg het toe onderaan Invoer › 2. Projecten.",
     "Tabblad Controle",
     "Onderaan Controle staan de lijsten met nieuwe, vervallen en verliesprojecten. Niets verdwijnt stil."),
    ("Aanvullen",
     "Vul aan wat niet in SAP staat: taakstellingen, eliminaties, AK-dekking en fees (Invoer › 3), toelichtingen "
     "(Invoer › 4) en de gele cellen op het tabblad Stoplicht.",
     "Tabbladen Invoer en Stoplicht",
     "Wat je niet invult, staat in de rapportage als 'n.b.' — nooit stil als nul."),
    ("Vastleggen",
     "Klik op 'Prognose vastleggen' en bevestig. De jaarprognose van deze versie (bijv. Q3) wordt als vaste waarden "
     "toegevoegd aan het tabblad Historie.",
     "Knop op Start of Historie",
     "Bestaat de versie al? Dan vraagt de knop om een toelichting en maakt een correctie (nieuwe revisie). "
     "Twee keer klikken geeft geen dubbeltelling."),
    ("Rapporteren",
     "Kies op het Dashboard het doeljaar en de versie. Klik op 'Maak PowerPoint' voor de presentatie.",
     "Dashboard en knop 'Maak PowerPoint'",
     "De PowerPoint wordt opgeslagen naast dit Excelbestand en direct geopend."),
]

TABBLADEN = [
    ("Start", "Knoppen, status en snelkoppelingen.", "Hier begin je."),
    ("Dashboard", "Kerncijfers, grafieken en de filters doeljaar, versie, opstal/deelneming en project.", "Filters kiezen (gele cellen)."),
    ("RDFO", "De volledige RDFO: Budget, Q1–Q4 vastgelegd, upload, realisatie, to go, verschillen.", "Alleen lezen."),
    ("Projecten", "Per project: prognoses per versie, realisatie, to go, BOR, marge, orderintake, stoplicht.", "Filteren met de pijltjes."),
    ("Projectdetail", "Eén project in detail: gegevens, verloop, planning, woningverkopen.", "Project kiezen bovenaan."),
    ("Woningen", "Verkopen per maand en kwartaal, jaarverwachting per versie, koperstype, segment, projectmatrix.", "Alleen lezen."),
    ("Orderintake", "Orderintake per kwartaal en per project (PEW-omzet en PEW-BOR).", "Alleen lezen."),
    ("Stoplicht", "Stoplicht per project: verkoop, omgevingsvergunning, RvS, netcongestie, totaal.", "Gele cellen invullen; uitgangspunten rechts."),
    ("Controle", "Alle controles met status en uitzonderingslijsten.", "Na elke inleesronde bekijken."),
    ("Invoer", "Instellingen, projectenlijst, aanvullende posten, toelichtingen, koperstype.", "Gele cellen invullen."),
    ("Historie", "Alle vastgelegde prognoses en de importlog.", "Nooit wijzigen: alleen de knop voegt toe."),
    ("Handleiding", "Deze uitleg en de bijlagen.", ""),
]

STOPLICHT = [
    "Per project vul je in: Omgevingsvergunning (Ja / In procedure / Nee), RvS (Geen / Laag / Midden / Hoog) en Netcongestie (Geen / Laag / Midden / Hoog).",
    "Verkoop kleurt automatisch op basis van het verkochte percentage uit de Woningetalage. Wijkt de beoordeling af? Vul dan 'Verkocht % handmatig' in.",
    "De uitgangspunten staan rechts op het tabblad: vanaf welk percentage verkoop groen of oranje is, en welke waarde welke kleur krijgt. Pas daar aan; alle lichten kleuren direct mee.",
    "Totaal = de slechtste kleur van de ingevulde lichten. Een leeg rondje betekent: niet ingevuld of niet van toepassing.",
    "Het stoplicht komt terug in Projecten, op het Dashboard en in de PowerPoint.",
]

CORRECTIE = [
    "Pas de bron of de aanvullende invoer aan zodat de cijfers kloppen.",
    "Klik op 'Prognose vastleggen'. De knop ziet dat de versie al bestaat en vraagt om een toelichting.",
    "De correctie krijgt een hogere revisie. Alleen de nieuwste revisie telt; de oude regel blijft bewaard. In de RDFO staat ● achter de versie.",
]

FAQ = [
    ("De knoppen doen niets / gele balk 'Macro's zijn uitgeschakeld'",
     "Klik op 'Inhoud inschakelen'. Komt het bestand uit e-mail of internet: sluit het, klik met rechts op het bestand › Eigenschappen › vink 'Blokkering opheffen' aan."),
    ("Mijn organisatie staat geen macro's toe",
     "Alles rekent ook zonder macro's. Rechtermuisknop op een tabbladnaam › Zichtbaar maken › Bron_ZPMW03 / Bron_Woningetalage: plak de export in A1. "
     "Vastleggen: kopieer op het verborgen tabblad Concept A8:R(laatste regel) en plak als waarden op de eerste vrije rij van Historie."),
    ("'Maak PowerPoint' geeft een foutmelding", "PowerPoint moet op deze computer geïnstalleerd zijn. Sluit open dialoogvensters in PowerPoint en probeer opnieuw."),
    ("Er staat 'n.b.' in de RDFO", "Er is (nog) geen vastgelegde waarde of aanvullende invoer. Dat is bewust geen nul."),
    ("Een oranje subtotaal in de RDFO", "Minstens één onderliggende post is n.b.; het subtotaal is dus onvolledig."),
    ("Nieuw rapportagejaar", "Zet in Invoer › 1 het rapportagejaar, het versiejaar van de upload en het versiejaar in de rapporten op het nieuwe jaar. Kolomkoppen schuiven mee."),
    ("Ik wil de rekenbladen zien", "Klik op Start op 'Toon / verberg rekenbladen'."),
    ("Excel meldt bij openen dat inhoud is hersteld en de knoppen werken niet",
     "Importeer de macro's handmatig: Alt+F11 › Bestand › Bestand importeren › macros/modRDFO.bas. Klik daarna met rechts op elke knop › Macro toewijzen."),
]

SPELREGELS = [
    "Wijzig nooit iets op het tabblad Historie. Alleen de knop 'Prognose vastleggen' voegt regels toe.",
    "Voeg nieuwe projecten altijd onderaan toe in Invoer › 2. Niet sorteren of tussenvoegen: het stoplicht hoort bij de regel.",
    "Bedragen in de bron zijn hele euro's; de rapportage toont EUR x 1.000 of EUR mln (instelbaar).",
    "Omzet en resultaat zijn positief, verlies is negatief (rood). Er wordt nooit een absolute waarde genomen.",
    "Woningaantallen en bedragen zijn al op aandeelbasis; het aandeel wordt niet nog eens toegepast.",
]

ONTWERPKEUZE = [
    ("Macro's", "Drie knoppen nemen het handwerk over: bronnen inlezen (bestandskiezer), prognose vastleggen en PowerPoint maken. "
                "Alle berekeningen blijven gewone Excel-formules, zodat het bestand ook zonder macro's rekent en controleerbaar blijft."),
    ("Waarom geen Power Query", "Met de knop 'Bronnen inlezen' is het inlezen al één klik. Power Query zou een tweede mechanisme toevoegen "
                                "zonder extra voordeel; de staginglaag zoekt kolommen al op naam."),
    ("Grootboek", "Elke vastgelegde waarde is één regel met versie, revisie, peildatum, doeljaar, kengetal en waarde. Nieuwe versies vragen geen nieuwe kolommen."),
]

# rapportageregel, bron, berekening, aanvullende invoer, openstaande definitie
MAPPING = [
    ("RDFO › Omzet opstal", "ZPMW03 + Invoer › 2 (Type=Opstal)",
     "FY RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ; FY RJ+1 = −OP.BOB RJ+1; per WBS gesommeerd. Versiekolommen: Historie 'Omzet FY'.",
     "—", "Is OP.BOB RJ in productie een restant of een volledig jaar? (Controle waarschuwt als Q1–Q3 ≠ 0.)"),
    ("RDFO › Omzet deelnemingen", "ZPMW03 + Invoer › 2 (Type=Deelneming)", "Idem; al op aandeelbasis.", "—", "Type in productie uit SAP-veld of mapping?"),
    ("RDFO › Taakstelling omzet", "—", "Som Invoer › 3.", "Verplicht", "Bron en eigenaar van taakstellingen."),
    ("RDFO › Interne eliminaties", "—", "Som Invoer › 3.", "Verplicht", "Welke projecten/bedrijven en op welk niveau?"),
    ("RDFO › Omzet old", "—", "Opstal + deelnemingen + taakstelling + interne eliminaties.", "—", "Bevestigen dat dit de bestaande opbouw is."),
    ("RDFO › IFRS-eliminaties", "—", "Som Invoer › 3. IFRS J/N alleen als memo.", "Verplicht", "Regel voor IFRS-eliminatie per project."),
    ("RDFO › Omzet IFRS", "—", "Omzet old + IFRS-eliminaties.", "—", ""),
    ("RDFO › Projectresultaten / BOR", "ZPMW03", "FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ; RJ+1 = −OP.BOR RJ+1. Geen ABS.", "—", "Behandeling verliesvoorzieningen."),
    ("RDFO › AK-dekking, fees, overige, taakstelling EBIT", "—", "Som Invoer › 3 (FY en Actual YTD apart).", "Verplicht", "Bestaande AK-methode — niet generiek toegepast."),
    ("RDFO › EBIT-bijdrage", "—", "BOR + AK-dekking + fees + overige + taakstelling EBIT.", "—", "Bevestigen opbouw."),
    ("RDFO › BOR % / EBIT %", "—", "Totale BOR / projectomzet; EBIT / Omzet old; deling door nul afgevangen.", "—", "Noemer EBIT% bevestigen."),
    ("RDFO › Budget / Q1–Q4", "Historie", "Hoogste vastgestelde revisie per project/kengetal.", "Budget/Q1/Q2: fictieve historie", "Volledige jaarbedragen of delta’s in oude bladen?"),
    ("RDFO › Upload concept", "Bron_ZPMW03 / Bron_Woningetalage", "Live uit de bron; pas een versie na 'Prognose vastleggen'.", "—", ""),
    ("RDFO › Realisatie YTD / to go", "Historie of upload", "Altijd van dezelfde versie/peildatum als de FY-prognose.", "—", ""),
    ("Orderintake", "ZPMW03 (kolom Orderintake, PEW opbrengst, PEW resultaat)",
     "Jaar/kwartaal = datum in 'Orderintake'; bedrag = −PEW opbrengst; BOR = −PEW resultaat.", "Budget: Invoer › 3 'Orderintake budget'",
     "Is 'Orderintake' de juiste datum en PEW-omzet het juiste bedrag? Orderintake wordt (nog) niet per versie vastgelegd."),
    ("Stoplicht › Verkoop", "Woningetalage", "Cumulatief verkocht / totaal woningen (of handmatig %); drempels instelbaar.", "Stoplicht", "Welke drempels gelden? Verkocht % t.o.v. start bouw?"),
    ("Stoplicht › Omgevingsvergunning, RvS, netcongestie", "—", "Handmatige invoer; kleur via tabel uitgangspunten.", "Stoplicht", ""),
    ("Woningen", "Woningetalage", "hjr 01..12 (≤ peilmaand = realisatie), Volgend jr Q1–Q4, jaarverwachting per versie uit Historie.", "Koperstype: Invoer › 5", "Is 'Aantal woningen' bruto of aandeelbasis?"),
]

PRODUCTIEVRAGEN = [
    ("Definitie", "Bevat OP.BOB <jaar> het restant na de peildatum of het volledige jaar? Idem OP.BOR.", "FY-formule."),
    ("Definitie", "Zijn ZPMW03-bedragen en Woningetalage-aantallen al op aandeelbasis (zoals in de demo)?", "Weging."),
    ("Definitie", "Tekenconventie: opbrengsten en winst negatief, kosten positief — voor alle bedrijven?", "Tekenomzetting."),
    ("Definitie", "Exacte opbouw van Omzet old, Omzet IFRS en EBIT-bijdrage; noemer van EBIT%?", "RDFO-opbouw."),
    ("Definitie", "Hoe wordt AK-dekking bepaald en wie levert die aan?", "Aanvullende invoer."),
    ("Definitie", "Regel voor IFRS-eliminaties per project?", "Omzet IFRS."),
    ("Definitie", "Is de kolom 'Orderintake' de orderintakedatum en is PEW-omzet het juiste orderintakebedrag?", "Orderintake."),
    ("Definitie", "Welke drempels gelden voor het verkoopstoplicht, en is 'verkocht %' t.o.v. totaal of t.o.v. start bouw?", "Stoplicht."),
    ("Definitie", "Betekenis van Woningetalage-kolommen 'Rond', 2029–2032 en negatieve jaarkoppen.", "Jaarkolommen."),
    ("Data", "Kan één projectdefinitie meerdere WBS-elementen met eigen woningen hebben?", "Koppeling."),
    ("Data", "Waar komen type (opstal/deelneming), segment en koperstype in productie vandaan?", "Mapping-onderhoud."),
    ("Proces", "Wie stelt een versie vast en is er een vier-ogenprincipe?", "Vastleggen."),
    ("Proces", "Hoe migreren we historische versies uit de huidige RDFO?", "Eerste vulling Historie."),
    ("Techniek", "Zijn macro's toegestaan op de werkplekken (vertrouwde locatie / ondertekening)?", "Knoppen."),
    ("Techniek", "Capaciteit: 500 bronregels, 100 projecten, 5.000 historieregels — voldoende?", "Formulebereiken."),
]
