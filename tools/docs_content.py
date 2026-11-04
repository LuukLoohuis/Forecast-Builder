"""Teksten voor de handleiding (werkmap én docs/). Eén bron, twee uitvoervormen."""

INLEIDING = (
    "Met dit bestand maak je elk kwartaal de RDFO, het dashboard, het stoplicht, de bandbreedte en de PowerPoint. "
    "Je hebt twee exports uit SAP nodig: ZPMW03 en Woningetalage. Je leest ze in met één knop, kijkt of er fouten zijn, "
    "vult aan wat niet in SAP staat en zet de prognose vast. Een vastgezette prognose verandert daarna niet meer."
)

# (titel, wat je doet, waar / welke knop, waar je op let)
STAPPEN = [
    ("Inlezen",
     "Klik op 'Bronnen inlezen'. Kies de ZPMW03-export, daarna de Woningetalage-export daarna (als je die hebt) de AK-percentages uit SAP en de AK-realisatie (grootboekdump). "
     "Vul de peildatum in (bijv. 30-09-2026) en de versie (bijv. Q3).",
     "Knop op Start",
     "Het bestand opent daarna het tabblad Controle. Eerder vastgezette prognoses blijven zoals ze waren."),
    ("Controleren",
     "Kijk bovenaan het tabblad Controle: staat er een FOUT? Los die eerst op. "
     "Staan er nieuwe projecten? Klik op 'Nieuwe projecten toevoegen' en kies in Invoer › 2 of het Opstal of Deelneming is.",
     "Tabblad Controle",
     "Achter elke controle staat wat je moet doen. Onderaan staan de lijsten met nieuwe, verdwenen en verliesgevende projecten."),
    ("Aanvullen",
     "Vul de gele cellen in: op het tabblad Bandbreedte het stoplicht (vergunning, netcongestie) en eventueel kans %; "
     "op Invoer wat niet in SAP staat (taakstellingen, toelichtingen, kansen en risico's, DAEB, reserves).",
     "Tabbladen Bandbreedte en Invoer",
     "Niet ingevuld? Dan staat er 'n.b.' (niet bekend). Het bestand vult nooit stiekem een nul in."),
    ("Vastzetten",
     "Klik op 'Prognose vastleggen'. Je ziet de omzet en het aantal woningen. Klik op OK. Daarna kun je meteen opslaan.",
     "Knop op Start",
     "Is deze versie al vastgezet? Dan vraagt het bestand waarom je hem wilt aanpassen. De oude cijfers blijven bewaard."),
    ("Presenteren",
     "Kies op het Dashboard het jaar en de versie. Klik op 'Maak PowerPoint'.",
     "Dashboard of Start",
     "De PowerPoint komt in dezelfde map als dit bestand. Alle tabellen zijn gewone PowerPoint-tabellen: je kunt ze nog aanpassen."),
]

TABBLADEN = [
    ("Start", "Knoppen en status.", "Hier begin je."),
    ("Dashboard", "Kerncijfers en grafieken. Kies hier jaar, versie, opstal/deelneming of één project.", "Gele cellen kiezen."),
    ("RDFO", "De RDFO: budget, Q1–Q4, upload, realisatie, to go en verschillen. Onderaan: per gebied.", "Alleen lezen."),
    ("Outlook", "De 4 jaren na het lopende jaar: omzet (ook DAEB, funnel, target), BOR, AK-dekking, kansen en risico's, reserves, bandbreedte, woningen, orderintake.", "Alleen lezen. Basis voor het budget."),
    ("Projecten", "Per project: alle versies, realisatie, to go, BOR, marge en stoplicht.", "Filteren met de pijltjes."),
    ("Projectdetail", "Eén project in detail.", "Project kiezen bovenaan."),
    ("Woningen", "Verkopen per maand en kwartaal, per project per kwartaal (dit en volgend jaar), per versie, per productsegment en koperstype.", "Alleen lezen."),
    ("Bandbreedte", "Bovenaan de vijf scenario's per jaar met grafiek. Daaronder per project het stoplicht en de scenario's.", "Gele cellen invullen."),
    ("Kansen en risico's", "Invoer per kans of risico (project, jaar, kwartaal, bedrag) en een grafiek per project; staat ook in de PowerPoint.", "Gele cellen invullen."),
    ("Projectkeuze", "Per project: telt het mee voor prognose, woningaantallen, AK-dekking en de PowerPoint? Plus eigen AK % en de reserves.", "Ja/Nee kiezen; leeg = Ja."),
    ("AK-dekking", "Per project de hele AK-berekening: AK %, kosten per kwartaal, geboekte AK, prognose rest jaar, FY, budget en volgende jaren.", "Alleen lezen; om na te lopen."),
    ("Controle", "Alle controles. Rood = eerst oplossen.", "Na elke keer inlezen bekijken."),
    ("Invoer", "Instellingen, projectlijst en alles wat niet in SAP staat.", "Gele cellen invullen."),
    ("Handleiding", "Deze uitleg.", ""),
]

STOPLICHT = [
    "Per project vul je in: omgevingsvergunning (Ja / In procedure / Nee) en netcongestie (Geen / Laag / Midden / Hoog).",
    "Verkoop kleurt vanzelf op het percentage verkochte woningen. Wil je het percentage bij start bouw gebruiken? Vul dan 'Verkocht % SB' in.",
    "Rechts op het tabblad staan de uitgangspunten: vanaf welk percentage het groen of oranje is, en welke keuze welke kleur krijgt. Pas je dat aan, dan kleuren alle lichten meteen mee.",
    "Een leeg rondje betekent: nog niet ingevuld of niet van toepassing.",
]

BANDBREEDTE = [
    "Het bestand rekent per project vijf scenario's uit, voor het lopende jaar en de twee jaren daarna. Bovenaan het tabblad Bandbreedte zie je het totaal per jaar; de grafiek toont per jaar een band van het laagste tot het hoogste scenario.",
    "Optimistisch: de verkoop haalt aan het eind van het jaar minstens het ingestelde percentage (standaard 90%). De omzet groeit dan mee.",
    "Basis: de omzet uit de prognose (upload).",
    "Afslag: maar een deel van de geplande verkopen van dat jaar lukt (standaard 50%). De omzet daalt dan mee.",
    "Gewogen: omzet x kans %. De kans komt uit ZPMW03; je kunt hem per project zelf invullen (kolom 'Kans %').",
    "0-100: een project telt helemaal mee of helemaal niet. Standaard telt het mee vanaf 50% kans; per project kun je Ja of Nee kiezen.",
    "Wat al verkocht en gerealiseerd is, verandert in geen enkel scenario. De instellingen staan op het tabblad onder de tabel.",
]

INVOER_EENMALIG = [
    ("2. Projecten", "Eén regel per project. Kies Type, Gebied, Strapo (Ja/Nee) en hoe de woningen verdeeld zijn over de productsegmenten en kopers. AK: het percentage komt uit de SAP-dump (kolom V); zet 'AK meenemen' op Nee als het niet mee moet, of vul in kolom K een eigen percentage in."),
    ("7. Gebieden", "De lijst met gebieden. Die kun je per project kiezen."),
    ("8. Funnel en target", "Per jaar de omzet uit de acquisitiefunnel en de target."),
    ("9. DAEB-omzet", "Per project de DAEB-omzet en welk deel in welk jaar valt (samen 100%)."),
]

CORRECTIE = [
    "Pas de bron of de aanvullende invoer aan zodat de cijfers kloppen.",
    "Klik op 'Prognose vastleggen'. De knop ziet dat de versie al bestaat en vraagt om een toelichting.",
    "De correctie krijgt een hogere revisie. Alleen de nieuwste revisie telt; de oude regel blijft bewaard. In de RDFO staat ● achter de versie.",
]

FAQ = [
    ("Hoe is de AK-dekking van een project berekend?",
     "Kijk op het tabblad AK-dekking. Per project zie je het AK %, de kosten per kwartaal, de geboekte AK t/m de peildatum (grootboekdump), "
     "de prognose voor de rest van het jaar (kosten na de peildatum / 1,05 x AK %), het totaal, het budget en de volgende jaren."),
    ("Hoe krijg ik mijn eerdere prognoses (Budget, Q1, Q2, Q3) erin?",
     "Eenmalig: klik op Start op 'Oude prognoses overzetten' en kies je oude RDFO-bestand. Het bestand leest elk tabblad met Budget, Q1, Q2, Q3 of Q4 in de naam (bijvoorbeeld 'Q1' of 'Prognose Q1') "
     "(omzet, BOR en AK-dekking per WBS, in EUR x 1.000) en zet ze als vastgestelde versies in de Historie. Een versie die er al in staat, wordt overgeslagen."),
    ("Ik wil met een leeg bestand beginnen", "Gebruik RDFO_template.xlsm (al leeg), of klik op Start op 'Bestand leegmaken' en typ LEEG. Daarna: Bronnen inlezen › Nieuwe projecten toevoegen › Type en gebied kiezen."),
    ("Na inlezen is alles leeg of staat 'Nieuwe projecten: … FOUT'", "De WBS-elementen uit je export staan nog niet in de projectlijst; alleen projecten in Invoer › 2 tellen mee. Klik op 'Nieuwe projecten toevoegen' en controleer per regel Type en Woningen meetellen."),
    ("Controle geeft waarschuwingen over woningen of bedragen (aantal, etalage, cumulatief, portefeuille, PEW)",
     "Dat zijn verschillen in de SAP-export zelf; jij hoeft ze niet op te lossen om vast te zetten. Onderaan Controle staat per project welke controle afwijkt. Meld die projecten bij de verkoopadministratie of projectcontrol."),
    ("Controle meldt 'ontbrekende kolommen'", "Het bestand zoekt kolommen op naam (zie Handleiding › Veldwoordenboek). Na inlezen noemt de melding welke kolomnamen niet gevonden zijn. Hernoem de kolom in de export of pas de naam aan in de export-layout in SAP."),
    ("De knoppen doen niets / gele balk 'Macro's zijn uitgeschakeld'",
     "Klik op 'Inhoud inschakelen'. Komt het bestand uit e-mail of internet: sluit het, klik met rechts op het bestand › Eigenschappen › vink 'Blokkering opheffen' aan."),
    ("Mijn organisatie staat geen macro's toe",
     "Alles rekent ook zonder macro's. Rechtermuisknop op een tabbladnaam › Zichtbaar maken › Bron_ZPMW03 / Bron_Woningetalage: plak de export in A1. "
     "Vastleggen: kopieer op het verborgen tabblad Concept A8:R(laatste regel) en plak als waarden op de eerste vrije rij van Historie."),
    ("'Maak PowerPoint' geeft een foutmelding", "PowerPoint moet op deze computer geïnstalleerd zijn. Sluit open dialoogvensters in PowerPoint en probeer opnieuw."),
    ("Er staat 'n.b.' in de RDFO", "Er is (nog) geen vastgelegde waarde of aanvullende invoer. Dat is bewust geen nul."),
    ("Een oranje subtotaal in de RDFO", "Minstens één onderliggende post is n.b.; het subtotaal is dus onvolledig."),
    ("Nieuw rapportagejaar", "Zet in Invoer › 1 het rapportagejaar, het versiejaar van de upload en het versiejaar in de rapporten op het nieuwe jaar. Kolomkoppen schuiven mee."),
    ("Ik wil de rekenbladen of de Historie zien", "Klik op Start op 'Toon / verberg rekenbladen'. Ook Historie (alle vastgezette prognoses) staat daar."),
    ("Waar is het tabblad Stoplicht?", "Het stoplicht staat nu op het tabblad Bandbreedte, samen met de scenario's per project."),
    ("Excel meldt bij openen dat inhoud is hersteld en de knoppen werken niet",
     "Importeer de macro's handmatig: Alt+F11 › Bestand › Bestand importeren › macros/modRDFO.bas. Klik daarna met rechts op elke knop › Macro toewijzen."),
]

SPELREGELS = [
    "Wijzig nooit iets in de Historie. Alleen de knop 'Prognose vastleggen' voegt regels toe.",
    "Voeg nieuwe projecten altijd onderaan toe in Invoer › 2 (of met de knop). Niet sorteren of tussenvoegen: het stoplicht hoort bij de regel.",
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
    ("RDFO › AK-dekking", "ZPMW03 + AK-percentages (SAP-dump) + Invoer › 2", "Invoer › 3 als die is ingevuld; anders per project kosten uit ZPMW03 (Kosten LJ + Kosten RJ, daarna Kosten per jaar) / 1,05 x AK % (zoals de oude RDFO). Realisatie = geboekte AK uit de grootboekdump (zonder dump: Kosten LJ / 1,05 x AK %); boekingen zonder project tellen mee op regioniveau. FY = realisatie + kosten restant / 1,05 x AK %. AK % = eigen % of anders de SAP-dump; niet als 'AK meenemen' = Nee. Wordt mee vastgelegd.", "AK meenemen / eigen %", ""),
    ("RDFO › fees, overige, taakstelling EBIT", "—", "Som Invoer › 3 (FY en Actual YTD apart).", "Verplicht", ""),
    ("Outlook › DAEB, funnel, target", "Invoer › 8 en 9", "DAEB-omzet x percentage per jaar; funnel en target per jaar.", "Ja", "Komt DAEB-omzet uit de ZPMW DAEB-export?"),
    ("Outlook › kansen en risico's, reserves", "Invoer › 10 en 11", "Saldo per jaar (risico negatief); reserves x % gereed (cumulatief).", "Ja", ""),
    ("Bandbreedte", "ZPMW03 + Woningetalage + kans %", "Per project en jaar: optimistisch, basis, afslag, gewogen, 0-100 (zie Handleiding).", "Kans %, 0-100, instellingen", "Gelden 50% afslag en 90% optimistisch?"),
    ("RDFO › per gebied", "Projecten + Invoer › 2 (Gebied)", "Som per gebied van budget, vorige versie, geselecteerde versie en BOR.", "Gebied per project", ""),
    ("RDFO › EBIT-bijdrage", "—", "BOR + AK-dekking + fees + overige + taakstelling EBIT.", "—", "Bevestigen opbouw."),
    ("RDFO › BOR % / EBIT %", "—", "BOR % = (BOR + fees + overige resultaten + taakstelling EBIT) / Omzet old; EBIT % = EBIT / Omzet old; deling door nul afgevangen.", "—", ""),
    ("RDFO › Budget / Q1–Q4", "Historie", "Hoogste vastgestelde revisie per project/kengetal.", "Budget/Q1/Q2: fictieve historie", "Volledige jaarbedragen of delta’s in oude bladen?"),
    ("RDFO › Upload concept", "Bron_ZPMW03 / Bron_Woningetalage", "Live uit de bron; pas een versie na 'Prognose vastleggen'.", "—", ""),
    ("RDFO › Realisatie YTD / to go", "Historie of upload", "Altijd van dezelfde versie/peildatum als de FY-prognose.", "—", ""),
    ("Outlook › omzet/BOR per jaar", "ZPMW03", "Vorig jaar = −Bedrijfsopbrengsten VJ; lopend jaar = FY; daarna −OP.BOB per jaar t/m RJ+6; later = RJ+7, RJ+8 en onverdeeld.",
     "Budget: Historie (Budget, versiejaar)", ""),
    ("Outlook › woningverkopen per jaar", "Woningetalage", "Huidig jaar, Volgend jaar en de positieve jaarkolommen (2028, 2029, …).", "—",
     "Bevestigen dat positieve jaarkolommen verkopen in dat jaar zijn."),
    ("Orderintake (Outlook)", "ZPMW03 (kolom Orderintake, PEW opbrengst, PEW resultaat)",
     "Jaar/kwartaal = datum in 'Orderintake'; bedrag = −PEW opbrengst; BOR = −PEW resultaat.", "Budget: Invoer › 3 'Orderintake budget'",
     "Is 'Orderintake' de juiste datum en PEW-omzet het juiste bedrag? Orderintake wordt (nog) niet per versie vastgelegd."),
    ("Stoplicht › Verkoop", "Woningetalage", "Cumulatief verkocht / totaal woningen (of handmatig %); drempels instelbaar.", "Stoplicht", "Welke drempels gelden? Verkocht % t.o.v. start bouw?"),
    ("Stoplicht › Omgevingsvergunning, netcongestie", "—", "Handmatige invoer; kleur via tabel uitgangspunten.", "Stoplicht", ""),
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
    ("Definitie", "Zijn de positieve jaarkolommen in de Woningetalage (2028, 2029, …) verkopen in dat jaar? Wat betekenen 'Rond' en de negatieve jaarkoppen?", "Jaarkolommen."),
    ("Data", "Kan één projectdefinitie meerdere WBS-elementen met eigen woningen hebben?", "Koppeling."),
    ("Data", "Waar komen type (opstal/deelneming), segment en koperstype in productie vandaan?", "Mapping-onderhoud."),
    ("Proces", "Wie stelt een versie vast en is er een vier-ogenprincipe?", "Vastleggen."),
    ("Proces", "Hoe migreren we historische versies uit de huidige RDFO?", "Eerste vulling Historie."),
    ("Techniek", "Zijn macro's toegestaan op de werkplekken (vertrouwde locatie / ondertekening)?", "Knoppen."),
    ("Techniek", "Capaciteit: 500 bronregels, 100 projecten, 5.000 historieregels — voldoende?", "Formulebereiken."),
]
