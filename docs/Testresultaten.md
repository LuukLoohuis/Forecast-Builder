# Testresultaten

Uitgevoerd met `tools/test_rdfo.py`: kopieën van de werkmap worden aangepast zoals een gebruiker dat doet, door LibreOffice herberekend en vergeleken met een onafhankelijke pandas-berekening. De macro's zelf (VBA) zijn hier niet in Excel uitgevoerd; getest is dat het VBA-project leesbaar is en de knoppen naar bestaande macro's verwijzen.

| Nr | Test | Uitkomst / bewijs | Resultaat |
|---|---|---|---|
| 1 | Basis: werkmap rekent zonder formulefouten en controles | 169383 formules, 0 fouten; Controle: 0 FOUT, 0 WAARSCHUWING | GESLAAGD |
| 2 | RDFO-versiekolommen = fictieve historie; upload = bron (onafhankelijk nagerekend) | Budget/Q1/Q2/Q3/upload (EUR x 1.000): 118,525.4 / 116,754.0 / 112,768.2 / 114,096.8 / 114,096.8 | GESLAAGD |
| 3 | Q3-snapshot sluit aan op de bronbestanden | Q3 vastgelegd 114,096.8 = upload 114,096.8; som \|verschil\| alle 280 conceptregels = 0 | GESLAAGD |
| 4 | Ontbrekende Q4 en ontbrekende historische actuals niet als nul | Q4-kolom toont 'n.b.'; realisatie YTD Q3 = 74,162.9; vastgelegde versies zonder realisatie (blijven n.b.): 3 | GESLAAGD |
| 5 | Nieuwe bronupload verandert Budget/Q1/Q2/Q3 niet; alleen upload-concept beweegt | Budget–Q3 ongewijzigd: True; upload +335.2 (verwacht +335.2); Controle meldt afwijking upload vs vastgelegde Q3 als WAARSCHUWING: True | GESLAAGD |
| 6 | Kwartaalvastlegging Q4: Concept → Historie (dezelfde kopieerstap als de macro) | 280 regels, revisie [0]; Q4-kolom = 114,096.8; Budget–Q3 ongewijzigd; geselecteerd = Vastgesteld | GESLAAGD |
| 7 | Vastgelegde Q4 is onveranderlijk: latere upload herberekent haar niet | Q4 vóór/na nieuwe upload: 114,096.8 / 114,096.8; upload-concept 114,432.0 | GESLAAGD |
| 8 | Dubbele import veroorzaakt geen dubbeltelling | Q4 na 2x plakken: 114,096.8 (1x: 114,096.8); 280 regels als dubbel genegeerd en gemeld | GESLAAGD |
| 9 | Correctie op vastgestelde versie = nieuwe revisie, oude regel blijft, geen dubbeltelling | Q2 opstal +100,000 EUR (correctie +100.000, niet +5,5 mln); kolomkop gemarkeerd met ●; overige versies ongewijzigd | GESLAAGD |
| 10 | Correctie zonder toelichting wordt als FOUT gemeld | Controle: 1 FOUT | GESLAAGD |
| 11 | Woningen worden niet dubbel gewogen (aandeelbasis blijft aandeelbasis) | Totaal 1160 = brontotaal 1160 (met nogmaals aandeel zou 1,106 zijn); deelneming Het Kompas 50%: 24 st. (bruto afgeleid 48) | GESLAAGD |
| 12 | Verlies blijft zichtbaar (geen ABS) | Lindehaven, De Brink BOR -58.59 (x1.000) gemarkeerd 'Verlies'; RDFO BOR 10,477,198 = signed som 10,477,198 (ABS zou 10,702,738 geven) | GESLAAGD |
| 13 | Parkeergebouw zonder woningen: geen deling door nul | Projectmatrix toont 'n.v.t.', % verkocht leeg; 0 formulefouten | GESLAAGD |
| 14 | Alias Bedrijfsnaam/Naam maatschappij en gewijzigde kolomvolgorde | Header hernoemd + 'OP.BOB 2026' en 'Huidig jaar' van positie gewisseld: uitkomsten identiek, 0 FOUT | GESLAAGD |
| 15 | Nieuwe en vervallen projecten verschijnen als uitzondering (niet stil weg) | nieuw: (1, 'FOUT'); vervallen: (1, 'WAARSCHUWING'); uitzonderingslijst bevat G.900099 en Boswaard, Parkeergebouw | GESLAAGD |
| 16 | Doeljaar 2027: jaarprognoses per versie, geen realisatie (n.v.t.), to go = FY | Q3 2027 = 218,828.5; upload 2027 = 218,828.5; realisatie 'n.v.t.' | GESLAAGD |
| 17 | Fictieve demo-aannames alleen zichtbaar na expliciete keuze (en dan gemeld) | P_Demo=Nee: AK-dekking 'n.b.'; P_Demo=Ja: AK-dekking -1,400 en Controle-waarschuwing | GESLAAGD |
| 18 | Stoplicht (verkoop, omgevingsvergunning, netcongestie): kleur volgt uit invoer en uitgangspunten | Watergaarde (0% verkocht, net Laag) → [1, 3, 2]; Parkeergebouw (geen woningen, net Hoog) → [0, 3, 1] (0 = leeg/n.v.t.) | GESLAAGD |
| 19 | Stoplicht: uitgangspunten aanpassen kleurt alle lichten direct mee | Groengrens 0% en netcongestie Laag = groen → Watergaarde [3, 3, 3] | GESLAAGD |
| 20 | Orderintake per jaar/kwartaal = bron (kolom Orderintake, PEW opbrengst) | 2026 t/m peildatum: 337,740,000 EUR in 15 projecten (onafhankelijk: 337,740,000) | GESLAAGD |
| 21 | Outlook: omzet, woningverkopen en orderintake voor elk jaar (2025 t/m 2032) | woningverkopen 2026/2028 = 360/252; omzet 2027/2028 = 218.8/117.8 mln (= bron); 2029+ in bron 0 | GESLAAGD |
| 22 | PowerPoint-inhoud staat klaar (tabbladen PPT en Dias) | 20 projecten in de projecttabel; kernboodschap: 'Omzet FY 2026: € 114.1 mln, -4.4 mln t.o.v. budget  ·  BOR 9.2%' | GESLAAGD |
| 23 | PowerPoint: lege toelichtingskolom en jaren zonder cijfers vallen weg | gewicht toelichting = 0; gewichten jaarkolommen 2025-2032 = [1, 1, 1, 1, 0, 0, 0, 0] | GESLAAGD |
| 24 | Macro-bestand: VBA-project leesbaar, alle knoppen verwijzen naar bestaande macro's | 24 VBA-modules (oletools); 7 knoppen → ['RDFO_BronnenInlezen', 'RDFO_MaakPowerPoint', 'RDFO_PrognoseVastleggen', 'RDFO_ToonRekenbladen']. Let op: uitvoering in Excel zelf is hier niet getest. | GESLAAGD |
