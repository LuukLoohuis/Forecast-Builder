# Van bron naar rapport

Elke rapportregel met bron, berekening, aanvullende invoer en open punt. Ook in de werkmap (Handleiding › bijlage B).

| Rapportregel | Bron | Berekening | Aanvullende invoer | Open punt |
|---|---|---|---|---|
| RDFO › Omzet opstal | ZPMW03 + Invoer › 2 (Type=Opstal) | FY RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ; FY RJ+1 = −OP.BOB RJ+1; per WBS gesommeerd. Versiekolommen: Historie 'Omzet FY'. | — | Is OP.BOB RJ in productie een restant of een volledig jaar? (Controle waarschuwt als Q1–Q3 ≠ 0.) |
| RDFO › Omzet deelnemingen | ZPMW03 + Invoer › 2 (Type=Deelneming) | Idem; al op aandeelbasis. | — | Type in productie uit SAP-veld of mapping? |
| RDFO › Taakstelling omzet | — | Som Invoer › 3. | Verplicht | Bron en eigenaar van taakstellingen. |
| RDFO › Interne eliminaties | — | Som Invoer › 3. | Verplicht | Welke projecten/bedrijven en op welk niveau? |
| RDFO › Omzet old | — | Opstal + deelnemingen + taakstelling + interne eliminaties. | — | Bevestigen dat dit de bestaande opbouw is. |
| RDFO › IFRS-eliminaties | — | Som Invoer › 3. IFRS J/N alleen als memo. | Verplicht | Regel voor IFRS-eliminatie per project. |
| RDFO › Omzet IFRS | — | Omzet old + IFRS-eliminaties. | — |  |
| RDFO › Projectresultaten / BOR | ZPMW03 | FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ; RJ+1 = −OP.BOR RJ+1. Geen ABS. | — | Behandeling verliesvoorzieningen. |
| RDFO › AK-dekking, fees, overige, taakstelling EBIT | — | Som Invoer › 3 (FY en Actual YTD apart). | Verplicht | Bestaande AK-methode — niet generiek toegepast. |
| RDFO › EBIT-bijdrage | — | BOR + AK-dekking + fees + overige + taakstelling EBIT. | — | Bevestigen opbouw. |
| RDFO › BOR % / EBIT % | — | Totale BOR / projectomzet; EBIT / Omzet old; deling door nul afgevangen. | — | Noemer EBIT% bevestigen. |
| RDFO › Budget / Q1–Q4 | Historie | Hoogste vastgestelde revisie per project/kengetal. | Budget/Q1/Q2: fictieve historie | Volledige jaarbedragen of delta’s in oude bladen? |
| RDFO › Upload concept | Bron_ZPMW03 / Bron_Woningetalage | Live uit de bron; pas een versie na 'Prognose vastleggen'. | — |  |
| RDFO › Realisatie YTD / to go | Historie of upload | Altijd van dezelfde versie/peildatum als de FY-prognose. | — |  |
| Orderintake | ZPMW03 (kolom Orderintake, PEW opbrengst, PEW resultaat) | Jaar/kwartaal = datum in 'Orderintake'; bedrag = −PEW opbrengst; BOR = −PEW resultaat. | Budget: Invoer › 3 'Orderintake budget' | Is 'Orderintake' de juiste datum en PEW-omzet het juiste bedrag? Orderintake wordt (nog) niet per versie vastgelegd. |
| Stoplicht › Verkoop | Woningetalage | Cumulatief verkocht / totaal woningen (of handmatig %); drempels instelbaar. | Stoplicht | Welke drempels gelden? Verkocht % t.o.v. start bouw? |
| Stoplicht › Omgevingsvergunning, RvS, netcongestie | — | Handmatige invoer; kleur via tabel uitgangspunten. | Stoplicht |  |
| Woningen | Woningetalage | hjr 01..12 (≤ peilmaand = realisatie), Volgend jr Q1–Q4, jaarverwachting per versie uit Historie. | Koperstype: Invoer › 5 | Is 'Aantal woningen' bruto of aandeelbasis? |
