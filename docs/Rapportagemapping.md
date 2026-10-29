# Bron-naar-rapportagemapping

Elke rapportageregel met bron, berekening, benodigde aanvullende invoer en openstaande definitie. Ook in de werkmap (tab *Rapportagemapping*).

| Rapportageregel | Bron | Berekening | Aanvullende invoer | Openstaande definitie |
|---|---|---|---|---|
| RDFO › Omzet opstal | ZPMW03 + Mapping (Type=Opstal) | FY RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ; FY RJ+1 = −OP.BOB RJ+1; per WBS gesommeerd (VERW_Project). Versiekolommen: Vastlegging 'Omzet FY'. | — | Is OP.BOB RJ in productie een restant of volledig jaar? (Controle waarschuwt als Q1–Q3 ≠ 0.) |
| RDFO › Omzet deelnemingen | ZPMW03 + Mapping (Type=Deelneming) | Idem; bedragen al op aandeelbasis, aandeel niet toegepast. | — | Is Type ‘Deelneming’ in productie af te leiden uit een SAP-veld of blijft het mapping? |
| RDFO › Taakstelling omzet | — | Som Invoer_Aanvullend (rubriek, soort FY/Actual YTD). | Verplicht | Bron en eigenaar van taakstellingen. |
| RDFO › Interne eliminaties | — | Som Invoer_Aanvullend. | Verplicht | Welke projecten/bedrijven worden geëlimineerd en op welk niveau? |
| RDFO › Omzet old | — | Opstal + deelnemingen + taakstelling + interne eliminaties. | — | Bevestigen dat dit de oorspronkelijke opbouw is. |
| RDFO › IFRS-eliminaties | — | Som Invoer_Aanvullend. IFRS J/N alleen als memo, geen automatische eliminatie. | Verplicht | Regel voor IFRS-eliminatie (volledig/proportioneel/equity) per project. |
| RDFO › Omzet IFRS | — | Omzet old + IFRS-eliminaties. | — |  |
| RDFO › Projectresultaten / BOR | ZPMW03 | FY RJ = −Bruto (omzet) result. LJ − OP.BOR RJ; RJ+1 = −OP.BOR RJ+1. Geen ABS: verlies blijft negatief. | — | Behandeling verliesvoorzieningen (direct volledig of via OP.BOR)? |
| RDFO › AK-dekking | — | Som Invoer_Aanvullend (FY en Actual YTD apart). | Verplicht | Bestaande AK-methode (percentage per project? vast bedrag?) — niet generiek toegepast. |
| RDFO › Fees | — | Som Invoer_Aanvullend. | Verplicht | Bron van fees. |
| RDFO › Overige resultaten | — | Som Invoer_Aanvullend. | Verplicht |  |
| RDFO › Taakstelling EBIT | — | Som Invoer_Aanvullend. | Verplicht |  |
| RDFO › EBIT-bijdrage | — | BOR + AK-dekking + fees + overige resultaten + taakstelling EBIT. | — | Bevestigen opbouw. |
| RDFO › BOR % | — | Totale BOR / totale projectomzet (opstal + deelnemingen); deling door nul afgevangen. | — |  |
| RDFO › EBIT % | — | EBIT-bijdrage / Omzet old; deling door nul afgevangen. | — | Noemer bevestigen: Omzet old of Omzet IFRS? |
| RDFO › Budget / Q1–Q4 FY | Vastlegging | Actieve vastgestelde regels (hoogste revisie, eerste voorkomen) voor versiejaar P_VJ. | Budget/Q1/Q2: fictieve historie | Bevatten oudere productiebladen volledige jaarbedragen of delta’s t.o.v. budget? |
| RDFO › Upload concept | SRC_ZPMW03 / SRC_Woningetalage | Live uit de huidige bron; geen versie tot handmatig vastgelegd. | — |  |
| RDFO › Realisatie YTD | Vastlegging ('Omzet actual YTD') of upload | Altijd van dezelfde versie/peildatum als de geselecteerde FY-prognose. | — | Historische actuals Budget/Q1/Q2 niet beschikbaar → n.b. |
| RDFO › To go | — | Geselecteerde FY − realisatie YTD van dezelfde peildatum. Doeljaar > versiejaar: to go = FY. | — |  |
| RDFO › Δ vs budget / vorige | — | Geselecteerde versie − Budget; geselecteerde versie − direct voorgaande versie. | — |  |
| Dashboard › KPI-tegels | CALC_Versies (gefilterd) | Zelfde logica als RDFO, met filter type/project. | — |  |
| Dashboard › Jaarprognoseverloop | CALC_Versies | FY-omzet/BOR per vastgelegde versie; ontbrekende versie = geen staaf + label '(n.b.)'. | — |  |
| Dashboard › Realisatie vs to go | Projecten | Top 10 projecten op geselecteerde FY; gestapeld realisatie + to go. | — |  |
| Dashboard › Grootste mutaties | Projecten | Top 8 op \|Δ FY-omzet\| geselecteerde vs vorige versie. | — |  |
| Dashboard › Meerjarendoorkijk | ZPMW03 (upload) | −OP.BOB/−OP.BOR per jaar RJ..RJ+8 + onverdeeld; niet vastgelegd. | — |  |
| Projecten › alle kolommen | Vastlegging + VERW_Project + Invoer_Toelichting | Per WBS, zelfde definities als RDFO. Marge = BOR/omzet. | Toelichting |  |
| Projectdetail › planning | ZPMW03 | Realisatie t/m peildatum, Q4 rest, RJ+1 per kwartaal, RJ+2..RJ+8, onverdeeld. | — |  |
| Woningen › verkopen per maand/kwartaal | Woningetalage | hjr 01..12 (≤ peilmaand = actual), Volgend jr Q1–Q4 = RJ+1. | — | Bevestigen dat hjr-maanden na peildatum altijd prognose zijn. |
| Woningen › jaarverwachting per versie | Vastlegging ('Woningverkopen FY') | Per versie; upload = 'Huidig jaar' / 'Volgend jaar'. | — |  |
| Woningen › totaal / cumulatief / etalage | Woningetalage | Aantal woningen, Cumulatief Totaal, Etalage Totaal; al op aandeelbasis, NIET opnieuw gewogen. | — | Is ‘Aantal woningen’ in productie bruto of aandeelbasis? |
| Woningen › bruto (afgeleid) | Woningetalage + Mapping | Aandeelbasis ÷ aandeel% — alleen informatief. | — |  |
| Woningen › koperstype | Woningetalage (actual) + Invoer_Koperstype (prognose) | Actual Beleggers/Particulier; prognose zonder invoer = Onbekend. | Optioneel |  |
| Woningen › segment | Mapping ('Segment demo') | Fictieve demo-indeling; vervangen door echte segmentbron. | Mapping | Bron voor segment in productie. |
