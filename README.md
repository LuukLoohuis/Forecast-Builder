# Forecast-Builder — RDFO-prototype

Een overzichtelijk RDFO-bestand met knoppen (macro's), gebouwd op **uitsluitend fictieve data**.

**Oplevering:**
- [`RDFO_prototype.xlsm`](RDFO_prototype.xlsm): gevuld met fictieve demodata.
- [`RDFO_template.xlsm`](RDFO_template.xlsm): dezelfde werkmap, maar leeg (geen data, geen historie). Hiermee begin je met echte data.

Open het bestand en klik op *Inhoud inschakelen*.

| Knop (tabblad Start) | Wat gebeurt er |
|---|---|
| 1 Bronnen inlezen | Je kiest de ZPMW03- en Woningetalage-export en vult de peildatum en versie in. Daarna rekent het bestand alles opnieuw en opent het tabblad Controle. |
| 2 Prognose vastleggen | Voegt de jaarprognose (bijv. Q3) toe aan Historie. Bestaat de versie al, dan wordt het een correctie met toelichting. |
| 3 Maak PowerPoint | Maakt een presentatie in geel/donkerblauw met kerncijfers, RDFO, projecten met stoplicht, prognoseverloop, outlook, orderintake, woningen en aandachtspunten. Tabellen zijn echte PowerPoint-tabellen (zelf aan te passen); een lege toelichtingskolom en jaren zonder cijfers worden weggelaten. Een voorbeeld staat in [docs/Voorbeeld_PowerPoint.pptx](docs/Voorbeeld_PowerPoint.pptx); de lege indeling (wat komt waar, uit welk tabblad) in [docs/RDFO_PowerPoint_template.pptx](docs/RDFO_PowerPoint_template.pptx). |
| Nieuwe projecten toevoegen | Zet WBS-elementen uit ZPMW03 die nog niet op de Projectlijst staan onderaan de Projectlijst (gebeurt ook vanzelf bij Bronnen inlezen). |
| Bestand leegmaken | Wist bronnen, projectlijst, invoer, stoplicht en Historie. Instellingen blijven staan. |
| Toon / verberg rekenbladen | Toont of verbergt de bron- en rekenbladen en de Historie. |

Er zijn 15 zichtbare tabbladen: Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Outlook, Bandbreedte, Kansen en risico's, Reserves, Projectlijst, AK-dekking, Controle, Invoer en Handleiding. Alle invoer per project (keuzes, details, AK %, reserves, stoplicht en scenario) staat op de Projectlijst; Invoer › 2 en de projecttabel op Bandbreedte verwijzen ernaar.
- **Bandbreedte** bevat het stoplicht en vijf scenario's per jaar: optimistisch, basis, afslag, gewogen en 0-100.
- **Outlook** toont per jaar ook DAEB, acquisitiefunnel, target, AK-dekking, kansen en risico's, reserves en de bandbreedte.
- **RDFO** toont onderaan het overzicht per gebied.
- **Woningen** toont de verkopen per productsegment en per koperstype.

Bronnen, tussenberekeningen en Historie staan op verborgen bladen. Je ziet ze via de knop *Toon / verberg rekenbladen*.

Grafieken volgen de IBCS/Zebra-opbouw in Heijmans-kleuren:
- realisatie donkerblauw;
- prognose geel;
- budget grijs;
- afwijkingen groen of rood;
- bandbreedte als gele band van het laagste naar het hoogste scenario.

| Document | Inhoud |
|---|---|
| [docs/Handleiding.md](docs/Handleiding.md) | Elk kwartaal in 5 stappen, tabbladen, stoplicht, correcties, problemen oplossen |
| [docs/Rapportagemapping.md](docs/Rapportagemapping.md) | Van bron naar rapport, per regel |
| [docs/Productievragen.md](docs/Productievragen.md) | Vragen vóór productie |
| [docs/Testresultaten.md](docs/Testresultaten.md) | Uitgevoerde scenario-tests |
| [docs/Overdracht_Copilot.md](docs/Overdracht_Copilot.md) | Overdracht voor Copilot/ontwikkelaar, inclusief opzet scenarioanalyse |
| [macros/](macros) | VBA-broncode (ook handmatig te importeren) |

## Opnieuw bouwen en testen

```bash
python tools/release.py <pad-naar-recalc.py>
```

Dit script doet het volgende:
1. Het bouwt de werkmap uit `data/` (`tools/build_rdfo.py`).
2. Het draait de scenario-tests (`tools/test_rdfo.py`).
3. Het zet de berekende waarden als cache in het bestand.
4. Het verpakt het geheel als `.xlsm`, met het VBA-project uit `macros/` (`tools/vbaproject.py`, `tools/package_xlsm.py`).

De formules zijn getest in LibreOffice. De macro's zijn niet in Excel uitgevoerd: getest is alleen dat het VBA-project leesbaar is en dat elke knop naar een bestaande macro verwijst.
