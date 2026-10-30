# Forecast-Builder — RDFO-prototype

Een overzichtelijk RDFO-bestand met knoppen (macro's), gebouwd op **uitsluitend fictieve data**.

**Oplevering:** [`RDFO_prototype.xlsm`](RDFO_prototype.xlsm). Open het bestand en klik op *Inhoud inschakelen*.

| Knop (tabblad Start) | Wat gebeurt er |
|---|---|
| 1 Bronnen inlezen | Je kiest de ZPMW03- en Woningetalage-export en vult de peildatum en versie in. Daarna rekent het bestand alles opnieuw en opent het tabblad Controle. |
| 2 Prognose vastleggen | Voegt de jaarprognose (bijv. Q3) toe aan Historie. Bestaat de versie al, dan wordt het een correctie met toelichting. |
| 3 Maak PowerPoint | Maakt een presentatie in geel/donkerblauw met kerncijfers, RDFO, projecten met stoplicht, prognoseverloop, outlook, orderintake, woningen en aandachtspunten. Tabellen zijn echte PowerPoint-tabellen (zelf aan te passen); een lege toelichtingskolom en jaren zonder cijfers worden weggelaten. Een voorbeeld staat in [docs/Voorbeeld_PowerPoint.pptx](docs/Voorbeeld_PowerPoint.pptx). |
| Toon / verberg rekenbladen | Toont of verbergt de bron- en rekenbladen. |

Er zijn 12 zichtbare tabbladen: Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Outlook, Stoplicht, Controle, Invoer, Historie en Handleiding. Bronnen en tussenberekeningen staan op verborgen bladen.

Grafieken volgen de IBCS/Zebra-notatie: realisatie massief, prognose gearceerd, budget grijs of als streep, afwijkingen groen of rood, en watervallen voor de brug van budget naar prognose.

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
