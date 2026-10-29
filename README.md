# Forecast-Builder — RDFO-prototype

Een onderhoudbaar RDFO-Excelbestand op **uitsluitend fictieve data**. Het herontwerp scheidt de bronnen, de projectmapping, de berekeningen, de vastgelegde historie en de rapportage van elkaar. Het bestand gebruikt geen macro's en geen draaitabellen.

**Oplevering:** [`RDFO_prototype.xlsx`](RDFO_prototype.xlsx)

| Document | Inhoud |
|---|---|
| [docs/Handleiding.md](docs/Handleiding.md) | Bronnen vervangen, verversen, een kwartaal vastleggen (handmatig), correcties, ontwerpkeuze |
| [docs/Rapportagemapping.md](docs/Rapportagemapping.md) | Per rapportageregel: bron, berekening, aanvullende invoer en openstaande definitie |
| [docs/Productievragen.md](docs/Productievragen.md) | Vragen die vóór productie beantwoord moeten worden |
| [docs/Testresultaten.md](docs/Testresultaten.md) | Uitgevoerde scenario-tests met uitkomst |

## Opbouw van de werkmap

| Laag | Tabbladen |
|---|---|
| Rapportage | Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Controle |
| Invoer | Parameters, Invoer_Aanvullend, Invoer_Toelichting, Invoer_Koperstype, Mapping |
| Versiebeheer | Vastlegging (grootboek waaraan alleen regels worden toegevoegd), Vastleggen (concept uit de upload), Importlog |
| Berekening | VERW_Project (één regel per WBS), CALC_Versies |
| Staging | STG_ZPMW03, STG_Woningetalage (kolommen worden op naam gezocht, met aliassen en dubbele headers) |
| Bron | SRC_ZPMW03, SRC_Woningetalage (de export ongewijzigd plakken) |
| Documentatie | Veldwoordenboek, Rapportagemapping, Productievragen, Handleiding, Testresultaten |

## Opnieuw bouwen en testen

```bash
python tools/release.py <pad-naar-recalc.py>
```

Dit script bouwt de werkmap uit `data/` en draait daarna de scenario-tests (`tools/test_rdfo.py`). Elke test past een kopie van de werkmap aan, laat LibreOffice herberekenen en vergelijkt de uitkomst met een onafhankelijke pandas-berekening. Tot slot zet het script de berekende waarden als cache in het bestand, zodat ook een preview de cijfers toont. Excel rekent bij het openen alles opnieuw.

Alle data in `data/` is synthetisch. De demo-aannames zijn geen vastgestelde SAP-definities.
