# Prompt: Forecast in één Excel-bestand (SF + SAP → CF → PowerPoint)

> Status: concept. Aannames staan tussen [vierkante haken] en moeten nog bevestigd worden met de echte templates.

---

## Rol & doel

Je bent een ervaren financial controller en Excel-modelbouwer. Bouw één Excel-werkmap
(`Forecast_[JJJJ-MM].xlsx`) die het huidige forecastproces vervangt. Nu werken we met
losse bestanden (Salesforce-dump, SAP-dump, vorige-periode-bestand, CF-bestand,
PowerPoint). Straks moet alles vanuit één bestand lopen: data erin plakken, voorstel
beoordelen, handmatig bijsturen, en de PowerPoint ververst vanzelf.

## Huidige situatie (input)

1. **SF-dump** – export uit Salesforce met [opportunities: account, omschrijving, fase,
   kans %, bedrag, verwachte closedatum, eigenaar]. Basis voor de opbrengsten-forecast.
2. **SAP-dump huidige periode** – export uit SAP met [boekingen/realisatie per
   grootboekrekening, kostenplaats, maand, bedrag].
3. **SAP-dump vorige periode** – nu kopiëren we elke periode handmatig de SAP-dump van
   de vorige periode naar het tabblad *Vorige periode*. Dat is foutgevoelig en moet
   eenvoudiger.
4. **CF-overzicht** – [cashflow/forecast-overzicht] met kosten en opbrengsten per maand.
5. **PowerPoint** – slides zijn gekoppeld aan tabellen/grafieken uit het CF-overzicht.

## Gewenste opbouw van de werkmap

| # | Tabblad | Doel |
|---|---------|------|
| 0 | `Start` | Periode kiezen (één cel), instructies stap-voor-stap, status van de checks |
| 1 | `IN_SF` | SF-dump 1-op-1 plakken (vaste kolomkoppen, als Excel-tabel) |
| 2 | `IN_SAP` | SAP-dump 1-op-1 plakken, **met kolom `Periode`** |
| 3 | `IN_SAP_Vorig` | Vorige periode – zie "Periode-overgang" hieronder |
| 4 | `Mapping` | Grootboekrekening → CF-regel; SF-fase → standaard kans %; kostenplaats → afdeling |
| 5 | `Voorstel` | Berekend forecastvoorstel per CF-regel per maand (geen handwerk) |
| 6 | `CF` | Het forecastoverzicht: realisatie + voorstel + handmatige override |
| 7 | `Analyse` | Verschil t.o.v. vorige forecast en budget (brug/waterval) met toelichting |
| 8 | `PPT_Output` | Vaste, benoemde bereiken en grafieken waar PowerPoint aan gekoppeld is |
| 9 | `Checks` | Controles: totalen SAP = CF-realisatie, niet-gemapte rekeningen, lege velden |

## Kern: CF met voorstel én handmatige forecast

Per CF-regel (kosten en opbrengsten) per toekomstige maand vier kolommen/lagen:

1. **Voorstel** (berekend, grijs, niet aanpasbaar)
2. **Methode** (keuzelijst per regel): `Run-rate (gem. laatste 3 mnd)` | `Vorige forecast` |
   `Budget` | `SF-pipeline gewogen` | `Handmatig`
3. **Override** (geel invoerveld) – als gevuld, wint dit altijd
4. **Definitief** = `ALS(Override<>""; Override; Voorstel volgens Methode)`

Plus een kolom **Toelichting** (verplicht als override > [X% of €Y] afwijkt van voorstel).
Gerealiseerde maanden komen altijd uit SAP en zijn niet overschrijfbaar.

### Voorstelregels
- **Opbrengsten**: SF-bedrag × kans % (uit SF of uit `Mapping` per fase), toegewezen aan
  de maand van de verwachte closedatum [eventueel gespreid over looptijd]. Gesloten-gewonnen
  deals die al in SAP staan niet dubbel tellen.
- **Kosten**: standaard run-rate (gemiddelde laatste 3 gerealiseerde maanden uit SAP),
  met optie voor vorige forecast of budget. Eenmalige posten markeerbaar zodat ze niet in de
  run-rate meetellen.

## Periode-overgang (vorige periode)

Vervang het handmatig kopiëren door één van deze opties (geef advies welke past):
- **A. Eén SAP-tabel met kolom `Periode`** – nieuwe dump onderaan plakken; "vorige periode"
  is een filter op `Periode = gekozen periode − 1`. Geen kopiëren meer.
- **B. Power Query uit map** – dumps opslaan als `SAP_JJJJ-MM.xlsx` in een vaste map;
  Power Query laadt automatisch huidige en vorige periode.
- **C. Knop/macro "Periode afsluiten"** – kopieert huidige naar vorige, bewaart de definitieve
  forecast als snapshot (`Snapshot_Forecast`) en leegt de invoertabellen.

Ook de **definitieve forecast** van elke periode wordt als snapshot bewaard, zodat
`Analyse` altijd "deze forecast vs. vorige forecast" kan tonen.

## Koppeling met PowerPoint

- Alles wat in de PowerPoint staat komt uit `PPT_Output`, op **benoemde bereiken** met vaste
  grootte (nooit rijen invoegen in die bereiken).
- Koppeling via Plakken speciaal → Koppeling (of [VBA-export]), zodat "Koppelingen bijwerken"
  in PowerPoint genoeg is.
- Geef een lijst: welke slide ↔ welk benoemd bereik/grafiek.

## Eisen

- Werkt in Excel 365, Nederlandse instellingen (puntkomma's in formules, € notatie).
- Invoercellen geel, formules grijs/beveiligd, één kleurlegenda op `Start`.
- Geen hardcoded getallen in formules; alle parameters op `Start` of `Mapping`.
- Gebruik Excel-tabellen en `SOMMEN.ALS`/`X.ZOEKEN`; macro's alleen als optie C gekozen wordt.
- Elke check op `Checks` geeft OK/FOUT en staat samengevat op `Start`.

## Wat ik terug wil

1. Korte bevestiging van de aannames en vragen over wat onduidelijk is (eerst vragen,
   niet gokken).
2. Advies voor de periode-overgang (A/B/C) met voor- en nadelen.
3. De werkmap zelf, gevuld met voorbeelddata in de structuur van de bijgevoegde templates.
4. Een werkinstructie van max. 1 A4: "maandafsluiting in 6 stappen".

## Bijlagen (nog toevoegen)

- [ ] Voorbeeld SF-dump (geanonimiseerd mag)
- [ ] Voorbeeld SAP-dump
- [ ] Huidig CF-bestand
- [ ] PowerPoint-template
