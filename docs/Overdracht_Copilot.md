# Overdracht RDFO-prototype → Copilot

Dit document is bedoeld als startcontext voor een AI-assistent (Copilot) of ontwikkelaar die verder bouwt aan het RDFO-bestand. Lees het helemaal voordat je iets wijzigt.

---

## 1. Wat is dit?

Een Excel-werkmap (`RDFO_prototype.xlsm`) voor de kwartaalprognose van een bouw-/ontwikkelregio (Heijmans-huisstijl: geel `#F5B800`, donkerblauw `#14284B`). De werkmap:

- leest twee SAP-exports in: **ZPMW03** (financieel per WBS-element) en **Woningetalage** (woningverkopen per Stufe 3);
- koppelt die via het tabblad **Projectlijst** (WBS, naam, type, gebied + 4 keuzes; sleutel WBS-element = Stufe 3). Invoer › 2 verwijst daarnaar en bevat alleen de details per project;
- rekent per project en per versie (Budget, Q1–Q4) omzet, BOR (bruto omzetresultaat), realisatie, to go, woningverkopen en orderintake;
- legt elke kwartaalprognose onveranderlijk vast in **Historie** (correcties = nieuwe revisie, nooit overschrijven);
- rapporteert in Dashboard, RDFO, Projecten, Woningen, Outlook, Bandbreedte, Kansen en risico's en AK-dekking (per kwartaal);
- maakt met één knop een PowerPoint (native, bewerkbare tabellen + grafieken).

Alle data in `data/` is **fictief**. De oorspronkelijke opdracht staat in `docs/opdracht.txt`.

---

## 2. Belangrijk: het Excelbestand wordt gegenereerd

`RDFO_prototype.xlsm` is een **bouwresultaat**. De bron is Python + VBA in deze repository:

| Pad | Inhoud |
|---|---|
| `tools/build_rdfo.py` | Bouwt de hele werkmap met openpyxl: bladen, formules, named ranges, opmaak, grafieken, Dias/PPT-specificatie (~3.300 regels). |
| `tools/xlhelpers.py` | Hulpfuncties (stijlen, `add_name`, `input_style`, `landscape`, …). |
| `tools/docs_content.py` | Teksten voor Handleiding, Start-pagina en docs (stappen, FAQ, mapping). |
| `macros/modRDFO.bas` | Alle VBA: bronnen inlezen, nieuwe projecten toevoegen, leegmaken, prognose vastleggen, PowerPoint maken, Outlook-jaren. |
| `macros/shStoplicht.cls` | Blad-event Stoplicht (stempelt WBS bij invoer). |
| `tools/vbaproject.py` | Eigen generator van `vbaProject.bin` (MS-CFB/MS-OVBA). Niet aanraken tenzij nodig. |
| `tools/package_xlsm.py` | Maakt van .xlsx een .xlsm: VBA-project + knoppen (vormen met `macro=`). Knoppen staan in `BUTTONS` in `build_rdfo.py`. |
| `tools/inject_cache.py` | Zet berekende waarden (uit LibreOffice) als cache in het bestand, zodat het direct goed opent. |
| `tools/test_rdfo.py` | 24 scenario-tests (bouwt varianten, herberekent in LibreOffice, controleert uitkomsten). |
| `tools/release.py` | Alles in één keer: tests → bouwen → herberekenen → cache → .xlsm → docs → voorbeeld-PPT. |
| `tools/preview.py`, `tools/preview_pptx.py` | Render bladen naar PNG; bouw `docs/Voorbeeld_PowerPoint.pptx` (python-pptx) uit dezelfde specificatie als de macro. |

**Bouwen en testen**

```bash
python tools/release.py <pad-naar-recalc.py> [werkmap]
```

`recalc.py` herberekent een xlsx met LibreOffice (headless) en rapporteert formulefouten. Enige toegestane fouten: `#N/A` in `Outlook!P..W` (bewuste gaten in grafieklijnen).

**Werk je alleen in Excel (geen Python)?** Dan kun je `RDFO_prototype.xlsm` direct aanpassen, maar die wijzigingen gaan verloren bij de volgende build. Neem ze dan later over in `build_rdfo.py` / `macros/`. VBA-wijzigingen: exporteer de module en zet hem in `macros/`.

---

## 3. Architectuur van de werkmap

**Zichtbare bladen (14):** Start, Dashboard, RDFO, Projecten, Woningen, Outlook, Bandbreedte, Kansen en risico's, Reserves (overzicht; invoer op Projectlijst L–Q; totaal ook onderaan RDFO, dia T_RESS/T_RES), Projectlijst, AK-dekking, Controle, Invoer, Handleiding. Historie en Projectdetail zijn verborgen (knop *Toon / verberg rekenbladen*). Controle toont na inlezen alleen FOUT/WAARSCHUWING (macro `RDFO_ControleWissel`).

**Verborgen bladen:** `Bron_ZPMW03`, `Bron_Woningetalage` (ruwe export, kop in rij 1) → `STG_ZPMW03`, `STG_Woningetalage` (kolommen op naam gezocht, afgeleide velden, controlekolommen `C_*`) → `VERW_Project` (1 rij per project in de projectlijst) → `CALC_Versies` (totalen per versie) → rapportbladen. Verder `Concept` (wat vastgelegd wordt), `Lijsten`, `PPT` (teksten + dia-specificatie `PPT_Spec`), `Dias` (tabellen/grafieken voor de PowerPoint).

**Gegevensstroom**

```
SAP-export ─► Bron_* ─► STG_* ─► VERW_Project ─► CALC_Versies ─► Dashboard / RDFO / Outlook / …
                                     ▲                 ▲
                         Invoer › 2 (projectlijst)   Historie (vastgelegde versies)
```

**Named ranges** zijn de interface tussen bladen en VBA. Voorvoegsels:
`P_*` instellingen (Invoer › 1), `Sel_*` selectie Dashboard, `Map_*` projectlijst, `A_*` aanvullende posten, `T_*` toelichtingen, `KT_*` koperstype, `L_*` Historie-grootboek, `V_*` verwerking, `SL_*` stoplicht, `OUT_*` Outlook, `PPT_*` / `K_*` / `T_RDFO`, `T_PROJ1..3`, `T_OUT`, `T_OI*` voor de PowerPoint.

**Capaciteit:** 500 bronregels per export, 100 projecten in de projectlijst, 5.000 regels Historie.

---

## 4. Vaste rekenafspraken (niet breken)

1. Bronbedragen in hele EUR; rapportage toont EUR × 1.000 of mln (alleen weergave).
2. In ZPMW03 zijn opbrengsten en resultaat **negatief**, kosten positief → expliciete tekenomzetting naar management (omzet positief). **Nooit `ABS`**: verlies moet negatief blijven (er is een test op).
3. FY omzet RJ = −Bedrijfsopbrengsten LJ − OP.BOB RJ (LJ = realisatie t/m peildatum, OP.BOB RJ = restant).
4. Bedragen én woningen zijn al op **aandeelbasis**; aandeel nooit opnieuw toepassen.
5. Ontbrekend = **`n.b.`**, nooit stil 0. Niet van toepassing = `n.v.t.`
6. Vastgelegde versies zijn onveranderlijk; correctie = nieuwe revisie met verplichte toelichting.
7. Kolommen worden op **naam** gezocht (aliassen mogelijk), niet op positie.
8. Grafieken in IBCS/Zebra-stijl: realisatie massief donkerblauw, prognose gearceerd, budget grijs/streep, afwijking groen/rood.

---

## 5. Huidige status (oktober 2026)

**Werkt (getest in LibreOffice, 24/24 tests):** import-formules, mapping-controles, versiebeheer en correcties, RDFO, Outlook 2025–2032 (jaren zonder cijfers vallen weg), stoplicht (verkoop, omgevingsvergunning, netcongestie) op basis van uitgangspunten, orderintake, PowerPoint-specificatie, Controle met 40+ checks.

**Niet getest:** de VBA-macro's zijn **nooit in echte Excel uitgevoerd** (alleen gecontroleerd dat het VBA-project leesbaar is en elke knop naar een bestaande macro wijst). Dit is het grootste risico.

**Bekende open punten**

| # | Punt | Toelichting |
|---|---|---|
| 1 | Inlezen met echte exports | Gebruiker meldde dat uploaden niet lukte. `LeesBron` is robuuster gemaakt (kopregel in eerste 30 rijen, koppen opgeschoond, geen klembord) maar nog niet met echte bestanden getest. Vraag de foutmelding en de kolomkoppen. |
| 2 | Kolomnamen Woningetalage | Bij echte data ontbraken 3 kolommen. Na inlezen toont de melding welke. Oplossen via aliassen in `e_fields()` / `z_fields()`. |
| 3 | Projectlijst | Echte WBS-elementen staan niet in de demo-projectlijst. Knop *Nieuwe projecten toevoegen* vult ze aan (Type standaard Opstal → handmatig Deelneming zetten). |
| 4 | Tekenconventie echte data | Controle 20 (OP.BOB + Kosten ≠ OP.BOR) gaf 45 FOUT bij echte data; vermoedelijk ander teken voor kosten. Verifiëren met een echte export. |
| 5 | Demo-aannames | Zie `docs/Productievragen.md` (definitie OP.BOB RJ, betekenis negatieve jaarkoppen Woningetalage, kolom "Rond", koperstype). |
| 6 | Weergave | Gebruiker vindt het "nog niet helemaal zoals het moet". Vraag concreet welke bladen/grafieken. |

---

## 6. Scenarioanalyse en uitbreidingen (november 2026, gebouwd)

Gebouwd op basis van de oorspronkelijke RDFO van de gebruiker (tabbladen BHTS, Bandbreedte 2027/2028, Kansen en risico's, Reserves, AK dekking, FY per gebied, Woningen):

| Onderdeel | Waar | Rekenregel |
|---|---|---|
| Bandbreedte (5 scenario's) | Tabblad **Bandbreedte** (vervangt Stoplicht) + verborgen **SCEN** | Per project en jaar RJ..RJ+4: optimistisch (verkoop haalt `SC_Opt`, max `SC_Max` x), basis (upload), afslag (`SC_Afslag` van de geplande verkopen), gewogen (x kans %), 0-100 (kans ≥ `SC_Grens` of Ja/Nee). Realisatie blijft altijd staan. |
| DAEB, funnel, target | Invoer › 8 en 9 → Outlook | DAEB-omzet x % per jaar; funnel en target per jaar. |
| AK-dekking | Projectlijst (AK % eigen) of SAP-dump → tabblad AK-dekking (per kwartaal: t/m `AK_PeilQ` geboekt uit STG_AKR `AKR_Kw`, daarna kosten kwartaal / `P_AKFactor` x AK %) → VERW `AKYTD`, `AK0..AK4` → Outlook en RDFO | Jaar = Q1..Q4; regio = projecten + `AK_NietToegewezen`; aanvullende invoer gaat voor; kengetal `AK-dekking FY` wordt mee vastgelegd. |
| Kansen en risico's | Invoer › 10 → Outlook, PowerPoint | Saldo per jaar, risico negatief. |
| Reserves / PCM | Invoer › 11 → Outlook | Reserves x % gereed (cumulatief) per jaar. |
| Gebieden | Invoer › 7 en 2 → RDFO (onderaan), PowerPoint | Som per gebied uit Projecten. |
| Strapo's | Invoer › 2 (Strapo Ja/Nee) → Outlook "waarvan strapo's" | — |
| Woningsegmentatie | Invoer › 2 (7 productsegmenten + % belegger) → Woningen | Controle meldt Verdeelfout. |
| Template | `RDFO_template.xlsm` (build met `template=True`) | Alle invoer, bronregels en Historie leeg. |

Bevestigd door de gebruiker: 50% afslag, 90% optimistisch, AK-deler 1,05. Open: komt DAEB-omzet uit de ZPMW DAEB-export (nu handmatig)?

## 7. Werkwijze voor wijzigingen

1. Wijzig `tools/build_rdfo.py` en/of `macros/*.bas`; teksten in `tools/docs_content.py`.
2. Voeg tests toe in `tools/test_rdfo.py`.
3. Draai `tools/release.py`; alle tests groen en geen nieuwe formulefouten.
4. Bekijk het resultaat (`tools/preview.py <xlsm> <map> <Blad>`).
5. Werk `README.md` en de Handleiding bij (Handleiding wordt gegenereerd uit `docs_content.py`).
6. Schrijf Nederlands voor de gebruiker: kort, concreet, geen jargon.

**Valkuilen**
- Gebruik openpyxl-formules met Engelse functienamen en `,` als scheidingsteken; nieuwere functies met prefix (`_xlfn.MAXIFS`, `_xlfn.NUMBERVALUE`).
- Projectlijst en stoplichtinvoer zijn rij-gekoppeld: niet sorteren of tussenvoegen.
- Nieuwe knop: tuple in `BUTTONS` + `Public Sub` in `modRDFO.bas`; test controleert het aantal macro's achter knoppen (`len(macros_in_buttons)`).
- Nieuwe verborgen bladen ook toevoegen aan `VERBORGEN` in `modRDFO.bas`.
- LibreOffice en Excel verschillen soms (grafiekweergave, `DisplayFormat`); test belangrijke dingen in Excel.
