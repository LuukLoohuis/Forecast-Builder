# Overdracht RDFO-prototype → Copilot

Dit document is bedoeld als startcontext voor een AI-assistent (Copilot) of ontwikkelaar die verder bouwt aan het RDFO-bestand. Lees het helemaal voordat je iets wijzigt.

---

## 1. Wat is dit?

Een Excel-werkmap (`RDFO_prototype.xlsm`) voor de kwartaalprognose van een bouw-/ontwikkelregio (Heijmans-huisstijl: geel `#F5B800`, donkerblauw `#14284B`). De werkmap:

- leest twee SAP-exports in: **ZPMW03** (financieel per WBS-element) en **Woningetalage** (woningverkopen per Stufe 3);
- koppelt die via een projectlijst (**Invoer › 2**, sleutel WBS-element = Stufe 3);
- rekent per project en per versie (Budget, Q1–Q4) omzet, BOR (bruto omzetresultaat), realisatie, to go, woningverkopen en orderintake;
- legt elke kwartaalprognose onveranderlijk vast in **Historie** (correcties = nieuwe revisie, nooit overschrijven);
- rapporteert in Dashboard, RDFO, Projecten, Projectdetail, Woningen, Outlook en Stoplicht;
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

**Zichtbare bladen (12):** Start, Dashboard, RDFO, Projecten, Projectdetail, Woningen, Outlook, Stoplicht, Controle, Invoer, Historie, Handleiding.

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

## 6. Volgende opdracht: scenarioanalyse met bandbreedte

**Wens gebruiker:** scenarioanalyse toevoegen, met een bandbreedte rond de prognose.

Hieronder een **voorstel**. Leg de open vragen (6.5) eerst voor aan de gebruiker.

### 6.1 Doel
Laten zien hoe zeker de prognose is: per jaar een **pessimistische – basis – optimistische** waarde voor omzet, BOR en woningverkopen (en eventueel orderintake), op regio- en projectniveau.

### 6.2 Invoer (nieuw: Invoer › 7 "Scenario's", gele cellen)
Per scenario (Pessimistisch / Optimistisch; Basis = huidige prognose) een paar eenvoudige knoppen:

| Parameter | Voorbeeld pessimistisch | Voorbeeld optimistisch |
|---|---|---|
| Woningverkopen (tempo, % van prognose) | 80% | 110% |
| Vertraging verkoop/start bouw (maanden) | 6 | 0 |
| Omzetrealisatie to go (% van prognose) | 90% | 105% |
| BOR-marge (procentpunt) | −1,0 | +0,5 |
| Alleen projecten met kans ≥ (Kans% uit ZPMW03) | 75% | 25% |

Plus per project een optionele afwijking (kolom in Invoer › 2 of aparte tabel: WBS, scenario, % omzet, % BOR, vertraging), zodat grote risicoprojecten apart te sturen zijn. Koppeling met het stoplicht ligt voor de hand: rood stoplicht → automatisch pessimistische factor (instelbaar).

### 6.3 Rekenregels (voorstel)
- Realisatie (actual) verandert nooit; alleen **to go** en toekomstige jaren worden geschaald.
- Vertraging schuift omzet/woningen naar volgende jaren (niet weg), zodat het totaal per project gelijk blijft; alleen de verdeling over jaren verandert.
- Bereken in een nieuw verborgen blad `SCEN_Project` (1 rij per project, zelfde rijvolgorde als `VERW_Project`) en tel op in `CALC_Versies` of een nieuw `CALC_Scenario`.
- Geen `ABS`, `n.b.` blijft `n.b.`, tekenafspraken als in §4.
- Scenario's worden **niet** vastgelegd in Historie (het zijn analyses), tenzij de gebruiker dat wil.

### 6.4 Uitvoer
- **Nieuw blad "Scenario's"** (zichtbaar, na Outlook): tabel per jaar met Pessimistisch / Basis / Optimistisch en budget; top-10 projecten met de grootste bandbreedte.
- **Grafiek bandbreedte** in Zebra-stijl: basis als kolom (gearceerd voor prognose), bandbreedte als foutbalk/whisker of gestapelde zwevende balk (min → max), budget als streep. Hulpfuncties: `ibcs_columns`, `hatched`, `solid` in `build_rdfo.py`.
- **Dashboard:** één tegel "Bandbreedte omzet FY: € x – y mln".
- **PowerPoint:** extra dia "Scenario's en bandbreedte" via een regel in `build_ppt_spec` (soort `tabel` + `grafiek`), tabel als blok in `build_dias` met `table_head(...)`; `preview_pptx.py` pakt het dan automatisch mee.
- **Controle:** check dat Basis = huidige prognose (aansluiting) en dat pessimistisch ≤ basis ≤ optimistisch.
- **Tests** in `tools/test_rdfo.py`: basis = upload, 100%/0-maanden-scenario = basis, vertraging verschuift maar behoudt totaal, realisatie ongewijzigd.

### 6.5 Open vragen voor de gebruiker
1. Welke kengetallen moeten een bandbreedte krijgen: omzet, BOR, woningverkopen, orderintake?
2. Twee scenario's (pessimistisch/optimistisch) of meer (bijv. ook "stress")?
3. Regio-brede parameters genoeg, of ook per project instelbaar?
4. Moet het stoplicht automatisch een project naar pessimistisch duwen?
5. Bandbreedte alleen voor het doeljaar, of voor alle Outlook-jaren?
6. Wil je scenario's kunnen vastleggen (bijv. "Q3 pessimistisch") voor vergelijking later?

---

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
