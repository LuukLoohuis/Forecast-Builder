Attribute VB_Name = "CashflowNaarPowerPoint"
Option Explicit

' =====================================================================================
'  Cashflow_scenario -> PowerPoint                       (late binding, Windows en Mac)
'  versie 10: typeblokken (maximaal MAX_TYPES woningtypes), knoppen Type invoegen / Type verwijderen
'  v11: scenariolijn als reeks in de cashflowgrafiek (dia 3); reeks weg als de kop op '(uit)' eindigt
'  v12: dia 5 bouwtermijnen-tijdlijn (BT_GRAFIEK), dia 6 verkoop en transport per kwartaal (VP1_/VP2_GRAFIEK),
'       dia 7 van verkoop naar omzet (VO_GRAFIEK)
'  v13: bouwtermijnen per woningtype: per termijn een baan per type (blok AA7, 108 kolommen, tot MAX_RIJEN_BT rijen,
'       aantal uit Model!BY96; tijd-as in kwartaalindex vast op Model!BY97..BY99; reeksen met kop '(uit)' gaan weg),
'       VP-blok vanaf EF7, VO-blok vanaf FG7
'  v14: dia 8 'Opbrengsten per woningtype' (PT_GRAFIEK, blok FM7, 8 kolommen): de macro maakt per woningtype met aantal > 0
'       een kopie van dia 8, zet het type in de keuzecel op het Dashboard (CEL_PT_KEUZE), rekent door en vult de kopie;
'       dia 8 zelf gaat daarna weg. Blok 5 op Woningtypes (soort DAEB en fees, rijen 48-63) bij Type invoegen/verwijderen.
'       Tabellen vanaf kolom FV; sjabloon v14
'  v14: UitFOOphalen (knop 'Ophalen uit FO' op tab Invoer): leest een FO-werkboek ('1. Cashflow', 'FO - actuals',
'       'CF - opbrengsten') en vult tab Invoer (cashflow, verkocht, transport), tab Woningtypes (typen, koopsom, grondtermijn,
'       start bouw, bouwtermijnen, extra's, DAEB-fees) en de FO-cellen op het Dashboard (zelfde logica als forecast_builder/fo.py)
'
'  NaarPowerPoint        opent het sjabloon als kopie, vult teksten, tabellen en grafieken
'                        vanaf tabblad "PowerPoint" en bewaart een nieuwe presentatie
'  ControleerSjabloon    kijkt of het sjabloon alle vormnamen heeft die de knop vult
'  KnoppenControleren    zet ontbrekende knoppen op de tabbladen (draait bij het openen)
'  TypeInvoegen          voegt een leeg typeblok in (drie kolommen op Woningtypes, twee op Invoer)
'  TypeVerwijderen       haalt een typeblok weg (zelfde kolommen); alles rechts schuift mee
'  SchikDia5             na het vullen van VT_TABEL (dia 5): bij meer dan 4 types rijen en letters
'                        kleiner, zo nodig de kaart met BT_GRAFIEK korter, en de tabelkaart sluit om de tabel
'  SchikDia2             na het vullen van de teksten (dia 2): kleine KPI-kaarten zonder regel op tabblad PowerPoint
'                        (of een lege lastenkaart) gaan weg, de overige verdelen de breedte (drie kaarten: precies onder
'                        de grote kaarten); de voetnoot en een leeg bijschrift 'vorige prognose' gaan weg
'  VulBalken             balken op dia 2: breedte = aandeel (kolom D) x breedte van het spoor, normstreep op de norm
'  SchikGrafiekDia2      kleine cashflowgrafiek op dia 2: waarde-as, grijs vlak t/m 'nu' en het nu-blok bij dat kwartaal
'
'  v15: grafiek per woningtype (dia 8): blok FM7 met 10 kolommen (totaal per kwartaal als label, eindpunt met het totaal),
'       alleen de periodes vanaf twee jaar vóór de eerste opbrengst (Model!BY128); 'Vorige prognose' op dia 3 lichtblauw (sinds v22: geel, scenario rood); sjabloon v15.
'       Knoppen op Woningtypes: 'Vorige prognose uit FO' (ouder FO -> blok 4 en 6) en 'Nu bewaren als vorige prognose'.
'  v22: dia 6 boven elke staaf de opsomming per type ('10 + 2', OpsommingLabels); percentages boven de lijn
'  v16: dia 6 (VP1_/VP2_GRAFIEK): cumulatief percentage verkocht/getransporteerd als lijn met labels; blok EF7 met 30 kolommen
'       (vier kolommen extra: cumulatief % en label, verkocht en transport), alleen de periodes vanaf het actuals-jaar (Model!BY132);
'       VO-blok vanaf FK7, blok per type (dia 8) vanaf FQ7, tabellen vanaf kolom GB; sjabloon v16.
'       Dia 2 met minder tekst: bandbreedte (KPI5) heeft geen regel meer op tabblad PowerPoint en SchikDia2 haalt die kaart
'       weg (werkt met sjabloon v15 en v16); KPI8 (was rente) = marge niet-DAEB / totaal, verkocht bij start bouw,
'       verkocht / getransporteerd t/m actuals; de voetnoot gaat van de dia (VOETNOOT, anders de tekstvorm die met 'Stip' begint).
'  v17: dia 2 in het nieuwe ontwerp (tools/pptx_v17_kerncijfers.py): drie tegels, een tijdlijnband met de kleine cashflowgrafiek
'       KPI_GRAFIEK (blok GK7, 8 kolommen: jaarlabel, cum. cashflow, vorige prognose, nu, dieptepunt, break-even, eindsaldo, €0;
'       de laatste zes als één-punts-reeksen met de reeksnaam als label; '(uit)' = reeks weg) en vier kaarten met balken:
'       verkocht bij start bouw (KPI6_SPOOR/_BALK/_NORM), verkocht / getransporteerd (KPI7_SPOOR_V/_BALK_V, _T), marge en de
'       grootste lasten (KPI9_NAAM1..3, _BEDRAG1..3, _SPOOR1..3/_BALK1..3; Dashboard-sectie 'GROOTSTE LASTEN'). VulBalken zet de
'       balkbreedtes (aandeel in kolom D), SchikGrafiekDia2 de waarde-as (Model!BY142..BY143), het grijze vlak GRAFIEK_REALISATIE
'       en het nu-blok (TL_NU_LABEL, KPI4_WAARDE, KPI4_SUB) op het 'nu'-kwartaal; instellingen op D76/D77; sjabloon v17.
'  v18: scenario's van tab Scenario's (start bouw, verkoop en transport per scenario): dia 4 'Verkoopscenario's' (sjabloon v18,
'       tools/pptx_v18_scenarios.py) met een kaart per scenario, legenda en VS_GRAFIEK (blok BLOK_VS, KOL_VS kolommen, tijdvak vanaf
'       het actuals-jaar zoals de verkoopgrafiek); de cashflowgrafiek (dia 3) krijgt de scenariolijnen erbij: blok BLOK_CF met KOL_CF
'       kolommen (tot v17: G7, 8 kolommen). SchikDiaScenarios haalt de dia weg als er geen scenario is ingevuld of Naar PowerPoint =
'       nee (Model!BY154 = 0), en anders de kaarten en legenda-regels van lege scenario's, en verdeelt de rest; SchikGrafiekVS zet de
'       waarde-as (Model!BY155..BY157) en het grijze vlak VS_REALISATIE t/m het 'nu'-kwartaal. Instellingen nu op D130/D131
'       (meer KPI-regels); de dia's na dia 3 schuiven een nummer op; sjabloon v18.
'  v19: tab CF kosten (kosten per post en per kwartaal): UitFOOphalen leest ook het FO-tabblad 'CF - kosten' (blok 'Onderdelen',
'       FOKosten; zelfde regels als fo.py _lees_kosten) en zet 'schuift mee met' op de naam van de post (StandaardMee), of houdt de
'       keuze die er voor een post met dezelfde naam al stond. De scenariolijnen in de cashflowgrafiek hebben een eigen schakelaar
'       op tab Scenario's (kop '(uit)' in blok BLOK_CF); de grootste lasten op dia 2 komen zonder Dashboard-invoer uit tab CF kosten
'       (tab PowerPoint, geen macrowijziging). Tab Woningtypes blok 7: bouwkosten per termijn (aannemer) uit het termijnblok met
'       dezelfde titel op 'CF - kosten' (KostenBlok, VoegBouwkostenToe; zoals fo.py _lees_bouwkosten). Meer dan tien termijnen na
'       het splitsen over kwartalen: kleinste delen samengevoegd (BinnenLimiet, zoals fo.py _binnen_limiet).
'  v20: DAEB-fees: ook een los DAEB-blok onderaan 'CF - opbrengsten' ('Opbrengsten DAEB', de componenten met hun totaal boven
'       de koprij) wordt een DAEB-type (FODaeb, LeesDaebBlok; aantal en verkoop van tabblad 'Verkooptempo', transport =
'       verkoop); fee-bedragen uit de cellen, per kwartaal gesplitst (FeeDelen, zoals fo.py _fee_delen); ook in
'       VorigePrognoseUitFO. Dia 'opbrengsten per woningtype' (VulDiasPerType): standaard één dia met het totaal van alle
'       niet-DAEB-typen (keuzecel 0 = 'Totaal (zonder DAEB)'); Dashboard CEL_PT_PP: 'totaal (zonder DAEB)', 'per type' of
'       'totaal en per type'. Geen sjabloonwijziging.
'  v21: bouwtermijnengrafiek zonder legenda (anders bleven kwartaalnummers als losse cijfers onder de grafiek staan); kortere teksten; vormen die geen tekst meer krijgen (ondertitels van dia 2, 3, 4 en 7, CF_VOET, KPI4_SUB, VS_Kk_DELTA,
'       VS_NOOT) haalt de macro van de dia (HaalVervallenVormenWeg), zodat ook sjabloon v18 werkt; sjabloon v19 (tools/
'       pptx_v19_kort.py) heeft ze niet meer en de scenariolijnen in contrastrijkere kleuren met een eigen lijnstijl.
'  v22: veel DAEB (meer dan de helft van de woningen): dia 7 met balken en % alleen voor koop en DAEB als gouden markering
'       (twee kolommen extra in het VP-blok, KOL_VP = 32; de blokken erna schuiven twee kolommen op; '(uit)' = geen markering,
'       ook in VP1_/VP2_GRAFIEK); drie kaartjes VP_STAT1..3 op dia 7; bouwtermijnen in twee groepen (koop / DAEB-fees) met een
'       kopregel en het bedrag per fee (alleen tab PowerPoint); sjabloon v20 (tools/pptx_v20_daeb.py).
'       Daarna, zoals de preview: VulTeksten zet de tekst van een grafiekvorm als grafiektitel, kleurt *_ACCENT-vormen (hex in
'       kolom D) en de blokjes/ruiten in een legenda (*_LEGENDA, BT_SUBTITEL; kleuren in kolom E, KleurLegenda); dikke gouden
'       DAEB-fee-balken met het bedrag ernaast (reeksen dfee, fgap, fl1..fl10 achteraan het BT-blok, KOL_BT = 120; de blokken erna
'       schuiven mee).
'       Daarna: projectnaam (Dashboard CEL_PROJECTNAAM, bij Ophalen uit FO uit de bestandsnaam als hij leeg is) op de titeldia
'       (COVER_PROJECT, COVER_KWARTAAL) en in de voettekst VOET_TEKST in de dia-indeling (ZoekVorm zoekt ook in de indelingen);
'       bouwtermijnen met drie koprijen (lege rij tussen jaar en kwartaal; MAX_RIJEN_BT = 133). Sjabloon v20 heeft geen tabel
'       woningtypes meer (VT_TABEL ontbreekt zonder melding; met een ouder sjabloon vult de knop hem zoals altijd).
'  v23: FO-werkbestand (FO en scenariomodel in een bestand): OphalenUitFO kan dit werkbestand zelf lezen (knop
'       'Scenario bijwerken' op tab Start, zonder vragen); FOKosten volgt per post de methode in kolom C van
'       CF - kosten (Voorstel = voorstelblok, VOORSTEL_RIJEN lager); het transport per type komt uit de regel
'       'koopsom' (die volgt de transportregel); 1. Cashflow heet nu 'Cashflow'; tab CF kosten heet 'Scenario kosten';
'       sjabloon v21 (v20 met vormnamen voor de datum op dia 1 en de kaarten van dia 10).
'       Sjabloon v22 (tools/pptx_v22.py): palet van het werkboek, tekst minimaal 8 pt, dia 3 zonder scenario 1-3 en met de
'       kans grondtermijn; 9 dia's (zonder 'Van verkoop naar omzet'); op de dia Update PEW het overzicht van 2. Ontw PEW (of 2B bij DAEB) als afbeelding in PEW_BEELD (geen top 3 meer).
'       Daarna: Invoer, Woningtypes, Scenario kosten en de FO-cellen op het Dashboard zijn formules op de FO-tabbladen
'       (tools/scenario_koppeling.py); OphalenUitFO blijft bestaan voor een los FO, maar heeft geen knop meer.
'  SchikDiaScenarios    dia 4: weg zonder scenario's; lege kaarten en legenda-regels weg, de rest verdeelt de breedte
'  SchikGrafiekVS       VS_GRAFIEK (dia 4): as en grijs vlak t/m 'nu'
'  UitFOOphalen         vult Invoer, Woningtypes en CF kosten vanuit een FO-werkboek (zie v14 en v19 hierboven)
'  VorigePrognoseUitFO  vult blok 4 en 6 (vorige prognose) vanuit een ouder FO
'  PlanningBewarenAlsVorig  zet de huidige planning in blok 4 en 6
'
'  Welke tekst naar welke vorm gaat staat op tabblad "PowerPoint", kolom F.
' =====================================================================================

Private Const SH_PP As String = "PowerPoint"
Private Const SH_DASH As String = "Dashboard"
Private Const SH_MODEL As String = "Model"
Private Const SH_TYPES As String = "Woningtypes"
Private Const SH_INVOER As String = "Invoer"
Private Const SH_KOSTEN As String = "Scenario kosten"  ' tab Scenario kosten (tot v23: CF kosten)
Private Const KOST_RIJ1 As Long = 8                    ' eerste kostenpost (layout.KF_R1)
Private Const VOORSTEL_RIJEN As Long = 41             ' werkbestand: voorstel per post op CF - kosten = rij post + 41
Private Const KOST_N As Long = 20                      ' aantal posten (layout.N_KF)
Private Const KOST_COL1 As Long = 9                    ' kolom I = periode 1 (layout.KF_COL1); B naam, C schuift mee met, E gerealiseerd
Private Const KPI_ROW1 As Long = 8            ' eerste regel met KPI-teksten
Private Const CEL_N As String = "BY338"            ' aantal getoonde periodes (tabblad Model; Dashboard: grafiek vanaf / t/m)
Private Const CEL_N_ALLE As String = "BY341"       ' kleine grafiek op dia 2: t/m het jaar van de laatste opbrengsten of kosten
Private Const CEL_VP_RIJEN As String = "BY339"     ' verkoop en transport: totaalrij + getoonde periodes
Private Const CEL_AS_MIN As String = "BY97"        ' ondergrens tijd-as BT_GRAFIEK: kwartaalindex 1 jan eerste jaar (Model, t_first)
Private Const CEL_AS_MAX As String = "BY99"        ' bovengrens tijd-as: kwartaalindex na het laatste jaar (Model, t_end)
Private Const CEL_SJABLOON As String = "D160"
Private Const CEL_NAAM As String = "D161"
Private Const SJABLOON_STANDAARD As String = "Kwartaal_Template_cashflow_v22.pptx"
Private Const SJABLOON_VORIG As String = "Kwartaal_Template_cashflow_v21.pptx"   ' als v22 er niet staat
Private Const BEWAAR_VRAAG As String = "De planning die nu in dit werkboek staat (start bouw per type en de kwartalen van de fee-termijnen) " & _
    "bewaren als vorige prognose (blok 4 en 6 op tab Woningtypes)?" & vbCrLf & vbCrLf & _
    "Ja = bewaren (lichtrood in de bouwtermijnengrafiek) · Nee = de vorige prognose laten zoals die is · Annuleren = niets ophalen"
Private Const MAX_RIJEN As Long = 60               ' periodes (kwartalen) per grafiekblok
Private Const MAX_RIJEN_BT As Long = 133           ' rijen in het bouwtermijnenblok (koprijen, banen type x termijn, scheidingsrijen)
Private Const KOL_BT As Long = 120                 ' kolommen in het bouwtermijnenblok (label + alle reeksen van layout.BT_REEKSEN; layout.PP_BT_KOL)
Private Const BLOK_PT As String = "GE7"            ' blok grafiek per woningtype (dia 8): kwartaal, acht reeksen, realisatie
Private Const KOL_PT As Long = 10
Private Const KOL_VP As Long = 32          ' kolommen van het blok van dia 6 (verkoop en transport, met cumulatief %)
Private Const CEL_PT_N As String = "BY128"
Private Const CEL_VP_N As String = "BY132"         ' aantal getoonde periodes in de verkoop- en transportgrafiek (vanaf het actuals-jaar)         ' aantal getoonde periodes in de grafiek per type (tijdvak vanaf twee jaar vóór de eerste opbrengst)
Private Const CEL_PT_KEUZE As String = "B120"      ' keuzecel op het Dashboard: naam van het type in de grafiek per woningtype
Private Const CEL_PT_PP As String = "W120"         ' Dashboard (sinds v20): dia per woningtype als totaal zonder DAEB, per type of beide
Private Const PT_TOTAAL As String = "Totaal (zonder DAEB)"   ' keuze in CEL_PT_KEUZE (of 0): alle niet-DAEB-typen samen
' dia 2 (sjabloon v17): blok van de kleine cashflowgrafiek, as-grenzen (Model, kc_min/kc_max), positie actuals, en het tekenvlak
' van KPI_GRAFIEK als fractie van het grafiekkader (layout.KC_PLOT_X / KC_PLOT_W; manualLayout in het sjabloon)
Private Const BLOK_KC As String = "GY7"
Private Const KOL_KC As Long = 8
Private Const CEL_KC_MIN As String = "BY142"
Private Const CEL_KC_MAX As String = "BY143"
Private Const CEL_POS_ACTUALS As String = "BY10"
Private Const KC_PLOT_X As Double = 0.0217
Private Const KC_PLOT_W As Double = 0.9217
' v18: cashflowgrafiek met de scenariolijnen (dia 3) en dia 4 'Verkoopscenario's'
Private Const BLOK_CF As String = "HH7"
Private Const KOL_CF As Long = 11
Private Const BLOK_VS As String = "HT7"
Private Const KOL_VS As Long = 9
Private Const N_VS As Long = 3                     ' scenario's naast de Basis (kaarten VS_K1..VS_K3, legenda VS_LEG1..3)
Private Const CEL_VS_PP As String = "BY154"        ' Model: dia 4 en de scenariolijnen in PowerPoint (1/0)
Private Const CEL_VS_MIN As String = "BY155"       ' Model: waarde-as van VS_GRAFIEK (€ mln) en de stap
Private Const CEL_VS_MAX As String = "BY156"
Private Const CEL_VS_STAP As String = "BY157"
Private Const CEL_VS_NU As String = "BY171"        ' Model: positie van het 'nu'-kwartaal in het tijdvak van VS_GRAFIEK
Private Const VS_PLOT_X As Double = 0.0437
Private Const VS_PLOT_W As Double = 0.9421
' FO-koppeling (Dashboard): bestand, actuals t/m, opbrengsten en kosten totaal uit '1. Cashflow', datum; actuals-knop W5/X5
Private Const CEL_FO_BESTAND As String = "W69"
Private Const CEL_PROJECTNAAM As String = "W10"    ' Dashboard: projectnaam (titeldia en voettekst van de dia's)
Private Const CEL_FO_ACTUALS As String = "W70"
Private Const CEL_FO_OPBR As String = "W71"
Private Const CEL_FO_KOSTEN As String = "W72"
Private Const CEL_FO_DATUM As String = "W73"
Private Const CEL_ACTUALS_JAAR As String = "W5"
Private Const CEL_ACTUALS_KW As String = "X5"
Private Const TYPE_RIJ_KOOPSOM As Long = 9
Private Const TYPE_RIJ_STARTJAAR As Long = 10
Private foStap As String                             ' laatste deelstap van 'Ophalen uit FO' (voor de foutmelding)
Private Const IDX_MIN As Long = 7961                 ' 1990*4+1: kleinste kwartaalindex die IsJaar toelaat
Private Const IDX_MAX As Long = 8404                 ' 2100*4+4
Private Const MAX_PERIODES As Long = 60
Private Const LIJST_MAX As Long = 200                ' Ophalen uit FO: regels per type vóór het samenvoegen (BinnenLimiet)
Private Const titel As String = "Naar PowerPoint"

' typeblokken (zelfde getallen als forecast_builder/layout.py)
Private Const MAX_TYPES As Long = 10
Private Const TYPE_COL1 As Long = 3           ' kolom C: eerste blok op tabblad Woningtypes
Private Const TYPE_BREEDTE As Long = 3        ' kolommen per type op Woningtypes (naam/%, %, kw)
Private Const TYPE_RIJ_NAAM As Long = 7
Private Const TYPE_RIJ_STARTKW As Long = 11
Private Const TYPE_RIJ_GROND As Long = 15
Private Const TYPE_RIJ_T1 As Long = 16
Private Const TYPE_RIJ_TN As Long = 25
Private Const TYPE_RIJ_X1 As Long = 33        ' extra opbrengsten per woning (blok 3 op tabblad Woningtypes)
Private Const TYPE_RIJ_XN As Long = 38
Private Const TYPE_RIJ_VJAAR As Long = 43     ' start bouw vorige prognose (blok 4): jaar / kwartaal
Private Const TYPE_RIJ_VKW As Long = 44
Private Const TYPE_RIJ_SOORT As Long = 48     ' blok 5: soort (niet-DAEB / DAEB)
Private Const TYPE_RIJ_FC1 As Long = 51       ' fee-componenten: naam (kolom 1), totaal (kolom 2)
Private Const TYPE_RIJ_FCN As Long = 52
Private Const TYPE_RIJ_F1 As Long = 54        ' fee-termijnen: mijlpaal (1), bedrag of % (2), kwartaal (3)
Private Const TYPE_RIJ_FN As Long = 63
Private Const TYPE_RIJ_VF1 As Long = 68       ' blok 6: kwartaal per fee-termijn volgens de vorige prognose (kolom 3)
Private Const TYPE_RIJ_VFN As Long = 77
Private Const TYPE_RIJ_B1 As Long = 82        ' blok 7 (sinds v19): bouwkosten per termijn (aannemer): termijn, bedrag €, bouwkwartaal
Private Const TYPE_RIJ_BN As Long = 91
Private Const TYPE_RIJ_AANTAL As Long = 8
Private Const INVOER_COL1 As Long = 12        ' kolom L: eerste paar (verkocht, transport) op tabblad Invoer
Private Const INVOER_BREEDTE As Long = 2
Private Const INVOER_RIJ1 As Long = 8
Private Const INVOER_RIJN As Long = 500        ' zelfde grens als de formules (layout.IN_ROWMAX)

Public OphalenGelukt As Boolean                    ' laatste OphalenUitFO gelukt (voor PowerPoint maken)

Private Function Grafieken() As Variant
    ' vormnaam op de dia, kopcel van het blok op tabblad PowerPoint, aantal kolommen,
    ' cel op tabblad Model met het aantal rijen (leeg = aantal periodes uit CEL_N)
    ' CF (dia 3): acht kolommen; de achtste (kolom N) is de scenariolijn, in het sjabloon als reeks op kolom H van het
    '   gegevensblad (staat de lijn uit, dan eindigt de kop op '(uit)' en haalt VulGrafiek de reeks weg)
    ' BT (dia 5): bouwtermijnen-tijdlijn: koprij jaartallen, koprij kwartaalnummers, dan per termijn een baan per type
    '   (categorie in kolom A, daarna de reeksen: stapel 1 = onzichtbaar, grijs, onzichtbaar, vorige prognose per type (10),
    '   grijs, onzichtbaar, jaarblokjes (16); stapel 2 = onzichtbaar, huidige planning per type (10), onzichtbaar,
    '   kwartaalblokjes (64)); tijd in kwartaalindex; aantal rijen in Model!BY96; reeksen met kop '(uit)' (typen zonder
    '   termijnen, jaren buiten het bereik) gaan weg; na het vullen zet ZetTijdAs beide waarde-assen op BY97..BY99
    ' VP1/VP2 (dia 6): verkocht en getransporteerd per kwartaal, beide panelen lezen hetzelfde blok van 26 kolommen
    ' KPI (dia 2): kleine cashflowgrafiek, blok BLOK_KC met KOL_KC kolommen (één-punts-reeksen voor nu, dieptepunt, break-even,
    '   eindsaldo en €0; na het vullen zet SchikGrafiekDia2 de as, het grijze vlak en het nu-blok)
    ' VS (dia 4, sinds v18): verkoopscenario's, blok BLOK_VS (jaartal, Basis, scenario 1..3, dieptepunten), tijdvak vanaf het
    '   actuals-jaar (aantal rijen CEL_VP_N); '(uit)' = scenario niet in PowerPoint; na het vullen zet SchikGrafiekVS as en grijs vlak.
    '   CF heeft sinds v18 elf kolommen (BLOK_CF): de laatste drie zijn de scenario's (in sjabloon v17 ongebruikt)
    Grafieken = Array(Array("CF_GRAFIEK", BLOK_CF, KOL_CF, ""), Array("SC_GRAFIEK", "P7", 10, ""), _
                      Array("BT_GRAFIEK", "AA7", KOL_BT, "BY96"), _
                      Array("VP1_GRAFIEK", "ER7", KOL_VP, CEL_VP_RIJEN), Array("VP2_GRAFIEK", "ER7", KOL_VP, CEL_VP_RIJEN), _
                      Array("KPI_GRAFIEK", BLOK_KC, KOL_KC, CEL_N_ALLE), _
                      Array("VS_GRAFIEK", BLOK_VS, KOL_VS, CEL_VP_N))
End Function

Private Function Tabellen() As Variant
    ' vormnaam op de dia, kopcel van de tabel, aantal rijen met kop, aantal kolommen, lege regels overslaan (1 = ja)
    Tabellen = Array(Array("SC_TABEL", "GP7", 9, 4, 0), Array("VT_TABEL", "GP18", MAX_TYPES + 1, 8, 1))
End Function

' -------------------------------------------------------------------------------------
'  1. Knoppen
' -------------------------------------------------------------------------------------
Public Sub KnoppenControleren()
    ' de knop 'Naar PowerPoint' op het Dashboard rekent eerst door (PowerPointMaken), zoals de knop op Start;
    ' op de verborgen rekenbladen komen geen knoppen meer (Nieuw kwartaal starten bewaart ook de planning)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    If Not BestaatVorm(ws, "btnNaarPowerPoint") Then
        MaakKnop ws, ws.Range("W35:X36"), "btnNaarPowerPoint", "Naar PowerPoint", "PowerPointMaken", RGB(23, 54, 93), 0, 0, 2
    End If
    ws.Shapes("btnNaarPowerPoint").OnAction = "PowerPointMaken"
    On Error GoTo 0
End Sub

Private Function BestaatVorm(ws As Worksheet, naam As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(naam)
    BestaatVorm = Not shp Is Nothing
    On Error GoTo 0
End Function

Private Sub MaakKnop(ws As Worksheet, plek As Range, naam As String, tekst As String, macro As String, kleur As Long, breedte As Double, links As Double, plaatsing As Long)
    Dim shp As Shape, w As Double, hg As Double
    w = plek.Width - 4
    If breedte > 0 Then w = breedte
    hg = plek.Height - 4
    If hg < 18 Then hg = 18
    Set shp = ws.Shapes.AddShape(5, plek.Left + 2 + links, plek.Top + 2, w, hg)     ' 5 = afgeronde rechthoek
    With shp
        .Name = naam
        .Fill.ForeColor.RGB = kleur
        .Line.Visible = 0
        .TextFrame2.TextRange.Text = tekst
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = -1
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2      ' gecentreerd
        .TextFrame2.VerticalAnchor = 3                           ' midden
        .TextFrame2.WordWrap = 0
        .Placement = plaatsing                                   ' 2 = meebewegen met de cellen, 3 = los
        .OnAction = macro
    End With
End Sub

' -------------------------------------------------------------------------------------
'  2. Woningtypes: blok invoegen of verwijderen
'     Het model leest de blokken op positie, dus na invoegen of verwijderen klopt alles weer.
' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------
'  3. Naar PowerPoint
' -------------------------------------------------------------------------------------
Public Sub NaarPowerPoint()
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim stap As String, n As Long, nG As Long, pad As String, mist As String, fouten As String
    Dim g As Variant, t As Variant, i As Long, bericht As String
    On Error GoTo Fout

    stap = "werkbladen zoeken"
    Set wsP = ThisWorkbook.Worksheets(SH_PP)
    Application.Calculate
    n = AantalPeriodes()
    If n < 1 Then
        MsgBox "Er staan nog geen periodes op tabblad Invoer.", vbExclamation, titel
        Exit Sub
    End If

    stap = "PowerPoint starten"
    Set pp = HaalPowerPoint()
    If pp Is Nothing Then
        MsgBox "PowerPoint kon niet worden gestart.", vbExclamation, titel
        Exit Sub
    End If

    stap = "sjabloon openen"
    Set pres = OpenSjabloon(pp, wsP)
    If pres Is Nothing Then Exit Sub

    stap = "teksten vullen"
    HaalVervallenVormenWeg pres
    VulTeksten pres, wsP, mist

    stap = "dia 2 schikken"
    SchikDia2 pres, wsP
    VulBalken pres, wsP

    stap = "dia verkoopscenario's schikken"
    SchikDiaScenarios pres, wsP

    stap = "tabellen vullen"
    t = Tabellen()
    For i = LBound(t) To UBound(t)
        VulTabel pres, CStr(t(i)(0)), wsP.Range(CStr(t(i)(1))), CLng(t(i)(2)), CLng(t(i)(3)), (CLng(t(i)(4)) = 1), mist, fouten
    Next i

    stap = "grafieken vullen"
    g = Grafieken()
    For i = LBound(g) To UBound(g)
        ' aantal rijen: uit de opgegeven Model-cel (bouwtermijnenblok), anders het aantal periodes;
        ' geen rijen (bijvoorbeeld geen bouwtermijnen): een lege rij, zodat de voorbeeldgegevens uit het sjabloon verdwijnen
        If Len(CStr(g(i)(3))) > 0 Then
            nG = AantalRijen(CStr(g(i)(3)), MAX_RIJEN_BT)
            If nG < 1 Then nG = 1
        Else
            nG = n
        End If
        ' de dia met VS_GRAFIEK is weg als er geen scenario in PowerPoint komt (SchikDiaScenarios)
        If CStr(g(i)(0)) <> "VS_GRAFIEK" Or VsDiaAan() Then
            VulGrafiek pres, CStr(g(i)(0)), wsP.Range(CStr(g(i)(1))), nG, CLng(g(i)(2)), mist, fouten
        End If
    Next i

    stap = "dia per woningtype"
    VulDiasPerType pres, wsP, n, mist, fouten

    stap = "2. Ontw PEW op de dia Update PEW"
    PlaatsPEW pres, fouten

    stap = "presentatie bewaren"
    pad = Bewaar(pres, wsP)

    On Error Resume Next
    pp.Activate
    pres.Windows(1).Activate
    On Error GoTo Fout

    If Len(pad) > 0 Then
        bericht = "Klaar. De presentatie is gemaakt en bewaard als:" & vbCrLf & vbCrLf & pad
    Else
        bericht = "Klaar. De presentatie staat open in PowerPoint, maar is nog niet bewaard."
    End If
    If Len(mist) > 0 Then bericht = bericht & vbCrLf & vbCrLf & "Niet gevonden in het sjabloon (overgeslagen):" & vbCrLf & mist
    If Len(fouten) > 0 Then
        bericht = bericht & vbCrLf & vbCrLf & "Niet gelukt (plak deze met de hand vanaf tabblad PowerPoint):" & vbCrLf & fouten
    End If
    MsgBox bericht, IIf(Len(mist) + Len(fouten) > 0, vbExclamation, vbInformation), titel
    Exit Sub

Fout:
    Application.ScreenUpdating = True
    MsgBox "Het ging mis bij: " & stap & vbCrLf & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, titel
End Sub

Private Sub PlaatsPEW(pres As Object, ByRef fouten As String)
    ' het overzicht van 2. Ontw PEW als afbeelding in het vak PEW_BEELD op de dia Update PEW (zoals het op het tabblad staat,
    ' met de toelichting); bij een gemengd project (DAEB-woningen, en A of B ingevuld op 2B. Ontw PEW) het overzicht van
    ' 2B. Ontw PEW.
    Dim vak As Object, dia As Object, sr As Object, pic As Object, tb As Object, bron As Range, wsB As Worksheet
    Dim marge As Single, hT3 As Single, tekst As String, nieuw As Boolean
    Set vak = ZoekVorm(pres, "PEW_BEELD")
    nieuw = Not vak Is Nothing
    If vak Is Nothing Then Set vak = ZoekVorm(pres, "PEW_TABEL")            ' sjabloon v21
    If vak Is Nothing Then Exit Sub
    Set dia = vak.Parent
    On Error GoTo Mis
    Set bron = ThisWorkbook.Worksheets("2. Ontw PEW").Range("C4:K31")
    On Error Resume Next
    Set wsB = ThisWorkbook.Worksheets("2B. Ontw PEW")
    On Error GoTo Mis
    If Not wsB Is Nothing Then
        If Val(CStr(ThisWorkbook.Worksheets(SH_MODEL).Range("BY135").Value)) > 0 And _
           Application.WorksheetFunction.Count(wsB.Range("F7:G26")) > 0 Then Set bron = wsB.Range("C4:N31")
    End If
    marge = IIf(nieuw, 0, 6)
    hT3 = 0
    ' als afbeelding (zoals het op het tabblad staat: kleuren, band, toelichting); een tabel komt in PowerPoint niet mooi over
    bron.CopyPicture Appearance:=1, Format:=-4147
    DoEvents
    Set sr = dia.Shapes.PasteSpecial(DataType:=2)                  ' als afbeelding (verbeterde metafile)
    Set pic = sr(1)
    pic.LockAspectRatio = True
    pic.Width = vak.Width - 2 * marge
    If pic.Height > vak.Height - hT3 - 2 * marge Then pic.Height = vak.Height - hT3 - 2 * marge
    pic.Left = vak.Left + (vak.Width - pic.Width) / 2
    pic.Top = vak.Top + marge
    pic.Name = "PEW_AFBEELDING"
    ' sinds deze versie geen top 3 kosten meer op de dia (staat op Start, blok 7); een oud vak PEW_TOP3 leegmaken
    Set tb = ZoekVorm(pres, "PEW_TOP3")
    If Not tb Is Nothing Then tb.Delete
    Exit Sub
Mis:
    fouten = fouten & "- 2. Ontw PEW op de dia Update PEW: " & Err.Description & vbCrLf
End Sub

Private Function Top3Tekst() As String
    ' top 3 kosten van het reviewkwartaal (tab Start, blok 7), een regel per kostensoort; leeg zonder top 3
    Dim wsS As Worksheet, r As Long, k As Long, v As Variant, tekst As String
    Set wsS = ThisWorkbook.Worksheets("Start")
    r = CLng(Val(FOWerkbestand.WerkbestandCfg("top3_rij", "0")))
    If r <= 0 Then Exit Function
    tekst = "Top 3 kosten " & CStr(wsS.Range("E12").Value) & " (actuals)"
    For k = 0 To 2
        v = wsS.Cells(r + k, 4).Value
        If Len(CStr(wsS.Cells(r + k, 3).Value)) > 0 And IsNumeric(v) Then
            tekst = tekst & vbCr & (k + 1) & ". " & CStr(wsS.Cells(r + k, 3).Value) & ":  " & ChrW(8364) & " " & Format(v, "#,##0")
        End If
    Next k
    Top3Tekst = tekst
End Function

Private Function AantalPeriodes() As Long
    AantalPeriodes = AantalRijen(CEL_N, MAX_RIJEN)
End Function

Private Function AantalRijen(cel As String, maxN As Long) As Long
    ' getal uit een cel op tabblad Model, begrensd op maxN; 0 bij leeg, fout of geen getal
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_MODEL).Range(cel).Value
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then AantalRijen = CLng(v)
    If AantalRijen < 0 Then AantalRijen = 0
    If AantalRijen > maxN Then AantalRijen = maxN
End Function

Private Function HaalPowerPoint() As Object
    Dim pp As Object
    On Error Resume Next
    Set pp = GetObject(, "PowerPoint.Application")
    If pp Is Nothing Then Set pp = CreateObject("PowerPoint.Application")
    If Not pp Is Nothing Then pp.Visible = True
    On Error GoTo 0
    Set HaalPowerPoint = pp
End Function

Private Function MapVanWerkmap() As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function
    If LCase$(Left$(p, 4)) = "http" Then
        MapVanWerkmap = p & "/"
    Else
        MapVanWerkmap = p & Application.PathSeparator
    End If
End Function

Private Function OpenSjabloon(pp As Object, wsP As Worksheet) As Object
    ' opent het sjabloon als naamloze kopie: het sjabloon zelf verandert nooit
    Dim pad As String, keuze As Variant, pres As Object
    pad = Trim$(CelTekst(wsP.Range(CEL_SJABLOON)))
    If Len(pad) = 0 Then pad = SJABLOON_STANDAARD
    If InStr(pad, "\") = 0 And InStr(pad, "/") = 0 And InStr(pad, ":") = 0 Then pad = MapVanWerkmap() & pad
    On Error Resume Next
    Set pres = pp.Presentations.Open(pad, 0, -1, -1)
    If pres Is Nothing And Len(Trim$(CelTekst(wsP.Range(CEL_SJABLOON)))) = 0 Then
        Set pres = pp.Presentations.Open(MapVanWerkmap() & SJABLOON_VORIG, 0, -1, -1)
    End If
    On Error GoTo 0
    If pres Is Nothing Then
        MsgBox "Het sjabloon is hier niet gevonden:" & vbCrLf & vbCrLf & pad & vbCrLf & vbCrLf & _
               "Kies zo het sjabloon. Tip: zet het in dezelfde map als dit bestand," & vbCrLf & _
               "of vul het pad in op tabblad PowerPoint (cel " & CEL_SJABLOON & ").", vbInformation, titel
#If Mac Then
        keuze = Application.GetOpenFilename()
#Else
        keuze = Application.GetOpenFilename("PowerPoint (*.pptx;*.potx),*.pptx;*.potx", , "Kies het sjabloon")
#End If
        If VarType(keuze) = vbBoolean Then Exit Function
        On Error Resume Next
        Set pres = pp.Presentations.Open(CStr(keuze), 0, -1, -1)
        On Error GoTo 0
        If pres Is Nothing Then MsgBox "Dit bestand kon niet worden geopend.", vbExclamation, titel
    End If
    Set OpenSjabloon = pres
End Function

Private Function Bewaar(pres As Object, wsP As Worksheet) As String
    ' bewaart naast dit bestand; lukt dat niet (bijvoorbeeld op SharePoint), dan vraagt de knop om een plek
    Dim naam As String, pad As String, keuze As Variant
    naam = SchoneNaam(CelTekst(wsP.Range(CEL_NAAM)))
    If Len(naam) = 0 Then naam = "Cashflow update"
    naam = naam & " " & Format$(Now, "yyyy-mm-dd hhnn") & ".pptx"
    pad = MapVanWerkmap() & naam
    On Error Resume Next
    pres.SaveAs pad, 24                                          ' 24 = pptx
    If Err.Number <> 0 Then
        Err.Clear
        pad = ""
#If Mac Then
        keuze = Application.GetSaveAsFilename(naam)
#Else
        keuze = Application.GetSaveAsFilename(naam, "PowerPoint (*.pptx),*.pptx", , "Presentatie bewaren")
#End If
        If VarType(keuze) <> vbBoolean Then
            pres.SaveAs CStr(keuze), 24
            If Err.Number = 0 Then pad = CStr(keuze)
            Err.Clear
        End If
    End If
    On Error GoTo 0
    Bewaar = pad
End Function

Private Function SchoneNaam(s As String) As String
    Dim i As Long, teken As String, uit As String
    For i = 1 To Len(s)
        teken = Mid$(s, i, 1)
        If InStr("\/:*?""<>|", teken) = 0 Then uit = uit & teken
    Next i
    SchoneNaam = Trim$(uit)
End Function

' -------------------------------------------------------------------------------------
'  4. Vullen
' -------------------------------------------------------------------------------------
Private Sub HaalVervallenVormenWeg(pres As Object)
    ' vormen uit oudere sjablonen die geen tekst meer krijgen (sinds v21): weg, anders blijft de voorbeeldtekst staan
    Dim namen As Variant, i As Long, j As Long, shp As Object
    namen = Array("KPI_SUBTITEL", "KPI4_SUB", "CF_SUBTITEL", "CF_VOET", "VS_SUBTITEL", "VS_NOOT", "VP_SUBTITEL", _
                  "VS_K0_DELTA", "VS_K1_DELTA", "VS_K2_DELTA", "VS_K3_DELTA")
    On Error Resume Next
    For i = LBound(namen) To UBound(namen)
        For j = 1 To 3                                   ' ook een dubbele (gekopieerde) vorm; nooit eindeloos
            Set shp = Nothing
            Set shp = ZoekVorm(pres, CStr(namen(i)))
            If shp Is Nothing Then Exit For
            shp.Delete
        Next j
    Next i
End Sub

Private Function ZoekVorm(pres As Object, naam As String) As Object
    ' eerst op de dia's, dan in de dia-indelingen (daar staat de voettekst VOET_TEKST van alle dia's)
    Dim s As Object, shp As Object, d As Object, lay As Object
    For Each s In pres.Slides
        For Each shp In s.Shapes
            If shp.Name = naam Then
                Set ZoekVorm = shp
                Exit Function
            End If
        Next shp
    Next s
    On Error Resume Next
    For Each d In pres.Designs
        For Each lay In d.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                If shp.Name = naam Then
                    Set ZoekVorm = shp
                    Exit Function
                End If
            Next shp
        Next lay
    Next d
End Function

Private Function CelTekst(cel As Range) As String
    Dim v As Variant
    v = cel.Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency
            CelTekst = Format$(v, "0")
        Case Else
            CelTekst = CStr(v)
    End Select
End Function

Private Function HexKleur(hexTekst As String) As Long
    ' 'EDA100' -> RGB; ongeldig: grijs
    Dim t As String
    t = Trim$(hexTekst)
    If Len(t) <> 6 Then
        HexKleur = RGB(122, 122, 122)
        Exit Function
    End If
    On Error Resume Next
    HexKleur = RGB(CLng("&H" & Mid$(t, 1, 2)), CLng("&H" & Mid$(t, 3, 2)), CLng("&H" & Mid$(t, 5, 2)))
End Function

Private Sub KleurLegenda(shp As Object, kleuren As String)
    ' kleurt elk blokje (ChrW 9632) en elke ruit (ChrW 9670) met de volgende kleur uit de lijst; de rest van de tekst blijft zoals hij is
    Dim lijst As Variant, t As String, i As Long, k As Long, c As String
    On Error Resume Next
    lijst = Split(kleuren, ",")
    t = shp.TextFrame.TextRange.Text
    k = LBound(lijst)
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c = ChrW(9632) Or c = ChrW(9670) Then
            If k <= UBound(lijst) Then shp.TextFrame.TextRange.Characters(i, 1).Font.Color.RGB = HexKleur(CStr(lijst(k)))
            k = k + 1
        End If
    Next i
End Sub

Private Function RichtingKleur(richting As String) As Long
    Select Case richting
        Case "daeb"
            RichtingKleur = RGB(122, 83, 0)
        Case "slechter", "onder norm"
            RichtingKleur = RGB(185, 28, 28)
        Case "beter", "op norm"
            RichtingKleur = RGB(21, 128, 61)
        Case Else
            RichtingKleur = RGB(122, 122, 122)
    End Select
End Function

Private Sub VulTeksten(pres As Object, wsP As Worksheet, ByRef mist As String)
    ' loopt de KPI-regels af tot de eerste lege regel; kolom D = tekst, E = richting (kleur), F = vormnaam
    ' (balkregels, vormnaam met _BALK of op _NORM: geen tekst maar een aandeel; die doet VulBalken na SchikDia2)
    Dim r As Long, vorm As String, richting As String, shp As Object, stip As Object, kleur As Long
    r = KPI_ROW1
    Do While Len(CelTekst(wsP.Range("C" & r))) > 0 And r < KPI_ROW1 + 300
        vorm = Trim$(CelTekst(wsP.Range("F" & r)))
        If IsBalkregel(vorm) Then
            ' niets: balk
        ElseIf Left$(vorm, 3) = "VS_" And Not VsDiaAan() Then
            ' niets: dia 4 (verkoopscenario's) gaat weg
        ElseIf Len(vorm) > 0 Then
            Set shp = ZoekVorm(pres, vorm)
            If shp Is Nothing Then
                mist = mist & vorm & vbCrLf
            ElseIf shp.HasChart Then
                ' grafiek: de tekst wordt de grafiektitel (dia 7: 'Verkocht per kwartaal · koop (36)')
                On Error Resume Next
                shp.Chart.HasTitle = True
                shp.Chart.ChartTitle.Text = CelTekst(wsP.Range("D" & r))
                On Error GoTo 0
            ElseIf Right$(vorm, 7) = "_ACCENT" Then
                ' accentstreep: kolom D = kleur als hex (goud bij DAEB)
                On Error Resume Next
                shp.Fill.ForeColor.RGB = HexKleur(CelTekst(wsP.Range("D" & r)))
                On Error GoTo 0
            ElseIf Right$(vorm, 8) = "_LEGENDA" Or vorm = "BT_SUBTITEL" Then
                ' legenda in één tekstvak: kolom E = kleuren van de blokjes en ruiten (ChrW 9632 / 9670) in volgorde (hex, komma's)
                shp.TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & r))
                KleurLegenda shp, CelTekst(wsP.Range("E" & r))
            Else
                shp.TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & r))
                richting = LCase$(Trim$(CelTekst(wsP.Range("E" & r))))
                If Len(richting) > 0 Then
                    kleur = RichtingKleur(richting)
                    shp.TextFrame.TextRange.Font.Color.RGB = kleur
                    If InStr(vorm, "_VERSCHIL") > 0 Then
                        Set stip = ZoekVorm(pres, Replace(vorm, "_VERSCHIL", "_STIP"))
                        If Not stip Is Nothing Then stip.Fill.ForeColor.RGB = kleur
                    End If
                End If
            End If
        End If
        r = r + 1
    Loop
End Sub

Private Function RegelVanVorm(wsP As Worksheet, naam As String) As Long
    ' rij van de KPI-regel op tabblad PowerPoint met deze vorm in kolom F; 0 = geen regel (dan vult de knop de vorm niet)
    Dim r As Long
    r = KPI_ROW1
    Do While Len(CelTekst(wsP.Range("C" & r))) > 0 And r < KPI_ROW1 + 300
        If StrComp(Trim$(CelTekst(wsP.Range("F" & r))), naam, vbTextCompare) = 0 Then
            RegelVanVorm = r
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Private Function VormOpDia(dia As Object, naam As String) As Object
    Dim shp As Object
    For Each shp In dia.Shapes
        If shp.Name = naam Then
            Set VormOpDia = shp
            Exit Function
        End If
    Next shp
End Function

Private Function IsBalkregel(vorm As String) As Boolean
    ' regel op tabblad PowerPoint voor een balk (KPI6_BALK, KPI7_BALK_V, KPI9_BALK1, ...) of de normstreep (KPI6_NORM)
    IsBalkregel = (InStr(vorm, "_BALK") > 0 Or Right$(vorm, 5) = "_NORM")
End Function

Private Function IsTekstVanKaart(naam As String) As Boolean
    ' tekstvakken van een kleine kaart die bij drie kaarten op de tekstrand van de grote kaarten komen te staan
    IsTekstVanKaart = (Right$(naam, 6) = "_LABEL" Or Right$(naam, 7) = "_WAARDE" Or Right$(naam, 4) = "_SUB" _
                       Or InStr(naam, "_NAAM") > 0 Or InStr(naam, "_BEDRAG") > 0)
End Function

Private Sub SchikDia2(pres As Object, wsP As Worksheet)
    ' na het vullen van de teksten, dia 2 (kerncijfers): een kleine KPI-kaart in de onderste rij (KPI4 en verder: KAART, LABEL,
    ' WAARDE, SUB; de lastenkaart KPI9: NAAM1..3, BEDRAG1..3, SPOOR1..3, BALK1..3) gaat van de dia als de WAARDE geen regel heeft
    ' op tabblad PowerPoint (bandbreedte KPI5 in sjabloon v15/v16) of als de eerste lastenregel (KPI9_NAAM1) leeg is. De overige
    ' kaarten verdelen daarna de breedte van het kaartenblok met dezelfde tussenruimte; blijven er drie over, dan staan ze precies
    ' onder de grote kaarten KPI1..KPI3, met de teksten op dezelfde afstand van de rand als in de grote kaarten. Sporen en de
    ' kleine labels schuiven mee en rekken met de kaart; balken en normstreep schuiven alleen mee (VulBalken zet ze daarna
    ' opnieuw). Vult de knop alle kaarten, dan blijft de dia zoals het sjabloon hem heeft. Weg gaan ook: de voetnoot (VOETNOOT,
    ' of de tekstvorm die met 'Stip' begint) en een leeg bijschrift 'vorige prognose' bij de grafiek (GRAFIEK_LEGENDA_VORIG + _LIJN).
    ' Alleen kaarten in de onderste rij tellen mee (zelfde Top als de laagste kaart, 7 pt speling).
    ' Hoogtes en letters veranderen niet; fouten worden stil overgeslagen (de dia blijft dan zoals het sjabloon hem had)
    Const MAX_KAART As Long = 20
    Dim dia As Object, kaart As Object, shp As Object, groot As Object, grootTekst As Object
    Dim nr(1 To MAX_KAART) As Long, links(1 To MAX_KAART) As Double, breed(1 To MAX_KAART) As Double, boven(1 To MAX_KAART) As Double
    Dim blijft(1 To MAX_KAART) As Boolean, doelL(1 To MAX_KAART) As Double, doelB(1 To MAX_KAART) As Double
    Dim n As Long, nBlijft As Long, k As Long, i As Long, j As Long, t As Long, hNr As Long, hBlijft As Boolean, rW As Long
    Dim hL As Double, hB As Double, bLinks As Double, bRechts As Double, tussen As Double, w As Double, rijTop As Double
    Dim binnen As Double, mL As Double, mr As Double, prefix As String, onderGroot As Boolean

    On Error Resume Next
    Set kaart = ZoekVorm(pres, "KPI1_KAART")
    If kaart Is Nothing Then Exit Sub
    Set dia = kaart.Parent
    If dia Is Nothing Then Exit Sub

    ' voetnoot weg: de vorm VOETNOOT, of anders de tekstvorm die met 'Stip' begint (de tekst uit het sjabloon)
    Set shp = VormOpDia(dia, "VOETNOOT")
    If shp Is Nothing Then
        For j = dia.Shapes.Count To 1 Step -1
            If dia.Shapes(j).HasTextFrame Then
                If Left$(Trim$(dia.Shapes(j).TextFrame.TextRange.Text), 4) = "Stip" Then
                    Set shp = dia.Shapes(j)
                    Exit For
                End If
            End If
        Next j
    End If
    If Not shp Is Nothing Then shp.Delete

    ' bijschrift 'vorige prognose' bij de grafiek: weg als de tekst leeg is (geen vorige prognose)
    Set shp = Nothing
    Set shp = VormOpDia(dia, "GRAFIEK_LEGENDA_VORIG")
    If Not shp Is Nothing Then
        If Len(Trim$(shp.TextFrame.TextRange.Text)) = 0 Then
            shp.Delete
            Set shp = Nothing
            Set shp = VormOpDia(dia, "GRAFIEK_LEGENDA_VORIG_LIJN")
            If Not shp Is Nothing Then shp.Delete
        End If
    End If

    ' de kleine kaarten op de dia (alleen de onderste rij) en of de knop ze vult
    rijTop = -1
    For k = 4 To MAX_KAART
        Set kaart = Nothing
        Set kaart = VormOpDia(dia, "KPI" & k & "_KAART")
        If Not kaart Is Nothing Then
            If kaart.Top > rijTop Then rijTop = kaart.Top
        End If
    Next k
    For k = 4 To MAX_KAART
        Set kaart = Nothing
        Set kaart = VormOpDia(dia, "KPI" & k & "_KAART")
        If Not kaart Is Nothing Then
            If kaart.Top >= rijTop - 7 Then
                n = n + 1
                nr(n) = k
                links(n) = kaart.Left
                breed(n) = kaart.Width
                boven(n) = kaart.Top
                rW = RegelVanVorm(wsP, "KPI" & k & "_WAARDE")
                If rW > 0 Then
                    blijft(n) = True
                Else
                    rW = RegelVanVorm(wsP, "KPI" & k & "_NAAM1")
                    blijft(n) = False
                    If rW > 0 Then blijft(n) = (Len(CelTekst(wsP.Range("D" & rW))) > 0)
                End If
                If blijft(n) Then nBlijft = nBlijft + 1
            End If
        End If
    Next k
    If n = 0 Or nBlijft = n Then Exit Sub

    ' van links naar rechts (invoegsorteren op Left)
    For i = 2 To n
        hNr = nr(i): hL = links(i): hB = breed(i): hBlijft = blijft(i)
        j = i - 1
        Do While j >= 1
            If links(j) <= hL Then Exit Do
            nr(j + 1) = nr(j): links(j + 1) = links(j): breed(j + 1) = breed(j): blijft(j + 1) = blijft(j)
            j = j - 1
        Loop
        nr(j + 1) = hNr: links(j + 1) = hL: breed(j + 1) = hB: blijft(j + 1) = hBlijft
    Next i

    ' kaartenblok: van de linkerrand van de eerste tot de rechterrand van de laatste kaart; tussenruimte = de kleinste
    bLinks = links(1)
    bRechts = links(1) + breed(1)
    tussen = -1
    For i = 1 To n
        If links(i) + breed(i) > bRechts Then bRechts = links(i) + breed(i)
        If i > 1 Then
            If tussen < 0 Or links(i) - (links(i - 1) + breed(i - 1)) < tussen Then tussen = links(i) - (links(i - 1) + breed(i - 1))
        End If
    Next i
    If tussen < 0 Then tussen = 0

    ' kaarten zonder regel: alle vormen waarvan de naam met KPIk_ begint gaan weg
    For i = 1 To n
        If Not blijft(i) Then
            prefix = "KPI" & nr(i) & "_"
            For j = dia.Shapes.Count To 1 Step -1
                Set shp = dia.Shapes(j)
                If Left$(shp.Name, Len(prefix)) = prefix Then shp.Delete
            Next j
        End If
    Next i
    If nBlijft = 0 Then Exit Sub

    ' nieuwe plek: drie kaarten onder de drie grote kaarten (van links naar rechts), anders gelijk verdeeld over het blok
    binnen = -1
    onderGroot = (nBlijft = 3)
    If onderGroot Then
        For t = 1 To 3
            Set groot = Nothing
            Set groot = VormOpDia(dia, "KPI" & t & "_KAART")
            If groot Is Nothing Then
                onderGroot = False
            Else
                doelL(t) = groot.Left
                doelB(t) = groot.Width
            End If
        Next t
    End If
    If onderGroot Then
        For i = 1 To 2
            For j = 1 To 3 - i
                If doelL(j) > doelL(j + 1) Then
                    hL = doelL(j): doelL(j) = doelL(j + 1): doelL(j + 1) = hL
                    hB = doelB(j): doelB(j) = doelB(j + 1): doelB(j + 1) = hB
                End If
            Next j
        Next i
        ' tekstrand van de grote kaart: van de kaartrand tot de tekst (vak plus binnenmarge van het vak)
        Set groot = VormOpDia(dia, "KPI1_KAART")
        Set grootTekst = VormOpDia(dia, "KPI1_LABEL")
        If Not grootTekst Is Nothing Then binnen = grootTekst.Left + grootTekst.TextFrame.MarginLeft - groot.Left
    Else
        w = (bRechts - bLinks - (nBlijft - 1) * tussen) / nBlijft
        For t = 1 To nBlijft
            doelL(t) = bLinks + (t - 1) * (w + tussen)
            doelB(t) = w
        Next t
    End If

    ' verplaatsen en verbreden: elke vorm van de kaart houdt zijn afstand tot de linker- en rechterrand van de kaart (en rekt dus
    ' mee); onder de grote kaarten staan label, waarde en toelichting op dezelfde afstand van de rand als de tekst in de grote
    ' kaart (de binnenmarge van het tekstvak telt mee); balken en normstreep schuiven alleen mee (VulBalken zet ze opnieuw)
    t = 0
    For i = 1 To n
        If blijft(i) Then
            t = t + 1
            prefix = "KPI" & nr(i) & "_"
            For j = 1 To dia.Shapes.Count
                Set shp = dia.Shapes(j)
                If Left$(shp.Name, Len(prefix)) = prefix Then
                    mL = shp.Left - links(i)
                    mr = links(i) + breed(i) - (shp.Left + shp.Width)
                    If IsBalkregel(shp.Name) Then
                        shp.Left = doelL(t) + mL
                    Else
                        If binnen >= 0 And IsTekstVanKaart(shp.Name) Then
                            mL = binnen - shp.TextFrame.MarginLeft
                            mr = binnen - shp.TextFrame.MarginRight
                        End If
                        If doelB(t) - mL - mr > 1 Then
                            shp.Left = doelL(t) + mL
                            shp.Width = doelB(t) - mL - mr
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub VulBalken(pres As Object, wsP As Worksheet)
    ' balken op dia 2 (sjabloon v17): een regel op tabblad PowerPoint met vormnaam ..._BALK (ook _BALK_V, _BALK_T, _BALK1..3) heeft
    ' in kolom D een aandeel 0..1; de balk krijgt de linkerrand en aandeel x de breedte van het spoor met dezelfde naam waarin
    ' _BALK door _SPOOR is vervangen (KPI7_BALK_V -> KPI7_SPOOR_V). Een regel ..._NORM (KPI6_NORM) zet de normstreep op aandeel x
    ' breedte van het spoor. Aandeel 0 = balk onzichtbaar, een klein aandeel = een rondje zo breed als het spoor hoog is; de balk
    ' KPIk_BALK krijgt de richtingkleur van KPIk_WAARDE (op norm = groen, onder norm = rood). Na SchikDia2, want die verbreedt de
    ' sporen bij drie kaarten; fouten worden per balk stil overgeslagen.
    Dim r As Long, vorm As String, v As Variant, aandeel As Double, shp As Object, spoor As Object, rW As Long, richting As String
    On Error Resume Next
    r = KPI_ROW1
    Do While Len(CelTekst(wsP.Range("C" & r))) > 0 And r < KPI_ROW1 + 300
        vorm = Trim$(CelTekst(wsP.Range("F" & r)))
        If IsBalkregel(vorm) Then
            v = wsP.Range("D" & r).Value
            aandeel = 0
            If Not IsError(v) Then
                If IsNumeric(v) Then aandeel = CDbl(v)
            End If
            If aandeel < 0 Then aandeel = 0
            If aandeel > 1 Then aandeel = 1
            Set shp = Nothing
            Set spoor = Nothing
            Set shp = ZoekVorm(pres, vorm)
            If InStr(vorm, "_BALK") > 0 Then
                Set spoor = ZoekVorm(pres, Replace(vorm, "_BALK", "_SPOOR"))
            Else
                Set spoor = ZoekVorm(pres, Replace(vorm, "_NORM", "_SPOOR"))
            End If
            If Not shp Is Nothing And Not spoor Is Nothing Then
                If Right$(vorm, 5) = "_NORM" Then
                    shp.Left = spoor.Left + aandeel * spoor.Width - shp.Width / 2
                Else
                    shp.Left = spoor.Left
                    shp.Top = spoor.Top
                    shp.Height = spoor.Height
                    If aandeel * spoor.Width < 0.5 Then
                        shp.Visible = 0
                    Else
                        shp.Visible = -1
                        If aandeel * spoor.Width < spoor.Height Then
                            shp.Width = spoor.Height
                        Else
                            shp.Width = aandeel * spoor.Width
                        End If
                    End If
                    If Right$(vorm, 5) = "_BALK" Then
                        rW = RegelVanVorm(wsP, Replace(vorm, "_BALK", "_WAARDE"))
                        If rW > 0 Then
                            richting = LCase$(Trim$(CelTekst(wsP.Range("E" & rW))))
                            If Len(richting) > 0 Then shp.Fill.ForeColor.RGB = RichtingKleur(richting)
                        End If
                    End If
                End If
            End If
        End If
        r = r + 1
    Loop
    On Error GoTo 0
End Sub

Private Sub SchikGrafiekDia2(shp As Object, ch As Object)
    ' KPI_GRAFIEK (dia 2, sjabloon v17), na het vullen: de waarde-as op Model!BY142..BY143 (as-grenzen uit het Model, gezet zoals
    ' in ZetTijdAs), het grijze vlak GRAFIEK_REALISATIE van de linkerrand van het tekenvlak tot het 'nu'-kwartaal en het nu-blok
    ' (TL_NU_LABEL, KPI4_WAARDE, KPI4_SUB) ernaast. x van kwartaal i (1..n) = links + breedte x (KC_PLOT_X + KC_PLOT_W x (i - 1) /
    ' (n - 1)): het tekenvlak staat vast in het sjabloon (manualLayout, layout.KC_PLOT). Het nu-blok staat links van de stip zoals in
    ' het sjabloon; past dat niet (nu in de eerste kwartalen), dan rechts ervan, links uitgelijnd. Geen gerealiseerd kwartaal: vlak
    ' en blok onzichtbaar. Fouten worden stil overgeslagen.
    Dim vMin As Variant, vMax As Variant, ax As Object, dia As Object, vlak As Object, blok As Object, namen As Variant
    Dim n As Long, posNu As Long, i As Long, plotL As Double, xNu As Double, xOud As Double, dx As Double, bLinks As Double
    On Error Resume Next
    vMin = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_KC_MIN).Value
    vMax = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_KC_MAX).Value
    If Not IsError(vMin) And Not IsError(vMax) Then
        If IsNumeric(vMin) And IsNumeric(vMax) Then
            If CDbl(vMax) > CDbl(vMin) Then
                Set ax = Nothing
                Set ax = ch.Axes(2, 1)
                If Not ax Is Nothing Then
                    ax.MaximumScale = CDbl(vMax)
                    Err.Clear
                    ax.MinimumScale = CDbl(vMin)
                    Err.Clear
                    ax.MaximumScale = CDbl(vMax)
                    Err.Clear
                    ax.MinimumScaleIsAuto = False
                    ax.MaximumScaleIsAuto = False
                    Err.Clear
                End If
            End If
        End If
    End If
    Set dia = shp.Parent
    If dia Is Nothing Then Exit Sub
    n = AantalRijen(CEL_N_ALLE, MAX_RIJEN)
    posNu = AantalRijen(CEL_POS_ACTUALS, MAX_RIJEN)
    Set vlak = VormOpDia(dia, "GRAFIEK_REALISATIE")
    namen = Array("TL_NU_LABEL", "KPI4_WAARDE", "KPI4_SUB")
    If posNu < 1 Or n < 2 Or posNu > n Then
        If Not vlak Is Nothing Then vlak.Visible = 0
        For i = LBound(namen) To UBound(namen)
            Set blok = Nothing
            Set blok = VormOpDia(dia, CStr(namen(i)))
            If Not blok Is Nothing Then blok.Visible = 0
        Next i
        Exit Sub
    End If
    If vlak Is Nothing Then Exit Sub
    plotL = shp.Left + shp.Width * KC_PLOT_X
    xNu = plotL + shp.Width * KC_PLOT_W * (posNu - 1) / (n - 1)
    xOud = vlak.Left + vlak.Width
    vlak.Left = plotL
    If xNu - plotL < 1 Then
        vlak.Visible = 0
    Else
        vlak.Visible = -1
        vlak.Width = xNu - plotL
    End If
    ' nu-blok meeschuiven; links van de stip zoals in het sjabloon, of rechts ervan als het links niet past
    dx = xNu - xOud
    bLinks = 1000000000#
    For i = LBound(namen) To UBound(namen)
        Set blok = Nothing
        Set blok = VormOpDia(dia, CStr(namen(i)))
        If Not blok Is Nothing Then
            blok.Left = blok.Left + dx
            If blok.Left < bLinks Then bLinks = blok.Left
        End If
    Next i
    If bLinks < plotL Then
        dx = (xNu + 7) - bLinks
        For i = LBound(namen) To UBound(namen)
            Set blok = Nothing
            Set blok = VormOpDia(dia, CStr(namen(i)))
            If Not blok Is Nothing Then
                blok.Left = blok.Left + dx
                blok.TextFrame.TextRange.ParagraphFormat.Alignment = 1    ' ppAlignLeft
            End If
        Next i
        ' het getal direct achter 'nu' (het nu-vak is rechts uitgelijnd en breder dan het woord)
        Set blok = Nothing
        Set blok = VormOpDia(dia, "KPI4_WAARDE")
        If Not blok Is Nothing Then blok.Left = xNu + 7 + 14.4
    End If
    On Error GoTo 0
End Sub

Private Function VsDiaAan() As Boolean
    ' dia 4 'Verkoopscenario's' en de scenariolijnen in de cashflowdia: alleen als er op tab Scenario's een scenario is ingevuld
    ' en Naar PowerPoint op ja staat (Model!BY154 = 1)
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_VS_PP).Value
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then VsDiaAan = (CDbl(v) = 1)
End Function

Private Function ModelGetal(cel As String, ByRef ok As Boolean) As Double
    ' getal uit een cel op tabblad Model; ok = False bij leeg, fout of tekst
    Dim v As Variant
    ok = False
    On Error Resume Next
    v = ThisWorkbook.Worksheets(SH_MODEL).Range(cel).Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    ModelGetal = CDbl(v)
    ok = True
End Function

Private Sub SchikDiaScenarios(pres As Object, wsP As Worksheet)
    ' dia 4 'Verkoopscenario's' (sjabloon v18), na het vullen van de teksten. Komt er geen scenario in PowerPoint (niets ingevuld op
    ' tab Scenario's, of Naar PowerPoint = nee), dan gaat de hele dia weg. Anders gaan de kaart (VS_Kk_*) en de legenda-regel
    ' (VS_LEGk, VS_LEGk_LIJN) weg van elk scenario waarvan de regel VS_Kk_WAARDE op tab PowerPoint leeg is; de overige kaarten
    ' verdelen de breedte van het kaartenblok (van de linkerrand van de eerste tot de rechterrand van de laatste kaart) met dezelfde
    ' tussenruimte, en de legenda-regels schuiven aan op de plekken uit het sjabloon. Teksten in een kaart houden hun afstand tot
    ' beide randen (en rekken dus mee), de accentstreep schuift alleen mee. Fouten worden stil overgeslagen.
    Dim dia As Object, shp As Object, k As Long, j As Long, rW As Long, n As Long, nBlijft As Long, t As Long
    Dim links(0 To N_VS) As Double, breed(0 To N_VS) As Double, legL(0 To N_VS) As Double
    Dim blijft(0 To N_VS) As Boolean, bestaat(0 To N_VS) As Boolean
    Dim bLinks As Double, bRechts As Double, tussen As Double, w As Double, doelL As Double, mL As Double, mr As Double, dx As Double
    Dim prefix As String
    On Error Resume Next
    Set shp = ZoekVorm(pres, "VS_GRAFIEK")
    If shp Is Nothing Then Set shp = ZoekVorm(pres, "VS_K0_KAART")
    If shp Is Nothing Then Exit Sub
    Set dia = shp.Parent
    If dia Is Nothing Then Exit Sub
    If Not VsDiaAan() Then
        dia.Delete
        Exit Sub
    End If
    bLinks = 1E+20
    bRechts = -1E+20
    tussen = -1
    For k = 0 To N_VS
        Set shp = Nothing
        Set shp = VormOpDia(dia, "VS_K" & k & "_KAART")
        If Not shp Is Nothing Then
            bestaat(k) = True
            links(k) = shp.Left
            breed(k) = shp.Width
            If links(k) < bLinks Then bLinks = links(k)
            If links(k) + breed(k) > bRechts Then bRechts = links(k) + breed(k)
            If k = 0 Then
                blijft(k) = True
            Else
                rW = RegelVanVorm(wsP, "VS_K" & k & "_WAARDE")
                blijft(k) = False
                If rW > 0 Then blijft(k) = (Len(CelTekst(wsP.Range("D" & rW))) > 0)
            End If
            n = n + 1
            If blijft(k) Then nBlijft = nBlijft + 1
        End If
        Set shp = Nothing
        Set shp = VormOpDia(dia, "VS_LEG" & k & "_LIJN")
        If Not shp Is Nothing Then legL(k) = shp.Left
    Next k
    For k = 1 To N_VS
        If bestaat(k) And bestaat(k - 1) Then
            If tussen < 0 Or links(k) - (links(k - 1) + breed(k - 1)) < tussen Then tussen = links(k) - (links(k - 1) + breed(k - 1))
        End If
    Next k
    If tussen < 0 Then tussen = 0
    If n = 0 Or nBlijft = 0 Or nBlijft = n Then Exit Sub

    ' kaarten en legenda-regels van lege scenario's weg
    For k = 1 To N_VS
        If bestaat(k) And Not blijft(k) Then
            prefix = "VS_K" & k & "_"
            For j = dia.Shapes.Count To 1 Step -1
                Set shp = dia.Shapes(j)
                If Left$(shp.Name, Len(prefix)) = prefix Or shp.Name = "VS_LEG" & k Or shp.Name = "VS_LEG" & k & "_LIJN" Then shp.Delete
            Next j
        End If
    Next k

    ' de overige kaarten gelijk verdelen, legenda-regels op de plekken van links af
    w = (bRechts - bLinks - (nBlijft - 1) * tussen) / nBlijft
    t = 0
    For k = 0 To N_VS
        If bestaat(k) And blijft(k) Then
            doelL = bLinks + t * (w + tussen)
            prefix = "VS_K" & k & "_"
            For j = 1 To dia.Shapes.Count
                Set shp = dia.Shapes(j)
                If Left$(shp.Name, Len(prefix)) = prefix Then
                    mL = shp.Left - links(k)
                    mr = links(k) + breed(k) - (shp.Left + shp.Width)
                    If shp.Name = prefix & "ACCENT" Then
                        shp.Left = doelL + mL
                    ElseIf w - mL - mr > 1 Then
                        shp.Left = doelL + mL
                        shp.Width = w - mL - mr
                    End If
                End If
            Next j
            dx = legL(t) - legL(k)
            If dx <> 0 Then
                Set shp = Nothing
                Set shp = VormOpDia(dia, "VS_LEG" & k & "_LIJN")
                If Not shp Is Nothing Then shp.Left = shp.Left + dx
                Set shp = Nothing
                Set shp = VormOpDia(dia, "VS_LEG" & k)
                If Not shp Is Nothing Then shp.Left = shp.Left + dx
            End If
            t = t + 1
        End If
    Next k
    On Error GoTo 0
End Sub

Private Sub SchikGrafiekVS(shp As Object, ch As Object)
    ' VS_GRAFIEK (dia 4, sjabloon v18), na het vullen: de waarde-as op Model!BY155..BY157 (minimum, maximum, stap), het grijze vlak
    ' VS_REALISATIE (met het label VS_REAL_LBL) van de linkerrand van het tekenvlak tot het 'nu'-kwartaal (positie in het tijdvak:
    ' Model!BY171; x van kwartaal i (1..n) = links + breedte x (VS_PLOT_X + VS_PLOT_W x (i - 1) / (n - 1)), het tekenvlak staat vast
    ' in het sjabloon (manualLayout, layout.VS_PLOT)). Geen gerealiseerd kwartaal in het tijdvak: vlak en label onzichtbaar.
    ' Fouten worden stil overgeslagen.
    Dim vMin As Double, vMax As Double, stap As Double, okMin As Boolean, okMax As Boolean, okStap As Boolean, ax As Object
    Dim dia As Object, vlak As Object, lbl As Object, n As Long, posNu As Long, plotL As Double, xNu As Double
    On Error Resume Next
    vMin = ModelGetal(CEL_VS_MIN, okMin)
    vMax = ModelGetal(CEL_VS_MAX, okMax)
    stap = ModelGetal(CEL_VS_STAP, okStap)
    If okMin And okMax Then
        If vMax > vMin Then
            Set ax = Nothing
            Set ax = ch.Axes(2, 1)
            If Not ax Is Nothing Then
                ax.MaximumScale = vMax
                Err.Clear
                ax.MinimumScale = vMin
                Err.Clear
                ax.MaximumScale = vMax
                Err.Clear
                ax.MinimumScaleIsAuto = False
                ax.MaximumScaleIsAuto = False
                If okStap And stap > 0 Then ax.MajorUnit = stap
                Err.Clear
            End If
        End If
    End If
    Set dia = shp.Parent
    If dia Is Nothing Then Exit Sub
    n = AantalRijen(CEL_VP_N, MAX_RIJEN)
    posNu = AantalRijen(CEL_VS_NU, MAX_RIJEN)
    plotL = shp.Left + shp.Width * VS_PLOT_X
    Set vlak = VormOpDia(dia, "VS_REALISATIE")
    Set lbl = VormOpDia(dia, "VS_REAL_LBL")
    If posNu < 1 Or n < 2 Or posNu > n Then
        If Not vlak Is Nothing Then vlak.Visible = 0
        If Not lbl Is Nothing Then lbl.Visible = 0
    ElseIf Not vlak Is Nothing Then
        xNu = plotL + shp.Width * VS_PLOT_W * (posNu - 1) / (n - 1)
        vlak.Left = plotL
        If xNu - plotL < 1 Then
            vlak.Visible = 0
            If Not lbl Is Nothing Then lbl.Visible = 0
        Else
            vlak.Visible = -1
            vlak.Width = xNu - plotL
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub VulTabel(pres As Object, vorm As String, kop As Range, nRijen As Long, nKol As Long, overslaan As Boolean, _
                     ByRef mist As String, ByRef fouten As String)
    Dim shp As Object, tbl As Object, r As Long, c As Long, rij As Long
    Set shp = ZoekVorm(pres, vorm)
    If shp Is Nothing Then
        ' sinds sjabloon v20 heeft het sjabloon geen tabel woningtypes meer: geen melding
        If vorm <> "VT_TABEL" Then mist = mist & vorm & vbCrLf
        Exit Sub
    End If
    On Error GoTo Mislukt
    Set tbl = shp.Table
    rij = 0
    For r = 0 To nRijen - 1
        If overslaan And r > 0 And Len(CelTekst(kop.Offset(r, 0))) = 0 Then
            ' lege regel (woningtype zonder aantal): niet op de dia
        Else
            rij = rij + 1
            If rij > tbl.Rows.Count Then tbl.Rows.Add
            For c = 0 To nKol - 1
                If c < tbl.Columns.Count Then tbl.Cell(rij, c + 1).Shape.TextFrame.TextRange.Text = CelTekst(kop.Offset(r, c))
            Next c
        End If
    Next r
    If overslaan Then
        Do While tbl.Rows.Count > rij And tbl.Rows.Count > 1
            tbl.Rows(tbl.Rows.Count).Delete
        Loop
    End If
    If vorm = "VT_TABEL" Then SchikDia5 pres, tbl
    Exit Sub
Mislukt:
    fouten = fouten & vorm & ": " & Err.Description & vbCrLf
End Sub

Private Function ZelfdeDia(a As Object, b As Object) As Boolean
    On Error Resume Next
    ZelfdeDia = (a.Parent.SlideID = b.Parent.SlideID)
End Function

Private Sub SchikDia5(pres As Object, tbl As Object)
    ' na het vullen van VT_TABEL (dia 5, onder de bouwtermijnengrafiek BT_GRAFIEK in kaart VT_KAART):
    ' bij meer dan 4 typeregels de tabel binnen de kaart en boven het dianummer houden
    '   a) rijen lager (tot 10 pt) en letters kleiner (tot 5,5 pt), celmarges boven en onder 1 pt
    '   b) past het dan nog niet: grafiekkaart en grafiek inkorten (grafiek minimaal 110 pt), tabelkaart en tabel omhoog
    '   c) tabelkaart eindigt 8 pt onder de tabel
    ' alles in punten; fouten worden stil overgeslagen (de dia blijft dan zoals het sjabloon hem had)
    Dim tabelVorm As Object, tabelKaart As Object, kaart As Object, grafiek As Object
    Dim nr As Long, r As Long, c As Long
    Dim onder As Double, beschikbaar As Double, rijH As Double, letter As Double
    Dim tekort As Double, extra As Double, krimp As Double, schuif As Double, tabelH As Double
    Const RIJ_MAX As Double = 14.4
    Const RIJ_MIN As Double = 10
    Const LETTER_MAX As Double = 7.5                           ' lettergrootte van de cellen in het sjabloon
    Const LETTER_MIN As Double = 5.5
    Const AFSTAND As Double = 8
    Const GRAFIEK_MIN As Double = 110

    On Error Resume Next
    nr = tbl.Rows.Count
    If nr <= 5 Then Exit Sub                                     ' kop + 4 types: het sjabloon past zo
    Set tabelVorm = ZoekVorm(pres, "VT_TABEL")
    Set tabelKaart = ZoekVorm(pres, "VT_TABELKAART")
    Set kaart = ZoekVorm(pres, "VT_KAART")
    Set grafiek = ZoekVorm(pres, "BT_GRAFIEK")
    If tabelVorm Is Nothing Then Exit Sub

    ' ondergrens: 30 pt boven de dia-rand, daar staat het dianummer
    onder = 0
    onder = pres.PageSetup.SlideHeight - 30
    If onder <= 0 Then onder = 375
    beschikbaar = onder - tabelVorm.Top

    ' b) past de tabel ook met de laagste rijen niet, dan ruimte halen bij de grafiekkaart:
    '    eerst de ruimte tussen de kaarten boven de 8 pt, de rest door grafiekkaart en grafiek in te korten
    '    (grafiek minimaal 110 pt); tabelkaart en tabel schuiven omhoog en houden 8 pt tot de grafiekkaart
    tekort = nr * RIJ_MIN - beschikbaar
    ' sinds sjabloon v20 staat de tabel op een eigen dia: dan de grafiek niet inkorten
    If tekort > 0 And Not kaart Is Nothing And Not grafiek Is Nothing And ZelfdeDia(tabelVorm, grafiek) Then
        extra = 0
        If Not tabelKaart Is Nothing Then extra = tabelKaart.Top - (kaart.Top + kaart.Height) - AFSTAND
        krimp = tekort - extra
        If krimp > grafiek.Height - GRAFIEK_MIN Then krimp = grafiek.Height - GRAFIEK_MIN
        If krimp < 0 Then krimp = 0
        schuif = krimp + extra
        If schuif > tekort Then schuif = tekort
        If krimp > 0 Then
            kaart.Height = kaart.Height - krimp
            grafiek.Height = grafiek.Height - krimp
        End If
        If schuif > 0 Then
            If Not tabelKaart Is Nothing Then tabelKaart.Top = tabelKaart.Top - schuif
            tabelVorm.Top = tabelVorm.Top - schuif
            beschikbaar = onder - tabelVorm.Top
        End If
    End If

    ' a) rijhoogte en lettergrootte
    rijH = beschikbaar / nr
    If rijH > RIJ_MAX Then rijH = RIJ_MAX
    If rijH < RIJ_MIN Then rijH = RIJ_MIN
    letter = LETTER_MAX * rijH / RIJ_MAX
    If letter < LETTER_MIN Then letter = LETTER_MIN
    ' eerst letters en marges, dan de hoogte: PowerPoint maakt een rij nooit lager dan zijn tekst
    For r = 1 To nr
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.TextFrame
                .MarginTop = 1
                .MarginBottom = 1
                .TextRange.Font.Size = letter
            End With
        Next c
    Next r
    For r = 1 To nr
        tbl.Rows(r).Height = rijH
    Next r

    ' c) tabelkaart om de tabel sluiten; de rijhoogten teruglezen, want een gezette hoogte is een minimum
    tabelH = 0
    For r = 1 To nr
        tabelH = tabelH + tbl.Rows(r).Height
    Next r
    If Not tabelKaart Is Nothing And tabelH > 0 Then
        tabelKaart.Height = tabelVorm.Top + tabelH + AFSTAND - tabelKaart.Top
    End If
    On Error GoTo 0
End Sub

Private Sub VulGrafiek(pres As Object, vorm As String, kop As Range, n As Long, nKol As Long, ByRef mist As String, ByRef fouten As String)
    ' zoekt het grafiekvak op naam (eerste dia waarop het staat) en vult het
    Dim shp As Object
    Set shp = ZoekVorm(pres, vorm)
    If shp Is Nothing Then
        mist = mist & vorm & vbCrLf
        Exit Sub
    End If
    VulGrafiekVorm shp, vorm, kop, n, nKol, fouten
End Sub

Private Sub VulGrafiekVorm(shp As Object, vorm As String, kop As Range, n As Long, nKol As Long, ByRef fouten As String)
    ' zet het blok (kop + n rijen) in het gegevensblad van de grafiek en past het bereik van elke reeks aan:
    ' X-bereik = kolom A, Y-bereik = de kolomletter uit de SERIES-formule van die reeks (KolomVanReeks). Reeksnamen die
    ' naar rij 1 van het gegevensblad verwijzen volgen zo de koptekst; namen als tekst in de formule (BT_GRAFIEK) blijven.
    Dim ch As Object, wbE As Object, wsE As Object
    Dim arr() As Variant, r As Long, c As Long, v As Variant, i As Long, kol As String, letter As Double
    On Error GoTo Mislukt
    Set ch = shp.Chart
    ReDim arr(0 To n, 0 To nKol - 1)
    For r = 0 To n
        For c = 0 To nKol - 1
            v = kop.Offset(r, c).Value
            If IsError(v) Then
                arr(r, c) = Empty
            ElseIf Len(CStr(v)) = 0 Then
                arr(r, c) = Empty
            Else
                arr(r, c) = v
            End If
        Next c
    Next r
    ' verkoop / transport: kolom V (VP1) of W (VP2) = som van de typen per kwartaal (ook de totaalrij), de hoogte waar
    ' het label 'Opsomming' boven de staaf staat
    If vorm = "VP1_GRAFIEK" Or vorm = "VP2_GRAFIEK" Then OpsommingSom arr, n, IIf(vorm = "VP1_GRAFIEK", 1, 11)
    ch.ChartData.Activate
    DoEvents
    Set wbE = ch.ChartData.Workbook
    Set wsE = wbE.Worksheets(1)
    ' wissen tot de grootste bloklengte (MAX_RIJEN_BT >= MAX_RIJEN): veilig voor alle grafieken, ook als het sjabloon
    ' meer voorbeeldrijen bevat dan het blok nu heeft
    wsE.Range("A1").Resize(MAX_RIJEN_BT + 1, nKol).ClearContents
    wsE.Range("A1").Resize(n + 1, nKol).Value = arr
    For i = 1 To ch.SeriesCollection.Count
        kol = KolomVanReeks(CStr(ch.SeriesCollection(i).Formula))
        If Len(kol) > 0 Then ZetBereik ch.SeriesCollection(i), wsE, kol, n
    Next i
    ' reeksen met '(uit)' in de kop gaan weg: de scenariolijn in de cashflowgrafiek, in de bouwtermijnengrafiek de typen
    ' zonder termijnen en de jaren buiten het bereik, in de grafiek per type de onderdelen die het type niet heeft
    If vorm = "CF_GRAFIEK" Or vorm = "BT_GRAFIEK" Or vorm = "PT_GRAFIEK" Or vorm = "KPI_GRAFIEK" Or vorm = "VS_GRAFIEK" _
       Or vorm = "VP1_GRAFIEK" Or vorm = "VP2_GRAFIEK" Then VerwijderUitgezetteReeksen ch, wsE
    If vorm = "BT_GRAFIEK" Then
        ZetTijdAs ch
        ' geen legenda: na het weghalen van reeksen verschuiven de verborgen legenda-regels en blijven de kwartaalnummers
        ' ('1'..'4') als losse cijfers onder de grafiek staan; de typekleuren staan in de tabel eronder (VT_TABEL)
        On Error Resume Next
        ch.HasLegend = False
        ' rijlabels passend bij het aantal rijen (met DAEB-fees en kopregels zijn er meer): er staat hooguit één label per drie rijen
        letter = 7.5
        If n > 0 Then letter = 2.4 * ch.PlotArea.InsideHeight / n
        If letter > 8 Then letter = 8
        If letter < 7 Then letter = 7
        ch.Axes(1).TickLabels.Font.Size = letter
        On Error GoTo Mislukt
    End If
    If vorm = "VP1_GRAFIEK" Or vorm = "VP2_GRAFIEK" Then OpsommingLabels ch, arr, n, IIf(vorm = "VP1_GRAFIEK", 1, 11)
    If vorm = "KPI_GRAFIEK" Then SchikGrafiekDia2 shp, ch
    If vorm = "VS_GRAFIEK" Then SchikGrafiekVS shp, ch
    wbE.Close
    Set wbE = Nothing
    Exit Sub
Mislukt:
    fouten = fouten & vorm & ": " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wbE Is Nothing Then wbE.Close
    On Error GoTo 0
End Sub

Private Sub OpsommingSom(arr() As Variant, n As Long, k0 As Long)
    ' som van de tien typekolommen (vanaf k0) per rij in kolom V (k0 = 1, verkocht) of W (k0 = 11, getransporteerd)
    Dim r As Long, k As Long, som As Double
    For r = 1 To n
        som = 0
        For k = k0 To k0 + 9
            If IsNumeric(arr(r, k)) And Not IsEmpty(arr(r, k)) Then
                If arr(r, k) > 0 Then som = som + arr(r, k)
            End If
        Next k
        If som > 0 Then arr(r, IIf(k0 = 1, 21, 22)) = som Else arr(r, IIf(k0 = 1, 21, 22)) = Empty
    Next r
End Sub

Private Sub OpsommingLabels(ch As Object, arr() As Variant, n As Long, k0 As Long)
    ' boven elke staaf de aantallen per type ('10 + 2'), zodat ook een dun staafdeel een leesbaar cijfer heeft
    Dim i As Long, r As Long, k As Long, tekst As String, ser As Object
    On Error Resume Next
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).Name = "Opsomming" Then Set ser = ch.SeriesCollection(i)
    Next i
    If ser Is Nothing Then Exit Sub
    ser.HasDataLabels = True
    For r = 1 To n
        tekst = ""
        For k = k0 To k0 + 9
            If IsNumeric(arr(r, k)) And Not IsEmpty(arr(r, k)) Then
                If arr(r, k) > 0 Then tekst = tekst & IIf(Len(tekst) > 0, " + ", "") & CStr(arr(r, k))
            End If
        Next k
        If Len(tekst) > 0 Then
            ser.Points(r).HasDataLabel = True
            ser.Points(r).DataLabel.Text = tekst
        Else
            ser.Points(r).HasDataLabel = False
        End If
    Next r
    ' in de staaf alleen cijfers die passen (staafdeel >= 6% van de hoogste staaf); de rest staat in de opsomming erboven
    Dim maxSom As Double, b As Object, w As Variant
    For r = 1 To n
        If IsNumeric(arr(r, IIf(k0 = 1, 21, 22))) And Not IsEmpty(arr(r, IIf(k0 = 1, 21, 22))) Then
            If arr(r, IIf(k0 = 1, 21, 22)) > maxSom Then maxSom = arr(r, IIf(k0 = 1, 21, 22))
        End If
    Next r
    If maxSom <= 0 Then Exit Sub
    For Each b In ch.SeriesCollection
        If b.ChartType = 52 Then                       ' xlColumnStacked
            w = b.Values
            For r = 1 To n
                If IsNumeric(w(r)) And Not IsEmpty(w(r)) Then
                    If w(r) > 0 And w(r) < 0.06 * maxSom Then b.Points(r).HasDataLabel = False
                End If
            Next r
        End If
    Next b
    On Error GoTo 0
End Sub

Private Sub VerwijderUitgezetteReeksen(ch As Object, wsE As Object)
    ' Een reeks waarvan de kop in het gegevensblad eindigt op '(uit)' (de scenariolijn bij 'Scenariolijn tonen = nee',
    ' tab PowerPoint kolom N; in de bouwtermijnengrafiek typen zonder termijnen en jaren buiten het bereik) wordt uit de
    ' grafiek verwijderd, zodat hij niet als lege vermelding in de legenda staat.
    ' Gekozen voor verwijderen in plaats van Legend.LegendEntries(k).Delete: LegendEntries telt alleen zichtbare entries
    ' (in sjabloon v11 zijn de entries van reeks 0 en 1, de vlakken voorfinanciering/positief saldo, verborgen, dus de
    ' scenariolijn zou entry 5 zijn) en die telling verschuift zodra iemand het sjabloon aanpast. Verwijderen hangt niet
    ' van een index af en is veilig: het sjabloon wordt per run opnieuw geopend, alleen de nieuwe presentatie verandert.
    ' Van achteren naar voren, omdat de nummering van SeriesCollection verschuift na een Delete.
    Dim i As Long, kol As String, kopTekst As String
    On Error Resume Next
    For i = ch.SeriesCollection.Count To 1 Step -1
        kol = KolomVanReeks(CStr(ch.SeriesCollection(i).Formula))
        If Len(kol) > 0 Then
            kopTekst = ""
            kopTekst = Trim$(CStr(wsE.Range(kol & "1").Value))
            If Right$(kopTekst, 5) = "(uit)" Then ch.SeriesCollection(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub ZetTijdAs(ch As Object)
    ' BT_GRAFIEK: beide waarde-assen (tijd in kwartaalindex) vast op Model!BY97 (1 januari van het jaar twee jaar vóór de eerste
    ' termijn, niet vóór de eerste periode) .. BY99 (einde laatste
    ' jaar), zodat de blokjes van vorige prognose en huidige planning op dezelfde schaal staan als de koprijen.
    ' Late binding: Axes(2 = xlValue, 1 = xlPrimary / 2 = xlSecondary). Eerst het maximum, dan het minimum en nog eens
    ' het maximum: zo lukt het ook als het nieuwe bereik helemaal boven of onder de voorbeeldwaarden van het sjabloon
    ' ligt (PowerPoint weigert een minimum boven het huidige maximum en andersom). Ontbreekt de tweede as, dan blijft
    ' de rest staan: alles onder On Error.
    Dim jaar1 As Variant, jaar2 As Variant, asNr As Long, ax As Object
    On Error Resume Next
    jaar1 = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_AS_MIN).Value
    jaar2 = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_AS_MAX).Value
    If IsError(jaar1) Or IsError(jaar2) Then Exit Sub
    If Not IsNumeric(jaar1) Or Not IsNumeric(jaar2) Then Exit Sub
    If CDbl(jaar2) <= CDbl(jaar1) Then Exit Sub
    For asNr = 1 To 2
        Set ax = Nothing
        Set ax = ch.Axes(2, asNr)
        If Not ax Is Nothing Then
            ax.MaximumScale = CDbl(jaar2)
            Err.Clear
            ax.MinimumScale = CDbl(jaar1)
            Err.Clear
            ax.MaximumScale = CDbl(jaar2)
            Err.Clear
            ax.MinimumScaleIsAuto = False
            ax.MaximumScaleIsAuto = False
            Err.Clear
        End If
    Next asNr
    On Error GoTo 0
End Sub

Private Sub VulDiasPerType(pres As Object, wsP As Worksheet, n As Long, ByRef mist As String, ByRef fouten As String)
    ' Dia 8 (PT_GRAFIEK): volgens het Dashboard (CEL_PT_PP) één dia met het totaal van alle niet-DAEB-typen ('totaal (zonder DAEB)',
    ' standaard; keuzecel = 0), per woningtype met aantal > 0 een kopie ('per type'), of eerst het totaal en dan per type ('totaal
    ' en per type'). Het Dashboard heeft een keuzecel met de typenaam; het Model rekent het selectieblok (tab PowerPoint vanaf BLOK_PT)
    ' voor dat type of het totaal uit. De macro zet de keuzecel, rekent door, kopieert de dia (achter de vorige kopie) en vult titel,
    ' ondertitel en grafiek op de kopie zelf (niet via ZoekVorm, dat de eerste dia met die vormnaam zou pakken). Daarna gaat dia 8
    ' zelf weg en krijgt de keuzecel zijn oude waarde terug. Geen koopwoningen: per type. Geen type: alleen dia 8 verwijderen.
    Dim wsT As Worksheet, wsD As Worksheet, shp As Object, sjabloon As Object
    Dim oud As Variant, pos As Long, c As Long, naam As String, aantal As Variant, idx As Long, nKopie As Long
    Dim rTitel As Long, rSub As Long, r As Long, modus As String, totaal As Boolean, perType As Boolean, nKoop As Long
    Set shp = ZoekVorm(pres, "PT_GRAFIEK")
    If shp Is Nothing Then
        mist = mist & "PT_GRAFIEK" & vbCrLf
        Exit Sub
    End If
    On Error GoTo Mislukt
    Set sjabloon = shp.Parent
    idx = sjabloon.SlideIndex
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)
    oud = wsD.Range(CEL_PT_KEUZE).Formula
    modus = LCase$(Trim$(CelTekst(wsD.Range(CEL_PT_PP))))
    perType = (Left$(modus, 8) = "per type") Or (InStr(modus, "en per type") > 0)
    totaal = (Left$(modus, 8) <> "per type")
    ' regels op tab PowerPoint met de titel en ondertitel van dia 8 (kolom F = vormnaam)
    For r = KPI_ROW1 To KPI_ROW1 + 300
        If Len(CelTekst(wsP.Range("C" & r))) = 0 Then Exit For
        If Trim$(CelTekst(wsP.Range("F" & r))) = "PT_TITEL" Then rTitel = r
        If Trim$(CelTekst(wsP.Range("F" & r))) = "PT_SUBTITEL" Then rSub = r
    Next r
    nKopie = 0
    If totaal Then
        For pos = 1 To MAX_TYPES
            c = TYPE_COL1 + (pos - 1) * TYPE_BREEDTE
            aantal = wsT.Cells(TYPE_RIJ_AANTAL, c).Value
            If IsNumeric(aantal) And Not IsEmpty(aantal) Then
                If CDbl(aantal) > 0 And UCase$(Trim$(CelTekst(wsT.Cells(TYPE_RIJ_SOORT, c)))) <> "DAEB" Then nKoop = nKoop + 1
            End If
        Next pos
        If nKoop > 0 Then
            naam = PT_TOTAAL
            wsD.Range(CEL_PT_KEUZE).Value = 0                ' 0 = totaal van alle niet-DAEB-typen
            Application.Calculate
            VulKopiePT sjabloon, idx, nKopie, wsP, rTitel, rSub, naam, fouten
        Else
            perType = True                                   ' geen koopwoningen (alleen DAEB): dan per type
        End If
    End If
    If perType Then
        For pos = 1 To MAX_TYPES
            c = TYPE_COL1 + (pos - 1) * TYPE_BREEDTE
            aantal = wsT.Cells(TYPE_RIJ_AANTAL, c).Value
            If IsNumeric(aantal) And Not IsEmpty(aantal) Then
                If CDbl(aantal) > 0 Then
                    naam = CStr(wsT.Cells(TYPE_RIJ_NAAM, c).Value)
                    If Len(naam) = 0 Then naam = "Type " & pos
                    wsD.Range(CEL_PT_KEUZE).Value = pos      ' typenummer: het Model kiest op nummer, dus ook bij dubbele namen goed
                    Application.Calculate
                    VulKopiePT sjabloon, idx, nKopie, wsP, rTitel, rSub, naam, fouten
                End If
            End If
        Next pos
    End If
    sjabloon.Delete
    wsD.Range(CEL_PT_KEUZE).Formula = oud
    Application.Calculate
    Exit Sub
Mislukt:
    fouten = fouten & "dia per woningtype (" & naam & "): " & Err.Description & vbCrLf
    On Error Resume Next
    wsD.Range(CEL_PT_KEUZE).Formula = oud
    Application.Calculate
    On Error GoTo 0
End Sub

Private Sub VulKopiePT(sjabloon As Object, idx As Long, ByRef nKopie As Long, wsP As Worksheet, rTitel As Long, rSub As Long, _
                       naam As String, ByRef fouten As String)
    ' kopie van dia 8 achter de vorige kopie, gevuld met het type (of het totaal) dat nu in de keuzecel staat; een fout gaat naar de
    ' foutafhandeling van VulDiasPerType
    Dim rngK As Object, kopie As Object, fouten2 As String
    Set rngK = sjabloon.Duplicate
    Set kopie = rngK.Item(1)
    nKopie = nKopie + 1
    kopie.MoveTo idx + nKopie
    On Error Resume Next                                   ' titel en ondertitel: alleen als de vorm op de dia staat
    If rTitel > 0 Then kopie.Shapes("PT_TITEL").TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & rTitel))
    If rSub > 0 Then kopie.Shapes("PT_SUBTITEL").TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & rSub))
    On Error GoTo 0
    VulGrafiekVorm kopie.Shapes("PT_GRAFIEK"), "PT_GRAFIEK", wsP.Range(BLOK_PT), AantalRijen(CEL_PT_N, MAX_RIJEN), KOL_PT, fouten2
    If Len(fouten2) > 0 Then fouten = fouten & naam & " · " & fouten2
End Sub

Private Sub ZetBereik(reeks As Object, wsE As Object, kol As String, n As Long)
    ' eerst als verwijzing; lukt dat niet in deze Office-versie, dan als bereik
    Dim blad As String
    blad = "='" & wsE.Name & "'!$"
    On Error Resume Next
    reeks.XValues = blad & "A$2:$A$" & (n + 1)
    If Err.Number <> 0 Then
        Err.Clear
        reeks.XValues = wsE.Range("A2:A" & (n + 1))
    End If
    reeks.Values = blad & kol & "$2:$" & kol & "$" & (n + 1)
    If Err.Number <> 0 Then
        Err.Clear
        reeks.Values = wsE.Range(kol & "2:" & kol & (n + 1))
    End If
    On Error GoTo 0
End Sub

Private Function KolomVanReeks(f As String) As String
    ' kolomletter uit het waardenbereik van =SERIES(naam,categorieen,waarden,volgorde)
    ' werkt voor naam als verwijzing (Sheet1!$C$1), als tekst ("Gerealiseerd"), leeg (=SERIES(,...): delen(0) is dan
    ' '=SERIES(' en delen(2) nog steeds de waarden) en met aanhalingstekens om de bladnaam ('Blad 1'!$C$2:$C$21)
    Dim delen() As String, s As String, p As Long, q As Long
    delen = Split(f, ",")
    If UBound(delen) < 2 Then Exit Function
    s = delen(2)
    p = InStr(s, "!$")
    If p = 0 Then Exit Function
    q = InStr(p + 2, s, "$")
    If q = 0 Then Exit Function
    KolomVanReeks = Mid$(s, p + 2, q - p - 2)
End Function

' -------------------------------------------------------------------------------------
'  4b. Ophalen uit FO: tab Invoer en tab Woningtypes vullen vanuit een FO-werkboek
'      Alles wordt op labels gezocht (zelfde regels als forecast_builder/fo.py):
'      '1. Cashflow'     koprij 'Jaar' in kolom B; daaronder jaar, Q, kosten, % kosten, omzet, % omzet, CF, CF x 1000, CF vorig
'      'FO - actuals'    laatste boekjaar/kwartaal met boekingen -> Dashboard 'Actuals t/m'
'      Vóór het leegmaken vraagt de knop of de planning die er nu staat als vorige prognose bewaard moet worden (blok 4 en 6);
'      de knop 'Vorige prognose uit FO' (tab Woningtypes, VorigePrognoseUitFO) haalt die uit een ouder FO.
'      'CF - opbrengsten' per type: blok 'Termijnen k - naam' (rij erboven '# won' en het aantal), bouwplanning (aantal
'                        woningen per termijn per kwartaal), blok 'Omzet HVG Termijnen' (Koopsom, termijnen %/EUR/EUR per wn,
'                        extra's; DAEB: fee-componenten met alleen een totaal en daaronder de termijnen), blok 'Verkooptempo'
'                        (verkocht per kwartaal; 'actuals' = al verkocht). Een los DAEB-blok (bv. 'Opbrengsten DAEB', de
'                        componenten met hun totaal boven de koprij) wordt een eigen DAEB-type (FODaeb): aantal en verkoop van
'                        tabblad 'Verkooptempo', transport = verkoop. Fee-bedragen uit de cellen, per kwartaal gesplitst (FeeDelen).
' -------------------------------------------------------------------------------------
Private Sub FOOpenen(pad As String, ByRef wbF As Workbook, ByRef zelfGeopend As Boolean, ByRef iteratie As Boolean)
    ' staat het FO al open, dan dat gebruiken (en open laten); anders alleen-lezen openen. Een FO kan zelf kringverwijzingen
    ' bevatten (bv. een formule die naar zijn eigen rij verwijst): zolang het hier open staat geen melding daarover
    ' (iteratief rekenen aan, meldingen uit); dit werkboek zelf heeft geen kringverwijzingen
    Dim wb As Workbook
    Set wbF = Nothing
    For Each wb In Workbooks
        If LCase$(wb.FullName) = LCase$(pad) Then Set wbF = wb
    Next wb
    zelfGeopend = False
    If wbF Is Nothing Then
        iteratie = Application.Iteration
        Application.Iteration = True
        Application.DisplayAlerts = False
        Set wbF = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True)
        Application.DisplayAlerts = True
        zelfGeopend = True
    End If
End Sub

Private Sub FOSluiten(ByRef wbF As Workbook, zelfGeopend As Boolean, iteratie As Boolean)
    If zelfGeopend And Not wbF Is Nothing Then wbF.Close SaveChanges:=False
    If zelfGeopend Then Application.Iteration = iteratie
    Set wbF = Nothing
End Sub

Private Function FOOpbrengstenBlad(wbF As Workbook) As Worksheet
    Set FOOpbrengstenBlad = ZoekBlad(wbF, "CF - opbrengsten")
    If FOOpbrengstenBlad Is Nothing Then Set FOOpbrengstenBlad = ZoekBlad(wbF, "CF-opbrengsten")
    If FOOpbrengstenBlad Is Nothing Then Set FOOpbrengstenBlad = ZoekBlad(wbF, "Opbrengsten")
End Function

Private Function KwTekst(idx As Long) As String
    ' kwartaalindex -> 'jaar Qk' (de schrijfwijze van blok 5 en 6)
    KwTekst = ((idx - 1) \ 4) & " Q" & (idx - ((idx - 1) \ 4) * 4)
End Function

Public Sub VorigePrognoseUitFO()
    ' knop 'Vorige prognose uit FO' (tab Woningtypes): uit een ouder FO per type de start bouw (blok 4) en per fee-termijn het
    ' kwartaal (blok 6) halen; typen op naam, anders op positie; fee-termijnen op mijlpaal binnen dezelfde component
    Dim pad As Variant, wbF As Workbook, wsO As Worksheet, wsT As Worksheet, zelfGeopend As Boolean, iteratie As Boolean
    Dim stap As String, foutNr As Long, foutOms As String, waarschuwing As String, n As Long, rekenmodus As Long, bericht As String
    foStap = ""
    On Error GoTo Fout
#If Mac Then
    pad = Application.GetOpenFilename()
#Else
    pad = Application.GetOpenFilename("Excel-werkboeken (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", , "Ouder FO-werkboek (vorige prognose) kiezen")
#End If
    If VarType(pad) = vbBoolean Then Exit Sub
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Application.ScreenUpdating = False
    rekenmodus = Application.Calculation
    Application.Calculation = -4135
    stap = "FO openen"
    FOOpenen CStr(pad), wbF, zelfGeopend, iteratie
    Set wsO = FOOpbrengstenBlad(wbF)
    If wsO Is Nothing Then Err.Raise 1002, , "Tabblad 'CF - opbrengsten' niet gevonden in " & wbF.Name & "."
    stap = "vorige prognose (CF - opbrengsten)"
    n = FOVorig(wsO, wsT, waarschuwing)
    stap = "afronden"
    foStap = ""
    FOSluiten wbF, zelfGeopend, iteratie
    Application.Calculation = rekenmodus
    Application.Calculate
    Application.ScreenUpdating = True
    bericht = "Klaar: vorige prognose (start bouw in blok 4, fee-kwartalen in blok 6) voor " & n & " woningtype(n) uit " & CStr(pad) & "."
    If Len(waarschuwing) > 0 Then bericht = bericht & vbCrLf & vbCrLf & "Let op:" & vbCrLf & waarschuwing
    MsgBox bericht, IIf(Len(waarschuwing) > 0, vbExclamation, vbInformation), "Vorige prognose uit FO"
    Exit Sub
Fout:
    foutNr = Err.Number
    foutOms = Err.Description
    On Error Resume Next
    If rekenmodus <> 0 Then Application.Calculation = rekenmodus
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    FOSluiten wbF, zelfGeopend, iteratie
    bericht = "Het ging mis bij: " & stap
    If Len(foStap) > 0 Then bericht = bericht & vbCrLf & "(" & foStap & ")"
    MsgBox bericht & vbCrLf & vbCrLf & "Fout " & foutNr & ": " & foutOms, vbExclamation, "Vorige prognose uit FO"
End Sub

Public Sub PlanningBewarenAlsVorig(Optional ByVal stil As Boolean = False)
    ' knop 'Nu bewaren als vorige prognose' (tab Woningtypes): de planning die er nu staat (start bouw per type, kwartalen van de
    ' fee-termijnen) naar blok 4 en 6, zodat je daarna de nieuwe planning kunt invullen en het verschil lichtrood in de grafiek ziet
    Dim wsT As Worksheet, k As Long, col As Long, i As Long, n As Long
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    For k = 1 To MAX_TYPES
        col = TYPE_COL1 + (k - 1) * TYPE_BREEDTE
        If Len(TekstVan(wsT.Cells(TYPE_RIJ_NAAM, col).Value)) > 0 Or IsGetal(wsT.Cells(TYPE_RIJ_STARTJAAR, col).Value) Then n = n + 1
    Next k
    If n = 0 Then
        If Not stil Then MsgBox "Er staat nog geen planning op tab Woningtypes.", vbInformation, "Nu bewaren als vorige prognose"
        Exit Sub
    End If
    If Not stil Then If MsgBox("De huidige start bouw per type en de kwartalen van de fee-termijnen worden de vorige prognose (blok 4 en 6). " & _
              "Wat daar nu staat wordt overschreven. Doorgaan?", vbQuestion + vbYesNo, "Nu bewaren als vorige prognose") <> vbYes Then Exit Sub
    For k = 1 To MAX_TYPES
        col = TYPE_COL1 + (k - 1) * TYPE_BREEDTE
        wsT.Cells(TYPE_RIJ_VJAAR, col).Value = wsT.Cells(TYPE_RIJ_STARTJAAR, col).Value
        wsT.Cells(TYPE_RIJ_VKW, col).Value = wsT.Cells(TYPE_RIJ_STARTKW, col).Value
        For i = 0 To TYPE_RIJ_FN - TYPE_RIJ_F1
            wsT.Cells(TYPE_RIJ_VF1 + i, col + 2).Value = wsT.Cells(TYPE_RIJ_F1 + i, col + 2).Value
        Next i
    Next k
    If Not stil Then MsgBox "Klaar: de huidige planning staat als vorige prognose in blok 4 en 6 van tab Woningtypes. Pas nu de planning aan; " & _
           "de bouwtermijnengrafiek toont de vorige prognose lichtrood.", vbInformation, "Nu bewaren als vorige prognose"
End Sub

Private Function FOVorig(wsO As Worksheet, wsT As Worksheet, ByRef waarschuwing As String) As Long
    ' blokken op 'CF - opbrengsten' van een ouder FO -> blok 4 (start bouw) en blok 6 (fee-kwartalen) op tab Woningtypes (ook het
    ' losse DAEB-blok)
    Dim rMax As Long, r As Long, b As String, nBlok As Long, nOmzet As Long, blok() As Long, omzet() As Long, k As Long, kT As Long, kk As Long
    Dim idxKol() As Long, cAct As Long, idxKolO() As Long, cActO As Long, idxStart As Long, rPlan1 As Long, rPlan2 As Long
    Dim naam As String, col As Long, ro As Long, lab As String, ll As String, compRij As Boolean, inFee As Boolean, overslaan As Boolean
    Dim nComp As Long, i As Long, curF As Long, pctD As Double, eurD As Double, alDaeb As Boolean
    Dim cNaam(1 To 2) As String, cTot(1 To 2) As Double, stil As String
    Dim aN(1 To LIJST_MAX) As String, aW(1 To LIJST_MAX) As Double, aK(1 To LIJST_MAX) As Variant, nA As Long
    Dim bN(1 To LIJST_MAX) As String, bW(1 To LIJST_MAX) As Double, bK(1 To LIJST_MAX) As Variant, nB As Long
    ReDim blok(1 To MAX_TYPES)
    ReDim omzet(1 To MAX_TYPES)
    foStap = "blokken zoeken op '" & wsO.Name & "'"
    rMax = wsO.Cells(wsO.Rows.Count, 2).End(xlUp).Row
    For r = 2 To rMax
        b = LCase$(TekstVan(wsO.Cells(r, 2).Value))
        If Left$(b, 10) = "termijnen " And InStr(b, " in " & ChrW(8364)) = 0 And LCase$(Left$(TekstVan(wsO.Cells(r - 1, 6).Value), 5)) = "# won" Then
            If nBlok < MAX_TYPES Then
                nBlok = nBlok + 1
                blok(nBlok) = r
            End If
        ElseIf Left$(b, 9) = "omzet hvg" Or Left$(b, 15) = "omzet termijnen" Then
            If nOmzet < MAX_TYPES Then
                nOmzet = nOmzet + 1
                omzet(nOmzet) = r
            End If
        End If
    Next r
    kT = 0
    For k = 1 To nBlok
        If Getal(wsO.Cells(blok(k) - 1, 7).Value) <= 0 Then GoTo Volgende
        kT = kT + 1
        naam = typeNaam(wsO, blok(k), kT)
        foStap = "vorige prognose: '" & naam & "' (blok vanaf rij " & blok(k) & " van '" & wsO.Name & "')"
        ' doelkolom: het type met dezelfde naam, anders het type op dezelfde positie
        col = 0
        For kk = 1 To MAX_TYPES
            If LCase$(TekstVan(wsT.Cells(TYPE_RIJ_NAAM, TYPE_COL1 + (kk - 1) * TYPE_BREEDTE).Value)) = LCase$(naam) Then col = TYPE_COL1 + (kk - 1) * TYPE_BREEDTE
        Next kk
        If col = 0 Then
            col = TYPE_COL1 + (kT - 1) * TYPE_BREEDTE
            If Len(TekstVan(wsT.Cells(TYPE_RIJ_NAAM, col).Value)) > 0 Then
                waarschuwing = waarschuwing & "- '" & naam & "' uit het vorige FO staat niet op tab Woningtypes; type " & kT & " ('" & TekstVan(wsT.Cells(TYPE_RIJ_NAAM, col).Value) & "') op dezelfde positie gebruikt." & vbCrLf
            Else
                waarschuwing = waarschuwing & "- '" & naam & "' uit het vorige FO staat niet op tab Woningtypes; overgeslagen." & vbCrLf
                GoTo Volgende
            End If
        End If
        ' start bouw: eerste telling in de bouwplanning
        KolomKaart wsO, blok(k), idxKol, cAct
        rPlan1 = blok(k) + 1
        rPlan2 = rPlan1
        Do While Len(TekstVan(wsO.Cells(rPlan2 + 1, 2).Value)) > 0
            rPlan2 = rPlan2 + 1
        Loop
        idxStart = 0
        For r = rPlan1 To rPlan2
            idxStart = EersteIdx(wsO, r, idxKol)
            If idxStart > 0 Then Exit For
        Next r
        wsT.Range(wsT.Cells(TYPE_RIJ_VJAAR, col), wsT.Cells(TYPE_RIJ_VKW, col)).ClearContents
        If idxStart > 0 Then
            wsT.Cells(TYPE_RIJ_VJAAR, col).Value = (idxStart - 1) \ 4
            wsT.Cells(TYPE_RIJ_VKW, col).Value = idxStart - ((idxStart - 1) \ 4) * 4
        End If
        ' fees: de termijnen van het oude FO zoals bij het ophalen (FeeDelen, FeesBinnenLimiet), dan per rij van blok 5 het kwartaal
        ' op mijlpaal (KwVorig, zoals fo.voeg_vorige_prognose_toe)
        If k <= nOmzet Then
            ro = omzet(k)
            KolomKaart wsO, ro, idxKolO, cActO
            inFee = False
            nComp = 0
            curF = 0
            nA = 0
            nB = 0
            r = ro + 1
            Do While Len(TekstVan(wsO.Cells(r, 2).Value)) > 0
                lab = TekstVan(wsO.Cells(r, 2).Value)
                ll = LCase$(lab)
                compRij = (SomRij(wsO, r, idxKolO) = 0 And ActWaarde(wsO, r, cActO) = 0 And Getal(wsO.Cells(r, 4).Value) <> 0 And Getal(wsO.Cells(r, 3).Value) = 0)
                If ll = "koopsom" Then
                    ' niets
                ElseIf compRij And (IsFeeLabel(ll) Or inFee) Then
                    inFee = True
                    curF = 0
                    If Left$(ll, 10) <> "onvoorzien" And nComp < 2 Then
                        nComp = nComp + 1
                        curF = nComp
                        alDaeb = True
                    End If
                ElseIf inFee And (Getal(wsO.Cells(r, 3).Value) <> 0 Or Getal(wsO.Cells(r, 4).Value) <> 0) And (SomRij(wsO, r, idxKolO) <> 0 Or ActWaarde(wsO, r, cActO) <> 0) Then
                    pctD = 0
                    eurD = 0
                    If IsGetal(wsO.Cells(r, 3).Value) Then pctD = CDbl(wsO.Cells(r, 3).Value)
                    If IsGetal(wsO.Cells(r, 4).Value) Then eurD = CDbl(wsO.Cells(r, 4).Value)
                    If curF = 1 Then
                        FeeDelen wsO, r, idxKolO, cActO, lab, pctD, eurD, aN, aW, aK, nA, stil
                    ElseIf curF = 2 Then
                        FeeDelen wsO, r, idxKolO, cActO, lab, pctD, eurD, bN, bW, bK, nB, stil
                    End If
                End If
                r = r + 1
            Loop
            If nComp > 0 Then
                FeesBinnenLimiet aN, aW, aK, nA, 5, naam, "fee-termijnen", stil
                FeesBinnenLimiet bN, bW, bK, nB, 5, naam, "fee-termijnen", stil
                KwVorig wsT, col, aN, aK, nA, bN, bK, nB
            End If
        End If
        FOVorig = FOVorig + 1
Volgende:
    Next k
    ' los DAEB-blok: start bouw (blok 4) en per fee-termijn het kwartaal (blok 6); op mijlpaal, anders op positie (zoals
    ' fo.voeg_vorige_prognose_toe); de meldingen over het oude FO zelf blijven weg
    If LeesDaebBlok(wsO, naam, idxStart, cNaam, cTot, nComp, aN, aW, aK, nA, bN, bW, bK, nB, alDaeb, stil) > 0 Then
        foStap = "vorige prognose: DAEB-blok '" & naam & "' van '" & wsO.Name & "'"
        col = 0
        For kk = 1 To MAX_TYPES
            If LCase$(TekstVan(wsT.Cells(TYPE_RIJ_NAAM, TYPE_COL1 + (kk - 1) * TYPE_BREEDTE).Value)) = LCase$(naam) Then col = TYPE_COL1 + (kk - 1) * TYPE_BREEDTE
        Next kk
        If col = 0 And kT < MAX_TYPES Then
            col = TYPE_COL1 + kT * TYPE_BREEDTE
            If Len(TekstVan(wsT.Cells(TYPE_RIJ_NAAM, col).Value)) > 0 Then
                waarschuwing = waarschuwing & "- '" & naam & "' uit het vorige FO staat niet op tab Woningtypes; type " & (kT + 1) & " ('" & TekstVan(wsT.Cells(TYPE_RIJ_NAAM, col).Value) & "') op dezelfde positie gebruikt." & vbCrLf
            Else
                waarschuwing = waarschuwing & "- '" & naam & "' uit het vorige FO staat niet op tab Woningtypes; overgeslagen." & vbCrLf
                col = 0
            End If
        End If
        If col > 0 Then
            wsT.Range(wsT.Cells(TYPE_RIJ_VJAAR, col), wsT.Cells(TYPE_RIJ_VKW, col)).ClearContents
            If idxStart > 0 Then
                wsT.Cells(TYPE_RIJ_VJAAR, col).Value = (idxStart - 1) \ 4
                wsT.Cells(TYPE_RIJ_VKW, col).Value = idxStart - ((idxStart - 1) \ 4) * 4
            End If
            KwVorig wsT, col, aN, aK, nA, bN, bK, nB
            FOVorig = FOVorig + 1
        End If
    End If
    foStap = ""
End Function

Private Function BasisFee(lab As String, ByRef nr As Long) As String
    ' 'na akkoord do (2/2)' -> 'na akkoord do', nr 2; 'x (ontvangen)' -> 'x', nr 0; anders het label, nr -1 (zoals fo.py _basis_fee)
    Dim b As String
    nr = -1
    BasisFee = lab
    If Len(lab) > 13 And Right$(lab, 12) = " (ontvangen)" Then
        BasisFee = Left$(lab, Len(lab) - 12)
        nr = 0
        Exit Function
    End If
    b = BasisNaam(lab)
    If Len(b) > 0 Then
        BasisFee = b
        nr = CLng(Val(Mid$(lab, Len(b) + 3)))
    End If
End Function

Private Sub KwVorig(wsT As Worksheet, col As Long, aN() As String, aK() As Variant, nA As Long, bN() As String, bK() As Variant, nB As Long)
    ' blok 6 van één type uit de fee-termijnen van het oude FO (per component aN../bN..; zoals fo.py _kw_vorig): per rij van blok 5,
    ' binnen dezelfde component, eerst het zelfde label, dan dezelfde mijlpaal zonder '(i/n)'/'(ontvangen)' (het deel met hetzelfde
    ' volgnummer, anders het eerste), dan de zelfde positie; alleen als blok 5 mijlpalen heeft
    Dim c As Long, i As Long, j As Long, lab As String, oudN(1 To 5) As String, oudK(1 To 5) As Variant, kw As Variant, heeft As Boolean
    Dim b As String, nr As Long, nr2 As Long, eerste As Long, zelfde As Long
    For i = 0 To TYPE_RIJ_FN - TYPE_RIJ_F1
        If Len(TekstVan(wsT.Cells(TYPE_RIJ_F1 + i, col).Value)) > 0 Then heeft = True
    Next i
    If Not heeft Then Exit Sub
    wsT.Range(wsT.Cells(TYPE_RIJ_VF1, col + 2), wsT.Cells(TYPE_RIJ_VFN, col + 2)).ClearContents
    For c = 1 To 2
        For i = 1 To 5
            oudN(i) = ""
            oudK(i) = Empty
            If c = 1 And i <= nA Then
                oudN(i) = LCase$(Trim$(aN(i)))
                oudK(i) = aK(i)
            ElseIf c = 2 And i <= nB Then
                oudN(i) = LCase$(Trim$(bN(i)))
                oudK(i) = bK(i)
            End If
        Next i
        For i = 1 To 5
            lab = LCase$(TekstVan(wsT.Cells(TYPE_RIJ_F1 + (c - 1) * 5 + i - 1, col).Value))
            If Len(lab) > 0 Then
                kw = Empty
                For j = 1 To 5
                    If Len(oudN(j)) > 0 And oudN(j) = lab Then
                        kw = oudK(j)
                        Exit For
                    End If
                Next j
                If IsEmpty(kw) Then
                    b = BasisFee(lab, nr)
                    eerste = 0
                    zelfde = 0
                    For j = 1 To 5
                        If Len(oudN(j)) > 0 Then
                            If BasisFee(oudN(j), nr2) = b Then
                                If eerste = 0 Then eerste = j
                                If zelfde = 0 And nr2 = nr Then zelfde = j
                            End If
                        End If
                    Next j
                    If zelfde > 0 Then
                        kw = oudK(zelfde)
                    ElseIf eerste > 0 Then
                        kw = oudK(eerste)
                    End If
                End If
                If IsEmpty(kw) And Len(oudN(i)) > 0 Then kw = oudK(i)
                If Not IsEmpty(kw) Then wsT.Cells(TYPE_RIJ_VF1 + (c - 1) * 5 + i - 1, col + 2).Value = kw
            End If
        Next i
    Next c
End Sub

Private Function ZoekBlad(wb As Workbook, naam As String) As Worksheet
    ' werkblad op naam, hoofdletters en spaties genegeerd
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Replace(LCase$(ws.Name), " ", "") = Replace(LCase$(naam), " ", "") Then
            Set ZoekBlad = ws
            Exit Function
        End If
    Next ws
End Function

Private Function IsJaar(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    IsJaar = (Val(v) >= 1990 And Val(v) <= 2100)
End Function

Private Function KwNummer(v As Variant) As Long
    ' 'Q1'..'Q4' of 1..4 -> 1..4; anders 0
    Dim t As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    t = UCase$(Replace(CStr(v), " ", ""))
    If Left$(t, 1) = "Q" Then t = Mid$(t, 2)
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Then KwNummer = CLng(t)
End Function

Private Function Getal(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Getal = CDbl(v)
End Function

Private Function IsGetal(v As Variant) As Boolean
    ' echte getallen (geen lege cel, tekst, datum of foutwaarde)
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: IsGetal = True
    End Select
End Function

Private Function typeNaam(ws As Worksheet, r As Long, k As Long) As String
    ' naam van het type: uit de bloktitel 'Termijnen k - naam', anders de cel naast de titel (bv. '<300K VON'),
    ' anders de tekst na 'Termijnen ' als dat geen nummer is, anders 'Type k'
    Dim t As String, rest As String
    t = TekstVan(ws.Cells(r, 2).Value)
    If InStr(t, " - ") > 0 Then
        typeNaam = Trim$(Mid$(t, InStr(t, " - ") + 3))
    ElseIf Len(TekstVan(ws.Cells(r, 3).Value)) > 0 And Not IsGetal(ws.Cells(r, 3).Value) Then
        typeNaam = TekstVan(ws.Cells(r, 3).Value)
    Else
        rest = Trim$(Mid$(t, 11))
        If Len(rest) > 0 And Not IsNumeric(rest) Then typeNaam = rest
    End If
    If Len(typeNaam) = 0 Then typeNaam = "Type " & k
End Function

Private Function TekstVan(v As Variant) As String
    ' celwaarde als tekst zonder spaties eromheen; foutwaarden (#N/A, #REF!) en lege cellen geven ""
    If IsError(v) Or IsEmpty(v) Then Exit Function
    TekstVan = Trim$(CStr(v))
End Function

Private Function IsFeeLabel(ll As String) As Boolean
    ' label (kleine letters) van een DAEB-fee: het woord 'fee'/'fees', 'bijkomende kosten' of 'onvoorzien'
    Dim t As String
    t = " " & Replace(Replace(Replace(Replace(Replace(ll, "-", " "), "(", " "), ")", " "), "/", " "), ":", " ") & " "
    IsFeeLabel = (InStr(t, " fee ") > 0) Or (InStr(t, " fees ") > 0) Or (Left$(ll, 17) = "bijkomende kosten") Or (Left$(ll, 10) = "onvoorzien")
End Function

Private Sub KolomKaart(ws As Worksheet, rijQ As Long, ByRef idxKol() As Long, ByRef cAct As Long)
    ' idxKol(c) = jaar*4+kw van kolom c (0 = geen kwartaalkolom), uit de rij met Q1..Q4 en de jaartallen in de rij erboven;
    ' cAct = kolom 'actuals' (0 = geen)
    Dim c As Long, jaar As Long, q As Long, nKol As Long
    nKol = ws.UsedRange.Columns.Count + ws.UsedRange.Column
    ReDim idxKol(1 To nKol)
    cAct = 0
    jaar = 0
    For c = 1 To nKol
        If rijQ > 1 Then
            If IsJaar(ws.Cells(rijQ - 1, c).Value) Then jaar = CLng(Val(ws.Cells(rijQ - 1, c).Value))
        End If
        If LCase$(Left$(TekstVan(ws.Cells(rijQ, c).Value), 6)) = "actual" Then cAct = c
        q = KwNummer(ws.Cells(rijQ, c).Value)
        If q > 0 And jaar > 0 Then idxKol(c) = jaar * 4 + q Else idxKol(c) = 0
    Next c
End Sub

Private Function EersteIdx(ws As Worksheet, r As Long, idxKol() As Long) As Long
    ' kwartaalindex van de eerste kwartaalcel in rij r met een getal <> 0 (0 = geen)
    Dim c As Long
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            If Getal(ws.Cells(r, c).Value) <> 0 Then
                EersteIdx = idxKol(c)
                Exit Function
            End If
        End If
    Next c
End Function

Private Function ActWaarde(ws As Worksheet, r As Long, cAct As Long) As Double
    ' waarde in de kolom 'actuals' van rij r (0 als die kolom er niet is)
    If cAct > 0 Then ActWaarde = Getal(ws.Cells(r, cAct).Value)
End Function

Private Function SomRij(ws As Worksheet, r As Long, idxKol() As Long) As Double
    Dim c As Long
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then SomRij = SomRij + Getal(ws.Cells(r, c).Value)
    Next c
End Function

Private Sub VoegToe(lNaam() As String, lW() As Double, lKw() As Variant, ByRef nL As Long, naam As String, waarde As Double, kw As Variant)
    If nL >= UBound(lNaam) Then Exit Sub
    nL = nL + 1
    lNaam(nL) = naam
    lW(nL) = waarde
    lKw(nL) = kw
End Sub

Private Function FeeDelen(ws As Worksheet, r As Long, idxKol() As Long, cAct As Long, lab As String, pct As Double, _
                          eur As Double, lNaam() As String, lW() As Double, lKw() As Variant, ByRef nL As Long, _
                          ByRef waarschuwing As String) As Long
    ' één fee-termijn -> regels voor blok 5 (zelfde regels als fo.py _fee_delen): de bedragen komen uit de cellen; de kolom actuals
    ' is één deel 'actuals' ('x (ontvangen)' als er ook kwartalen zijn), elk kwartaal een eigen deel met zijn kalenderkwartaal
    ' ('x (1/2)' ...). Een deel van 1 euro of minder (correctie of afronding, ook negatief) gaat op in het vorige deel (anders het
    ' volgende). Tellen de cellen samen op tot 1 of minder (alleen een teken), dan het % of EUR van de rij, naar rato.
    ' Geeft het aantal toegevoegde regels (0 = geen bedrag in een kwartaal of in actuals).
    Dim c As Long, v As Double, som As Double, nD As Long, nLater As Long, i As Long, j As Long, k As Long, basis As Double, opRato As Boolean
    Dim dKw() As Long, dV() As Double
    ReDim dKw(1 To UBound(idxKol) + 1)
    ReDim dV(1 To UBound(idxKol) + 1)
    If cAct > 0 Then
        If IsGetal(ws.Cells(r, cAct).Value) Then
            If CDbl(ws.Cells(r, cAct).Value) <> 0 Then
                nD = 1
                dKw(1) = -1                                  ' -1 = actuals
                dV(1) = CDbl(ws.Cells(r, cAct).Value)
            End If
        End If
    End If
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            If IsGetal(ws.Cells(r, c).Value) Then
                v = CDbl(ws.Cells(r, c).Value)
                If v <> 0 Then
                    nD = nD + 1
                    dKw(nD) = idxKol(c)
                    dV(nD) = v
                End If
            End If
        End If
    Next c
    If nD = 0 Then Exit Function
    For i = 1 To nD
        som = som + dV(i)
    Next i
    opRato = (Abs(som) <= 1 And som <> 0 And (pct <> 0 Or eur <> 0))
    If Not opRato Then
        Do While nD > 1
            i = 0
            For k = 1 To nD
                If dV(k) <= 1 Then
                    i = k
                    Exit For
                End If
            Next k
            If i = 0 Then Exit Do
            If i > 1 Then j = i - 1 Else j = i + 1
            waarschuwing = waarschuwing & "- " & lab & ": " & Round(dV(i), 2) & " in " & IIf(dKw(i) = -1, "actuals", KwTekst(dKw(i))) & _
                           " (correctie of afronding) is opgeteld bij " & IIf(dKw(j) = -1, "actuals", KwTekst(dKw(j))) & "." & vbCrLf
            dV(j) = dV(j) + dV(i)
            For k = i To nD - 1
                dKw(k) = dKw(k + 1)
                dV(k) = dV(k + 1)
            Next k
            nD = nD - 1
        Loop
    End If
    If pct <> 0 Then basis = pct Else basis = eur
    For k = 1 To nD
        If dKw(k) <> -1 Then nLater = nLater + 1
    Next k
    i = 0
    For k = 1 To nD
        If dKw(k) = -1 Then
            VoegToe lNaam, lW, lKw, nL, lab & IIf(nLater > 0, " (ontvangen)", ""), FeeWaarde(dV(k), opRato, basis, som), "actuals"
        Else
            i = i + 1
            VoegToe lNaam, lW, lKw, nL, lab & IIf(nLater > 1, " (" & i & "/" & nLater & ")", ""), FeeWaarde(dV(k), opRato, basis, som), KwTekst(dKw(k))
        End If
        FeeDelen = FeeDelen + 1
    Next k
End Function

Private Function FeeWaarde(v As Double, opRato As Boolean, basis As Double, som As Double) As Double
    ' bedrag op centen, of (alleen een teken in de cellen) het % of EUR van de rij naar rato op 6 decimalen
    If opRato Then FeeWaarde = Round(basis * v / som, 6) Else FeeWaarde = Round(v, 2)
End Function

Private Sub FeesBinnenLimiet(lNaam() As String, lW() As Double, lKw() As Variant, ByRef nL As Long, nMax As Long, typeNaam As String, _
                             wat As String, ByRef waarschuwing As String)
    ' fee-termijnen van één component terugbrengen tot nMax regels zonder bedrag te verliezen (zoals fo.py _fees_binnen_limiet):
    ' eerst alles wat al ontvangen is in één regel ('ontvangen t/m actuals'), dan regels in hetzelfde kwartaal samen ('a + b'),
    ' dan gesplitste delen samenvoegen en de rest in de laatste regel (BinnenLimiet)
    Dim i As Long, j As Long, k As Long, eerste As Long, nAct As Long, samen As Boolean, gevonden As Boolean
    If nL <= nMax Then Exit Sub
    For i = 1 To nL
        If CStr(lKw(i)) = "actuals" Then nAct = nAct + 1
    Next i
    If nAct > 1 Then
        eerste = 0
        i = 1
        Do While i <= nL
            If CStr(lKw(i)) = "actuals" Then
                If eerste = 0 Then
                    eerste = i
                    i = i + 1
                Else
                    lW(eerste) = lW(eerste) + lW(i)
                    For k = i To nL - 1
                        lNaam(k) = lNaam(k + 1)
                        lW(k) = lW(k + 1)
                        lKw(k) = lKw(k + 1)
                    Next k
                    nL = nL - 1
                End If
            Else
                i = i + 1
            End If
        Loop
        lNaam(eerste) = "ontvangen t/m actuals"
        lW(eerste) = Round(lW(eerste), 6)
        waarschuwing = waarschuwing & "- " & typeNaam & ": meer dan " & nMax & " " & wat & "; wat al ontvangen is staat in één rij ('ontvangen t/m actuals')." & vbCrLf
    End If
    Do While nL > nMax
        gevonden = False
        For i = 1 To nL - 1
            For j = i + 1 To nL
                If Not IsEmpty(lKw(i)) And CStr(lKw(i)) = CStr(lKw(j)) Then
                    gevonden = True
                    Exit For
                End If
            Next j
            If gevonden Then Exit For
        Next i
        If Not gevonden Then Exit Do
        lNaam(i) = lNaam(i) & " + " & lNaam(j)
        lW(i) = Round(lW(i) + lW(j), 6)
        For k = j To nL - 1
            lNaam(k) = lNaam(k + 1)
            lW(k) = lW(k + 1)
            lKw(k) = lKw(k + 1)
        Next k
        nL = nL - 1
        samen = True
    Loop
    If samen Then waarschuwing = waarschuwing & "- " & typeNaam & ": meer dan " & nMax & " " & wat & "; termijnen in hetzelfde kwartaal staan in één rij." & vbCrLf
    BinnenLimiet lNaam, lW, lKw, nL, nMax, typeNaam, wat, waarschuwing, True
End Sub

Private Function EnkeleSpaties(s As String) As String
    ' spatie, tab, regeleinde en harde spatie worden één spatie (zoals fo.py _woorden)
    EnkeleSpaties = Replace(Replace(Replace(Replace(s, Chr(9), " "), Chr(10), " "), Chr(13), " "), Chr(160), " ")
    EnkeleSpaties = Trim$(EnkeleSpaties)
    Do While InStr(EnkeleSpaties, "  ") > 0
        EnkeleSpaties = Replace(EnkeleSpaties, "  ", " ")
    Loop
End Function

Private Function SleutelWoord(s As String) As String
    ' het eerste woord (kleine letters) dat niet 'fee'/'fees' is; zonder zo'n woord het eerste woord (zoals fo.py _sleutel)
    Dim w() As String, i As Long
    w = Split(EnkeleSpaties(s), " ")
    For i = 0 To UBound(w)
        If LCase$(w(i)) <> "fee" And LCase$(w(i)) <> "fees" Then
            SleutelWoord = LCase$(w(i))
            Exit Function
        End If
    Next i
    If UBound(w) >= 0 Then SleutelWoord = LCase$(w(0))
End Function

Private Function ZonderVoorvoegsel(lab As String, compNaam As String) As String
    ' 'AK fee na akkoord SO' bij component 'AK' -> 'na akkoord SO' (zoals fo.py _zonder_voorvoegsel): de woorden vooraan die ook in de
    ' componentnaam staan (of 'fee'/'fees') vallen weg, en een scheidingsteken ('-', ':') daarna ook; blijft er niets over (of valt
    ' er niets weg), dan het label zoals het is
    Dim w() As String, namen As String, i As Long, j As Long, t As String
    namen = " " & LCase$(EnkeleSpaties(compNaam)) & " fee fees "
    w = Split(EnkeleSpaties(lab), " ")
    i = 0
    Do While i <= UBound(w)
        If InStr(namen, " " & LCase$(w(i)) & " ") = 0 Then Exit Do
        i = i + 1
    Loop
    If i > 0 And i <= UBound(w) Then
        For j = i To UBound(w)
            t = t & IIf(Len(t) > 0, " ", "") & w(j)
        Next j
        Do While Len(t) > 0
            If InStr("-:+ " & ChrW(8211), Left$(t, 1)) = 0 Then Exit Do
            t = Mid$(t, 2)
        Loop
    End If
    If Len(t) > 0 Then ZonderVoorvoegsel = t Else ZonderVoorvoegsel = lab
End Function

Private Function LeesDaebBlok(wsO As Worksheet, ByRef naam As String, ByRef idxStart As Long, cNaam() As String, _
                              cTot() As Double, ByRef nComp As Long, aN() As String, aW() As Double, aK() As Variant, ByRef nA As Long, _
                              bN() As String, bW() As Double, bK() As Variant, ByRef nB As Long, alDaeb As Boolean, _
                              ByRef waarschuwing As String) As Long
    ' los DAEB-blok op 'CF - opbrengsten' (zelfde regels als fo.py _lees_daeb): een titel met 'DAEB' (bv. 'Opbrengsten DAEB') op een
    ' koprij met 'actuals' en Q1..Q4 (jaartallen erboven), geen typeblok of omzetblok; boven de koprij per fee-component de naam
    ' (kolom C) en het totaal (kolom E, anders D), rijen r0-3..r0-1 ('Onvoorzien' doet niet mee); daaronder per termijn een rij,
    ' groepen gescheiden door een lege rij, tot 'Totaal ...' of drie lege rijen (een subtotaal telt niet mee). Een termijn hoort
    ' bij de component met hetzelfde sleutelwoord (SleutelWoord), anders bij de component van zijn groep; de woorden die al in de
    ' componentnaam staan vallen weg uit de mijlpaal. Start bouw: 'Startbouw' boven een kwartaalkolom. Het eerste blok met
    ' componenten telt; alDaeb (er is al een DAEB-type met fees uit een omzetblok): alleen een melding. Geeft de koprij (0 = geen).
    ' cNaam/cTot(1..2) en per component de termijnen (aN.. en bN.., al binnen de vijf rijen van blok 5)
    Dim rMax As Long, r As Long, rr As Long, b As String, idxKol() As Long, cAct As Long, c As Long, nQ As Long, r0 As Long
    Dim lab As String, ll As String, tot As Double, titel As String, groep As Long, leeg As Long, inGroep As Boolean, ci As Long
    Dim i As Long, pct As Double, eur As Double, w() As String, nD As Long, mijl As String, nC As Long, nTreffer As Long
    Dim cn(1 To 3) As String, ct(1 To 3) As Double, co(1 To 3) As Boolean, sl(1 To 3) As String, isSub As Boolean, sw As String
    Dim kn(1 To 3) As String, kT(1 To 3) As Double, ko(1 To 3) As Boolean, kc As Long, echt As Boolean, eersteKand As Long
    Dim xN(1 To LIJST_MAX) As String, xW(1 To LIJST_MAX) As Double, xK(1 To LIJST_MAX) As Variant, nX As Long
    nA = 0
    nB = 0
    nComp = 0
    idxStart = 0
    naam = ""
    rMax = wsO.Cells(wsO.Rows.Count, 2).End(xlUp).Row
    For r = 2 To rMax
        b = LCase$(TekstVan(wsO.Cells(r, 2).Value))
        If InStr(b, "daeb") > 0 And Left$(b, 9) <> "omzet hvg" And Left$(b, 15) <> "omzet termijnen" Then
            If Not (Left$(b, 10) = "termijnen " And LCase$(Left$(TekstVan(wsO.Cells(r - 1, 6).Value), 5)) = "# won") Then
                KolomKaart wsO, r, idxKol, cAct
                nQ = 0
                For c = LBound(idxKol) To UBound(idxKol)
                    If idxKol(c) > 0 Then nQ = nQ + 1
                Next c
                If nQ >= 4 Then
                    If eersteKand = 0 Then eersteKand = r
                    ' componenten boven deze koprij
                    kc = 0
                    echt = False
                    For rr = IIf(r - 3 < 1, 1, r - 3) To r - 1
                        lab = TekstVan(wsO.Cells(rr, 3).Value)
                        tot = 0
                        If IsGetal(wsO.Cells(rr, 5).Value) Then
                            tot = CDbl(wsO.Cells(rr, 5).Value)
                        ElseIf IsGetal(wsO.Cells(rr, 4).Value) Then
                            tot = CDbl(wsO.Cells(rr, 4).Value)
                        End If
                        If Len(lab) > 0 And Not IsGetal(wsO.Cells(rr, 3).Value) And tot <> 0 Then
                            kc = kc + 1
                            kn(kc) = lab
                            kT(kc) = Round(tot, 2)
                            ko(kc) = (LCase$(Left$(lab, 10)) = "onvoorzien")
                            If Not ko(kc) Then echt = True
                        End If
                    Next rr
                    If echt Then
                        If r0 = 0 Then
                            r0 = r
                            nC = kc
                            For i = 1 To kc
                                cn(i) = kn(i)
                                ct(i) = kT(i)
                                co(i) = ko(i)
                            Next i
                        Else
                            waarschuwing = waarschuwing & "- '" & wsO.Name & "': nog een DAEB-blok '" & TekstVan(wsO.Cells(r, 2).Value) & "' (rij " & r & "); alleen het eerste ('" & TekstVan(wsO.Cells(r0, 2).Value) & "', rij " & r0 & ") is overgenomen." & vbCrLf
                        End If
                    End If
                End If
            End If
        End If
    Next r
    If r0 = 0 Then
        If eersteKand > 0 Then waarschuwing = waarschuwing & "- '" & wsO.Name & "': blok '" & TekstVan(wsO.Cells(eersteKand, 2).Value) & "' (rij " & eersteKand & ") zonder fee-componenten (naam in kolom C en totaal in kolom E boven de koprij): DAEB-fees niet overgenomen." & vbCrLf
        Exit Function
    End If
    titel = TekstVan(wsO.Cells(r0, 2).Value)
    If alDaeb Then
        waarschuwing = waarschuwing & "- '" & wsO.Name & "': het blok '" & titel & "' (rij " & r0 & ") is niet overgenomen: er is al een DAEB-type met fees uit een omzetblok." & vbCrLf
        Exit Function
    End If
    KolomKaart wsO, r0, idxKol, cAct
    For i = 1 To nC
        sl(i) = SleutelWoord(cn(i))
    Next i
    groep = 1
    For r = r0 + 1 To r0 + 40
        lab = TekstVan(wsO.Cells(r, 2).Value)
        If Len(lab) = 0 Then
            leeg = leeg + 1
            If leeg >= 3 Then Exit For
            If inGroep Then
                groep = groep + 1
                inGroep = False
            End If
        Else
            leeg = 0
            ll = LCase$(lab)
            ' subtotaal van een component ('Subtotaal ...', 'Totaal AK fee') telt niet; 'Totaal ...' sluit het blok af
            isSub = (Left$(ll, 9) = "subtotaal")
            If Not isSub And Left$(ll, 6) = "totaal" Then
                sw = " " & LCase$(Mid$(EnkeleSpaties(lab), InStr(EnkeleSpaties(lab) & " ", " ") + 1)) & " "
                For i = 1 To nC
                    If InStr(sw, " " & sl(i) & " ") > 0 Then isSub = True
                Next i
                If Not isSub Then Exit For
            End If
            If Not isSub Then
                inGroep = True
                nTreffer = 0
                ci = 0
                For i = 1 To nC
                    If sl(i) = SleutelWoord(lab) Then
                        nTreffer = nTreffer + 1
                        ci = i
                    End If
                Next i
                If nTreffer <> 1 Then
                    If groep <= nC Then ci = groep Else ci = 0
                End If
                If ci = 0 Then
                    waarschuwing = waarschuwing & "- '" & wsO.Name & "': rij '" & lab & "' in het blok '" & titel & "' hoort bij geen fee-component en is overgeslagen." & vbCrLf
                ElseIf Not co(ci) Then
                    pct = 0
                    eur = 0
                    If IsGetal(wsO.Cells(r, 3).Value) Then pct = CDbl(wsO.Cells(r, 3).Value)
                    If IsGetal(wsO.Cells(r, 4).Value) Then eur = CDbl(wsO.Cells(r, 4).Value)
                    mijl = ZonderVoorvoegsel(lab, cn(ci))
                    ' de echte componenten (zonder 'Onvoorzien') in volgorde: de eerste naar aN.., de tweede naar bN.., de rest vervalt
                    kc = 0
                    For i = 1 To ci
                        If Not co(i) Then kc = kc + 1
                    Next i
                    Select Case kc
                        Case 1: nD = FeeDelen(wsO, r, idxKol, cAct, mijl, pct, eur, aN, aW, aK, nA, waarschuwing)
                        Case 2: nD = FeeDelen(wsO, r, idxKol, cAct, mijl, pct, eur, bN, bW, bK, nB, waarschuwing)
                        Case Else: nD = FeeDelen(wsO, r, idxKol, cAct, mijl, pct, eur, xN, xW, xK, nX, waarschuwing)
                    End Select
                    If nD = 0 And (pct <> 0 Or eur <> 0) Then
                        waarschuwing = waarschuwing & "- '" & wsO.Name & "': rij '" & lab & "' in het blok '" & titel & "' heeft geen bedrag in een kwartaal of in 'actuals' en is overgeslagen." & vbCrLf
                    End If
                End If
            End If
        End If
    Next r
    ' naam: de titel zonder 'opbrengsten', 'omzet', 'fee', 'fees'
    w = Split(EnkeleSpaties(titel), " ")
    For i = 0 To UBound(w)
        Select Case LCase$(w(i))
            Case "opbrengsten", "omzet", "fee", "fees"
            Case Else
                naam = naam & IIf(Len(naam) > 0, " ", "") & w(i)
        End Select
    Next i
    If Len(naam) = 0 Then naam = "DAEB"
    For rr = IIf(r0 - 3 < 1, 1, r0 - 3) To r0 - 1
        For c = LBound(idxKol) To UBound(idxKol)
            If idxKol(c) > 0 Then
                If Left$(Replace(LCase$(TekstVan(wsO.Cells(rr, c).Value)), " ", ""), 9) = "startbouw" Then
                    idxStart = idxKol(c)
                    Exit For
                End If
            End If
        Next c
        If idxStart > 0 Then Exit For
    Next rr
    If idxStart = 0 Then waarschuwing = waarschuwing & "- " & naam & ": geen 'Startbouw' boven een kwartaal van het blok '" & titel & "'; vul start bouw in op tab Woningtypes." & vbCrLf
    kc = 0
    For i = 1 To nC
        If Not co(i) Then
            kc = kc + 1
            If kc <= 2 Then
                cNaam(kc) = cn(i)
                cTot(kc) = ct(i)
            End If
        End If
    Next i
    If kc > 2 Then waarschuwing = waarschuwing & "- " & naam & ": " & kc & " fee-componenten in het FO, alleen de eerste 2 zijn overgenomen." & vbCrLf
    nComp = IIf(kc > 2, 2, kc)
    FeesBinnenLimiet aN, aW, aK, nA, 5, naam, "fee-termijnen van '" & cNaam(1) & "'", waarschuwing
    If nComp >= 2 Then FeesBinnenLimiet bN, bW, bK, nB, 5, naam, "fee-termijnen van '" & cNaam(2) & "'", waarschuwing
    LeesDaebBlok = r0
End Function

Private Function BasisNaam(lab As String) As String
    ' 'Casco (2/3)' -> 'Casco'; zonder achtervoegsel (i/n): ""
    Dim p As Long, t As String, s As Long
    p = InStrRev(lab, " (")
    If p = 0 Or Right$(lab, 1) <> ")" Then Exit Function
    t = Mid$(lab, p + 2, Len(lab) - p - 2)
    s = InStr(t, "/")
    If s < 2 Or s = Len(t) Then Exit Function
    If Not IsNumeric(Left$(t, s - 1)) Or Not IsNumeric(Mid$(t, s + 1)) Then Exit Function
    If InStr(t, ".") > 0 Or InStr(t, ",") > 0 Or InStr(t, " ") > 0 Then Exit Function
    BasisNaam = Left$(lab, p - 1)
End Function

Private Sub BinnenLimiet(lNaam() As String, lW() As Double, lKw() As Variant, ByRef nL As Long, nMax As Long, typeNaam As String, _
                         wat As String, ByRef waarschuwing As String, laatste As Boolean)
    ' de lijst terugbrengen tot nMax regels, zoals fo.py _binnen_limiet: van de termijn met de meeste delen (bij gelijk aantal de
    ' eerste) gaat het kleinste deel op in het grootste deel (kwartaal van het grootste), daarna nummert hij de delen opnieuw; zonder
    ' gesplitste termijnen om samen te voegen vallen de laatste regels weg (met een melding)
    Dim i As Long, j As Long, b As String, bBest As String, nBest As Long, cnt As Long, eerste As Boolean
    Dim klein As Long, groot As Long, k As Long, n As Long, samen As String
    Do While nL > nMax
        bBest = ""
        nBest = 0
        For i = 1 To nL
            b = BasisNaam(lNaam(i))
            If Len(b) > 0 Then
                eerste = True
                For j = 1 To i - 1
                    If BasisNaam(lNaam(j)) = b Then eerste = False
                Next j
                If eerste Then
                    cnt = 0
                    For j = i To nL
                        If BasisNaam(lNaam(j)) = b Then cnt = cnt + 1
                    Next j
                    If cnt > 1 And cnt > nBest Then
                        bBest = b
                        nBest = cnt
                    End If
                End If
            End If
        Next i
        If nBest = 0 Then
            If laatste Then
                ' fees: de rest gaat op in de laatste regel (' e.a.'; bedrag blijft, kwartaal van die regel), zoals fo.py afkappen='laatste'
                samen = ""
                For i = nMax + 1 To nL
                    samen = samen & IIf(Len(samen) > 0, ", ", "") & lNaam(i)
                    lW(nMax) = lW(nMax) + lW(i)
                Next i
                waarschuwing = waarschuwing & "- " & typeNaam & ": " & nL & " " & wat & " in het FO, ruimte voor " & nMax & "; " & samen & " staan bij '" & lNaam(nMax) & "' (in " & CStr(lKw(nMax)) & ")." & vbCrLf
                lW(nMax) = Round(lW(nMax), 6)
                lNaam(nMax) = lNaam(nMax) & " e.a."
                samen = ""
            Else
                waarschuwing = waarschuwing & "- " & typeNaam & ": " & nL & " " & wat & " in het FO, ruimte voor " & nMax & "; de laatste zijn niet overgenomen." & vbCrLf
            End If
            nL = nMax
            Exit Do
        End If
        klein = 0
        For i = 1 To nL
            If BasisNaam(lNaam(i)) = bBest Then
                If klein = 0 Then
                    klein = i
                ElseIf Abs(lW(i)) < Abs(lW(klein)) Then
                    klein = i
                End If
            End If
        Next i
        groot = 0
        For i = 1 To nL
            If i <> klein And BasisNaam(lNaam(i)) = bBest Then
                If groot = 0 Then
                    groot = i
                ElseIf Abs(lW(i)) > Abs(lW(groot)) Then
                    groot = i
                End If
            End If
        Next i
        lW(groot) = Round(lW(groot) + lW(klein), 6)
        For i = klein To nL - 1
            lNaam(i) = lNaam(i + 1)
            lW(i) = lW(i + 1)
            lKw(i) = lKw(i + 1)
        Next i
        nL = nL - 1
        If InStr(samen, "'" & bBest & "'") = 0 Then samen = samen & IIf(Len(samen) > 0, ", ", "") & "'" & bBest & "'"
        n = 0
        For i = 1 To nL
            If BasisNaam(lNaam(i)) = bBest Then n = n + 1
        Next i
        k = 0
        For i = 1 To nL
            If BasisNaam(lNaam(i)) = bBest Then
                k = k + 1
                If n = 1 Then lNaam(i) = bBest Else lNaam(i) = bBest & " (" & k & "/" & n & ")"
            End If
        Next i
    Loop
    If Len(samen) > 0 Then
        waarschuwing = waarschuwing & "- " & typeNaam & ": meer dan " & nMax & " " & wat & " na het splitsen over kwartalen; kleinste delen samengevoegd bij het grootste deel van " & samen & "." & vbCrLf
    End If
End Sub

' -------------------------------------------------------------------------------------
'  5. Sjabloon controleren: staan alle vormnamen erin?
' -------------------------------------------------------------------------------------
Public Sub ControleerSjabloon()
    ' vormnamen uit kolom F van de KPI-regels (o.a. BT_TITEL, BT_SUBTITEL, VP_TITEL, VP_SUBTITEL, VO_TITEL, CF_*, SC_*,
    ' KPI*), uit Grafieken() en uit Tabellen(); daarnaast de hulpvormen van dia 5 die SchikDia5 gebruikt
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim r As Long, vorm As String, mist As String, aantal As Long, i As Long, lijst As Variant, hulp As String
    On Error GoTo Fout
    Set wsP = ThisWorkbook.Worksheets(SH_PP)
    Set pp = HaalPowerPoint()
    If pp Is Nothing Then
        MsgBox "PowerPoint kon niet worden gestart.", vbExclamation, "Sjabloon controleren"
        Exit Sub
    End If
    Set pres = OpenSjabloon(pp, wsP)
    If pres Is Nothing Then Exit Sub
    r = KPI_ROW1
    Do While Len(CelTekst(wsP.Range("C" & r))) > 0 And r < KPI_ROW1 + 300
        vorm = Trim$(CelTekst(wsP.Range("F" & r)))
        If Len(vorm) > 0 Then
            aantal = aantal + 1
            If ZoekVorm(pres, vorm) Is Nothing Then mist = mist & vorm & vbCrLf
        End If
        r = r + 1
    Loop
    lijst = Grafieken()
    For i = LBound(lijst) To UBound(lijst)
        aantal = aantal + 1
        If ZoekVorm(pres, CStr(lijst(i)(0))) Is Nothing Then mist = mist & CStr(lijst(i)(0)) & vbCrLf
    Next i
    lijst = Tabellen()
    For i = LBound(lijst) To UBound(lijst)
        If CStr(lijst(i)(0)) <> "VT_TABEL" Then                   ' sinds v20 geen tabel woningtypes in het sjabloon
            aantal = aantal + 1
            If ZoekVorm(pres, CStr(lijst(i)(0))) Is Nothing Then mist = mist & CStr(lijst(i)(0)) & vbCrLf
        End If
    Next i
    ' hulpvormen van dia 5 (kaarten) en dia 2 (grijs vlak, nu-blok, sporen): zonder die vormen vult de knop wel, maar schikt hij niet
    lijst = Array("VT_KAART", "GRAFIEK_REALISATIE", "TL_NU_LABEL", "KPI6_SPOOR", "KPI7_SPOOR_V", "KPI7_SPOOR_T", "KPI9_SPOOR1", _
                  "VS_REALISATIE", "VS_K0_KAART", "VS_LEG0_LIJN")
    For i = LBound(lijst) To UBound(lijst)
        If ZoekVorm(pres, CStr(lijst(i))) Is Nothing Then hulp = hulp & CStr(lijst(i)) & vbCrLf
    Next i
    On Error Resume Next
    pres.Saved = -1
    pres.Close
    On Error GoTo 0
    If Len(hulp) > 0 Then
        hulp = vbCrLf & "Hulpvormen (dia 2, 4 en 6) die niet gevonden zijn (die dia's worden dan niet automatisch geschikt):" & vbCrLf & hulp
    End If
    If Len(mist) = 0 Then
        MsgBox "Alles compleet: alle " & aantal & " vormnamen staan in het sjabloon." & hulp, _
               IIf(Len(hulp) > 0, vbExclamation, vbInformation), "Sjabloon controleren"
    Else
        MsgBox "Deze vormen ontbreken in het sjabloon (die slaat de knop over):" & vbCrLf & vbCrLf & mist & vbCrLf & _
               "Hernoem de vormen in PowerPoint via Start > Selecteren > Selectiedeelvenster." & vbCrLf & hulp, _
               vbExclamation, "Sjabloon controleren"
    End If
    Exit Sub
Fout:
    MsgBox "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Sjabloon controleren"
End Sub
