Attribute VB_Name = "FOWerkbestand"
Option Explicit

' =====================================================================================
'  FO-werkbestand                                           (late binding, Windows en Mac)
'  Knoppen op tab Start:
'    Plakken / Inlezen...   per invoertabblad (FO - vorig, FO - huidig, FO - actuals, Prijslijst dump):
'                           leegmaken, export erin (klembord of bestand), hulpkolommen doortrekken, status op Start
'    Oud FO inlezen...      neemt alle tabbladen van een FO in de oude opzet over (formules en waarden), zet de nieuwe
'                           onderdelen terug (tab _sjabloon: transportregels, methode, voorstel, ...), vult tab Vorige
'                           prognose (kolom CF vorig kwartaal) en tab Controle oud FO
'    Nieuw kwartaal starten cashflow bewaren als vorige prognose, FO - huidig -> FO - vorig, huidig en actuals leeg
'    Scenario bijwerken     (module CashflowNaarPowerPoint) scenariomodel uit de FO-tabbladen van dit bestand
'    PowerPoint maken       eerst Scenario bijwerken, dan Naar PowerPoint
'  Vaste plekken: zie tools/nieuwe_bladen.py.
' =====================================================================================

Private Const SH_START As String = "Start"
Private Const SH_CF As String = "Cashflow"            ' vervangt 1. Cashflow (zelfde opbouw)
Private Const SH_CF_OUD As String = "1. Cashflow"     ' in een oud FO
Private Const SH_VORIGE As String = "Vorige prognose"
Private Const SH_CONTROLE As String = "Controle oud FO"
Private Const SH_SJABLOON As String = "_sjabloon"     ' blad / cel / inhoud: na Oud FO inlezen terugzetten
Private Const TITEL As String = "FO-werkbestand"

Private Const RIJ_VORIG As Long = 15                 ' rijen op tab Start (blok 2)
Private Const RIJ_HUIDIG As Long = 16
Private Const RIJ_ACTUALS As Long = 17
Private Const RIJ_PRIJS As Long = 18
Private Const KOL_DATUM As Long = 6                  ' Start kolom F: laatst ingelezen
Private Const KOL_BESTAND As Long = 7                ' Start kolom G: bestand / klembord

Private Const TRANSPORTRIJEN As String = "13|26|36|46"   ' CF - opbrengsten: transportregel per woningtype

' tabbladen die 'Oud FO inlezen' overneemt (zelfde naam in het oude FO)
Private Const FO_BLADEN As String = "Toelichting onderlegger|Behind the scens|Prijslijst dump|Prijslijst|Aansluiting opbrengsten|" & _
    "Notities FO gesprek|CF - opbrengsten|CF - kosten|Verkooptempo|FO - vorig|FO - huidig|FO - actuals|1. Ontw businessc FD|" & _
    "2. Ontw PEW|5. K&R businessc|CF - hulp VENA"

' =====================================================================================
'  knoppen
' =====================================================================================

Public Sub StartKnoppen()
    ' maakt alleen knoppen die nog niet bestaan (draait bij het openen)
    Dim ws As Worksheet, donker As Long, grijs As Long, geel As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_START)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    donker = RGB(17, 31, 57)
    grijs = RGB(91, 97, 105)
    geel = RGB(254, 185, 0)
    Knop ws, "J6:J8", "btnOudFO", "Oud FO inlezen...", "OudFOInlezen", grijs
    Knop ws, "K6:K8", "btnNieuwKwartaal", "Nieuw kwartaal starten", "NieuwKwartaalStarten", grijs
    Knop ws, "J9:J11", "btnScenario", "Scenario bijwerken", "ScenarioBijwerken", donker
    Knop ws, "K9:K11", "btnPowerPoint", "PowerPoint maken", "PowerPointMaken", donker
    Knop ws, "J" & RIJ_VORIG, "btnPlakVorig", "Plakken", "PlakkenVorig", donker
    Knop ws, "K" & RIJ_VORIG, "btnLeesVorig", "Inlezen...", "InlezenVorig", grijs
    Knop ws, "J" & RIJ_HUIDIG, "btnPlakHuidig", "Plakken", "PlakkenHuidig", donker
    Knop ws, "K" & RIJ_HUIDIG, "btnLeesHuidig", "Inlezen...", "InlezenHuidig", grijs
    Knop ws, "J" & RIJ_ACTUALS, "btnPlakActuals", "Plakken", "PlakkenActuals", donker
    Knop ws, "K" & RIJ_ACTUALS, "btnLeesActuals", "Inlezen...", "InlezenActuals", grijs
    Knop ws, "J" & RIJ_PRIJS, "btnPlakPrijs", "Plakken", "PlakkenPrijslijst", donker
    Knop ws, "K" & RIJ_PRIJS, "btnLeesPrijs", "Inlezen...", "InlezenPrijslijst", grijs
End Sub

Private Sub Knop(ws As Worksheet, plek As String, naam As String, tekst As String, macro As String, kleur As Long)
    Dim shp As Shape, r As Range
    On Error Resume Next
    Set shp = ws.Shapes(naam)
    On Error GoTo 0
    If Not shp Is Nothing Then Exit Sub
    Set r = ws.Range(plek)
    Set shp = ws.Shapes.AddShape(5, r.Left + 2, r.Top + 2, r.Width - 4, IIf(r.Height - 4 < 16, 16, r.Height - 4))
    With shp
        .Name = naam
        .Fill.ForeColor.RGB = kleur
        .Line.Visible = 0
        .TextFrame2.TextRange.Text = tekst
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = -1
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame2.VerticalAnchor = 3
        .TextFrame2.WordWrap = -1
        .TextFrame2.MarginLeft = 2
        .TextFrame2.MarginRight = 2
        .Placement = 2
        .OnAction = macro
    End With
End Sub

' =====================================================================================
'  invoer: plakken of inlezen
' =====================================================================================

Public Sub PlakkenVorig()
    Plakken RIJ_VORIG
End Sub
Public Sub PlakkenHuidig()
    Plakken RIJ_HUIDIG
End Sub
Public Sub PlakkenActuals()
    Plakken RIJ_ACTUALS
End Sub
Public Sub PlakkenPrijslijst()
    Plakken RIJ_PRIJS
End Sub
Public Sub InlezenVorig()
    Inlezen RIJ_VORIG
End Sub
Public Sub InlezenHuidig()
    Inlezen RIJ_HUIDIG
End Sub
Public Sub InlezenActuals()
    Inlezen RIJ_ACTUALS
End Sub
Public Sub InlezenPrijslijst()
    Inlezen RIJ_PRIJS
End Sub

Private Sub InvoerInfo(rij As Long, ByRef blad As String, ByRef nKol As Long, ByRef hulp As String, ByRef hulpRij As Long)
    ' nKol = kolommen van de export (A..), hulp = hulpkolommen met formules rechts ervan, hulpRij = eerste rij met de formule
    Select Case rij
        Case RIJ_VORIG
            blad = "FO - vorig": nKol = 21: hulp = "X|Y": hulpRij = 17
        Case RIJ_HUIDIG
            blad = "FO - huidig": nKol = 21: hulp = "X|Y": hulpRij = 17
        Case RIJ_ACTUALS
            blad = "FO - actuals": nKol = 16: hulp = "Q|R": hulpRij = 2
        Case Else
            blad = "Prijslijst dump": nKol = 9: hulp = "": hulpRij = 2
    End Select
End Sub

Private Sub Plakken(rij As Long)
    Dim blad As String, nKol As Long, hulp As String, hulpRij As Long, ws As Worksheet, w As Long, h As Long, gelukt As Boolean
    InvoerInfo rij, blad, nKol, hulp, hulpRij
    Set ws = ThisWorkbook.Worksheets(blad)
    If MsgBox("Tabblad '" & blad & "' wordt leeggemaakt (kolom A t/m " & KolomLetter(nKol) & ") en de export op het klembord " & _
              "wordt vanaf A1 geplakt (alleen waarden)." & vbCrLf & vbCrLf & _
              "Kopieer eerst de hele export, vanaf cel A1. Doorgaan?", vbQuestion + vbYesNo, "Plakken in " & blad) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    LeegMaken ws, nKol
    ws.Activate
    ws.Range("A1").Select
    On Error Resume Next
    ws.Range("A1").PasteSpecial Paste:=-4163                 ' xlPasteValues (gekopieerd uit Excel)
    If Err.Number <> 0 Then
        Err.Clear
        ws.Paste Destination:=ws.Range("A1")                 ' tekst uit een ander programma
    End If
    gelukt = (Err.Number = 0)
    On Error GoTo 0
    Application.CutCopyMode = 0
    If Not gelukt Or Application.WorksheetFunction.CountA(ws.Range(ws.Cells(1, 1), ws.Cells(50, nKol))) = 0 Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Er stond niets bruikbaars op het klembord. Kopieer de export en probeer het opnieuw.", vbExclamation, TITEL
        Exit Sub
    End If
    ' wat buiten de exportkolommen is geplakt weg (de hulpkolommen komen hieronder terug)
    w = 0
    On Error Resume Next
    w = Selection.Column + Selection.Columns.Count - 1
    On Error GoTo 0
    h = LaatsteRij(ws, nKol)
    If w > nKol Then ws.Range(ws.Cells(1, nKol + 1), ws.Cells(h, w)).ClearContents
    HulpBijwerken ws, hulp, hulpRij, h
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ZetStatus rij, "klembord"
    ThisWorkbook.Worksheets(SH_START).Activate
    MsgBox "Klaar: " & h & " rijen in '" & blad & "'. Controleer de status op tab Start.", vbInformation, TITEL
End Sub

Private Sub Inlezen(rij As Long)
    Dim blad As String, nKol As Long, hulp As String, hulpRij As Long, ws As Worksheet
    Dim pad As Variant, wbE As Workbook, wsE As Worksheet, wsX As Worksheet, data As Variant, nR As Long, nC As Long
    Dim foutNr As Long, foutOms As String
    InvoerInfo rij, blad, nKol, hulp, hulpRij
    Set ws = ThisWorkbook.Worksheets(blad)
#If Mac Then
    pad = Application.GetOpenFilename()
#Else
    pad = Application.GetOpenFilename("Exports (*.xlsx;*.xlsm;*.xls;*.csv;*.txt),*.xlsx;*.xlsm;*.xls;*.csv;*.txt", , "Export kiezen voor " & blad)
#End If
    If VarType(pad) = vbBoolean Then Exit Sub
    On Error GoTo Fout
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wbE = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True)
    Application.DisplayAlerts = True
    ' het werkblad met de meeste gegevens
    Set wsE = wbE.Worksheets(1)
    For Each wsX In wbE.Worksheets
        If wsX.UsedRange.Rows.Count * wsX.UsedRange.Columns.Count > wsE.UsedRange.Rows.Count * wsE.UsedRange.Columns.Count Then Set wsE = wsX
    Next wsX
    nR = wsE.UsedRange.Row + wsE.UsedRange.Rows.Count - 1
    nC = wsE.UsedRange.Column + wsE.UsedRange.Columns.Count - 1
    If nC > nKol Then nC = nKol
    data = wsE.Range(wsE.Cells(1, 1), wsE.Cells(nR, nC)).Value
    wbE.Close SaveChanges:=False
    Set wbE = Nothing
    LeegMaken ws, nKol
    ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Value = data
    HulpBijwerken ws, hulp, hulpRij, LaatsteRij(ws, nKol)
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ZetStatus rij, BestandsNaam(CStr(pad))
    MsgBox "Klaar: " & nR & " rijen in '" & blad & "' uit " & BestandsNaam(CStr(pad)) & ". Controleer de status op tab Start.", _
           vbInformation, TITEL
    Exit Sub
Fout:
    foutNr = Err.Number
    foutOms = Err.Description
    On Error Resume Next
    If Not wbE Is Nothing Then wbE.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Inlezen in '" & blad & "' is niet gelukt." & vbCrLf & vbCrLf & "Fout " & foutNr & ": " & foutOms, vbExclamation, TITEL
End Sub

Private Sub LeegMaken(ws As Worksheet, nKol As Long)
    Dim h As Long
    h = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If h < 1 Then h = 1
    ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol)).ClearContents
End Sub

Private Function LaatsteRij(ws As Worksheet, nKol As Long) As Long
    ' laatste rij met iets in de exportkolommen
    Dim c As Long, r As Long, m As Long
    For c = 1 To nKol
        r = ws.Cells(ws.Rows.Count, c).End(-4162).Row     ' xlUp
        If r > m Then m = r
    Next c
    LaatsteRij = m
End Function

Private Sub HulpBijwerken(ws As Worksheet, hulp As String, hulpRij As Long, laatste As Long)
    ' de formule van de eerste rij doortrekken t/m de laatste exportrij (rijen daaronder houden hun formule)
    Dim kolommen() As String, i As Long, f As String
    If Len(hulp) = 0 Or laatste <= hulpRij Then Exit Sub
    kolommen = Split(hulp, "|")
    For i = LBound(kolommen) To UBound(kolommen)
        f = ws.Range(kolommen(i) & hulpRij).FormulaR1C1
        If Left$(f, 1) = "=" Then ws.Range(kolommen(i) & hulpRij & ":" & kolommen(i) & laatste).FormulaR1C1 = f
    Next i
End Sub

Private Sub ZetStatus(rij As Long, bron As String)
    With ThisWorkbook.Worksheets(SH_START)
        .Cells(rij, KOL_DATUM).Value = Now
        .Cells(rij, KOL_BESTAND).Value = bron
    End With
End Sub

' =====================================================================================
'  nieuw kwartaal
' =====================================================================================

Public Sub NieuwKwartaalStarten()
    Dim wsH As Worksheet, wsV As Worksheet, wsA As Worksheet, h As Long, n As Long
    If MsgBox("Nieuw kwartaal starten:" & vbCrLf & vbCrLf & _
              "1. de cashflow (tab Cashflow) wordt bewaard als vorige prognose (tab Vorige prognose); " & _
              "die verschijnt in kolom J 'CF vorig kwartaal';" & vbCrLf & _
              "2. FO - huidig wordt FO - vorig;" & vbCrLf & _
              "3. FO - huidig en FO - actuals worden leeggemaakt voor de nieuwe exports." & vbCrLf & vbCrLf & _
              "Doe dit pas als het vorige FO helemaal klaar is. Doorgaan?", vbQuestion + vbYesNo, "Nieuw kwartaal starten") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    n = BewaarVorigePrognose()
    Set wsH = ThisWorkbook.Worksheets("FO - huidig")
    Set wsV = ThisWorkbook.Worksheets("FO - vorig")
    Set wsA = ThisWorkbook.Worksheets("FO - actuals")
    h = LaatsteRij(wsH, 21)
    LeegMaken wsV, 21
    wsV.Range("A1:U" & h).Value = wsH.Range("A1:U" & h).Value
    HulpBijwerken wsV, "X|Y", 17, h
    LeegMaken wsH, 21
    LeegMaken wsA, 16
    With ThisWorkbook.Worksheets(SH_START)
        .Cells(RIJ_VORIG, KOL_DATUM).Value = Now
        .Cells(RIJ_VORIG, KOL_BESTAND).Value = "uit FO - huidig (nieuw kwartaal)"
        .Range(.Cells(RIJ_HUIDIG, KOL_DATUM), .Cells(RIJ_ACTUALS, KOL_BESTAND)).ClearContents
    End With
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_START).Activate
    MsgBox "Klaar: " & n & " periodes bewaard als vorige prognose en FO - huidig staat nu in FO - vorig." & vbCrLf & vbCrLf & _
           "Volgende stap: plak of lees de nieuwe exports in (FO - huidig, FO - actuals en de prijslijst)." & vbCrLf & _
           "Wil je ook de planning (start bouw per type) als vorige prognose bewaren? Gebruik dan op tab Woningtypes " & _
           "de knop 'Nu bewaren als vorige prognose'.", vbInformation, "Nieuw kwartaal starten"
End Sub

Private Function BewaarVorigePrognose() As Long
    ' tab Cashflow -> tab Vorige prognose (waarden); kolom J (CF vorig kwartaal) op tab Cashflow leest daaruit
    Dim wsC As Worksheet, wsV As Worksheet, kop As Long, kolJ As Long, r As Long, c As Long, i As Long
    Set wsC = ThisWorkbook.Worksheets(SH_CF)
    Set wsV = ThisWorkbook.Worksheets(SH_VORIGE)
    For r = 1 To 30
        For c = 1 To 6
            If LCase$(Trim$(CStr(wsC.Cells(r, c).Value))) = "jaar" Then
                kop = r
                kolJ = c
                Exit For
            End If
        Next c
        If kop > 0 Then Exit For
    Next r
    If kop = 0 Then Exit Function
    wsV.Range("B8:G67").ClearContents
    r = kop + 1
    Do While IsNumeric(wsC.Cells(r, kolJ).Value) And Not IsEmpty(wsC.Cells(r, kolJ).Value) And i < 60
        wsV.Cells(8 + i, 2).Value = wsC.Cells(r, kolJ).Value
        wsV.Cells(8 + i, 3).Value = CStr(wsC.Cells(r, kolJ + 1).Value)
        wsV.Cells(8 + i, 4).Value = wsC.Cells(r, kolJ + 2).Value
        wsV.Cells(8 + i, 5).Value = wsC.Cells(r, kolJ + 4).Value
        wsV.Cells(8 + i, 6).Value = wsC.Cells(r, kolJ + 6).Value
        wsV.Cells(8 + i, 7).Value = wsC.Cells(r, kolJ + 7).Value
        i = i + 1
        r = r + 1
    Loop
    wsV.Range("D5").Value = Now
    BewaarVorigePrognose = i
End Function

' =====================================================================================
'  oud FO inlezen
' =====================================================================================

Public Sub OudFOInlezen()
    Dim pad As Variant, wbO As Workbook, wsO As Worksheet, wsN As Worksheet, bladen() As String, i As Long
    Dim waarschuwing As String, rekenmodus As Long, stap As String, foutNr As Long, foutOms As String, n As Long
    Dim zelfGeopend As Boolean, wb As Workbook, verschil As Long, bericht As String
#If Mac Then
    pad = Application.GetOpenFilename()
#Else
    pad = Application.GetOpenFilename("Excel-werkboeken (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", , "Oud FO kiezen")
#End If
    If VarType(pad) = vbBoolean Then Exit Sub
    If LCase$(CStr(pad)) = LCase$(ThisWorkbook.FullName) Then
        MsgBox "Kies een ander FO-bestand dan dit werkbestand.", vbExclamation, TITEL
        Exit Sub
    End If
    If MsgBox("Alle FO-tabbladen van dit werkbestand (FO - vorig/huidig/actuals, prijslijst, Aansluiting opbrengsten, " & _
              "CF - opbrengsten, CF - kosten, 1. Cashflow, enz.) worden overschreven met de inhoud van:" & vbCrLf & CStr(pad) & _
              vbCrLf & vbCrLf & "De nieuwe onderdelen (transportregels, methode per kostenpost, voorstel, correctieregel) worden daarna " & _
              "teruggezet; zonder aanpassingen is de uitkomst gelijk aan het oude FO (zie tab Controle oud FO). Doorgaan?", _
              vbQuestion + vbYesNo, "Oud FO inlezen") <> vbYes Then Exit Sub
    On Error GoTo Fout
    stap = "oud FO openen"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    rekenmodus = Application.Calculation
    Application.Calculation = -4135                      ' handmatig rekenen tijdens het kopieren
    For Each wb In Workbooks
        If LCase$(wb.FullName) = LCase$(CStr(pad)) Then Set wbO = wb
    Next wb
    If wbO Is Nothing Then
        Application.DisplayAlerts = False
        Set wbO = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True)
        Application.DisplayAlerts = True
        zelfGeopend = True
    End If
    wbO.Application.Calculate

    bladen = Split(FO_BLADEN, "|")
    For i = LBound(bladen) To UBound(bladen)
        stap = "tabblad '" & bladen(i) & "'"
        Set wsO = ZoekBlad(wbO, bladen(i))
        If wsO Is Nothing Then
            waarschuwing = waarschuwing & "- '" & bladen(i) & "' staat niet in het oude FO: niet overgenomen." & vbCrLf
        Else
            Set wsN = ThisWorkbook.Worksheets(bladen(i))
            KopieerBlad wsO, wsN, waarschuwing
            n = n + 1
        End If
    Next i

    stap = "opzet controleren"
    OpzetControleren waarschuwing
    stap = "nieuwe onderdelen terugzetten (_sjabloon)"
    SjabloonTerugzetten waarschuwing
    stap = "hulpkolommen"
    HulpBijwerken ThisWorkbook.Worksheets("FO - vorig"), "X|Y", 17, LaatsteRij(ThisWorkbook.Worksheets("FO - vorig"), 21)
    HulpBijwerken ThisWorkbook.Worksheets("FO - huidig"), "X|Y", 17, LaatsteRij(ThisWorkbook.Worksheets("FO - huidig"), 21)
    HulpBijwerken ThisWorkbook.Worksheets("FO - actuals"), "Q|R", 2, LaatsteRij(ThisWorkbook.Worksheets("FO - actuals"), 16)
    stap = "vorige prognose en controle vullen"
    ControleVullen wbO, CStr(pad)
    If zelfGeopend Then wbO.Close SaveChanges:=False
    Set wbO = Nothing

    stap = "doorrekenen"
    Application.Calculation = rekenmodus
    Application.EnableEvents = True
    Application.CalculateFull
    With ThisWorkbook.Worksheets(SH_START)
        For i = RIJ_VORIG To RIJ_PRIJS
            .Cells(i, KOL_DATUM).Value = Now
            .Cells(i, KOL_BESTAND).Value = "oud FO: " & BestandsNaam(CStr(pad))
        Next i
    End With
    verschil = Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets(SH_CONTROLE).Range("M9:M60"), "VERSCHIL")

    stap = "scenario bijwerken"
    Application.ScreenUpdating = True
    CashflowNaarPowerPoint.OphalenUitFO "", True, True

    ThisWorkbook.Worksheets(SH_CONTROLE).Activate
    bericht = "Klaar: " & n & " tabbladen overgenomen uit " & BestandsNaam(CStr(pad)) & "." & vbCrLf & vbCrLf
    If verschil = 0 Then
        bericht = bericht & "De cashflow is per periode gelijk aan het oude FO (tab Controle oud FO)."
    Else
        bericht = bericht & verschil & " periode(s) wijken af van het oude FO: zie tab Controle oud FO."
    End If
    If Not CashflowNaarPowerPoint.OphalenGelukt Then bericht = bericht & vbCrLf & "Het scenario is niet bijgewerkt (zie de melding)."
    If Len(waarschuwing) > 0 Then bericht = bericht & vbCrLf & vbCrLf & "Let op:" & vbCrLf & waarschuwing
    MsgBox bericht, IIf(Len(waarschuwing) > 0 Or verschil > 0, vbExclamation, vbInformation), "Oud FO inlezen"
    Exit Sub
Fout:
    foutNr = Err.Number
    foutOms = Err.Description
    On Error Resume Next
    If zelfGeopend And Not wbO Is Nothing Then wbO.Close SaveChanges:=False
    If rekenmodus <> 0 Then Application.Calculation = rekenmodus
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Oud FO inlezen ging mis bij: " & stap & vbCrLf & vbCrLf & "Fout " & foutNr & ": " & foutOms, vbExclamation, "Oud FO inlezen"
End Sub

Private Sub KopieerBlad(wsO As Worksheet, wsN As Worksheet, ByRef waarschuwing As String)
    ' alle formules en waarden van het oude tabblad op dezelfde plek; formules die naar een tabblad verwijzen dat hier niet
    ' bestaat (of naar een ander bestand) worden waarden, zodat Excel niet om een bestand gaat vragen
    Dim nR As Long, nC As Long, f As Variant, v As Variant, i As Long, j As Long, s As String, mist As Long, r As Long
    nR = wsO.UsedRange.Row + wsO.UsedRange.Rows.Count - 1
    nC = wsO.UsedRange.Column + wsO.UsedRange.Columns.Count - 1
    If nR < 2 Then nR = 2
    If nC < 2 Then nC = 2
    f = wsO.Range(wsO.Cells(1, 1), wsO.Cells(nR, nC)).Formula
    v = wsO.Range(wsO.Cells(1, 1), wsO.Cells(nR, nC)).Value
    For i = 1 To nR
        For j = 1 To nC
            s = CStr(f(i, j))
            If Left$(s, 1) = "=" Then
                If Not VerwijzingenBestaan(s) Then
                    If IsError(v(i, j)) Then f(i, j) = "" Else f(i, j) = v(i, j)
                    mist = mist + 1
                End If
            End If
        Next j
    Next i
    wsN.UsedRange.ClearContents
    On Error Resume Next
    wsN.Range(wsN.Cells(1, 1), wsN.Cells(nR, nC)).Formula = f
    If Err.Number <> 0 Then
        ' bijvoorbeeld samengevoegde cellen die anders liggen: per cel
        Err.Clear
        For i = 1 To nR
            For j = 1 To nC
                If Len(CStr(f(i, j))) > 0 Then
                    wsN.Cells(i, j).Formula = f(i, j)
                    If Err.Number <> 0 Then
                        Err.Clear
                        r = r + 1
                    End If
                End If
            Next j
        Next i
    End If
    On Error GoTo 0
    If mist > 0 Then waarschuwing = waarschuwing & "- '" & wsN.Name & "': " & mist & " formule(s) met een verwijzing naar een " & _
                                    "ander bestand of tabblad zijn als waarde overgenomen." & vbCrLf
    If r > 0 Then waarschuwing = waarschuwing & "- '" & wsN.Name & "': " & r & " cel(len) niet overgenomen (samengevoegde cellen)." & vbCrLf
End Sub

Private Function VerwijzingenBestaan(f As String) As Boolean
    ' True als elke bladverwijzing (Blad!A1 of 'Blad naam'!A1) in deze formule naar een tabblad van dit werkbestand wijst
    Dim p As Long, q As Long, naam As String, ch As String
    VerwijzingenBestaan = True
    If InStr(f, "[") > 0 Then
        VerwijzingenBestaan = False
        Exit Function
    End If
    p = InStr(f, "!")
    Do While p > 0
        If BinnenTekst(f, p) Then
            naam = ""
        ElseIf p > 1 And Mid$(f, p - 1, 1) = "'" Then
            q = p - 2
            Do While q > 0
                If Mid$(f, q, 1) = "'" Then
                    If q > 1 Then
                        If Mid$(f, q - 1, 1) = "'" Then
                            q = q - 2
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    q = q - 1
                End If
            Loop
            naam = Replace(Mid$(f, q + 1, p - q - 2), "''", "'")
        Else
            q = p - 1
            Do While q > 0
                ch = Mid$(f, q, 1)
                If ch Like "[A-Za-z0-9_.]" Then q = q - 1 Else Exit Do
            Loop
            naam = Mid$(f, q + 1, p - q - 1)
        End If
        If Len(naam) > 0 Then
            If ZoekBlad(ThisWorkbook, naam) Is Nothing Then
                VerwijzingenBestaan = False
                Exit Function
            End If
        End If
        p = InStr(p + 1, f, "!")
    Loop
End Function

Private Function BinnenTekst(f As String, p As Long) As Boolean
    ' staat positie p tussen dubbele aanhalingstekens (een tekst in de formule)?
    Dim i As Long, n As Long
    For i = 1 To p - 1
        If Mid$(f, i, 1) = Chr$(34) Then n = n + 1
    Next i
    BinnenTekst = (n Mod 2 = 1)
End Function

Private Sub SjabloonTerugzetten(ByRef waarschuwing As String)
    ' tab _sjabloon: per regel blad, cel en inhoud ('=formule', waarde of leeg) van de onderdelen die een oud FO niet heeft
    Dim ws As Worksheet, r As Long, blad As String, cel As String, inhoud As String, mis As Long
    Set ws = ThisWorkbook.Worksheets(SH_SJABLOON)
    r = 2
    Do While Len(CStr(ws.Cells(r, 1).Value)) > 0
        blad = CStr(ws.Cells(r, 1).Value)
        cel = CStr(ws.Cells(r, 2).Value)
        inhoud = CStr(ws.Cells(r, 3).Value)
        On Error Resume Next
        If Len(inhoud) = 0 Then
            ThisWorkbook.Worksheets(blad).Range(cel).ClearContents
        Else
            ThisWorkbook.Worksheets(blad).Range(cel).Formula = inhoud
        End If
        If Err.Number <> 0 Then
            Err.Clear
            mis = mis + 1
        End If
        On Error GoTo 0
        r = r + 1
    Loop
    If mis > 0 Then waarschuwing = waarschuwing & "- " & mis & " cel(len) van het sjabloon konden niet worden teruggezet." & vbCrLf
End Sub

Private Sub OpzetControleren(ByRef waarschuwing As String)
    ' de prognosebladen volgen de rijen van de standaardopzet; wijkt het oude FO daarvan af, dan melden
    Dim fout As String
    fout = fout & Controleer("CF - kosten", "B58", "onderdelen")
    fout = fout & Controleer("CF - kosten", "B75", "totaal kosten")
    fout = fout & Controleer("CF - kosten", "B77", "periode in cashflow sap")
    fout = fout & Controleer("CF - opbrengsten", "B128", "termijnen 1 in")
    fout = fout & Controleer("CF - opbrengsten", "B189", "overige opbrengsten in")
    fout = fout & Controleer("CF - opbrengsten", "B194", "totaal opbrengsten")
    fout = fout & Controleer("CF - opbrengsten", "B196", "periode in cashflow sap")
    ' de transportregels komen in lege rijen onder de bouwtermijnen; stond daar iets, dan melden
    Dim t As Variant
    For Each t In Split(TRANSPORTRIJEN, "|")
        If Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets("CF - opbrengsten").Range("B" & t & ":Z" & t)) > 0 Then
            fout = fout & "    CF - opbrengsten rij " & t & " was niet leeg; daar staat nu de transportregel" & vbCrLf
        End If
    Next t
    If Len(fout) > 0 Then
        waarschuwing = waarschuwing & "- de opzet van het oude FO wijkt af van de standaard; controleer:" & vbCrLf & fout
    End If
End Sub

Private Function Controleer(blad As String, cel As String, begin As String) As String
    Dim t As String
    t = LCase$(Trim$(CStr(ThisWorkbook.Worksheets(blad).Range(cel).Value)))
    If Left$(t, Len(begin)) <> begin Then Controleer = "    " & blad & "!" & cel & " (verwacht '" & begin & "...')" & vbCrLf
End Function

Private Sub ControleVullen(wbO As Workbook, pad As String)
    ' cashflow van het oude FO (1. Cashflow, waarden) naar tab Controle oud FO (per jaar en kwartaal) en kolom
    ' 'CF vorig kwartaal' naar tab Vorige prognose, zodat tab Cashflow kolom J doorloopt
    Dim wsO As Worksheet, wsK As Worksheet, wsV As Worksheet, i As Long, r As Long
    Set wsK = ThisWorkbook.Worksheets(SH_CONTROLE)
    Set wsV = ThisWorkbook.Worksheets(SH_VORIGE)
    wsK.Range("B9:D60,G9:G60,J9:J60").ClearContents
    wsK.Range("D5").Value = BestandsNaam(pad)
    wsK.Range("D6").Value = Now
    Set wsO = ZoekBlad(wbO, SH_CF_OUD)
    If wsO Is Nothing Then Set wsO = ZoekBlad(wbO, SH_CF)
    If wsO Is Nothing Then Exit Sub
    wsV.Range("B8:G67").ClearContents
    wsV.Range("D5").Value = "oud FO: " & BestandsNaam(pad)
    For i = 0 To 51
        r = 4 + i
        If Not IsNumeric(wsO.Cells(r, 2).Value) Or IsEmpty(wsO.Cells(r, 2).Value) Then Exit For
        wsK.Cells(9 + i, 2).Value = wsO.Cells(r, 2).Value
        wsK.Cells(9 + i, 3).Value = CStr(wsO.Cells(r, 3).Value)
        If Not IsError(wsO.Cells(r, 4).Value) Then wsK.Cells(9 + i, 4).Value = wsO.Cells(r, 4).Value
        If Not IsError(wsO.Cells(r, 6).Value) Then wsK.Cells(9 + i, 7).Value = wsO.Cells(r, 6).Value
        If Not IsError(wsO.Cells(r, 8).Value) Then wsK.Cells(9 + i, 10).Value = wsO.Cells(r, 8).Value
        If i < 60 Then
            wsV.Cells(8 + i, 2).Value = wsO.Cells(r, 2).Value
            wsV.Cells(8 + i, 3).Value = CStr(wsO.Cells(r, 3).Value)
            If Not IsError(wsO.Cells(r, 10).Value) Then wsV.Cells(8 + i, 7).Value = wsO.Cells(r, 10).Value
        End If
    Next i
End Sub

' =====================================================================================
'  PowerPoint
' =====================================================================================

Public Sub PowerPointMaken()
    ' eerst het scenario bijwerken uit dit werkbestand, dan de presentatie maken
    CashflowNaarPowerPoint.OphalenUitFO "", True, True
    If Not CashflowNaarPowerPoint.OphalenGelukt Then Exit Sub
    CashflowNaarPowerPoint.NaarPowerPoint
End Sub

' =====================================================================================
'  hulpjes
' =====================================================================================

Private Function ZoekBlad(wb As Workbook, naam As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(naam), vbTextCompare) = 0 Then
            Set ZoekBlad = ws
            Exit Function
        End If
    Next ws
End Function

Private Function BestandsNaam(pad As String) As String
    Dim p As Long
    p = InStrRev(pad, "\")
    If InStrRev(pad, "/") > p Then p = InStrRev(pad, "/")
    If InStrRev(pad, ":") > p Then p = InStrRev(pad, ":")
    BestandsNaam = Mid$(pad, p + 1)
End Function

Private Function KolomLetter(n As Long) As String
    KolomLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, n).Address(True, False), "$")(0)
End Function
