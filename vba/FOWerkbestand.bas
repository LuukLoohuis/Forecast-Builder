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
'    PowerPoint maken       doorrekenen en Naar PowerPoint (Invoer, Woningtypes en Scenario kosten zijn formules op het FO)
'  Vaste plekken: zie tools/nieuwe_bladen.py.
' =====================================================================================

Private Const SH_START As String = "Start"
Private Const SH_AANSL As String = "Aansluiting opbrengsten"   ' was: Prijslijst (de aansluiting Salesforce <-> SAP)
Private Const SH_CF As String = "Invoer"              ' cashflow per periode (vervangt 1. Cashflow; kop 'Jaar' in B7)
Private Const SH_CF_OUD As String = "1. Cashflow"     ' in een oud FO
Private Const SH_VORIGE As String = "CF vorig en huidig"
Private Const SH_CONTROLE As String = "Controle oud FO"
Private Const SH_SJABLOON As String = "_sjabloon"     ' blad / cel / inhoud: na Oud FO inlezen terugzetten
Private Const TITEL As String = "FO-werkbestand"

Private Const RIJ_VORIG As Long = 15                 ' rijen op tab Start (blok 2)
Private Const RIJ_HUIDIG As Long = 16
Private Const RIJ_ACTUALS As Long = 17
Private Const RIJ_PRIJS As Long = 18
Private Const KOL_DATUM As Long = 6                  ' Start kolom F: laatst ingelezen
Private Const KOL_BESTAND As Long = 7                ' Start kolom G: bestand / klembord

' de opzet van het FO waarop dit werkbestand is gebouwd staat op _sjabloon kolom E:F (sleutel, waarde); de waarden
' hieronder zijn alleen de terugval voor de standaardopzet
Private Const TRANSPORTRIJEN As String = "13|26|36|46"   ' CF - opbrengsten: transportregel per woningtype
'@@CONSTANTEN@@

' tabbladen die 'Oud FO inlezen' overneemt (zelfde naam in het oude FO)
Private Const FO_BLADEN As String = "Toelichting onderlegger|Behind the scens|Prijslijst dump|Prijslijst|Aansluiting opbrengsten|" & _
    "CF - opbrengsten|CF - kosten|Verkooptempo|FO - vorig|FO - huidig|FO - actuals|1. Ontw businessc FD|" & _
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
    Knop ws, "J9:K11", "btnPowerPoint", "PowerPoint maken", "PowerPointMaken", donker
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
            blad = "FO - vorig": nKol = 21: hulpRij = 17
        Case RIJ_HUIDIG
            blad = "FO - huidig": nKol = 21: hulpRij = 17
        Case RIJ_ACTUALS
            blad = "FO - actuals": nKol = 16: hulpRij = 2
        Case Else
            blad = "Prijslijst dump": nKol = 9: hulpRij = 2
    End Select
    hulp = HulpKolommen(blad)
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
    ' bedragen die als tekst in Nederlandse notatie zijn geplakt ("157.397,00") als getal
    If h > 1 Then
        Dim plak As Variant, nPl As Long
        plak = ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol)).Value
        NLGetallenOmzetten plak, nPl
        If nPl > 0 Then ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol)).Value = plak
    End If
    Dim omgezet As String
    If rij = RIJ_VORIG Or rij = RIJ_HUIDIG Then omgezet = ExportNormaliseren(ws, nKol): h = LaatsteRij(ws, nKol)
    HulpBijwerken ws, hulp, hulpRij, h
    If rij = RIJ_ACTUALS Then WeergaveBijwerken
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ZetStatus rij, "klembord"
    ThisWorkbook.Worksheets(SH_START).Activate
    MsgBox "Klaar: " & h & " rijen in '" & blad & "'. Controleer de status op tab Start." & _
           IIf(Len(omgezet) > 0, vbCrLf & vbCrLf & omgezet, ""), vbInformation, TITEL
End Sub

Public Function ExportNormaliserenTest(ByVal blad As String) As String
    ' zonder dialogen (voor de test): de export op het tabblad naar de vaste indeling en de hulpkolommen bijwerken
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(blad)
    ExportNormaliserenTest = ExportNormaliseren(ws, 21)
    HulpBijwerken ws, HulpKolommen(blad), 17, LaatsteRij(ws, 21)
    Application.Calculate
End Function

Private Function IsNLGetalTekst(ByVal v As Variant) As Boolean
    ' tekst die zeker een getal in Nederlandse notatie is: duizendtallen met een punt ("157.397", "3.387.028") en/of
    ' decimalen met een komma ("157.397,00", "12,5"), eventueel met een minteken voor of achter. Gewone cijfers zonder
    ' scheidingsteken ("202609", "5100") en tekst met letters niet.
    Dim t As String, delen() As String, groepen() As String, i As Long
    If VarType(v) <> vbString Then Exit Function
    t = Replace(Replace(Trim$(CStr(v)), ChrW(160), ""), " ", "")
    If Len(t) = 0 Then Exit Function
    If Left$(t, 1) = "-" Or Left$(t, 1) = ChrW(8722) Then t = Mid$(t, 2)
    If Right$(t, 1) = "-" Then t = Left$(t, Len(t) - 1)
    If InStr(t, ".") = 0 And InStr(t, ",") = 0 Then Exit Function
    delen = Split(t, ",")
    If UBound(delen) > 1 Then Exit Function
    If UBound(delen) = 1 Then
        If Len(delen(1)) = 0 Or Not (delen(1) Like String(Len(delen(1)), "#")) Then Exit Function
    End If
    groepen = Split(delen(0), ".")
    For i = 0 To UBound(groepen)
        If Len(groepen(i)) = 0 Or Not (groepen(i) Like String(Len(groepen(i)), "#")) Then Exit Function
        If i = 0 And UBound(groepen) > 0 And Len(groepen(i)) > 3 Then Exit Function
        If i > 0 And Len(groepen(i)) <> 3 Then Exit Function
    Next i
    IsNLGetalTekst = True
End Function

Private Sub NLGetallenOmzetten(ByRef data As Variant, Optional ByRef n As Long)
    ' in een matrix (uit een export) alle teksten in Nederlandse getalnotatie omzetten naar getallen, vóórdat ze op een
    ' tabblad komen: Excel leest tekst die je vanuit VBA in een cel zet met Engelse instellingen (157.397 -> 157,397)
    Dim r As Long, c As Long, g As Double
    If Not IsArray(data) Then Exit Sub
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            If IsNLGetalTekst(data(r, c)) Then
                If NLGetal(data(r, c), g) Then data(r, c) = g: n = n + 1
            End If
        Next c
    Next r
End Sub

Private Function NLGetal(ByVal v As Variant, ByRef uit As Double) As Boolean
    ' tekst in Nederlandse notatie ("3.387.028", "798.210,00", "-2.483.005", "1.234-") naar een getal
    Dim t As String, i As Long, ch As String, neg As Boolean, schoon As String, komma As Long
    t = Trim$(CStr(v))
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "-" Then neg = True: t = Left$(t, Len(t) - 1)
    If Left$(t, 1) = "-" Or Left$(t, 1) = ChrW(8722) Then neg = Not neg: t = Mid$(t, 2)
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            schoon = schoon & ch
        ElseIf ch = "," Then
            komma = komma + 1
            schoon = schoon & "."
        ElseIf ch <> "." And ch <> " " Then
            Exit Function
        End If
    Next i
    If Len(schoon) = 0 Or komma > 1 Or schoon = "." Then Exit Function
    uit = Val(schoon)
    If neg Then uit = -uit
    NLGetal = True
End Function

Private Function ExportNormaliseren(ByVal ws As Worksheet, ByVal nKol As Long) As String
    ' SAP-export (FO - vorig / FO - huidig) naar de vaste indeling van het werkbestand: de kopregel 'Business object ...
    ' Plan Budget ...' op rij 15 met de bedragen vanaf kolom K, de boekingsperiode in E4 en de bedragen als getal.
    ' Een export die al zo staat blijft ongewijzigd. Geeft een zin terug over wat er is omgezet (leeg = niets).
    Const KOP_RIJ As Long = 15
    Const KOP_KOL As Long = 11
    Dim h As Long, a As Variant, b() As Variant, r As Long, c As Long, rA As Long, cP As Long, dr As Long, dc As Long
    Dim r2 As Long, c2 As Long, nG As Long, g As Double, t As String
    h = LaatsteRij(ws, nKol + 2)
    If h < 2 Then Exit Function
    a = ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol + 2)).Value
    For r = 1 To IIf(h < 40, h, 40)
        For c = 1 To 6
            If LCase$(Left$(Trim$(CStr(a(r, c))), 15)) = "business object" Then rA = r: Exit For
        Next c
        If rA > 0 Then Exit For
    Next r
    If rA = 0 Then
        ExportNormaliseren = "LET OP: de kopregel 'Business object' is niet gevonden; de export is niet omgezet."
        Exit Function
    End If
    For c = 1 To nKol + 2
        If LCase$(Left$(Trim$(CStr(a(rA, c))), 11)) = "plan budget" Then cP = c: Exit For
    Next c
    If cP = 0 Then
        ExportNormaliseren = "LET OP: de kolom 'Plan Budget' is niet gevonden; de export is niet omgezet."
        Exit Function
    End If
    dr = KOP_RIJ - rA
    dc = KOP_KOL - cP
    ReDim b(1 To h + IIf(dr > 0, dr, 0), 1 To nKol)
    For r = 1 To h
        r2 = r + dr
        If r2 >= 1 Then
            For c = 1 To nKol + 2
                ' t/m de kopregel blijft kolom A (de labels) staan; de waarden en de hele tabel schuiven mee
                If r <= rA And c = 1 Then c2 = 1 Else c2 = c + dc
                If c2 >= 1 And c2 <= nKol Then
                    If Not IsEmpty(a(r, c)) Then
                        If Len(CStr(a(r, c))) > 0 Then b(r2, c2) = a(r, c)
                    End If
                End If
            Next c
        End If
    Next r
    ' bedragen (K t/m S) als getal, en in alle kolommen tekst in Nederlandse getalnotatie (anders maakt Excel er bij het
    ' terugschrijven met Engelse instellingen 157,397 van)
    For r = KOP_RIJ + 1 To UBound(b, 1)
        For c = KOP_KOL To KOP_KOL + 8
            If VarType(b(r, c)) = vbString Then
                If NLGetal(b(r, c), g) Then b(r, c) = g: nG = nG + 1
            End If
        Next c
    Next r
    NLGetallenOmzetten b, nG
    If dr = 0 And dc = 0 And nG = 0 Then Exit Function
    ws.Range(ws.Cells(1, 1), ws.Cells(h + 2, nKol + 2)).ClearContents
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(b, 1), nKol)).Value = b
    If dr <> 0 Or dc <> 0 Then
        t = "De export stond in de SAP-indeling en is omgezet ("
        If dr <> 0 Then t = t & Abs(dr) & IIf(dr < 0, " rij omhoog", " rij omlaag")
        If dr <> 0 And dc <> 0 Then t = t & ", "
        If dc <> 0 Then t = t & Abs(dc) & IIf(dc > 0, " kolom naar rechts", " kolom naar links")
        t = t & IIf(nG > 0, "; " & nG & " bedragen naar getallen", "") & ")."
    Else
        t = nG & " bedragen stonden als tekst en zijn getallen geworden."
    End If
    ExportNormaliseren = t
End Function

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
    ' Local:=True: een SAP-export is vaak tekst (.xls als HTML/tab, of .csv). Zonder Local leest VBA die met Engelse
    ' instellingen en wordt "157.397,00" het getal 157,397 (punt = decimaal). Met Local gelden de Nederlandse instellingen.
    Set wbE = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True, Local:=True)
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
    NLGetallenOmzetten data                     ' vóór het schrijven, anders leest Excel "157.397" als 157,397
    LeegMaken ws, nKol
    ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Value = data
    Dim omgezet As String
    If rij = RIJ_VORIG Or rij = RIJ_HUIDIG Then omgezet = ExportNormaliseren(ws, nKol)
    HulpBijwerken ws, hulp, hulpRij, LaatsteRij(ws, nKol)
    If rij = RIJ_ACTUALS Then WeergaveBijwerken
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ZetStatus rij, BestandsNaam(CStr(pad))
    MsgBox "Klaar: " & nR & " rijen in '" & blad & "' uit " & BestandsNaam(CStr(pad)) & ". Controleer de status op tab Start." & _
           IIf(Len(omgezet) > 0, vbCrLf & vbCrLf & omgezet, ""), vbInformation, TITEL
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

Private Sub LeegMaken(ByVal ws As Worksheet, ByVal nKol As Long)
    Dim h As Long
    h = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If h < 1 Then h = 1
    ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol)).ClearContents
End Sub

Private Function LaatsteRij(ByVal ws As Worksheet, ByVal nKol As Long) As Long
    ' laatste rij met iets in de exportkolommen: vanaf de onderkant van het gebruikte bereik omhoog
    Dim m As Long
    m = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While m > 1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(m, 1), ws.Cells(m, nKol))) > 0 Then Exit Do
        m = m - 1
    Loop
    LaatsteRij = m
End Function

Private Function TypeVan(ByVal ws As Worksheet, ByVal r As Long, ByVal kType As Long, ByVal kBlok As Long, ByRef sl() As String, ByVal nS As Long) As Long
    ' zoals kolom U op Aansluiting opbrengsten: elk bouwnummer bij precies een type; exact op Type, dan Type in een lijst (met ;),
    ' dan exact op Termijnen / blok, dan blok in een lijst, dan '*'; bij gelijkspel het eerste type
    Dim k As Long, kol As Variant, w As String, x As Variant
    For Each kol In Array(kType, kBlok)
        If CLng(kol) > 0 Then
            w = LCase$(Tk(ws.Cells(r, CLng(kol)).Value))
            If Len(w) > 0 Then
                For k = 1 To nS
                    If LCase$(Trim$(sl(k))) = w Then TypeVan = k: Exit Function
                Next k
                For k = 1 To nS
                    For Each x In Split(sl(k), ";")
                        If LCase$(Trim$(CStr(x))) = w Then TypeVan = k: Exit Function
                    Next x
                Next k
            End If
        End If
    Next kol
    For k = 1 To nS
        If Trim$(sl(k)) = "*" Then TypeVan = k: Exit Function
    Next k
End Function

Private Function HulpKolommen(ByVal blad As String) As String
    ' hulpkolommen rechts van de export met de eerste rij met een formule, bv. "X:17|Y:30"
    Select Case blad
        Case "FO - vorig", "FO - huidig": HulpKolommen = Cfg("hulp|" & blad, "X:17|Y:17")
        Case "FO - actuals": HulpKolommen = Cfg("hulp|" & blad, "Q:2|R:2")
        Case Else: HulpKolommen = ""
    End Select
End Function

Private Sub HulpBijwerken(ByVal ws As Worksheet, ByVal hulp As String, ByVal hulpRij As Long, ByVal laatste As Long)
    ' per hulpkolom ("X:17|Y:30", kolom:eerste formulerij) de formule van de eerste rij doortrekken t/m de laatste
    ' exportrij (rijen daaronder houden hun formule); hulpRij geldt voor kolommen zonder eigen eerste rij
    Dim kolommen() As String, i As Long, f As String, k As String, r1 As Long
    If Len(hulp) = 0 Then Exit Sub
    kolommen = Split(hulp, "|")
    For i = LBound(kolommen) To UBound(kolommen)
        k = kolommen(i)
        r1 = hulpRij
        If InStr(k, ":") > 0 Then
            r1 = CLng(Split(k, ":")(1))
            k = Split(k, ":")(0)
        End If
        If laatste > r1 Then
            f = ws.Range(k & r1).FormulaR1C1
            If Left$(f, 1) = "=" Then ws.Range(k & r1 & ":" & k & laatste).FormulaR1C1 = f
        End If
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
    Dim n As Long
    If MsgBox("Nieuw kwartaal starten:" & vbCrLf & vbCrLf & _
              "1. de cashflow wordt bewaard als vorige prognose (stippellijn in de grafiek, tabel vorige/huidig op het Dashboard);" & vbCrLf & _
              "2. de planning (start bouw per type) wordt bewaard als vorige planning;" & vbCrLf & _
              "3. FO - huidig wordt FO - vorig;" & vbCrLf & _
              "4. FO - huidig en FO - actuals worden leeggemaakt voor de nieuwe exports." & vbCrLf & vbCrLf & _
              "Je invoer op Aansluiting opbrengsten, CF - opbrengsten en CF - kosten blijft staan." & vbCrLf & vbCrLf & _
              "Doe dit pas als het vorige FO helemaal klaar is. Doorgaan?", vbQuestion + vbYesNo, "Nieuw kwartaal starten") <> vbYes Then Exit Sub
    n = NieuwKwartaalUitvoeren()
    MsgBox "Klaar: " & n & " periodes en de planning bewaard als vorige prognose; FO - huidig staat nu in FO - vorig." & vbCrLf & vbCrLf & _
           "Volgende stappen:" & vbCrLf & _
           "1. lees de nieuwe exports in (FO - huidig, FO - actuals en de prijslijst);" & vbCrLf & _
           "2. werk de planning, verkoop en transport bij op CF - opbrengsten en de kosten op CF - kosten;" & vbCrLf & _
           "3. kijk op Start of alle controles OK zijn.", vbInformation, "Nieuw kwartaal starten"
End Sub

Public Function NieuwKwartaalUitvoeren() As Long
    ' het werk van Nieuw kwartaal starten, zonder vragen en meldingen (ook voor de test buiten Excel)
    Dim wsH As Worksheet, wsV As Worksheet, wsA As Worksheet, h As Long, n As Long
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    n = BewaarVorigePrognose()
    On Error Resume Next
    CashflowNaarPowerPoint.PlanningBewarenAlsVorig True
    On Error GoTo 0
    Set wsH = ThisWorkbook.Worksheets("FO - huidig")
    Set wsV = ThisWorkbook.Worksheets("FO - vorig")
    Set wsA = ThisWorkbook.Worksheets("FO - actuals")
    h = LaatsteRij(wsH, 21)
    LeegMaken wsV, 21
    wsV.Range("A1:U" & h).Value = wsH.Range("A1:U" & h).Value
    HulpBijwerken wsV, HulpKolommen("FO - vorig"), 17, h
    LeegMaken wsH, 21
    LeegMaken wsA, 16
    With ThisWorkbook.Worksheets(SH_START)
        .Cells(RIJ_VORIG, KOL_DATUM).Value = Now
        .Cells(RIJ_VORIG, KOL_BESTAND).Value = "uit FO - huidig (nieuw kwartaal)"
        .Range(.Cells(RIJ_HUIDIG, KOL_DATUM), .Cells(RIJ_ACTUALS, KOL_BESTAND)).ClearContents
        ' huidig kwartaal (rij 8): zelf ingevuld = een kwartaal verder; de formule volgt straks de nieuwe exports
        If Not .Range("C8").HasFormula And Not .Range("D8").HasFormula Then
            If IsNumeric(.Range("C8").Value) And IsNumeric(.Range("D8").Value) And Len(CStr(.Range("C8").Value)) > 0 Then
                If CLng(.Range("D8").Value) >= 4 Then
                    .Range("C8").Value = CLng(.Range("C8").Value) + 1: .Range("D8").Value = 1
                Else
                    .Range("D8").Value = CLng(.Range("D8").Value) + 1
                End If
            End If
        End If
    End With
    Application.EnableEvents = True
    Application.Calculate
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_START).Activate
    WeergaveBijwerken
    NieuwKwartaalUitvoeren = n
End Function

Private Function BewaarVorigePrognose() As Long
    ' tab Invoer -> tab Vorige prognose (waarden); kolom J (CF vorig kwartaal) op tab Invoer leest daaruit
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
    If MsgBox("Het FO wordt omgezet naar dit werkbestand: exports, prijslijst, woningtypes en termijnschema, planning, " & _
              "verkocht en getransporteerd, kosten, DAEB en de FD-kolom van 2. Ontw PEW. De invoer van dit werkbestand wordt " & _
              "overschreven met die uit:" & vbCrLf & CStr(pad) & vbCrLf & vbCrLf & _
              "Tab Controle oud FO laat daarna per periode zien of de cashflow gelijk is. Doorgaan?", _
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
    ' omzetten naar de vaste opzet van dit werkbestand (exports, prijslijst, woningtypes, planning, verkoop, transport,
    ' kosten, DAEB, PEW); het oude bestand verandert niet
    stap = "omzetten"
    waarschuwing = OmzettenOudFO(wbO)
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
    WeergaveBijwerken

    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets(SH_CONTROLE).Visible = -1
    ThisWorkbook.Worksheets(SH_CONTROLE).Activate
    bericht = "Klaar: " & BestandsNaam(CStr(pad)) & " is omgezet." & vbCrLf & vbCrLf
    If verschil = 0 Then
        bericht = bericht & "De cashflow is per periode gelijk aan het oude FO (tab Controle oud FO)."
    Else
        bericht = bericht & verschil & " periode(s) wijken af van het oude FO: zie tab Controle oud FO."
    End If
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

Private Sub ControleVullen(ByVal wbO As Workbook, ByVal pad As String)
    ' cashflow van het oude FO (1. Cashflow, waarden) naar tab Controle oud FO (per jaar en kwartaal) en kolom
    ' 'CF vorig kwartaal' naar tab Vorige prognose, zodat tab Invoer kolom J doorloopt
    Dim wsO As Worksheet, wsK As Worksheet, wsV As Worksheet, i As Long, r As Long
    Set wsK = ThisWorkbook.Worksheets(SH_CONTROLE)
    Set wsV = ThisWorkbook.Worksheets(SH_VORIGE)
    wsK.Range("B9:D60,G9:G60,J9:J60").ClearContents
    wsK.Range("D5").Value = BestandsNaam(pad)
    wsK.Range("D6").Value = Now
    Set wsO = ZoekBlad(wbO, SH_CF_OUD)
    If wsO Is Nothing Then Set wsO = ZoekBlad(wbO, "Cashflow")
    If wsO Is Nothing Then Exit Sub
    wsV.Range("B8:G67").ClearContents
    wsV.Range("D5").Value = "oud FO: " & BestandsNaam(pad)
    For i = 0 To 51
        r = CLng(Cfg("cashflow_rij1", "4")) + i
        If Not IsNumeric(wsO.Cells(r, 2).Value) Or IsEmpty(wsO.Cells(r, 2).Value) Then Exit For
        wsK.Cells(9 + i, 2).Value = wsO.Cells(r, 2).Value
        wsK.Cells(9 + i, 3).Value = Kwartaal(wsO.Cells(r, 3).Value)
        If Not IsError(wsO.Cells(r, 4).Value) Then wsK.Cells(9 + i, 4).Value = wsO.Cells(r, 4).Value
        If Not IsError(wsO.Cells(r, 6).Value) Then wsK.Cells(9 + i, 7).Value = wsO.Cells(r, 6).Value
        If Not IsError(wsO.Cells(r, 8).Value) Then wsK.Cells(9 + i, 10).Value = wsO.Cells(r, 8).Value
        If i < 60 Then
            wsV.Cells(8 + i, 2).Value = wsO.Cells(r, 2).Value
            wsV.Cells(8 + i, 3).Value = Kwartaal(wsO.Cells(r, 3).Value)
            If Not IsError(wsO.Cells(r, 10).Value) Then wsV.Cells(8 + i, 7).Value = wsO.Cells(r, 10).Value
        End If
    Next i
End Sub

' =====================================================================================
'  PowerPoint
' =====================================================================================

Public Sub PowerPointMaken()
    ' het scenariomodel rekent met formules mee met het FO: doorrekenen en de presentatie maken. Gewoon doorrekenen is genoeg
    ' (geen vluchtige functies of externe koppelingen; bij het openen rekent Excel alles door); staat het rekenen op
    ' handmatig, dan alles opnieuw. De rekentijd staat in de eindmelding.
    Dim t0 As Double
    t0 = Timer
    If Application.Calculation = -4135 Then Application.CalculateFull Else Application.Calculate
    CashflowNaarPowerPoint.RekenTijd = Timer - t0
    CashflowNaarPowerPoint.NaarPowerPoint
End Sub

' =====================================================================================
'  hulpjes
' =====================================================================================

Public Function WerkbestandCfg(ByVal sleutel As String, ByVal standaard As String) As String
    WerkbestandCfg = Cfg(sleutel, standaard)
End Function

Private Function Cfg(ByVal sleutel As String, ByVal standaard As String) As String
    ' opzet van het FO (_sjabloon kolom E: sleutel, F: waarde), anders de standaardwaarde
    Dim ws As Worksheet, r As Variant
    Cfg = standaard
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SJABLOON)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    r = Application.Match(sleutel, ws.Range("E1:E200"), 0)
    If Not IsError(r) Then
        If Len(CStr(ws.Cells(r, 6).Value)) > 0 Then Cfg = CStr(ws.Cells(r, 6).Value)
    End If
End Function

Private Function Kwartaal(ByVal v As Variant) As String
    ' 'Q1'..'Q4'; een jaarrij (leeg of het jaartal in de kolom Q) wordt leeg
    Dim t As String
    If IsError(v) Then Exit Function
    t = UCase$(Trim$(CStr(v)))
    If t Like "Q[1-4]" Then Kwartaal = t
End Function

Private Function ZoekBlad(ByVal wb As Workbook, ByVal naam As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(naam), vbTextCompare) = 0 Then
            Set ZoekBlad = ws
            Exit Function
        End If
    Next ws
End Function

Private Function BestandsNaam(ByVal pad As String) As String
    Dim p As Long
    p = InStrRev(pad, "\")
    If InStrRev(pad, "/") > p Then p = InStrRev(pad, "/")
    If InStrRev(pad, ":") > p Then p = InStrRev(pad, ":")
    BestandsNaam = Mid$(pad, p + 1)
End Function

Private Function KolomLetter(ByVal n As Long) As String
    KolomLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, n).Address(True, False), "$")(0)
End Function

' =====================================================================================
'  Oud FO omzetten naar de vaste opzet van dit werkbestand (zelfde regels als tools/omzetten.py)
' =====================================================================================

Public Sub OudFOOmzettenTest(ByVal pad As String)
    ' zonder vragen en meldingen (voor de controle buiten Excel)
    Dim wbO As Workbook, logTekst As String
    Set wbO = Workbooks.Open(pad, UpdateLinks:=0, ReadOnly:=True)
    logTekst = OmzettenOudFO(wbO)
    ControleVullen wbO, pad
    wbO.Close SaveChanges:=False
    WeergaveBijwerken
    ThisWorkbook.Worksheets(SH_START).Range("B4").Value = logTekst
End Sub

Private Function Tk(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Tk = Trim$(CStr(v))
End Function

Private Function Nk(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then Nk = CDbl(v)
End Function

Private Function IsGetalW(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Or VarType(v) = vbBoolean Then Exit Function
    IsGetalW = IsNumeric(v)
End Function

Private Function KwSleutels(ByVal ws As Worksheet, ByVal rj As Long, ByVal rq As Long, ByVal c1 As Long, ByVal c2 As Long) As Variant
    ' kolom -> kwartaalindex (jaar*4+q), 0 voor kolommen zonder kwartaal
    Dim a() As Long, c As Long, jaar As Long, q As String
    ReDim a(1 To c2)
    For c = c1 To c2
        If IsGetalW(ws.Cells(rj, c).Value) Then jaar = CLng(ws.Cells(rj, c).Value)
        q = UCase$(Tk(ws.Cells(rq, c).Value))
        If jaar > 0 And q Like "Q[1-4]" Then a(c) = jaar * 4 + CLng(Mid$(q, 2, 1))
    Next c
    KwSleutels = a
End Function

Private Function KolStd(ByVal sleutel As Long, ByVal k1 As Long) As Long
    ' kolom in dit werkbestand: actuals (G) voor alles vóór het eerste kwartaal, 0 = valt na het laatste kwartaal
    If sleutel = 0 Or sleutel < k1 Then
        KolStd = C_ACT
    ElseIf sleutel - k1 < NQ Then
        KolStd = C_Q1 + sleutel - k1
    End If
End Function

Private Sub RijOver(ByVal wsB As Worksheet, ByVal rB As Long, ByVal sl As Variant, ByVal c1 As Long, ByVal c2 As Long, ByVal wsD As Worksheet, ByVal rD As Long, _
                    ByVal k1 As Long, ByVal alleenQ As Boolean, ByVal deler As Double, ByRef logTekst As String)
    ' waarden van een rij in het oude FO naar rij rD (per kwartaal; niet-kwartalen opgeteld in actuals); deel <> 0 = delen
    Dim som() As Double, gezet() As Boolean, c As Long, k As Long, v As Variant
    ReDim som(1 To C_Q1 + NQ): ReDim gezet(1 To C_Q1 + NQ)
    For c = c1 To c2
        v = wsB.Cells(rB, c).Value
        If IsGetalW(v) Then
            k = KolStd(CLng(sl(c)), k1)
            If k = 0 Then
                If CDbl(v) <> 0 Then logTekst = logTekst & "- " & wsD.Name & " rij " & rD & ": " & v & " na het laatste kwartaal valt weg" & vbCrLf
            ElseIf Not (alleenQ And k = C_ACT) Then
                som(k) = som(k) + CDbl(v)
                gezet(k) = True
            End If
        End If
    Next c
    For k = C_ACT To C_Q1 + NQ - 1
        If gezet(k) And som(k) <> 0 Then
            If deler <> 0 Then wsD.Cells(rD, k).Value = som(k) / deler Else wsD.Cells(rD, k).Value = som(k)
        End If
    Next k
End Sub

Private Function RijMet(ByVal ws As Worksheet, ByVal kol As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal begin As String) As Long
    Dim r As Long
    For r = r1 To r2
        If LCase$(Tk(ws.Cells(r, kol).Value)) Like begin Then
            RijMet = r
            Exit Function
        End If
    Next r
End Function

Private Function OmzettenOudFO(ByVal wbO As Workbook) As String
    Dim lg As String, wO As Worksheet, wK As Worksheet, wP As Worksheet, wD As Worksheet, ws As Worksheet
    Dim maxR As Long, vk As Long, r As Long, c As Long, i As Long, j As Long, k As Long, n As Long
    Dim kops(1 To 20) As Long, omz(1 To 20) As Long, nKop As Long, nOmz As Long, laatste As Long, euro As Long, per As Long
    Dim sl As Variant, jaar As Long, k1 As Long, cf As Worksheet, hb As Long, won As Double, d As Long
    Dim nPlan As Long, oRij(1 To 30) As Long, nO As Long, lab As String, term(1 To 20) As Long, aantT As Long
    Dim kmw25 As Long, kmw75 As Long, kad As Long, ovr(1 To 10) As Long, nOv As Long, a0 As Long
    Dim oudW(1 To 5) As Double, plW(1 To 5) As Variant, sleutel As String, kopP As Long, kType As Long
    Dim bn() As Long, nBn As Long, x As Variant, basis As Double, tr As Long, t As Long, ov2() As Double, ov1() As Double, ovp() As Double, ovc() As Double, _
        euR(1 To 20) As Long, nEu As Long, eind As Long, heeftOv1 As Boolean, heeftOv2 As Boolean
    Dim regels As Variant, ond As Long, kr As Long, kb As Long, ksl As Variant, kc2 As Long

    Set wO = ZoekBlad(wbO, "CF - opbrengsten")
    If wO Is Nothing Then
        OmzettenOudFO = "- CF - opbrengsten ontbreekt in het oude FO: niets omgezet." & vbCrLf
        Exit Function
    End If
    Set wK = ThisWorkbook.Worksheets("CF - kosten")
    Set wP = ThisWorkbook.Worksheets(SH_AANSL)
    Set wD = ThisWorkbook.Worksheets("CF - opbrengsten")

    ' ---- exports: de export als waarden, de hulpkolommen houden hun formule -----------------------------------
    Dim ex As Variant, nKol As Long, h As Long, arr As Variant
    For Each ex In Array("FO - vorig|21", "FO - huidig|21", "FO - actuals|16", "Prijslijst dump|9")
        Set ws = ZoekBlad(wbO, Split(ex, "|")(0))
        nKol = CLng(Split(ex, "|")(1))
        If ws Is Nothing Then
            lg = lg & "- '" & Split(ex, "|")(0) & "' staat niet in het oude FO." & vbCrLf
        Else
            With ThisWorkbook.Worksheets(ws.Name)
                LeegMaken ThisWorkbook.Worksheets(ws.Name), nKol
                h = LaatsteRij(ws, nKol)
                If h >= 1 Then
                    arr = ws.Range(ws.Cells(1, 1), ws.Cells(h, nKol)).Value
                    .Range(.Cells(1, 1), .Cells(h, nKol)).Value = arr
                    ' tekst die op een getal lijkt (periode, jaar uit SAP) als tekst laten, zoals in de export
                    For r = 1 To h
                        For c = 1 To nKol
                            If VarType(arr(r, c)) = vbString Then
                                If Len(arr(r, c)) > 0 And IsNumeric(arr(r, c)) Then .Cells(r, c).Value = "'" & arr(r, c)
                            End If
                        Next c
                    Next r
                End If
                HulpBijwerken ThisWorkbook.Worksheets(ws.Name), HulpKolommen(ws.Name), IIf(nKol = 21, 17, 2), h
            End With
        End If
    Next ex
    Set ws = ZoekBlad(wbO, "Behind the scens")
    If Not ws Is Nothing Then
        With ThisWorkbook.Worksheets("Behind the scens")
            .UsedRange.ClearContents
            .Range(ws.UsedRange.Address).Formula = ws.UsedRange.Formula
            ' CODA-posten die als tekst staan ("5100") als tekst houden: kolom Y van FO - vorig/huidig zoekt op tekst
            Dim bc As Range
            For Each bc In ws.UsedRange.Cells
                If Not bc.HasFormula And VarType(bc.Value) = vbString Then
                    If Len(bc.Value) > 0 And IsNumeric(bc.Value) Then .Range(bc.Address).Value = "'" & bc.Value
                End If
            Next bc
        End With
    End If

    ' ---- opzet van het oude CF - opbrengsten ---------------------------------------------------------------
    maxR = wO.UsedRange.Row + wO.UsedRange.Rows.Count - 1
    vk = RijMet(wO, 2, 2, maxR, "verkooptempo")
    For r = 2 To maxR
        If LCase$(Tk(wO.Cells(r - 1, 6).Value)) = "# won" Then
            If r < vk Then
                nKop = nKop + 1: kops(nKop) = r
            Else
                nOmz = nOmz + 1: omz(nOmz) = r
            End If
        End If
    Next r
    If nKop > 0 And nOmz < nKop Then
        ' terugval: een omzetblok zonder '# won' in kolom F (wel het aantal in G erboven), herkend aan de kop '%' / EUR,
        ' na het verkooptempo en voor het financiele blok ('Financieel' / 'Termijnen 1 in EUR')
        eind = maxR
        For r = vk + 1 To maxR
            lab = LCase$(Tk(wO.Cells(r, 2).Value))
            If Left$(lab, 10) = "financieel" Or (Left$(lab, 9) = "termijnen" And InStr(lab, ChrW(8364)) > 0) Then eind = r: Exit For
        Next r
        nOmz = 0
        For r = vk + 1 To eind - 1
            If LCase$(Tk(wO.Cells(r - 1, 6).Value)) = "# won" Then
                nOmz = nOmz + 1: omz(nOmz) = r
            ElseIf Len(Tk(wO.Cells(r, 2).Value)) > 0 And Left$(Tk(wO.Cells(r, 3).Value), 1) = "%" And _
                   Tk(wO.Cells(r, 4).Value) = ChrW(8364) And IsGetalW(wO.Cells(r - 1, 7).Value) Then
                nOmz = nOmz + 1: omz(nOmz) = r
            End If
        Next r
    End If
    If nKop = 0 Or nKop <> nOmz Then
        OmzettenOudFO = lg & "- CF - opbrengsten: " & nKop & " bouwplanningsblokken en " & nOmz & " omzetblokken; niet omgezet." & vbCrLf
        Exit Function
    End If
    If nKop > NT Then lg = lg & "- " & nKop & " woningtypes; dit werkbestand heeft er " & NT & ": de rest valt weg." & vbCrLf
    For c = 7 To 80
        If UCase$(Tk(wO.Cells(kops(1), c).Value)) Like "Q[1-4]" Then laatste = c
    Next c
    For r = omz(nOmz) To maxR
        lab = LCase$(Tk(wO.Cells(r, 2).Value))
        If Left$(lab, 9) = "termijnen" And InStr(lab, ChrW(8364)) > 0 Then euro = r: Exit For
    Next r
    per = RijMet(wO, 2, euro, maxR, "periode in cashflow*")
    sl = KwSleutels(wO, euro - 1, euro, 7, laatste)
    ' startjaar: eerste kwartaal in 1. Cashflow, anders van CF - opbrengsten
    Set cf = ZoekBlad(wbO, "1. Cashflow")
    If Not cf Is Nothing Then
        For r = 4 To 80
            If Tk(cf.Cells(r, 3).Value) Like "Q[1-4]" And IsGetalW(cf.Cells(r, 2).Value) Then jaar = CLng(cf.Cells(r, 2).Value): Exit For
        Next r
    End If
    If jaar = 0 Then
        For c = 7 To laatste
            If sl(c) > 0 Then jaar = (sl(c) - 1) \ 4: Exit For
        Next c
    End If
    k1 = jaar * 4 + 1
    wD.Range("D3").Value = jaar
    ' start bouw en oplevering: de markeringen boven de kwartalen in het oude FO -> tab Start (rij 7 en 9)
    For r = IIf(kops(1) > 3, kops(1) - 3, 1) To kops(1) - 1
        For c = 7 To laatste
            lab = LCase$(Replace(Tk(wO.Cells(r, c).Value), " ", ""))
            If (lab = "startbouw" Or lab = "oplevering") And sl(c) > 0 Then
                With ThisWorkbook.Worksheets(SH_START)
                    .Cells(IIf(lab = "startbouw", 7, 9), 3).Value = (sl(c) - 1) \ 4
                    .Cells(IIf(lab = "startbouw", 7, 9), 4).Value = sl(c) - 4 * ((sl(c) - 1) \ 4)
                End With
            End If
        Next c
    Next r
    If per > 0 Then wD.Cells(PER_RIJ, 5).Value = wO.Cells(per, 5).Value

    ' ---- prijslijst per bouwnummer -------------------------------------------------------------------------
    Set ws = ZoekBlad(wbO, "Prijslijst")
    If Not ws Is Nothing Then
        kopP = RijMet(ws, 1, 1, 30, "bouwnummer")
        If kopP > 0 Then
            kType = KolNaam(ws, kopP, "type")
            If LCase$(Tk(ws.Range("A1").Value)) Like "datum*" Then wP.Range("C1").Value = ws.Range("C1").Value
            wP.Range(wP.Cells(P_BN1, 1), wP.Cells(P_BNN, 3)).ClearContents
            ReDim bn(1 To ws.UsedRange.Rows.Count + 10)
            For r = kopP + 1 To ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
                If Len(Tk(ws.Cells(r, 1).Value)) > 0 And nBn < P_BNN - P_BN1 + 1 Then
                    nBn = nBn + 1: bn(nBn) = r
                    wP.Cells(P_BN1 + nBn - 1, 1).Value = ws.Cells(r, 1).Value
                    For Each x In Array("termijnen / blok|2", "type|3", "kadastrale kosten|10", "bouwrente|17", "aanvullende rente|18", "korting|19")
                        c = KolNaam(ws, kopP, CStr(Split(x, "|")(0)))
                        If c > 0 Then wP.Cells(P_BN1 + nBn - 1, CLng(Split(x, "|")(1))).Value = ws.Cells(r, c).Value
                    Next x
                End If
            Next r
            lg = lg & "- prijslijst: " & nBn & " bouwnummers" & vbCrLf
        End If
    End If

    ' ---- per woningtype -------------------------------------------------------------------------------------
    Dim wA As Worksheet, sleutels(1 To 20) As String, nS As Long, f As String, p1 As Long, p2 As Long
    Set wA = ZoekBlad(wbO, "Aansluiting opbrengsten")
    If Not wA Is Nothing Then
        For r = 1 To 40
            f = wA.Cells(r, 3).Formula
            p1 = InStr(1, f, "SUMIF(Prijslijst!", vbTextCompare)
            If p1 > 0 Then
                p1 = InStr(p1, f, Chr$(34))
                p2 = InStr(p1 + 1, f, Chr$(34))
                If p1 > 0 And p2 > p1 And nS < 20 Then nS = nS + 1: sleutels(nS) = Mid$(f, p1 + 1, p2 - p1 - 1)
            End If
        Next r
    End If
    Dim typeNaam As String, viaParke As Boolean, viaCommerc As Boolean, overslaan As Boolean
    ReDim ov1(1 To C_Q1 + NQ): ReDim ov2(1 To C_Q1 + NQ): ReDim ovp(1 To C_Q1 + NQ): ReDim ovc(1 To C_Q1 + NQ)
    ' EUR-blokken per type ('Termijnen n in EUR'): een type met woningen maar zonder bedragen daar (parkeren, commercieel)
    ' heeft zijn omzet onder 'Overige opbrengsten in EUR'; dan prijs 0 (de kosten blijven)
    nEu = 0
    For r = euro To IIf(per > 0, per, maxR)
        lab = LCase$(Tk(wO.Cells(r, 2).Value))
        If Left$(lab, 9) = "termijnen" And InStr(lab, ChrW(8364)) > 0 And nEu < 20 Then nEu = nEu + 1: euR(nEu) = r
    Next r
    For i = 1 To nKop
        If i > NT Then Exit For
        hb = kops(i)
        won = Nk(wO.Cells(hb - 1, 7).Value)
        If won = 0 Then
            lg = lg & "- type " & i & ": geen woningen, niet overgenomen" & vbCrLf
        Else
            ' naam: na ' - ' in de titel, anders kolom C, anders na 'Termijnen '
            lab = Tk(wO.Cells(hb, 2).Value)
            If InStr(lab, " - ") > 0 Then
                lab = Trim$(Mid$(lab, InStr(lab, " - ") + 3))
            ElseIf Len(Tk(wO.Cells(hb, 3).Value)) > 0 Then
                lab = Tk(wO.Cells(hb, 3).Value)
            ElseIf LCase$(Left$(lab, 10)) = "termijnen " Then
                lab = Trim$(Mid$(lab, 11))
            End If
            wP.Cells(P_TYPE1 + i - 1, 2).Value = lab
            sleutel = ""
            If i <= nS Then sleutel = sleutels(i)
            wP.Cells(P_TYPE1 + i - 1, 3).Value = IIf(sleutel = "", Empty, sleutel)
            ' planningsrijen en omzetrijen
            nPlan = 0
            r = hb + 1
            Do While Len(Tk(wO.Cells(r, 2).Value)) > 0 And r < hb + 40
                nPlan = nPlan + 1: r = r + 1
            Loop
            nO = 0
            r = omz(i) + 1
            Do While Len(Tk(wO.Cells(r, 2).Value)) > 0 And nO < 30
                nO = nO + 1: oRij(nO) = r: r = r + 1
            Loop
            d = oRij(2) - (hb + 1)
            aantT = 0: kmw25 = 0: kmw75 = 0: kad = 0: nOv = 0
            For j = 2 To nO
                lab = LCase$(Tk(wO.Cells(oRij(j), 2).Value))
                If InStr(lab, "meerwerk") > 0 And InStr(lab, "25") > 0 Then
                    kmw25 = oRij(j)
                ElseIf InStr(lab, "meerwerk") > 0 And InStr(lab, "75") > 0 Then
                    kmw75 = oRij(j)
                ElseIf oRij(j) - d >= hb + 1 And oRij(j) - d <= hb + nPlan Then
                    If aantT < NTERM Then aantT = aantT + 1: term(aantT) = oRij(j)
                ElseIf InStr(lab, "kadast") > 0 Or InStr(lab, "rente") > 0 Then
                    kad = oRij(j)
                ElseIf nOv < 10 Then
                    nOv = nOv + 1: ovr(nOv) = oRij(j)
                End If
            Next j
            ' bedragen zoals het oude FO rekende; handmatig als de prijslijst ze niet verklaart
            oudW(1) = won
            oudW(2) = Nk(wO.Cells(oRij(1), 4).Value)
            basis = 0
            For j = 1 To aantT
                If Nk(wO.Cells(term(j), 3).Value) <> 0 Then basis = Nk(wO.Cells(term(j), 4).Value) / Nk(wO.Cells(term(j), 3).Value): Exit For
            Next j
            oudW(3) = basis
            If kmw25 > 0 Then
                oudW(4) = Nk(wO.Cells(kmw25, 4).Value) / 0.25
            ElseIf kmw75 > 0 Then
                oudW(4) = Nk(wO.Cells(kmw75, 4).Value) / 0.75
            Else
                oudW(4) = 0
            End If
            If kad > 0 Then oudW(5) = Nk(wO.Cells(kad, 4).Value) Else oudW(5) = 0
            For j = 1 To 5: plW(j) = Empty: Next j
            If sleutel <> "" And kType > 0 Then
                plW(1) = 0: plW(2) = 0: plW(3) = 0: plW(4) = 0: plW(5) = 0
                For n = 1 To nBn
                    If TypeVan(ws, bn(n), kType, KolNaam(ws, kopP, "termijnen / blok"), sleutels, nS) = i Then
                        plW(1) = plW(1) + 1
                        plW(2) = plW(2) + PrijsW(ws, bn(n), kopP, "grondkosten start verkoop (facturatie)")
                        plW(3) = plW(3) + PrijsW(ws, bn(n), kopP, "aanneemsom start verkoop") + PrijsW(ws, bn(n), kopP, "totaal woonwensen in von")
                        plW(4) = plW(4) + PrijsW(ws, bn(n), kopP, "totaal woonwensen in ob")
                        plW(5) = plW(5) + PrijsW(ws, bn(n), kopP, "kadastrale kosten") + PrijsW(ws, bn(n), kopP, "bouwrente") + _
                                 PrijsW(ws, bn(n), kopP, "aanvullende rente") + PrijsW(ws, bn(n), kopP, "korting")
                    End If
                Next n
            End If
            For j = 1 To 5
                If IsEmpty(plW(j)) Then
                    wP.Cells(P_TYPE1 + i - 1, 13 + j).Value = oudW(j)
                ElseIf Abs(plW(j) - oudW(j)) > 1 Then
                    wP.Cells(P_TYPE1 + i - 1, 13 + j).Value = oudW(j)
                    lg = lg & "- type " & i & ": bedrag " & j & " uit het oude FO wijkt af van de prijslijst; handmatig gezet" & vbCrLf
                End If
            Next j
            ' termijnschema
            a0 = P_AANS1 + P_BLOK * (i - 1)
            For j = 1 To aantT
                wP.Cells(a0 + 1 + j, 12).Value = Tk(wO.Cells(term(j), 2).Value)
                wP.Cells(a0 + 1 + j, 13).Value = Nk(wO.Cells(term(j), 3).Value)
            Next j
            ' planning, verkocht, transport, kadastrale kosten
            hb = 6 + BLOK * (i - 1)                 ' titelrij in dit werkbestand
            tr = hb + NTERM + 1
            For j = 1 To aantT
                wD.Range(wD.Cells(hb + j, C_ACT), wD.Cells(hb + j, C_Q1 + NQ - 1)).ClearContents
                RijOver wO, term(j) - d, sl, 7, laatste, wD, hb + j, k1, False, 0, lg
            Next j
            ' tabel verkoop en transport: per type verkocht (VK_RIJ + 2i - 1) en getransporteerd (VK_RIJ + 2i); kolom 1
            ' (t/m nu) is een formule op de Salesforce-status, de kwartalen komen uit het oude FO
            wD.Range(wD.Cells(VK_RIJ + 2 * i - 1, C_Q1), wD.Cells(VK_RIJ + 2 * i - 1, C_Q1 + NQ - 1)).ClearContents
            wD.Range(wD.Cells(VK_RIJ + 2 * i, C_Q1), wD.Cells(VK_RIJ + 2 * i, C_Q1 + NQ - 1)).ClearContents
            RijOver wO, vk + i, sl, 7, laatste, wD, VK_RIJ + 2 * i - 1, k1, True, 0, lg      ' verkoop in aantallen (verkooptempo)
            RijOver wO, oRij(1), sl, 7, laatste, wD, VK_RIJ + 2 * i, k1, True, 0, lg         ' transport in aantallen
            If sleutel = "" Then wD.Cells(VK_RIJ + 2 * i - 1, 4).Value = Nk(wO.Cells(vk + i, 4).Value)
            If kad > 0 Then RijOver wO, kad, sl, 7, laatste, wD, OZ1 + OBLOK * (i - 1) + 1 + NTERM + 3, k1, False, 0, lg
            For j = 1 To nOv
                basis = Nk(wO.Cells(ovr(j), 5).Value)
                For c = 7 To laatste
                    k = KolStd(CLng(sl(c)), k1)
                    If k > 0 Then ov2(k) = ov2(k) + Nk(wO.Cells(ovr(j), c).Value) * basis
                Next c
            Next j
            If i <= nEu Then
                basis = 0: r = euR(i) + 1
                Do While Len(Tk(wO.Cells(r, 2).Value)) > 0 And r < euR(i) + 30
                    For c = 7 To laatste: basis = basis + Abs(Nk(wO.Cells(r, c).Value)): Next c
                    r = r + 1
                Loop
                If basis = 0 Then
                    ' parkeren / commercieel met een eigen termijnschema: als echt type via de termijnen (de regel onder
                    ' Overige opbrengsten van het oude FO vervalt dan, anders telt het dubbel)
                    typeNaam = LCase$(Tk(wP.Cells(P_TYPE1 + i - 1, 2).Value))
                    If (InStr(typeNaam, "parke") > 0 Or InStr(typeNaam, "commerc") > 0) And Nk(wP.Cells(P_TYPE1 + i - 1, 16).Value) > 0 Then
                        If InStr(typeNaam, "parke") > 0 Then viaParke = True Else viaCommerc = True
                        lg = lg & "- type " & i & ": omzet via het termijnschema van het type; de regel onder Overige opbrengsten van het oude FO is niet overgenomen, de kwartalen kunnen afwijken van het oude FO" & vbCrLf
                    Else
                        For j = 2 To 5: wP.Cells(P_TYPE1 + i - 1, 13 + j).Value = 0: Next j
                        lg = lg & "- type " & i & ": geen bedragen in het " & ChrW(8364) & "-blok van het oude FO; de omzet staat onder Overige opbrengsten (prijs 0, de kosten blijven)" & vbCrLf
                    End If
                End If
            End If
            lg = lg & "- type " & i & " (" & wP.Cells(P_TYPE1 + i - 1, 2).Value & "): " & aantT & " termijnen, " & won & " woningen" & vbCrLf
            ' kosten: budget aanneemsom, KAO, % per termijn en eigen planning als die afwijkt
            KostenOmzetten wbO, wO, i, kops(i), term, aantT, d, sl, laatste, k1, lg
        End If
    Next i
    ' overige opbrengsten
    r = 0
    For j = euro To maxR
        If LCase$(Tk(wO.Cells(j, 2).Value)) Like "overige opbrengsten in*" Then r = j: Exit For
    Next j
    ' parkeren en commercieel naar hun eigen regel; de rest opgeteld als 'Overige opbrengsten (uit het oude FO)'
    If r > 0 Then
        For j = r + 1 To r + 3
            If Len(Tk(wO.Cells(j, 2).Formula)) = 0 Then Exit For
            lab = LCase$(Tk(wO.Cells(j, 2).Value) & " " & wO.Cells(j, 2).Formula)
            overslaan = (viaParke And InStr(LCase$(Tk(wO.Cells(j, 2).Value)), "parke") > 0) Or (viaCommerc And InStr(LCase$(Tk(wO.Cells(j, 2).Value)), "commerc") > 0)
            If Not overslaan Then
            For c = 7 To laatste
                k = KolStd(CLng(sl(c)), k1)
                If k > 0 Then
                    If InStr(lab, "parkeer") > 0 Then
                        ovp(k) = ovp(k) + Nk(wO.Cells(j, c).Value)
                    ElseIf InStr(lab, "commerc") > 0 Then
                        ovc(k) = ovc(k) + Nk(wO.Cells(j, c).Value)
                    Else
                        ov1(k) = ov1(k) + Nk(wO.Cells(j, c).Value)
                    End If
                End If
            Next c
            End If
        Next j
    End If
    ' de regel parkeren is bezet als er verschillen per type zijn; commercieel als er andere overige opbrengsten zijn
    For k = C_ACT To C_Q1 + NQ - 1
        If ov2(k) <> 0 Then heeftOv2 = True
        If ov1(k) <> 0 Then heeftOv1 = True
    Next k
    For k = C_ACT To C_Q1 + NQ - 1
        If heeftOv2 Then ov1(k) = ov1(k) + ovp(k): ovp(k) = 0
        If heeftOv1 Then ov1(k) = ov1(k) + ovc(k): ovc(k) = 0
    Next k
    For k = C_ACT To C_Q1 + NQ - 1
        If ov1(k) <> 0 Then wD.Cells(OV1_RIJ, k).Value = ov1(k): wD.Cells(OV1_RIJ, 2).Value = "Overige opbrengsten (uit het oude FO)": wD.Cells(OV1_RIJ - 2, 3).Value = "handmatig"
        If ovc(k) <> 0 Then wD.Cells(OV1_RIJ, k).Value = ovc(k): wD.Cells(OV1_RIJ, 2).Value = "Opbrengsten commercieel (uit het oude FO)": wD.Cells(OV1_RIJ - 2, 3).Value = "handmatig"
        If ov2(k) <> 0 Then wD.Cells(OV2_RIJ, k).Value = ov2(k): wD.Cells(OV2_RIJ, 2).Value = "Overige / verschillen per type (uit het oude FO)": wD.Cells(OV2_RIJ - 2, 3).Value = "handmatig"
        If ovp(k) <> 0 Then wD.Cells(OV2_RIJ, k).Value = ovp(k): wD.Cells(OV2_RIJ, 2).Value = "Opbrengsten parkeren (uit het oude FO)": wD.Cells(OV2_RIJ - 2, 3).Value = "handmatig"
    Next k
    ' aansluiting: SKO per regel (kolom B van de vergelijkingsblokken)
    If Not wA Is Nothing Then
        regels = Split(P_REGELS, "|")
        n = 0
        For r = 1 To wA.UsedRange.Row + wA.UsedRange.Rows.Count
            If LCase$(Left$(Tk(wA.Cells(r, 1).Value), 12)) = "vergelijking" Then
                n = n + 1
                If n <= NT Then
                    j = r + 2
                    Do While Len(Tk(wA.Cells(j, 1).Value)) > 0
                        For c = 0 To UBound(regels)
                            If Tk(wA.Cells(j, 1).Value) = regels(c) And IsGetalW(wA.Cells(j, 2).Value) Then
                                wP.Cells(P_AANS1 + P_BLOK * (n - 1) + 2 + c, 2).Value = wA.Cells(j, 2).Value
                            End If
                            ' regels zonder koppeling met de prijslijst: het bedrag uit de kolom Prijslijst van het oude FO
                            If Tk(wA.Cells(j, 1).Value) = regels(c) And InStr(1, "|VON verhogingen|Opbrengsten commercieel|Opbrengsten parkeren|", "|" & regels(c) & "|") > 0 Then
                                If IsGetalW(wA.Cells(j, 4).Value) Then
                                    If Nk(wA.Cells(j, 4).Value) <> 0 Then wP.Cells(P_AANS1 + P_BLOK * (n - 1) + 2 + c, 4).Value = wA.Cells(j, 4).Value
                                End If
                            End If
                        Next c
                        j = j + 1
                    Loop
                End If
            End If
        Next r
    End If
    ' DAEB
    DaebOmzetten wbO, wO, kops, nKop, sl, laatste, k1, lg
    ' kostenposten (prognose; Aanneemsom en Kopersmeerwerk volgen de termijnen)
    Dim wKo As Worksheet, posten As Long, toel As Long, pAct As Long, q As Long, weg As Double, verschoven As Double
    Set wKo = ZoekBlad(wbO, "CF - kosten")
    ' laatste afgesloten kwartaal volgens de SAP-export (FO - actuals is al ingelezen): jaar (P) en maand (O), zoals Python
    Dim wAc As Worksheet, jA As Variant, mA As Variant, perMax As Long
    Set wAc = ThisWorkbook.Worksheets("FO - actuals")
    For r = 2 To wAc.UsedRange.Row + wAc.UsedRange.Rows.Count - 1
        jA = wAc.Cells(r, 16).Value: mA = wAc.Cells(r, 15).Value
        If IsNumeric(jA) And IsNumeric(mA) And Len(Trim$(CStr(jA))) > 0 And Len(Trim$(CStr(mA))) > 0 Then
            If CLng(mA) >= 1 And CLng(mA) <= 12 Then
                If CLng(jA) * 4 + (CLng(mA) - 1) \ 3 + 1 > pAct Then pAct = CLng(jA) * 4 + (CLng(mA) - 1) \ 3 + 1
                If CLng(jA) * 100 + CLng(mA) > perMax Then perMax = CLng(jA) * 100 + CLng(mA)
            End If
        End If
    Next r
    ' loopt het kwartaal van de laatste boekingsperiode nog, dan is dat het huidige kwartaal en is het kwartaal ervoor afgesloten
    If (perMax Mod 100) Mod 3 <> 0 Then pAct = pAct - 1
    If Not wKo Is Nothing Then
        ond = RijMet(wKo, 2, 2, 300, "onderdelen")
        If ond > 0 Then
            kc2 = 0
            For c = 7 To 80
                If UCase$(Tk(wKo.Cells(ond, c).Value)) Like "Q[1-4]" Then kc2 = c
                If LCase$(Tk(wKo.Cells(ond, c).Value)) = "toelichting" Then toel = c
            Next c
            ksl = KwSleutels(wKo, ond - 1, ond, 7, kc2)
            r = ond + 1
            Do While Len(Tk(wKo.Cells(r, 2).Value)) > 0 And LCase$(Left$(Tk(wKo.Cells(r, 2).Value), 6)) <> "totaal" And posten < NPOST
                If toel > 0 Then wK.Cells(K_POST1 + posten, C_TOEL).Value = wKo.Cells(r, toel).Value
                lab = LCase$(Tk(wKo.Cells(r, 2).Value))
                If lab <> "aanneemsom" And lab <> "kopersmeerwerk" Then
                    wK.Range(wK.Cells(K_POST1 + posten, C_Q1), wK.Cells(K_POST1 + posten, C_Q1 + NQ - 1)).ClearContents
                    RijOver wKo, r, ksl, 7, kc2, wK, K_POST1 + posten, k1, True, 0, lg
                    ' prognose in een kwartaal dat volgens de SAP-export al is afgesloten (de cashflow neemt daar de
                    ' actuals): naar het eerste prognosekwartaal, anders telt die prognose nergens mee
                    If pAct >= k1 And KolStd(pAct + 1, k1) > 0 Then
                        weg = 0
                        For q = k1 To pAct
                            k = KolStd(q, k1)
                            If IsGetalW(wK.Cells(K_POST1 + posten, k).Value) Then
                                weg = weg + CDbl(wK.Cells(K_POST1 + posten, k).Value)
                                wK.Cells(K_POST1 + posten, k).ClearContents
                            End If
                        Next q
                        If weg <> 0 Then
                            k = KolStd(pAct + 1, k1)
                            wK.Cells(K_POST1 + posten, k).Value = Nk(wK.Cells(K_POST1 + posten, k).Value) + weg
                            verschoven = verschoven + weg
                        End If
                    End If
                End If
                posten = posten + 1
                r = r + 1
            Loop
        End If
    End If
    If verschoven <> 0 Then lg = lg & "- CF - kosten: " & Format(verschoven, "0") & " gepland in kwartalen die volgens de SAP-export al zijn afgesloten (t/m Q" & _
                               ((pAct - 1) Mod 4 + 1) & " " & ((pAct - 1) \ 4) & ") verschoven naar het eerste prognosekwartaal" & vbCrLf
    ' PEW: meest recente FD en toelichting
    PewOmzetten wbO, lg
    OmzettenOudFO = lg
End Function

Private Function PrijsW(ByVal ws As Worksheet, ByVal r As Long, ByVal kopP As Long, ByVal naam As String) As Double
    Dim c As Long
    c = KolNaam(ws, kopP, naam)
    If c > 0 Then PrijsW = Nk(ws.Cells(r, c).Value)
End Function

Private Function KolNaam(ByVal ws As Worksheet, ByVal r As Long, ByVal naam As String) As Long
    Dim c As Long
    For c = 1 To 40
        If LCase$(Tk(ws.Cells(r, c).Value)) = naam Then
            KolNaam = c
            Exit Function
        End If
    Next c
End Function

Private Sub KostenOmzetten(ByVal wbO As Workbook, ByVal wO As Worksheet, ByVal i As Long, ByVal hbOud As Long, term() As Long, ByVal aantT As Long, ByVal d As Long, _
                           ByVal sl As Variant, ByVal laatste As Long, ByVal k1 As Long, ByRef lg As String)
    Dim wKo As Worksheet, wK As Worksheet, kopTekst As String, r As Long, kr As Long, kRij(1 To 20) As Long, n As Long
    Dim ksl As Variant, kc2 As Long, c As Long, kb As Long, j As Long, afwijk As Boolean, c2 As Long, sK As Double, sP As Double
    Set wKo = ZoekBlad(wbO, "CF - kosten")
    If wKo Is Nothing Then Exit Sub
    Set wK = ThisWorkbook.Worksheets("CF - kosten")
    kopTekst = Tk(wO.Cells(hbOud, 2).Value)
    For r = 2 To 120
        If Tk(wKo.Cells(r, 2).Value) = kopTekst And Len(kopTekst) > 0 Then kr = r: Exit For
        If LCase$(Tk(wKo.Cells(r, 2).Value)) = "onderdelen" Then Exit For
    Next r
    If kr = 0 Then Exit Sub
    r = kr + 1
    Do While Len(Tk(wKo.Cells(r, 2).Value)) > 0 And n < 20
        If Left$(LCase$(Tk(wKo.Cells(r, 2).Value)), 10) = "aanneemsom" Or Left$(LCase$(Tk(wKo.Cells(r, 2).Value)), 8) = "meerwerk" Then Exit Do
        If InStr(LCase$(Tk(wKo.Cells(r, 2).Value)), "meerwerk") = 0 Then n = n + 1: kRij(n) = r
        r = r + 1
    Loop
    For c = 7 To 80
        If UCase$(Tk(wKo.Cells(kr, c).Value)) Like "Q[1-4]" Then kc2 = c
    Next c
    ksl = KwSleutels(wKo, kr - 1, kr, 7, kc2)
    kb = 3 + KBLOK * (i - 1)
    wK.Cells(kb, 4).Value = Nk(wKo.Cells(kr - 1, 4).Value)
    wK.Cells(kb, 6).Value = Nk(wKo.Cells(kr - 1, 6).Value)
    For j = 1 To n
        If j > NTERM Then Exit For
        wK.Cells(kb + 1 + j, 3).Value = Nk(wKo.Cells(kRij(j), 3).Value)
        ' eigen planning (HWB) alleen als die afwijkt van de planning van de opbrengsten
        afwijk = False
        sK = 0: sP = 0
        For c = 7 To kc2
            sK = sK + Nk(wKo.Cells(kRij(j), c).Value)
        Next c
        If j <= aantT Then
            For c = 7 To laatste
                sP = sP + Nk(wO.Cells(term(j) - d, c).Value)
                For c2 = 7 To kc2
                    If ksl(c2) = sl(c) And sl(c) > 0 Then
                        If Abs(Nk(wKo.Cells(kRij(j), c2).Value) - Nk(wO.Cells(term(j) - d, c).Value)) > 0.000000001 Then afwijk = True
                    End If
                Next c2
            Next c
        End If
        If Abs(sK - sP) > 0.000000001 Then afwijk = True
        If afwijk Then
            wK.Range(wK.Cells(kb + 1 + j, C_ACT), wK.Cells(kb + 1 + j, C_Q1 + NQ - 1)).ClearContents
            RijOver wKo, kRij(j), ksl, 7, kc2, wK, kb + 1 + j, k1, False, 0, lg
            lg = lg & "- type " & i & ": kostenplanning termijn " & j & " wijkt af van de opbrengsten: overgenomen" & vbCrLf
        End If
    Next j
End Sub

Private Sub DaebOmzetten(ByVal wbO As Workbook, ByVal wO As Worksheet, kops() As Long, ByVal nKop As Long, ByVal sl As Variant, ByVal laatste As Long, _
                         ByVal k1 As Long, ByRef lg As String)
    Dim wD As Worksheet, r0 As Long, r As Long, lab As String, comp As Long, g As Long, i As Long, aantal As Variant
    Dim rv As Long, c As Long, wv As Worksheet, kop As Long, rij As Long, acol As Long, vsl As Variant, vk As Long
    Set wD = ThisWorkbook.Worksheets("CF - opbrengsten")
    For r = 2 To wO.UsedRange.Row + wO.UsedRange.Rows.Count
        lab = LCase$(Tk(wO.Cells(r, 2).Value))
        If InStr(lab, "daeb") > 0 And Left$(lab, 9) <> "omzet hvg" And Left$(lab, 10) <> "termijnen " Then r0 = r: Exit For
    Next r
    If r0 = 0 Then Exit Sub
    For r = r0 - 3 To r0 - 1
        If Len(Tk(wO.Cells(r, 3).Value)) > 0 And Nk(wO.Cells(r, 5).Value) <> 0 And comp < 2 Then
            wD.Cells(DAEB_KOP_RIJ - 3 + comp, 3).Value = Tk(wO.Cells(r, 3).Value)
            wD.Cells(DAEB_KOP_RIJ - 3 + comp, 5).Value = Nk(wO.Cells(r, 5).Value)
            comp = comp + 1
        End If
    Next r
    g = 0: i = 0
    For r = r0 + 1 To r0 + 40
        lab = Tk(wO.Cells(r, 2).Value)
        If LCase$(Left$(lab, 6)) = "totaal" Then Exit For
        If Len(lab) = 0 Then
            If i > 0 Then g = g + 1: i = 0
            If Len(Tk(wO.Cells(r + 1, 2).Value)) = 0 Then Exit For
        ElseIf g < 2 And i < 8 Then
            wD.Cells(DAEB_KOP_RIJ + 1 + g * 9 + i, 2).Value = lab
            If IsGetalW(wO.Cells(r, 3).Value) Then wD.Cells(DAEB_KOP_RIJ + 1 + g * 9 + i, 3).Value = wO.Cells(r, 3).Value
            RijOver wO, r, sl, 7, laatste, wD, DAEB_KOP_RIJ + 1 + g * 9 + i, k1, False, 0, lg
            i = i + 1
        End If
    Next r
    ' aantal en verkoop: regels in het blok, anders Verkooptempo (de DAEB-rij of de rij na de koopwoningtypes)
    aantal = Empty
    For r = r0 To r0 + 60
        lab = LCase$(Tk(wO.Cells(r, 2).Value))
        If Left$(lab, 11) = "aantal daeb" Then aantal = wO.Cells(r, 4).Value
        If Left$(lab, 24) = "verkoop en levering daeb" Then rv = r
    Next r
    If rv > 0 Then RijOver wO, rv, sl, 7, laatste, wD, DAEB_VERKOOP_RIJ, k1, True, 0, lg
    Set wv = ZoekBlad(wbO, "Verkooptempo")
    If (IsEmpty(aantal) Or rv = 0) And Not wv Is Nothing Then
        kop = RijMet(wv, 1, 1, 30, "verkooptempo")
        If kop > 0 Then
            For r = kop + 1 To kop + 20
                If InStr(LCase$(Tk(wv.Cells(r, 1).Value)), "daeb") > 0 Then rij = r: Exit For
            Next r
            If rij = 0 Then
                r = kop + 1 + nKop
                If Len(Tk(wv.Cells(r, 1).Value)) > 0 And LCase$(Left$(Tk(wv.Cells(r, 1).Value), 6)) <> "totaal" Then rij = r
            End If
            If rij > 0 Then
                acol = 3
                For c = 1 To 40
                    If LCase$(Tk(wv.Cells(kop, c).Value)) = "aantal" Then acol = c: Exit For
                Next c
                If IsEmpty(aantal) Then aantal = wv.Cells(rij, acol).Value
                If rv = 0 Then
                    vsl = KwSleutels(wv, kop - 1, kop, 2, 40)
                    For c = 2 To 40
                        If vsl(c) > 0 And c <> acol Then
                            vk = KolStd(CLng(vsl(c)), k1)
                            If vk > 0 And Nk(wv.Cells(rij, c).Value) <> 0 Then wD.Cells(DAEB_VERKOOP_RIJ, vk).Value = wv.Cells(rij, c).Value
                        End If
                    Next c
                End If
            End If
        End If
    End If
    wD.Cells(DAEB_AANTAL_RIJ, 4).Value = aantal
    lg = lg & "- DAEB-blok overgenomen (" & comp & " componenten, aantal " & aantal & ")" & vbCrLf
End Sub

Private Sub PewOmzetten(ByVal wbO As Workbook, ByRef lg As String)
    Dim wO As Worksheet, wN As Worksheet, kop As Long, r As Long, c As Long, fd As Long, toel As Long, eind As Long
    Set wO = ZoekBlad(wbO, "2. Ontw PEW")
    If wO Is Nothing Then Exit Sub
    Set wN = ThisWorkbook.Worksheets("2. Ontw PEW")
    For r = 1 To 12
        For c = 3 To 16
            If LCase$(Tk(wO.Cells(r, c).Value)) = "huidig" Then kop = r
        Next c
        If kop > 0 Then Exit For
    Next r
    If kop = 0 Then Exit Sub
    For c = 3 To 20
        If LCase$(Tk(wO.Cells(kop, c).Value)) = "meest recente fd" Then fd = c
        If LCase$(Tk(wO.Cells(kop - 1, c).Value)) = "toelichting" Then toel = c
    Next c
    If fd = 0 Then fd = 5
    For r = kop + 1 To kop + 40
        If LCase$(Tk(wO.Cells(r, 3).Value)) = "totaal opbrengsten" Then eind = r: Exit For
    Next r
    For r = kop + 2 To eind - 1
        If LCase$(Tk(wO.Cells(r, 3).Value)) <> "totaal kosten" And Len(Tk(wO.Cells(r, 3).Value)) > 0 Then
            If IsGetalW(wO.Cells(r, fd).Value) Then wN.Cells(r, 5).Value = wO.Cells(r, fd).Value
            If toel > 0 Then wN.Cells(r, 11).Value = wO.Cells(r, toel).Value
        End If
    Next r
    ' 2B. Ontw PEW (gemengd project): A = concept SKO DAEB, B = concept SKO niet-DAEB en de toelichting, per post
    Dim wB As Worksheet, wB2 As Worksheet, ka As Long, kb As Long, kt As Long, kr As Long, lab As String, rr As Long, n2 As Long, t As String
    Set wB = ZoekBlad(wbO, "2B. Ontw PEW")
    If wB Is Nothing Then Exit Sub
    Set wB2 = ThisWorkbook.Worksheets("2B. Ontw PEW")
    For r = 1 To 11
        For c = 3 To 19
            t = LCase$(Tk(wB.Cells(r, c).Value))
            ' alleen kolomkoppen ('Concept SKO ...'), niet een toelichting die toevallig 'niet-DAEB' noemt
            If InStr(t, "sko") > 0 And (InStr(t, "niet-daeb") > 0 Or InStr(t, "niet daeb") > 0) And kb = 0 Then
                kb = c: kr = r
            ElseIf InStr(t, "sko") > 0 And InStr(t, "daeb") > 0 And InStr(t, "niet") = 0 And ka = 0 Then
                ka = c: kr = r
            ElseIf t = "toelichting" And kt = 0 Then
                kt = c
            End If
        Next c
    Next r
    If ka = 0 Or kb = 0 Then Exit Sub
    For r = kr + 1 To kr + 29
        lab = LCase$(Tk(wB.Cells(r, 3).Value))
        If Len(lab) > 0 And Left$(lab, 6) <> "totaal" And Left$(lab, 9) <> "resultaat" And Left$(lab, 5) <> "check" Then
            For rr = 6 To 28
                If LCase$(Tk(wB2.Cells(rr, 3).Value)) = lab Then
                    If IsGetalW(wB.Cells(r, ka).Value) Then wB2.Cells(rr, 6).Value = wB.Cells(r, ka).Value
                    If IsGetalW(wB.Cells(r, kb).Value) Then wB2.Cells(rr, 7).Value = wB.Cells(r, kb).Value
                    If kt > 0 Then wB2.Cells(rr, 14).Value = wB.Cells(r, kt).Value
                    n2 = n2 + 1
                    Exit For
                End If
            Next rr
        End If
    Next r
    lg = lg & "- 2B. Ontw PEW: A (DAEB) en B (niet-DAEB) overgenomen voor " & n2 & " posten" & vbCrLf
End Sub

Public Sub WeergaveBijwerken()
    ' CF - opbrengsten en CF - kosten: woningtypes zonder woningen ingeklapt; de kwartalen vóór de laatste actuals
    ' (tab Start) als ingeklapte kolomgroep (met + weer open te klappen)
    Dim wsO As Worksheet, wsK As Worksheet, ws As Variant, k As Long, leeg As Boolean, r1 As Long, n As Long, i As Long
    Dim k1 As Long, ka As Long
    ' 2B. Ontw PEW (DAEB / niet-DAEB): zichtbaar bij een project met DAEB-woningen of als A of B is ingevuld
    On Error Resume Next
    With ThisWorkbook.Worksheets("2B. Ontw PEW")
        .Visible = IIf(Nk(ThisWorkbook.Worksheets("Model").Range("BY135").Value) > 0 Or _
                       Application.WorksheetFunction.Count(.Range("F7:G26")) > 0, -1, 0)
    End With
    On Error GoTo 0
    Set wsO = ThisWorkbook.Worksheets("CF - opbrengsten")
    Set wsK = ThisWorkbook.Worksheets("CF - kosten")
    For k = 1 To NT
        leeg = (Nk(wsO.Cells(5 + BLOK * (k - 1), 7).Value) = 0)
        r1 = 6 + BLOK * (k - 1)
        wsO.Range(wsO.Rows(r1 + 1), wsO.Rows(r1 + BLOK - 2)).EntireRow.Hidden = leeg
        wsO.Rows(KANS_RIJ + k).Hidden = leeg
        wsO.Range(wsO.Rows(VK_RIJ + 2 * k - 1), wsO.Rows(VK_RIJ + 2 * k)).EntireRow.Hidden = leeg
        With ThisWorkbook.Worksheets(SH_AANSL)
            .Rows(P_TYPE1 + k - 1).Hidden = leeg
            .Range(.Rows(P_AANS1 + P_BLOK * (k - 1)), .Rows(P_AANS1 + P_BLOK * k - 2)).EntireRow.Hidden = leeg
        End With
        r1 = 3 + KBLOK * (k - 1)
        wsK.Range(wsK.Rows(r1 + 2), wsK.Rows(r1 + NTERM + 5)).EntireRow.Hidden = leeg
    Next k
    If Not IsGetalW(wsO.Range("D3").Value) Then Exit Sub
    k1 = CLng(wsO.Range("D3").Value) * 4 + 1
    ka = CLng(Nk(ThisWorkbook.Worksheets(SH_START).Range("F12").Value))
    n = ka - k1          ' het kwartaal van de laatste actuals (vaak het lopende) blijft zichtbaar
    If n > NQ Then n = NQ
    For Each ws In Array(wsO, wsK)
        On Error Resume Next
        For i = 1 To 3
            ws.Range(ws.Columns(C_Q1), ws.Columns(C_Q1 + NQ - 1)).EntireColumn.Ungroup
        Next i
        Err.Clear
        On Error GoTo 0
        ws.Range(ws.Columns(C_Q1), ws.Columns(C_Q1 + NQ - 1)).EntireColumn.Hidden = False
        If n > 0 Then
            ws.Range(ws.Columns(C_Q1), ws.Columns(C_Q1 + n - 1)).EntireColumn.Group
            ws.Range(ws.Columns(C_Q1), ws.Columns(C_Q1 + n - 1)).EntireColumn.Hidden = True
        End If
    Next ws
End Sub
