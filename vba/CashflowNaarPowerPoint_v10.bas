Attribute VB_Name = "CashflowNaarPowerPoint"
Option Explicit

' =====================================================================================
'  Cashflow_scenario -> PowerPoint                       (late binding, Windows en Mac)
'  versie 10: typeblokken (maximaal MAX_TYPES woningtypes), knoppen Type invoegen / Type verwijderen
'
'  NaarPowerPoint        opent het sjabloon als kopie, vult teksten, tabellen en grafieken
'                        vanaf tabblad "PowerPoint" en bewaart een nieuwe presentatie
'  ControleerSjabloon    kijkt of het sjabloon alle vormnamen heeft die de knop vult
'  KnoppenControleren    zet ontbrekende knoppen op de tabbladen (draait bij het openen)
'  TypeInvoegen          voegt een leeg typeblok in (drie kolommen op Woningtypes, twee op Invoer)
'  TypeVerwijderen       haalt een typeblok weg (zelfde kolommen); alles rechts schuift mee
'  SchikDia5             na het vullen van VT_TABEL (dia 5): bij meer dan 4 types rijen en letters
'                        kleiner, zo nodig de grafiekkaart korter, en de tabelkaart sluit om de tabel
'
'  Welke tekst naar welke vorm gaat staat op tabblad "PowerPoint", kolom F.
' =====================================================================================

Private Const SH_PP As String = "PowerPoint"
Private Const SH_DASH As String = "Dashboard"
Private Const SH_MODEL As String = "Model"
Private Const SH_TYPES As String = "Woningtypes"
Private Const SH_INVOER As String = "Invoer"
Private Const KPI_ROW1 As Long = 8            ' eerste regel met KPI-teksten
Private Const CEL_N As String = "BY8"              ' aantal periodes (tabblad Model)
Private Const CEL_SJABLOON As String = "D52"
Private Const CEL_NAAM As String = "D53"
Private Const SJABLOON_STANDAARD As String = "Kwartaal_Template_cashflow_v10.pptx"
Private Const MAX_RIJEN As Long = 60
Private Const TITEL As String = "Naar PowerPoint"

' typeblokken (zelfde getallen als forecast_builder/layout.py)
Private Const MAX_TYPES As Long = 10
Private Const TYPE_COL1 As Long = 3           ' kolom C: eerste blok op tabblad Woningtypes
Private Const TYPE_BREEDTE As Long = 3        ' kolommen per type op Woningtypes (naam/%, %, kw)
Private Const TYPE_RIJ_NAAM As Long = 7
Private Const TYPE_RIJ_STARTKW As Long = 11
Private Const TYPE_RIJ_GROND As Long = 15
Private Const TYPE_RIJ_T1 As Long = 16
Private Const TYPE_RIJ_TN As Long = 25
Private Const INVOER_COL1 As Long = 12        ' kolom L: eerste paar (verkocht, transport) op tabblad Invoer
Private Const INVOER_BREEDTE As Long = 2
Private Const INVOER_RIJ1 As Long = 8
Private Const INVOER_RIJN As Long = 500        ' zelfde grens als de formules (layout.IN_ROWMAX)

Private Function Grafieken() As Variant
    ' vormnaam op de dia, kopcel van het blok op tabblad PowerPoint, aantal kolommen
    Grafieken = Array(Array("CF_GRAFIEK", "G7", 7), Array("SC_GRAFIEK", "O7", 10), Array("VT_GRAFIEK", "Z7", 5), Array("VO_GRAFIEK", "AF7", 5))
End Function

Private Function Tabellen() As Variant
    ' vormnaam op de dia, kopcel van de tabel, aantal rijen met kop, aantal kolommen, lege regels overslaan (1 = ja)
    Tabellen = Array(Array("SC_TABEL", "AL7", 9, 4, 0), Array("VT_TABEL", "AL18", MAX_TYPES + 1, 8, 1))
End Function

' -------------------------------------------------------------------------------------
'  1. Knoppen
' -------------------------------------------------------------------------------------
Public Sub KnoppenControleren()
    ' maakt alleen knoppen die nog niet bestaan, zodat verplaatste knoppen blijven staan
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    If Not BestaatVorm(ws, "btnNaarPowerPoint") Then
        MaakKnop ws, ws.Range("W32:X33"), "btnNaarPowerPoint", "Naar PowerPoint", "NaarPowerPoint", RGB(23, 54, 93), 0, 0, 2
    End If
    Set ws = ThisWorkbook.Worksheets(SH_PP)
    If Not BestaatVorm(ws, "btnNaarPowerPoint2") Then
        MaakKnop ws, ws.Range("C4:C5"), "btnNaarPowerPoint2", "Naar PowerPoint", "NaarPowerPoint", RGB(23, 54, 93), 150, 0, 2
    End If
    If Not BestaatVorm(ws, "btnControle") Then
        MaakKnop ws, ws.Range("D4:D5"), "btnControle", "Sjabloon controleren", "ControleerSjabloon", RGB(91, 97, 105), 150, 0, 2
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(SH_TYPES)
    If Not ws Is Nothing Then
        ' los van de cellen (plaatsing 3): kolommen invoegen of verwijderen schuift de knoppen niet weg
        If Not BestaatVorm(ws, "btnTypeInvoegen") Then
            MaakKnop ws, ws.Range("B4"), "btnTypeInvoegen", "Type invoegen", "TypeInvoegen", RGB(23, 54, 93), 92, 0, 3
        End If
        If Not BestaatVorm(ws, "btnTypeVerwijderen") Then
            MaakKnop ws, ws.Range("B4"), "btnTypeVerwijderen", "Type verwijderen", "TypeVerwijderen", RGB(91, 97, 105), 96, 96, 3
        End If
    End If
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
Public Sub TypeInvoegen()
    Dim wsT As Worksheet, wsI As Worksheet, k As Variant, pos As Long
    On Error GoTo Fout
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Set wsI = ThisWorkbook.Worksheets(SH_INVOER)
    If Not BlokLeeg(wsT, wsI, MAX_TYPES) Then
        MsgBox "Alle " & MAX_TYPES & " typeblokken zijn in gebruik. Maak eerst het laatste blok (type " & MAX_TYPES & ") leeg, ook op tabblad Invoer.", vbExclamation, "Type invoegen"
        Exit Sub
    End If
    k = Application.InputBox("Op welke positie komt het nieuwe type? (1 t/m " & MAX_TYPES & ")", "Type invoegen", 1, , , , , 1)
    If VarType(k) = vbBoolean Then Exit Sub
    pos = CLng(k)
    If pos < 1 Or pos > MAX_TYPES Or pos <> k Then
        MsgBox "Kies een heel getal van 1 t/m " & MAX_TYPES & ".", vbExclamation, "Type invoegen"
        Exit Sub
    End If
    If MsgBox("Op positie " & pos & " komt een leeg type; de types erachter schuiven op (ook op tabblad Invoer). " & _
              "Dit kan niet ongedaan worden gemaakt. Doorgaan?", vbQuestion + vbYesNo, "Type invoegen") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    BlokInvoegen wsT, TYPE_COL1 + (pos - 1) * TYPE_BREEDTE, TYPE_BREEDTE
    BlokInvoegen wsI, INVOER_COL1 + (pos - 1) * INVOER_BREEDTE, INVOER_BREEDTE
    ' het (lege) blok dat voorbij de capaciteit is geschoven weghalen
    BlokVerwijderen wsT, TYPE_COL1 + MAX_TYPES * TYPE_BREEDTE, TYPE_BREEDTE
    BlokVerwijderen wsI, INVOER_COL1 + MAX_TYPES * INVOER_BREEDTE, INVOER_BREEDTE
    Application.ScreenUpdating = True
    wsT.Activate
    wsT.Cells(TYPE_RIJ_NAAM, TYPE_COL1 + (pos - 1) * TYPE_BREEDTE).Select
    MsgBox "Type " & pos & " is ingevoegd. Vul het blok hier en de twee kolommen op tabblad Invoer.", vbInformation, "Type invoegen"
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Type invoegen"
End Sub

Public Sub TypeVerwijderen()
    Dim wsT As Worksheet, wsI As Worksheet, k As Variant, pos As Long, naam As String
    On Error GoTo Fout
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Set wsI = ThisWorkbook.Worksheets(SH_INVOER)
    k = Application.InputBox("Welk type wil je verwijderen? (positie 1 t/m " & MAX_TYPES & ")", "Type verwijderen", 1, , , , , 1)
    If VarType(k) = vbBoolean Then Exit Sub
    pos = CLng(k)
    If pos < 1 Or pos > MAX_TYPES Or pos <> k Then
        MsgBox "Kies een heel getal van 1 t/m " & MAX_TYPES & ".", vbExclamation, "Type verwijderen"
        Exit Sub
    End If
    naam = CStr(wsT.Cells(TYPE_RIJ_NAAM, TYPE_COL1 + (pos - 1) * TYPE_BREEDTE).Value)
    If MsgBox("Type " & pos & IIf(Len(naam) > 0, " (" & naam & ")", "") & " verwijderen, inclusief verkoop en transport op tabblad Invoer? " & _
              "Dit kan niet ongedaan worden gemaakt.", vbQuestion + vbYesNo, "Type verwijderen") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    ' eerst achteraan een leeg blok bijmaken, dan het gekozen blok weghalen: het aantal blokken blijft gelijk
    BlokKopieAanEind wsT, TYPE_COL1, TYPE_BREEDTE
    BlokKopieAanEind wsI, INVOER_COL1, INVOER_BREEDTE
    BlokVerwijderen wsT, TYPE_COL1 + (pos - 1) * TYPE_BREEDTE, TYPE_BREEDTE
    BlokVerwijderen wsI, INVOER_COL1 + (pos - 1) * INVOER_BREEDTE, INVOER_BREEDTE
    Application.ScreenUpdating = True
    MsgBox "Type " & pos & " is verwijderd; de types erachter zijn opgeschoven.", vbInformation, "Type verwijderen"
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Type verwijderen"
End Sub

Private Function BlokLeeg(wsT As Worksheet, wsI As Worksheet, pos As Long) As Boolean
    Dim c As Long
    c = TYPE_COL1 + (pos - 1) * TYPE_BREEDTE
    If Len(CStr(wsT.Cells(TYPE_RIJ_NAAM, c).Value)) > 0 Then Exit Function
    If Application.WorksheetFunction.CountA(wsT.Range(wsT.Cells(TYPE_RIJ_NAAM + 1, c), wsT.Cells(TYPE_RIJ_STARTKW, c))) > 0 Then Exit Function
    If Application.WorksheetFunction.Count(wsT.Range(wsT.Cells(TYPE_RIJ_GROND, c + 1), wsT.Cells(TYPE_RIJ_TN, c + 2))) > 0 Then Exit Function
    c = INVOER_COL1 + (pos - 1) * INVOER_BREEDTE
    If Application.WorksheetFunction.Count(wsI.Range(wsI.Cells(INVOER_RIJ1, c), wsI.Cells(INVOER_RIJN, c + INVOER_BREEDTE - 1))) > 0 Then Exit Function
    BlokLeeg = True
End Function

Private Sub BlokInvoegen(ws As Worksheet, col As Long, breedte As Long)
    ' voegt hele kolommen in en neemt opmaak, koppen en formules over van het blok dat erdoor naar rechts schoof
    ws.Range(ws.Columns(col), ws.Columns(col + breedte - 1)).Insert Shift:=xlToRight
    ws.Range(ws.Columns(col + breedte), ws.Columns(col + 2 * breedte - 1)).Copy ws.Columns(col)
    Application.CutCopyMode = False
    InvoerLeegmaken ws, col, breedte
End Sub

Private Sub BlokVerwijderen(ws As Worksheet, col As Long, breedte As Long)
    ws.Range(ws.Columns(col), ws.Columns(col + breedte - 1)).Delete Shift:=xlToLeft
End Sub

Private Sub BlokKopieAanEind(ws As Worksheet, col1 As Long, breedte As Long)
    ' kopie van het laatste blok op de plek erachter (opmaak, koppen en formules), zonder invoer
    Dim laatste As Long
    laatste = col1 + (MAX_TYPES - 1) * breedte
    ws.Range(ws.Columns(laatste), ws.Columns(laatste + breedte - 1)).Copy ws.Columns(laatste + breedte)
    Application.CutCopyMode = False
    InvoerLeegmaken ws, laatste + breedte, breedte
End Sub

Private Sub InvoerLeegmaken(ws As Worksheet, col As Long, breedte As Long)
    ' maakt de invoercellen van één blok leeg; koppen, formules en termijnnamen blijven staan
    If ws.Name = SH_TYPES Then
        ws.Range(ws.Cells(TYPE_RIJ_NAAM, col), ws.Cells(TYPE_RIJ_STARTKW, col)).ClearContents
        ws.Cells(TYPE_RIJ_GROND, col + 1).ClearContents
        ws.Range(ws.Cells(TYPE_RIJ_T1, col + 1), ws.Cells(TYPE_RIJ_TN, col + 2)).ClearContents
    Else
        ws.Range(ws.Cells(INVOER_RIJ1, col), ws.Cells(INVOER_RIJN, col + breedte - 1)).ClearContents
    End If
End Sub

' -------------------------------------------------------------------------------------
'  3. Naar PowerPoint
' -------------------------------------------------------------------------------------
Public Sub NaarPowerPoint()
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim stap As String, n As Long, pad As String, mist As String, fouten As String
    Dim g As Variant, t As Variant, i As Long, bericht As String
    On Error GoTo Fout

    stap = "werkbladen zoeken"
    Set wsP = ThisWorkbook.Worksheets(SH_PP)
    Application.Calculate
    n = AantalPeriodes()
    If n < 1 Then
        MsgBox "Er staan nog geen periodes op tabblad Invoer.", vbExclamation, TITEL
        Exit Sub
    End If

    stap = "PowerPoint starten"
    Set pp = HaalPowerPoint()
    If pp Is Nothing Then
        MsgBox "PowerPoint kon niet worden gestart.", vbExclamation, TITEL
        Exit Sub
    End If

    stap = "sjabloon openen"
    Set pres = OpenSjabloon(pp, wsP)
    If pres Is Nothing Then Exit Sub

    stap = "teksten vullen"
    VulTeksten pres, wsP, mist

    stap = "tabellen vullen"
    t = Tabellen()
    For i = LBound(t) To UBound(t)
        VulTabel pres, CStr(t(i)(0)), wsP.Range(CStr(t(i)(1))), CLng(t(i)(2)), CLng(t(i)(3)), (CLng(t(i)(4)) = 1), mist, fouten
    Next i

    stap = "grafieken vullen"
    g = Grafieken()
    For i = LBound(g) To UBound(g)
        VulGrafiek pres, CStr(g(i)(0)), wsP.Range(CStr(g(i)(1))), n, CLng(g(i)(2)), mist, fouten
    Next i

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
    MsgBox bericht, IIf(Len(mist) + Len(fouten) > 0, vbExclamation, vbInformation), TITEL
    Exit Sub

Fout:
    Application.ScreenUpdating = True
    MsgBox "Het ging mis bij: " & stap & vbCrLf & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, TITEL
End Sub

Private Function AantalPeriodes() As Long
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_N).Value
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then AantalPeriodes = CLng(v)
    If AantalPeriodes > MAX_RIJEN Then AantalPeriodes = MAX_RIJEN
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
    On Error GoTo 0
    If pres Is Nothing Then
        MsgBox "Het sjabloon is hier niet gevonden:" & vbCrLf & vbCrLf & pad & vbCrLf & vbCrLf & _
               "Kies zo het sjabloon. Tip: zet het in dezelfde map als dit bestand," & vbCrLf & _
               "of vul het pad in op tabblad PowerPoint (cel " & CEL_SJABLOON & ").", vbInformation, TITEL
#If Mac Then
        keuze = Application.GetOpenFilename()
#Else
        keuze = Application.GetOpenFilename("PowerPoint (*.pptx;*.potx),*.pptx;*.potx", , "Kies het sjabloon")
#End If
        If VarType(keuze) = vbBoolean Then Exit Function
        On Error Resume Next
        Set pres = pp.Presentations.Open(CStr(keuze), 0, -1, -1)
        On Error GoTo 0
        If pres Is Nothing Then MsgBox "Dit bestand kon niet worden geopend.", vbExclamation, TITEL
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
Private Function ZoekVorm(pres As Object, naam As String) As Object
    Dim s As Object, shp As Object
    For Each s In pres.Slides
        For Each shp In s.Shapes
            If shp.Name = naam Then
                Set ZoekVorm = shp
                Exit Function
            End If
        Next shp
    Next s
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

Private Function RichtingKleur(richting As String) As Long
    Select Case richting
        Case "slechter", "onder norm"
            RichtingKleur = RGB(192, 57, 47)
        Case "beter", "op norm"
            RichtingKleur = RGB(18, 134, 78)
        Case Else
            RichtingKleur = RGB(122, 122, 122)
    End Select
End Function

Private Sub VulTeksten(pres As Object, wsP As Worksheet, ByRef mist As String)
    ' loopt de KPI-regels af tot de eerste lege regel; kolom D = tekst, E = richting (kleur), F = vormnaam
    Dim r As Long, vorm As String, richting As String, shp As Object, stip As Object, kleur As Long
    r = KPI_ROW1
    Do While Len(CelTekst(wsP.Range("C" & r))) > 0 And r < KPI_ROW1 + 300
        vorm = Trim$(CelTekst(wsP.Range("F" & r)))
        If Len(vorm) > 0 Then
            Set shp = ZoekVorm(pres, vorm)
            If shp Is Nothing Then
                mist = mist & vorm & vbCrLf
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

Private Sub VulTabel(pres As Object, vorm As String, kop As Range, nRijen As Long, nKol As Long, overslaan As Boolean, _
                     ByRef mist As String, ByRef fouten As String)
    Dim shp As Object, tbl As Object, r As Long, c As Long, rij As Long
    Set shp = ZoekVorm(pres, vorm)
    If shp Is Nothing Then
        mist = mist & vorm & vbCrLf
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

Private Sub SchikDia5(pres As Object, tbl As Object)
    ' na het vullen van VT_TABEL: bij meer dan 4 typeregels de tabel binnen de kaart en boven het dianummer houden
    '   a) rijen lager (tot 10 pt) en letters kleiner (tot 5,5 pt), celmarges boven en onder 1 pt
    '   b) past het dan nog niet: grafiekkaart en grafiek inkorten (grafiek minimaal 110 pt), tabelkaart en tabel omhoog
    '   c) tabelkaart eindigt 8 pt onder de tabel
    ' alles in punten; fouten worden stil overgeslagen (de dia blijft dan zoals het sjabloon hem had)
    Dim tabelVorm As Object, tabelKaart As Object, kaart As Object, grafiek As Object
    Dim nR As Long, r As Long, c As Long
    Dim onder As Double, beschikbaar As Double, rijH As Double, letter As Double
    Dim tekort As Double, extra As Double, krimp As Double, schuif As Double, tabelH As Double
    Const RIJ_MAX As Double = 14.4
    Const RIJ_MIN As Double = 10
    Const LETTER_MAX As Double = 7.5                           ' lettergrootte van de cellen in het sjabloon
    Const LETTER_MIN As Double = 5.5
    Const AFSTAND As Double = 8
    Const GRAFIEK_MIN As Double = 110

    On Error Resume Next
    nR = tbl.Rows.Count
    If nR <= 5 Then Exit Sub                                     ' kop + 4 types: het sjabloon past zo
    Set tabelVorm = ZoekVorm(pres, "VT_TABEL")
    Set tabelKaart = ZoekVorm(pres, "VT_TABELKAART")
    Set kaart = ZoekVorm(pres, "VT_KAART")
    Set grafiek = ZoekVorm(pres, "VT_GRAFIEK")
    If tabelVorm Is Nothing Then Exit Sub

    ' ondergrens: 30 pt boven de dia-rand, daar staat het dianummer
    onder = 0
    onder = pres.PageSetup.SlideHeight - 30
    If onder <= 0 Then onder = 375
    beschikbaar = onder - tabelVorm.Top

    ' b) past de tabel ook met de laagste rijen niet, dan ruimte halen bij de grafiekkaart:
    '    eerst de ruimte tussen de kaarten boven de 8 pt, de rest door grafiekkaart en grafiek in te korten
    '    (grafiek minimaal 110 pt); tabelkaart en tabel schuiven omhoog en houden 8 pt tot de grafiekkaart
    tekort = nR * RIJ_MIN - beschikbaar
    If tekort > 0 And Not kaart Is Nothing And Not grafiek Is Nothing Then
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
    rijH = beschikbaar / nR
    If rijH > RIJ_MAX Then rijH = RIJ_MAX
    If rijH < RIJ_MIN Then rijH = RIJ_MIN
    letter = LETTER_MAX * rijH / RIJ_MAX
    If letter < LETTER_MIN Then letter = LETTER_MIN
    ' eerst letters en marges, dan de hoogte: PowerPoint maakt een rij nooit lager dan zijn tekst
    For r = 1 To nR
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.TextFrame
                .MarginTop = 1
                .MarginBottom = 1
                .TextRange.Font.Size = letter
            End With
        Next c
    Next r
    For r = 1 To nR
        tbl.Rows(r).Height = rijH
    Next r

    ' c) tabelkaart om de tabel sluiten; de rijhoogten teruglezen, want een gezette hoogte is een minimum
    tabelH = 0
    For r = 1 To nR
        tabelH = tabelH + tbl.Rows(r).Height
    Next r
    If Not tabelKaart Is Nothing And tabelH > 0 Then
        tabelKaart.Height = tabelVorm.Top + tabelH + AFSTAND - tabelKaart.Top
    End If
    On Error GoTo 0
End Sub

Private Sub VulGrafiek(pres As Object, vorm As String, kop As Range, n As Long, nKol As Long, ByRef mist As String, ByRef fouten As String)
    ' zet het blok (kop + n periodes) in het gegevensblad van de grafiek en past het bereik van elke reeks aan
    Dim shp As Object, ch As Object, wbE As Object, wsE As Object
    Dim arr() As Variant, r As Long, c As Long, v As Variant, i As Long, kol As String
    Set shp = ZoekVorm(pres, vorm)
    If shp Is Nothing Then
        mist = mist & vorm & vbCrLf
        Exit Sub
    End If
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
    ch.ChartData.Activate
    DoEvents
    Set wbE = ch.ChartData.Workbook
    Set wsE = wbE.Worksheets(1)
    wsE.Range("A1").Resize(MAX_RIJEN + 1, nKol).ClearContents
    wsE.Range("A1").Resize(n + 1, nKol).Value = arr
    For i = 1 To ch.SeriesCollection.Count
        kol = KolomVanReeks(CStr(ch.SeriesCollection(i).Formula))
        If Len(kol) > 0 Then ZetBereik ch.SeriesCollection(i), wsE, kol, n
    Next i
    wbE.Close
    Set wbE = Nothing
    Exit Sub
Mislukt:
    fouten = fouten & vorm & ": " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wbE Is Nothing Then wbE.Close
    On Error GoTo 0
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
'  5. Sjabloon controleren: staan alle vormnamen erin?
' -------------------------------------------------------------------------------------
Public Sub ControleerSjabloon()
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim r As Long, vorm As String, mist As String, aantal As Long, i As Long, lijst As Variant
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
        aantal = aantal + 1
        If ZoekVorm(pres, CStr(lijst(i)(0))) Is Nothing Then mist = mist & CStr(lijst(i)(0)) & vbCrLf
    Next i
    On Error Resume Next
    pres.Saved = -1
    pres.Close
    On Error GoTo 0
    If Len(mist) = 0 Then
        MsgBox "Alles compleet: alle " & aantal & " vormnamen staan in het sjabloon.", vbInformation, "Sjabloon controleren"
    Else
        MsgBox "Deze vormen ontbreken in het sjabloon (die slaat de knop over):" & vbCrLf & vbCrLf & mist & vbCrLf & _
               "Hernoem de vormen in PowerPoint via Start > Selecteren > Selectiedeelvenster.", vbExclamation, "Sjabloon controleren"
    End If
    Exit Sub
Fout:
    MsgBox "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Sjabloon controleren"
End Sub
