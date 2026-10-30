Attribute VB_Name = "modRDFO"
Option Explicit

' RDFO-prototype - knoppen op het tabblad Start.
'   RDFO_BronnenInlezen      : ZPMW03 en Woningetalage inlezen via een bestandskiezer
'   RDFO_PrognoseVastleggen  : jaarprognose van de upload toevoegen aan Historie (nooit overschrijven)
'   RDFO_MaakPowerPoint      : presentatie in huisstijl maken (geel / donkerblauw)
'   RDFO_ToonRekenbladen     : verborgen bron- en rekenbladen tonen of verbergen
' Alle berekeningen staan in gewone Excel-formules; de macro's doen alleen kopieren, plakken en opmaken.

Private Const HISTORIE_LAATSTE_RIJ As Long = 5005
Private Const BRON_CAPACITEIT As Long = 500
Private Const IMPORTLOG_KOLOM As Long = 24        ' kolom X op Historie
Private Const VERBORGEN As String = "Concept,VERW_Project,CALC_Versies,STG_ZPMW03,STG_Woningetalage,Bron_ZPMW03,Bron_Woningetalage,Lijsten,PPT,Dias"

' ---------------------------------------------------------------- huisstijl
Private Function NAVY() As Long
    NAVY = RGB(20, 40, 75)
End Function
Private Function GEEL() As Long
    GEEL = RGB(245, 184, 0)
End Function
Private Function GRIJS() As Long
    GRIJS = RGB(107, 115, 133)
End Function
Private Function LICHT() As Long
    LICHT = RGB(243, 245, 249)
End Function
Private Function LIJN() As Long
    LIJN = RGB(217, 222, 232)
End Function
Private Function WIT() As Long
    WIT = RGB(255, 255, 255)
End Function

Private Function Nm(ByVal naam As String) As Range
    Set Nm = ThisWorkbook.Names(naam).RefersToRange
End Function

Private Function T(ByVal naam As String) As String
    T = CStr(Nm(naam).Text)
End Function

' ================================================================= 1. BRONNEN INLEZEN
Public Sub RDFO_BronnenInlezen()
    Dim fZ As Variant, fE As Variant, peilTxt As String, peil As Date, versie As String
    Dim nZ As Long, nE As Long, impId As String, flt As String

    flt = "Excel-bestanden (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls"
    fZ = Application.GetOpenFilename(flt, , "Stap 1 van 2: kies de ZPMW03-export (Annuleren = overslaan)")
    fE = Application.GetOpenFilename(flt, , "Stap 2 van 2: kies de Woningetalage-export (Annuleren = overslaan)")
    If VarType(fZ) = vbBoolean And VarType(fE) = vbBoolean Then Exit Sub

    peilTxt = InputBox("Peildatum van deze export (dd-mm-jjjj):", "Peildatum", Format(StandaardPeildatum(), "dd-mm-yyyy"))
    If peilTxt = "" Then Exit Sub
    If Not IsDate(peilTxt) Then
        MsgBox "'" & peilTxt & "' is geen geldige datum. Er is niets ingelezen.", vbExclamation
        Exit Sub
    End If
    peil = CDate(peilTxt)
    versie = InputBox("Prognoseversie van deze export:" & vbCr & "Budget, Q1, Q2, Q3 of Q4", "Versie", VersieBijDatum(peil))
    If versie = "" Then Exit Sub
    versie = Trim(versie)
    If LCase(versie) = "budget" Then versie = "Budget" Else versie = UCase(versie)
    If InStr(1, ",Budget,Q1,Q2,Q3,Q4,", "," & versie & ",", vbBinaryCompare) = 0 Then
        MsgBox "Versie moet Budget, Q1, Q2, Q3 of Q4 zijn. Er is niets ingelezen.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    impId = "IMP-" & Format(Now, "yyyymmdd-hhnn")
    nZ = -2: nE = -2
    If VarType(fZ) <> vbBoolean Then nZ = LeesBron(CStr(fZ), "Bron_ZPMW03", "WBS-element")
    If VarType(fE) <> vbBoolean Then nE = LeesBron(CStr(fE), "Bron_Woningetalage", "Stufe 3")

    Nm("P_Peildatum").Value = peil
    Nm("P_RJ").Value = Year(peil)
    Nm("P_UploadVersie").Value = versie
    If versie = "Budget" Then Nm("P_UploadVJ").Value = Year(peil) + 1 Else Nm("P_UploadVJ").Value = Year(peil)
    Nm("P_VJ").Value = Nm("P_UploadVJ").Value
    Nm("P_ImportID").Value = impId
    Nm("P_Importdatum").Value = Date
    Nm("Sel_Versie").Value = versie
    Nm("Sel_Doeljaar").Value = Nm("P_UploadVJ").Value
    If nZ >= 0 Then
        Nm("P_BronZ").Value = BestandsNaam(CStr(fZ))
        Nm("P_RegelsZ").Value = nZ
        LogRegel impId, "ZPMW03", BestandsNaam(CStr(fZ)), peil, Nm("P_UploadVJ").Value, versie, nZ, "Ingelezen", ""
    End If
    If nE >= 0 Then
        Nm("P_BronE").Value = BestandsNaam(CStr(fE))
        Nm("P_RegelsE").Value = nE
        LogRegel impId, "Woningetalage", BestandsNaam(CStr(fE)), peil, Nm("P_UploadVJ").Value, versie, nE, "Ingelezen", ""
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    RDFO_OutlookJaren
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Controle").Activate
    ThisWorkbook.Worksheets("Controle").Range("A1").Select
    MsgBox "Ingelezen: " & IIf(nZ >= 0, nZ & " regels ZPMW03", "ZPMW03 overgeslagen") & ", " & _
           IIf(nE >= 0, nE & " regels Woningetalage", "Woningetalage overgeslagen") & "." & vbCr & vbCr & _
           "Controle: " & ThisWorkbook.Worksheets("Controle").Range("D4").Value & " fout(en), " & _
           ThisWorkbook.Worksheets("Controle").Range("D5").Value & " waarschuwing(en)." & vbCr & _
           "Los eventuele fouten op en klik daarna op 'Prognose vastleggen'.", vbInformation, "Bronnen inlezen"
    Exit Sub
Fout:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Inlezen is mislukt: " & Err.Description, vbExclamation
End Sub

Private Function LeesBron(ByVal pad As String, ByVal doelblad As String, ByVal sleutel As String) As Long
    Dim wbS As Workbook, ws As Worksheet, src As Worksheet, f As Range, gevonden As Range
    Dim laatsteRij As Long, laatsteKol As Long, n As Long

    Set wbS = Workbooks.Open(Filename:=pad, ReadOnly:=True, UpdateLinks:=0)
    For Each ws In wbS.Worksheets
        Set f = ws.Rows(1).Find(What:=sleutel, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not f Is Nothing Then
            Set src = ws
            Set gevonden = f
            Exit For
        End If
    Next ws
    If src Is Nothing Then
        wbS.Close SaveChanges:=False
        MsgBox "De kolom '" & sleutel & "' staat niet in rij 1 van " & BestandsNaam(pad) & "." & vbCr & _
               "Is dit de juiste export? Dit bestand is overgeslagen.", vbExclamation
        LeesBron = -1
        Exit Function
    End If
    laatsteRij = src.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    laatsteKol = src.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    If laatsteRij - 1 > BRON_CAPACITEIT Then
        MsgBox "Let op: de export heeft " & laatsteRij - 1 & " regels; het bestand rekent met maximaal " & _
               BRON_CAPACITEIT & ". Controle meldt dit als fout.", vbExclamation
    End If
    With ThisWorkbook.Worksheets(doelblad)
        .Cells.Clear
        src.Range(src.Cells(1, 1), src.Cells(laatsteRij, laatsteKol)).Copy
        .Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
    End With
    n = Application.WorksheetFunction.CountA(src.Range(src.Cells(2, gevonden.Column), src.Cells(laatsteRij, gevonden.Column)))
    wbS.Close SaveChanges:=False
    LeesBron = n
End Function

Private Function StandaardPeildatum() As Date
    Dim d As Date, q As Integer
    d = Date
    q = (Month(d) - 1) \ 3
    StandaardPeildatum = DateSerial(Year(d), q * 3 + 1, 1) - 1
End Function

Private Function VersieBijDatum(ByVal d As Date) As String
    Select Case Month(d)
        Case 1 To 3: VersieBijDatum = "Q1"
        Case 4 To 6: VersieBijDatum = "Q2"
        Case 7 To 9: VersieBijDatum = "Q3"
        Case Else: VersieBijDatum = "Q4"
    End Select
End Function

Private Function BestandsNaam(ByVal pad As String) As String
    BestandsNaam = Mid(pad, InStrRev(pad, Application.PathSeparator) + 1)
End Function

Private Sub LogRegel(ByVal id As String, ByVal bron As String, ByVal bestand As String, ByVal peil As Variant, _
                     ByVal vj As Variant, ByVal versie As String, ByVal n As Long, ByVal status As String, ByVal toel As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Historie")
    r = 6
    Do While ws.Cells(r, IMPORTLOG_KOLOM).Value <> ""
        r = r + 1
    Loop
    ws.Cells(r, IMPORTLOG_KOLOM).Value = id
    ws.Cells(r, IMPORTLOG_KOLOM + 1).Value = bron
    ws.Cells(r, IMPORTLOG_KOLOM + 2).Value = bestand
    ws.Cells(r, IMPORTLOG_KOLOM + 3).Value = Now
    ws.Cells(r, IMPORTLOG_KOLOM + 3).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Cells(r, IMPORTLOG_KOLOM + 4).Value = peil
    ws.Cells(r, IMPORTLOG_KOLOM + 4).NumberFormat = "dd-mm-yyyy"
    ws.Cells(r, IMPORTLOG_KOLOM + 5).Value = vj
    ws.Cells(r, IMPORTLOG_KOLOM + 6).Value = versie
    ws.Cells(r, IMPORTLOG_KOLOM + 7).Value = n
    ws.Cells(r, IMPORTLOG_KOLOM + 9).Value = status
    ws.Cells(r, IMPORTLOG_KOLOM + 10).Value = toel
End Sub

' ================================================================= 2. PROGNOSE VASTLEGGEN
Public Sub RDFO_PrognoseVastleggen()
    Dim wsC As Worksheet, wsH As Worksheet, n As Long, r0 As Long, fouten As Long, bestaand As Long
    Dim versie As String, toel As String, antw As VbMsgBoxResult

    Application.CalculateFull
    Set wsC = ThisWorkbook.Worksheets("Concept")
    Set wsH = ThisWorkbook.Worksheets("Historie")
    fouten = ThisWorkbook.Worksheets("Controle").Range("D4").Value
    If fouten > 0 Then
        MsgBox "Controle bevat " & fouten & " fout(en). Los die eerst op; daarna kun je vastleggen.", vbExclamation, "Prognose vastleggen"
        ThisWorkbook.Worksheets("Controle").Activate
        Exit Sub
    End If
    versie = Nm("P_UploadVersie").Value & " " & Nm("P_UploadVJ").Value
    n = wsC.Range("E3").Value
    bestaand = wsC.Range("N3").Value
    If n = 0 Then
        MsgBox "Er zijn geen regels om vast te leggen (zijn de bronnen ingelezen?).", vbExclamation
        Exit Sub
    End If

    If bestaand > 0 Then
        antw = MsgBox("Versie " & versie & " is al vastgelegd." & vbCr & vbCr & _
                      "Wil je een CORRECTIE vastleggen? Er komt een nieuwe revisie bij;" & vbCr & _
                      "de oude regels blijven bewaard en tellen niet meer mee.", vbYesNo + vbQuestion, "Correctie")
        If antw <> vbYes Then Exit Sub
        toel = InputBox("Reden van de correctie (verplicht):", "Correctie " & versie)
        If Trim(toel) = "" Then
            MsgBox "Zonder toelichting wordt niets vastgelegd.", vbExclamation
            Exit Sub
        End If
        Nm("P_ToelCorrectie").Value = toel
        Application.CalculateFull
    Else
        antw = MsgBox("Prognose " & versie & " vastleggen?" & vbCr & vbCr & n & " regels worden toegevoegd aan Historie " & _
                      "(status: " & Nm("P_StatusVastleg").Value & ")." & vbCr & "Vastgelegde waarden veranderen daarna niet meer.", _
                      vbOKCancel + vbQuestion, "Prognose vastleggen")
        If antw <> vbOK Then Exit Sub
    End If

    r0 = wsH.Range("D3").Value
    If r0 + n - 1 > HISTORIE_LAATSTE_RIJ Then
        MsgBox "Historie is vol (maximaal rij " & HISTORIE_LAATSTE_RIJ & "). Vraag om de formules in Historie verder door te trekken.", vbCritical
        Exit Sub
    End If
    Application.ScreenUpdating = False
    wsH.Cells(r0, 1).Resize(n, 18).Value = wsC.Range("A8").Resize(n, 18).Value
    wsH.Cells(r0, 3).Resize(n, 1).NumberFormat = "dd-mm-yyyy"
    wsH.Cells(r0, 8).Resize(n, 1).NumberFormat = "dd-mm-yyyy"
    wsH.Cells(r0, 1).Resize(n, 18).Font.Size = 8
    LogRegel Nm("P_ImportID").Value, "Vastlegging", ThisWorkbook.Name, Nm("P_Peildatum").Value, Nm("P_UploadVJ").Value, _
             CStr(Nm("P_UploadVersie").Value), n, CStr(Nm("P_StatusVastleg").Value), toel
    Nm("P_ToelCorrectie").Value = ""
    Application.CalculateFull
    Application.ScreenUpdating = True
    wsH.Activate
    wsH.Cells(r0, 1).Select
    MsgBox "Klaar: " & n & " regels vastgelegd voor " & versie & "." & vbCr & "Sla het bestand op.", vbInformation, "Prognose vastleggen"
End Sub

' ================================================================= 3. REKENBLADEN
Public Sub RDFO_ToonRekenbladen()
    Dim naam As Variant, tonen As Boolean
    tonen = (ThisWorkbook.Worksheets("Concept").Visible <> xlSheetVisible)
    For Each naam In Split(VERBORGEN, ",")
        If tonen Then
            ThisWorkbook.Worksheets(CStr(naam)).Visible = xlSheetVisible
        Else
            ThisWorkbook.Worksheets(CStr(naam)).Visible = xlSheetHidden
        End If
    Next naam
    ThisWorkbook.Worksheets("Start").Activate
    MsgBox IIf(tonen, "De bron- en rekenbladen staan nu achteraan zichtbaar. Wijzig daar niets.", "De bron- en rekenbladen zijn weer verborgen."), vbInformation
End Sub

' ================================================================= 4. POWERPOINT
' De opbouw staat in de dia-specificatie (verborgen tabblad PPT, bereik PPT_Spec): per dia de titel, kernboodschap en
' elementen. Tabellen worden echte (bewerkbare) PowerPoint-tabellen; grafieken worden als afbeelding geplakt.
Public Sub RDFO_MaakPowerPoint()
    Dim pp As Object, pres As Object, sld As Object, spec As Range, r As Long, dia As Long, vorige As Long
    Dim soort As String, wsD As Worksheet, terug As Worksheet, zicht As XlSheetVisibility, pad As String, melding As String

    On Error GoTo Fout
    Application.CalculateFull
    RDFO_OutlookJaren
    Set terug = ActiveSheet
    Set wsD = ThisWorkbook.Worksheets("Dias")
    zicht = wsD.Visible
    wsD.Visible = xlSheetVisible

    Set pp = CreateObject("PowerPoint.Application")
    pp.Visible = True
    Set pres = pp.Presentations.Add(True)
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    TitelDia pres

    Set spec = Nm("PPT_Spec")
    vorige = -1
    For r = 1 To spec.Rows.Count
        If spec.Cells(r, 11).Value = 1 Then
            dia = spec.Cells(r, 1).Value
            If dia <> vorige Then
                Set sld = NieuweDia(pres, spec.Cells(r, 2).Text, spec.Cells(r, 3).Text)
                vorige = dia
            End If
            soort = spec.Cells(r, 4).Text
            Select Case soort
                Case "kpi": Tegels sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value
                Case "tabel": Tabel sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 10).Text
                Case "grafiek": Grafiek sld, wsD, CLng(spec.Cells(r, 5).Value), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value, spec.Cells(r, 10).Text
                Case "opsomming": Opsomming sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value
            End Select
        End If
    Next r

    wsD.Visible = zicht
    terug.Activate
    pad = OpslagMap() & Application.PathSeparator & "RDFO " & T("P_Regio") & " " & T("Sel_Versie") & " " & T("P_VJ") & " " & Format(Now, "yyyy-mm-dd hhnn") & ".pptx"
    pres.SaveAs pad
    MsgBox "PowerPoint gemaakt en opgeslagen:" & vbCr & pad, vbInformation, "Maak PowerPoint"
    Exit Sub
Fout:
    melding = Err.Description
    On Error Resume Next
    wsD.Visible = zicht
    terug.Activate
    MsgBox "PowerPoint maken lukte niet: " & melding & vbCr & "(Is PowerPoint geinstalleerd? Sluit open dialoogvensters in PowerPoint.)", vbExclamation
End Sub

' Laat in Outlook en in de outlook-grafieken alleen de jaren zien waarvoor cijfers zijn (OUT_N = laatste jaar met cijfers).
Public Sub RDFO_OutlookJaren()
    Dim wsO As Worksheet, n As Long, i As Long, co As ChartObject, s As Long, basis As Long, rij As Long
    Dim doel As Variant, delen() As String
    Set wsO = ThisWorkbook.Worksheets("Outlook")
    n = Nm("OUT_N").Value
    For i = 1 To 8
        wsO.Columns(2 + i).Hidden = (i > n)
    Next i
    wsO.Columns(11).Hidden = (Nm("OUT_Later").Value <> 1)
    ' grafiek: blad|index|basisrij (rijen basis..basis+3 = realisatie, prognose, totaal, budget op Outlook vanaf kolom P)
    For Each doel In Split(Nm("OUT_Grafieken").Value, ";")
        delen = Split(doel, "|")
        Set co = ThisWorkbook.Worksheets(delen(0)).ChartObjects(CLng(delen(1)))
        basis = CLng(delen(2))
        For s = 1 To co.Chart.SeriesCollection.Count
            rij = basis + s - 1
            co.Chart.SeriesCollection(s).Values = wsO.Range(wsO.Cells(rij, 16), wsO.Cells(rij, 15 + n))
            co.Chart.SeriesCollection(s).XValues = wsO.Range(wsO.Cells(5, 16), wsO.Cells(5, 15 + n))
        Next s
    Next doel
End Sub

Private Function OpslagMap() As String
    OpslagMap = ThisWorkbook.Path
    If OpslagMap = "" Or LCase(Left(OpslagMap, 4)) = "http" Then OpslagMap = Environ("TEMP")
End Function

Private Sub Tekst(ByVal sh As Object, ByVal txt As String, ByVal grootte As Single, ByVal kleur As Long, ByVal vet As Boolean)
    With sh.TextFrame
        .WordWrap = True
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = txt
        With .TextRange.Font
            .Name = "Arial"
            .Size = grootte
            .Color.RGB = kleur
            .Bold = vet
        End With
    End With
End Sub

Private Function Vlak(ByVal sld As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal kleur As Long) As Object
    Dim sh As Object
    Set sh = sld.Shapes.AddShape(1, x, y, w, h)
    sh.Fill.ForeColor.RGB = kleur
    sh.Line.Visible = False
    sh.Shadow.Visible = False
    Set Vlak = sh
End Function

Private Sub TitelDia(ByVal pres As Object)
    Dim sld As Object, sh As Object
    Set sld = pres.Slides.Add(1, 12)
    Vlak sld, 0, 0, 300, 540, GEEL
    Vlak sld, 300, 0, 8, 540, NAVY
    Set sh = sld.Shapes.AddTextbox(1, 40, 60, 230, 30)
    Tekst sh, T("P_Regio"), 14, NAVY, True
    Set sh = sld.Shapes.AddTextbox(1, 40, 470, 230, 40)
    Tekst sh, "Gemaakt op " & Format(Now, "dd-mm-yyyy hh:nn"), 9, NAVY, False
    Set sh = sld.Shapes.AddTextbox(1, 360, 190, 560, 110)
    Tekst sh, T("PPT_Titel"), 38, NAVY, True
    Vlak sld, 362, 304, 90, 6, GEEL
    Set sh = sld.Shapes.AddTextbox(1, 360, 324, 560, 30)
    Tekst sh, T("PPT_Sub"), 16, NAVY, False
    Set sh = sld.Shapes.AddTextbox(1, 360, 356, 560, 20)
    Tekst sh, "Bron: " & ThisWorkbook.Name, 10, GRIJS, False
End Sub

Private Function NieuweDia(ByVal pres As Object, ByVal titel As String, ByVal kern As String) As Object
    Dim sld As Object, sh As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, 12)
    Vlak sld, 0, 0, 960, 6, NAVY
    Vlak sld, 36, 30, 6, 30, GEEL
    Set sh = sld.Shapes.AddTextbox(1, 52, 28, 860, 34)
    Tekst sh, titel, 24, NAVY, True
    If kern <> "" Then
        Set sh = sld.Shapes.AddTextbox(1, 52, 64, 860, 22)
        Tekst sh, kern, 12, GRIJS, False
    End If
    Set sh = sld.Shapes.AddLine(36, 510, 924, 510)
    sh.Line.ForeColor.RGB = LIJN
    sh.Line.Weight = 0.75
    Set sh = sld.Shapes.AddTextbox(1, 36, 516, 700, 16)
    Tekst sh, T("PPT_Footer"), 9, GRIJS, False
    Set sh = sld.Shapes.AddTextbox(1, 880, 516, 44, 16)
    Tekst sh, CStr(sld.SlideIndex), 9, NAVY, True
    sh.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    Set NieuweDia = sld
End Function

Private Sub Kop(ByVal sld As Object, ByVal tekstKop As String, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Dim sh As Object
    If tekstKop = "" Then Exit Sub
    Set sh = sld.Shapes.AddTextbox(1, x, y - 20, w, 16)
    Tekst sh, tekstKop, 11, NAVY, True
End Sub

' Zes kerncijfertegels uit een blok van 6 regels x 3 kolommen (label, waarde, toelichting).
Private Sub Tegels(ByVal sld As Object, ByVal blok As Range, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim k As Long, tw As Single, tx As Single, sh As Object
    tw = (w - 5 * 8) / 6
    For k = 1 To 6
        tx = x + (k - 1) * (tw + 8)
        Vlak sld, tx, y, tw, h, LICHT
        Vlak sld, tx, y, tw, 3, GEEL
        Set sh = sld.Shapes.AddTextbox(1, tx + 6, y + 8, tw - 12, 14)
        Tekst sh, blok.Cells(k, 1).Text, 8.5, NAVY, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 6, y + 24, tw - 12, 28)
        Tekst sh, blok.Cells(k, 2).Text, 19, NAVY, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 6, y + h - 18, tw - 12, 14)
        Tekst sh, blok.Cells(k, 3).Text, 8, GRIJS, False
    Next k
End Sub

Private Function CelTekst(ByVal c As Range) As String
    Dim s As String
    s = c.Text
    If Len(s) > 0 And Left(s, 1) = "#" And IsNumeric(c.Value2) Then s = Format(c.Value2, "#,##0")
    CelTekst = Replace(Replace(s, vbCr, ""), vbLf, Chr(11))
End Function

' Bewerkbare PowerPoint-tabel uit een bereik (eerste regel = kop). De regel boven het bereik bevat kolomgewichten:
' gewicht 0 of leeg = kolom weglaten. Regels met een lege eerste cel worden overgeslagen.
' Kleuren (rood bij verlies, stoplicht, geselecteerde versie) komen uit de weergave in Excel (DisplayFormat).
Private Sub Tabel(ByVal sld As Object, ByVal bron As Range, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal tekstKop As String)
    Dim gew As Range, kol() As Long, nk As Long, k As Long, reg() As Long, nr As Long, r As Long, som As Double
    Dim shp As Object, tbl As Object, i As Long, j As Long, c As Range, txt As String, grootte As Single, bol As Boolean

    Set gew = bron.Rows(1).Offset(-1, 0)
    ReDim kol(1 To bron.Columns.Count)
    For k = 1 To bron.Columns.Count
        If IsNumeric(gew.Cells(1, k).Value) And gew.Cells(1, k).Value <> "" Then
            If gew.Cells(1, k).Value > 0 Then
                nk = nk + 1
                kol(nk) = k
                som = som + gew.Cells(1, k).Value
            End If
        End If
    Next k
    ReDim reg(1 To bron.Rows.Count)
    For r = 1 To bron.Rows.Count
        If r = 1 Or Len(bron.Cells(r, 1).Text) > 0 Then
            nr = nr + 1
            reg(nr) = r
        End If
    Next r
    If nk = 0 Or nr = 0 Then Exit Sub
    grootte = IIf(nr > 15, 8, 9)
    Kop sld, tekstKop, x, y, w

    Set shp = sld.Shapes.AddTable(nr, nk, x, y, w, nr * (grootte + 6))
    Set tbl = shp.Table
    On Error Resume Next
    tbl.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", False
    On Error GoTo 0
    tbl.FirstRow = True
    tbl.HorizBanding = False
    For j = 1 To nk
        tbl.Columns(j).Width = w * gew.Cells(1, kol(j)).Value / som
    Next j

    For i = 1 To nr
        For j = 1 To nk
            Set c = bron.Cells(reg(i), kol(j))
            txt = CelTekst(c)
            bol = (txt = ChrW(9679) Or txt = ChrW(9675))
            With tbl.Cell(i, j).Shape
                .TextFrame.MarginLeft = 3: .TextFrame.MarginRight = 3
                .TextFrame.MarginTop = 1: .TextFrame.MarginBottom = 1
                .TextFrame.VerticalAnchor = 3
                .TextFrame.WordWrap = True
                .TextFrame.TextRange.Text = txt
                With .TextFrame.TextRange.Font
                    .Name = "Arial"
                    .Size = IIf(bol, grootte + 5, grootte)
                    If i = 1 Then
                        .Color.RGB = WIT
                        .Bold = True
                    Else
                        .Color.RGB = c.DisplayFormat.Font.Color
                        .Bold = c.Font.Bold
                        .Italic = c.Font.Italic
                    End If
                End With
                If j = 1 Then
                    .TextFrame.TextRange.ParagraphFormat.Alignment = 1
                ElseIf bol Or i = 1 Or c.HorizontalAlignment = xlCenter Then
                    .TextFrame.TextRange.ParagraphFormat.Alignment = 2
                ElseIf c.HorizontalAlignment = xlLeft Then
                    .TextFrame.TextRange.ParagraphFormat.Alignment = 1
                Else
                    .TextFrame.TextRange.ParagraphFormat.Alignment = 3
                End If
                If i = 1 Then
                    .Fill.ForeColor.RGB = NAVY
                ElseIf c.DisplayFormat.Interior.ColorIndex <> xlColorIndexNone And c.DisplayFormat.Interior.Color <> WIT Then
                    .Fill.ForeColor.RGB = c.DisplayFormat.Interior.Color
                Else
                    .Fill.ForeColor.RGB = WIT
                End If
            End With
            With tbl.Cell(i, j).Borders(3)
                .Visible = True
                .ForeColor.RGB = LIJN
                .Weight = 0.5
            End With
        Next j
        tbl.Rows(i).Height = grootte + 6
    Next i
End Sub

Private Sub Grafiek(ByVal sld As Object, ByVal wsD As Worksheet, ByVal nr As Long, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal tekstKop As String)
    Dim co As ChartObject, sh As Object, poging As Integer, bestand As String
    Set co = wsD.ChartObjects(nr)
    Kop sld, tekstKop, x, y, w
    For poging = 1 To 5
        On Error Resume Next
        Set sh = Nothing
        co.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        Set sh = sld.Shapes.PasteSpecial(DataType:=2)(1)
        On Error GoTo 0
        If Not sh Is Nothing Then Exit For
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next poging
    If sh Is Nothing Then
        bestand = Environ("TEMP") & Application.PathSeparator & "rdfo_grafiek_" & nr & ".png"
        co.Chart.Export bestand
        Set sh = sld.Shapes.AddPicture(bestand, False, True, x, y)
    End If
    sh.LockAspectRatio = True
    If sh.Width / sh.Height > w / h Then sh.Width = w Else sh.Height = h
    sh.Left = x + (w - sh.Width) / 2
    sh.Top = y
End Sub

Private Sub Opsomming(ByVal sld As Object, ByVal bron As Range, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim sh As Object, s As String, c As Range
    For Each c In bron.Cells
        If Len(Trim(c.Text)) > 0 Then s = s & IIf(Len(s) > 0, vbCr, "") & c.Text
    Next c
    If s = "" Then s = "Geen bijzonderheden."
    Vlak sld, x, y, w, h, LICHT
    Set sh = sld.Shapes.AddTextbox(1, x + 24, y + 20, w - 48, h - 40)
    Tekst sh, s, 16, NAVY, False
    With sh.TextFrame.TextRange.ParagraphFormat
        .Bullet.Visible = True
        .Bullet.Character = 9632
        .Bullet.Font.Color.RGB = GEEL
        .SpaceAfter = 14
    End With
End Sub
