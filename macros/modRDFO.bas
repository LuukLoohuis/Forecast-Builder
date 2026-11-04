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
Private Const VERBORGEN As String = "Historie,SCEN,Concept,VERW_Project,CALC_Versies,STG_ZPMW03,STG_Woningetalage,Bron_ZPMW03,Bron_Woningetalage,Bron_AK,Lijsten,PPT,Dias"

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
    Dim nZ As Long, nE As Long, nA As Long, impId As String, flt As String, ontbreekt As String, nieuw As Long, fA As Variant

    flt = "Excel-bestanden (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls"
    fZ = Application.GetOpenFilename(flt, , "Stap 1 van 3: kies de ZPMW03-export (Annuleren = overslaan)")
    fE = Application.GetOpenFilename(flt, , "Stap 2 van 3: kies de Woningetalage-export (Annuleren = overslaan)")
    fA = Application.GetOpenFilename(flt, , "Stap 3 van 3: kies de AK-percentages uit SAP (Annuleren = overslaan)")
    If VarType(fZ) = vbBoolean And VarType(fE) = vbBoolean And VarType(fA) = vbBoolean Then Exit Sub

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
    nA = -2
    If VarType(fA) <> vbBoolean Then nA = LeesBron(CStr(fA), "Bron_AK", "WBS-element")

    Nm("CT_Bekeken").ClearContents   ' nieuwe bronnen = waarschuwingen opnieuw bekijken
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
    ontbreekt = ""
    If nZ >= 0 Then ontbreekt = ontbreekt & OntbrekendeKolommen("STG_ZPMW03")
    If nE >= 0 Then ontbreekt = ontbreekt & OntbrekendeKolommen("STG_Woningetalage")
    nieuw = 0
    If nZ >= 0 Then nieuw = NieuweProjecten()
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Controle").Activate
    ThisWorkbook.Worksheets("Controle").Range("A1").Select
    MsgBox "Ingelezen: " & IIf(nZ >= 0, nZ & " regels ZPMW03", "ZPMW03 overgeslagen") & ", " & _
           IIf(nE >= 0, nE & " regels Woningetalage", "Woningetalage overgeslagen") & ", " & _
           IIf(nA >= 0, nA & " regels AK-percentages", "AK-percentages ongewijzigd") & "." & vbCr & vbCr & _
           "Controle: " & ThisWorkbook.Worksheets("Controle").Range("D4").Value & " fout(en), " & _
           ThisWorkbook.Worksheets("Controle").Range("D5").Value & " waarschuwing(en)." & vbCr & _
           IIf(ontbreekt <> "", vbCr & "Deze kolommen zijn niet gevonden (kolomnaam in de export wijkt af):" & ontbreekt & vbCr, "") & _
           IIf(nieuw > 0, vbCr & nieuw & " WBS-element(en) staan nog niet in de projectlijst. Klik op 'Nieuwe projecten toevoegen'." & vbCr, "") & _
           vbCr & "Los eventuele fouten op en klik daarna op 'Prognose vastleggen'.", vbInformation, "Bronnen inlezen"
    Exit Sub
Fout:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Inlezen is mislukt: " & Err.Description, vbExclamation
End Sub

Private Function LeesBron(ByVal pad As String, ByVal doelblad As String, ByVal sleutel As String) As Long
    ' Zoekt de kopregel (eerste 30 rijen, alle tabbladen), neemt alles vanaf die kopregel over als waarden
    ' (zonder klembord) en zet de kopregel in rij 1. Spaties rond kolomkoppen worden verwijderd.
    Dim wbS As Workbook, ws As Worksheet, src As Worksheet, kopRij As Long, sleutelKol As Long
    Dim laatsteRij As Long, laatsteKol As Long, n As Long, r As Long, k As Long, data As Variant, f As Range

    Set wbS = Workbooks.Open(Filename:=pad, ReadOnly:=True, UpdateLinks:=0)
    For Each ws In wbS.Worksheets
        For r = 1 To 30
            For k = 1 To 200
                If LCase(Left(SchoneKop(ws.Cells(r, k).Value), Len(sleutel))) = LCase(sleutel) Then
                    Set src = ws: kopRij = r: sleutelKol = k
                    Exit For
                End If
            Next k
            If Not src Is Nothing Then Exit For
        Next r
        If Not src Is Nothing Then Exit For
    Next ws
    If src Is Nothing Then
        wbS.Close SaveChanges:=False
        MsgBox "De kolom '" & sleutel & "' is niet gevonden in de eerste 30 rijen van " & BestandsNaam(pad) & "." & vbCr & _
               "Is dit de juiste export? Dit bestand is overgeslagen.", vbExclamation
        LeesBron = -1
        Exit Function
    End If
    Set f = src.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious)
    laatsteRij = f.Row
    Set f = src.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious)
    laatsteKol = f.Column
    If laatsteRij - kopRij > BRON_CAPACITEIT Then
        MsgBox "Let op: de export heeft " & laatsteRij - kopRij & " regels; het bestand rekent met maximaal " & _
               BRON_CAPACITEIT & ". Controle meldt dit als fout.", vbExclamation
    End If
    data = src.Range(src.Cells(kopRij, 1), src.Cells(laatsteRij, laatsteKol)).Value
    wbS.Close SaveChanges:=False
    For k = 1 To UBound(data, 2)
        If Not IsEmpty(data(1, k)) Then data(1, k) = SchoneKop(data(1, k))
    Next k
    For r = 2 To UBound(data, 1)
        If Trim(CStr(data(r, sleutelKol))) <> "" Then n = n + 1
    Next r
    With ThisWorkbook.Worksheets(doelblad)
        .Cells.Clear
        .Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    End With
    LeesBron = n
End Function

Private Function SchoneKop(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    SchoneKop = Application.WorksheetFunction.Trim(Replace(Replace(CStr(v), Chr(160), " "), vbLf, " "))
End Function

Private Function OntbrekendeKolommen(ByVal stg As String) As String
    ' Kolommen die het bestand zoekt maar niet vindt (STG-blad rij 4 = ONTBREEKT; rij 2 = gezochte kop).
    Dim ws As Worksheet, k As Long, t As String
    Set ws = ThisWorkbook.Worksheets(stg)
    For k = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If CStr(ws.Cells(4, k).Text) = "ONTBREEKT" Then t = t & vbCr & "   - " & ws.Cells(2, k).Text
    Next k
    OntbrekendeKolommen = t
End Function

Private Function KolomNr(ByVal ws As Worksheet, ByVal kop As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=kop, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then KolomNr = f.Column
End Function

Private Function NieuweProjecten() As Long
    Dim wsZ As Worksheet, c As Long, r As Long, map As Range, w As String
    Set wsZ = ThisWorkbook.Worksheets("Bron_ZPMW03")
    Set map = Nm("Map_WBS")
    c = KolomNr(wsZ, "WBS-element")
    If c = 0 Then Exit Function
    For r = 2 To wsZ.Cells(wsZ.Rows.Count, c).End(xlUp).Row
        w = Trim(CStr(wsZ.Cells(r, c).Value))
        If w <> "" Then If Application.WorksheetFunction.CountIf(map, w) = 0 Then NieuweProjecten = NieuweProjecten + 1
    Next r
End Function

' ================================================================= 1b. NIEUWE PROJECTEN NAAR MAPPING
Public Sub RDFO_ProjectenToevoegen()
    Dim wsZ As Worksheet, wsE As Worksheet, map As Range, r As Long, i As Long, n As Long, vrij As Long
    Dim cW As Long, cP As Long, cO As Long, eS As Long, eA As Long, eH As Long, wbs As String, m As Variant, eerste As Range, naam As String

    Set wsZ = ThisWorkbook.Worksheets("Bron_ZPMW03")
    Set wsE = ThisWorkbook.Worksheets("Bron_Woningetalage")
    Set map = Nm("Map_WBS")
    cW = KolomNr(wsZ, "WBS-element"): cP = KolomNr(wsZ, "Projectdefinitie"): cO = KolomNr(wsZ, "Omschrijving")
    eS = KolomNr(wsE, "Stufe 3"): eA = KolomNr(wsE, "Aantal woningen"): eH = KolomNr(wsE, "Aandeel Heijm.")
    If cW = 0 Then
        MsgBox "Er is nog geen ZPMW03 ingelezen (kolom 'WBS-element' niet gevonden). Klik eerst op 'Bronnen inlezen'.", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    vrij = 1
    For r = 2 To wsZ.Cells(wsZ.Rows.Count, cW).End(xlUp).Row
        wbs = Trim(CStr(wsZ.Cells(r, cW).Value))
        If wbs <> "" And Application.WorksheetFunction.CountIf(map, wbs) = 0 Then
            Do While vrij <= map.Rows.Count
                If Trim(CStr(map.Cells(vrij, 1).Value)) = "" Then Exit Do
                vrij = vrij + 1
            Loop
            If vrij > map.Rows.Count Then
                MsgBox "De projectlijst op Invoer is vol (" & map.Rows.Count & " regels). " & n & " project(en) toegevoegd; de rest niet.", vbExclamation
                Exit For
            End If
            With map.Cells(vrij, 1)
                .Value = wbs
                If cP > 0 Then .Offset(0, 1).Value = wsZ.Cells(r, cP).Value
                If cO > 0 Then
                    naam = Trim(CStr(wsZ.Cells(r, cO).Value))
                    If Application.WorksheetFunction.CountIf(map.Offset(0, 2), naam) > 0 Then naam = naam & " (" & wbs & ")"
                    .Offset(0, 2).Value = naam
                End If
                .Offset(0, 5).Value = "Opstal"
                .Offset(0, 6).Value = "Nee"
                .Offset(0, 9).Value = "Nee"
                m = CVErr(xlErrNA)
                If eS > 0 Then m = Application.Match(wbs, wsE.Columns(eS), 0)
                If Not IsError(m) Then
                    If eA > 0 Then .Offset(0, 3).Value = wsE.Cells(m, eA).Value
                    If eH > 0 Then .Offset(0, 4).Value = wsE.Cells(m, eH).Value
                    .Offset(0, 6).Value = "Ja"
                End If
                .Resize(1, 19).Interior.Color = RGB(255, 235, 156)
            End With
            If eerste Is Nothing Then Set eerste = map.Cells(vrij, 1)
            n = n + 1
        End If
    Next r
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "Alle WBS-elementen uit ZPMW03 staan al in de projectlijst.", vbInformation, "Nieuwe projecten"
    Else
        ThisWorkbook.Worksheets("Invoer").Activate
        eerste.Select
        MsgBox n & " nieuw(e) project(en) toegevoegd aan Invoer › 2 (oranje gemarkeerd)." & vbCr & vbCr & _
               "Controleer per regel:" & vbCr & "   - Projectnaam" & vbCr & "   - Type: Opstal of Deelneming (standaard Opstal)" & vbCr & _
               "   - Woningen meetellen (Ja als het project in de Woningetalage staat)", vbInformation, "Nieuwe projecten"
    End If
End Sub

' ================================================================= 1d. OUDE PROGNOSES OVERZETTEN (eenmalig)
' Leest uit een oud RDFO-bestand de tabbladen Budget en Prognose Q1..Q4 (opzet: rij 2 'BOB 2026' / 'BOR 2026' / 'AK dekking 2026',
' rij 3 Q1..Q4 + Totaal, kopregel met 'WBS-element' in kolom A, bedragen x 1.000) en legt ze vast in Historie als vastgestelde versies.
Public Sub RDFO_OudePrognosesOverzetten()
    Dim f As Variant, wbS As Workbook, ws As Worksheet, tabs As Variant, tabblad As Variant, versie As String, verslag As String
    Dim kopRij As Long, r As Long, c As Long, jaar As Long, leeg As Long, n As Long, totN As Long
    Dim cBOB As Long, cBOR As Long, cAK As Long, wbs As String, wsH As Worksheet, r0 As Long, peil As Date
    Dim waarden(1 To 3) As Double, heeft(1 To 3) As Boolean, kg As Variant, k As Long, mi As Variant, naam As String, typ As String

    f = Application.GetOpenFilename("Excel-bestanden (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                    "Kies het oude RDFO-bestand (met tabbladen Budget en Prognose Q1/Q2/Q3)")
    If VarType(f) = vbBoolean Then Exit Sub
    On Error GoTo Fout
    Application.ScreenUpdating = False
    Set wsH = ThisWorkbook.Worksheets("Historie")
    Set wbS = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    tabs = Array("Budget", "Prognose Q1", "Prognose Q2", "Prognose Q3", "Prognose Q4")
    kg = Array("Omzet FY", "BOR FY", "AK-dekking FY")
    With Nm("L_WBS")
        r0 = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        If Trim(CStr(.Cells(.Rows.Count, 1).Value)) <> "" Then r0 = .Row + .Rows.Count
        If r0 < .Row Then r0 = .Row
    End With
    For Each tabblad In tabs
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbS.Worksheets(CStr(tabblad))
        On Error GoTo Fout
        If ws Is Nothing Then GoTo Volgende
        versie = IIf(tabblad = "Budget", "Budget", Right(CStr(tabblad), 2))
        ' kopregel en kolommen zoeken
        kopRij = 0: cBOB = 0: cBOR = 0: cAK = 0: jaar = 0
        For r = 1 To 15
            If LCase(Trim(CStr(ws.Cells(r, 1).Value))) Like "wbs-element*" Then kopRij = r: Exit For
        Next r
        For r = 1 To 5
            For c = 1 To 40
                If LCase(CStr(ws.Cells(r, c).Value)) Like "bob 20##*" And cBOB = 0 Then cBOB = TotaalKolom(ws, r, c): jaar = CLng(Mid(CStr(ws.Cells(r, c).Value), 5, 4))
                If LCase(CStr(ws.Cells(r, c).Value)) Like "bor 20##*" And cBOR = 0 Then cBOR = TotaalKolom(ws, r, c)
                If LCase(CStr(ws.Cells(r, c).Value)) Like "ak dekking 20##*" And cAK = 0 Then cAK = TotaalKolom(ws, r, c)
            Next c
        Next r
        If kopRij = 0 Or cBOB = 0 Or jaar = 0 Then
            verslag = verslag & vbCr & "   - " & tabblad & ": opzet niet herkend, overgeslagen"
            GoTo Volgende
        End If
        If Application.WorksheetFunction.CountIfs(Nm("L_Versiejaar"), jaar, Nm("L_Versie"), versie, Nm("L_Status"), "Vastgesteld") > 0 Then
            verslag = verslag & vbCr & "   - " & versie & " " & jaar & ": staat al in Historie, overgeslagen"
            GoTo Volgende
        End If
        If versie = "Budget" Then
            peil = DateSerial(jaar - 1, 12, 31)
        Else
            peil = DateSerial(jaar, CLng(Right(versie, 1)) * 3 + 1, 1) - 1
        End If
        n = 0: leeg = 0
        r = kopRij + 1
        Do While leeg < 5 And r < kopRij + 3000
            wbs = Trim(CStr(ws.Cells(r, 1).Value))
            If wbs = "" Or LCase(Left(wbs, 6)) = "totaal" Then
                leeg = leeg + 1
            Else
                leeg = 0
                waarden(1) = Waarde(ws, r, cBOB, heeft(1))
                waarden(2) = Waarde(ws, r, cBOR, heeft(2))
                waarden(3) = Waarde(ws, r, cAK, heeft(3))
                mi = Application.Match(wbs, Nm("Map_WBS"), 0)
                If IsError(mi) Then
                    naam = CStr(ws.Cells(r, 3).Value)
                    typ = IIf(InStr(1, CStr(ws.Cells(r, 2).Value), "Deel", vbTextCompare) > 0, "Deelneming", "Opstal")
                Else
                    naam = CStr(Nm("Map_Naam").Cells(mi, 1).Value)
                    typ = CStr(Nm("Map_Type").Cells(mi, 1).Value)
                End If
                For k = 1 To 3
                    If heeft(k) And Abs(waarden(k)) > 0.5 Then
                        If r0 > HISTORIE_LAATSTE_RIJ Then Err.Raise vbObjectError + 1, , "Historie is vol (maximaal rij " & HISTORIE_LAATSTE_RIJ & ")."
                        wsH.Cells(r0, 1).Resize(1, 18).Value = Array("MIGR-" & versie & "-" & jaar, BestandsNaam(CStr(f)), Date, "Vastgesteld", jaar, versie, 0, _
                            peil, jaar, wbs, Left(wbs, 8), naam, typ, "", kg(k - 1), waarden(k), "EUR", "Overgezet uit oude RDFO, tabblad " & tabblad)
                        wsH.Cells(r0, 3).NumberFormat = "dd-mm-yyyy"
                        wsH.Cells(r0, 8).NumberFormat = "dd-mm-yyyy"
                        wsH.Cells(r0, 1).Resize(1, 18).Font.Size = 8
                        r0 = r0 + 1
                        n = n + 1
                    End If
                Next k
            End If
            r = r + 1
        Loop
        LogRegel "MIGR-" & versie & "-" & jaar, "Oude RDFO (" & tabblad & ")", BestandsNaam(CStr(f)), peil, jaar, versie, n, "Vastgesteld", "Eenmalig overgezet"
        verslag = verslag & vbCr & "   - " & versie & " " & jaar & ": " & n & " regels"
        totN = totN + n
Volgende:
    Next tabblad
    wbS.Close SaveChanges:=False
    Application.CalculateFull
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("RDFO").Activate
    MsgBox "Oude prognoses overgezet (" & totN & " regels in Historie):" & verslag & vbCr & vbCr & _
           "Kijk op het tabblad RDFO of de kolommen Budget/Q1/Q2/Q3 kloppen met je oude bestand." & vbCr & _
           "Klopt iets niet? Sluit dit bestand zonder opslaan en probeer opnieuw.", vbInformation, "Oude prognoses overzetten"
    Exit Sub
Fout:
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Overzetten is niet gelukt: " & Err.Description, vbExclamation
End Sub

Private Function TotaalKolom(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    ' kolom 'Totaal' in de regel onder de groepskop; anders 0 (dan worden Q1..Q4 opgeteld, kolom wordt dan negatief teruggegeven)
    Dim k As Long
    For k = c To c + 5
        If LCase(Trim(CStr(ws.Cells(r + 1, k).Value))) = "totaal" Then TotaalKolom = k: Exit Function
    Next k
    TotaalKolom = -c
End Function

Private Function Waarde(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByRef gevonden As Boolean) As Double
    ' bedragen in de oude tabbladen staan in EUR x 1.000
    Dim k As Long, v As Variant
    gevonden = False
    If c = 0 Then Exit Function
    If c > 0 Then
        v = ws.Cells(r, c).Value
        If IsNumeric(v) And Not IsEmpty(v) Then Waarde = CDbl(v) * 1000: gevonden = True
    Else
        For k = -c To -c + 3
            v = ws.Cells(r, k).Value
            If IsNumeric(v) And Not IsEmpty(v) Then Waarde = Waarde + CDbl(v) * 1000: gevonden = True
        Next k
    End If
End Function

' ================================================================= 1c. BESTAND LEEGMAKEN
Public Sub RDFO_LeegMaken()
    Dim antw As String, il As Range
    If MsgBox("Hiermee maak je een leeg format. Dit wordt gewist:" & vbCr & _
              "   - ingelezen ZPMW03, Woningetalage en AK-percentages" & vbCr & _
              "   - projectlijst (Invoer › 2), aanvullende posten, toelichtingen, koperstype," & vbCr & _
              "     gebieden, funnel en target, DAEB, kansen en risico's, reserves" & vbCr & _
              "   - stoplicht- en scenario-invoer (tabblad Bandbreedte)" & vbCr & _
              "   - Historie: alle vastgelegde prognoses en het importlog" & vbCr & vbCr & _
              "Instellingen (regio, uitgangspunten stoplicht) blijven staan." & vbCr & _
              "Tip: sla eerst een kopie op via 'Opslaan als'. Doorgaan?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Bestand leegmaken") <> vbYes Then Exit Sub
    antw = InputBox("Typ LEEG om te bevestigen.", "Bestand leegmaken")
    If UCase(Trim(antw)) <> "LEEG" Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    With ThisWorkbook.Worksheets("Bron_ZPMW03")
        .Rows("2:" & .Rows.Count).ClearContents
    End With
    With ThisWorkbook.Worksheets("Bron_Woningetalage")
        .Rows("2:" & .Rows.Count).ClearContents
    End With
    With ThisWorkbook.Worksheets("Bron_AK")
        .Rows("2:" & .Rows.Count).ClearContents
    End With
    ThisWorkbook.Worksheets("Projectdetail").Range("C4").Value = ""
    With Nm("Map_WBS").Resize(, 19)
        .ClearContents
        .Interior.Color = Nm("P_Regio").Interior.Color
    End With
    With Nm("Map_AKMee")
        .ClearContents
        .Interior.Color = Nm("P_Regio").Interior.Color
    End With
    Nm("Gebieden").ClearContents
    Nm("MJ_Funnel").ClearContents
    Nm("MJ_Target").ClearContents
    Nm("DB_Omzet").Resize(, 7).ClearContents
    Nm("KR_Project").Resize(, 7).ClearContents
    Nm("RS_Project").Resize(, 6).ClearContents
    Nm("A_Doel").Resize(, 12).ClearContents
    Nm("T_Key").Offset(0, -8).Resize(, 5).ClearContents
    Nm("T_Txt").Resize(, 3).ClearContents
    Nm("KT_WBS").Resize(, 6).ClearContents
    Nm("KT_WBS").Offset(0, 10).ClearContents
    Nm("SL_Invoer1").ClearContents
    Nm("SL_Invoer2").ClearContents
    Nm("SL_Stempel").ClearContents
    Nm("L_ImportID").Resize(, 18).ClearContents
    Set il = ThisWorkbook.Worksheets("Historie").Range(ThisWorkbook.Worksheets("Historie").Cells(6, IMPORTLOG_KOLOM), _
             ThisWorkbook.Worksheets("Historie").Cells(HISTORIE_LAATSTE_RIJ, IMPORTLOG_KOLOM + 10))
    il.ClearContents
    Nm("P_BronZ").Value = "": Nm("P_BronE").Value = ""
    Nm("P_RegelsZ").Value = 0: Nm("P_RegelsE").Value = 0
    Nm("P_ImportID").Value = "": Nm("P_ToelCorrectie").Value = ""
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Start").Activate
    MsgBox "Het bestand is leeg. Volgende stappen:" & vbCr & _
           "1. Bronnen inlezen (ZPMW03 en Woningetalage)" & vbCr & _
           "2. Nieuwe projecten toevoegen en Type controleren" & vbCr & _
           "3. Controle nalopen, daarna Prognose vastleggen", vbInformation, "Bestand leegmaken"
    Exit Sub
Fout:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Leegmaken is niet helemaal gelukt: " & Err.Description, vbExclamation
End Sub

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
        antw = MsgBox("Prognose " & versie & " vastleggen?" & vbCr & vbCr & _
                      "Omzet " & Nm("P_UploadVJ").Value & ": € " & Format(Application.WorksheetFunction.Sum(Nm("V_OmzFY")) / 1000000, "0.0") & " mln" & vbCr & _
                      "Woningverkopen " & Nm("P_UploadVJ").Value & ": " & Format(Application.WorksheetFunction.SumIfs(Nm("V_W_HuidigJaar"), Nm("V_Meetellen"), "Ja"), "0") & vbCr & vbCr & _
                      "Daarna veranderen deze cijfers niet meer.", vbOKCancel + vbQuestion, "Prognose vastleggen")
        If antw <> vbOK Then Exit Sub
    End If

    ' eerste vrije rij: onder de laatste gevulde WBS-cel van het grootboek (niet afhankelijk van een formule)
    With Nm("L_WBS")
        r0 = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        If Trim(CStr(.Cells(.Rows.Count, 1).Value)) <> "" Then r0 = .Row + .Rows.Count
        If r0 < .Row Then r0 = .Row
    End With
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
    ThisWorkbook.Worksheets("RDFO").Activate
    If MsgBox("Klaar: " & versie & " is vastgelegd (" & n & " regels in Historie)." & vbCr & vbCr & _
              "Bestand nu opslaan?", vbYesNo + vbInformation, "Prognose vastleggen") = vbYes Then ThisWorkbook.Save
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
    Dim overgeslagen As String, stap As String

    On Error GoTo Fout
    stap = "herberekenen"
    Application.CalculateFull
    stap = "grafieken bijwerken"
    On Error Resume Next   ' niet essentieel voor de presentatie
    RDFO_OutlookJaren
    On Error GoTo Fout
    Set terug = ActiveSheet
    Set wsD = ThisWorkbook.Worksheets("Dias")
    zicht = wsD.Visible
    wsD.Visible = xlSheetVisible

    stap = "PowerPoint starten"
    Set pp = CreateObject("PowerPoint.Application")
    pp.Visible = True
    Set pres = pp.Presentations.Add(True)
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    stap = "titeldia"
    TitelDia pres
    stap = "dia's"

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
            ' elk onderdeel apart: lukt er één niet, dan gaat de rest gewoon door en volgt aan het eind een melding
            On Error Resume Next
            Err.Clear
            Select Case soort
                Case "kpi": Tegels sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value
                Case "tabel": Tabel sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 10).Text
                Case "grafiek": Grafiek sld, wsD, CLng(spec.Cells(r, 5).Value), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value, spec.Cells(r, 10).Text
                Case "opsomming": Opsomming sld, Nm(spec.Cells(r, 5).Text), spec.Cells(r, 6).Value, spec.Cells(r, 7).Value, spec.Cells(r, 8).Value, spec.Cells(r, 9).Value
            End Select
            If Err.Number <> 0 Then
                overgeslagen = overgeslagen & vbCr & "   - dia " & sld.SlideIndex & " (" & soort & " " & spec.Cells(r, 5).Text & "): " & Err.Description
                Err.Clear
            End If
            On Error GoTo Fout
        End If
    Next r

    stap = "opslaan"
    wsD.Visible = zicht
    terug.Activate
    pad = OpslagMap() & Application.PathSeparator & "RDFO " & T("P_Regio") & " " & T("Sel_Versie") & " " & T("P_VJ") & " " & Format(Now, "yyyy-mm-dd hhnn") & ".pptx"
    pres.SaveAs pad
    If overgeslagen = "" Then
        MsgBox "PowerPoint gemaakt en opgeslagen:" & vbCr & pad, vbInformation, "Maak PowerPoint"
    Else
        MsgBox "PowerPoint gemaakt en opgeslagen:" & vbCr & pad & vbCr & vbCr & _
               "Deze onderdelen zijn overgeslagen:" & overgeslagen & vbCr & vbCr & _
               "Stuur deze melding door als je wilt dat het wordt opgelost.", vbExclamation, "Maak PowerPoint"
    End If
    Exit Sub
Fout:
    melding = Err.Description & IIf(stap <> "", " (bij: " & stap & ")", "")
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
    ' getoond: jaar 3..6 (= RJ+1 .. RJ+4), en alleen jaren met cijfers
    n = Nm("OUT_N").Value
    If n > 6 Then n = 6
    If n < 3 Then n = 3
    For i = 1 To 8
        wsO.Columns(2 + i).Hidden = (i < 3 Or i > n)
    Next i
    wsO.Columns(11).Hidden = True
    ' grafiek: blad|index|basisrij (rijen basis..basis+3 = realisatie, prognose, totaal, budget op Outlook vanaf kolom P)
    For Each doel In Split(Nm("OUT_Grafieken").Value, ";")
        delen = Split(doel, "|")
        Set co = ThisWorkbook.Worksheets(delen(0)).ChartObjects(CLng(delen(1)))
        basis = CLng(delen(2))
        For s = 1 To co.Chart.SeriesCollection.Count
            rij = basis + s - 1
            co.Chart.SeriesCollection(s).Values = wsO.Range(wsO.Cells(rij, 18), wsO.Cells(rij, 15 + n))
            co.Chart.SeriesCollection(s).XValues = wsO.Range(wsO.Cells(5, 18), wsO.Cells(5, 15 + n))
        Next s
    Next doel
    ' bandbreedte (liggend): jaren op de verticale as = rapportagejaar .. +2
    Dim rj As Long
    rj = Nm("P_RJ").Value
    For Each doel In Split(Nm("BB_Grafieken").Value, ";")
        delen = Split(doel, "|")
        Set co = ThisWorkbook.Worksheets(delen(0)).ChartObjects(CLng(delen(1)))
        With co.Chart.Axes(2)
            .MinimumScale = rj - 1
            .MaximumScale = rj + 3
            .TickLabels.NumberFormat = "[<" & rj & "]"""";[>" & (rj + 2) & "]"""";0"
        End With
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
    ' volledig geel met donkerblauwe letters
    Dim sld As Object, sh As Object
    Set sld = pres.Slides.Add(1, 12)
    Vlak sld, 0, 0, 960, 540, GEEL
    Vlak sld, 0, 0, 960, 10, NAVY
    Set sh = sld.Shapes.AddTextbox(1, 72, 64, 600, 30)
    Tekst sh, T("P_Regio"), 16, NAVY, True
    Set sh = sld.Shapes.AddTextbox(1, 72, 180, 820, 130)
    Tekst sh, T("PPT_Titel"), 46, NAVY, True
    Vlak sld, 74, 322, 120, 8, NAVY
    Set sh = sld.Shapes.AddTextbox(1, 72, 346, 820, 34)
    Tekst sh, T("PPT_Sub"), 20, NAVY, False
    Set sh = sld.Shapes.AddTextbox(1, 72, 470, 820, 20)
    Tekst sh, "Gemaakt op " & Format(Now, "dd-mm-yyyy") & "  ·  bron: " & ThisWorkbook.Name, 10, NAVY, False
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
        Tekst sh, CelTekst(blok.Cells(k, 1)), 8.5, NAVY, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 6, y + 24, tw - 12, 28)
        Tekst sh, CelTekst(blok.Cells(k, 2)), 19, NAVY, True
        Set sh = sld.Shapes.AddTextbox(1, tx + 6, y + h - 18, tw - 12, 14)
        Tekst sh, CelTekst(blok.Cells(k, 3)), 8, GRIJS, False
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
    Dim breedte As Single, kleur As Variant

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

    ' zonder hoogte: PowerPoint kiest zelf een geldige hoogte; na het verkleinen van de letters krimpen de rijen
    Set shp = sld.Shapes.AddTable(nr, nk, x, y, w)
    Set tbl = shp.Table
    On Error Resume Next
    tbl.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", False
    On Error GoTo 0
    tbl.FirstRow = True
    tbl.HorizBanding = False
    On Error Resume Next   ' PowerPoint weigert te smalle kolommen: dan blijft de standaardbreedte staan
    For j = 1 To nk
        breedte = w * gew.Cells(1, kol(j)).Value / som
        If breedte < 20 Then breedte = 20
        tbl.Columns(j).Width = breedte
    Next j
    On Error GoTo 0

    On Error Resume Next   ' opmaak die PowerPoint niet accepteert wordt overgeslagen; de tekst komt er altijd in
    For i = 1 To nr
        For j = 1 To nk
            Set c = bron.Cells(reg(i), kol(j))
            txt = CelTekst(c)
            tbl.Cell(i, j).Shape.TextFrame.TextRange.Text = txt
            bol = (txt = ChrW(9679) Or txt = ChrW(9675))
            With tbl.Cell(i, j).Shape
                .TextFrame.MarginLeft = 3: .TextFrame.MarginRight = 3
                .TextFrame.MarginTop = 1: .TextFrame.MarginBottom = 1
                .TextFrame.VerticalAnchor = 3
                With .TextFrame.TextRange.Font
                    .Name = "Arial"
                    .Size = IIf(bol, grootte + 5, grootte)
                    If i = 1 Then
                        .Color.RGB = WIT
                        .Bold = True
                    Else
                        kleur = c.DisplayFormat.Font.Color
                        If IsNull(kleur) Then kleur = 0
                        If kleur < 0 Or kleur > 16777215 Then kleur = 0
                        .Color.RGB = kleur
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
                ElseIf c.DisplayFormat.Interior.ColorIndex <> xlColorIndexNone And c.DisplayFormat.Interior.Color <> WIT _
                       And c.DisplayFormat.Interior.Color >= 0 And c.DisplayFormat.Interior.Color <= 16777215 Then
                    .Fill.ForeColor.RGB = c.DisplayFormat.Interior.Color
                Else
                    .Fill.ForeColor.RGB = WIT
                End If
            End With
            With tbl.Cell(i, j).Borders(3)
                .ForeColor.RGB = LIJN
                .Weight = 0.5
            End With
        Next j
        tbl.Rows(i).Height = grootte + 6
    Next i
    On Error GoTo 0
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
    If sh.Height <= 0 Or h <= 0 Then
        sh.Width = w
    ElseIf sh.Width / sh.Height > w / h Then
        sh.Width = w
    Else
        sh.Height = h
    End If
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
