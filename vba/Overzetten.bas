Attribute VB_Name = "Overzetten"
Option Explicit
' =====================================================================================
'  FO overzetten: werkbestanden van een eerdere versie van het FO-werkbestand in het nieuwste sjabloon zetten.
'  Per gekozen werkbestand: een kopie van het sjabloon naast het oude bestand (<naam>_nieuw.xlsm), daarin de macro
'  FOOvernemenTest van het sjabloon (exports en alle invoer), opslaan en sluiten. Het oude bestand verandert niet.
' =====================================================================================
Private Const BLAD As String = "Overzetten"
Private Const CEL_SJABLOON As String = "C6"
Private Const RIJ1 As Long = 12                    ' eerste resultaatrij (kop op rij 11)
Private Const TITEL As String = "FO overzetten"

Public Sub SjabloonKiezen()
    Dim pad As Variant
#If Mac Then
    pad = Application.GetOpenFilename()
#Else
    pad = Application.GetOpenFilename("FO-werkbestand (*.xlsm),*.xlsm", , "Nieuwste sjabloon kiezen")
#End If
    If VarType(pad) = vbBoolean Then Exit Sub
    ThisWorkbook.Worksheets(BLAD).Range(CEL_SJABLOON).Value = CStr(pad)
End Sub

Public Sub Overzetten()
    Dim sjabloon As String, keuze As Variant, i As Long, n As Long, fout As Long
    sjabloon = Trim$(CStr(ThisWorkbook.Worksheets(BLAD).Range(CEL_SJABLOON).Value))
    If Len(sjabloon) = 0 Or Len(Dir$(sjabloon)) = 0 Then
        SjabloonKiezen
        sjabloon = Trim$(CStr(ThisWorkbook.Worksheets(BLAD).Range(CEL_SJABLOON).Value))
        If Len(sjabloon) = 0 Then Exit Sub
    End If
#If Mac Then
    keuze = Application.GetOpenFilename(, , , , True)
#Else
    keuze = Application.GetOpenFilename("FO-werkbestanden (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Werkbestanden kiezen (meerdere mag)", , True)
#End If
    If Not IsArray(keuze) Then Exit Sub
    For i = LBound(keuze) To UBound(keuze)
        If ZetOver(sjabloon, CStr(keuze(i))) Then n = n + 1 Else fout = fout + 1
    Next i
    MsgBox n & " werkbestand(en) overgezet" & IIf(fout > 0, ", " & fout & " met een fout", "") & "." & vbCrLf & _
           "Zie de tabel op dit tabblad.", IIf(fout > 0, vbExclamation, vbInformation), TITEL
End Sub

Public Sub OverzettenTest(ByVal sjabloon As String, ByVal oud As String)
    ' zonder vragen (voor de controle buiten Excel)
    ZetOver sjabloon, oud
End Sub

Private Function NieuwePad(ByVal oud As String) As String
    Dim basis As String, p As Long, i As Long
    p = InStrRev(oud, ".")
    basis = IIf(p > 0, Left$(oud, p - 1), oud)
    NieuwePad = basis & "_nieuw.xlsm"
    i = 2
    Do While Len(Dir$(NieuwePad)) > 0
        NieuwePad = basis & "_nieuw" & i & ".xlsm"
        i = i + 1
    Loop
End Function

Private Function ZetOver(ByVal sjabloon As String, ByVal oud As String) As Boolean
    Dim nieuw As String, wb As Workbook, melding As String, status As String, r As Long, ws As Worksheet
    Dim rekenmodus As Long
    Set ws = ThisWorkbook.Worksheets(BLAD)
    r = RIJ1
    Do While Len(CStr(ws.Cells(r, 2).Value)) > 0
        r = r + 1
    Loop
    ws.Cells(r, 2).Value = oud
    rekenmodus = Application.Calculation
    On Error GoTo Fout
    If LCase$(oud) = LCase$(sjabloon) Then Err.Raise vbObjectError + 1, , "dit is het sjabloon zelf"
    nieuw = NieuwePad(oud)
    FileCopy sjabloon, nieuw
    ws.Cells(r, 3).Value = nieuw
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(nieuw, UpdateLinks:=0)
    Application.DisplayAlerts = True
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!FOOvernemenTest", oud
    If Err.Number <> 0 Then
        melding = "het gekozen sjabloon kan dit niet (kies het nieuwste sjabloon): " & Err.Description
        Err.Clear
        status = "FOUT"
    Else
        melding = CStr(wb.Worksheets("Start").Range("B4").Value)
        status = IIf(InStr(melding, "- FOUT") > 0, "FOUT", "OK")
    End If
    On Error GoTo Fout
    Application.DisplayAlerts = False
    wb.Save                                            ' de kopie is al een .xlsm: opslaan in dezelfde indeling
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ws.Cells(r, 4).Value = status
    ws.Cells(r, 5).Value = Replace(Replace(melding, vbCrLf, " "), vbLf, " ")
    ZetOver = (status = "OK")
    Exit Function
Fout:
    ws.Cells(r, 4).Value = "FOUT"
    ws.Cells(r, 5).Value = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = rekenmodus
End Function
