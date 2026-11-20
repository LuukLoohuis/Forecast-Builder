Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    On Error Resume Next
    CashflowNaarPowerPoint.KnoppenControleren
    FOWerkbestand.StartKnoppen
    MeldingMacros False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    MeldingMacros True     ' bewaard bestand toont de melding als de macro's later uit staan
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    MeldingMacros False
End Sub

Private Sub MeldingMacros(ByVal tonen As Boolean)
    ' Start J6: melding 'geen knoppen? macro's staan uit'; met macro's aan weg (de knoppen liggen erover)
    Dim t As String
    On Error Resume Next
    t = "Geen knoppen? Dan staan de macro's uit. Sluit het bestand, klik er met rechts op > Eigenschappen > " & _
        "vink Blokkering opheffen aan > OK. Open het opnieuw en klik Inhoud inschakelen."
    With ThisWorkbook.Worksheets("Start").Range("J6")
        If tonen Then .Value = t Else .Value = ""
    End With
    ThisWorkbook.Saved = True
End Sub
