Attribute VB_Name = "shStart"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' projectsoort (D6): standaard of 100% DAEB -> tabbladen tonen of verbergen
    If Intersect(Target, Me.Range("D6")) Is Nothing Then Exit Sub
    On Error Resume Next
    FOWerkbestand.WeergaveBijwerken
End Sub
