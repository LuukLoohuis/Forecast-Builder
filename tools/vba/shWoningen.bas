Attribute VB_Name = "shWoningen"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then Application.Run "RDFO_DeelgebiedSync", Me.Range("C3").Value, "Woningetalage"
End Sub
