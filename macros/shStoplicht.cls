Option Explicit

' Bij elke invoer in de gele kolommen (stoplicht en scenario) wordt het WBS-element van die regel vastgelegd (verborgen kolom AF).
' Controle meldt het als de projectenlijst later verschuift en de invoer bij een ander project terechtkomt.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim isect As Range, c As Range, wbsKol As Long, stempelKol As Long
    Set isect = Intersect(Target, Union(ThisWorkbook.Names("SL_Invoer1").RefersToRange, ThisWorkbook.Names("SL_Invoer2").RefersToRange))
    If isect Is Nothing Then Exit Sub
    wbsKol = ThisWorkbook.Names("SL_WBS").RefersToRange.Column
    stempelKol = ThisWorkbook.Names("SL_Stempel").RefersToRange.Column
    Application.EnableEvents = False
    For Each c In isect.Cells
        Me.Cells(c.Row, stempelKol).Value = Me.Cells(c.Row, wbsKol).Value
    Next c
    Application.EnableEvents = True
End Sub
