Option Explicit

' Bij elke invoer in de gele kolommen wordt het WBS-element van die regel vastgelegd (verborgen kolom P).
' Controle meldt het als de projectenlijst later verschuift en de invoer bij een ander project terechtkomt.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim isect As Range, c As Range
    Set isect = Intersect(Target, Me.Range("E8:H107"))
    If isect Is Nothing Then Exit Sub
    Application.EnableEvents = False
    For Each c In isect.Cells
        Me.Cells(c.Row, "P").Value = Me.Cells(c.Row, "O").Value
    Next c
    Application.EnableEvents = True
End Sub
