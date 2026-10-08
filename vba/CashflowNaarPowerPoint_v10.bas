Attribute VB_Name = "CashflowNaarPowerPoint"
Option Explicit

' =====================================================================================
'  Cashflow_scenario -> PowerPoint                       (late binding, Windows en Mac)
'  versie 10: typeblokken (maximaal MAX_TYPES woningtypes), knoppen Type invoegen / Type verwijderen
'  v11: scenariolijn als reeks in de cashflowgrafiek (dia 3); reeks weg als de kop op '(uit)' eindigt
'  v12: dia 5 bouwtermijnen-tijdlijn (BT_GRAFIEK), dia 6 verkoop en transport per kwartaal (VP1_/VP2_GRAFIEK),
'       dia 7 van verkoop naar omzet (VO_GRAFIEK)
'  v13: bouwtermijnen per woningtype: per termijn een baan per type (blok AA7, 108 kolommen, tot MAX_RIJEN_BT rijen,
'       aantal uit Model!BY96; tijd-as in kwartaalindex vast op Model!BY97..BY99; reeksen met kop '(uit)' gaan weg),
'       VP-blok vanaf EF7, VO-blok vanaf FG7
'  v14: dia 8 'Opbrengsten per woningtype' (PT_GRAFIEK, blok FM7, 8 kolommen): de macro maakt per woningtype met aantal > 0
'       een kopie van dia 8, zet het type in de keuzecel op het Dashboard (CEL_PT_KEUZE), rekent door en vult de kopie;
'       dia 8 zelf gaat daarna weg. Blok 5 op Woningtypes (soort DAEB en fees, rijen 48-63) bij Type invoegen/verwijderen.
'       Tabellen vanaf kolom FV; sjabloon v14
'  v14: UitFOOphalen (knop 'Ophalen uit FO' op tab Invoer): leest een FO-werkboek ('1. Cashflow', 'FO - actuals',
'       'CF - opbrengsten') en vult tab Invoer (cashflow, verkocht, transport), tab Woningtypes (typen, koopsom, grondtermijn,
'       start bouw, bouwtermijnen, extra's, DAEB-fees) en de FO-cellen op het Dashboard (zelfde logica als forecast_builder/fo.py)
'
'  NaarPowerPoint        opent het sjabloon als kopie, vult teksten, tabellen en grafieken
'                        vanaf tabblad "PowerPoint" en bewaart een nieuwe presentatie
'  ControleerSjabloon    kijkt of het sjabloon alle vormnamen heeft die de knop vult
'  KnoppenControleren    zet ontbrekende knoppen op de tabbladen (draait bij het openen)
'  TypeInvoegen          voegt een leeg typeblok in (drie kolommen op Woningtypes, twee op Invoer)
'  TypeVerwijderen       haalt een typeblok weg (zelfde kolommen); alles rechts schuift mee
'  SchikDia5             na het vullen van VT_TABEL (dia 5): bij meer dan 4 types rijen en letters
'                        kleiner, zo nodig de kaart met BT_GRAFIEK korter, en de tabelkaart sluit om de tabel
'
'  UitFOOphalen         vult Invoer en Woningtypes vanuit een FO-werkboek (zie v14 hierboven)
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
Private Const CEL_AS_MIN As String = "BY97"        ' ondergrens tijd-as BT_GRAFIEK: kwartaalindex 1 jan eerste jaar (Model, t_first)
Private Const CEL_AS_MAX As String = "BY99"        ' bovengrens tijd-as: kwartaalindex na het laatste jaar (Model, t_end)
Private Const CEL_SJABLOON As String = "D56"
Private Const CEL_NAAM As String = "D57"
Private Const SJABLOON_STANDAARD As String = "Kwartaal_Template_cashflow_v14.pptx"
Private Const MAX_RIJEN As Long = 60               ' periodes (kwartalen) per grafiekblok
Private Const MAX_RIJEN_BT As Long = 132           ' rijen in het bouwtermijnenblok (koprijen, banen type x termijn, scheidingsrijen)
Private Const KOL_BT As Long = 108                 ' kolommen in het bouwtermijnenblok (label + 107 reeksen; layout.PP_BT_KOL)
Private Const BLOK_PT As String = "FM7"            ' blok grafiek per woningtype (dia 8): kwartaal, zes reeksen, realisatie
Private Const KOL_PT As Long = 8
Private Const CEL_PT_KEUZE As String = "B120"      ' keuzecel op het Dashboard: naam van het type in de grafiek per woningtype
' FO-koppeling (Dashboard): bestand, actuals t/m, opbrengsten en kosten totaal uit '1. Cashflow', datum; actuals-knop W5/X5
Private Const CEL_FO_BESTAND As String = "W69"
Private Const CEL_FO_ACTUALS As String = "W70"
Private Const CEL_FO_OPBR As String = "W71"
Private Const CEL_FO_KOSTEN As String = "W72"
Private Const CEL_FO_DATUM As String = "W73"
Private Const CEL_ACTUALS_JAAR As String = "W5"
Private Const CEL_ACTUALS_KW As String = "X5"
Private Const TYPE_RIJ_KOOPSOM As Long = 9
Private Const TYPE_RIJ_STARTJAAR As Long = 10
Private Const MAX_PERIODES As Long = 60
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
Private Const TYPE_RIJ_X1 As Long = 33        ' extra opbrengsten per woning (blok 3 op tabblad Woningtypes)
Private Const TYPE_RIJ_XN As Long = 38
Private Const TYPE_RIJ_VJAAR As Long = 43     ' start bouw vorige prognose (blok 4): jaar / kwartaal
Private Const TYPE_RIJ_VKW As Long = 44
Private Const TYPE_RIJ_SOORT As Long = 48     ' blok 5: soort (niet-DAEB / DAEB)
Private Const TYPE_RIJ_FC1 As Long = 51       ' fee-componenten: naam (kolom 1), totaal (kolom 2)
Private Const TYPE_RIJ_FCN As Long = 52
Private Const TYPE_RIJ_F1 As Long = 54        ' fee-termijnen: mijlpaal (1), bedrag of % (2), kwartaal (3)
Private Const TYPE_RIJ_FN As Long = 63
Private Const TYPE_RIJ_AANTAL As Long = 8
Private Const INVOER_COL1 As Long = 12        ' kolom L: eerste paar (verkocht, transport) op tabblad Invoer
Private Const INVOER_BREEDTE As Long = 2
Private Const INVOER_RIJ1 As Long = 8
Private Const INVOER_RIJN As Long = 500        ' zelfde grens als de formules (layout.IN_ROWMAX)

Private Function Grafieken() As Variant
    ' vormnaam op de dia, kopcel van het blok op tabblad PowerPoint, aantal kolommen,
    ' cel op tabblad Model met het aantal rijen (leeg = aantal periodes uit CEL_N)
    ' CF (dia 3): acht kolommen; de achtste (kolom N) is de scenariolijn, in het sjabloon als reeks op kolom H van het
    '   gegevensblad (staat de lijn uit, dan eindigt de kop op '(uit)' en haalt VulGrafiek de reeks weg)
    ' BT (dia 5): bouwtermijnen-tijdlijn: koprij jaartallen, koprij kwartaalnummers, dan per termijn een baan per type
    '   (categorie in kolom A, daarna de reeksen: stapel 1 = onzichtbaar, grijs, onzichtbaar, vorige prognose per type (10),
    '   grijs, onzichtbaar, jaarblokjes (16); stapel 2 = onzichtbaar, huidige planning per type (10), onzichtbaar,
    '   kwartaalblokjes (64)); tijd in kwartaalindex; aantal rijen in Model!BY96; reeksen met kop '(uit)' (typen zonder
    '   termijnen, jaren buiten het bereik) gaan weg; na het vullen zet ZetTijdAs beide waarde-assen op BY97..BY99
    ' VP1/VP2 (dia 6): verkocht en getransporteerd per kwartaal, beide panelen lezen hetzelfde blok van 26 kolommen
    Grafieken = Array(Array("CF_GRAFIEK", "G7", 8, ""), Array("SC_GRAFIEK", "P7", 10, ""), _
                      Array("BT_GRAFIEK", "AA7", KOL_BT, "BY96"), _
                      Array("VP1_GRAFIEK", "EF7", 26, ""), Array("VP2_GRAFIEK", "EF7", 26, ""), _
                      Array("VO_GRAFIEK", "FG7", 5, ""))
End Function

Private Function Tabellen() As Variant
    ' vormnaam op de dia, kopcel van de tabel, aantal rijen met kop, aantal kolommen, lege regels overslaan (1 = ja)
    Tabellen = Array(Array("SC_TABEL", "FV7", 9, 4, 0), Array("VT_TABEL", "FV18", MAX_TYPES + 1, 8, 1))
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
        MaakKnop ws, ws.Range("W35:X36"), "btnNaarPowerPoint", "Naar PowerPoint", "NaarPowerPoint", RGB(23, 54, 93), 0, 0, 2
    End If
    Set ws = ThisWorkbook.Worksheets(SH_PP)
    If Not BestaatVorm(ws, "btnNaarPowerPoint2") Then
        MaakKnop ws, ws.Range("C4:C5"), "btnNaarPowerPoint2", "Naar PowerPoint", "NaarPowerPoint", RGB(23, 54, 93), 150, 0, 2
    End If
    If Not BestaatVorm(ws, "btnControle") Then
        MaakKnop ws, ws.Range("D4:D5"), "btnControle", "Sjabloon controleren", "ControleerSjabloon", RGB(91, 97, 105), 150, 0, 2
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(SH_INVOER)
    If Not ws Is Nothing Then
        If Not BestaatVorm(ws, "btnUitFO") Then
            MaakKnop ws, ws.Range("L2:N3"), "btnUitFO", "Ophalen uit FO", "UitFOOphalen", RGB(23, 54, 93), 150, 0, 2
        End If
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
    If Application.WorksheetFunction.Count(wsT.Range(wsT.Cells(TYPE_RIJ_X1, c + 1), wsT.Cells(TYPE_RIJ_XN, c + 2))) > 0 Then Exit Function
    If Application.WorksheetFunction.Count(wsT.Range(wsT.Cells(TYPE_RIJ_VJAAR, c), wsT.Cells(TYPE_RIJ_VKW, c))) > 0 Then Exit Function
    If Len(CStr(wsT.Cells(TYPE_RIJ_SOORT, c).Value)) > 0 Then Exit Function
    If Application.WorksheetFunction.Count(wsT.Range(wsT.Cells(TYPE_RIJ_FC1, c + 1), wsT.Cells(TYPE_RIJ_FCN, c + 1))) > 0 Then Exit Function
    If Application.WorksheetFunction.CountA(wsT.Range(wsT.Cells(TYPE_RIJ_F1, c + 1), wsT.Cells(TYPE_RIJ_FN, c + 2))) > 0 Then Exit Function
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
        ws.Range(ws.Cells(TYPE_RIJ_X1, col + 1), ws.Cells(TYPE_RIJ_XN, col + 2)).ClearContents
        ws.Range(ws.Cells(TYPE_RIJ_VJAAR, col), ws.Cells(TYPE_RIJ_VKW, col)).ClearContents
        ws.Cells(TYPE_RIJ_SOORT, col).ClearContents
        ws.Range(ws.Cells(TYPE_RIJ_FC1, col + 1), ws.Cells(TYPE_RIJ_FCN, col + 1)).ClearContents
        ws.Range(ws.Cells(TYPE_RIJ_F1, col + 1), ws.Cells(TYPE_RIJ_FN, col + 2)).ClearContents
    Else
        ws.Range(ws.Cells(INVOER_RIJ1, col), ws.Cells(INVOER_RIJN, col + breedte - 1)).ClearContents
    End If
End Sub

' -------------------------------------------------------------------------------------
'  3. Naar PowerPoint
' -------------------------------------------------------------------------------------
Public Sub NaarPowerPoint()
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim stap As String, n As Long, nG As Long, pad As String, mist As String, fouten As String
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
        ' aantal rijen: uit de opgegeven Model-cel (bouwtermijnenblok), anders het aantal periodes;
        ' geen rijen (bijvoorbeeld geen bouwtermijnen): een lege rij, zodat de voorbeeldgegevens uit het sjabloon verdwijnen
        If Len(CStr(g(i)(3))) > 0 Then
            nG = AantalRijen(CStr(g(i)(3)), MAX_RIJEN_BT)
            If nG < 1 Then nG = 1
        Else
            nG = n
        End If
        VulGrafiek pres, CStr(g(i)(0)), wsP.Range(CStr(g(i)(1))), nG, CLng(g(i)(2)), mist, fouten
    Next i

    stap = "dia per woningtype"
    VulDiasPerType pres, wsP, n, mist, fouten

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
    AantalPeriodes = AantalRijen(CEL_N, MAX_RIJEN)
End Function

Private Function AantalRijen(cel As String, maxN As Long) As Long
    ' getal uit een cel op tabblad Model, begrensd op maxN; 0 bij leeg, fout of geen getal
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_MODEL).Range(cel).Value
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then AantalRijen = CLng(v)
    If AantalRijen < 0 Then AantalRijen = 0
    If AantalRijen > maxN Then AantalRijen = maxN
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
    ' na het vullen van VT_TABEL (dia 5, onder de bouwtermijnengrafiek BT_GRAFIEK in kaart VT_KAART):
    ' bij meer dan 4 typeregels de tabel binnen de kaart en boven het dianummer houden
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
    Set grafiek = ZoekVorm(pres, "BT_GRAFIEK")
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
    ' zoekt het grafiekvak op naam (eerste dia waarop het staat) en vult het
    Dim shp As Object
    Set shp = ZoekVorm(pres, vorm)
    If shp Is Nothing Then
        mist = mist & vorm & vbCrLf
        Exit Sub
    End If
    VulGrafiekVorm shp, vorm, kop, n, nKol, fouten
End Sub

Private Sub VulGrafiekVorm(shp As Object, vorm As String, kop As Range, n As Long, nKol As Long, ByRef fouten As String)
    ' zet het blok (kop + n rijen) in het gegevensblad van de grafiek en past het bereik van elke reeks aan:
    ' X-bereik = kolom A, Y-bereik = de kolomletter uit de SERIES-formule van die reeks (KolomVanReeks). Reeksnamen die
    ' naar rij 1 van het gegevensblad verwijzen volgen zo de koptekst; namen als tekst in de formule (BT_GRAFIEK) blijven.
    Dim ch As Object, wbE As Object, wsE As Object
    Dim arr() As Variant, r As Long, c As Long, v As Variant, i As Long, kol As String
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
    ' wissen tot de grootste bloklengte (MAX_RIJEN_BT >= MAX_RIJEN): veilig voor alle grafieken, ook als het sjabloon
    ' meer voorbeeldrijen bevat dan het blok nu heeft
    wsE.Range("A1").Resize(MAX_RIJEN_BT + 1, nKol).ClearContents
    wsE.Range("A1").Resize(n + 1, nKol).Value = arr
    For i = 1 To ch.SeriesCollection.Count
        kol = KolomVanReeks(CStr(ch.SeriesCollection(i).Formula))
        If Len(kol) > 0 Then ZetBereik ch.SeriesCollection(i), wsE, kol, n
    Next i
    ' reeksen met '(uit)' in de kop gaan weg: de scenariolijn in de cashflowgrafiek, in de bouwtermijnengrafiek de typen
    ' zonder termijnen en de jaren buiten het bereik, in de grafiek per type de onderdelen die het type niet heeft
    If vorm = "CF_GRAFIEK" Or vorm = "BT_GRAFIEK" Or vorm = "PT_GRAFIEK" Then VerwijderUitgezetteReeksen ch, wsE
    If vorm = "BT_GRAFIEK" Then ZetTijdAs ch
    wbE.Close
    Set wbE = Nothing
    Exit Sub
Mislukt:
    fouten = fouten & vorm & ": " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wbE Is Nothing Then wbE.Close
    On Error GoTo 0
End Sub

Private Sub VerwijderUitgezetteReeksen(ch As Object, wsE As Object)
    ' Een reeks waarvan de kop in het gegevensblad eindigt op '(uit)' (de scenariolijn bij 'Scenariolijn tonen = nee',
    ' tab PowerPoint kolom N; in de bouwtermijnengrafiek typen zonder termijnen en jaren buiten het bereik) wordt uit de
    ' grafiek verwijderd, zodat hij niet als lege vermelding in de legenda staat.
    ' Gekozen voor verwijderen in plaats van Legend.LegendEntries(k).Delete: LegendEntries telt alleen zichtbare entries
    ' (in sjabloon v11 zijn de entries van reeks 0 en 1, de vlakken voorfinanciering/positief saldo, verborgen, dus de
    ' scenariolijn zou entry 5 zijn) en die telling verschuift zodra iemand het sjabloon aanpast. Verwijderen hangt niet
    ' van een index af en is veilig: het sjabloon wordt per run opnieuw geopend, alleen de nieuwe presentatie verandert.
    ' Van achteren naar voren, omdat de nummering van SeriesCollection verschuift na een Delete.
    Dim i As Long, kol As String, kopTekst As String
    On Error Resume Next
    For i = ch.SeriesCollection.Count To 1 Step -1
        kol = KolomVanReeks(CStr(ch.SeriesCollection(i).Formula))
        If Len(kol) > 0 Then
            kopTekst = ""
            kopTekst = Trim$(CStr(wsE.Range(kol & "1").Value))
            If Right$(kopTekst, 5) = "(uit)" Then ch.SeriesCollection(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub ZetTijdAs(ch As Object)
    ' BT_GRAFIEK: beide waarde-assen (tijd in kwartaalindex) vast op Model!BY97 (1 januari eerste jaar) .. BY99 (einde laatste
    ' jaar), zodat de blokjes van vorige prognose en huidige planning op dezelfde schaal staan als de koprijen.
    ' Late binding: Axes(2 = xlValue, 1 = xlPrimary / 2 = xlSecondary). Eerst het maximum, dan het minimum en nog eens
    ' het maximum: zo lukt het ook als het nieuwe bereik helemaal boven of onder de voorbeeldwaarden van het sjabloon
    ' ligt (PowerPoint weigert een minimum boven het huidige maximum en andersom). Ontbreekt de tweede as, dan blijft
    ' de rest staan: alles onder On Error.
    Dim jaar1 As Variant, jaar2 As Variant, asNr As Long, ax As Object
    On Error Resume Next
    jaar1 = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_AS_MIN).Value
    jaar2 = ThisWorkbook.Worksheets(SH_MODEL).Range(CEL_AS_MAX).Value
    If IsError(jaar1) Or IsError(jaar2) Then Exit Sub
    If Not IsNumeric(jaar1) Or Not IsNumeric(jaar2) Then Exit Sub
    If CDbl(jaar2) <= CDbl(jaar1) Then Exit Sub
    For asNr = 1 To 2
        Set ax = Nothing
        Set ax = ch.Axes(2, asNr)
        If Not ax Is Nothing Then
            ax.MaximumScale = CDbl(jaar2)
            Err.Clear
            ax.MinimumScale = CDbl(jaar1)
            Err.Clear
            ax.MaximumScale = CDbl(jaar2)
            Err.Clear
            ax.MinimumScaleIsAuto = False
            ax.MaximumScaleIsAuto = False
            Err.Clear
        End If
    Next asNr
    On Error GoTo 0
End Sub

Private Sub VulDiasPerType(pres As Object, wsP As Worksheet, n As Long, ByRef mist As String, ByRef fouten As String)
    ' Dia 8 (PT_GRAFIEK): per woningtype met aantal > 0 een kopie van de dia, gevuld met dat type. Het Dashboard heeft een
    ' keuzecel met de typenaam; het Model rekent het selectieblok (tab PowerPoint vanaf BLOK_PT) voor dat type uit. De macro
    ' zet per type de keuzecel, rekent door, kopieert de dia (achter de vorige kopie) en vult titel, ondertitel en grafiek
    ' op de kopie zelf (niet via ZoekVorm, dat de eerste dia met die vormnaam zou pakken). Daarna gaat dia 8 zelf weg en
    ' krijgt de keuzecel zijn oude waarde terug. Geen type: alleen dia 8 verwijderen.
    Dim wsT As Worksheet, wsD As Worksheet, shp As Object, sjabloon As Object, kopie As Object, rngK As Object
    Dim oud As Variant, pos As Long, c As Long, naam As String, aantal As Variant, idx As Long, nKopie As Long
    Dim rTitel As Long, rSub As Long, r As Long, fouten2 As String
    Set shp = ZoekVorm(pres, "PT_GRAFIEK")
    If shp Is Nothing Then
        mist = mist & "PT_GRAFIEK" & vbCrLf
        Exit Sub
    End If
    On Error GoTo Mislukt
    Set sjabloon = shp.Parent
    idx = sjabloon.SlideIndex
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)
    oud = wsD.Range(CEL_PT_KEUZE).Formula
    ' regels op tab PowerPoint met de titel en ondertitel van dia 8 (kolom F = vormnaam)
    For r = KPI_ROW1 To KPI_ROW1 + 300
        If Len(CelTekst(wsP.Range("C" & r))) = 0 Then Exit For
        If Trim$(CelTekst(wsP.Range("F" & r))) = "PT_TITEL" Then rTitel = r
        If Trim$(CelTekst(wsP.Range("F" & r))) = "PT_SUBTITEL" Then rSub = r
    Next r
    nKopie = 0
    For pos = 1 To MAX_TYPES
        c = TYPE_COL1 + (pos - 1) * TYPE_BREEDTE
        aantal = wsT.Cells(TYPE_RIJ_AANTAL, c).Value
        If IsNumeric(aantal) And Not IsEmpty(aantal) Then
            If CDbl(aantal) > 0 Then
                naam = CStr(wsT.Cells(TYPE_RIJ_NAAM, c).Value)
                If Len(naam) = 0 Then naam = "Type " & pos
                wsD.Range(CEL_PT_KEUZE).Value = pos          ' typenummer: het Model kiest op nummer, dus ook bij dubbele namen goed
                Application.Calculate
                Set rngK = sjabloon.Duplicate
                Set kopie = rngK.Item(1)
                nKopie = nKopie + 1
                kopie.MoveTo idx + nKopie
                On Error Resume Next                           ' titel en ondertitel: alleen als de vorm op de dia staat
                If rTitel > 0 Then kopie.Shapes("PT_TITEL").TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & rTitel))
                If rSub > 0 Then kopie.Shapes("PT_SUBTITEL").TextFrame.TextRange.Text = CelTekst(wsP.Range("D" & rSub))
                On Error GoTo Mislukt
                fouten2 = ""
                VulGrafiekVorm kopie.Shapes("PT_GRAFIEK"), "PT_GRAFIEK", wsP.Range(BLOK_PT), n, KOL_PT, fouten2
                If Len(fouten2) > 0 Then fouten = fouten & naam & " · " & fouten2
            End If
        End If
    Next pos
    sjabloon.Delete
    wsD.Range(CEL_PT_KEUZE).Formula = oud
    Application.Calculate
    Exit Sub
Mislukt:
    fouten = fouten & "dia per woningtype (" & naam & "): " & Err.Description & vbCrLf
    On Error Resume Next
    wsD.Range(CEL_PT_KEUZE).Formula = oud
    Application.Calculate
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
    ' werkt voor naam als verwijzing (Sheet1!$C$1), als tekst ("Gerealiseerd"), leeg (=SERIES(,...): delen(0) is dan
    ' '=SERIES(' en delen(2) nog steeds de waarden) en met aanhalingstekens om de bladnaam ('Blad 1'!$C$2:$C$21)
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
'  4b. Ophalen uit FO: tab Invoer en tab Woningtypes vullen vanuit een FO-werkboek
'      Alles wordt op labels gezocht (zelfde regels als forecast_builder/fo.py):
'      '1. Cashflow'     koprij 'Jaar' in kolom B; daaronder jaar, Q, kosten, % kosten, omzet, % omzet, CF, CF x 1000, CF vorig
'      'FO - actuals'    laatste boekjaar/kwartaal met boekingen -> Dashboard 'Actuals t/m'
'      'CF - opbrengsten' per type: blok 'Termijnen k - naam' (rij erboven '# won' en het aantal), bouwplanning (aantal
'                        woningen per termijn per kwartaal), blok 'Omzet HVG Termijnen' (Koopsom, termijnen %/€/€ per wn,
'                        extra's), blok 'Verkooptempo' (verkocht per kwartaal; 'actuals' = al verkocht)
' -------------------------------------------------------------------------------------
Public Sub UitFOOphalen()
    Dim pad As Variant, wbF As Workbook, wsCF As Worksheet, wsO As Worksheet, wsA As Worksheet
    Dim wsI As Worksheet, wsT As Worksheet, wsD As Worksheet
    Dim stap As String, bericht As String, nPer As Long, nTyp As Long, k As Long, waarschuwing As String
    Dim idxAct As Long, rijVan As Object, aJaar As Long, aKw As Long
    On Error GoTo Fout
    pad = Application.GetOpenFilename("Excel-werkboeken (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", , "FO-werkboek kiezen")
    If VarType(pad) = vbBoolean Then Exit Sub
    If MsgBox("Tab Invoer (cashflow, verkocht, transport) en tab Woningtypes worden overschreven met de gegevens uit:" & vbCrLf & pad & _
              vbCrLf & vbCrLf & "De knoppen op het Dashboard blijven staan. Doorgaan?", vbQuestion + vbYesNo, "Ophalen uit FO") <> vbYes Then Exit Sub
    Set wsI = ThisWorkbook.Worksheets(SH_INVOER)
    Set wsT = ThisWorkbook.Worksheets(SH_TYPES)
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)
    Application.ScreenUpdating = False
    stap = "FO openen"
    Set wbF = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True)
    Set wsCF = ZoekBlad(wbF, "1. Cashflow")
    If wsCF Is Nothing Then Set wsCF = ZoekBlad(wbF, "Cashflow")
    Set wsO = ZoekBlad(wbF, "CF - opbrengsten")
    Set wsA = ZoekBlad(wbF, "FO - actuals")
    If wsCF Is Nothing Or wsO Is Nothing Then
        wbF.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Dit lijkt geen FO-werkboek: tabblad '1. Cashflow' of 'CF - opbrengsten' ontbreekt.", vbExclamation, "Ophalen uit FO"
        Exit Sub
    End If

    stap = "cashflow (1. Cashflow)"
    Set rijVan = CreateObject("Scripting.Dictionary")
    nPer = FOCashflow(wsCF, wsI, wsD, rijVan, waarschuwing)
    If nPer = 0 Then Err.Raise 1000, , "Geen periodes gevonden onder de koprij 'Jaar' op '1. Cashflow'."

    stap = "actuals (FO - actuals)"
    idxAct = 0
    If Not wsA Is Nothing Then idxAct = FOActuals(wsA)
    If idxAct > 0 Then
        aJaar = (idxAct - 1) \ 4
        aKw = idxAct - aJaar * 4
        wsD.Range(CEL_ACTUALS_JAAR).Value = aJaar
        wsD.Range(CEL_ACTUALS_KW).Value = aKw
        wsD.Range(CEL_FO_ACTUALS).Value = "Q" & aKw & " " & aJaar
    Else
        wsD.Range(CEL_FO_ACTUALS).Value = ""
        idxAct = CLng(Val(wsD.Range(CEL_ACTUALS_JAAR).Value)) * 4 + CLng(Val(wsD.Range(CEL_ACTUALS_KW).Value))
    End If

    stap = "woningtypes (CF - opbrengsten)"
    For k = 1 To MAX_TYPES
        InvoerLeegmaken wsT, TYPE_COL1 + (k - 1) * TYPE_BREEDTE, TYPE_BREEDTE
        InvoerLeegmaken wsI, INVOER_COL1 + (k - 1) * INVOER_BREEDTE, INVOER_BREEDTE
    Next k
    nTyp = FOTypen(wsO, wsT, wsI, rijVan, idxAct, waarschuwing)

    stap = "afronden"
    wsD.Range(CEL_FO_BESTAND).Value = wbF.Name
    wsD.Range(CEL_FO_DATUM).Value = Format$(Now, "dd-mm-yyyy hh:nn")
    wbF.Close SaveChanges:=False
    Set wbF = Nothing
    Application.Calculate
    Application.ScreenUpdating = True
    bericht = "Klaar: " & nPer & " periodes op tab Invoer en " & nTyp & " woningtype(n) op tab Woningtypes uit " & CStr(pad) & "."
    If idxAct > 0 Then bericht = bericht & vbCrLf & "Actuals t/m Q" & aKw & " " & aJaar & " (uit 'FO - actuals')."
    bericht = bericht & vbCrLf & vbCrLf & "Controleer op het Dashboard de controle 'Invoer sluit aan op het FO' en de dekking van het verkooptempo-model."
    If Len(waarschuwing) > 0 Then bericht = bericht & vbCrLf & vbCrLf & "Let op:" & vbCrLf & waarschuwing
    MsgBox bericht, IIf(Len(waarschuwing) > 0, vbExclamation, vbInformation), "Ophalen uit FO"
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wbF Is Nothing Then wbF.Close SaveChanges:=False
    MsgBox "Het ging mis bij: " & stap & vbCrLf & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Ophalen uit FO"
End Sub

Private Function ZoekBlad(wb As Workbook, naam As String) As Worksheet
    ' werkblad op naam, hoofdletters en spaties genegeerd
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Replace(LCase$(ws.Name), " ", "") = Replace(LCase$(naam), " ", "") Then
            Set ZoekBlad = ws
            Exit Function
        End If
    Next ws
End Function

Private Function IsJaar(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    IsJaar = (Val(v) >= 1990 And Val(v) <= 2100)
End Function

Private Function KwNummer(v As Variant) As Long
    ' 'Q1'..'Q4' of 1..4 -> 1..4; anders 0
    Dim t As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    t = UCase$(Replace(CStr(v), " ", ""))
    If Left$(t, 1) = "Q" Then t = Mid$(t, 2)
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Then KwNummer = CLng(t)
End Function

Private Function Getal(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Getal = CDbl(v)
End Function

Private Function FOCashflow(wsCF As Worksheet, wsI As Worksheet, wsD As Worksheet, rijVan As Object, ByRef waarschuwing As String) As Long
    ' '1. Cashflow' -> Invoer B8:J..; rijVan(jaar*4+kw) = rij op Invoer; totalen -> Dashboard FO-cellen
    Dim kop As Long, r As Long, i As Long, c As Long, q As Long
    For r = 1 To 15
        If LCase$(Trim$(CStr(wsCF.Cells(r, 2).Value))) = "jaar" Then
            kop = r
            Exit For
        End If
    Next r
    If kop = 0 Then Err.Raise 1001, , "Koprij 'Jaar' niet gevonden op '1. Cashflow'."
    wsI.Range(wsI.Cells(INVOER_RIJ1, 2), wsI.Cells(INVOER_RIJN, 10)).ClearContents
    r = kop + 1
    i = 0
    Do While IsJaar(wsCF.Cells(r, 2).Value)
        If i >= MAX_PERIODES Then
            waarschuwing = waarschuwing & "- '1. Cashflow' heeft meer dan " & MAX_PERIODES & " periodes; alleen de eerste " & MAX_PERIODES & " zijn overgenomen." & vbCrLf
            Exit Do
        End If
        For c = 2 To 10
            If c = 3 Then
                q = KwNummer(wsCF.Cells(r, 3).Value)
                If q > 0 Then wsI.Cells(INVOER_RIJ1 + i, 3).Value = "Q" & q Else wsI.Cells(INVOER_RIJ1 + i, 3).ClearContents
            ElseIf IsNumeric(wsCF.Cells(r, c).Value) And Not IsEmpty(wsCF.Cells(r, c).Value) Then
                wsI.Cells(INVOER_RIJ1 + i, c).Value = CDbl(wsCF.Cells(r, c).Value)
            End If
        Next c
        If q > 0 Then rijVan(CLng(Val(wsCF.Cells(r, 2).Value)) * 4 + q) = INVOER_RIJ1 + i
        i = i + 1
        r = r + 1
    Loop
    ' totalen onder de tabel ('Totaal:' in kolom C)
    wsD.Range(CEL_FO_OPBR).ClearContents
    wsD.Range(CEL_FO_KOSTEN).ClearContents
    For c = r To r + 6
        If LCase$(Left$(Trim$(CStr(wsCF.Cells(c, 3).Value)), 6)) = "totaal" Then
            wsD.Range(CEL_FO_KOSTEN).Value = Getal(wsCF.Cells(c, 4).Value)
            wsD.Range(CEL_FO_OPBR).Value = Getal(wsCF.Cells(c, 6).Value)
            Exit For
        End If
    Next c
    FOCashflow = i
End Function

Private Function FOActuals(wsA As Worksheet) As Long
    ' laatste boekjaar x 4 + kwartaal met een boeking (kolommen op koptekst: Boekjaar, Q, Waarde/TrVal)
    Dim c As Long, cj As Long, cq As Long, cw As Long, r As Long, laatste As Long, idx As Long, kop As String, rMax As Long
    For c = 1 To wsA.UsedRange.Columns.Count + wsA.UsedRange.Column
        kop = LCase$(Trim$(CStr(wsA.Cells(1, c).Value)))
        If kop = "boekjaar" Then cj = c
        If kop = "q" Then cq = c
        If kop = "waarde/trval" Then cw = c
    Next c
    If cj = 0 Or cq = 0 Then Exit Function
    rMax = wsA.Cells(wsA.Rows.Count, cj).End(xlUp).Row
    For r = 2 To rMax
        If IsJaar(wsA.Cells(r, cj).Value) And KwNummer(wsA.Cells(r, cq).Value) > 0 Then
            If cw = 0 Or Not IsEmpty(wsA.Cells(r, cw).Value) Then
                idx = CLng(Val(wsA.Cells(r, cj).Value)) * 4 + KwNummer(wsA.Cells(r, cq).Value)
                If idx > laatste Then laatste = idx
            End If
        End If
    Next r
    FOActuals = laatste
End Function

Private Sub KolomKaart(ws As Worksheet, rijQ As Long, ByRef idxKol() As Long, ByRef cAct As Long)
    ' idxKol(c) = jaar*4+kw van kolom c (0 = geen kwartaalkolom), uit de rij met Q1..Q4 en de jaartallen in de rij erboven;
    ' cAct = kolom 'actuals' (0 = geen)
    Dim c As Long, jaar As Long, q As Long, nKol As Long
    nKol = ws.UsedRange.Columns.Count + ws.UsedRange.Column
    ReDim idxKol(1 To nKol)
    cAct = 0
    jaar = 0
    For c = 1 To nKol
        If IsJaar(ws.Cells(rijQ - 1, c).Value) Then jaar = CLng(Val(ws.Cells(rijQ - 1, c).Value))
        If LCase$(Left$(Trim$(CStr(ws.Cells(rijQ, c).Value)), 6)) = "actual" Then cAct = c
        q = KwNummer(ws.Cells(rijQ, c).Value)
        If q > 0 And jaar > 0 Then idxKol(c) = jaar * 4 + q Else idxKol(c) = 0
    Next c
End Sub

Private Function EersteIdx(ws As Worksheet, r As Long, idxKol() As Long) As Long
    ' kwartaalindex van de eerste kwartaalcel in rij r met een getal <> 0 (0 = geen)
    Dim c As Long
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            If Getal(ws.Cells(r, c).Value) <> 0 Then
                EersteIdx = idxKol(c)
                Exit Function
            End If
        End If
    Next c
End Function

Private Function ActWaarde(ws As Worksheet, r As Long, cAct As Long) As Double
    ' waarde in de kolom 'actuals' van rij r (0 als die kolom er niet is)
    If cAct > 0 Then ActWaarde = Getal(ws.Cells(r, cAct).Value)
End Function

Private Function SomRij(ws As Worksheet, r As Long, idxKol() As Long) As Double
    Dim c As Long
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then SomRij = SomRij + Getal(ws.Cells(r, c).Value)
    Next c
End Function

Private Function ZoekPlanningRij(ws As Worksheet, r1 As Long, r2 As Long, naam As String, idxKol() As Long) As Long
    ' rij r1..r2 in de bouwplanning met dezelfde naam en minstens één kwartaalcel; 0 = geen
    Dim r As Long
    For r = r1 To r2
        If LCase$(Trim$(CStr(ws.Cells(r, 2).Value))) = LCase$(Trim$(naam)) Then
            If SomRij(ws, r, idxKol) <> 0 Then
                ZoekPlanningRij = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub SchrijfVerdeeld(wsT As Worksheet, ws As Worksheet, rBron As Long, idxKol() As Long, idxStart As Long, naam As String, _
                            waarde As Double, col As Long, ByRef rij As Long, rMax As Long, nf As String, ByRef waarschuwing As String)
    ' schrijft naam / waarde / bouwkwartaal in rij 'rij' van het typeblok (kolommen col, col+1, col+2); staat de termijn in het FO
    ' over meer kwartalen verdeeld, dan evenredig gesplitst in meer rijen (naam (1/2), (2/2) ...)
    Dim c As Long, som As Double, n As Long, i As Long, deel As Double
    som = SomRij(ws, rBron, idxKol)
    If som = 0 Or idxStart = 0 Then
        If rij > rMax Then
            waarschuwing = waarschuwing & "- te veel termijnen voor '" & naam & "': niet alles overgenomen." & vbCrLf
            Exit Sub
        End If
        wsT.Cells(rij, col).Value = naam
        wsT.Cells(rij, col + 1).Value = waarde
        wsT.Cells(rij, col + 1).NumberFormat = nf
        rij = rij + 1
        Exit Sub
    End If
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            If Getal(ws.Cells(rBron, c).Value) <> 0 Then n = n + 1
        End If
    Next c
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            If Getal(ws.Cells(rBron, c).Value) <> 0 Then
                If rij > rMax Then
                    waarschuwing = waarschuwing & "- te veel termijnen voor '" & naam & "': niet alles overgenomen." & vbCrLf
                    Exit Sub
                End If
                i = i + 1
                deel = Getal(ws.Cells(rBron, c).Value) / som
                wsT.Cells(rij, col).Value = naam & IIf(n > 1, " (" & i & "/" & n & ")", "")
                wsT.Cells(rij, col + 1).Value = waarde * deel
                wsT.Cells(rij, col + 1).NumberFormat = nf
                wsT.Cells(rij, col + 2).Value = idxKol(c) - idxStart + 1
                rij = rij + 1
            End If
        End If
    Next c
End Sub

Private Sub SchrijfAantallen(ws As Worksheet, r As Long, idxKol() As Long, cAct As Long, wsI As Worksheet, colI As Long, rijVan As Object, idxAct As Long)
    ' kwartaalcellen (en de kolom 'actuals' in het actuals-kwartaal) van rij r naar de kolom colI op tab Invoer, opgeteld bij wat er staat
    Dim c As Long, v As Double, rI As Variant
    If cAct > 0 Then
        v = Getal(ws.Cells(r, cAct).Value)
        If v <> 0 Then
            rI = RijVoor(rijVan, idxAct)
            If Not IsEmpty(rI) Then wsI.Cells(rI, colI).Value = Getal(wsI.Cells(rI, colI).Value) + v
        End If
    End If
    For c = LBound(idxKol) To UBound(idxKol)
        If idxKol(c) > 0 Then
            v = Getal(ws.Cells(r, c).Value)
            If v <> 0 And rijVan.Exists(idxKol(c)) Then
                wsI.Cells(rijVan(idxKol(c)), colI).Value = Getal(wsI.Cells(rijVan(idxKol(c)), colI).Value) + v
            End If
        End If
    Next c
End Sub

Private Function RijVoor(rijVan As Object, idx As Long) As Variant
    ' rij op Invoer van kwartaalindex idx; bestaat die niet, de eerste periode met een kwartaal
    Dim sleutel As Variant, kleinste As Long
    If rijVan.Exists(idx) Then
        RijVoor = rijVan(idx)
        Exit Function
    End If
    kleinste = 0
    For Each sleutel In rijVan.Keys
        If kleinste = 0 Or CLng(sleutel) < kleinste Then kleinste = CLng(sleutel)
    Next sleutel
    If kleinste > 0 Then RijVoor = rijVan(kleinste) Else RijVoor = Empty
End Function

Private Function FOTypen(wsO As Worksheet, wsT As Worksheet, wsI As Worksheet, rijVan As Object, idxAct As Long, ByRef waarschuwing As String) As Long
    ' blokken op 'CF - opbrengsten' -> typeblokken op Woningtypes en verkocht/transport op Invoer
    Dim rMax As Long, r As Long, b As String, nBlok As Long, nOmzet As Long, rVerkoop As Long
    Dim blok() As Long, omzet() As Long, k As Long, kT As Long, col As Long, naam As String, aantal As Double
    Dim idxKol() As Long, cAct As Long, idxKolO() As Long, cActO As Long, idxStart As Long, rPlan1 As Long, rPlan2 As Long
    Dim ro As Long, lab As String, ll As String, pct As Variant, eur As Variant, pw As Variant, koopPw As Double, koopsom As Double
    Dim rTermijn As Long, rExtra As Long, rp As Long, rFee As Long, nComp As Long, feeComp As Long, isDaeb As Boolean
    Dim termNaam(1 To 20) As String, termPct(1 To 20) As Double, termPw(1 To 20) As Double, termRij(1 To 20) As Long, nTerm As Long, i As Long
    ReDim blok(1 To MAX_TYPES)
    ReDim omzet(1 To MAX_TYPES)
    rMax = wsO.Cells(wsO.Rows.Count, 2).End(xlUp).Row
    For r = 2 To rMax
        b = LCase$(Trim$(CStr(wsO.Cells(r, 2).Value)))
        If Left$(b, 10) = "termijnen " And InStr(b, " in ") = 0 And LCase$(Left$(Trim$(CStr(wsO.Cells(r - 1, 6).Value)), 5)) = "# won" Then
            If nBlok < MAX_TYPES Then
                nBlok = nBlok + 1
                blok(nBlok) = r
            End If
        ElseIf Left$(b, 9) = "omzet hvg" Or Left$(b, 15) = "omzet termijnen" Then
            If nOmzet < MAX_TYPES Then
                nOmzet = nOmzet + 1
                omzet(nOmzet) = r
            End If
        ElseIf b = "verkooptempo" And rVerkoop = 0 Then
            rVerkoop = r
        End If
    Next r
    kT = 0
    For k = 1 To nBlok
        aantal = Getal(wsO.Cells(blok(k) - 1, 7).Value)
        If aantal <= 0 Then GoTo Volgende
        kT = kT + 1
        col = TYPE_COL1 + (kT - 1) * TYPE_BREEDTE
        naam = Trim$(CStr(wsO.Cells(blok(k), 2).Value))
        If InStr(naam, " - ") > 0 Then naam = Trim$(Mid$(naam, InStr(naam, " - ") + 3)) Else naam = "Type " & kT
        wsT.Cells(TYPE_RIJ_NAAM, col).Value = naam
        wsT.Cells(TYPE_RIJ_AANTAL, col).Value = aantal
        ' bouwplanning: eerste rij met een telling = start bouw
        KolomKaart wsO, blok(k), idxKol, cAct
        rPlan1 = blok(k) + 1
        rPlan2 = rPlan1
        Do While Len(Trim$(CStr(wsO.Cells(rPlan2 + 1, 2).Value))) > 0
            rPlan2 = rPlan2 + 1
        Loop
        idxStart = 0
        For r = rPlan1 To rPlan2
            idxStart = EersteIdx(wsO, r, idxKol)
            If idxStart > 0 Then Exit For
        Next r
        If idxStart > 0 Then
            wsT.Cells(TYPE_RIJ_STARTJAAR, col).Value = (idxStart - 1) \ 4
            wsT.Cells(TYPE_RIJ_STARTKW, col).Value = idxStart - ((idxStart - 1) \ 4) * 4
        End If
        ' termijnnamen en extra-namen van het sjabloon leegmaken (het FO levert zijn eigen namen)
        wsT.Range(wsT.Cells(TYPE_RIJ_T1, col), wsT.Cells(TYPE_RIJ_TN, col + 2)).ClearContents
        wsT.Range(wsT.Cells(TYPE_RIJ_X1, col), wsT.Cells(TYPE_RIJ_XN, col + 2)).ClearContents
        If k > nOmzet Then GoTo Volgende
        ' omzetblok: koopsom, termijnen, extra's, fees
        ro = omzet(k)
        KolomKaart wsO, ro, idxKolO, cActO
        koopPw = 0
        nTerm = 0
        rTermijn = TYPE_RIJ_T1
        rExtra = TYPE_RIJ_X1
        isDaeb = False
        nComp = 0
        rFee = 0
        r = ro + 1
        Do While Len(Trim$(CStr(wsO.Cells(r, 2).Value))) > 0
            lab = Trim$(CStr(wsO.Cells(r, 2).Value))
            ll = LCase$(lab)
            pct = wsO.Cells(r, 3).Value
            eur = wsO.Cells(r, 4).Value
            pw = wsO.Cells(r, 5).Value
            If ll = "koopsom" Then
                koopPw = Getal(pw)
                SchrijfAantallen wsO, r, idxKolO, cActO, wsI, INVOER_COL1 + (kT - 1) * INVOER_BREEDTE + 1, rijVan, idxAct
            ElseIf Left$(ll, 6) = "ak fee" Or Left$(ll, 6) = "ak-fee" Or Left$(ll, 17) = "bijkomende kosten" Or Left$(ll, 10) = "onvoorzien" Then
                ' DAEB: componentrij (totaal in €, zonder kwartalen) of termijnrij (% of € in een kwartaal / actuals)
                If SomRij(wsO, r, idxKolO) = 0 And ActWaarde(wsO, r, cActO) = 0 And Getal(eur) <> 0 And (IsEmpty(pct) Or Not IsNumeric(pct)) Then
                    If nComp < 2 Then
                        nComp = nComp + 1
                        isDaeb = True
                        wsT.Cells(TYPE_RIJ_FC1 + nComp - 1, col).Value = lab
                        wsT.Cells(TYPE_RIJ_FC1 + nComp - 1, col + 1).Value = Getal(eur)
                        rFee = TYPE_RIJ_F1 + (nComp - 1) * 5
                    End If
                ElseIf nComp > 0 And rFee > 0 And rFee < TYPE_RIJ_F1 + nComp * 5 Then
                    wsT.Cells(rFee, col).Value = lab
                    If IsNumeric(pct) And Not IsEmpty(pct) Then wsT.Cells(rFee, col + 1).Value = CDbl(pct) Else wsT.Cells(rFee, col + 1).Value = Getal(eur)
                    If EersteIdx(wsO, r, idxKolO) > 0 Then
                        wsT.Cells(rFee, col + 2).Value = ((EersteIdx(wsO, r, idxKolO) - 1) \ 4) & " Q" & (EersteIdx(wsO, r, idxKolO) - ((EersteIdx(wsO, r, idxKolO) - 1) \ 4) * 4)
                    ElseIf cActO > 0 Then
                        If Getal(wsO.Cells(r, cActO).Value) <> 0 Then wsT.Cells(rFee, col + 2).Value = "actuals"
                    End If
                    rFee = rFee + 1
                End If
            ElseIf IsNumeric(pct) And Not IsEmpty(pct) Then
                If CDbl(pct) <> 0 And nTerm < 20 Then
                    nTerm = nTerm + 1
                    termNaam(nTerm) = lab
                    termPct(nTerm) = CDbl(pct)
                    termPw(nTerm) = Getal(pw)
                    termRij(nTerm) = r
                End If
            ElseIf Getal(pw) <> 0 Then
                ' extra per woning (kopersmeerwerk, kadastrale kosten, overige): bij transport als het FO de telling in 'actuals' zet
                rp = ZoekPlanningRij(wsO, rPlan1, rPlan2, lab, idxKol)
                If rp = 0 And SomRij(wsO, r, idxKolO) <> 0 Then rp = r
                If rp = 0 Or ActWaarde(wsO, r, cActO) <> 0 Then
                    If rExtra <= TYPE_RIJ_XN Then
                        wsT.Cells(rExtra, col).Value = lab
                        wsT.Cells(rExtra, col + 1).Value = Getal(pw)
                        rExtra = rExtra + 1
                    End If
                ElseIf rp = r Then
                    SchrijfVerdeeld wsT, wsO, rp, idxKolO, idxStart, lab, Getal(pw), col, rExtra, TYPE_RIJ_XN, "#,##0", waarschuwing
                Else
                    SchrijfVerdeeld wsT, wsO, rp, idxKol, idxStart, lab, Getal(pw), col, rExtra, TYPE_RIJ_XN, "#,##0", waarschuwing
                End If
            End If
            r = r + 1
        Loop
        If isDaeb Then
            wsT.Cells(TYPE_RIJ_SOORT, col).Value = "DAEB"
            wsT.Cells(TYPE_RIJ_KOOPSOM, col).ClearContents
            wsT.Cells(TYPE_RIJ_GROND, col + 1).ClearContents
        Else
            koopsom = koopPw
            For i = 1 To nTerm
                koopsom = koopsom + termPw(i)
            Next i
            If koopsom > 0 Then
                wsT.Cells(TYPE_RIJ_KOOPSOM, col).Value = Round(koopsom, 2)
                wsT.Cells(TYPE_RIJ_GROND, col + 1).Value = koopPw / koopsom
            End If
            For i = 1 To nTerm
                rp = ZoekPlanningRij(wsO, rPlan1, rPlan2, termNaam(i), idxKol)
                If rp > 0 Then
                    SchrijfVerdeeld wsT, wsO, rp, idxKol, idxStart, termNaam(i), termPct(i), col, rTermijn, TYPE_RIJ_TN, "0.0%", waarschuwing
                Else
                    SchrijfVerdeeld wsT, wsO, termRij(i), idxKolO, idxStart, termNaam(i), termPct(i), col, rTermijn, TYPE_RIJ_TN, "0.0%", waarschuwing
                End If
            Next i
        End If
Volgende:
    Next k
    ' verkooptempo: rijen in dezelfde volgorde als de blokken (ook blokken zonder woningen tellen als rij)
    If rVerkoop > 0 Then
        KolomKaart wsO, rVerkoop, idxKol, cAct
        kT = 0
        r = rVerkoop + 1
        k = 0
        Do While Len(Trim$(CStr(wsO.Cells(r, 2).Value))) > 0 And k < nBlok
            k = k + 1
            If Getal(wsO.Cells(blok(k) - 1, 7).Value) > 0 Then
                kT = kT + 1
                SchrijfAantallen wsO, r, idxKol, cAct, wsI, INVOER_COL1 + (kT - 1) * INVOER_BREEDTE, rijVan, idxAct
            End If
            r = r + 1
        Loop
    Else
        waarschuwing = waarschuwing & "- blok 'Verkooptempo' niet gevonden op 'CF - opbrengsten': verkocht per kwartaal is leeg gebleven." & vbCrLf
    End If
    FOTypen = kT
End Function

' -------------------------------------------------------------------------------------
'  5. Sjabloon controleren: staan alle vormnamen erin?
' -------------------------------------------------------------------------------------
Public Sub ControleerSjabloon()
    ' vormnamen uit kolom F van de KPI-regels (o.a. BT_TITEL, BT_SUBTITEL, VP_TITEL, VP_SUBTITEL, VO_TITEL, CF_*, SC_*,
    ' KPI*), uit Grafieken() en uit Tabellen(); daarnaast de hulpvormen van dia 5 die SchikDia5 gebruikt
    Dim pp As Object, pres As Object, wsP As Worksheet
    Dim r As Long, vorm As String, mist As String, aantal As Long, i As Long, lijst As Variant, hulp As String
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
    ' hulpvormen van dia 5 (kaarten): zonder die vormen vult de knop wel, maar schikt SchikDia5 de dia niet
    lijst = Array("VT_KAART", "VT_TABELKAART")
    For i = LBound(lijst) To UBound(lijst)
        If ZoekVorm(pres, CStr(lijst(i))) Is Nothing Then hulp = hulp & CStr(lijst(i)) & vbCrLf
    Next i
    On Error Resume Next
    pres.Saved = -1
    pres.Close
    On Error GoTo 0
    If Len(hulp) > 0 Then
        hulp = vbCrLf & "Hulpvormen van dia 5 die niet gevonden zijn (de tabel wordt dan niet automatisch geschikt):" & vbCrLf & hulp
    End If
    If Len(mist) = 0 Then
        MsgBox "Alles compleet: alle " & aantal & " vormnamen staan in het sjabloon." & hulp, _
               IIf(Len(hulp) > 0, vbExclamation, vbInformation), "Sjabloon controleren"
    Else
        MsgBox "Deze vormen ontbreken in het sjabloon (die slaat de knop over):" & vbCrLf & vbCrLf & mist & vbCrLf & _
               "Hernoem de vormen in PowerPoint via Start > Selecteren > Selectiedeelvenster." & vbCrLf & hulp, _
               vbExclamation, "Sjabloon controleren"
    End If
    Exit Sub
Fout:
    MsgBox "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Sjabloon controleren"
End Sub
