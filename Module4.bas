Attribute VB_Name = "Module4"
Sub Bouton4_Cliquer()
 Call MajFeuillePhase(Sheets("Gros oeuvre"), "Gros oeuvre")

End Sub
Sub Bouton2_Cliquer()
Call MajFeuillePhase(Sheets("Second oeuvre"), "Second oeuvre")

End Sub
Sub Finitions_Bouton2_Cliquer()
Call MajFeuillePhase(Sheets("Finitions"), "Finitions")

End Sub

' Module : Module_GanttComplet


Sub GenererGantt(ws As Worksheet, phase As String)
    Dim projetChoisi As String
    projetChoisi = Trim(ws.Range("C7").Value)
    If projetChoisi = "" Then
        MsgBox "Veuillez sélectionner un ID projet dans la cellule C7.", vbExclamation
        Exit Sub
    End If

    Dim filtreResp As String, filtreAv As Double
    filtreResp = Trim(ws.Range("C8").Value)
    If IsNumeric(ws.Range("C9").Value) Then filtreAv = CDbl(ws.Range("C9").Value) Else filtreAv = 0

    ws.Range("C12:Z1000").ClearContents
    ws.Range("L12:L20").ClearContents ' Effacer ancien résumé
    On Error Resume Next
    ws.ChartObjects("DiagrammeGantt").Delete
    On Error GoTo 0

    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Sheets("BdD Tâches")
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, ligne As Long: ligne = 12
    Dim dateMin As Date: dateMin = Date
    Dim trouvee As Boolean: trouvee = False

    For i = 2 To lastRow
        If NettoyerTexte(wsData.Cells(i, 1).Value) = NettoyerTexte(projetChoisi) And _
           RemplacerOe(LCase(Trim(wsData.Cells(i, 2).Value))) = LCase(phase) And _
           (filtreResp = "" Or wsData.Cells(i, 7).Value = filtreResp) And _
           (Not IsEmpty(wsData.Cells(i, 6).Value) And wsData.Cells(i, 6).Value >= filtreAv) Then

            trouvee = True
            ws.Cells(ligne, 3).Resize(1, 6).Value = wsData.Cells(i, 3).Resize(1, 6).Value ' Tâche à Email
            If IsDate(ws.Cells(ligne, 4).Value) Then If ws.Cells(ligne, 4).Value < dateMin Then dateMin = ws.Cells(ligne, 4).Value
            ligne = ligne + 1
        End If
    Next i

    If Not trouvee Then
        MsgBox "Aucune tâche trouvée pour ce projet et ces filtres.", vbInformation
        Exit Sub
    End If

    ' Création du tableau structuré
    Dim tblRange As Range, derCol As Long: derCol = 8
    Set tblRange = ws.Range(ws.Cells(12, 3), ws.Cells(ligne - 1, derCol))

    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("TableauTaches"): If Not lo Is Nothing Then lo.Delete
    On Error GoTo 0
    Set lo = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    lo.Name = "TableauTaches": lo.TableStyle = "TableStyleMedium9"

    ' Mise en couleur des lignes selon avancement
    Dim cell As Range
    For Each cell In lo.ListColumns(4).DataBodyRange ' Avancement = colonne 4
        Select Case cell.Value
            Case Is < 30: cell.EntireRow.Interior.Color = RGB(255, 204, 204)
Case 30 To 70: cell.EntireRow.Interior.Color = RGB(255, 255, 204)
            Case 71 To 99: cell.EntireRow.Interior.Color = RGB(204, 255, 204)
            Case 100: cell.EntireRow.Interior.Color = RGB(0, 176, 80)
        End Select
    Next cell

    ' Zone de résumé
    Dim nbTotal As Long, nbFini As Long, sommeAv As Double, dureeTot As Long
    nbTotal = 0: nbFini = 0: sommeAv = 0: dureeTot = 0
    For i = 12 To ligne - 1
        If IsNumeric(ws.Cells(i, 6).Value) Then
            nbTotal = nbTotal + 1
            If ws.Cells(i, 6).Value >= 100 Then nbFini = nbFini + 1
            sommeAv = sommeAv + ws.Cells(i, 6).Value
        End If
        If IsDate(ws.Cells(i, 4).Value) And IsDate(ws.Cells(i, 5).Value) Then
            dureeTot = dureeTot + DateDiff("d", ws.Cells(i, 4).Value, ws.Cells(i, 5).Value)
        End If
    Next i
    If nbTotal > 0 Then
        ws.Range("L12").Value = "¦ Résumé du projet : " & projetChoisi
        ws.Range("L13").Value = "Nombre de tâches : " & nbTotal
        ws.Range("L14").Value = "Tâches terminées : " & nbFini
        ws.Range("L15").Value = "Avancement moyen : " & Format(sommeAv / nbTotal, "0.0") & "%"
        ws.Range("L16").Value = "Durée totale (jours) : " & dureeTot
    End If

    ' Création de la feuille temporaire pour le diagramme
    Dim wsTemp As Worksheet
    On Error Resume Next: Application.DisplayAlerts = False: Worksheets("TempGantt").Delete
    Application.DisplayAlerts = True: On Error GoTo 0
    Set wsTemp = Worksheets.Add: wsTemp.Name = "TempoGantt"
 Dim j As Long: j = 2
    wsTemp.Range("A1:D1").Value = Array("Tâche", "Décalage", "Avancement", "Reste")

    Dim tache As String, dateDebut As Date, dateFin As Date
    Dim duree As Long, avancement As Double, reste As Long, decalage As Long

    For i = 12 To ligne - 1
        tache = ws.Cells(i, 3).Value
        If IsDate(ws.Cells(i, 4).Value) And IsDate(ws.Cells(i, 5).Value) Then
            dateDebut = ws.Cells(i, 4).Value: dateFin = ws.Cells(i, 5).Value
            duree = DateDiff("d", dateDebut, dateFin): decalage = DateDiff("d", dateMin, dateDebut)
        Else: duree = 0: decalage = 0
        End If

        If IsNumeric(ws.Cells(i, 6).Value) Then avancement = CDbl(ws.Cells(i, 6).Value) Else avancement = 0
        If avancement > 100 Then avancement = 100 Else If avancement < 0 Then avancement = 0
        reste = duree - Round(duree * avancement / 100): If reste < 0 Then reste = 0

        wsTemp.Cells(j, 1).Value = tache
        wsTemp.Cells(j, 2).Value = decalage
        wsTemp.Cells(j, 3).Value = duree - reste
        wsTemp.Cells(j, 4).Value = reste
        j = j + 1
    Next i

    ' Création du diagramme Gantt
    Dim chartObj As ChartObject: Set chartObj = ws.ChartObjects.Add(Left:=350, Top:=80, Width:=700, Height:=300)
    chartObj.Name = "DiagrammeGantt"
    With chartObj.Chart
        .ChartType = xlBarStacked
        .SetSourceData Source:=wsTemp.Range("A1:D" & j - 1)
        .HasTitle = True: .ChartTitle.Text = "Diagramme Gantt - Phase " & phase
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Tâches"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Temps"
        .Axes(xlValue).MajorUnit = 7
        .Axes(xlValue).TickLabels.NumberFormat = "dd mmm yyyy"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(144, 238, 144)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(211, 211, 211)
        .Legend.Position = xlLegendPositionBottom
    End With

    Application.DisplayAlerts = False: wsTemp.Delete: Application.DisplayAlerts = True
End Sub

Function NettoyerTexte(txt As String) As String
    Dim i As Long, c As String, res As String
    For i = 1 To Len(txt)
        c = Mid(txt, i, 1)
        If Asc(c) >= 32 And Asc(c) <= 126 Then res = res & c
    Next i
    NettoyerTexte = Trim(res)
End Function

Function RemplacerOe(txt As String) As String
    RemplacerOe = Replace(txt, "œ", "oe")
End Function

Sub Bouton3_Cliquer()
Call GenererGantt(Sheets("Finitions"), "Finitions")
End Sub
