Attribute VB_Name = "Module1"
Public ClientAModifierID As String
Public NomUtilisateurConnecte As String



Sub basededonnées()
Attribute basededonnées.VB_ProcData.VB_Invoke_Func = " \n14"
'
' basededonnées Macro
'

'
    Selection.Characters.Text = "base de données Clients"
    With Selection.Characters(Start:=1, Length:=23).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes("Button 7").ScaleWidth 1.2330097087, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Button 7").IncrementLeft -7.2
    ActiveSheet.Shapes("Button 7").IncrementTop -5.4
    ActiveSheet.Shapes("Button 7").ScaleWidth 1.188976378, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 7")).Select
    ActiveSheet.Shapes("Button 7").IncrementLeft 3
    ActiveSheet.Shapes("Button 7").IncrementTop -0.6
    ActiveSheet.Buttons.Add(142.8, 324, 94.2, 42).Select
    Selection.OnAction = "AjoutC"
    Selection.Characters.Text = "Ajouter nouveaux clients"
    With Selection.Characters(Start:=1, Length:=24).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Buttons.Add(268.2, 325.2, 96.6, 41.4).Select
    Selection.OnAction = "RecherchC"
    Selection.Characters.Text = "Rechercher un client"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Buttons.Add(397.8, 326.4, 93, 40.2).Select
    Selection.OnAction = "SupprimC"
    Selection.Characters.Text = "Supprimer un client"
    With Selection.Characters(Start:=1, Length:=19).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Buttons.Add(529.8, 327, 85.8, 39.6).Select
    Selection.OnAction = "ModifC"
    Selection.Characters.Text = "Modifier les données d'un client"
    With Selection.Characters(Start:=1, Length:=32).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes("Button 11").ScaleWidth 1.2867132867, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 11")).Select
    Selection.Characters.Text = "Modifier les données d'un client"
    With Selection.Characters(Start:=1, Length:=32).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 11")).Select
    ActiveSheet.Shapes("Button 11").IncrementLeft 59.4
    ActiveSheet.Shapes("Button 11").IncrementTop 10.8
    ActiveSheet.Shapes.Range(Array("Button 7")).Select
    Sheets("ESPACE CLIENT").Select
    ActiveSheet.Shapes("Button 7").IncrementLeft 17.4
    ActiveSheet.Shapes("Button 7").IncrementTop 63.6
    ActiveSheet.Shapes.Range(Array("Half Frame 16")).Select
    Selection.ShapeRange.IncrementLeft 18
    Selection.ShapeRange.IncrementTop 60.6
    Selection.ShapeRange.ScaleWidth 0.9911504425, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.7666666667, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 0.8125, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.IncrementLeft 1.8
    Selection.ShapeRange.IncrementTop 4.8
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 7")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    ActiveSheet.Shapes("Button 8").IncrementLeft 7.8
    ActiveSheet.Shapes("Button 8").IncrementTop 63
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 9")).Select
    ActiveSheet.Shapes("Button 9").IncrementLeft -2.4
    ActiveSheet.Shapes("Button 9").IncrementTop 62.4
    Sheets("ESPACE CLIENT").Select
    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    ActiveSheet.Shapes("Button 10").IncrementLeft -13.8
    ActiveSheet.Shapes("Button 10").IncrementTop 61.2
    ActiveSheet.Shapes.Range(Array("Button 11")).Select
    ActiveSheet.Shapes("Button 11").IncrementLeft -64.2
    ActiveSheet.Shapes("Button 11").IncrementTop 51.6
    ActiveSheet.Shapes.Range(Array("Half Frame 17")).Select
    Selection.ShapeRange.IncrementLeft -64.2
    Selection.ShapeRange.IncrementTop 54.6
    Selection.ShapeRange.ScaleWidth 0.7946859903, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.9361111111, msoFalse, _
        msoScaleFromBottomRight
    Selection.ShapeRange.ScaleWidth 0.9604879177, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.7596439169, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Half Frame 17")).Select
    Selection.ShapeRange.ScaleWidth 0.9873434544, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleWidth 0.961535226, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Half Frame 17")).Select
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    Sheets("ESPACE CLIENT").Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.OnAction = "Ajouterbtn"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    ActiveWindow.SmallScroll Down:=0
    Sheets("ACCUEIL").Select
    ActiveWindow.Zoom = 25
    ActiveWindow.Zoom = 40
    ActiveWindow.SmallScroll Down:=0
    ActiveSheet.Shapes.Range(Array("Picture 144")).Select
    ActiveSheet.Shapes.Range(Array("Picture 143")).Select
    ActiveSheet.Shapes.Range(Array("Picture 142")).Select
    ActiveSheet.Shapes.Range(Array("Picture 145")).Select
    Sheets("ESPACE CLIENT").Select
    ActiveWorkbook.Save
    Selection.Characters.Text = "Ajouter nouveaux clients"
    With Selection.Characters(Start:=1, Length:=24).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.OnAction = "OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Application.GoTo Reference:="OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.OnAction = "OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.OnAction = "OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.OnAction = "OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 9")).Select
    Selection.OnAction = "OuvrirFormulaireClient"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveSheet.Shapes.Range(Array("Button 9")).Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    Range("M33").Select
    ActiveWindow.SmallScroll Down:=-16
    ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Range("N20").Select
    ActiveWindow.SmallScroll Down:=5
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Select
    Selection.ShapeRange.IncrementLeft 9.6
    Selection.ShapeRange.IncrementTop 184.8
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset58
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ActiveWindow.SmallScroll Down:=6
    Sheets("GESTION ADMINISTRATIVE").Select
    ActiveSheet.Shapes.Range(Array("Pentagon 9")).Select
    ActiveSheet.Shapes.Range(Array("Picture 14")).Select
    ActiveSheet.Shapes.Range(Array("Pentagon 9")).Select
    ActiveSheet.Shapes.Range(Array("Picture 14")).Select
    Sheets("GESTION DES TERRAINS").Select
    ActiveWindow.SmallScroll Down:=9
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 9")).Select
    Selection.ShapeRange.IncrementLeft 8.4
    Selection.ShapeRange.IncrementTop 69.6
    Sheets("GESTION DU BUDGET").Select
    ActiveWorkbook.Save
    Application.WindowState = xlNormal
    Sheets("ACCUEIL").Select
    ActiveSheet.Shapes.Range(Array("TextBox 139")).Select
    ActiveSheet.Shapes.Range(Array("Picture 4")).Select
    Range("AF23").Select
    ActiveSheet.Shapes.Range(Array("Picture 4")).Select
    ActiveWindow.SmallScroll Down:=14
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 14")).Select
    Selection.ShapeRange.IncrementLeft 36
    Selection.ShapeRange.IncrementTop -19.5
    Range("AI58").Select
    ActiveWindow.SmallScroll Down:=-14
    Sheets("ACCUEIL").Select
    ActiveSheet.Shapes.Range(Array("TextBox 139")).Select
    Sheets("ACCUEIL").Select
    Selection.OnAction = "GestionAdmin"
    ActiveSheet.Shapes.Range(Array("Picture 4")).Select
    ActiveSheet.Shapes.Range(Array("TextBox 139")).Select
    Selection.OnAction = "GestionAdmin"
    ActiveSheet.Shapes.Range(Array("Picture 4")).Select
    ActiveSheet.Shapes.Range(Array("TextBox 139")).Select
    Selection.OnAction = "GestionAdmin"
    Sheets("ESPACE CLIENT").Select
    ActiveSheet.Shapes.Range(Array("Button 7")).Select
    Selection.OnAction = "afficherBdDClients"
    ActiveSheet.Shapes.Range(Array("Flowchart: Document 21")).Select
    ChDir _
        "C:\Users\HP\OneDrive\Bureau\ensea\SEMESTRE 1\LOGICIEL EXCEL ET VBA\PROJET VBA 2025"
End Sub

