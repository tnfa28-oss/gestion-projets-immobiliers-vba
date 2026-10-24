Attribute VB_Name = "Module3"
Sub CreerModeleDevisAvecLogo()

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim cheminModele As String
    Dim cheminLogo As String

    ' Définir les chemins
    cheminModele = ThisWorkbook.Path & "\Modele_Devis.docx"
    cheminLogo = ThisWorkbook.Path & "\logo.png" ' ? adapte ici si le nom de ton logo est différent

    ' Lancer Word
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add

    ' Insérer le logo en haut
    If Dir(cheminLogo) <> "" Then
        wordDoc.Shapes.AddPicture Filename:=cheminLogo, _
            LinkToFile:=False, SaveWithDocument:=True, _
            Left:=20, Top:=20, Width:=80, Height:=80
    Else
        MsgBox "Le fichier logo.png n'a pas été trouvé dans le dossier du classeur Excel.", vbExclamation
    End If

    ' Ajouter le contenu texte
    With wordDoc.Content
        .InsertParagraphAfter
        .InsertAfter vbCrLf & " DEVIS DE DÉPENSES – Projet <ID-Projet>" & vbCrLf & vbCrLf & _
                     " Client : <NomClient>" & vbCrLf & _
                    "Responsable : <Responsable>" & vbCrLf & _
                     " Email : <Email>" & vbCrLf & vbCrLf & _
                     " Coût d'achat du terrain : <AchatTerrain> FCFA" & vbCrLf & _
                     " Travaux de construction : <Travaux> FCFA" & vbCrLf & _
                     " Finitions : <Finitions> FCFA" & vbCrLf & _
                     " Prestation de l'entreprise : <Prestation> FCFA" & vbCrLf & vbCrLf & _
                     " Budget Total : <Total> FCFA" & vbCrLf & vbCrLf & _
                     " Date : <Date>"
    End With

    ' Enregistrer et fermer
    wordDoc.SaveAs2 Filename:=cheminModele
    wordDoc.Close False
    wordApp.Quit

    MsgBox "Modèle de devis avec logo créé avec succès.", vbInformation

End Sub

