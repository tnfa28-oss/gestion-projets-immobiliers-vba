Attribute VB_Name = "ModFonctions"
Public Sub OuvrirModifierClient()
 FormModifierClient.Show
End Sub

Sub OuvrirModifierBudget()
 ModifierBudgets.Show
End Sub

Sub OuvrirModifierCollabo()
 ModifierCollabo.Show
End Sub

Sub OuvrirModifierDemande()
 ModifierDemande.Show
End Sub

Sub OuvrirModifierProjet()
 ModifierProjets.Show
End Sub

Sub OuvrirModifierTache()
 ModifierTache.Show
End Sub

Sub OuvrirModifierTerrainsPossession()
 ModifierTerrainsPossession.Show
End Sub

Sub OuvrirModifierTerrainsRecherche()
 ModifierTerrainsRecherches.Show
End Sub
Sub OuvrirListeClients()
 ListeClients.Show
End Sub

Sub OuvrirListeDemande()
 ListeDemande.Show
End Sub

Sub OuvrirListeProjets()
 ListeProjets.Show
End Sub

Sub OuvrirListeTaches()
 ListeTache.Show
End Sub

Sub OuvrirListeTerrainsPossession()
 ListeTerrainsPossession.Show
End Sub

Sub OuvrirListeTerrainsRecherche()
 ListeTerrainsRecherche.Show
End Sub
Sub OuvrirListeCollabo()
 ListeCollabo.Show
End Sub
Sub OuvrirListeBudget()
 ListeBudgets.Show
End Sub


Sub OuvrirFormulaireClient()
 formClients.Show
End Sub

Sub OuvrirFormulaireBudget()
 FormBudget.Show
End Sub

Sub OuvrirFormulaireDemande()
 FormDemande.Show
End Sub

Sub OuvrirFormulaireTerrain()
 FormlTerrains.Show
End Sub

Sub OuvrirFormulaireProjets()
 FormProjets.Show
End Sub
Sub OuvrirFormulaireCollabo()
 FormCollabo.Show
End Sub
Sub GestionAdmin()
   Sheets("GESTION ADMINISTRATIVE").Activate
End Sub

Sub afficherBdDClients()
   Sheets("BdD Clients").Activate
End Sub

Sub afficherCollaborateur()
   Sheets("Collaborateurs").Activate
End Sub

Sub afficherEspaceClient()
 Sheets("ESPACE CLIENT").Activate
End Sub

Sub afficherGestionTerrains()
 Sheets("GESTION DES TERRAINS").Activate
End Sub

Sub afficherSuiviPermis()
 Sheets("Suivi Permis").Activate
End Sub

Sub afficherGestionBudget()
 Sheets("GESTION DU BUDGET").Activate
End Sub

Sub afficherDemandePermis()
 Sheets("DEMANDE DE PERMIS").Activate
End Sub

Sub afficherSuiViProjets()
 Sheets("SUIVI T. PROJETS").Activate
End Sub

Sub afficherFonctionnalites()
 Sheets("FONCTIONNALITES COLLABO").Activate
End Sub

Sub afficherSyntheR()
 Sheets("SYNTHESE ET RAPPORT").Activate
End Sub

Sub afficherAccueil()
 Sheets("ACCUEIL").Activate
End Sub
Sub afficherBdDBudget()
 Sheets("BdD Budget").Activate
End Sub
Sub afficherBdDTerrainsRecherche()
 Sheets("BdD TerrainsRecherche").Activate
End Sub
Sub afficherBdDTerrainsPossession()
 Sheets("BdD TerrainsPossession").Activate
End Sub
Sub afficherBdDRapports()
 Sheets("BdD Rapports").Activate
End Sub
Sub afficherFinitions()
 Sheets("Finitions").Activate
End Sub
Sub afficherGrosoeuvre()
 Sheets("Gros oeuvre").Activate
End Sub
Sub afficherSecondoeuvre()
 Sheets("Second oeuvre").Activate
End Sub
Sub afficherJA()
 Sheets("Journal des activities").Activate
End Sub
Sub afficherDashboard()
 Sheets("Dashboard").Activate
End Sub
Sub afficherBdDProjets()
 Sheets("BdD Projets").Activate
End Sub
Sub afficherListeCollabo()
 ListeCollabo.Show
End Sub
Sub afficherBdDPermis()
 Sheets("feuil Permis").Activate
End Sub
Public Function AjouterAuJournal(nomUtilisateur As String, action As String, element As String, commentaire As String)

    Dim feuilleJournal As Worksheet
    Dim ligne As Long

    Set feuilleJournal = ThisWorkbook.Sheets("Journal des Activities")
    ligne = feuilleJournal.Cells(feuilleJournal.Rows.Count, 1).End(xlUp).Row + 1

    With feuilleJournal
        .Cells(ligne, 1).Value = Now                            ' Colonne A : Date + Heure
        .Cells(ligne, 2).Value = nomUtilisateur                 ' Colonne B : Utilisateur
        .Cells(ligne, 3).Value = action                         ' Colonne C : Action
        .Cells(ligne, 4).Value = element                        ' Colonne D : Élément
        .Cells(ligne, 5).Value = commentaire                    ' Colonne E : Commentaire
    End With

End Function

Public Function UtilisateurExiste(nom As String, email As String, mdp As String) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Set ws = Worksheets("Utilisateurs")

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LCase(Trim(ws.Cells(i, 1).Value)) = LCase(Trim(nom)) And _
       LCase(Trim(ws.Cells(i, 2).Value)) = LCase(Trim(email)) And _
       Trim(ws.Cells(i, 3).Value) = Trim(mdp) Then
       
        UtilisateurExiste = True
        Exit Function
    End If
Next i
UtilisateurExiste = False
End Function

Public Function EstEmailValide(email As String) As Boolean
    Dim aro As Long, point As Long
    email = Trim(email)
    aro = InStr(1, email, "@")
    point = InStrRev(email, ".")
    If aro > 1 And point > aro + 1 And point < Len(email) Then
        EstEmailValide = True
    Else
        EstEmailValide = False
    End If
End Function

Sub OuvrirEspaceConexion()
 EspaceConnexion.Show
End Sub

Private Sub btnExporterDevis()

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim cheminModele As String
    Dim cheminFinal As String
    
    ' Lien vers ton modèle Word
    cheminModele = ThisWorkbook.Path & "\Modele_Devis.docx"
    cheminFinal = ThisWorkbook.Path & "\Devis_" & txtNomClient.Value & ".docx"
    
    ' Ouvrir Word en arrière-plan
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(cheminModele)
    
    ' Remplacement des balises dans le modèle
    With wordDoc.Content.Find
        .Execute FindText:="<ID-Projet>", ReplaceWith:="PRJ-" & Format(Now, "yyyymmddhhmmss"), Replace:=2
        .Execute FindText:="<NomClient>", ReplaceWith:=txtNomClient.Value, Replace:=2
        .Execute FindText:="<Responsable>", ReplaceWith:=txtResponsable.Value, Replace:=2
        .Execute FindText:="<Email>", ReplaceWith:=txtEmail.Value, Replace:=2
        .Execute FindText:="<AchatTerrain>", ReplaceWith:=txtTotalAchat.Value, Replace:=2
        .Execute FindText:="<Travaux>", ReplaceWith:=txtTotalTravaux.Value, Replace:=2
        .Execute FindText:="<Finitions>", ReplaceWith:=txtTotalFinitions.Value, Replace:=2
        .Execute FindText:="<Prestation>", ReplaceWith:=txtTotalPrestation.Value, Replace:=2
        .Execute FindText:="<Total>", ReplaceWith:=txtBudgetTotal.Value, Replace:=2
        .Execute FindText:="<Date>", ReplaceWith:=Format(Date, "dd/mm/yyyy"), Replace:=2
    End With

    ' Enregistrer sous nouveau nom
    wordDoc.SaveAs2 Filename:=cheminFinal
    
    ' Imprimer directement (optionnel)
    wordDoc.PrintOut
    
    ' Fermer le document
    wordDoc.Close False
    wordApp.Quit
    
    MsgBox "Le devis a été généré et imprimé avec succès.", vbInformation
    With ThisWorkbook.Sheets("BdD Rapports")
    Dim ligneRapport As Long
    ligneRapport = .Cells(.Rows.Count, 1).End(xlUp).Row + 1

    .Cells(ligneRapport, 1).Value = Now
    .Cells(ligneRapport, 2).Value = NomUtilisateurConnecte
    .Cells(ligneRapport, 3).Value = "Fiche Permis"
    .Cells(ligneRapport, 4).Value = "Permis_" & IDPermis & "_" & nomClient & ".pdf"
    .Cells(ligneRapport, 5).Value = cheminPDF
End With
End Sub

Sub OuvrirFormulaireTache()
 FormTache.Show
End Sub

Sub TestFonction()
    MsgBox UtilisateurExiste("Alice", "alice@email.com", "motdepasse")
End Sub
Sub ExporterFichePermisPDF()
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Suivi Permis")

    Dim IDPermis As String
    IDPermis = InputBox("Entrez l'ID du permis à exporter :", "Exporter Fiche Permis")

    If IDPermis = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long, trouve As Boolean
    trouve = False

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = IDPermis Then
            trouve = True
            Exit For
        End If
    Next i

    If Not trouve Then
        MsgBox "ID Permis non trouvé.", vbExclamation
        Exit Sub
    End If

    ' Créer le document Word
    Dim wordApp As Object, doc As Object
    Set wordApp = CreateObject("Word.Application")
    Set doc = wordApp.Documents.Add

    ' Récupérer les données importantes
    Dim nomClient As String: nomClient = ws.Cells(i, 2).Value
    Dim dateMAJ As String: dateMAJ = ws.Cells(i, 19).Value
 ' Contenu du document
    With doc.Content
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Text = "FICHE PERMIS DE CONSTRUIRE" & vbCrLf & String(35, "=") & vbCrLf & vbCrLf & _
                "ID Permis : " & ws.Cells(i, 1).Value & vbCrLf & _
                "Nom du client : " & nomClient & vbCrLf & _
                "Email du client : " & ws.Cells(i, 3).Value & vbCrLf & _
                "Nom de l'architecte : " & ws.Cells(i, 4).Value & vbCrLf & _
                "Email de l'architecte : " & ws.Cells(i, 5).Value & vbCrLf & _
                "Lieu de construction : " & ws.Cells(i, 6).Value & vbCrLf & _
                "Date de soumission : " & ws.Cells(i, 7).Value & vbCrLf & _
                "État d'avancement : " & ws.Cells(i, 8).Value & vbCrLf & _
                "Problèmes rencontrés : " & ws.Cells(i, 9).Value & vbCrLf & _
                "Date limite complément : " & ws.Cells(i, 10).Value & vbCrLf & _
                "Date limite instruction : " & ws.Cells(i, 11).Value & vbCrLf & _
                "Date validation permis : " & ws.Cells(i, 12).Value & vbCrLf & _
                "Date de délivrance : " & ws.Cells(i, 13).Value & vbCrLf & _
                "Date début travaux : " & ws.Cells(i, 14).Value & vbCrLf & _
                "Date fin prévue : " & ws.Cells(i, 15).Value & vbCrLf & _
                "Date limite DAACT : " & ws.Cells(i, 16).Value & vbCrLf & _
                "Responsable dossier : " & ws.Cells(i, 17).Value & vbCrLf & _
                "Email responsable : " & ws.Cells(i, 18).Value & vbCrLf & _
                "Dernière mise à jour : " & dateMAJ
    End With

    ' Chemin et nom du fichier PDF
    Dim cheminPDF As String
    cheminPDF = ThisWorkbook.Path & "\Permis_" & IDPermis & "_" & nomClient & ".pdf"
    
    ' Exporter en PDF
    doc.ExportAsFixedFormat OutputFileName:=cheminPDF, ExportFormat:=17

    ' Fermer Word ou le garder visible
    wordApp.Visible = True ' False si on veut qu'il reste en arrière-plan

    ' Journal des activités
    On Error Resume Next
    Call AjouterAuJournal(NomUtilisateurConnecte, "Exportation PDF", "Permis ID " & IDPermis, "Fiche exportée le " & dateMAJ)
    On Error GoTo 0

    ' Message de confirmation
    MsgBox "Fiche PDF exportée avec succès : " & vbCrLf & cheminPDF, vbInformation


    End Sub

' Mise à jour tableau dynamique selon projet et phase, début à C12
Public Sub MajTableauPhase(ws As Worksheet, phase As String)
    Dim wsData As Worksheet
    Dim lastRow As Long, i As Long, ligneAff As Long
    Dim projetChoisi As String
    
    Set wsData = ThisWorkbook.Sheets("BdD Tâches")
    projetChoisi = ws.Range("C7").Value
    
    If projetChoisi = "" Then
        MsgBox "Veuillez sélectionner un projet dans la liste déroulante en C7.", vbExclamation
        Exit Sub
    End If
    
    ' Effacer tableau précédent (C12:H1000)
    ws.Range("C12:H1000").ClearContents
    
    ' Entêtes
    ws.Range("C12").Value = "Tâche"
    ws.Range("D12").Value = "Début"
    ws.Range("E12").Value = "Fin"
    ws.Range("F12").Value = "Avancement (%)"
    ws.Range("G12").Value = "Responsable"
    ws.Range("H12").Value = "Email"
    
    ligneAff = 13
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsData.Cells(i, 1).Value) = Trim(projetChoisi) And LCase(Trim(wsData.Cells(i, 2).Value)) = LCase(phase) Then
            ws.Cells(ligneAff, 3).Value = wsData.Cells(i, 3).Value        ' Tâche
            ws.Cells(ligneAff, 4).Value = wsData.Cells(i, 4).Value        ' Début
            ws.Cells(ligneAff, 5).Value = wsData.Cells(i, 5).Value        ' Fin
            ws.Cells(ligneAff, 6).Value = wsData.Cells(i, 6).Value        ' Avancement
            ws.Cells(ligneAff, 7).Value = wsData.Cells(i, 7).Value        ' Responsable
            ws.Cells(ligneAff, 8).Value = wsData.Cells(i, 8).Value        ' Email
            ligneAff = ligneAff + 1
        End If
    Next i
    
    ' Format dates et % avancement
    ws.Range("D13:D" & ligneAff - 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("E13:E" & ligneAff - 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("F13:F" & ligneAff - 1).NumberFormat = "0%"
    If ligneAff = 13 Then
    MsgBox "Aucune tâche trouvée pour le projet " & projetChoisi & " dans la phase " & phase & ".", vbInformation
    End If
End Sub

' Dessiner le diagramme de Gantt dans la feuille, couleurs selon phase
Public Sub DessinerDiagrammeGantt(ws As Worksheet, phase As String)
    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim i As Long, nbTaches As Long
    Dim dateMin As Date, dateMax As Date
    Dim ligneDebut As Long, ligneFin As Long
    Dim arrTaches() As String
    Dim arrOffsets() As Double, arrAvancements() As Double, arrDurees() As Double
    Dim projetChoisi As String
    Dim tempSheet As Worksheet, tempRange As Range
    
    projetChoisi = ws.Range("C7").Value
    If projetChoisi = "" Then Exit Sub
    
    ligneDebut = 13
    ligneFin = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If ligneFin < ligneDebut Then Exit Sub
    
    ' Trouver dateMin et dateMax selon colonne Début (D) et Fin (E)
    dateMin = ws.Cells(ligneDebut, 4).Value
    dateMax = ws.Cells(ligneDebut, 5).Value
    For i = ligneDebut + 1 To ligneFin
     If ws.Cells(i, 4).Value < dateMin Then dateMin = ws.Cells(i, 4).Value
     If ws.Cells(i, 5).Value > dateMax Then dateMax = ws.Cells(i, 5).Value
     Next i
    
    ' Appliquer filtre par date visible si défini en H3 (date début) et H4 (date fin)
    If IsDate(ws.Range("H3").Value) Then dateMin = ws.Range("H3").Value
    If IsDate(ws.Range("H4").Value) Then dateMax = ws.Range("H4").Value
    If dateMax < dateMin Then
        MsgBox "La date de fin (H4) doit être supérieure ou égale à la date de début (H3).", vbExclamation
        Exit Sub
    End If
    
    nbTaches = ligneFin - ligneDebut + 1
    
    ' Charger données dans tableaux VBA
    ReDim arrTaches(1 To nbTaches)
    ReDim arrOffsets(1 To nbTaches)
    ReDim arrAvancements(1 To nbTaches)
    ReDim arrDurees(1 To nbTaches)
    
    Dim j As Long
    j = 1
    For i = ligneDebut To ligneFin
        arrTaches(j) = ws.Cells(i, 3).Value
        Dim debutTache As Date, finTache As Date, avancement As Double
        debutTache = ws.Cells(i, 4).Value
        finTache = ws.Cells(i, 5).Value
        avancement = ws.Cells(i, 6).Value
        
        ' Ajuster les dates au filtre dateMin/dateMax
        If debutTache < dateMin Then debutTache = dateMin
        If finTache > dateMax Then finTache = dateMax
        
        arrOffsets(j) = debutTache - dateMin
        arrDurees(j) = finTache - debutTache + 1
        arrAvancements(j) = avancement
        
        j = j + 1
    Next i
    
    ' Préparer feuille temporaire pour les données du graphique
    On Error Resume Next
    Set tempSheet = ThisWorkbook.Sheets("TempGantt")
    On Error GoTo 0
    
    If tempSheet Is Nothing Then
        Set tempSheet = ThisWorkbook.Sheets.Add
        tempSheet.Name = "TempGantt"
        tempSheet.Visible = xlSheetVeryHidden
    Else
        tempSheet.Cells.Clear
    End If
    
    tempSheet.Range("A1:D1").Value = Array("Tâche", "Décalage", "Avancement", "Reste")
    
    For i = 1 To nbTaches
        tempSheet.Cells(i + 1, 1).Value = arrTaches(i)
        tempSheet.Cells(i + 1, 2).Value = arrOffsets(i)
        tempSheet.Cells(i + 1, 3).Value = arrDurees(i) * arrAvancements(i)
        tempSheet.Cells(i + 1, 4).Value = arrDurees(i) * (1 - arrAvancements(i))
    Next i
    
    Set tempRange = tempSheet.Range("A1").CurrentRegion
    
    ' Supprimer anciens graphiques
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
 ' Créer le graphique
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Range("J12").Left, Width:=600, Top:=ws.Range("C12").Top, Height:=nbTaches * 22 + 80)
    Set ch = chartObj.Chart
    ch.ChartType = xlBarStacked
    ch.SetSourceData Source:=tempRange
    
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    
    ' Masquer la série Décalage (1ère série)
    ch.SeriesCollection(1).Format.Fill.Visible = msoFalse
    
    ' Couleurs selon phase
    Dim couleurAvancement As Long, couleurReste As Long
    
    Select Case phase
        Case "Gros oeuvre"
            couleurAvancement = RGB(173, 216, 230) ' bleu clair
            couleurReste = RGB(135, 206, 250)       ' bleu moyen
        Case "Second oeuvre"
            couleurAvancement = RGB(255, 255, 224) ' jaune clair
            couleurReste = RGB(255, 255, 102)       ' jaune plus soutenu
        Case "Finitions"
            couleurAvancement = RGB(144, 238, 144) ' vert clair
            couleurReste = RGB(60, 179, 113)        ' vert moyen
        Case Else
            couleurAvancement = RGB(200, 200, 200)
            couleurReste = RGB(150, 150, 150)
    End Select
    
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = couleurAvancement
    ch.SeriesCollection(3).Format.Fill.ForeColor.RGB = couleurReste
    
    ' Titres axes
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = "Tâches"
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "Jours"
    
    ' Inverser ordre tâches pour correspondre au tableau Excel
    ch.Axes(xlCategory).ReversePlotOrder = True
    
    ' Axe X : valeurs de 0 à nb jours du filtre
    ch.Axes(xlValue).MinimumScale = 0
    ch.Axes(xlValue).MaximumScale = dateMax - dateMin + 1
    
    ' Masquer grille verticale pour plus de lisibilité
    ch.Axes(xlValue).MajorGridlines.Format.Line.Visible = msoFalse
End Sub

Public Sub MajFeuillePhase(ws As Worksheet, phase As String)
    Call MajTableauPhase(ws, phase)
    Call DessinerDiagrammeGantt(ws, phase)
End Sub

Sub ConnexionApplication()
    Dim ws As Worksheet
    Dim nomSaisi As String
    Dim emailSaisi As String
    
    ' --- MES INFORMATIONS DE CONCEPTEUR ---
    Dim nomConcepteur As String
    Dim emailConcepteur As String
    nomConcepteur = "TUO Nompan"
    emailConcepteur = "snompan@gmail.com"
    '---------------------------------------------------
    
    
    On Error Resume Next
    nomSaisi = EspaceConnexion.txtNom.Text
    emailSaisi = EspaceConnexion.txtEmail.Text
    On Error GoTo 0
    
    ' Vérifier si c'est bien le concepteur (on nettoie les espaces et les majuscules/minuscules)
    Dim estConcepteur As Boolean
    If Trim(LCase(nomSaisi)) = Trim(LCase(nomConcepteur)) And Trim(LCase(emailSaisi)) = Trim(LCase(emailConcepteur)) Then
        estConcepteur = True
    Else
        estConcepteur = False
    End If
    
    ' Parcourir toutes les feuilles du classeur pour gérer leur visibilité
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Ouverture" Then
            ws.Visible = xlSheetVisible
            
        ' Liste des feuilles STRICTEMENT réservées au concepteur
        ElseIf ws.Name = "Utilisateurs" Or ws.Name = "bdD Rapports" Then
            If estConcepteur Then
                ws.Visible = xlSheetVisible  ' Visible uniquement si c'est moi
            Else
                ws.Visible = xlSheetVeryHidden ' Masqué pour les autres utilisateurs
            End If
            
        ' Pour toutes les autres feuilles "normales" de l'application immobilière
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    
    ' Message de confirmation
    If estConcepteur Then
        MsgBox "Connexion réussie. Bienvenue, Maître Concepteur !", vbInformation, "Accès Administrateur"
    Else
        MsgBox "Connexion réussie. Bienvenue !", vbInformation, "Accès Utilisateur"
    End If
    
End Sub


Sub Auto_Open()
    Dim ws As Worksheet

    ' Parcourir toutes les feuilles
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Ouverture" Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' Activer la feuille d'ouverture
    Worksheets("Ouverture").Activate
End Sub


Sub RechercherDansBase()
    Dim critere As String
    critere = InputBox("Entrez un mot-clé à rechercher (ex : nom du client)")

    If critere = "" Then Exit Sub

    Dim sh As Worksheet
    Set sh = ActiveSheet  ' S’applique à la feuille active (BdD Clients, BdD Projets, etc.)

    sh.Range("A1").AutoFilter Field:=2, Criteria1:="" & critere & ""
End Sub


Sub ReinitialiserRecherche()
    ' *** Effacer les filtres automatiques ***
    On Error Resume Next
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    On Error GoTo 0

    ' *** Afficher toutes les lignes/colonnes masquées ***
    Cells.EntireRow.Hidden = False
    Cells.EntireColumn.Hidden = False

    ' *** Supprimer surlignage manuel (couleurs) ***
    Cells.Interior.ColorIndex = xlNone

    ' *** Supprimer toute mise en forme conditionnelle ***
    If Cells.FormatConditions.Count > 0 Then
        Cells.FormatConditions.Delete
    End If
End Sub
Public Sub EnvoyerAlerteMail(destinataire As String, sujet As String, corps As String)
    On Error Resume Next
    Dim OutlookApp As Object, MailItem As Object
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

    With MailItem
        .To = destinataire
        .Subject = sujet
        .Body = corps
        .Send
    End With
End Sub


Public Sub VerifierEcheances()
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim dateDebut As Date, dateFin As Date
    Dim nomTache As String, responsable As String, msg As String
    Dim dateAuj As Date
    Dim alerteTrouvee As Boolean
    dateAuj = Date
    alerteTrouvee = False

    Set ws = Sheets("BdD Tâches")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 4).Value) And IsDate(ws.Cells(i, 5).Value) Then
            dateDebut = ws.Cells(i, 4).Value ' Colonne D
            dateFin = ws.Cells(i, 5).Value   ' Colonne E
            nomTache = ws.Cells(i, 3).Value  ' Colonne C
            responsable = ws.Cells(i, 7).Value ' Colonne G

            If dateFin - dateAuj <= 3 And dateFin >= dateAuj Then
                msg = "La tâche """ & nomTache & """ assignée à """ & responsable & _
                      """ se termine le " & Format(dateFin, "dd/mm/yyyy") & "." & vbCrLf & _
                      "Merci de prendre les mesures nécessaires."

                Call EnvoyerAlerteMail(ws.Cells(i, 8).Value, "Échéance proche", msg)
                alerteTrouvee = True
            End If
        End If
    Next i

    If alerteTrouvee Then
        MsgBox "Des tâches sont proches de leur échéance !", vbExclamation, "Alerte"
End If
End Sub

Public Sub VerifierAlertes()
    Dim wsBudget As Worksheet, wsPermis As Worksheet
    Dim i As Long, lastRow As Long
    Dim ecart As Double, dateLimite As Date
    Dim alerte As String
    Dim dateAuj As Date
    Dim dateLimite2 As Date, dateLimite3 As Date, dateLimite4 As Date, dateLimite5 As Date
    Dim alerte2 As String, alerte3 As String, alerte4 As String, alerte5 As String
    dateAuj = Date

    ' Vider la ListBox
    frmAlertes.lstAlertes.Clear

    ' Vérification des budgets déficitaires
    Set wsBudget = Sheets("BdD Budget")
    lastRow = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        ecart = Val(wsBudget.Cells(i, 8).Value) ' colonne Écart budgétaire
        If ecart < 0 Then
            alerte = "!!Budget insuffisant!! - Projet: " & wsBudget.Cells(i, 2).Value & _
                     " | Ecart: " & ecart
            frmAlertes.lstAlertes.AddItem alerte
        End If
    Next i

    ' Vérification des échéances
    Set wsPermis = Sheets("Suivi Permis")
    lastRow = wsPermis.Cells(wsPermis.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
       
         '  vérifier la date limite de complément des dossiers
        If IsDate(wsPermis.Cells(i, 10).Value) Then
            dateLimite2 = wsPermis.Cells(i, 10).Value
            If dateLimite2 - dateAuj <= 5 And dateLimite2 >= dateAuj Then
                alerte2 = "!!Échéance proche - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Date limite: " & Format(dateLimite2, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte2
            ElseIf dateLimite2 < dateAuj Then

                alerte2 = "Échéance dépassée - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Était prévue le: " & Format(dateLimite2, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte2
            End If
        End If
        
        'Vérifier la date limite instruction
        If IsDate(wsPermis.Cells(i, 11).Value) Then
            dateLimite3 = wsPermis.Cells(i, 11).Value
            If dateLimite3 - dateAuj <= 5 And dateLimite3 >= dateAuj Then
                alerte3 = "!!Échéance proche - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Date limite: " & Format(dateLimite3, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte3
            ElseIf dateLimite3 < dateAuj Then

                alerte3 = "Échéance dépassée - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Était prévue le: " & Format(dateLimite3, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte3
            End If
        End If
        'Vérifier la date de début des travaux
        If IsDate(wsPermis.Cells(i, 14).Value) Then
            dateLimite4 = wsPermis.Cells(i, 14).Value
            If dateLimite4 - dateAuj <= 5 And dateLimite4 >= dateAuj Then
                alerte4 = "!!Échéance proche - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Date limite: " & Format(dateLimite4, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte4
            ElseIf dateLimite4 < dateAuj Then

                alerte4 = "Échéance dépassée - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Était prévue le: " & Format(dateLimite4, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte4
            End If
        End If
        'Vérifier la date de fin des travaux
        If IsDate(wsPermis.Cells(i, 15).Value) Then
            dateLimite5 = wsPermis.Cells(i, 15).Value
            If dateLimite5 - dateAuj <= 5 And dateLimite5 >= dateAuj Then
                alerte5 = "!!Échéance proche - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Date limite: " & Format(dateLimite5, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte5
            ElseIf dateLimite5 < dateAuj Then

                alerte5 = "Échéance dépassée - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Était prévue le: " & Format(dateLimite5, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte5
            End If
        End If
         '  vérifier la date limite de déclaration d'achèvement
        If IsDate(wsPermis.Cells(i, 16).Value) Then
            dateLimite = wsPermis.Cells(i, 16).Value
            If dateLimite - dateAuj <= 5 And dateLimite >= dateAuj Then
                alerte = "!!Échéance proche - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Date limite: " & Format(dateLimite, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte
            ElseIf dateLimite < dateAuj Then

                alerte = "Échéance dépassée - Client: " & wsPermis.Cells(i, 2).Value & _
                         " | Était prévue le: " & Format(dateLimite, "dd/mm/yyyy")
                frmAlertes.lstAlertes.AddItem alerte
            End If
        End If
    Next i

    ' Affichage de la boîte d’alerte seulement s’il y a des alertes
    If frmAlertes.lstAlertes.ListCount > 0 Then
        frmAlertes.Show
    End If
End Sub

Public Function VerifierMotDePasse(nomUtilisateur As String, mdpSaisi As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Set ws = Sheets("Utilisateurs")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If LCase(Trim(ws.Cells(i, 1).Value)) = LCase(Trim(nomUtilisateur)) Then
            If Trim(ws.Cells(i, 3).Value) = mdpSaisi Then
                VerifierMotDePasse = True
                Exit Function
            End If
        End If
    Next i

    VerifierMotDePasse = False
End Function

Public Sub AfficherDateHeureAccueil()
    Dim texte As String
    Dim maForme As Shape
    Dim feuilleAccueil As Worksheet

    ' Définir la feuille
    Set feuilleAccueil = ThisWorkbook.Sheets("ACCUEIL")

    ' Construire le texte avec date et heure
    texte = Format(Now, "dddd dd mmmm yyyy – hh:mm:ss")

    ' Essayer de récupérer la forme et modifier son texte
    On Error Resume Next
    Set maForme = feuilleAccueil.Shapes("horlogeAccueil")
    If Not maForme Is Nothing Then
        maForme.TextFrame.Characters.Text = texte
    End If
    On Error GoTo 0
End Sub

Public Sub MettreAJourHorloge()
    Call AfficherDateHeureAccueil
    Application.OnTime Now + TimeValue("00:00:01"), "MettreAJourHorloge" ' Actualise toutes les 30 sec
End Sub

Sub AfficherAlertes()
  frmAlertes.Show
End Sub





