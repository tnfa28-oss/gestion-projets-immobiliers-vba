Attribute VB_Name = "Module2"
Sub TestUtilisateur()
    Dim nom As String, mail As String, motDePasse As String
    nom = "Alice"
    mail = "alice@email.com"
    motDePasse = "secret"

    If UtilisateurExiste(nom, mail, motDePasse) Then
        MsgBox "Utilisateur trouvé"
    Else
        MsgBox "Utilisateur introuvable"
    End If
End Sub
