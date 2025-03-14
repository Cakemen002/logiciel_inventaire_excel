VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVisualiser 
   Caption         =   "Visualiser le produit"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "FormVisualiser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVisualiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MLigneProduit As Long
Public Property Let LigneProduit(ByVal NouvelleLigne As Long)
    MLigneProduit = NouvelleLigne
End Property

Private Sub ButtonModifierNomProduit_Click()
    Dim frm As New FormChangementNumero
    frm.MNomProduit = TextBoxProduit.Value
    frm.Show
    Call EntrerInfosDuProduit
End Sub

Private Sub ButtonSupprimer_Click()
    Call SupprimerLeProduit
End Sub

Private Sub UserForm_Activate()
    Call EntrerInfosComboBox
    Call EntrerInfosDuProduit
End Sub

Private Sub ButtonAnnuler_Click()
    Unload Me
End Sub

Private Sub ButtonEnregistrer_Click()
    Call VérificationSiInformationProduitValide
End Sub

Private Sub ButtonRendreInactif_Click()
    Call RendreProduitInactif
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Call Filtrer(ShProduits, 1)
End Sub
Private Sub RendreProduitInactif()
    If ShProduits.Cells(MLigneProduit + 2, 5) <> 0 Then
        If MsgBox("La quantité en inventaire pour ce produit est différente de 0. Voulez-vous toujours rendre ce produit inactif ?", vbYesNo + vbQuestion, "Quantité différente de 0") = vbNo Then
            Exit Sub
        End If
    End If
    With ShProduits.Range("A2").Offset(MLigneProduit)
        DL = ShProduitsInactifs.Cells(ShProduitsInactifs.Rows.Count, 1).End(xlUp).Row + 1
        ShProduitsInactifs.Cells(DL, 1).Value = .Cells(1, 1).Value
        ShProduitsInactifs.Cells(DL, 2).Value = .Cells(1, 2).Value
        ShProduitsInactifs.Cells(DL, 3).Value = .Cells(1, 3).Value
        ShProduitsInactifs.Cells(DL, 4).Value = .Cells(1, 4).Value
        ShProduitsInactifs.Cells(DL, 5).Value = .Cells(1, 5).Value
        ShProduitsInactifs.Cells(DL, 6).Value = .Cells(1, 7).Value
    End With
    ShProduits.Range(MLigneProduit + 2 & ":" & MLigneProduit + 2).Delete
End Sub
Private Sub EntrerInfosComboBox()
    ComboBoxGroupe.List = AvoirListe("TbGroupe")
End Sub

Private Sub EntrerInfosDuProduit()
    With ShProduits.Range("A2").Offset(MLigneProduit)
        TextBoxProduit.Value = .Cells(1, 1).Value
        Me.Caption = "Visualiser le produit " & .Cells(1, 1).Value
        TextBoxDescription.Value = .Cells(1, 2).Value
        ComboBoxGroupe.Value = .Cells(1, 3).Value
        TextBoxLocalisation.Value = .Cells(1, 4).Value
        TextBoxPrix.Value = .Cells(1, 7).Value
    End With
End Sub
Private Sub VérificationSiInformationProduitValide()
    If TextBoxDescription.Value = "" Then
        MsgBox "Le produit doit avoir une description", vbOKOnly + vbExclamation, "Aucune description au produit"
    ElseIf IsNumeric(TextBoxPrix.Value) = False Then
        MsgBox "Le prix du produit doit être une somme", vbOKOnly + vbExclamation, "Prix non numérique"
    ElseIf TextBoxPrix.Value < 0 Then
        MsgBox "Le prix doit avoir une quantité positive", vbOKOnly + vbExclamation, "Prix négatif"
    Else
        Call EntrerInfoModifié
        Unload Me
    End If
End Sub

Private Sub EntrerInfoModifié()
    If MsgBox("Voulez-vous vraiment modifier le produit ?", vbYesNo + vbQuestion, "Modifier le produit") = vbYes Then
        With ShProduits.Range("A2").Offset(MLigneProduit)
            .Cells(1, 2).Value = TextBoxDescription.Value
            .Cells(1, 3).Value = ComboBoxGroupe.Value
            .Cells(1, 4).Value = TextBoxLocalisation.Value
            .Cells(1, 7).Value = TextBoxPrix.Value
        End With
    End If
End Sub
Private Sub SupprimerLeProduit()
    If MsgBox("Voulez-vous vraiment supprimer ce produit définitivement ?", vbQuestion + vbYesNo, "Supprimer le produit") = vbYes Then
        Call SupprimerUnProduit(ShProduits, MLigneProduit + 2)
        Unload Me
    End If
End Sub
