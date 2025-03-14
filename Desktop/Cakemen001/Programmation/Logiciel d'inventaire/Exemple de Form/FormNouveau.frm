VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormNouveau 
   Caption         =   "Créer un produit"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "FormNouveau.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormNouveau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonAnnuler_Click()
    Unload Me
End Sub

Private Sub ButtonCréer_Click()
    Call VérifierSiInformationProduitValide
End Sub

Private Sub TextBoxProduit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call VérifierNuméroProduitValide
End Sub

Private Sub UserForm_Initialize()
    Call EntrerInfoComboBox
End Sub

Private Sub EntrerInfoComboBox()
    Me.ComboBoxGroupe.List = AvoirListe("TbGroupe")
    Me.ComboBoxGroupe.ListIndex = 0
End Sub

Private Sub VérifierNuméroProduitValide()
    If TextBoxProduit.Value = "" Then
        MsgBox "Vous devez donner un nom au produit", vbOKOnly + vbExclamation, "Aucun nom au produit"
    ElseIf TrouverInfo(TextBoxProduit.Value, ShProduits, 1) > 0 Then
        MsgBox "Le produit existe déjà", vbOKOnly + vbExclamation, "Produit déjà Existant"
    ElseIf TrouverInfo(TextBoxProduit.Value, ShProduitsInactifs, 1) > 0 Then
        MsgBox "Le produit existe déjà en tant que produit inactif", vbOKOnly + vbExclamation, "Produit déjà Existant"
    Else
        Call UnlockTextBox
    End If
End Sub

Private Sub UnlockTextBox()
    Me.Caption = "Création du produit " & TextBoxProduit.Value
    TextBoxDescription.Enabled = True
    TextBoxLocalisation.Enabled = True
    ComboBoxGroupe.Enabled = True
    TextBoxPrix.Enabled = True
    ButtonCréer.Enabled = True
End Sub

Private Sub VérifierSiInformationProduitValide()
    If TextBoxDescription.Value = "" Then
        MsgBox "Le produit doit avoir une description", vbOKOnly + vbExclamation, "Aucune description au produit"
    ElseIf IsNumeric(TextBoxPrix.Value) = False Then
        MsgBox "Le prix du produit doit être une somme", vbOKOnly + vbExclamation, "Prix non numérique"
    ElseIf TextBoxPrix.Value < 0 Then
        MsgBox "Le prix doit avoir un quantité positive", vbOKOnly + vbExclamation, "Prix négatif"
    Else
        Call CréerUnNouveauProduit
        Unload Me
    End If
End Sub

Private Sub CréerUnNouveauProduit()
    'Prendre la valeur de la dernière ligne
    Dim ligne As Integer
    ligne = DerniereLigne(ShProduits) + 1
    'Insérer les informations dans cette lignes
    With ShProduits
        .Cells(ligne, 1) = Me.TextBoxProduit.Value
        .Cells(ligne, 2) = Me.TextBoxDescription.Value
        .Cells(ligne, 3) = Me.ComboBoxGroupe.Value
        .Cells(ligne, 4) = Me.TextBoxLocalisation.Value
        .Cells(ligne, 5) = 0
        .Cells(ligne, 6) = 0
        .Cells(ligne, 7) = Me.TextBoxPrix.Value
    End With
End Sub
