VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVisualiserProduitChoisiFac 
   Caption         =   "Visualiser le produit"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   OleObjectBlob   =   "FormVisualiserProduitChoisiFac.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVisualiserProduitChoisiFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MLigneProduit As Long
Public Suppression As Boolean
Private Sub UserForm_Activate()
    If ShProduitChoisi.Range("A2").Offset(MLigneProduit) = "" Then
        Call AppelerProduit
    Else
        Call OuvrirProduit
    End If
    TextBoxPrix = Format(TextBoxPrix, "# ##0.00 $")
End Sub

Private Sub TextBoxPrix_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBoxPrix = Format(TextBoxPrix, "# ##0.00 $")
End Sub

Private Sub ButtonEnregistrer_Click()
    Call Vérification
End Sub

Private Sub ButtonSupprimer_Click()
    Call SupprimerProduit
    Unload Me
End Sub

Private Sub ButtonAnnuler_Click()
    Call SupprimerFeuilleEntière(ShAjout)
    Unload Me
End Sub

'#########################################################################################################################
Private Sub AppelerProduit()
    Dim frm As New FormProduits
    frm.MInsertion = True
    frm.Show
    With ShProduits.Range("A2").Offset(frm.MLigneProduit)
        TextBoxProduit.Value = .Cells(1, 1).Value
        TextBoxDescription.Value = .Cells(1, 2).Value
        TextBoxQuantité.Value = 1
        TextBoxPrix.Value = .Cells(1, 7).Value
    End With
End Sub

Private Sub OuvrirProduit()
    With ShProduitChoisi.Range("A2").Offset(MLigneProduit)
        TextBoxProduit.Value = .Cells(1, 1).Value
        TextBoxDescription.Value = .Cells(1, 2).Value
        TextBoxQuantité.Value = .Cells(1, 3).Value
        TextBoxPrix.Value = .Cells(1, 6).Value
    End With
End Sub

Private Sub Vérification()
    If TextBoxQuantité.Value = "" Or TextBoxPrix = "" Then
        MsgBox "Vous devez remplir toutes les zones de texte", vbExclamation, "Zones de texte non remplie"
    ElseIf IsNumeric(TextBoxQuantité Or TextboxProix) = False Then
        MsgBox "Les zones de textes doivent être numérique", vbExclamation, "Zones de texte non numérique"
    ElseIf TextBoxPrix < 0 Then
        MsgBox "Le prix doit être positif", vbExclamation, "Prix négatif"
    Else
        Call Enregistrerinformations
    End If
End Sub

Private Sub Enregistrerinformations()
    Dim DL As Long
    With ShProduitChoisi
        If TrouverInfo(TextBoxProduit.Value, ShProduitChoisi, 1) = 0 Then
            DL = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        Else
            DL = TrouverInfo(TextBoxProduit.Value, ShProduitChoisi, 1)
        End If
        .Cells(DL, 1) = TextBoxProduit.Value
        .Cells(DL, 2) = TextBoxDescription.Value
        .Cells(DL, 3) = TextBoxQuantité.Value
        If .Cells(DL, 4) = "" Then
            .Cells(DL, 4) = 0
        End If
        .Cells(DL, 5) = .Cells(DL, 3) - .Cells(DL, 4)
        .Cells(DL, 6) = TextBoxPrix.Value
        .Cells(DL, 7) = .Cells(DL, 5) * .Cells(DL, 6)
    End With
    Unload Me
End Sub

Private Sub SupprimerProduit()
    Call Supprimer(ShProduitChoisi, MLigneProduit + 2)
End Sub
