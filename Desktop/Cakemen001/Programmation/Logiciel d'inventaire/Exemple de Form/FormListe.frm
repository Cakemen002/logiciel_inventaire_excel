VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormListbox 
   Caption         =   "Liste des clients"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   OleObjectBlob   =   "FormListe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonFermer_Click()
    Unload Me
End Sub

Private Sub ButtonNouveau_Click()
    Dim frm As New FormNouveauClient
    frm.Show
    Call AddDataToListBox
End Sub

Private Sub TextBoxChercher_Change()
    Call Trouver(ShClients, ListeClients, ComboBoxGroupe, TextBoxChercher)
End Sub

Private Sub ComboBoxGroupe_Change()
    Call Trouver(ShClients, ListeClients, ComboBoxGroupe, TextBoxChercher)
End Sub

Private Sub ListeClients_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ModifierClient
End Sub

Private Sub ListeClients_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        Call ModifierClient
    End If
End Sub

Private Sub UserForm_Initialize()
    Call AddDataToListBox
    Call AddDataToComboBox
End Sub

Private Sub ModifierClient()
    If ShClients.Range("A2") = "" Then
        MsgBox "Vous devez créer un client pour en ouvrir un", vbExclamation, "Aucun client créé"
        Exit Sub
    End If
    Dim Index As Long
    Index = ListeClients.ListIndex
    Dim frm As New FormVisualiserClient
    frm.LigneClient = Index
    frm.Show
    Call AddDataToListBox
    ListeClients.ListIndex = Index
End Sub

Private Sub Trouver(ByVal Feuille, ByVal ListBox, ByVal ComboBox, ByVal TextBox)
    Dim DL As Long
    Dim colonne As Long
    Select Case ComboBox.Value
        Case "Numéro"
            colonne = 1
        Case "Prénom"
            colonne = 2
        Case "Nom"
            colonne = 3
        Case "Adresse"
            colonne = 4
        Case "Code Postale"
            colonne = 5
        Case "Entreprise"
            colonne = 6
        Case "No de Tél."
            colonne = 7
        Case "Courriel"
            colonne = 8
    End Select
    Call Filtrer(Feuille, colonne)
    DL = Feuille.Cells(Feuille.Rows.Count, 1).End(xlUp).Row
    For i = 2 To DL
        If TrouverDansListe(Feuille.Cells(i, colonne).Value, TextBox.Value) = True Then
            ListBox.ListIndex = i - 2
            Exit Sub
        End If
    Next
End Sub

Private Sub AddDataToListBox()
    'Obtenir les infos
    Dim rg As Range
    Set rg = GetRange(ShClients)
    Call Filtrer(ShClients, 1)
    'Mettre les infos dans le listbox Produits
    With ListeClients
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnWidths = "20;80;80;200;55;150;87;50;0"
        .ColumnHeads = True
        .ListIndex = 0
    End With
End Sub

Private Sub AddDataToComboBox()
    With ComboBoxGroupe
        .List = AvoirListe("TbInfoClients")
        .ListIndex = 0
    End With
End Sub

