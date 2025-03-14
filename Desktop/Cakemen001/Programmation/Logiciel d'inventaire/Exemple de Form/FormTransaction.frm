VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTransaction 
   Caption         =   "Soumission #"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12255
   OleObjectBlob   =   "FormTransaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ShProduitChoisi.Cells(1, 3) = "Soumis"
    TextBoxTransport = 0
    Call EntrerInfoListbox
    Call CalculerSousTotal
End Sub

Private Sub LabelNuméro_Click()
    Call AppelerSoumission
End Sub

Private Sub TextBoxNuméro_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBoxNuméro = "" Then
        Call NouvelleSoumission
    ElseIf IsNumeric(TextBoxNuméro) = False Then
        MsgBox "Les soumissions sont seulement des nombres", vbExclamation, "Valeur non numérique"
        TextBoxNuméro = ""
    ElseIf TrouverInfo(TextBoxNuméro, ShSoumissions, 1) = 0 Or TextBoxNuméro = ProchainNuméro(ShSoumissions) Then
        MsgBox "Ce numéro de soumission n'existe pas", vbExclamation, "Soumission inconnue"
        TextBoxNuméro = ""
    Else
        Call OuvrirSoumission
    End If
    TextBoxNuméro.TabStop = False
End Sub

Private Sub LabelClient_Click()
    Call AppelerClient
End Sub

Private Sub LabelVendeur_Click()
    Call AppelerVendeur
End Sub

Private Sub TextBoxNoClient_Change()
    Call EntrerInfoClient
End Sub

Private Sub TextBoxNoClient_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TrouverInfo(TextBoxNoClient, ShClients, 1) = 0 Then
        TextBoxNoClient = ""
    End If
End Sub

Private Sub TextBoxNoVendeur_Change()
    Call EntrerInfoVendeur
End Sub

Private Sub TextBoxNoVendeur_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TrouverInfo(TextBoxNoVendeur, ShEmployés, 1) = 0 Then
        TextBoxNoVendeur = ""
    End If
End Sub

Private Sub ListBoxProduit_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call OuvrirProduit
End Sub

Private Sub ListBoxProduit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        Call OuvrirProduit
    End If
End Sub

Private Sub TextBoxTransport_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBoxTransport = "" Or IsNumeric(TextBoxTransport) = False Then
        TextBoxTransport = 0
    End If
    Call CalculerSousTotal
End Sub

Private Sub TextBoxSousTotal_Change()
    Call CalculerSousTotal
End Sub

Private Sub TextBoxTPS_Change()
    Call CalculerSousTotal
End Sub

Private Sub TextBoxTVQ_Change()
    Call CalculerSousTotal
End Sub

Private Sub TextBoxTotal_Change()
    Call CalculerSousTotal
End Sub

Private Sub ButtonEnregistrer_Click()
    Call VérificationAvantEnregistrement
End Sub

Private Sub ButtonCommander_Click()
    Call Commander
End Sub

Private Sub ButtonSupprimer_Click()
    If TextBoxNuméro = "" Or TextBoxNuméro = ProchainNuméro(ShSoumissions) Then
        MsgBox "La supression est impossible puisque aucune soumission est en cours ou celle en cours n'a jamais été enregistré", vbExclamation, "Suppression impossible"
    Else
        Call SupprimerEnregistrement(ShSoumissions, TextBoxNuméro)
        Call ViderSoumission
        TextBoxNuméro = ""
    End If
End Sub

Private Sub ButtonImprimer_Click()
    Call VérificationAvantImpression
End Sub

Private Sub ButtonFermer_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Call SupprimerFeuilleEntière(ShProduitChoisi)
End Sub

'#########################################################################################################################
Private Sub EntrerInfoListbox()
    Dim rg As Range
    Set rg = GetRange(ShProduitChoisi)
    If ShProduitChoisi.Range("A2") <> "" Then
        Set rg = rg.Resize(rg.Rows.Count + 1)
    End If
    With ListBoxProduit
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnWidths = "60;180;50;0;0;80;80"
        .ColumnHeads = True
        .ListIndex = 0
    End With
End Sub

Private Sub CalculerSousTotal()
    Dim SousTotal As Single
    Dim Transport As Single
    Dim TPS As Single
    Dim TVQ As Single
    Dim Total As Single
    SousTotal = WorksheetFunction.Sum(ShProduitChoisi.Range("G:G"))
    Transport = TextBoxTransport.Value
    TPS = (SousTotal + Transport) * 0.05
    TVQ = (SousTotal + Transport) * 0.0975
    Total = (SousTotal + Transport) * 1.14975
    TextBoxSousTotal = Format(SousTotal, "# ##0.00 $")
    TextBoxTransport = Format(Transport, "# ##0.00 $")
    TextBoxTPS = Format(TPS, "# ##0.00 $")
    TextBoxTVQ = Format(TVQ, "# ##0.00 $")
    TextBoxTotal = Format(Total, "# ##0.00 $")
End Sub

Private Sub AppelerSoumission()
    Dim frm As New FormListeSouComFac
    frm.MTypeTransaction = "Sou"
    frm.Show
    If frm.MNuméro > 0 Then
        TextBoxNuméro = frm.MNuméro
        Call OuvrirSoumission
    End If
End Sub

Private Sub NouvelleSoumission()
    Call ViderSoumission
    TextBoxNuméro = ProchainNuméro(ShSoumissions)
    TextBoxDate = Format(Date, "dd-mm-yyyy")
    Me.Caption = "Soumission #" & TextBoxNuméro
End Sub

Private Sub OuvrirSoumission()
    Call ViderSoumission
    Dim Ligne As Long
    Dim L As Long
    Me.Caption = "Soumission #" & TextBoxNuméro
    Ligne = TrouverInfo(TextBoxNuméro, ShSoumissions, 1)
    With ShSoumissions.Range("A" & Ligne)
        TextBoxNoClient = .Cells(1, 2)
        TextBoxRéference = .Cells(1, 9)
        TextBoxDescription = .Cells(1, 8)
        TextBoxDate = Format(.Cells(1, 4), "dd-mm-yyyy")
        TextBoxNoVendeur = .Cells(1, 3)
        TextBoxTransport = .Cells(1, 7)
        If .Cells(2, 1) = "" Then
            L = 2
            Do While .Cells(L, 1) = ""
                ShProduitChoisi.Cells(L, 1) = .Cells(L, 2)
                ShProduitChoisi.Cells(L, 2) = ShProduits.Cells(TrouverInfo(ShProduitChoisi.Cells(L, 1), ShProduits, 1), 2)
                ShProduitChoisi.Cells(L, 3) = .Cells(L, 3)
                ShProduitChoisi.Cells(L, 6) = .Cells(L, 5)
                ShProduitChoisi.Cells(L, 7) = .Cells(L, 3) * .Cells(L, 5)
                L = L + 1
            Loop
        End If
        Call EntrerInfoListbox
        Call Filtrer(ShProduitChoisi, 1)
        Call CalculerSousTotal
    End With
End Sub

Private Sub AppelerClient()
    Dim frm As New FormClient
    frm.MInsertion = True
    frm.Show
    If frm.MNuméroClient > 0 Then
        TextBoxNoClient = frm.MNuméroClient
    End If
End Sub

Private Sub AppelerVendeur()
    Dim frm As New FormEmployes
    frm.MInsertion = True
    frm.Show
    TextBoxNoVendeur = frm.MNuméroEmployé
End Sub

Private Sub EntrerInfoClient()
    If TrouverInfo(TextBoxNoClient.Value, ShClients, 1) > 1 Then
        TextBoxNoClient2 = TextBoxNoClient
        TextBoxNomClient = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 2) & " " & ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 3)
        TextBoxNomClient2 = TextBoxNomClient
        TextBoxAdresse = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 4)
        TextBoxCodePostale = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 5)
        TextBoxEntreprise = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 6)
        TextBoxTéléphone = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 7)
        TextBoxInformation = ShClients.Cells(TrouverInfo(TextBoxNoClient, ShClients, 1), 8)
    Else
        TextBoxNoClient2 = ""
        TextBoxNomClient = ""
        TextBoxNomClient2 = ""
        TextBoxAdresse = ""
        TextBoxCodePostale = ""
        TextBoxEntreprise = ""
        TextBoxTéléphone = ""
        TextBoxInformation = ""
    End If
End Sub

Private Sub EntrerInfoVendeur()
    If TrouverInfo(TextBoxNoVendeur.Value, ShEmployés, 1) > 1 Then
        TextBoxNomVendeur = ShEmployés.Cells(TrouverInfo(TextBoxNoVendeur, ShEmployés, 1), 2) & " " & ShEmployés.Cells(TrouverInfo(TextBoxNoVendeur, ShEmployés, 1), 3)
    Else
        TextBoxNomVendeur = ""
    End If
End Sub

Private Sub OuvrirProduit()
    Dim frm As New FormVisualiserProduitChoisi
    Dim Index As Long
    Index = ListBoxProduit.ListIndex
    frm.MLigneProduit = Index
    frm.Show
    Call EntrerInfoListbox
    Call CalculerSousTotal
    Call Filtrer(ShProduitChoisi, 1)
    If frm.Suppression = True Then
        Index = Index - 1
    End If
    ListBoxProduit.ListIndex = -1
    ListBoxProduit.ListIndex = Index
End Sub

Private Sub VérificationAvantEnregistrement()
    If TextBoxNuméro = "" Then
        MsgBox "Il n'y a pas de soumission ouverte", vbExclamation, "Aucune soumission ouverte"
    ElseIf TextBoxNoClient = "" Then
        MsgBox "Il n'y a pas de client attaché à la soumission", vbExclamation, "Aucun client choisi"
    ElseIf TextBoxNoVendeur = "" Then
        MsgBox "Il n'y a pas de vendeur attaché à la soumission", vbExclamation, "Aucun vendeur choisi"
    ElseIf ShProduitChoisi.Range("A2") = "" Then
        MsgBox "Il n'y a pas de produit dans la soumission", vbExclamation, "Aucun produit choisi"
    Else
        Call EnregistrerSoumission
    End If
End Sub

Private Sub EnregistrerSoumission()
    Dim Ligne As Long
    Dim L As Long
    Ligne = TrouverInfo(TextBoxNuméro, ShSoumissions, 1)
    With ShSoumissions.Range("A" & Ligne)
        If TextBoxNuméro = ProchainNuméro(ShSoumissions) Then
            .Offset(1) = TextBoxNuméro + 1
        End If
        .Cells(1, 2) = Format(TextBoxNoClient, "0")
        .Cells(1, 3) = Format(TextBoxNoVendeur, "0")
        .Cells(1, 4) = Format(TextBoxDate, "0")
        .Cells(1, 5) = Format(TextBoxTotal, "0.00 $")
        .Cells(1, 7) = Format(TextBoxTransport, "0.00 $")
        .Cells(1, 8) = TextBoxDescription
        .Cells(1, 9) = TextBoxRéference
        Do While .Offset(1) = ""
            .Offset(1).EntireRow.Delete
        Loop
        L = 2
        Do While ShProduitChoisi.Cells(L, 2) <> ""
            .Offset(1).EntireRow.Insert
            .Cells(2, 2) = ShProduitChoisi.Cells(L, 1)
            .Cells(2, 3) = ShProduitChoisi.Cells(L, 3)
            .Cells(2, 5) = ShProduitChoisi.Cells(L, 6)
            L = L + 1
        Loop
    End With
End Sub

Private Sub Commander()
    If MsgBox("Voulez-vous vraiment commander cette soumission ?", vbQuestion + vbYesNo, "Commander la soumission") = vbYes Then
        Dim frm As New FormCommande
        frm.TextBoxNuméro = ProchainNuméro(ShCommandes)
        frm.TextBoxSoumission = TextBoxNuméro
        frm.TextBoxNoClient = TextBoxNoClient
        frm.TextBoxRéference = TextBoxRéference
        frm.TextBoxDescription = TextBoxDescription
        frm.TextBoxNoVendeur = TextBoxNoVendeur
        frm.TextBoxTransport = TextBoxTransport
        With ShCommandes.Cells(ShCommandes.Rows.Count, 1).End(xlUp)
            .Cells(2, 1) = frm.TextBoxNuméro + 1
            .Cells(1, 2) = Format(TextBoxNoClient, "0")
            .Cells(1, 3) = Format(TextBoxNoVendeur, "0")
            .Cells(1, 4) = Format(TextBoxDate, "0")
            .Cells(1, 5) = Format(TextBoxTotal, "0.00 $")
            .Cells(1, 6) = TextBoxNuméro
            .Cells(1, 7) = Format(TextBoxTransport, "0.00 $")
            .Cells(1, 8) = TextBoxDescription
            .Cells(1, 9) = TextBoxRéference
            L = 2
            Do While ShProduitChoisi.Cells(L, 2) <> ""
                .Offset(1).EntireRow.Insert
                .Cells(2, 2) = ShProduitChoisi.Cells(L, 1)
                .Cells(2, 3) = ShProduitChoisi.Cells(L, 3)
                .Cells(2, 4) = 0
                .Cells(2, 5) = ShProduitChoisi.Cells(L, 6)
                L = L + 1
            Loop
        End With
        Call SupprimerEnregistrement(ShSoumissions, TextBoxNuméro)
        Unload Me
        frm.Mcommandé = True
        frm.Show
    End If
End Sub

Private Sub ViderSoumission()
    TextBoxNoClient = ""
    TextBoxRéference = ""
    TextBoxDescription = ""
    TextBoxDate = ""
    TextBoxNoVendeur = ""
    TextBoxTransport = 0
    Call SupprimerFeuilleEntière(ShProduitChoisi)
    Call EntrerInfoListbox
    Call CalculerSousTotal
End Sub

Private Sub VérificationAvantImpression()
    If TextBoxNuméro = "" Then
        MsgBox "Il n'y a pas de soumission ouverte", vbExclamation, "Aucune soumission ouverte"
    ElseIf TextBoxNoClient = "" Then
        MsgBox "Il n'y a pas de client attaché à la soumission", vbExclamation, "Aucun client choisi"
    ElseIf TextBoxNoVendeur = "" Then
        MsgBox "Il n'y a pas de vendeur attaché à la soumission", vbExclamation, "Aucun vendeur choisi"
    ElseIf ShProduitChoisi.Range("A2") = "" Then
        MsgBox "Il n'y a pas de produit dans la soumission", vbExclamation, "Aucun produit choisi"
    Else
        Call Imprimercopie
    End If
End Sub

Private Sub Imprimercopie()
    Dim Ncopie As String
    Do
        Ncopie = InputBox("Combien de copie désirez-vous imprimer ?", "Impression de la soumission", 1)
        If Ncopie = "" Then
            MsgBox "Vous avez rien écrit dans la zone de texte", vbExclamation, "Champ de texte vide"
        ElseIf IsNumeric(Ncopie) = False Then
            MsgBox "Vous devez entrer une quantité numérique dans la zone de texte", vbExclamation, "Valeur non numérique"
        ElseIf Ncopie < 0 Then
            MsgBox "Vous devez entrer une quantitée positive", vbExclamation, "Valeur négative"
        End If
    Loop While IsNumeric(Ncopie) = False Or Ncopie < 0 Or Ncopie = ""
    If Ncopie > 0 Then
        With ShImpressionSoumission
            .Range("G2") = TextBoxNuméro
            .Range("H31") = TextBoxSousTotal
            .Range("H32") = TextBoxTransport
            .Range("H33") = TextBoxTPS
            .Range("H34") = TextBoxTVQ
            .Range("G35") = TextBoxTotal
            For I = 5 To 77 Step 36
                .Range("B" & I) = TextBoxNoClient
                .Range("B" & I + 1) = TextBoxNomClient
                .Range("B" & I + 2) = TextBoxEntreprise
                .Range("B" & I + 3) = TextBoxAdresse
                .Range("B" & I + 5) = TextBoxCodePostale
                .Range("E" & I) = TextBoxNomVendeur
                .Range("E" & I + 1) = TextBoxDate
                .Range("E" & I + 3) = TextBoxCourriel
                .Range("E" & I + 4) = TextBoxTéléphone
                .Range("E" & I + 5) = TextBoxDescription
            Next
            Dim NProduit As Long
            Dim LChoisi As Long
            Dim NImpr As Long
            NProduit = ShProduitChoisi.Cells(ShProduitChoisi.Rows.Count, 1).End(xlUp).Row - 1
            LChoisi = 2
            .Range("C36") = "Page 1 de 1"
            NImpr = 1
            For I = 13 To 29
                .Cells(I, 1) = ShProduitChoisi.Cells(LChoisi, 1)
                .Cells(I, 2) = ShProduitChoisi.Cells(LChoisi, 2)
                .Cells(I, 3) = ShProduitChoisi.Cells(LChoisi, 3)
                .Cells(I, 7) = ShProduitChoisi.Cells(LChoisi, 6)
                .Cells(I, 8) = ShProduitChoisi.Cells(LChoisi, 7)
                LChoisi = LChoisi + 1
            Next
            If NProduit > 17 Then
                .Range("C36") = "Page 1 de 2"
                .Range("C72") = "Page 2 de 2"
                NImpr = 2
                For I = 49 To 65
                    .Cells(I, 1) = ShProduitChoisi.Cells(LChoisi, 1)
                    .Cells(I, 2) = ShProduitChoisi.Cells(LChoisi, 2)
                    .Cells(I, 3) = ShProduitChoisi.Cells(LChoisi, 3)
                    .Cells(I, 7) = ShProduitChoisi.Cells(LChoisi, 6)
                    .Cells(I, 8) = ShProduitChoisi.Cells(LChoisi, 7)
                    LChoisi = LChoisi + 1
                Next
            End If
            If NProduit > 34 Then
                .Range("C36") = "Page 1 de 3"
                .Range("C72") = "Page 2 de 3"
                NImpr = 3
                For I = 85 To 101
                    .Cells(I, 1) = ShProduitChoisi.Cells(LChoisi, 1)
                    .Cells(I, 2) = ShProduitChoisi.Cells(LChoisi, 2)
                    .Cells(I, 3) = ShProduitChoisi.Cells(LChoisi, 3)
                    .Cells(I, 7) = ShProduitChoisi.Cells(LChoisi, 6)
                    .Cells(I, 8) = ShProduitChoisi.Cells(LChoisi, 7)
                    LChoisi = LChoisi + 1
                Next
            End If
            .PrintOut 1, NImpr, Ncopie
        End With
    End If
End Sub


