VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChoixFluides 
   Caption         =   "Fluids "
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5328
   OleObjectBlob   =   "ChoixFluides.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChoixFluides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AjoutFluid_Click()


    'Lorsque que l'on clique sur le bouton Rajout QStream
    Dim nbLignes As Integer
    If TextBoxFluid.Text = "" Then
        MsgBox "Veuillez rentrer un fluide"
    Else
        'On met à jour liste dans feuille cachée
        If Sheets("GT Specs").Range("B11") = "" Then
            'si aucun QStream ajouté, combine pour pas que ça bug
            Sheets("GT Specs").Range("B11") = "QTest"
            nbLignes = Sheets("GT Specs").Range("B10").End(xlDown).Row - 1
            Sheets("GT Specs").Range("B11") = ""
        Else
            nbLignes = Sheets("GT Specs").Range("B10").End(xlDown).Row
        End If
        'On ajoute et on met à jour
        Sheets("GT Specs").Cells(nbLignes + 1, 2) = TextBoxFluid.Text
        Sheets("GT Specs").Cells(nbLignes + 1, 2).Borders.Weight = xlThin
        ModifFluid
        TextBoxFluid.Text = ""
    End If


End Sub


Private Sub ModifFluid()
    
    If Sheets("GT Specs").Cells(11, 1) <> "" Then
        ListBoxFluid.Clear
        
        Dim Fluid As String
        Dim I As Integer
        Dim nbLignes As Integer
        
        nbLignes = Sheets("GT Specs").Range("B9").End(xlDown).Row
        
        For I = 10 To nbLignes
            'On met à jour les 3 listes de streams
            Fluid = Sheets("GT Specs").Cells(I + 1, 2)
            If Fluid <> "SUM of componens fraction" Then
                ListBoxFluid.AddItem Fluid
            End If
            
        Next
    End If
End Sub


Private Sub BoutonAnnuler_Click()

    Unload ChoixFluides
    
End Sub

Private Sub BoutonTerminer_Click()

    If Sheets("GT Specs").Range("B11") = "" Then
        MsgBox "Enter at least one fluid"
    Else
        Unload ChoixFluides
        InfoStream.Show
    End If
    
End Sub



Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()

ClearFluides
ModifFluid

Sheets("GT Specs").Cells(7, 2) = "Pressure (kPa)"
Sheets("GT Specs").Cells(7, 2).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(7, 2).Font.Bold = True
Sheets("GT Specs").Cells(8, 2) = "Temperature (K)"
Sheets("GT Specs").Cells(8, 2).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(8, 2).Font.Bold = True
Sheets("GT Specs").Cells(9, 2) = "Mass Flow (kg/s)"
Sheets("GT Specs").Cells(9, 2).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(9, 2).Font.Bold = True
Sheets("GT Specs").Cells(10, 2) = "Name"
Sheets("GT Specs").Cells(10, 2).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(10, 2).Font.Bold = True
Sheets("GT Specs").Cells(11, 1) = "Components"
Sheets("GT Specs").Cells(11, 1).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(11, 1).Font.Bold = True
Sheets("GT Specs").Cells(7, 1) = "Gas"
Sheets("GT Specs").Cells(7, 1).Borders.Weight = xlMedium
Sheets("GT Specs").Cells(7, 1).Font.Bold = True

End Sub
