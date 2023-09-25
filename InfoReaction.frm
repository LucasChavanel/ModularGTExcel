VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoReaction 
   Caption         =   "Reaction Parameters"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5856
   OleObjectBlob   =   "InfoReaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoReaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public NombreReac As Integer


Private Sub CommandButton1_Click()
If ComboBox_Reac.ListIndex = -1 Or TextBoxStoch.Text = "" Then
    MsgBox "The Field is empty"
ElseIf IsNumeric(TextBoxStoch.Text) = False Then
    MsgBox "The coefficient is not a number"
Else
    Enregistrer_Reac
    Unload InfoReaction
    Sheets("ListCompStream").Range("F1") = "Same"
    InfoReaction.Show
End If

End Sub

Private Sub CommandButton2_Click()
If ComboBox_Reac.ListIndex = -1 Or TextBoxStoch.Text = "" Then
    MsgBox "The Field is empty"
ElseIf IsNumeric(TextBoxStoch.Text) = False Then
    MsgBox "The coefficient is not a number"
Else
    Enregistrer_Reac
    Unload InfoReaction
    Sheets("ListCompStream").Range("F1") = "New"

       
End If
End Sub

Private Sub CommandButton3_Click()
'Ajouter Clear
Unload InfoReaction
End Sub

Private Sub UserForm_Initialize()

Dim I As Integer
Dim nbLignes As Integer, col As Integer
Dim Reac As String
'If Sheets("ListCompStream").Range("F1") = "New" Then
'
'
'col = Sheets("GT Specs").Range("J9").End(xlToRight).column
'If Sheets("GT Specs").Cells(7, col) <> "" Then
'    col = Sheets("GT Specs").Cells(7, col).End(xlToRight).column + 1
'End If
'
'
'
'Sheets("GT Specs").Cells(7, col) = "Reactives"
'Sheets("GT Specs").Cells(7, col).Borders.Weight = xlMedium
'Sheets("GT Specs").Cells(7, col).Font.Bold = True
'Sheets("GT Specs").Cells(7, col + 1) = "Stochiometric coefficients"
'Sheets("GT Specs").Cells(7, col + 1).Borders.Weight = xlMedium
'Sheets("GT Specs").Cells(7, col + 1).Font.Bold = True
'
'col = Sheets("GT Specs").Cells(7, col).End(xlToRight).column
'NombreReac = (col - Sheets("GT Specs").Range("A7").End(xlToRight).column - 2) / 2
'Sheets("GT Specs").Cells(6, col - 1) = "Reaction" & NombreReac
'Sheets("GT Specs").Cells(6, col - 1).Borders.Weight = xlMedium
'Sheets("GT Specs").Cells(6, col - 1).Font.Bold = True
'
'End If

nbLignes = Sheets("GT Specs").Range("J9").End(xlDown).Row
'On ajoute les composants à la comboBox, la liste de composant doit être modifiée
'A la main en cas de rajout
ComboBox_Reac.AddItem "Oxygen"
ComboBox_Reac.AddItem "Nitrogen"
ComboBox_Reac.AddItem "H2O"
ComboBox_Reac.AddItem "CO2"
ComboBox_Reac.AddItem "CO"

For I = 13 To nbLignes
    Reac = Sheets("GT Specs").Cells(I, 10)
    ComboBox_Reac.AddItem Reac
Next
End Sub

Sub Enregistrer_Reac()

Dim ligne As Integer, col As Integer


ligne = Sheets("GT Specs").Range("N8").End(xlDown).Row + 1
Sheets("GT Specs").Cells(ligne, 14) = ComboBox_Reac.Value
Sheets("GT Specs").Cells(ligne, 14).Borders.Weight = xlThin
Sheets("GT Specs").Cells(ligne, 15) = TextBoxStoch.Text
Sheets("GT Specs").Cells(ligne, 15).Borders.Weight = xlThin

End Sub
