VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NouveauReac 
   Caption         =   "UserForm1"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "NouveauReac.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NouveauReac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload NouveauReac
Sheets("GT Specs").Range("M4") = "Reaction 2"
Sheets("GT Specs").Range("M4").Borders.Weight = xlMedium
Sheets("GT Specs").Range("M4").Font.Bold = True
Sheets("GT Specs").Range("M5") = "Reactif"
Sheets("GT Specs").Range("M5").Borders.Weight = xlMedium
Sheets("GT Specs").Range("M5").Font.Bold = True
Sheets("GT Specs").Range("N5") = "Stochio Coeff"
Sheets("GT Specs").Range("N5").Borders.Weight = xlMedium
Sheets("GT Specs").Range("N5").Font.Bold = True

InfoReaction.Show

End Sub

Private Sub CommandButton2_Click()
Unload NouveauReac
End Sub


