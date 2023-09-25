VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NouveauGas 
   Caption         =   "Do you want to add another Gas?"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "NouveauGas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NouveauGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload NouveauGas
InfoGas.Show
End Sub

Private Sub CommandButton2_Click()
Unload NouveauGas
ClearReac
InfoReaction.Show
End Sub
