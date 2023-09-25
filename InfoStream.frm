VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoStream 
   Caption         =   "Stream Parameters"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3828
   OleObjectBlob   =   "InfoStream.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BoutonNext_Click()
If NomStream = "" Or Pressure = "" Or Temperature = "" Or MassFlow = "" Then
    MsgBox "A field is empty"
ElseIf IsNumeric(Pressure.Text) = False Or IsNumeric(MassFlow.Text) = False Or IsNumeric(Temperature.Text) = False Then
    MsgBox "A field is not a number"
Else
  Dim col As Integer

        col = Sheets("GT Specs").Range("A7").End(xlToRight).column + 1
     'On met en forme les cases contenant les info Stream
        Sheets("GT Specs").Cells(10, col) = NomStream.Text
        Sheets("GT Specs").Cells(10, col).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(7, col) = Pressure.Text
        Sheets("GT Specs").Cells(7, col).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(8, col) = Temperature.Text
        Sheets("GT Specs").Cells(8, col).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(9, col) = MassFlow.Text
        Sheets("GT Specs").Cells(9, col).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(6, col) = "Stream" & col - 2
        Sheets("GT Specs").Cells(6, col).Borders.Weight = xlMedium
        Sheets("GT Specs").Cells(6, col).Font.Bold = True

        
        Dim ligne As Integer 'We add the new Stream to the hidden list of Stream
        If Sheets("ListCompStream").Range("C2") = "" Then
            ligne = 1
        Else
            ligne = Sheets("ListCompStream").Range("C1").End(xlDown).Row
        End If
        Sheets("ListCompStream").Cells(ligne + 1, 3) = NomStream.Text
    
 
    
    Unload InfoStream
    CompoStream.Show
End If
End Sub


Private Sub CommandButton2_Click()
Unload InfoStream
End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()


End Sub
