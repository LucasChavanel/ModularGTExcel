VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoGas 
   Caption         =   "Gas Parameters"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5808
   OleObjectBlob   =   "InfoGas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BoutonNext_Click()

Dim ligne As Integer, I As Integer


If NomStream = "" Or Pressure = "" Or Temperature = "" Then
    MsgBox "A Field is empty"
ElseIf IsNumeric(Pressure.Text) = False Or IsNumeric(Temperature.Text) = False Then
    MsgBox "A Field is not a number"
Else

    Sheets("GT Specs").Cells(11, 11) = NomStream.Text
    Sheets("GT Specs").Cells(11, 11).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(12, 11) = TextBox1.Text
    Sheets("GT Specs").Cells(12, 11).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(9, 11) = Pressure.Text
    Sheets("GT Specs").Cells(9, 11).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(10, 11) = Temperature.Text
    Sheets("GT Specs").Cells(10, 11).Borders.Weight = xlThin
    
    Sheets("GT Specs").Cells(11, 12) = TextBox2.Text
    Sheets("GT Specs").Cells(11, 12).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(12, 12) = TextBox3.Text
    Sheets("GT Specs").Cells(12, 12).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(9, 12) = TextBox4.Text
    Sheets("GT Specs").Cells(9, 12).Borders.Weight = xlThin
    Sheets("GT Specs").Cells(10, 12) = TextBox5.Text
    Sheets("GT Specs").Cells(10, 12).Borders.Weight = xlThin
    
    
    'We add the new Stream to the hidden list of Stream
        If Sheets("ListCompStream").Range("C2") = "" Then
            ligne = 1
        Else
            ligne = Sheets("ListCompStream").Range("C1").End(xlDown).Row
        End If
        Sheets("ListCompStream").Cells(ligne + 1, 3) = NomStream.Text
        
    Unload InfoGas
    CompoGas.Show
End If
End Sub





Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

If Sheets("GT Specs").Range("D24") = True Or Sheets("GT Specs").Range("D25") = True Then
    NomStream.Visible = True
    Pressure.Visible = True
    Temperature.Visible = True
    TextBox1.Visible = True
    Label5.Visible = True
End If
If Sheets("GT Specs").Range("D27") = True Then
    TextBox2.Visible = True
    TextBox3.Visible = True
    TextBox4.Visible = True
    TextBox5.Visible = True
    Label6.Visible = True
End If


End Sub
