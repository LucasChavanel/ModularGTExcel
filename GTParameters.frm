VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GTParameters 
   Caption         =   "Succesive run"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10092
   OleObjectBlob   =   "GTParameters.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GTParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()

Unload GTParameters

End Sub

Private Sub CommandButton2_Click()


If TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Or TextBox14.Text = "" Then
    MsgBox "At least One field is missing"
ElseIf IsNumeric(TextBox4.Text) = False Or IsNumeric(TextBox5.Text) = False Then
    MsgBox "At least one field is not a number"
Else



Sheets("GT Specs").Range("D9") = TextBox4.Text
'Sheets("GT Specs").Range("D9").Borders.Weight = xlMedium

Sheets("GT Specs").Range("D10") = TextBox5.Text
'Sheets("GT Specs").Range("D10").Borders.Weight = xlMedium

Sheets("GT Specs").Range("D11") = TextBox6.Text
'Sheets("GT Specs").Range("D11").Borders.Weight = xlMedium

Sheets("GT Specs").Range("D12") = TextBox7.Text
'Sheets("GT Specs").Range("D12").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G9") = TextBox8.Text
'Sheets("GT Specs").Range("G9").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G10") = TextBox9.Text
'Sheets("GT Specs").Range("G10").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G11") = TextBox10.Text
'Sheets("GT Specs").Range("G11").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G12") = TextBox11.Text
'Sheets("GT Specs").Range("G12").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G13") = TextBox12.Text
'Sheets("GT Specs").Range("G13").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G14") = TextBox13.Text
'Sheets("GT Specs").Range("G14").Borders.Weight = xlMedium

Sheets("GT Specs").Range("G15") = TextBox14.Text
'Sheets("GT Specs").Range("G15").Borders.Weight = xlMedium

Sheets("GT Specs").Columns(3).AutoFit
Sheets("GT Specs").Columns(4).AutoFit
Sheets("GT Specs").Columns(6).AutoFit
Sheets("GT Specs").Columns(7).AutoFit

Unload GTParameters
End If

End Sub







Private Sub Label10_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub UserForm_Click()

End Sub
