VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SolarStream 
   Caption         =   "Solar Heater Input Stream"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "SolarStream.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SolarStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload SolarStream
End Sub

Private Sub Done_Click()
Dim col As Integer
col = Sheets("Constant Parameters").Range("A19").End(xlToRight).column + 1
If TextBox1.Text = "" Or TextBox1.Text = "" Or TextBox1.Text = "" Then
    MsgBox "At least one field is missing"
ElseIf IsNumeric(TextBox1.Text) = False Or IsNumeric(TextBox4.Text) = False Or IsNumeric(TextBox5.Text) = False Then
    MsgBox "At least One field is not a number"
Else
        Sheets("Constant Parameters").Cells(23, col) = CDbl(TextBox1.Text)
        Sheets("Constant Parameters").Cells(24, col) = CDbl(TextBox4.Text)
        Sheets("Constant Parameters").Cells(25, col) = CDbl(TextBox5.Text)
        Unload SolarStream
        ChoixComp.LabelP4.Visible = False
End If
End Sub

Private Sub UserForm_Click()

End Sub
