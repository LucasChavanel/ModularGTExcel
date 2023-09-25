VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TurbSpec 
   Caption         =   "UserForm2"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8892.001
   OleObjectBlob   =   "TurbSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TurbSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Efficiency As Double

Private Sub CommandButton2_Click()

EfficiencyCalcul

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub UserForm_Initialize()

TextBox1 = 0.6
TextBox2 = 1.6
TextBox3 = 0.5
TextBox4 = 0.8
TextBox5 = 460

End Sub


Private Sub EfficiencyCalcul()

Efficiency = 0
If (0.3 < TextBox1 And TextBox1 <= 0.4) And (0.9 < TextBox2 And TextBox2 <= 1.1) Then
    Efficiency = 0.94
ElseIf (0.4 < TextBox1 And TextBox1 <= 0.43) And (0.75 < TextBox2 And TextBox2 <= 1.24) Then
    Efficiency = 0.94
ElseIf (0.43 < TextBox1 And TextBox1 <= 0.74) And (0.75 < TextBox2 And TextBox2 <= 1.37) Then
    Efficiency = 0.94
    
ElseIf (0.74 < TextBox1 And TextBox1 <= 0.86) And (0.75 < TextBox2 And TextBox2 <= 1.37) Then
    Efficiency = 0.93
    
    
ElseIf (0.4 < TextBox1 And TextBox1 <= 0.43) And (1.24 < TextBox2 And TextBox2 <= 1.37) Then
    Efficiency = 0.92
ElseIf (0.43 < TextBox1 And TextBox1 <= 0.5) And (1.37 < TextBox2 And TextBox2 <= 1.6) Then
    Efficiency = 0.92
ElseIf (0.5 < TextBox1 And TextBox1 <= 0.9) And (1.37 < TextBox2 And TextBox2 <= 1.9) Then
    If (0.57 < TextBox1 And TextBox1 <= 0.8) And (1.37 < TextBox2 And TextBox2 <= 1.65) Then
        Efficiency = 0.93
    Else
        Efficiency = 0.92
    End If
ElseIf (0.86 < TextBox1 And TextBox1 <= 0.9) And (0.8 < TextBox2 And TextBox2 <= 1.37) Then
    Efficiency = 0.92
    
ElseIf (0.9 < TextBox1 And TextBox1 <= 1.02) And (0.8 < TextBox2 And TextBox2 <= 1.9) Then
    Efficiency = 0.91
    
ElseIf (0.43 < TextBox1 And TextBox1 <= 0.5) And (1.6 < TextBox2 And TextBox2 <= 1.9) Then
    Efficiency = 0.9
ElseIf (0.5 < TextBox1 And TextBox1 <= 1.02) And (1.9 < TextBox2 And TextBox2 <= 2.28) Then
    Efficiency = 0.9
ElseIf (0.96 < TextBox1 And TextBox1 <= 1.02) And (0.8 < TextBox2 And TextBox2 <= 1.9) Then
    Efficiency = 0.9
    
ElseIf (0.55 < TextBox1 And TextBox1 <= 1.07) And (2.28 < TextBox2 And TextBox2 <= 2.4) Then
    Efficiency = 0.89
ElseIf (1.02 < TextBox1 And TextBox1 <= 1.07) And (0.8 < TextBox2 And TextBox2 <= 2.28) Then
    Efficiency = 0.89
    
ElseIf (0.6 < TextBox1 And TextBox1 <= 1.15) And (2.4 < TextBox2 And TextBox2 <= 2.6) Then
    Efficiency = 0.88
ElseIf (1.07 < TextBox1 And TextBox1 <= 1.15) And (1.2 < TextBox2 And TextBox2 <= 2.4) Then
    Efficiency = 0.88
    
ElseIf (0.62 < TextBox1 And TextBox1 <= 1.15) And (2.6 < TextBox2 And TextBox2 <= 2.7) Then
    Efficiency = 0.87
     
ElseIf (0.7 < TextBox1 And TextBox1 <= 1.27) And (2.7 < TextBox2 And TextBox2 <= 3) Then
    Efficiency = 0.86
ElseIf (1.15 < TextBox1 And TextBox1 <= 1.27) And (2.4 < TextBox2 And TextBox2 <= 2.7) Then
    Efficiency = 0.86
Else
    Efficiency = 0
End If

If Efficiency = 0 Then
    Label7.Caption = "Parameters don't allow to find a correct efficiency"
Else
    Label7.Caption = Efficiency * 100
End If

End Sub



Private Sub CommandButton1_Click()
If IsNumeric(TextBox1) = False Or IsNumeric(TextBox2) = False Or IsNumeric(TextBox3) = False Or IsNumeric(TextBox4) = False Then
    MsgBox "One field is not a number"
Else
    Dim col As Integer
    col = Sheets("Constant Parameters").Range("A6").End(xlToRight).column + 1
    Sheets("Constant Parameters").Cells(7, col) = TextBox1.Text
    Sheets("Constant Parameters").Cells(8, col) = TextBox2.Text
    Sheets("Constant Parameters").Cells(9, col) = TextBox3.Text
    Sheets("Constant Parameters").Cells(10, col) = TextBox4.Text
    Sheets("Constant Parameters").Cells(11, col) = TextBox5.Text
    ChoixComp.Parameter4 = Label7.Caption
    Unload TurbSpec
End If
End Sub
