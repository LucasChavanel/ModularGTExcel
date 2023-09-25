VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompSpec 
   Caption         =   "UserForm2"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7740
   OleObjectBlob   =   "CompSpec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim ImpellerType As String
Dim Etah As Double
Dim Phi As Double


If TextBox1 = "" Or ComboBox1.Text = "" Then
    MsgBox "A field is missing"
ElseIf IsNumeric(TextBox1) = False Then
    MsgBox "Phi is not a number"
ElseIf Phi > 0.1 Then
    MsgBox "Phi must be inferior to 0.12"

Else
    Phi = TextBox1
    
    If ComboBox1.Text = "Axial Comp" Then
        
        ImpellerType = "Typical"
        Etah = -1631.2 * (Phi) ^ 3 + 335.23 * (Phi) ^ 2 - 20.469 * (Phi) + 1.1862
    Else
        
        
        
        If 0.036 < Phi Then
            MsgBox "Phi Value is too high for a radial compressor"
            ImpellerType = "No"
        ElseIf 0.028 < Phi Then
            ImpellerType = "Erad"
        ElseIf 0.022 < Phi Then
            ImpellerType = "Ep"
        ElseIf 0.02 < Phi Then
            ImpellerType = "Fct"
        ElseIf 0.018 < Phi Then
            ImpellerType = "Frad"
        ElseIf 0.0165 < Phi Then
            ImpellerType = "Fp"
        ElseIf 0.014 < Phi Then
            ImpellerType = "Gct"
        ElseIf 0.012 < Phi Then
            ImpellerType = "Grad"
        ElseIf 0 < Phi Then
            ImpellerType = "Gp"
        End If
        
        If ImpellerType = "No" Then
            Etah = 0
        ElseIf ImpellerType = "Erad" Then
            Etah = -18763 * (Phi) ^ 3 + 759.14 * (Phi) ^ 2 - 1.2546 * (Phi) + 0.717

        ElseIf ImpellerType = "Ep" Then
            Etah = -2335241.55 * (Phi) ^ 4 + 198568 * (Phi) ^ 3 - 6833.9 * (Phi) ^ 2 + 115.03 * (Phi) + 0.0717
            
        ElseIf ImpellerType = "Fct" Then
            Etah = -3984766.09 * (Phi) ^ 4 + 295808 * (Phi) ^ 3 - 8990.5 * (Phi) ^ 2 + 133.22 * (Phi) + 0.0499
            
        ElseIf ImpellerType = "Frad" Then
            Etah = -3784920.62 * (Phi) ^ 4 + 237995 * (Phi) ^ 3 - 6347.2 * (Phi) ^ 2 + 87.785 * (Phi) + 0.3445
        
        ElseIf ImpellerType = "Fp" Then
            Etah = -4419820.8 * (Phi) ^ 4 + 263946 * (Phi) ^ 3 - 6881.2 * (Phi) ^ 2 + 94.47 * (Phi) + 0.3052
        
        ElseIf ImpellerType = "Gct" Then
            Etah = -39236574.18 * (Phi) ^ 4 + 2528667.17 * (Phi) ^ 3 - 62850 * (Phi) ^ 2 + 709.69 * (Phi) - 2.2347
            
        ElseIf ImpellerType = "Grad" Then
            Etah = -17165080.11 * (Phi) ^ 4 + 831391 * (Phi) ^ 3 - 16638 * (Phi) ^ 2 + 166.13 * (Phi) + 0.1465
            
        ElseIf ImpellerType = "Gp" Then
            Etah = -55927417.21 * (Phi) ^ 4 + 3079406.48 * (Phi) ^ 3 - 65600.62 * (Phi) ^ 2 + 637.42 * (Phi) - 1.55
        End If
    End If
    TextBox2 = Etah * 100
End If

End Sub

Private Sub CommandButton2_Click()

Dim col As Integer
col = Sheets("Constant Parameters").Range("A14").End(xlToRight).column + 1
Sheets("Constant Parameters").Cells(15, col) = TextBox1.Text
Sheets("Constant Parameters").Cells(16, col) = ComboBox1.Text
Sheets("Constant Parameters").Cells(17, col) = TextBox3.Text
ChoixComp.Parameter4 = TextBox2
Unload CompSpec

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Initialize()

ComboBox1.AddItem "Axial Comp"
ComboBox1.AddItem "Radial Comp"
TextBox1 = 0.09
TextBox3 = 280

End Sub

