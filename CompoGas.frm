VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompoGas 
   Caption         =   "Gas Composition (Number between 0 and 1)"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5892
   OleObjectBlob   =   "CompoGas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompoGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Annuler_Click()
Unload CompoGas
End Sub

Private Sub CommandButton1_Click()

   Dim ligne As Integer
    Dim Sum As Double
    
    
        
        ligne = Sheets("GT Specs").Range("J9").End(xlDown).Row
        
       
    'Securité si le champ est vide
    If TextBox1.Text = "" Or (TextBox2.Text = "" And TextBox2.Visible = True) Or (TextBox3.Text = "" And TextBox3.Visible = True) Then
        MsgBox "The field is empty"
    ElseIf (IsNumeric(TextBox2.Text) = False And TextBox2.Visible = True) Or (IsNumeric(TextBox3.Text) = False And TextBox3.Visible = True) Then
        MsgBox "The coefficient is not a number"
        
        
    Else
        
        
    
    
       

        Sheets("GT Specs").Cells(ligne + 1, 10) = TextBox1.Text
        Sheets("GT Specs").Cells(ligne + 1, 10).Borders.Weight = xlThin
        
        
        Sheets("GT Specs").Cells(ligne + 1, 11) = TextBox2.Text
        Sheets("GT Specs").Cells(ligne + 1, 11).Borders.Weight = xlThin
        
        Sheets("GT Specs").Cells(ligne + 1, 12) = TextBox3.Text
        Sheets("GT Specs").Cells(ligne + 1, 12).Borders.Weight = xlThin
        
        Unload CompoGas
        


            Sum = 0
            Sum2 = 0
            ligne = Sheets("GT Specs").Range("K8").End(xlDown).Row
            For I = 13 To ligne
                Sum = Sum + Sheets("GT Specs").Cells(I, 11)
                Sum2 = Sum2 + Sheets("GT Specs").Cells(I, 12)
            Next
'            Sheets("GT Specs").Cells(ligne + 1, 8) = "Sum of Compositions"
'            Sheets("GT Specs").Cells(ligne + 1, 8).Borders.Weight = xlThin
'            Sheets("GT Specs").Cells(ligne + 1, 9) = Sum
'            Sheets("GT Specs").Cells(ligne + 1, 9).Borders.Weight = xlThin
            
            If Sum <> 1 And TextBox1.Visible = True Then
                MsgBox "The Sum of the Brayton Gas percentage isn't 1, please correct manually"
            End If
            
            If Sum2 <> 1 And TextBox2.Visible = True Then
                MsgBox "The Sum of the Rankine Gas percentage isn't 1, please correct manually"
            End If
            
            'On passe au paramétrage de la réaction
            Sheets("ListCompStream").Range("F1") = "New"
            

    
    End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Suivant_Click()
    Dim ligne As Integer
    Dim Sum As Double
    
    
        
        ligne = Sheets("GT Specs").Range("J9").End(xlDown).Row
        
       
    'Securité si le champ est vide
    If TextBox1.Text = "" Or (TextBox2.Text = "" And TextBox2.Visible = True) Or (TextBox3.Text = "" And TextBox3.Visible = True) Then
        MsgBox "The field is empty"
    ElseIf (IsNumeric(TextBox2.Text) = False And TextBox2.Visible = True) Or (IsNumeric(TextBox3.Text) = False And TextBox3.Visible = True) Then
        MsgBox "The coefficient is not a number"
        
        
    Else
    
       

        Sheets("GT Specs").Cells(ligne + 1, 10) = TextBox1.Text
        Sheets("GT Specs").Cells(ligne + 1, 10).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(ligne + 1, 11) = TextBox2.Text
        Sheets("GT Specs").Cells(ligne + 1, 11).Borders.Weight = xlThin
        Sheets("GT Specs").Cells(ligne + 1, 12) = TextBox3.Text
        Sheets("GT Specs").Cells(ligne + 1, 12).Borders.Weight = xlThin
        Unload CompoGas
        CompoGas.Show



            
        
    
    
    End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()

ligne = Sheets("GT Specs").Range("K8").End(xlDown).Row
If ligne > 12 Then
    Sheets("GT Specs").Range("H13:I" & ligne).Clear
End If

If Sheets("GT Specs").Range("D24") = True Or Sheets("GT Specs").Range("D25") = True Then
    TextBox2.Visible = True
    Label2.Visible = True
End If
If Sheets("GT Specs").Range("D27") = True Then
    TextBox3.Visible = True
    Label3.Visible = True
End If



End Sub

