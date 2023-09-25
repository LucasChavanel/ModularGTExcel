VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompoStream 
   Caption         =   "Stream Composition (Number between 0 and 1)"
   ClientHeight    =   1830
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4428
   OleObjectBlob   =   "CompoStream.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompoStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Annuler_Click()
Unload CompoStream
End Sub

Private Sub Suivant_Click()
Dim ligne As Integer, ligne2 As Integer
Dim SumComp As Double

col = Sheets("GT Specs").Range("A7").End(xlToRight).column
ligne = Sheets("GT Specs").Cells(9, col).End(xlDown).Row
ligne2 = Sheets("GT Specs").Cells(9, col).End(xlDown).Row

If TextBox1.Text = "" Then
    MsgBox "The field is empty"
ElseIf IsNumeric(TextBox1.Text) = False Then
    MsgBox "The coefficient is not a number"
Else

    Sheets("GT Specs").Cells(ligne + 1, col) = TextBox1.Text
    Sheets("GT Specs").Cells(ligne + 1, col).Borders.Weight = xlThin
    Unload CompoStream

    ligne = Sheets("GT Specs").Cells(9, col).End(xlDown).Row
    ligne2 = Sheets("GT Specs").Cells(9, 2).End(xlDown).Row
    If (ligne + 1) <> ligne2 Then
        
        CompoStream.Show
    Else
        SumComp = 0
        
        For I = 11 To ligne + 1
            SumComp = SumComp + Sheets("GT Specs").Cells(I, col)
        Next
        

        Sheets("GT Specs").Cells(ligne + 1, col) = SumComp
        Sheets("GT Specs").Cells(ligne + 1, col).Borders.Weight = xlThin
        
        If SumComp <> 1 Then
            MsgBox "The Sum of the Fluids percentage isn't 1, please correct manually"
        End If
    
        NouveauStream.Show
    End If

        
       
  
End If
End Sub

Private Sub UserForm_Initialize()
Dim ligne As Integer, col As Integer, ligne2 As Integer
col = Sheets("GT Specs").Range("A7").End(xlToRight).column
ligne = Sheets("GT Specs").Cells(9, col).End(xlDown).Row
ligne2 = Sheets("GT Specs").Cells(9, 2).End(xlDown).Row

Label1.Caption = Sheets("GT Specs").Cells(ligne + 1, 2)


If Sheets("GT Specs").Cells(ligne2 + 1, 2) = "" And (Sheets("GT Specs").Cells(ligne2, 2) <> "SUM of componens fraction") Then
    Sheets("GT Specs").Cells(ligne2 + 1, 2) = "SUM of componens fraction"
    Sheets("GT Specs").Cells(ligne2 + 1, 2).Borders.Weight = xlThin
End If
End Sub
