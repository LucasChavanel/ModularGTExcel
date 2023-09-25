VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CyclesChoose 
   Caption         =   "Type of Cycle in play"
   ClientHeight    =   1350
   ClientLeft      =   72
   ClientTop       =   312
   ClientWidth     =   5580
   OleObjectBlob   =   "CyclesChoose.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CyclesChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If ComboBox1.Text = "" Then
    MsgBox "One Field is not complete"
Else
    If Sheets("ListCompStream").Cells(2, 13) = "" Then
        'si aucun QStream ajouté, combine pour pas que ça bug
        Sheets("ListCompStream").Cells(2, 13) = "QTest"
        nbLignes = Sheets("ListCompStream").Range("M1").End(xlDown).Row - 1
        Sheets("ListCompStream").Cells(2, 13) = ""
    Else
        nbLignes = Sheets("ListCompStream").Range("M1").End(xlDown).Row
    End If
    
    Sheets("ListCompStream").Cells(nbLignes + 1, 13) = ComboBox1.Text
    Sheets("ListCompStream").Cells(nbLignes + 1, 14) = ComboBox2.Text
    
    nblignes2 = Sheets("ListCompStream").Range("L1").End(xlDown).Row
    
    If nblignes2 <= nbLignes + 1 Then
        Unload CyclesChoose
        GTParameters.Show
    Else
        Unload CyclesChoose
        CyclesChoose.Show
    End If
End If
End Sub

Private Sub CommandButton2_Click()

Unload CyclesChoose

End Sub


Private Sub Label4_Click()

End Sub

Private Sub UserForm_Initialize()
    If Sheets("ListCompStream").Cells(2, 13) = "" Then
        nbLignes = 1
    Else
        nbLignes = Sheets("ListCompStream").Range("M1").End(xlDown).Row
    End If
    
    Label2.Caption = Sheets("ListCompStream").Cells(nbLignes + 1, 12)
    
    nbLignes = Sheets("ListCompStream").Range("J1").End(xlDown).Row
    For I = 2 To nbLignes
        ComboBox1.AddItem Sheets("ListCompStream").Cells(I, 10)
    Next
    
    col = Sheets("Fluids").Range("A7").End(xlToRight).column
    For I = 3 To col
        ComboBox2.AddItem Sheets("Fluids").Cells(10, I)
    Next
    

End Sub
