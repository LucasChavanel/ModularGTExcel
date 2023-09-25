VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Considered 
   Caption         =   "Considered Cycles"
   ClientHeight    =   5415
   ClientLeft      =   156
   ClientTop       =   588
   ClientWidth     =   5724
   OleObjectBlob   =   "Considered.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Considered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()

If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False Then
    MsgBox "You need to select at least One Cycle"
Else
    Sheets("GT SPecs").Range("D19") = CheckBox1.Value
    Sheets("GT SPecs").Range("D20") = CheckBox2.Value
    Sheets("GT SPecs").Range("D21") = CheckBox3.Value
    Sheets("GT SPecs").Range("D24") = CheckBox4.Value
    Sheets("GT SPecs").Range("D25") = CheckBox5.Value
    Sheets("GT SPecs").Range("D26") = CheckBox6.Value
    Sheets("GT SPecs").Range("D27") = CheckBox7.Value
    Sheets("GT SPecs").Range("D28") = CheckBox8.Value
    Unload Considered
    
    ligne = 19
    
    If CheckBox1.Value = True Then
        Sheets("GT SPecs").Cells(ligne, 6) = "Brayton"
        ligne = ligne + 1
        Sheets("GT SPecs").Cells(ligne, 6) = "2Comp Brayton"
        ligne = ligne + 1
        Sheets("GT SPecs").Cells(ligne, 6) = "2Turb Brayton"
        ligne = ligne + 1
        Sheets("GT SPecs").Cells(ligne, 6) = "Regeneration Brayton"
        ligne = ligne + 1
         Sheets("GT SPecs").Cells(ligne, 6) = "2Comp3Turb Regeneration Brayton"
        ligne = ligne + 1
        If CheckBox8.Value = True Then
            Sheets("GT SPecs").Cells(ligne, 6) = "Solar Brayton"
            ligne = ligne + 1
            Sheets("GT SPecs").Cells(ligne, 6) = "Solar 2Comp Brayton"
            ligne = ligne + 1
            Sheets("GT SPecs").Cells(ligne, 6) = "Solar 2Turb Brayton"
            ligne = ligne + 1
            Sheets("GT SPecs").Cells(ligne, 6) = "Solar Regeneration Brayton"
            ligne = ligne + 1
        End If
    End If
    
    If CheckBox2.Value = True Then
        If CheckBox6.Value = True Then
            Sheets("GT SPecs").Cells(ligne, 6) = "Rankine Boiler"
            ligne = ligne + 1
        End If
        If CheckBox7.Value = True Then
            Sheets("GT SPecs").Cells(ligne, 6) = "Rankine Fired Heater"
            ligne = ligne + 1
        End If
        
        Sheets("GT SPecs").Cells(ligne, 6) = "Rankine ORC"
        ligne = ligne + 1

        If CheckBox8.Value = True Then
        
            If CheckBox6.Value = True Then
                Sheets("GT SPecs").Cells(ligne, 6) = "Solar Rankine Boiler"
                ligne = ligne + 1
            End If
            If CheckBox7.Value = True Then
                Sheets("GT SPecs").Cells(ligne, 6) = "Solar Rankine Fired Heater"
                ligne = ligne + 1
            End If
            
            Sheets("GT SPecs").Cells(ligne, 6) = "Solar Rankine ORC"
            ligne = ligne + 1
        End If
    End If
    
    If CheckBox3.Value = True Then
        Sheets("GT SPecs").Cells(ligne, 6) = "Combined Cycle"
        Sheets("GT SPecs").Cells(ligne + 1, 6) = "Combined Cycle 2Comp"
        Sheets("GT SPecs").Cells(ligne + 2, 6) = "Combined Cycle 2Turb"
        Sheets("GT SPecs").Cells(ligne + 3, 6) = "Combined Cycle Regeneration"
        ligne = ligne + 4
                
        If CheckBox8.Value = True Then
                Sheets("GT SPecs").Cells(ligne, 6) = "Solar Combined Cycle"
                ligne = ligne + 1
        End If
    End If
End If

End Sub

Private Sub CommandButton2_Click()
Unload Considered
End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
If Sheets("GT Specs").Range("F19") <> "" Then
    ligne = Sheets("GT Specs").Range("F18").End(xlDown).Row
    Sheets("GT Specs").Range("F19:F" & ligne).Clear
End If
End Sub
