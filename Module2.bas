Attribute VB_Name = "Module2"
'--------------------This module contains the subs for managing the excel data (deleting streams components, etc...)

'Call all the deleting subs. It resets the excel
Sub ClearAll()

ClearDatas
ClearFluides
ClearReac
ClearGas
ClearGT
DeleteSheets

End Sub

Sub ClearFluidPage()
ClearReac
ClearGas
ClearFluides

End Sub



Public Sub ClearReac()
'Sub to clear the reaction parameters

col = Sheets("GT Specs").Range("A7").End(xlToRight).column + 3
col2 = Sheets("GT Specs").Cells(7, col).End(xlToRight).column
ligne = Sheets("GT Specs").Range("B7").End(xlDown).Row

If Sheets("GT Specs").Cells(7, col) <> "" Then
    Sheets("GT Specs").Range(Cells(4, col), Cells(ligne, col2)).Clear
    Sheets("GT Specs").Range(Cells(4, col), Cells(ligne, col2)).Interior.Color = RGB(255, 255, 255)
End If

End Sub


Public Sub ClearFluides()

'Sub to delete the input streams


col = Sheets("GT Specs").Range("A7").End(xlToRight).column
ligne = Sheets("GT Specs").Range("B7").End(xlDown).Row

If col > 1 And ligne > 1 Then
    Sheets("GT Specs").Range(Cells(6, 1), Cells(ligne, col)).Clear
    Sheets("GT Specs").Range(Cells(6, 1), Cells(ligne, col)).Interior.Color = RGB(255, 255, 255)

End If

End Sub

Public Sub ClearGas()


ligne = Sheets("GT Specs").Range("B7").End(xlDown).Row + 4
col = Sheets("GT Specs").Cells(ligne, 1).End(xlToRight).column
ligne2 = Sheets("GT Specs").Cells(ligne, 2).End(xlDown).Row

If Sheets("GT Specs").Cells(ligne, col) <> "" Then
    Sheets("GT Specs").Range(Cells(ligne - 1, 1), Cells(ligne2, col)).Clear
    Sheets("GT Specs").Range(Cells(ligne - 1, 1), Cells(ligne2, col)).Interior.Color = RGB(255, 255, 255)
End If

End Sub

Sub ClearGT()
Sheets("Components").Range(Cells(28, 3), Cells(31, 3)).Clear
Sheets("Components").Range(Cells(28, 3), Cells(31, 3)).Interior.Color = RGB(255, 255, 255)
End Sub
Sub ClearDatas()

'Sub to delete the component table in the "Components" sheet
Dim colonne As Integer
Dim ligneStream As Integer
Dim ligneQStream As Integer

'Size of the table
colonne = Sheets("Components").Range("A2").End(xlToRight).column + 1
Sheets("Components").Range(Cells(2, 3), Cells(18, colonne)).Clear
Sheets("Components").Range(Cells(2, 3), Cells(18, colonne)).Interior.Color = RGB(255, 255, 255)
ligneStream = Sheets("ListCompStream").Range("C1").End(xlDown).Row

'If there is components, delete them
If ligneStream > 3 Then
Sheets("ListCompStream").Range("C4:C" & ligneStream).Clear
End If

'We clear the hidden Sheet from the created streams (between components)
If Sheets("ListCompStream").Range("D2") <> "" Then
ligneQStream = Sheets("ListCompStream").Range("D1").End(xlDown).Row
Sheets("ListCompStream").Range("D2:D" & ligneQStream).Clear
End If

End Sub

Sub DeleteSheets()
Dim NombreSheets As Integer

NombreSheets = ActiveWorkbook.Sheets.count

'If there is more than the parameters sheet, we delete them (results sheets and concatenation of results)
If NombreSheets > 6 Then
    For I = 6 To NombreSheets - 1
        Sheets(6).Delete
    Next
End If

End Sub


'----------------Some subs use with the buttons to open userforms-------------

Public Sub OuvrirReac()


    InfoReaction.Show

End Sub

Sub LancerUserForm()

ChoixComp.Show

End Sub

Sub lancerUserCycle()
    
    Sheets("ListCompStream").Range("M2:M60").Clear
    CyclesChoose.Show
   
End Sub

Sub LancerFormFluid()

ChoixFluides.Show

End Sub

Sub LancerCompo()
CompoGas.Show
End Sub
Sub LancerReac()
NouveauStream.Show
End Sub


Sub LancerGas()


    InfoGas.Show

End Sub

Sub LaunchConsider()

Considered.Show
End Sub
