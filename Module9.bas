Attribute VB_Name = "Module9"
Function Eff(I As Integer, X As Variant) As Double

Eff = X(1) * ((I) ^ 4) + X(2) * ((I) ^ 3) + X(3) * ((I) ^ 2) + X(4) * (I) + X(5)

End Function

Function Cost(I As Integer, X As Variant) As Double

Cost = X(1) * ((I) ^ 3) + X(2) * ((I) ^ 2) + X(3) * (I) + X(4)

End Function

Sub Test()

cycleName = "SolarRankine"
NumResults = 6
ligne = 5

Do While Sheets("Results").Cells(ligne, 1) <> cycleName
    ligne = ligne + 1
    If ligne > 500 Then
        Exit Do
    End If
Loop

ligne2 = ligne + NumResults - 1





Effrange = "N" & ligne & ":N" & ligne2
PRrange = "M" & ligne & ":M" & ligne2
CostRange = "R" & ligne & ":R" & ligne2

EffP = "=linest(" & Effrange & "," & PRrange & "^{1,2,3,4})"
CostP = "=linest(" & CostRange & "," & PRrange & "^{1})"




Dim EffE As Variant
Dim CostE As Variant

EffE = Application.Evaluate(EffP)
CostE = Application.Evaluate(CostP)
MsgBox "Equation is y=" & Format(EffE(1), "0.00000000") & "x3+" & Format(EffE(2), "0.0000000") & "x2+" & Format(EffE(3), "0.0000000") & "x+" & Format(EffE(4), "0.0000000")
 I = Sheets("Results").Cells(ligne, 13)
 j = 12
 MaxPR = 0
 MaxEff = 0
 CostOpti = 0
 While I < Sheets("Results").Cells(ligne2, 13)
    
    If Eff(I, EffE) > Eff(I - 1, EffE) And Cost(I, CostE) < Sheets(cycleName).Range("C42") Then
        MaxEff = Eff(I, EffE)
        MaxPR = I
        CostOpti = Cost(I, CostE)
    End If
    j = j + 1
    I = I + 1
    
 Wend

Sheets("Results").Cells(ligne - 1, 20) = "MaxPR"
Sheets("Results").Cells(ligne - 1, 21) = "MaxEFF"
Sheets("Results").Cells(ligne - 1, 22) = "CostOpti"
Sheets("Results").Cells(ligne, 20) = MaxPR
Sheets("Results").Cells(ligne, 21) = MaxEff
Sheets("Results").Cells(ligne, 22) = CostOpti
End Sub
