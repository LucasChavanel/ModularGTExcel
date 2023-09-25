Attribute VB_Name = "Module8"
'Modification of the cycle pressure Ratio to find optimum cost/Efficiency point

Function ChangePressureRatio(myCompoCollec As Collection, strcase As String, myCycleCollec As Collection, Datas() As Variant, cycleName As String) As String
Dim myComp As Collection
Dim myTurb As Collection
Dim oCycle As cCycle
Dim oComponent As cComponent
Dim PressureRatio As Double
Dim NumResults As Integer
Dim Yes As String
PressureRatio = 1
Dim simCase As SimulationCase
Dim StreamA As ProcessStream
Dim hyFluidPkg As HYSYS.FluidPackage
Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim Intervalle As Double
Dim NumberCompressor As Integer, NumberTurbine As Integer, NumberPump As Integer
    Dim RatioPerComp As Double, RatioPerTurb As Double, RatioPerPump As Double, NewPressureRatio As Double

NumResults = 1
For Each oCycle In myCycleCollec
    
    If (oCycle.CType = "Brayton" Or oCycle.CType = "Reheat Brayton" Or oCycle.CType = "Regeneration Brayton" Or oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine") And Sheets(cycleName).Range("C37") <> "Combined Cycle" Then

        PressureRatio = oCycle.PressureRatio
        
        If oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine" Then
            Intervalle = 40 'Number of results wanted
            NewPressureRatio = 10
        End If
        
        If oCycle.CType = "Brayton" Or oCycle.CType = "Reheat Brayton" Or oCycle.CType = "Regeneration Brayton" Then
            Intervalle = 7 'Number of results wanted
            NewPressureRatio = 3
        End If
            
        NumberCompressor = oCycle.NumberCompressor
        NumberTurbine = oCycle.NumberTurbine
        NumberPump = oCycle.NumberPump
        X = UBound(Datas, 2) - LBound(Datas, 2) + 1
        
        If oCycle.CType = "Brayton" Or oCycle.CType = "Reheat Brayton" Or oCycle.CType = "Regeneration Brayton" Then
            MaxPR = 34
        End If
        
        If oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine" Then
            MaxPR = 181
        End If
        
        While NewPressureRatio < MaxPR
        
            If cycleName = "Fired Rankine" Or cycleName = "Solar Fired Rankine" Then
                
    
                
    
                RatioPerTurb = (NewPressureRatio * 0.8) ^ (1 / NumberTurbine)
                If 1 / RatioPerTurb > 1 Then
                    RatioPerTurb = 1.1
                End If
                
                For I = 0 To X - 2
    
                    If Datas(1, I) = "Pump1" Then
                        Datas(8, I) = (NewPressureRatio) ^ (1 / 4)
                        hyFlwSht.MaterialStreams.Item("Mix1").Pressure = Sheets(cycleName).Range("E82") * Datas(8, I) * 0.99
                    ElseIf Datas(1, I) = "Pump2" Then
                        Datas(8, I) = (NewPressureRatio) ^ (3 / 4)
                    ElseIf Datas(0, I) = "Steam Turbine" Then
                        If Datas(1, I) <> "STurb4" Then
                            Datas(8, I) = 1 / RatioPerTurb
                        End If
                    End If
                Next
            Else
            
                
                If NumberCompressor <> 0 Then
                    RatioPerComp = (NewPressureRatio) ^ (1 / NumberCompressor)
                End If
                If NumberPump <> 0 Then
                    RatioPerPump = (NewPressureRatio) ^ (1 / NumberPump)
                End If
                If NumberTurbine <> 0 Then
                    RatioPerTurb = (NewPressureRatio * 0.82) ^ (1 / NumberTurbine)
                    If 1 / RatioPerTurb > 1 Then
                        RatioPerTurb = 1.1
                    End If
                End If
                For I = 0 To X - 2
                    If Datas(0, I) = "Compressor" Then
                        Datas(8, I) = RatioPerComp
                    ElseIf Datas(0, I) = "Pump" Then
                        Datas(8, I) = RatioPerPump
                    ElseIf Datas(0, I) = "Gas Turbine" Then
                        Datas(8, I) = 1 / RatioPerTurb
                    ElseIf Datas(0, I) = "Steam Turbine" Then
                        Datas(8, I) = 1 / RatioPerTurb
                    End If
                Next
                
             End If
            
            
            Set myCompoCollec = Modif_Turbine(Datas(), strcase, cycleName)
            Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
            
            If cycleName <> "Fired Rankine" Then
               Aye = HXRecalibration(Datas(), strcase, cycleName)
            End If
            
             If Sheets(cycleName).Range("C40") = "Fixed Power" Then
                Aye = ApproxGTPower(myCompoCollec, myCycleCollec, Datas(), strcase, cycleName)
                If cycleName <> "Fired Rankine" Then
                    Aye = HXRecalibration(Datas(), strcase, cycleName)
                End If
            End If

            Set myComp = CompDesign(Datas(), strcase, cycleName)


            Set myTurb = TurbDesign(Datas(), strcase, cycleName)

            Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
            Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
            
            Yes = ResultsCycle(myCompoCollec, strcase, myCycleCollec, Datas(), cycleName, NumResults)
            NumResults = NumResults + 1
            NewPressureRatio = NewPressureRatio + Intervalle

        Wend
        
        ChangePressureRatio = "Youhouooo"
        
     ElseIf (oCycle.CType = "Brayton" Or oCycle.CType = "Reheat Brayton" Or oCycle.CType = "Regeneration Brayton") And Sheets(cycleName).Range("C37") = "Combined Cycle" Then

        PressureRatio = oCycle.PressureRatio
    
        Intervalle = 7 'Number of results wanted
        NewPressureRatio = 3
        
            
        NumberCompressor = oCycle.NumberCompressor
        NumberTurbine = oCycle.NumberTurbine
        
        X = UBound(Datas, 2) - LBound(Datas, 2) + 1
        
        While NewPressureRatio < 34
        
            If NumberCompressor <> 0 Then
                RatioPerComp = (NewPressureRatio) ^ (1 / NumberCompressor)
            End If
    
            If NumberTurbine <> 0 Then
                RatioPerTurb = (NewPressureRatio * 0.85) ^ (1 / NumberTurbine)
                If 1 / RatioPerTurb > 1 Then
                    RatioPerTurb = 1.1
                End If
            End If
            
            For I = 0 To X - 2
                If Datas(0, I) = "Compressor" Then
                    Datas(8, I) = RatioPerComp

                ElseIf Datas(0, I) = "Gas Turbine" Then
                    Datas(8, I) = 1 / RatioPerTurb

                End If
            Next
            
            
            
            
            Set myCompoCollec = Modif_Turbine(Datas(), strcase, cycleName)
            Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
            Aye = HXRecalibration(Datas(), strcase, cycleName)
            
             If Sheets(cycleName).Range("C40") = "Fixed Power" Then
                Aye = ApproxGTPower(myCompoCollec, myCycleCollec, Datas(), strcase, cycleName)
                Aye = HXRecalibration(Datas(), strcase, cycleName)
            End If

            Set myComp = CompDesign(Datas(), strcase, cycleName)


            Set myTurb = TurbDesign(Datas(), strcase, cycleName)

            Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
            Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
            
            Yes = ResultsCycle(myCompoCollec, strcase, myCycleCollec, Datas(), cycleName, NumResults)
            NumResults = NumResults + 1
            NewPressureRatio = NewPressureRatio + Intervalle

        Wend
        
        ChangePressureRatio = "Youhouooo"
        
        
    
    End If
    
   
    
    
Next

NumResults = NumResults - 1

ligne = 5

Do While Sheets("Results").Cells(ligne, 1) <> cycleName
    ligne = ligne + 1
    If ligne > 500 Then
        Exit Do
    End If
Loop

ligne2 = ligne + NumResults - 1


Effrange = "O" & ligne & ":O" & ligne2
PRrange = "N" & ligne & ":N" & ligne2
CostRange = "T" & ligne & ":T" & ligne2
FCostRange = "Q" & ligne & ":Q" & ligne2

EffP = "=linest(" & Effrange & "," & PRrange & "^{1,2,3,4})"
CostP = "=linest(" & CostRange & "," & PRrange & "^{1,2,3})"


Ij = Sheets(cycleName).Range("E40")
Ny = Sheets(cycleName).Range("E41")
PhiM = Sheets(cycleName).Range("E42")
OH = Sheets(cycleName).Range("C43")
NA = Sheets(cycleName).Range("E43")
'MaxOCost = Sheets("GT Specs").Range("G9") '/ Sheets("GT Specs").Range("D9") * (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1)) * PhiM / OH
'MaxFuelCost = Sheets("GT Specs").Range("G10") / Sheets("GT Specs").Range("D9") * (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1)) / OH
Sheets("Results").Range("Y5") = MaxOCost

Dim EffE As Variant
Dim CostE As Variant
Dim CostF As Variant

EffE = Application.Evaluate(EffP)
CostE = Application.Evaluate(CostP)
CostF = Application.Evaluate(CostFu)
'MsgBox "Equation is y=" & Format(EffE(1), "0.00000000") & "x3+" & Format(EffE(2), "0.0000000") & "x2+" & Format(EffE(3), "0.0000000") & "x+" & Format(EffE(4), "0.0000000")
 I = Sheets("Results").Cells(ligne, 14)
 j = 12
 MaxPR = 0
 MaxEff = 0
 CostOpti = 1
 While I < Sheets("Results").Cells(ligne2, 14)
   
    If Cost(I, CostE) < CostOpti Then
        MaxEff = Eff(I, EffE)
        MaxPR = I
        CostOpti = Cost(I, CostE)
    End If
    j = j + 1
    I = I + 1
    
 Wend

Sheets("Results").Cells(ligne - 1, 22) = "MaxPR"
Sheets("Results").Cells(ligne - 1, 23) = "MaxEFF"
Sheets("Results").Cells(ligne - 1, 24) = "CostOpti"
Sheets("Results").Cells(ligne, 22) = MaxPR
Sheets("Results").Cells(ligne, 23) = MaxEff
Sheets("Results").Cells(ligne, 24) = CostOpti

ligne = Sheets("Results").Range("A4").End(xlDown).Row + 1
Sheets("Results").Cells(ligne, 1) = "Next"

End Function


