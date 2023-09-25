Attribute VB_Name = "Module4"
'Function who extract the results from Aspen and concatenate them in a sheet, sorted by components
Function ExtractResults(myCompoCollec As Collection, strcase As String, myCycleCollec As Collection, Datas() As Variant) As String

'//////////////////////////Results extraction////////////////////////////////////


NombreResults = ActiveWorkbook.Sheets.count - 5
Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim name As String
name = "Results" & NombreResults + 1
ActiveWorkbook.Sheets.Add(Before:=Worksheets("ListCompStream")).name = name

'Format of the Result sheet
Sheets(name).Cells(1, 1) = "N° Results"
Sheets(name).Cells(1, 2) = "Name of the cycle"
Sheets(name).Cells(1, 3) = "Type of cycle"
Sheets(name).Cells(1, 4) = "Name of the component"
Sheets(name).Cells(1, 5) = "Type of Component"
Sheets(name).Cells(1, 6) = "Power (kW)"
Sheets(name).Cells(1, 7) = "Isentropic Efficiency"
Sheets(name).Cells(1, 8) = "Pressure Ratio"
Sheets(name).Cells(1, 9) = "Flaming Temperature (K)"
Sheets(name).Cells(1, 10) = "Fuel Mass Flow (kg/s)"
Sheets(name).Cells(1, 11) = "Number of Stages"
Sheets(name).Cells(1, 12) = "Tip Speed (m/s)"
Sheets(name).Cells(1, 13) = "Rotating Speed (RPM)"
Sheets(name).Cells(1, 14) = "Mean Diameter (m)"
Sheets(name).Cells(1, 15) = "Cost ($)"
Sheets(name).Rows(1).Font.Bold = True

'Writing the results

Dim oComponent As cComponent
ligne = 2
For Each oComponent In myCompoCollec

    Sheets(name).Cells(ligne, 1) = name
    Sheets(name).Cells(ligne, 2) = oComponent.cycleName
    
    For Each oCycle In myCycleCollec
        If oCycle.name = oComponent.cycleName Then
            Sheets(name).Cells(ligne, 3) = oCycle.CType
        End If
    Next
    Sheets(name).Cells(ligne, 4) = oComponent.CompName
    Sheets(name).Cells(ligne, 5) = oComponent.CompType
    Sheets(name).Cells(ligne, 6) = oComponent.power
    Sheets(name).Cells(ligne, 7) = oComponent.Efficiency
    Sheets(name).Cells(ligne, 8) = oComponent.PressureRatio
    
    If oComponent.CompType = "Combustion Chamber" Or oComponent.CompType = "Fired Heater" Then
        Sheets(name).Cells(ligne, 9) = oComponent.Tout
    Else
        Sheets(name).Cells(ligne, 9) = 0
    End If
    
    
    
    If oComponent.CompType = "Combustion Chamber" Or oComponent.CompType = "Fired Heater" Then
        Sheets(name).Cells(ligne, 10) = oComponent.Fin2
    Else
        Sheets(name).Cells(ligne, 10) = 0
    End If
    
    Sheets(name).Cells(ligne, 11) = oComponent.NumberStage
    Sheets(name).Cells(ligne, 12) = oComponent.TipSpeed
    Sheets(name).Cells(ligne, 13) = oComponent.RotatingSpeed
    Sheets(name).Cells(ligne, 14) = oComponent.Diameter * 2.54 / 100
    Sheets(name).Cells(ligne, 15) = oComponent.PEC
    ligne = ligne + 1
    
Next
    
Dim power As Double, Cost As Double, h As Integer, endSheet As Integer
power = 0
Cost = 0
endSheet = ligne - 1
For h = 2 To ligne - 1
    power = power + Sheets(name).Cells(h, 6)
    Cost = Cost + Sheets(name).Cells(h, 15)
Next
Sheets(name).Cells(ligne, 6) = power
Sheets(name).Cells(ligne, 15) = Cost
ligne = ligne + 1
Sheets(name).Cells(ligne, 1) = "Name of cycle"
Sheets(name).Cells(ligne, 2) = "Type of Cycle"
Sheets(name).Cells(ligne, 3) = "Power Produced by Cycle (kW)"
Sheets(name).Cells(ligne, 4) = "Efficiency of the cycle"
Sheets(name).Cells(ligne, 5) = "Pressure Ratio"
Sheets(name).Cells(ligne, 6) = "Piloting Feed Name"
Sheets(name).Cells(ligne, 7) = "Feed Mass Flow (kg/s)"
Sheets(name).Cells(ligne, 8) = "Fuel Mass Flow (kg/s)"
Sheets(name).Cells(ligne, 9) = "Cost ($)"
Sheets(name).Rows(ligne).Font.Bold = True
ligne = ligne + 1


For Each oCycle In myCycleCollec
    Cost = 0
    Sheets(name).Cells(ligne, 1) = oCycle.name
    Sheets(name).Cells(ligne, 2) = oCycle.CType
    Sheets(name).Cells(ligne, 3) = oCycle.power
    Sheets(name).Cells(ligne, 4) = oCycle.Efficiency
    Sheets(name).Cells(ligne, 5) = oCycle.PressureRatio
    Sheets(name).Cells(ligne, 6) = oCycle.StreamPilot
    Sheets(name).Cells(ligne, 7) = oCycle.FeedFlow
    Sheets(name).Cells(ligne, 8) = oCycle.FuelFlow
    
    For h = 2 To endSheet
        If Sheets(name).Cells(h, 2) = oCycle.name Then
            Cost = Cost + Sheets(name).Cells(h, 15)
        End If
    Next
    Sheets(name).Cells(ligne, 9) = Cost
    ligne = ligne + 1
Next


'Resize of the sheet
With Sheets(name)
    For I = 1 To 16
        .Columns(I).AutoFit
    Next
End With

ExtractResults = "youououo"

End Function

'Sub to concatenate results into one sheet and readable by R
Sub AnalysisData()
Dim NombreResults As Integer

NombreResults = ActiveWorkbook.Sheets.count - 5
Dim ligne As Integer, ligne2 As Integer, col As Integer
ligne = Sheets("Results1").Range("A1").End(xlDown).Row + 2
ligne2 = Sheets("Results1").Cells(ligne, 1).End(xlDown).Row

ActiveWorkbook.Sheets.Add(Before:=Worksheets("ListCompStream")).name = "CondensedResults"

Sheets("CondensedResults").Cells(1, 1) = "Gas Turbine Optimisation"

col = 1
For I = ligne + 1 To ligne2

    Sheets("CondensedResults").Cells(4, col) = Sheets("Results1").Cells(I, 1)
    Sheets("CondensedResults").Cells(5, col) = Sheets("Results1").Cells(I, 2)
    Sheets("CondensedResults").Cells(3, col + 2) = "Pressure Ratio"
    Sheets("CondensedResults").Cells(3, col + 3) = "Efficiency"
    Sheets("CondensedResults").Cells(3, col + 4) = "Power (kW)"
    Sheets("CondensedResults").Cells(3, col + 5) = "Mass Flow (kg/s)"
    Sheets("CondensedResults").Cells(3, col + 6) = "Fuel Mass Flow (kg/s)"
    Sheets("CondensedResults").Cells(3, col + 7) = "Cost ($)"
    
    
    For j = 1 To NombreResults
        name = "Results" & j
        Sheets("CondensedResults").Cells(3 + j, col + 1) = Sheets(name).Cells(2, 1)
        Sheets("CondensedResults").Cells(3 + j, col + 2) = Sheets(name).Cells(I, 5)
        Sheets("CondensedResults").Cells(3 + j, col + 3) = Sheets(name).Cells(I, 4)
        Sheets("CondensedResults").Cells(3 + j, col + 4) = Sheets(name).Cells(I, 3)
        Sheets("CondensedResults").Cells(3 + j, col + 5) = Sheets(name).Cells(I, 7)
        Sheets("CondensedResults").Cells(3 + j, col + 6) = Sheets(name).Cells(I, 8)
        Sheets("CondensedResults").Cells(3 + j, col + 7) = Sheets(name).Cells(I, 9)
    
    Next
    
    col = col + 9
Next

For j = 1 To NombreResults
        name = "Results" & j
        Sheets("CondensedResults").Cells(3, col + 1) = "Pressure Ratio"
        Sheets("CondensedResults").Cells(3, col + 2) = "Efficiency"
        Sheets("CondensedResults").Cells(3, col + 3) = "Power"
        Sheets("CondensedResults").Cells(3, col + 6) = "Cost"
        Sheets("CondensedResults").Cells(3, col + 4) = "Fuel Flow"
        Sheets("CondensedResults").Cells(3, col + 5) = "Fuel Cost"
        
        Sheets("CondensedResults").Cells(3 + j, col + 1) = Sheets("CondensedResults").Cells(3 + j, 3)
        Sheets("CondensedResults").Cells(3 + j, col + 2) = Sheets("CondensedResults").Cells(3 + j, 4) + Sheets("CondensedResults").Cells(3 + j, col - 9 + 3) - Sheets("CondensedResults").Cells(3 + j, col - 9 + 3) * Sheets("CondensedResults").Cells(3 + j, 4)
        Sheets("CondensedResults").Cells(3 + j, col + 3) = Sheets("CondensedResults").Cells(3 + j, 5) + Sheets("CondensedResults").Cells(3 + j, col - 9 + 4)
        Sheets("CondensedResults").Cells(3 + j, col + 6) = Sheets("CondensedResults").Cells(3 + j, 8) + Sheets("CondensedResults").Cells(3 + j, col - 9 + 7) + Sheets("CondensedResults").Cells(3 + j, col + 5)
        Sheets("CondensedResults").Cells(3 + j, col + 4) = CDbl(Sheets("CondensedResults").Cells(3 + j, 7) + Sheets("CondensedResults").Cells(3 + j, col - 9 + 6))
        Sheets("CondensedResults").Cells(3 + j, col + 5) = CDbl(14600) * CDbl(4) * CDbl(Sheets("CondensedResults").Cells(3 + j, col + 4))
        
        'Input of Fuel cost
        
        
Next
Dim EffE
Dim CostE
EffE = Application.Evaluate("=linest(u4:u9,t4:t9^{1,2,3,4})")
CostE = Application.Evaluate("=linest(y4:y9,t4:t9^{1})")
'MsgBox "Equation is y=" & Format(X(1), "0.00000000") & "x3+" & Format(X(2), "0.0000000") & "x2+" & Format(X(3), "0.0000000") & "x+" & Format(X(4), "0.0000000")
 I = Sheets("CondensedResults").Range("t4")
 j = 12
 MaxPR = 0
 MaxEff = 0
 CostOpti = 0
 While I < Sheets("CondensedResults").Range("t9")
    
    If Eff(I, EffE) > Eff(I - 1, EffE) And Cost(I, CostE) < Sheets("Components").Range("C37") Then
        MaxEff = Eff(I, EffE)
        MaxPR = I
        CostOpti = Cost(I, CostE)
    End If
    j = j + 1
    I = I + 1
    
 Wend

Sheets("CondensedResults").Range("T13") = "MaxPR"
Sheets("CondensedResults").Range("U13") = "MaxEFF"
Sheets("CondensedResults").Range("V13") = "CostOpti"
Sheets("CondensedResults").Range("T14") = MaxPR
Sheets("CondensedResults").Range("U14") = MaxEff
Sheets("CondensedResults").Range("V14") = CostOpti
Dim oGraph As Shape
Dim oSerie As Series
Dim oAxes As Axes
Dim oTrendline As Trendline
Dim strEquation As String

ligne = Sheets("CondensedResults").Range("C3").End(xlDown).Row

Set oGraph = Sheets("CondensedResults").Shapes.AddChart

With oGraph.Chart
    .ChartType = xlXYScatter
    .HasTitle = True
    .ChartTitle.Text = "Effect of Pressure Ratio on Efficiency for fixed power"
    .SetElement msoElementPrimaryCategoryAxisShow
    .SetElement msoElementPrimaryCategoryAxisTitleHorizontal
    .SetElement msoElementPrimaryValueAxisShow
    .SetElement msoElementPrimaryValueAxisTitleHorizontal
    
'    Set oAxes = .Axes(xlCategory)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Pressure Ratio"
'    EndWith
'
'    Set oAxes = .Axes(xlValue)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Efficiency"
'
'    EndWith
        
    Set oSerie = .SeriesCollection.NewSeries
    With oSerie
        .XValues = Sheets("CondensedResults").Range("C4:C" & ligne)
        If Sheets("Components").Range("C32") = "Combined Cycle" Then
            .Values = Sheets("CondensedResults").Range("U4:U" & ligne)
        Else
            .Values = Sheets("CondensedResults").Range("D4:D" & ligne)
        End If
        .Trendlines.Add

        Set oTrendline = .Trendlines(1)

        With oTrendline
        
            .Type = xlPolynomial
            .Order = 3
            .DisplayEquation = True
            .DisplayRSquared = False
            .DataLabel.NumberFormat = "0.0000E+00"

        End With
        
    End With
   
   'Sheets("CondensedResults").Range("D12") = oSerie.Trendlines(1).DataLabel.
End With


Set oGraph = Sheets("CondensedResults").Shapes.AddChart

With oGraph.Chart
    .ChartType = xlXYScatter
    .HasTitle = True
    .ChartTitle.Text = "Effect of Pressure Ratio on Feed Flow for fixed power"
    .SetElement msoElementPrimaryCategoryAxisShow
    .SetElement msoElementPrimaryCategoryAxisTitleHorizontal
    .SetElement msoElementPrimaryValueAxisShow
    .SetElement msoElementPrimaryValueAxisTitleHorizontal
    
'    Set oAxes = .Axes(xlCategory)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Pressure Ratio"
'    EndWith
'
'    Set oAxes = .Axes(xlValue)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Efficiency"
'
'    EndWith
        
    Set oSerie = .SeriesCollection.NewSeries
    With oSerie
        .XValues = Sheets("CondensedResults").Range("C4:C" & ligne)
       .Values = Sheets("CondensedResults").Range("F4:F" & ligne)
    End With
End With


Set oGraph = Sheets("CondensedResults").Shapes.AddChart

With oGraph.Chart
    .ChartType = xlXYScatter
    .HasTitle = True
    .ChartTitle.Text = "Effect of Pressure Ratio on Fuel Flow for fixed power"
    .SetElement msoElementPrimaryCategoryAxisShow
    .SetElement msoElementPrimaryCategoryAxisTitleHorizontal
    .SetElement msoElementPrimaryValueAxisShow
    .SetElement msoElementPrimaryValueAxisTitleHorizontal
    
'    Set oAxes = .Axes(xlCategory)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Pressure Ratio"
'    EndWith
'
'    Set oAxes = .Axes(xlValue)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Efficiency"
'
'    EndWith
        
    Set oSerie = .SeriesCollection.NewSeries
    With oSerie
        .XValues = Sheets("CondensedResults").Range("C4:C" & ligne)
       .Values = Sheets("CondensedResults").Range("W4:W" & ligne)
    End With
End With

Set oGraph = Sheets("CondensedResults").Shapes.AddChart

With oGraph.Chart
    .ChartType = xlXYScatter
    .HasTitle = True
    .ChartTitle.Text = "Effect of Pressure Ratio on Cost for fixed power"
    .SetElement msoElementPrimaryCategoryAxisShow
    .SetElement msoElementPrimaryCategoryAxisTitleHorizontal
    .SetElement msoElementPrimaryValueAxisShow
    .SetElement msoElementPrimaryValueAxisTitleHorizontal
    
'    Set oAxes = .Axes(xlCategory)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Pressure Ratio"
'    EndWith
'
'    Set oAxes = .Axes(xlValue)
'    With oAxes
'        .HasTitle = True
'        .AxisTitle.text = "Efficiency"
'
'    EndWith
        
    Set oSerie = .SeriesCollection.NewSeries
    With oSerie
        .XValues = Sheets("CondensedResults").Range("T4:T" & ligne)
        If Sheets("Components").Range("C32") = "Combined Cycle" Then
            .Values = Sheets("CondensedResults").Range("Y4:Y" & ligne)
        Else
            .Values = Sheets("CondensedResults").Range("H4:H" & ligne)
        End If
    End With
End With



With Sheets("CondensedResults")
    For I = 1 To 26
        .Columns(I).AutoFit
    Next
End With


End Sub

Function ResultsCycle(myCompoCollec As Collection, strcase As String, myCycleCollec As Collection, Datas() As Variant, cycleName As String, NumResults As Integer) As String

'//////////////////////////Results extraction////////////////////////////////////


NombreResults = ActiveWorkbook.Sheets.count - 5
Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet


power = 0
RCost = 0
FuelCost = 0
    Ij = Sheets(cycleName).Range("E40")
    Ny = Sheets(cycleName).Range("E41")
    PhiM = Sheets(cycleName).Range("E42")
    OH = Sheets(cycleName).Range("C43")
    
'Writing the results

Dim oComponent As cComponent

If Sheets("Results").Range("A5") = "" Then
    ligne = 5
Else
    ligne = Sheets("Results").Range("A4").End(xlDown).Row + 1
End If

ligneStart = ligne - 1
'
For Each oCycle In myCycleCollec

    
    Sheets("Results").Cells(ligne, 1) = cycleName
    Sheets("Results").Cells(ligne, 2) = "Results n°" & NumResults
    Sheets("Results").Cells(ligne, 3) = oCycle.name
    Sheets("Results").Cells(ligne, 4) = oCycle.CType
    Sheets("Results").Cells(ligne, 5) = oCycle.power
    Sheets("Results").Cells(ligne, 6) = oCycle.Efficiency
    Sheets("Results").Cells(ligne, 7) = oCycle.PressureRatio
    Sheets("Results").Cells(ligne, 8) = oCycle.FiringTemp
    'Sheets("Results").Cells(ligne, 8) = oCycle.HeatRate
    Sheets("Results").Cells(ligne, 9) = oCycle.FeedFlow
    Sheets("Results").Cells(ligne, 10) = oCycle.FuelFlow
    
    ligne3 = Sheets(cycleName).Range("C82").End(xlDown).Row + 7
    If Sheets("Results").Cells(ligne, 10) <> 0 Then
        If oCycle.CType = "Brayton" Or oCycle.CType = "Regeneration Brayton" Or oCycle.CType = "Reheat Brayton" Then
           Sheets("Results").Cells(ligne, 11) = 3412 / oCycle.HeatRate * Sheets(cycleName).Cells(ligne3, 4) / 1000 / 1000
           FC = 3412 / oCycle.HeatRate * Sheets(cycleName).Cells(ligne3, 4) / 1000 / 1000
        ElseIf oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine" Then
            If Sheets("Results").Cells(ligne, 10) <> 0 Then
                Sheets("Results").Cells(ligne, 11) = 3412 / oCycle.HeatRate * Sheets(cycleName).Cells(ligne3, 5) / 1000 / 1000
                FC = 3412 / oCycle.HeatRate * Sheets(cycleName).Cells(ligne3, 5) / 1000 / 1000
            End If
        End If
    Else
        Sheets("Results").Cells(ligne, 11) = 0
        FC = 0
    End If
    
    Sheets("Results").Cells(ligne, 12) = oCycle.Cost  ' oCycle.costKWH


    
    If oCycle.CType = "Brayton" Or oCycle.CType = "Regeneration Brayton" Or oCycle.CType = "Reheat Brayton" Then
        EffB = oCycle.Efficiency
        PR = oCycle.PressureRatio
    ElseIf oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine" Then
        EffR = oCycle.Efficiency
        If Sheets(cycleName).Range("C37") = "Rankine" Then
            PR = oCycle.PressureRatio
        End If
    ElseIf oCycle.CType = "Solar" Then
        EffS = oCycle.Efficiency
    End If
    
    power = power + oCycle.power
    RCost = RCost + oCycle.Cost
    FuelCost = FuelCost + FC
        
    ligne = ligne + 1
Next


If Sheets(cycleName).Range("C37") = "Combined Cycle" Then

    Sheets("Results").Cells(ligne - 1 - NumResults, 15) = EffB + EffR - EffB * EffR
    Sheets("Results").Cells(ligne - 1 - NumResults, 16) = power
    Sheets("Results").Cells(ligne - 1 - NumResults, 14) = PR
    Sheets("Results").Cells(ligne - 1 - NumResults, 17) = FuelCost
    Sheets("Results").Cells(ligne - 1 - NumResults, 18) = RCost * PhiM * (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1))
    Sheets("Results").Cells(ligne - 1 - NumResults, 19) = Sheets("Results").Cells(ligne - 1 - NumResults, 18) / power / OH
    Sheets("Results").Cells(ligne - 1 - NumResults, 20) = Sheets("Results").Cells(ligne - 1 - NumResults, 19) + FuelCost

 Else
    If Sheets(cycleName).Range("C37") = "Brayton" Or Sheets(cycleName).Range("C37") = "Regeneration Brayton" Then
        Sheets("Results").Cells(ligne - 1, 15) = EffB
    ElseIf Sheets(cycleName).Range("C37") = "Rankine" Then
        Sheets("Results").Cells(ligne - 1, 15) = EffR

    End If
    

Sheets("Results").Cells(ligne - 1, 16) = power
Sheets("Results").Cells(ligne - 1, 14) = PR
Sheets("Results").Cells(ligne - 1, 17) = FuelCost
Sheets("Results").Cells(ligne - 1, 18) = RCost * PhiM * (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1)) ' oCycle.costKWH
Sheets("Results").Cells(ligne - 1, 19) = Sheets("Results").Cells(ligne - 1, 18) / power / OH
Sheets("Results").Cells(ligne - 1, 20) = Sheets("Results").Cells(ligne - 1, 19) + FuelCost
Sheets("Results").Cells(ligne - 1, 25) = FuelCost * power / (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1)) * OH
 End If
 
 
    
 
'Resize of the sheet
With Sheets("Results")
    For I = 1 To 16
        .Columns(I).AutoFit
    Next
End With

ResultsCycle = "youououo"

End Function

