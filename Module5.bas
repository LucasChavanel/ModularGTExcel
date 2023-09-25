Attribute VB_Name = "Module5"
Function CompDesign(Datas() As Variant, strcase As String, cycleName As String) As Collection


Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Set CompDesign = New Collection
Dim OComp As cCompressor
Dim l As Integer, k As Integer
Dim NumberCompressor As Integer
NumberCompressor = 0
Dim X As Integer
X = Sheets(cycleName).Range("A10").End(xlToRight).column

Dim v As Integer
Dim last_row As Integer, last_column As Integer
last_row = Sheets(cycleName).Range("D58").End(xlDown).Row
last_column = Sheets(cycleName).Range("C58").End(xlToRight).column
Dim CompSpec()


ReDim CompSpec(last_row - 57, last_column - 3)
For k = 58 To last_row
    For j = 4 To last_column
        CompSpec(k - 58, j - 4) = Sheets(cycleName).Cells(k, j)
     Next
Next

'We add each of our component in different collections : Compressor, Turbine and Combustion Chamber
 For l = 0 To X - 2
        If Datas(0, l) = "Compressor" Then
        
            'We increment the number of compressor, it will help us navigate in the the collection
            NumberCompressor = NumberCompressor + 1
            Set COmp = hyFlwSht.Operations.Item(Datas(1, l))
            Set OComp = New cCompressor
            OComp.index = NumberCompressor
            CompDesign.Add Item:=OComp
            'Debug.Print NumberCompressor
            
            'We add the Aspen Datas of the compressor to our VBA object
            CompDesign.Item(NumberCompressor).Tempo = (COmp.FeedTemperature + 273.15) * 1.8 'Conversion of SI units to Imperial units for the method
            CompDesign.Item(NumberCompressor).Tempd = (COmp.ProductTemperature + 273.15) * 1.8
            CompDesign.Item(NumberCompressor).PressureRatio = COmp.ProductPressure / COmp.FeedPressure
            CompDesign.Item(NumberCompressor).IsenEfficiency = Datas(13, l)
            CompDesign.Item(NumberCompressor).CompName = COmp.name
            CompDesign.Item(NumberCompressor).Press = COmp.FeedPressure
            CompDesign.Item(NumberCompressor).RSpeed = Datas(14, l)

            CompDesign.Item(NumberCompressor).Z = COmp.FeedStream.Compressibility
            CompDesign.Item(NumberCompressor).GasC = 8.314 / COmp.FeedStream.MolecularWeight * 185.915

            
            CompDesign.Item(NumberCompressor).MassFlow = COmp.MassFlow

             
            CompDesign.Item(NumberCompressor).AssComp = l
            
            For v = 0 To last_column - 3
                If CompSpec(0, v) = Datas(1, l) Then
                    CompDesign.Item(NumberCompressor).PhiAve = CompSpec(1, v)
                    
                    CompDesign.Item(NumberCompressor).MaxTipSpeed = CompSpec(3, v) * 3.28084
                End If
            Next
            
            

            
            '-----------Compressor Stage Design : Cf Guide --------------------------------------
            

            Dim StageMini As Double
           ' Debug.Print "Compressor : ", Datas(1, l)
            StageMini = WorksheetFunction.RoundUp(CompDesign.Item(NumberCompressor).Hp / CompDesign.Item(NumberCompressor).HeadPerStage, 0)
            'Debug.Print "Number Of Stages", StageMini
            CompDesign.Item(NumberCompressor).NumberStages = StageMini
                
           ' Debug.Print "HReq", "HperStage ", "Max TipSpeed", " TipSpeed ", " N ", " D ", "Z", "GasC", "To"
            'Debug.Print CompDesign.Item(NumberCompressor).Hp, CompDesign.Item(NumberCompressor).HeadPerStage, CompDesign.Item(NumberCompressor).MaxTipSpeed, CompDesign.Item(NumberCompressor).TipS, CompDesign.Item(NumberCompressor).RSpeed, CompDesign.Item(NumberCompressor).Diameter, CompDesign.Item(NumberCompressor).Z, CompDesign.Item(NumberCompressor).GasC, CompDesign.Item(NumberCompressor).Tempo
           ' Dim nsurn As Double
            'nsurn = 0.286 / Datas(13, l)
            'Dim Pr As Double
           ' Pr = ((CompDesign.Item(NumberCompressor).HeadPerStage * StageMini * nsurn / CompDesign.Item(NumberCompressor).Z / CompDesign.Item(NumberCompressor).GasC / CompDesign.Item(NumberCompressor).Tempo + 1) ^ (1 / nsurn))
            
           'MsgBox Pr

        End If
        
 Next l



    


End Function

Function TurbDesign(Datas() As Variant, strcase As String, cycleName As String) As Collection

Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim OTurb As cTurbine
Dim NumberTurbine As Integer
Set TurbDesign = New Collection
NumberTurbine = 0
Dim l As Integer, X As Integer, j As Integer, k As Integer
Dim pI As Double, Ti As Double, Pri As Double, DT As Double

Dim TurbSpec As Variant
Dim last_row As Integer, last_column As Integer

If Sheets(cycleName).Range("D49") <> "" Then

last_row = Sheets(cycleName).Range("D49").End(xlDown).Row
last_column = Sheets(cycleName).Range("C49").End(xlToRight).column


'We create a database with the Components table, we will use this database to input components into Aspen

Dim v As Integer
ReDim TurbSpec(last_row - 18, last_column - 3)
For k = 0 To last_row - 48
    For j = 0 To last_column - 3
        TurbSpec(k, j) = Sheets(cycleName).Cells(k + 49, j + 4)
     Next
Next

X = Sheets(cycleName).Range("A10").End(xlToRight).column
For l = 0 To X - 2
    If Datas(0, l) = "Gas Turbine" Then
           ' Debug.Print Datas(1, l)
            'We increment the number of turbines, it will help us navigate in the the collection
            NumberTurbine = NumberTurbine + 1
            
            Set Turb = hyFlwSht.Operations.Item(Datas(1, l))
            Set OTurb = New cTurbine
            OTurb.index = NumberTurbine
            TurbDesign.Add Item:=OTurb
            
            'We add our initial parameters to the VBA object
            TurbDesign.Item(NumberTurbine).AssTurb = l
            TurbDesign.Item(NumberTurbine).PressIn = Turb.FeedPressure
            TurbDesign.Item(NumberTurbine).TempIn = Turb.FeedTemperature + 273.15
            TurbDesign.Item(NumberTurbine).MassFlow = Turb.MassFlow
            TurbDesign.Item(NumberTurbine).Efficiency = Turb.ExpAdiabaticEff
            TurbDesign.Item(NumberTurbine).gamma = (Turb.FeedStream.CpCv + Turb.ProductStream.CpCv) / 2
            TurbDesign.Item(NumberTurbine).Cp = (Turb.FeedStream.MassHeatCapacity + Turb.ProductStream.MassHeatCapacity) / 2
            
            For v = 0 To last_column - 3
                If TurbSpec(0, v) = Datas(1, l) Then
                    TurbDesign.Item(NumberTurbine).NumberStage = TurbSpec(6, v)
                    TurbDesign.Item(NumberTurbine).FlowCoeff = TurbSpec(1, v)
                    TurbDesign.Item(NumberTurbine).LoadingCoeff = TurbSpec(2, v)
                    TurbDesign.Item(NumberTurbine).HTRatio = TurbSpec(3, v)
                    TurbDesign.Item(NumberTurbine).ReactionDegree = TurbSpec(4, v)
                    TurbDesign.Item(NumberTurbine).RotatingSpeed = Datas(14, l)
                    TurbDesign.Item(NumberTurbine).PressureRatio = 1 / Datas(8, l)
                    TurbDesign.Item(NumberTurbine).MaxTipSpeed = TurbSpec(5, v)
                End If
            Next
            pI = TurbDesign.Item(NumberTurbine).PressIn
            Ti = TurbDesign.Item(NumberTurbine).TempIn
           
            DT = TurbDesign.Item(NumberTurbine).DTstage
            
            'Debug.Print DT
            'Debug.Print "Number of stages", "Temperature", "Pressure", "Pri"
            'Debug.Print 0, Ti, Pi
            
            For j = 1 To TurbDesign.Item(NumberTurbine).NumberStage
            
                Pri = (1 - DT / Ti) ^ (TurbDesign.Item(NumberTurbine).gamma / (1 - TurbDesign.Item(NumberTurbine).gamma) / TurbDesign.Item(NumberTurbine).PolytropicEff)
                Ti = Ti - DT
                pI = pI / Pri
               ' Debug.Print j, Ti, Pi, Pri
            Next
            
           ' Debug.Print "Pressure Ratio : ", Pi / TurbDesign.Item(NumberTurbine).PressIn
            TurbDesign.Item(NumberTurbine).power = TurbDesign.Item(NumberTurbine).Cp * (TurbDesign.Item(NumberTurbine).TempIn - Ti) * TurbDesign.Item(NumberTurbine).MassFlow
            'Debug.Print "Power : ", TurbDesign.Item(NumberTurbine).Power
           ' Debug.Print "Tip Speed : ", TurbDesign.Item(NumberTurbine).Ut
            'Debug.Print "Mean Diameter : ", TurbDesign.Item(NumberTurbine).Dm
            TurbDesign.Item(NumberTurbine).PressOut = pI
            
            While TurbDesign.Item(NumberTurbine).Ut > TurbDesign.Item(NumberTurbine).MaxTipSpeed
             
                TurbDesign.Item(NumberTurbine).NumberStage = TurbDesign.Item(NumberTurbine).NumberStage + 1
                pI = TurbDesign.Item(NumberTurbine).PressIn
                 Ti = TurbDesign.Item(NumberTurbine).TempIn
                
                 DT = TurbDesign.Item(NumberTurbine).DTstage
                 
                 'Debug.Print DT
                 'Debug.Print "Number of stages", "Temperature", "Pressure", "Pri"
                ' Debug.Print 0, Ti, Pi
                 
                 For j = 1 To TurbDesign.Item(NumberTurbine).NumberStage
                 
                     Pri = (1 - DT / Ti) ^ (TurbDesign.Item(NumberTurbine).gamma / (1 - TurbDesign.Item(NumberTurbine).gamma) / TurbDesign.Item(NumberTurbine).PolytropicEff)
                     Ti = Ti - DT
                     pI = pI / Pri
                  '   Debug.Print j, Ti, Pi, Pri
                 Next
                ' Debug.Print "Pressure Ratio : ", Pi / TurbDesign.Item(NumberTurbine).PressIn
                 TurbDesign.Item(NumberTurbine).power = TurbDesign.Item(NumberTurbine).Cp * (TurbDesign.Item(NumberTurbine).TempIn - Ti) * TurbDesign.Item(NumberTurbine).MassFlow
                 'Debug.Print "Power : ", TurbDesign.Item(NumberTurbine).power
                ' Debug.Print "Tip Speed : ", TurbDesign.Item(NumberTurbine).Ut
                 'Debug.Print "Mean Diameter : ", TurbDesign.Item(NumberTurbine).Dm
                 TurbDesign.Item(NumberTurbine).PressOut = pI
            Wend
        End If
    Next l
End If
End Function





Function ApproxGTPower(myCompoCollec As Collection, myCycleCollec As Collection, Datas() As Variant, strcase As String, cycleName As String) As String
'With this method, we will modify our parameters to approximate the wanted power to the best
'Two scenarios :
'If Power too low : We will increase the firing temperature until either the power or the temperature limit is reached


'//////////////////////////GT Approximation to the right power///////////////////
Dim PowerWanted As Double
Dim GTPOwer As Double
Dim FiringTemp As Double
Dim TempMax As Double
Dim oComponent As cComponent
Dim oCycle As cCycle
Dim oCycle2 As cCycle
Dim simCase As SimulationCase
Dim myCycleCollec2 As Collection
Dim StreamName As String
Dim Continue As Boolean
Dim Ecart As Double
Dim count As Integer
count = 0
Dim myTurb As Collection
Dim myComp As Collection
Continue = True
Dim hyFluidPkg As HYSYS.FluidPackage
Dim X As Integer
X = Sheets(cycleName).Range("A10").End(xlToRight).column
Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim GTType As String
GTType = Sheets(cycleName).Range("C37")

If GTType = "Brayton" Or GTType = "Reheat Brayton" Or GTType = "Regeneration Brayton" Or GTType = "Rankine" Or GTType = "ORC Rankine" Then

    GTPOwer = 0
        
    For Each oCycle In myCycleCollec
        GTPOwer = GTPOwer + oCycle.power
    
       
       
        Do While Abs((GTPOwer - Sheets(cycleName).Range("C33")) / Sheets(cycleName).Range("C33")) > 0.01 And Continue = True
         Ecart = GTPOwer / Sheets(cycleName).Range("C33")
            GTPOwer = 0
            
            StreamName = oCycle.StreamPilot
            Set Stream = hyFlwSht.MaterialStreams.Item(StreamName)
            
            count = count + 1
            If count > 20 Then
                Continue = False
            End If
            
            Stream.MassFlowValue = Stream.MassFlow * (1 / Ecart)
            
            For Each oComponent In myCompoCollec
                If oComponent.CompType = "Combustion Chamber" And oComponent.cycleName = oCycle.name Then
                    Set CC = hyFlwSht.Operations.Item(oComponent.CompName)
                    For I = 0 To X - 2
                        If Datas(1, I) = oComponent.CompName Then

                            FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
                            If FuelFlow < 0 Then
                                MsgBox "FuelFlow is negative, please correct your gas turbine"
                                ApproxGTPower = "aye"
                                Exit Function
                            Else
                                CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
                            End If
'                            While CC.AttachedProducts.Item(1).Temperature < (Datas(12, I) - 273.15) * 0.995
'                                CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 1.005
'                            Wend
'                            While CC.AttachedProducts.Item(1).Temperature > (Datas(12, I) - 273.15) * 1.005
'                                CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 0.995
'                            Wend
                        End If
                    Next
                End If
            Next
    
    
            Set myComp = CompDesign(Datas(), strcase, cycleName)
            Set myTurb = TurbDesign(Datas(), strcase, cycleName)
            Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
            Set myCycleCollec2 = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
            
            GTPOwer = 0
            For Each oCycle2 In myCycleCollec2
                GTPOwer = GTPOwer + oCycle2.power
            Next
            
            Debug.Print GTPOwer
            If Abs((GTPOwer - Sheets(cycleName).Range("C33")) / Sheets(cycleName).Range("C33")) < 0.01 Then
                Exit For
            End If
             
             
             Set myCycleCollec = myCycleCollec2
        Loop
               

    Next

ElseIf GTType = "Combined Cycle" Then
    GTPOwer = 0
        

    For Each oCycle In myCycleCollec
        If oCycle.name = "Rankine1" Then
            CPower = oCycle.power
            While CPower > Sheets(cycleName).Range("C33") / 2
                
                    hyFlwSht.MaterialStreams.Item("Mix1").MassFlow = hyFlwSht.MaterialStreams.Item("Mix1").MassFlow / 2
                    
                    Set myComp = CompDesign(Datas(), strcase, cycleName)
                    Set myTurb = TurbDesign(Datas(), strcase, cycleName)
                    Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
                    Set myCycleCollec2 = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
                    For Each oCycle2 In myCycleCollec2
                        If oCycle2.name = "Rankine1" Then
                            CPower = oCycle2.power
                        End If
                    Next
               
            Wend
        End If
    Next
     Set myComp = CompDesign(Datas(), strcase, cycleName)
    Set myTurb = TurbDesign(Datas(), strcase, cycleName)
    Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
    Set myCycleCollec2 = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
    For Each oCycle2 In myCycleCollec2
        
        GTPOwer = GTPOwer + oCycle2.power
    Next
       
    Do While Abs((GTPOwer - Sheets(cycleName).Range("C33")) / Sheets(cycleName).Range("C33")) > 0.01 And Continue = True
        Ecart = GTPOwer / Sheets(cycleName).Range("C33")
        GTPOwer = 0
        
        For Each oCycle2 In myCycleCollec2
            If oCycle2.CType = "Brayton" Or oCycle2.CType = "Reheat Brayton" Or oCycle2.CType = "Regeneration Brayton" Then
                StreamName = oCycle2.StreamPilot
                Set Stream = hyFlwSht.MaterialStreams.Item(StreamName)
                
                count = count + 1
                If count > 20 Then
                    Continue = False
                End If
'                If Ecart > 0.1 Then
                    Stream.MassFlowValue = Stream.MassFlow * (1 / Ecart)
'                Else
'                    Stream.MassFlowValue = Stream.MassFlow * 1.01
'                End If
                
                For Each oComponent In myCompoCollec
                    If oComponent.CompType = "Combustion Chamber" And oComponent.cycleName = oCycle2.name Then
                        Set CC = hyFlwSht.Operations.Item(oComponent.CompName)
                        For I = 0 To X - 2
                            If Datas(1, I) = oComponent.CompName Then
                                FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
                                If FuelFlow < 0 Then
                                    MsgBox "FuelFlow is negative, please correct your gas turbine"
                                    ApproxGTPower = "aye"
                                    Exit Function
                                Else
                                    CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
                                End If
                                
'                            While CC.VapourProduct.Temperature > 1.05 * (Datas(12, I) - 273.15)
'                                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 0.99
'                            Wend
'
'                            While CC.VapourProduct.Temperature < 0.95 * (Datas(12, I) - 273.15)
'                                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 1.01
'                            Wend
                            End If
                        Next
                    End If
                Next
        
        
                Set myComp = CompDesign(Datas(), strcase, cycleName)
                Set myTurb = TurbDesign(Datas(), strcase, cycleName)
                Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
                Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
                
                GTPOwer = 0
                For Each oCycle In myCycleCollec
                    GTPOwer = GTPOwer + oCycle.power
                Next
           
                If Abs((GTPOwer - Sheets(cycleName).Range("C33")) / Sheets(cycleName).Range("C33")) < 0.01 Then
                    Exit For
                End If
            End If
         Next
         
         Set myCycleCollec = myCycleCollec2
    Loop
           

               

ElseIf GTType = "More Than Combined Cycle" Then

GTPOwer = 0
        
    For Each oCycle In myCycleCollec
        If oCycle.CType = Sheets(cycleName).Range("C37") Then
            GTPOwer = GTPOwer + oCycle.power
        
           
           
            Do While Abs((GTPOwer - Sheets(cycleName).Range("C33")) / Sheets(cycleName).Range("C33")) > 0.01 And Continue = True
                GTPOwer = 0
                
                StreamName = oCycle.StreamPilot
                Set Stream = hyFlwSht.MaterialStreams.Item(StreamName)
                
                count = count + 1
                If count > 20 Then
                    Continue = False
                End If
                
                Stream.MassFlowValue = Stream.MassFlow * 1.001
                Debug.Print count, Stream.MassFlow
                For Each oComponent In myCompoCollec
                    If oComponent.CompType = "Combustion Chamber" And oComponent.cycleName = oCycle.name Then
                        Set CC = hyFlwSht.Operations.Item(oComponent.CompName)
                        For I = 0 To X - 2
                            If Datas(1, I) = oComponent.CompName Then
                                FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
                                If FuelFlow < 0 Then
                                    MsgBox "FuelFlow is negative, please correct your gas turbine"
                                    ApproxGTPower = "aye"
                                    Exit Function
                                Else
                                    CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
                                End If
                                
'                            While CC.VapourProduct.Temperature > 1.05 * (Datas(12, I) - 273.15)
'                                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 0.99
'                            Wend
'
'                            While CC.VapourProduct.Temperature < 0.95 * (Datas(12, I) - 273.15)
'                                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 1.01
'                            Wend
'                                                            While CC.AttachedProducts.Item(1).Temperature < (Datas(12, I) - 273.15)
'                                CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 1.001
'                            Wend
'                            While CC.AttachedProducts.Item(1).Temperature > (Datas(12, I) - 273.15) * 1.001
'                                CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 0.999
'                            Wend
'
                            End If
                        Next
                    End If
                Next
        
                Aye = HXRecalibration(Datas(), strcase, cycleName)
                Set myComp = CompDesign(Datas(), strcase, cycleName)
                Set myTurb = TurbDesign(Datas(), strcase, cycleName)
                Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
                Set myCycleCollec2 = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
                
                GTPOwer = 0
                For Each oCycle2 In myCycleCollec2
                    GTPOwer = GTPOwer + oCycle2.power
                Next
                
                Debug.Print GTPOwer
                If GTPOwer > Sheets(cycleName).Range("C33") Then
                    Exit For
                End If
                 
                 
                 Set myCycleCollec = myCycleCollec2
            Loop
                   

        End If
    Next
End If


ApproxGTPower = "Yes"
'Add condition for multiple combined cycles





End Function



Function HXRecalibration(Datas() As Variant, strcase As String, cycleName As String) As String
'We recalibrate the HX



Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim X As Integer
Dim Y As Integer
X = UBound(Datas, 2) - LBound(Datas, 2) + 1
Y = UBound(Datas, 1) - LBound(Datas, 1) + 1
last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
Dim HXSpec As Variant
Dim v As Integer
ReDim HXSpec(2, last_column - 3)
For k = 0 To 1
    For j = 0 To last_column - 3
        HXSpec(k, j) = Sheets(cycleName).Cells(k + 73, j + 4)
     Next
Next
If cycleName <> "Solar Combined" And cycleName <> "Combined Cycle" Then
For I = 0 To X - 2
     If Datas(0, I) = "Heat Exchanger" Then
'
      Set HX = hyFlwSht.Operations.Item(Datas(1, I))
        last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
        For v = 0 To last_column - 4
            If Datas(1, I) = HXSpec(0, v) Then
                TypeHX = HXSpec(1, v)
            End If
        Next
        
        
        If HX.TubeSideProduct.Temperature = -32767 And (TypeHX = "Superheated steam" Or TypeHX = "Regeneration" Or TypeHX = "Reheat" Or TypeHX = "Heater") And HX.TubeSideProduct.Temperature > HX.ShellSideFeed.Temperature Then
            HX.TubeSideProduct.Temperature = HX.ShellSideFeed.Temperature * 0.95

        End If
        
        If (TypeHX = "Superheated steam" Or TypeHX = "Regeneration" Or TypeHX = "Reheat" Or TypeHX = "Heater") And HX.TubeSideFeed.Temperature > HX.ShellSideFeed.Temperature Then
            HX.TubeSideProduct.Temperature = HX.TubeSideFeed.Temperature
        End If
       
       If Datas(1, I) = "HXSolar" Then
            Set Solar = hyFlwSht.Operations.Item("Solar")
            While HX.TubeSideFeed.Temperature > HX.ShellSideProduct.Temperature
            
                HX.TubeSideProduct.Temperature = HX.TubeSideProduct.Temperature * 0.99
            Wend
            If HX.TubeSideFeed.Temperature > HX.TubeSideProduct.Temperature Then
                HX.TubeSideProduct.Temperature = HX.TubeSideFeed.Temperature
            End If
       End If
'        'Alert the user if the Output Temp is too low an fix it (10% of temperature increase)
'        If Datas(12, i) - 273.15 < HX.TubeSideFeed.TemperatureValue Then
'            HX.TubeSideProduct.TemperatureValue = 1.1 * HX.TubeSideFeed.TemperatureValue
'            If Sheets(cycleName).Range("C33") = "Design One Cycle" Then
'                MsgBox "Cold Output Temp < Input Temp in : " & HX.name
'                MsgBox "Correction made"
'            End If
'        'else we add the temp choosen by the user
'        Else
'
'            HX.TubeSideProduct.TemperatureValue = HX.ShellSideFeed.Temperature * 0.95
'        End If
'
'
'        'Alert the user if the Temperature cross is the HX wich mean an error in the design
'        If HX.ShellSideFeed.Temperature < HX.TubeSideProduct.Temperature Then
'            If Sheets(cycleName).Range("C33") = "Design One Cycle" Then
'                MsgBox "Temperature Cross in :" & HX.name
'                MsgBox "Please Correct your cycle"
'            End If
'        Else
'            'If the cycle is properly designed, we increase the number of shell tube (HOT Fluid) to get a proper HX
            While HX.FtFactor < 0.75 And HX.ShellSeriesValue < 10

                HX.ShellSeriesValue = HX.ShellSeriesValue + 1

           Wend
        End If
    If Datas(0, I) = "Cooler" Then
        Set Cool = hyFlwSht.Operations.Item(Datas(1, I))
        
        If Cool.FeedTemperatureValue = -32767 Then
             Cool.FeedTemperature.SetValue 288, "K"
            End If
        If Cool.FeedPressureValue = -32767 Then
             Cool.FeedPressure.SetValue 101, "kPa"
            End If
        Cool.ProductTemperature.SetValue Datas(12, I), "K"
        Cool.ProductPressure.SetValue (Cool.FeedPressureValue * (100 - Datas(9, I)) / 100)
        
    ElseIf Datas(0, I) = "Heater" Then
            Set Heat = hyFlwSht.Operations.Item(Datas(1, I))
            Heat.ProductTemperature.SetValue Datas(12, I), "K"
            Heat.ProductPressure.SetValue (Heat.FeedPressureValue * (1 - Datas(9, I) / 100))
    End If
    
     If Datas(0, I) = "Combustion Chamber" Then

         Set CC = hyFlwSht.Operations.Item(Datas(1, I))
    
        FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
        If FuelFlow < 0 Then
            MsgBox "FuelFlow is negative, please correct your gas turbine"
            HXRecalibration = strcase
            Exit Function
        Else
            CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
            While CC.VapourProduct.Temperature > 1.01 * (Datas(12, I) - 273.15)
                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 0.99
            Wend

            While CC.VapourProduct.Temperature < 0.99 * (Datas(12, I) - 273.15)
                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 1.01
            Wend
            
        End If
'        While CC.AttachedProducts.Item(1).Temperature < (Datas(12, I) - 273.15)
'            CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 1.001
'        Wend
'        While CC.AttachedProducts.Item(1).Temperature > (Datas(12, I) - 273.15) * 1.001
'            CC.Feeds.Item(1).MassFlow = CC.Feeds.Item(1).MassFlow * 0.999
'        Wend
     End If
     
     If Datas(0, I) = "Gas Turbine" Or Datas(0, I) = "Steam Turbine" Then
        Set Turb = hyFlwSht.Operations.Item(Datas(1, I))
        If Turb.FeedPressure = -32767 Then
            Turb.FeedPressureValue = 700
        End If
        If Turb.ProductPressure = -32767 Then
            Turb.ProductPressure = Turb.FeedPressure * Datas(8, I)
        End If
    End If
    
         
     If Datas(0, I) = "Compressor" Then
        Set COmp = hyFlwSht.Operations.Item(Datas(1, I))
        COmp.ProductPressure = COmp.FeedPressure * Datas(8, I)
      End If
         
     If Datas(0, I) = "Pump" Then
        Set Pump = hyFlwSht.Operations.Item(Datas(1, I))
        Pump.ProductPressure = Pump.FeedPressure * Datas(8, I)
    End If
 Next
 End If
 HXRecalibration = strcase

End Function

Function Modif_Turbine(Datas() As Variant, strcase As String, cycleName As String) As Collection

Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet
Dim X As Integer
Dim Y As Integer
X = UBound(Datas, 2) - LBound(Datas, 2) + 1
Dim HXSpec As Variant
Dim v As Integer
last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
ReDim HXSpec(2, last_column - 3)
For k = 0 To 1
    For j = 0 To last_column - 3
        HXSpec(k, j) = Sheets(cycleName).Cells(k + 73, j + 4)
     Next
Next

For I = 0 To X - 2
       
        
        '/////////////////////////////Construction of diffrrents components////////////////////////////
        
        'We dissociate different components
'        If Datas(0, I) = "Valve" Then
'                Set Valve = hyFlwSht.Operations.Item(Datas(1, I))
'                If cycleName = "Fired Rankine" Then
'                    Valve.ProductStream.PressureValue = Sheets(cycleName).Range("D82") * 1.5
'               End If
'        End If
        If Datas(0, I) = "Tank" Then
        
           
            Set Tank = hyFlwSht.Operations.Item(Datas(1, I))
            

            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
        
        If Datas(0, I) = "Flash" Then
        
            
            Set Tank = hyFlwSht.Operations.Item(Datas(1, I))

            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
        
        If Datas(0, I) = "Fired Heater" Then
          
            Set Fired = hyFlwSht.Operations.Item(Datas(1, I))
            Fired.ExcessAirPercentValue = Datas(9, I)
            Fired.CombustionEfficiencyValue = Datas(13, I)
            
            If Fired.RadInlet.Item(0).Temperature > Datas(12, I) - 273.15 Then
                Fired.AttachedProducts.Item(0).Temperature.SetValue Fired.RadInlet.Item(0).Temperature + 1
            Else
                Fired.AttachedProducts.Item(0).Temperature.SetValue Datas(12, I), "K"
            End If
            
            If Fired.CombustionProduct.Temperature < Fired.AttachedProducts.Item(0).Temperature Then
                MsgBox "Temperature Cross in Fired Heater : " & Datas(1, I) & ". Please Correct your cycle."
            End If
        End If
        
        If Datas(0, I) = "Compressor" Then
      
            Set COmp = hyFlwSht.Operations.Item(Datas(1, I))

            'We cant add the pressure ratio with vba, so we set the Product pressure
            COmp.ProductPressure.SetValue (COmp.FeedPressure * Datas(8, I))
            COmp.CompAdiabaticEff = Datas(13, I)
            
            
        End If
            
            'Same as the compressor
        If Datas(0, I) = "Gas Turbine" Then
        
            Set Turb = hyFlwSht.Operations.Item(Datas(1, I))
            
           
    
            Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            
            Turb.ExpAdiabaticEffValue = Datas(13, I)
        End If
        
        
        
        
        'Same as the compressor
        If Datas(0, I) = "Pump" Then
     
            Set Pump = hyFlwSht.Operations.Item(Datas(1, I))
            
            Pump.AdiabaticEfficiency = Datas(13, I)
            Pump.ProductPressure.SetValue (Pump.FeedPressureValue * Datas(8, I))
        End If
    
        
        'Same as Compressor but for temperature
        If Datas(0, I) = "Cooler" Then
        
        
        
            Set Cool = hyFlwSht.Operations.Item(Datas(1, I))
          
         
          
            Cool.ProductTemperature.SetValue Datas(12, I), "K"
            

            Cool.ProductPressure.SetValue (Cool.FeedPressureValue * (100 - Datas(9, I)) / 100)
        End If
    
        If Datas(0, I) = "Heater" Then
        
        
            Set Heat = hyFlwSht.Operations.Item(Datas(1, I))
            
      
            
            Heat.ProductTemperature.SetValue Datas(12, I), "K"
            
            Heat.ProductPressure.SetValue (Heat.FeedPressure * (100 - Datas(9, I)) / 100)
            
        End If
        

        If Datas(0, I) = "Heat Exchanger" Then
        
            Set HX = hyFlwSht.Operations.Item(Datas(1, I))
            last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
            For v = 0 To last_column - 4
                If Datas(1, I) = HXSpec(0, v) Then
                    TypeHX = HXSpec(1, v)
                End If
            Next
            
            If TypeHX = "Saturated Liquid" Then
                HX.TubeSideProduct.VapourFraction = 0
                
                
            ElseIf TypeHX = "Saturated Steam" Then
                HX.TubeSideProduct.VapourFraction = 1
                
            ElseIf TypeHX = "SuperHeated Steam" Or TypeHX = "Regeneration" Or TypeHX = "Reheat" Or TypeHX = "Heater" Then
                If Datas(12, I) - 273.15 < HX.TubeSideFeed.Temperature Then
                    HX.TubeSideProduct.Temperature = HX.TubeSideFeed.Temperature + 50
                Else
                    HX.TubeSideProduct.Temperature = Datas(12, I) - 273.15
                End If
                
                
             ElseIf TypeHX = "Regeneration" Then
                HX.TubeSideProduct.TemperatureValue = Datas(12, I) - 273.15
            End If

           If HX.ShellSideFeed.Pressure <> -32767 Then
                HX.ShellSidePressureDropValue = Datas(10, I) / 100 * HX.ShellSideFeed.Pressure
           Else
                HX.ShellSidePressureDropValue = 4
           End If
           
           If HX.TubeSideFeed.Pressure <> -32767 And HX.TubeSidePressureDrop = -32767 Then
                HX.TubeSidePressureDropValue = Datas(10, I) / 100 * HX.TubeSideFeed.Pressure
           ElseIf HX.TubeSidePressureDrop = -32767 Then
                HX.TubeSidePressureDropValue = 40
           End If
            
            
        End If
        
        If Datas(0, I) = "Combustion Chamber" Then
        
        
            NombreCC = NombreCC + 1
            
            Set CC = hyFlwSht.Operations.Item(Datas(1, I))
            

            

            
            
            
            'We cannot implement the Flaming temperature in Aspen so we need to pilot the Fuel Mass Flow
            
            'Formula to calculate the Fuel Mass Flow to add in the CC given by the Flaming Temperature wanted
            FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
            
            
            
            If FuelFlow < 0 Then
                MsgBox "FuelFlow is negative, please correct your gas turbine. The error probably come from Entry mass flow"
                Exit Function
            Else
                CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
            End If
            
            
            CC.PressureDrop = CC.Feeds.Item(0).Pressure * Datas(9, I) / 100
            While CC.VapourProduct.Temperature > 1.01 * (Datas(12, I) - 273.15)
                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 0.99
            Wend

            While CC.VapourProduct.Temperature < 0.99 * (Datas(12, I) - 273.15)
                CC.Feeds.Item(1).MassFlow.SetValue CC.Feeds.Item(1).MassFlow * 1.01
            Wend
  
        End If
    
   
    
        If Datas(0, I) = "Splitter" Then
   
    
    
            Set Split = hyFlwSht.Operations.Item(Datas(1, I))
        

        
        
            
            ratio = Split.SplitsValue
            ratio(0) = Datas(9, I)
            ratio(1) = Datas(10, I)
            
            Split.Splits.SetValues ratio, ""
        
        End If
        
 
 
    
        
                    'Same as the compressor
        If Datas(0, I) = "Steam Turbine" Then
            Datas(16, I) = "True"
            Set Turb = hyFlwSht.Operations.Item(Datas(1, I))
            
           
            Turb.ExpAdiabaticEffValue = Datas(13, I)

            If Turb.FeedPressure = -32767 Then
                Turb.FeedPressureValue = 500
            End If
            
            If Datas(8, I) <> 0 And Datas(1, I) <> "STurb4" Then
                a = Datas(1, I)
           ' If Turb.ProductPressure = -32767 Or Turb.ProductPressure = 154.3 Then
                Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            End If
            
            
        End If
        


        
        If Datas(0, I) = "Stream Saturator" Then
            Datas(16, I) = "True"
            

            
                Set Sat = hyFlwSht.Operations.Item(Datas(1, I))

                Sat.RelativeHumidity = Datas(13, I)
        
        End If
Next
    If cycleName = "Fired Rankine" Or cycleName = "Solar Fired Rankine" Then
        Set HX = hyFlwSht.Operations.Item("Boiler1")
        While HX.TubeSideFeed.Temperature > HX.ShellSideProduct.Temperature
            Set Stream = hyFlwSht.MaterialStreams.Item("Feed")
            Stream.MassFlow = Stream.MassFlow * 1.05
        Wend
    End If
    Dim myComp As Collection
    Set myComp = CompDesign(Datas(), strcase, cycleName)
    
    Dim myTurb As Collection
    Set myTurb = TurbDesign(Datas(), strcase, cycleName)
    
    Aye = HXRecalibration(Datas(), strcase, cycleName)
    Set Modif_Turbine = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
     
End Function
