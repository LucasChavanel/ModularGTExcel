Attribute VB_Name = "Module6"
Function StreamCreation(Datas() As Variant, strcase As String, cycleName As String) As String


Dim simCase As SimulationCase
Dim StreamA As ProcessStream
Dim hyFluidPkg As HYSYS.FluidPackage
Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet
'We add the first stream
Dim X As Integer
Dim Y As Integer
X = UBound(Datas, 2) - LBound(Datas, 2) + 1
    col1 = Sheets(cycleName).Range("B82").End(xlToRight).column
    ligne = Sheets(cycleName).Range("C82").End(xlDown).Row
    For I = 4 To col1
    
        
        Set StreamA = hyFlwSht.MaterialStreams.Add(Sheets(cycleName).Cells(85, I))
        With StreamA
            'Datas from the excel
            .Pressure.SetValue Sheets(cycleName).Cells(82, I).Value, "kPa"
            .Temperature.SetValue Sheets(cycleName).Cells(83, I).Value, "K"
            .MassFlow.SetValue Sheets(cycleName).Cells(84, I).Value, "kg/s"
            
            'We define the composition of the stream
            varComps = .ComponentMolarFraction.Values
            
            ligne = Sheets(cycleName).Cells(82, I).End(xlDown).Row - 1
            For j = 0 To ligne - 86
                varComps(j) = Sheets(cycleName).Cells(j + 86, I)
            Next
            .ComponentMolarFraction.Values = varComps
        End With
    
    Next
    
    Dim ligne1 As Integer, ligne2 As Integer
    ligne1 = Sheets(cycleName).Range("C82").End(xlDown).Row + 4
    ligne2 = Sheets(cycleName).Cells(ligne1, 3).End(xlDown).Row
    col1 = Sheets(cycleName).Cells(ligne1, 3).End(xlToRight).column
    
    For I = 4 To col1
            'We create the stream for the gas
             Dim Gas As ProcessStream
             Set Gas = hyFlwSht.MaterialStreams.Add(Sheets(cycleName).Cells(ligne1 + 2, I))
             With Gas
                'Datas from the excel
                .Pressure.SetValue Sheets(cycleName).Cells(ligne1, I).Value, "kPa"
                .Temperature.SetValue Sheets(cycleName).Cells(ligne1 + 1, I).Value, "K"
                '.MassFlow.SetValue 0.2, "kg/s" 'Security Value, tbd later on the program
                
                'We define the composition of the stream
                varComps = .ComponentMolarFraction.Values
                ligne = ligne2 - ligne1 - 4
                Dim v As Integer
                For v = 0 To ligne
                    varComps(v) = Sheets(cycleName).Cells(v + ligne1 + 4, I)
                Next
                
                .ComponentMolarFraction.Values = varComps
            End With
    
    Next
    
    simCase.Solver.CanSolve = True
    
    
    
    Dim Op As Operations
    
    For I = 0 To X - 2
    
        'We add the stream one to one
        Dim Stream As ProcessStream
        
        'Security if null value
        If Datas(2, I) <> 0 Then
        Set Stream = hyFlwSht.MaterialStreams.Add(Datas(2, I))
        End If
        
        'if the stream is already added, it automatically overrides the code
        If Datas(3, I) <> 0 Then
            Set Stream = hyFlwSht.MaterialStreams.Add(Datas(3, I))
        End If
        
        If Datas(4, I) <> 0 Then
            Set Stream = hyFlwSht.MaterialStreams.Add(Datas(4, I))
        End If
        
        If Datas(5, I) <> 0 Then
            Set Stream = hyFlwSht.MaterialStreams.Add(Datas(5, I))
        End If
        
    
    Next
    
    StreamCreation = "Done"

End Function

Function ComponentCreation(Datas() As Variant, strcase As String, cycleName As String) As String


Dim simCase As SimulationCase
Dim hyFluidPkg As HYSYS.FluidPackage
Dim hyBasis As HYSYS.BasisManager

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet
Set hyBasis = simCase.BasisManager
Set hyFluidPkg = hyBasis.FluidPackages.Item(0)

Dim colS As Integer
colS = Sheets(cycleName).Range("B82").End(xlToRight).column - 4

Dim X As Integer
X = UBound(Datas, 2) - LBound(Datas, 2) + 1

Dim FirstStream()

colS = Sheets(cycleName).Range("B82").End(xlToRight).column - 4
ReDim FirstStream(colS)
For I = 0 To colS
    
        FirstStream(I) = Sheets(cycleName).Cells(85, I + 4)
     
Next


Dim column As Integer


last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
'Sheets(cycleName).Range("28A").End(xlToRight).column
Dim HXSpec As Variant
Dim v As Integer
ReDim HXSpec(2, last_column - 3)
For k = 0 To 1
    For j = 0 To last_column - 3
        HXSpec(k, j) = Sheets(cycleName).Cells(k + 73, j + 4)
     Next
Next
Dim HXType As String

 For j = 0 To colS
        For I = 0 To X - 2
       
        
        '/////////////////////////////Construction of diffrrents components////////////////////////////
        
        'We dissociate different components
        If Datas(0, I) = "Valve" And FirstStream(j) = Datas(2, I) Then
        
            Datas(16, I) = "True"
            Set Valve = hyFlwSht.Operations.Add("valve1", "valveop")
            Valve.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Valve.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            If Valve.ProductStream.Pressure = -32767 Then
                Valve.ProductStream.PressureValue = Valve.FeedStream.Pressure * (1 - Datas(9, I) / 100)
           End If
           

        End If
        
        
        If Datas(0, I) = "Tank" And FirstStream(j) = Datas(2, I) Then
        
            Datas(16, I) = "True"
            Set Tank = hyFlwSht.Operations.Add(Datas(1, I), "FlashTank")
            
            Tank.SeparatorType = stTank
            Tank.Feeds.Add (Datas(2, I))
            Tank.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Tank.LiquidProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
        
        If Datas(0, I) = "Flash" And FirstStream(j) = Datas(2, I) Then
        
            Datas(16, I) = "True"
            Set Tank = hyFlwSht.Operations.Add(Datas(1, I), "FlashTank")
            Tank.Feeds.Add (Datas(2, I))
            Tank.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Tank.LiquidProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
        
        If Datas(0, I) = "Fired Heater" And FirstStream(j) = Datas(2, I) Then
            Datas(16, I) = "True"
            Set Fired = hyFlwSht.Operations.Add(Datas(1, I), "dynamicfiredheaterop")
            Fired.RadInlet.Add Datas(2, I)
            Fired.ExcessAirPercentValue = Datas(9, I)
            Fired.FuelsIn.Add Datas(3, I)
            Fired.CombustionEfficiencyValue = Datas(13, I)
            Fired.RadOutlet.Add Datas(4, I)
        
             Set Air = hyFlwSht.MaterialStreams.Add("Air" & Datas(1, I))
             With Air
            'Datas from the excel
            .Pressure.SetValue Sheets(cycleName).Range("C36").Value, "kPa"
            .Temperature.SetValue Sheets(cycleName).Range("C35").Value, "K"
            '.MassFlow.SetValue 0.2, "kg/s" 'Security Value, tbd later on the program
            
            'We define the composition of the stream
            varComps = .ComponentMolarFraction.Values
             ligne = Sheets(cycleName).Range("D83").End(xlDown).Row
            For coucou = 0 To ligne - 86
                If hyFluidPkg.Components.Item(coucou) = "Nitrogen" Then
                    varComps(coucou) = 0.79
                ElseIf hyFluidPkg.Components.Item(coucou) = "Oxygen" Then
                    varComps(coucou) = 0.21
                Else
                    varComps(coucou) = 0
                End If
            Next
            .ComponentMolarFraction.Values = varComps
            
            End With
            If Sheets(cycleName).Range("O32") <> "" Then
            
                Fired.RadOutlet.Item(0).VapourFraction = 1
                Fired.ConvInlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("O33")))
                Fired.ConvOutlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("P33")))
                Fired.ConvOutlet.Item(0).Temperature = Sheets(cycleName).Range("Q33") - 273.15
                Fired.EconOutlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("P34")))
                Fired.EconInlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("O34")))
                Fired.EconOutlet.Temperature = Sheets(cycleName).Range("Q34") - 273.15
            Else
                        
                If Fired.RadInlet.Item(0).Temperature > Datas(12, I) - 273.15 Then
                    Fired.AttachedProducts.Item(0).Temperature.SetValue Fired.RadInlet.Item(0).Temperature + 1
                Else
                    Fired.AttachedProducts.Item(0).Temperature.SetValue Datas(12, I), "K"
                End If
            End If
            Fired.BurnerFeed = Air
            Set Flare = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Fired.CombustionProduct = Flare

            
            If Fired.CombustionProduct.Temperature < Fired.AttachedProducts.Item(0).Temperature Then
                MsgBox "Temperature Cross in Fired Heater : " & Datas(1, I) & ". Please Correct your cycle."
            End If
        End If
        
        If Datas(0, I) = "Compressor" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
            Set COmp = hyFlwSht.Operations.Add(Datas(1, I), "compressor")
            
            COmp.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            COmp.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            COmp.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'Comp.CompPolytropicEffValue = Datas(13, i) * 100
            COmp.CompAdiabaticEffValue = Datas(13, I)
            

            'If there is no feed pressure, avoid the error by manually adding a feed pressure
            If COmp.FeedPressure = -32767 Then
                COmp.FeedPressureValue = 101
            End If
            

            'We cant add the pressure ratio with vba, so we set the Product pressure
            COmp.ProductPressure.SetValue (COmp.FeedPressure * Datas(8, I))
            
            COmp.SpeedInCompressorValue = Datas(14, I)
            
        End If
            
            'Same as the compressor
        If Datas(0, I) = "Gas Turbine" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
            Set Turb = hyFlwSht.Operations.Add(Datas(1, I), "expandop")
            
            Turb.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Turb.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Turb.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))

            
            'If there is no feed pressure, avoid the error
            If Turb.FeedPressure = -32767 Then
                Turb.FeedPressureValue = 505
            End If
            
            Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            
            Turb.ExpAdiabaticEffValue = Datas(13, I)
        End If
        
        
        
        
        'Same as the compressor
        If Datas(0, I) = "Pump" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
            Set Pump = hyFlwSht.Operations.Add(Datas(1, I), "pumpop")
            
            Pump.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Pump.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Pump.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Pump.FeedPressure = -32767 Then
                Pump.FeedPressureValue = 101
            End If
            Pump.AdiabaticEfficiency = Datas(13, I)
            Pump.ProductPressure.SetValue (Pump.FeedPressureValue * Datas(8, I))
        End If
    
        
        'Same as Compressor but for temperature
        If Datas(0, I) = "Cooler" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
        
        
            Set Cool = hyFlwSht.Operations.Add(Datas(1, I), "coolerop")
            
         
                Cool.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            
            Cool.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Cool.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Cool.FeedTemperature = -32767 Then
                Cool.FeedTemperatureValue = 17
            End If
            
            'If there is no feed pressure, avoid the error
            If Cool.FeedPressure = -32767 Then
                Cool.FeedPressureValue = 101
            End If
            
            
            Cool.ProductTemperature.SetValue Datas(12, I), "K"
            

            Cool.ProductPressure.SetValue (Cool.FeedPressureValue * (100 - Datas(9, I)) / 100)
        End If
    
        If Datas(0, I) = "Heater" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
        
            Set Heat = hyFlwSht.Operations.Add(Datas(1, I), "heaterop")
            
      
            Heat.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            
            Heat.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Heat.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Heat.FeedTemperature = -32767 Then
                Heat.FeedTemperatureValue = 17
            End If
            
            'If there is no feed pressure, avoid the error
            If Heat.FeedPressure = -32767 Then
                Heat.FeedPressureValue = 101
            End If
            
            Heat.ProductTemperature.SetValue Datas(12, I), "K"
            
            Heat.ProductPressure.SetValue (Heat.FeedPressure * (100 - Datas(9, I)) / 100)
            
        End If
        
        
        If Datas(0, I) = "Heat Exchanger" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
                       Set HX = hyFlwSht.Operations.Add(Datas(1, I), "heatexop")
            last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
            For v = 0 To last_column - 4
                If Datas(1, I) = HXSpec(0, v) Then
                    TypeHX = HXSpec(1, v)
                End If
            Next
            
            HX.TubeSideFeed = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            HX.TubeSideProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            HX.ShellSideFeed = hyFlwSht.MaterialStreams.Item(Datas(3, I))
            HX.ShellSideProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            
            If TypeHX = "Saturated Liquid" Then
                HX.TubeSideProduct.VapourFraction = 0
                
                
            ElseIf TypeHX = "Saturated Steam" Then
                HX.TubeSideProduct.VapourFraction = 1
                
            ElseIf TypeHX = "Superheated Steam" Or TypeHX = "Regeneration" Or TypeHX = "Reheat" Or TypeHX = "Heater" Then
                HX.TubeSideProduct.Temperature = Datas(12, I) - 273.15
            End If
                
             
 
           If HX.ShellSideFeed.Pressure <> -32767 Then
                HX.ShellSidePressureDropValue = Datas(10, I) / 100 * HX.ShellSideFeed.Pressure
           Else
                HX.ShellSidePressureDropValue = 4
           End If
           
           If HX.TubeSideFeed.Pressure <> -32767 Then
                HX.TubeSidePressureDropValue = Datas(10, I) / 100 * HX.TubeSideFeed.Pressure
           Else
                HX.TubeSidePressureDropValue = 40
           End If
            

        End If
        
        If Datas(0, I) = "Combustion Chamber" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
        
            NombreCC = NombreCC + 1
            
            Set CC = hyFlwSht.Operations.Add(Datas(1, I), "conversionreactorop")
            
            CC.Feeds.Add (Datas(2, I))
            CC.Feeds.Add (Datas(3, I))
            

            
            
            
            'We cannot implement the Flaming temperature in Aspen so we need to pilot the Fuel Mass Flow
            
            'Formula to calculate the Fuel Mass Flow to add in the CC given by the Flaming Temperature wanted
            FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
            
            
            
            If FuelFlow < 0 Then
                MsgBox "FuelFlow is negative, please correct your gas turbine. The error probably come from Entry mass flow"
                Exit Function
            Else
                CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
            End If
        
            
        
            
            CC.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            CC.PressureDrop = CC.Feeds.Item(0).Pressure * Datas(9, I) / 100
        
  
            Set CCL = hyFlwSht.MaterialStreams.Add("CCL" & Datas(1, I))
            CC.LiquidProduct = hyFlwSht.MaterialStreams.Item(CCL)
 
            
            
            If Datas(7, I) = "Reaction 1" Then
                CC.ReactionSet = SetReaction
            ElseIf Datas(7, I) = "Reaction 2" Then
                CC.ReactionSet = SetReaction2
            End If
        End If
    
   
    
        If Datas(0, I) = "Splitter" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
    
    
            Set Split = hyFlwSht.Operations.Add(Datas(1, I), "teeop")
        
        
            Split.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Split.Products.Add (Datas(4, I))
            Split.Products.Add (Datas(5, I))
        
        
            
           
            
            ratio = Split.SplitsValue
            ratio(0) = Datas(9, I)
            ratio(1) = Datas(10, I)
            
            Split.Splits.SetValues ratio, ""
        
        End If
        
        If Datas(0, I) = "Mixer" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
    
    
            Set Mixer = hyFlwSht.Operations.Add(Datas(1, I), "mixerop")
        
        
            Mixer.Feeds.Add (Datas(2, I))
            Mixer.Feeds.Add (Datas(3, I))
            If Datas(5, I) <> 0 Then
                Mixer.Feeds.Add (Datas(5, I))
            End If
            Mixer.Product = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            If Datas(9, I) <> 0 Then
                Mixer.Product.Pressure = Datas(9, I)
            End If
        End If
        
                    'Same as the compressor
        If Datas(0, I) = "Steam Turbine" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
            Set Turb = hyFlwSht.Operations.Add(Datas(1, I), "expandop")
            
            Turb.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Turb.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Turb.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))

            Turb.ExpAdiabaticEffValue = Datas(13, I)
            'If there is no feed pressure, avoid the error
            If Turb.FeedPressure = -32767 Then
                Turb.FeedPressureValue = 505
            End If
            
            Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            
            
        End If
        
            
        If Datas(0, I) = "Solar Heater" And FirstStream(j) = Datas(2, I) Then
        Datas(16, I) = "True"
        
            Set Solar = hyFlwSht.Operations.Add(Datas(1, I), "heaterop")
            
            Set TankCold = hyFlwSht.Operations.Add("CT" & Datas(1, I), "FlashTank")
            TankCold.SeparatorType = stTank
            
            Set TankHot = hyFlwSht.Operations.Add("HT" & Datas(1, I), "FlashTank")
            TankHot.SeparatorType = stTank
            
            
            Set Water = hyFlwSht.MaterialStreams.Add(Datas(2, I))
            
            
            col = Sheets(cycleName).Range("A19").End(xlToRight).column
            For colo = 3 To col
                If Datas(1, I) = Sheets(cycleName).Cells(19, colo) Then
                    bonnecol = colo
                End If
            Next
            
             With Water
            'Datas from the excel
            
            .Pressure.SetValue Sheets(cycleName).Cells(69, bonnecol).Value, "kPa"
            .Temperature.SetValue Sheets(cycleName).Cells(68, bonnecol).Value, "K"
            .MassFlow.SetValue 5, "kg/s" 'Sheets(cycleName).Cells(70, bonnecol).Value, "kg/s"
            
            'We define the composition of the stream
            
            varComps = .ComponentMolarFraction.Values

            ligne = Sheets(cycleName).Range("D83").End(xlDown).Row - 1
            For coucou = 0 To ligne - 86
                If hyFluidPkg.Components.Item(coucou) = "H2O" Or hyFluidPkg.Components.Item(coucou) = "Water" Then
                    varComps(coucou) = 1
                    
                Else
                    varComps(coucou) = 0
                End If
            Next
            .ComponentMolarFraction.Values = varComps
        End With
            
            
            TankCold.Feeds.Add Water
            Set Stream = hyFlwSht.MaterialStreams.Add("CTS" & Datas(1, I))
            
            If Water.Temperature > 100 Then
                TankCold.VapourProduct = Stream
                TankCold.LiquidProduct = hyFlwSht.MaterialStreams.Add("CTL" & Datas(1, I))
            Else
                TankCold.LiquidProduct = Stream
                TankCold.VapourProduct = hyFlwSht.MaterialStreams.Add("CTL" & Datas(1, I))
            End If
            TankCold.VesselPressureDropValue = 5 / 100 * Water.Pressure
            Set Stream = hyFlwSht.MaterialStreams.Add("Rec" & Datas(1, I))
            Solar.ProductStream = Stream
              Solar.EnergyStream = Datas(6, I)
            
            Solar.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            'Solar.DutyValue = Sheets(cycleName).Cells(65, bonnecol) / 100 * Sheets(cycleName).Cells(66, bonnecol) * Sheets(cycleName).Cells(67, bonnecol) / 1000
            Solar.ProductStream.Temperature = Sheets(cycleName).Cells(68, bonnecol) - 273.15
            Solar.ProductPressure.SetValue (Solar.FeedPressureValue) * 1.3
            Water.MolarFlow.SetValue Sheets(cycleName).Cells(69, bonnecol) * 3600 / (Solar.ProductStream.MolarEnthalpy - Solar.FeedStream.MolarEnthalpy)
            'Solar.ProductStream.MolarEnthalpyValue = Solar.FeedStream.MolarEnthalpy + Sheets(cycleName).Cells(65, bonnecol) / 100 * Sheets(cycleName).Cells(66, bonnecol) * Sheets(cycleName).Cells(67, bonnecol) / Sheets(cycleName).Cells(70, bonnecol) / 18.015
            
            'Solar.ProductTemperatureValue = Solar.FeedTemperature + Sheets(cycleName).Cells(20, bonnecol) / 100 * Sheets(cycleName).Cells(21, bonnecol) * Sheets(cycleName).Cells(22, bonnecol) / 0.01 / 4100
            
            
            TankHot.Feeds.Add Stream
            
            If Stream.Temperature > 100 Then
                Set Stream = hyFlwSht.MaterialStreams.Add(Datas(4, I))
                TankHot.VapourProduct = Stream
                TankHot.LiquidProduct = hyFlwSht.MaterialStreams.Add("HTL" & Datas(1, I))
            Else
                Set Stream = hyFlwSht.MaterialStreams.Add(Datas(4, I))
                TankHot.LiquidProduct = Stream
                TankHot.VapourProduct = hyFlwSht.MaterialStreams.Add("HTL" & Datas(1, I))
            End If
            TankHot.VesselPressureDropValue = 5 / 100 * Stream.Pressure
            
        End If
        
        If Datas(0, I) = "Stream Saturator" And FirstStream(j) = Datas(2, I) Then
            Datas(16, I) = "True"
            

            
                Set Sat = hyFlwSht.Operations.Add(Datas(1, I), "SaturatorOp")
                Sat.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
                Sat.WaterStream = hyFlwSht.MaterialStreams.Item(Datas(3, I))
                Sat.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
                Sat.RelativeHumidity = Datas(13, I)
        
        End If
     Next
    Next
     
     
     
     
    For I = 0 To X - 2
        
        
        '/////////////////////////////Construction of diffrrents components////////////////////////////
        
        'We dissociate different components
        
        If Datas(0, I) = "Valve" And Datas(16, I) = "False" Then
        
            Datas(16, I) = "True"
            Set Valve = hyFlwSht.Operations.Add("valve1", "valveop")
            Valve.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Valve.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            If Valve.ProductStream.Pressure = -32767 Then
                'Valve.ProductStream.PressureValue = Valve.FeedStream.Pressure * (1 - Datas(9, I) / 100)
                'MsgBox Valve.ProductStream.PressureValue
           End If

        End If
        
        If Datas(0, I) = "Compressor" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
            Set COmp = hyFlwSht.Operations.Add(Datas(1, I), "compressor")
            
            COmp.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            COmp.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            COmp.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'Comp.CompPolytropicEffValue = Datas(13, i) * 100
            COmp.CompAdiabaticEffValue = Datas(13, I)
            
            'If there is no feed pressure, avoid the error by manually adding a feed pressure
            If COmp.FeedPressure = -32767 Then
                COmp.FeedPressureValue = 101
            End If
            
            'We cant add the pressure ratio with vba, so we set the Product pressure
            COmp.ProductPressure.SetValue (COmp.FeedPressure * Datas(8, I))
            
            COmp.SpeedInCompressorValue = Datas(14, I)
            
        End If
            
            'Same as the compressor
        If Datas(0, I) = "Gas Turbine" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
            Set Turb = hyFlwSht.Operations.Add(Datas(1, I), "expandop")
            
            Turb.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Turb.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Turb.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))

            
            'If there is no feed pressure, avoid the error
            If Turb.FeedPressure = -32767 Then
                Turb.FeedPressureValue = 505
            End If
            
            Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            
            Turb.ExpAdiabaticEffValue = Datas(13, I)
        End If
        
        
        
        
        'Same as the compressor
        If Datas(0, I) = "Pump" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
            Set Pump = hyFlwSht.Operations.Add(Datas(1, I), "pumpop")
            
            Pump.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Pump.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Pump.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Pump.FeedPressure = -32767 Then
                Pump.FeedPressureValue = 101
            End If
            Pump.AdiabaticEfficiency = Datas(13, I)
            Pump.ProductPressure.SetValue (Pump.FeedPressureValue * Datas(8, I))
        End If
    
        
        'Same as Compressor but for temperature
        If Datas(0, I) = "Cooler" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
        
        
            Set Cool = hyFlwSht.Operations.Add(Datas(1, I), "coolerop")
            
         
                Cool.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            
            Cool.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Cool.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Cool.FeedTemperature = -32767 Then
                Cool.FeedTemperatureValue = 17
            End If
            
            'If there is no feed pressure, avoid the error
            If Cool.FeedPressure = -32767 Then
                Cool.FeedPressureValue = 101
            End If
            
            
            Cool.ProductTemperature.SetValue Datas(12, I), "K"
            

            Cool.ProductPressure.SetValue (Cool.FeedPressureValue * (100 - Datas(9, I)) / 100)
        End If
    
        If Datas(0, I) = "Heater" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
        
            Set Heat = hyFlwSht.Operations.Add(Datas(1, I), "heaterop")
            
      
            Heat.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            
            
            Heat.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            Heat.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
            'If there is no feed pressure, avoid the error
            If Heat.FeedTemperature = -32767 Then
                Heat.FeedTemperatureValue = 17
            End If
            
            'If there is no feed pressure, avoid the error
            If Heat.FeedPressure = -32767 Then
                Heat.FeedPressureValue = 101
            End If
            
            Heat.ProductTemperature.SetValue Datas(12, I), "K"
            
            Heat.ProductPressure.SetValue (Heat.FeedPressure * (100 - Datas(9, I)) / 100)
            
        End If
        
        
        If Datas(0, I) = "Heat Exchanger" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
            Set HX = hyFlwSht.Operations.Add(Datas(1, I), "heatexop")
            last_column = Sheets(cycleName).Range("B73").End(xlToRight).column
            For v = 0 To last_column - 4
                If Datas(1, I) = HXSpec(0, v) Then
                    TypeHX = HXSpec(1, v)
                End If
            Next
            
            HX.TubeSideFeed = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            HX.TubeSideProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            
            HX.ShellSideFeed = hyFlwSht.MaterialStreams.Item(Datas(3, I))
            HX.ShellSideProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            
            If TypeHX = "Saturated Liquid" Then
                HX.TubeSideProduct.VapourFraction = 0
                
                
            ElseIf TypeHX = "Saturated Steam" Then
                HX.TubeSideProduct.VapourFraction = 1
                
            ElseIf TypeHX = "Superheated Steam" Or TypeHX = "Regeneration" Or TypeHX = "Reheat" Or TypeHX = "Heater" Or TypeHX = "Condenser" Then
                 HX.TubeSideProduct.Temperature = Datas(12, I) - 273.15
            End If
                
             
 
           If HX.ShellSideFeed.Pressure <> -32767 Then
                HX.ShellSidePressureDropValue = Datas(10, I) / 100 * HX.ShellSideFeed.Pressure
           Else
                HX.ShellSidePressureDropValue = 4
           End If
           If TypeHX = "Condenser" Then
                If HX.TubeSideProduct.Pressure <> -32767 And HX.TubeSidePressureDrop = -32767 Then
                     HX.TubeSidePressureDropValue = (1 + Datas(10, I) / 100) * HX.TubeSideProduct.Pressure
                End If
            Else
                If HX.TubeSideFeed.Pressure <> -32767 Then
                     HX.TubeSidePressureDropValue = Datas(10, I) / 100 * HX.TubeSideFeed.Pressure
                Else
                     HX.TubeSidePressureDropValue = 40
                End If
           End If


            

        End If
        
        If Datas(0, I) = "Combustion Chamber" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
        
            NombreCC = NombreCC + 1
            
            Set CC = hyFlwSht.Operations.Add(Datas(1, I), "conversionreactorop")
            
            CC.Feeds.Add (Datas(2, I))
            CC.Feeds.Add (Datas(3, I))
            

            
            
            
            'We cannot implement the Flaming temperature in Aspen so we need to pilot the Fuel Mass Flow
            
            'Formula to calculate the Fuel Mass Flow to add in the CC given by the Flaming Temperature wanted
            FuelFlow = 1.275 * CC.Feeds.Item(0).MassFlow * CC.Feeds.Item(0).MassHeatCapacity * (Datas(12, I) - 273.15 - CC.Feeds.Item(0).Temperature) / CC.Feeds.Item(1).HigherHeatValue * CC.Feeds.Item(1).MolecularWeight
            
            
            
            If FuelFlow < 0 Then
                MsgBox "FuelFlow is negative, please correct your gas turbine. The error probably come from Entry mass flow"
                Exit Function
            Else
                CC.Feeds.Item(1).MassFlow.SetValue FuelFlow, "kg/s"
            End If
        
            
        
            
            CC.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            CC.PressureDrop = CC.Feeds.Item(0).Pressure * Datas(9, I) / 100

            Set CCL = hyFlwSht.MaterialStreams.Add("CCL" & Datas(1, I))
            CC.LiquidProduct = hyFlwSht.MaterialStreams.Item(CCL)
 
            
            
            
                CC.ReactionSet = SetReaction
               
        End If
    
   
    
        If Datas(0, I) = "Splitter" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
    
    
            Set Split = hyFlwSht.Operations.Add(Datas(1, I), "teeop")
        
        
            Split.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Split.Products.Add (Datas(4, I))
            Split.Products.Add (Datas(5, I))
        
        
            
            
            
            ratio = Split.SplitsValue
            ratio(0) = Datas(9, I)
            ratio(1) = Datas(10, I)
            
            Split.Splits.SetValues ratio, ""
        
        End If
        
        If Datas(0, I) = "Mixer" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
    
    
            Set Mixer = hyFlwSht.Operations.Add(Datas(1, I), "mixerop")
        
        
            Mixer.Feeds.Add (Datas(2, I))
            Mixer.Feeds.Add (Datas(3, I))
            If Datas(5, I) <> 0 Then
                Mixer.Feeds.Add (Datas(5, I))
            End If
           
            Mixer.Product = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            If Datas(9, I) <> 0 Then
                Mixer.Product.Pressure = Datas(9, I)
            End If
        
        End If
        
                    'Same as the compressor
        If Datas(0, I) = "Steam Turbine" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
            Set Turb = hyFlwSht.Operations.Add(Datas(1, I), "expandop")
            
            Turb.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
            Turb.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            

            Turb.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))

            Turb.ExpAdiabaticEffValue = Datas(13, I)
            'If there is no feed pressure, avoid the error
            If Turb.FeedPressure = -32767 Then
                Turb.FeedPressureValue = 505
            End If
            

            If Datas(8, I) <> 0 Then
           ' If Turb.ProductPressure = -32767 Or Turb.ProductPressure = 154.3 Then
                Turb.ProductPressure.SetValue (Turb.FeedPressureValue * Datas(8, I))
            End If
            

            
        End If
        
            
        If Datas(0, I) = "Solar Heater" And Datas(16, I) = "False" Then
        Datas(16, I) = "True"
        
            Set Solar = hyFlwSht.Operations.Add(Datas(1, I), "heaterop")
            Set Pump = hyFlwSht.Operations.Add("Pump" & Datas(1, I), "pumpop")
            Set Cool = hyFlwSht.Operations.Add("Cool" & Datas(1, I), "coolerop")
            Set TankCold = hyFlwSht.Operations.Add("CT" & Datas(1, I), "FlashTank")
            TankCold.SeparatorType = stTank
            
            Set TankHot = hyFlwSht.Operations.Add("HT" & Datas(1, I), "FlashTank")
            TankHot.SeparatorType = stTank
            
            
            Set Water = hyFlwSht.MaterialStreams.Add("Pump" & Datas(2, I))
            
            
            col = Sheets(cycleName).Range("B64").End(xlToRight).column
            For colo = 4 To col
                If Datas(1, I) = Sheets(cycleName).Cells(64, colo) Then
                    bonnecol = colo
                End If
            Next
            
             With Water
            'Datas from the excel
            
           ' .Pressure.SetValue Sheets(cycleName).Cells(24, bonnecol).Value, "kPa"
            .Temperature.SetValue Sheets(cycleName).Cells(35, 3).Value, "K"
            .MolarFlow.SetValue 1035, "kgmole/h" 'Sheets(cycleName).Cells(70, bonnecol).Value, "kg/s"
            
            'We define the composition of the stream
            
            varComps = .ComponentMolarFraction.Values
         
            ligne = Sheets(cycleName).Range("D83").End(xlDown).Row - 1
            For coucou = 0 To ligne - 86
                If hyFluidPkg.Components.Item(coucou) = "H2O" Or hyFluidPkg.Components.Item(coucou) = "Water" Then
                    varComps(coucou) = 1
                    
                Else
                    varComps(coucou) = 0
                End If
            Next
            .ComponentMolarFraction.Values = varComps
        End With
            
            TankCold.Feeds.Add Water
            Pump.ProductStream = Water
            Cool.FeedStream = hyFlwSht.MaterialStreams.Add(Datas(2, I))
            Cool.ProductStream = hyFlwSht.MaterialStreams.Add("Cool" & Datas(1, I))
            Pump.FeedStream = Cool.ProductStream
            Pump.EnergyStream = hyFlwSht.EnergyStreams.Add("QPump" & Datas(1, I))
            Cool.EnergyStream = hyFlwSht.EnergyStreams.Add("QCool" & Datas(1, I))
            Cool.PressureDrop = 4
            Set Stream = hyFlwSht.MaterialStreams.Add("CTS" & Datas(1, I))
            

                
            If Water.Temperature > 100 Then
                TankCold.VapourProduct = Stream
                TankCold.LiquidProduct = hyFlwSht.MaterialStreams.Add("CTL" & Datas(1, I))
                TankCold.VapourProduct.Pressure = 101.3

            Else
                TankCold.LiquidProduct = Stream
                TankCold.VapourProduct = hyFlwSht.MaterialStreams.Add("CTV" & Datas(1, I))
                TankCold.LiquidProduct.Pressure = 101.3
            End If
            
            'TankCold.LiquidProduct.Pressure.SetValue 101.3

                
            Solar.FeedStream = Stream
            
            Set Stream = hyFlwSht.MaterialStreams.Add("Rec" & Datas(1, I))
            Solar.ProductStream = Stream
           Solar.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            
                Solar.ProductStream.Temperature = Sheets(cycleName).Cells(68, bonnecol) - 273.15
                Solar.ProductPressure.SetValue (Solar.FeedPressureValue)
               
            Water.MolarFlow.SetValue Sheets(cycleName).Cells(69, bonnecol) * 3600 / (Solar.ProductStream.MolarEnthalpy - Solar.FeedStream.MolarEnthalpy), "kgmole/h"
            
            'Solar.EnergyStream = hyFlwSht.EnergyStreams.Add(Datas(6, I))
            'Solar.DutyValue = Sheets(cycleName).Cells(65, bonnecol) / 100 * Sheets(cycleName).Cells(66, bonnecol) * Sheets(cycleName).Cells(67, bonnecol) / 1000
            'Solar.ProductStream.MolarEnthalpyValue = Solar.FeedStream.MolarEnthalpy + Sheets(cycleName).Cells(65, bonnecol) * Sheets(cycleName).Cells(66, bonnecol) * Sheets(cycleName).Cells(67, bonnecol) / Sheets(cycleName).Cells(70, bonnecol) * 18.015 / 1000
            
            'Solar.ProductTemperatureValue = Solar.FeedTemperature + Sheets(cycleName).Cells(65, bonnecol) / 100 * Sheets(cycleName).Cells(66, bonnecol) * Sheets(cycleName).Cells(67, bonnecol) / 0.01 / 4100
            
            
            
            TankHot.Feeds.Add Stream
            TankHot.VesselPressureDropValue = 5 / 100 * Stream.Pressure
            If Stream.Temperature > 100 Then
                Set Stream = hyFlwSht.MaterialStreams.Add(Datas(4, I))
                TankHot.VapourProduct = Stream
                TankHot.LiquidProduct = hyFlwSht.MaterialStreams.Add("HTL" & Datas(1, I))
            Else
                Set Stream = hyFlwSht.MaterialStreams.Add(Datas(4, I))
                TankHot.LiquidProduct = Stream
                TankHot.VapourProduct = hyFlwSht.MaterialStreams.Add("HTL" & Datas(1, I))
            End If
            
            
        End If
        
        If Datas(0, I) = "Stream Saturator" And Datas(16, I) = "False" Then
            Datas(16, I) = "True"
            

            
                Set Sat = hyFlwSht.Operations.Add(Datas(1, I), "SaturatorOp")
                Sat.FeedStream = hyFlwSht.MaterialStreams.Item(Datas(2, I))
                Sat.WaterStream = hyFlwSht.MaterialStreams.Item(Datas(3, I))
                Sat.ProductStream = hyFlwSht.MaterialStreams.Item(Datas(4, I))
                Sat.RelativeHumidity = Datas(13, I)
        
        End If
        
                If Datas(0, I) = "Fired Heater" And Datas(16, I) = "False" Then
            Datas(16, I) = "True"
            Set Fired = hyFlwSht.Operations.Add(Datas(1, I), "dynamicfiredheaterop")
            Fired.RadInlet.Add Datas(2, I)
            Fired.ExcessAirPercentValue = Datas(9, I)
            Fired.FuelsIn.Add Datas(3, I)
            Fired.CombustionEfficiencyValue = Datas(13, I)
            Fired.RadOutlet.Add Datas(4, I)
        
             Set Air = hyFlwSht.MaterialStreams.Add("Air" & Datas(1, I))
             With Air
            'Datas from the excel
            .Pressure.SetValue Sheets(cycleName).Range("C36").Value, "kPa"
            .Temperature.SetValue Sheets(cycleName).Range("C35").Value, "K"
            '.MassFlow.SetValue 0.2, "kg/s" 'Security Value, tbd later on the program
            
            'We define the composition of the stream
            varComps = .ComponentMolarFraction.Values
             ligne = Sheets(cycleName).Range("D83").End(xlDown).Row - 1
            For coucou = 0 To ligne - 86
                If hyFluidPkg.Components.Item(coucou) = "Nitrogen" Then
                    varComps(coucou) = 0.79
                ElseIf hyFluidPkg.Components.Item(coucou) = "Oxygen" Then
                    varComps(coucou) = 0.21
                Else
                    varComps(coucou) = 0
                End If
            Next
            .ComponentMolarFraction.Values = varComps
            
            End With
            
            If Sheets(cycleName).Range("O32") <> "" Then
            
                Fired.AttachedProducts.Item(0).Temperature.SetValue Datas(12, I), "K"
               ' Fired.ConvInlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("O33")))
                'Fired.ConvOutlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("P33")))
                'Fired.ConvOutlet.Item(0).Temperature = Sheets(cycleName).Range("Q33") - 273.15
                Fired.EconOutlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("P34")))
                Fired.EconInlet.Add (hyFlwSht.MaterialStreams.Item(Sheets(cycleName).Range("O34")))
                Fired.AttachedProducts.Item(1).Temperature.SetValue Sheets(cycleName).Range("Q34"), "K"
            Else
                        
                If Fired.RadInlet.Item(0).Temperature > Datas(12, I) - 273.15 Then
                    Fired.AttachedProducts.Item(0).Temperature.SetValue Fired.RadInlet.Item(0).Temperature + 1
                Else
                    Fired.AttachedProducts.Item(0).Temperature.SetValue Datas(12, I), "K"
                End If
            End If
            'Fired.AttachedProducts.Item(0).Pressure.SetValue Fired.RadInlet.Item(0).Pressure * 0.9
            Fired.BurnerFeed = Air
            Set Flare = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Fired.CombustionProduct = Flare
            

            
            'If Fired.CombustionProduct.Temperature < Fired.AttachedProducts.Item(0).Temperature Then
               ' MsgBox "Temperature Cross in Fired Heater : " & Datas(1, I) & ". Please Correct your cycle."
            'End If
        End If
        
                
        If Datas(0, I) = "Flash" And Datas(16, I) = "False" Then
            Datas(16, I) = "True"
            Set Tank = hyFlwSht.Operations.Add(Datas(1, I), "FlashTank")
            Tank.Feeds.Add (Datas(2, I))
            Tank.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Tank.LiquidProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
        
        If Datas(0, I) = "Tank" And Datas(16, I) = "False" Then
            Datas(16, I) = "True"
            Set Tank = hyFlwSht.Operations.Add(Datas(1, I), "FlashTank")
            Tank.SeparatorType = stType
            Tank.Feeds.Add (Datas(2, I))
            Tank.VapourProduct = hyFlwSht.MaterialStreams.Item(Datas(5, I))
            Tank.LiquidProduct = hyFlwSht.MaterialStreams.Item(Datas(4, I))
            Tank.VesselPressureDropValue = Datas(10, I) / 100 * Tank.AttachedFeeds.Item(0).Pressure
           

        End If
       
     Next
     
     ComponentCreation = "Done"
    '///////////////////////Fin des composants//////////////////////////////////////////////////////
End Function

