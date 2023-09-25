Attribute VB_Name = "Module7"
Function DefCycle(Datas() As Variant, strcase As String, myCompoCollec As Collection, cycleName As String) As Collection

Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet
Dim Ta As Double, Tb As Double, Tc As Double, Td As Double, Te As Double, Qb As Double, Qc As Double
Dim power As Double

Qb = 0.001
Qc = 0.001
Set DefCycle = New Collection

Dim oComponent As cComponent
Dim oCycle As cCycle

Dim NumberCycle As Integer, NumberCompressor As Integer, NumberTurbine As Integer, PR As Double
PR = 1
NumberCycle = 0
NumberCompressor = 0
NumberTurbine = 0
NumberPump = 0

ligne = Sheets(cycleName).Range("F32").End(xlDown).Row
For I = 33 To ligne
    power = 0
    Qin = 0
    NumberCompressor = 0
    NumberTurbine = 0
    NumberPump = 0
    PR = 1
    Set oCycle = New cCycle
    DefCycle.Add Item:=oCycle
    NumberCycle = NumberCycle + 1
    oCycle.index = NumberCycle
    oCycle.name = Sheets(cycleName).Cells(I, 6)
    oCycle.CType = Sheets(cycleName).Cells(I, 7)
    oCycle.StreamPilot = Sheets(cycleName).Cells(I, 8)
    
    For Each Stream In hyFlwSht.MaterialStreams
        If Stream.name = oCycle.StreamPilot Then
            oCycle.FeedFlow = Stream.MassFlow
        End If
    Next
    
    If oCycle.CType = "Brayton" Or oCycle.CType = "Regeneration Brayton" Or oCycle.CType = "Reheat Brayton" Then
        For Each oComponent In myCompoCollec
        
            If oComponent.cycleName = oCycle.name Then
            
                If oComponent.CompType = "Compressor" Then
                   
                    power = power + oComponent.power
                    PR = PR * oComponent.Pout / oComponent.Pin
                    NumberCompressor = NumberCompressor + 1
                    
                ElseIf (oComponent.CompType = "Combustion Chamber" Or oComponent.CompType = "Fired Heater") Then
                   
                    oCycle.FuelFlow = oComponent.Fin2
                    oCycle.FiringTemp = oComponent.Tout
                    Qin = Qin + oComponent.HHV * oComponent.Fin2
                    QFuel = QFuel + oComponent.HHV * oComponent.Fin2

                ElseIf oComponent.CompType = "Gas Turbine" Then
                    
                    power = power + oComponent.power
                    NumberTurbine = NumberTurbine + 1
                    
                ElseIf oComponent.CompType = "Heat Exchanger" And oComponent.HXType = "Heater" Then
                   
                    'Qin = Qin + oComponent.Fin * (oComponent.hout - oComponent.hIn)
                    
                
                
                ElseIf oComponent.CompType = "Heater" Then
                    power = power + oComponent.power
                End If
                
            End If

         Next
         
         

            oCycle.Efficiency = power / Qin
            'If cycleName <> "Solar Regeneration Brayton" Then
                oCycle.HeatRate = power / QFuel

            oCycle.HeatPower = QFuel
            oCycle.power = power
            oCycle.PressureRatio = PR
            oCycle.NumberCompressor = NumberCompressor
            oCycle.NumberTurbine = NumberTurbine
            
            
    ElseIf oCycle.CType = "Rankine" Or oCycle.CType = "ORC Rankine" Then
        
        For Each oComponent In myCompoCollec
    
             If oComponent.cycleName = oCycle.name Then

                If oComponent.CompType = "Heater" Or oComponent.HXType = "Heater" Or oComponent.HXType = "Heater" Or oComponent.HXType = "Saturated Steam" Or oComponent.HXType = "Superheated Steam" Or oComponent.HXType = "Reheat" Then
                    If cycleName <> "Fired Rankine Test" And cycleName <> "Solar Fired Rankine Test" And cycleName <> "ORC Rankine" Then
                        Qb = Qb + oComponent.Fin * (oComponent.hout - oComponent.hIn)
                    End If
                   

                ElseIf oComponent.CompType = "Fired Heater" Then
                
                    oCycle.FuelFlow = oCycle.FuelFlow + oComponent.Fin2
                    oCycle.FiringTemp = oComponent.Tout
                    QFuel = QFuel + oComponent.Fin * (oComponent.hout - oComponent.hIn)
                    'QFuel = QFuel + oComponent.HHV * oComponent.Fin2
                    If cycleName = "Fired Rankine Test" Or cycleName = "Solar Fired Rankine Test" Or cycleName = "ORC Rankine" Then
                        Qb = Qb + oComponent.Fin * (oComponent.hout - oComponent.hIn)
                        
                    End If

  
                ElseIf oComponent.CompType = "Pump" Then
                    power = power + oComponent.power
                    NumberPump = NumberPump + 1
                    PR = PR * oComponent.Pout / oComponent.Pin
                ElseIf oComponent.CompType = "Steam Turbine" Then
                    power = power + oComponent.power
                    NumberTurbine = NumberTurbine + 1
                End If
            End If

         Next
         
         oCycle.Efficiency = Abs(power) / Qb
         oCycle.HeatRate = Abs(power) / QFuel
         oCycle.HeatPower = QFuel
         oCycle.power = power
         oCycle.PressureRatio = PR
         oCycle.NumberPump = NumberPump
         oCycle.NumberTurbine = NumberTurbine
         
'    ElseIf oCycle.CType = "Heat Source" Then
'        For Each oComponent In myCompoCollec
'
'
'            If oComponent.CompType = "Fired Heater" Then
'                oCycle.FuelFlow = oComponent.Fin2
'                Qb = Qb + oComponent.Fin * (oComponent.hout - oComponent.hIn)
'                QFuel = QFuel + oComponent.HHV * oComponent.Fin2
'            End If
'
'         Next
'
'         oCycle.power = power
'         oCycle.HeatPower = QFuel
'         oCycle.PressureRatio = PR
'         oCycle.NumberPump = NumberPump
'         oCycle.NumberTurbine = NumberTurbine
'
'    ElseIf oCycle.CType = "Solar" Then
'        For Each oComponent In myCompoCollec
'
'             If oComponent.cycleName = oCycle.name Then
'                If oComponent.CompType = "Solar Heater" Then
'                    Qb = oComponent.power
'                ElseIf oComponent.CompType = "Heat Exchanger" Then
'                    power = Fin * (hout - hIn)
'                End If
'              End If
'        Next
'
'        oCycle.Efficiency = Abs(power) / Qb
'
'
  End If



    
            oCycle.Cost = 0
            For Each oComponent In myCompoCollec
                If oCycle.name = oComponent.cycleName Then
                    oCycle.Cost = oCycle.Cost + oComponent.PEC
                End If
            Next
            
            
            If oCycle.power = 0 Then
                oCycle.power = 1
            End If
            Ij = Sheets(cycleName).Range("E40")
            Ny = Sheets(cycleName).Range("E41")
            PhiM = Sheets(cycleName).Range("E42")
            OH = Sheets(cycleName).Range("C43")
            NA = Sheets(cycleName).Range("E43")
            'oCycle.costKWH = oCycle.Cost / oCycle.power * (Ij / 100 * (1 + Ij / 100) ^ (Ny) / ((1 + Ij / 100) ^ (Ny) - 1)) * (PhiM) ^ (NA) / OH / NA
            
Next

End Function

Function CreateCompObject(Datas() As Variant, strcase As String, myComp As Collection, myTurb As Collection, cycleName As String) As Collection

Dim simCase As SimulationCase

Dim hyFluidPkg As HYSYS.FluidPackage

Set simCase = GetObject(strcase)
Set hyFlwSht = simCase.Flowsheet

Dim oComponent As cComponent
Dim OComp As cCompressor
Dim OTurb As cTurbine
Dim colonne3 As Integer
Set CreateCompObject = New Collection

Dim NumberComponent As Integer
NumberComponent = 0

X = Sheets(cycleName).Range("A10").End(xlToRight).column


    For I = 0 To X - 2
       
        
        '/////////////////////////////Construction of diffrrents components////////////////////////////
        
        'We dissociate different components
        
        If Datas(0, I) = "Tank" Then
            Set Tank = hyFlwSht.Operations.Item(Datas(1, I))
            
            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = Tank.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = Tank.LiquidProduct.Pressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = Tank.VapourProduct.Pressure
            oComponent.Tin = Tank.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = Tank.LiquidProduct.Temperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = Tank.VapourProduct.Temperature + 273.15
            oComponent.Fin = Tank.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = Tank.LiquidProduct.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = Tank.VapourProduct.MassFlow
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = Tank.VesselPressureDrop * 100 / Tank.AttachedFeeds.Item(0).Pressure
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            
        ElseIf Datas(0, I) = "Flash" Then
            
            Set Tank = hyFlwSht.Operations.Item(Datas(1, I))
            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            
            oComponent.Pin = Tank.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = Tank.LiquidProduct.Pressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = Tank.VapourProduct.Pressure
            oComponent.Tin = Tank.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = Tank.LiquidProduct.Temperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = Tank.VapourProduct.Temperature + 273.15
            oComponent.Fin = Tank.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = Tank.LiquidProduct.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = Tank.VapourProduct.MassFlow
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = Tank.VesselPressureDrop * 100 / Tank.AttachedFeeds.Item(0).Pressure
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
        
        ElseIf Datas(0, I) = "Fired Heater" Then

            Set Fired = hyFlwSht.Operations.Item(Datas(1, I))
            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = Fired.RadInlet.Item(0).Pressure
            oComponent.Pout = Fired.RadOutlet.Item(0).Pressure
            oComponent.Pin2 = Fired.FuelsIn.Item(0).Pressure
            oComponent.Pout2 = Fired.CombustionProduct.Pressure
            oComponent.Tin = Fired.RadInlet.Item(0).Temperature + 273.15
            oComponent.Tout = Fired.RadOutlet.Item(0).Temperature + 273.15
            oComponent.Tin2 = Fired.FuelsIn.Item(0).Temperature + 273.15
            oComponent.Tout2 = Fired.CombustionProduct.Temperature + 273.15
            oComponent.Fin = Fired.RadInlet.Item(0).MassFlow
            oComponent.Fout = Fired.RadOutlet.Item(0).MassFlow
            oComponent.Fin2 = Fired.FuelsIn.Item(0).MassFlow
            oComponent.Fout2 = Fired.CombustionProduct.MassFlow
            oComponent.power = Fired.AttachedProducts.Item(0).HeatFlow - Fired.RadInlet.Item(0).HeatFlow
            oComponent.Efficiency = Fired.CombustionEfficiency
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = Fired.ExcessAirPercent
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.Cp = Fired.RadOutlet.Item(0).MassHeatCapacity
            oComponent.LastComponent = False
            oComponent.HHV = Fired.FuelsIn.Item(0).MassHigherHeatValue
            oComponent.hIn = Fired.RadInlet.Item(0).MassEnthalpy
            oComponent.hout = Fired.RadOutlet.Item(0).MassEnthalpy
            If Sheets(cycleName).Range("J33") <> "" Then
                ligne = Sheets(cycleName).Range("J32").End(xlDown).Row
                For j = 33 To ligne
                    If oComponent.cycleName = Sheets(cycleName).Cells(j, 10) And oComponent.CompName = Sheets(cycleName).Cells(j, 11) And oComponent.CompType = Sheets(cycleName).Cells(j, 12) Then
                        oComponent.LastComponent = True
                    End If
                Next
            End If

        ElseIf Datas(0, I) = "Compressor" Then
            Set COmp = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = COmp.FeedPressure
            oComponent.Pout = COmp.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = COmp.FeedTemperature + 273.15
            oComponent.Tout = COmp.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = COmp.FeedStream.MassFlow

            oComponent.Fout = COmp.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = -COmp.EnergyValue
            oComponent.Efficiency = COmp.CompAdiabaticEff
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = COmp.ProductPressure / COmp.FeedPressure
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            
            If Sheets(cycleName).Range("J33") <> "" Then
                ligne = Sheets(cycleName).Range("J32").End(xlDown).Row
                For j = 33 To ligne
                    If oComponent.cycleName = Sheets(cycleName).Cells(j, 10) And oComponent.CompName = Sheets(cycleName).Cells(j, 11) And oComponent.CompType = Sheets(cycleName).Cells(j, 12) Then
                        oComponent.LastComponent = True
                    End If
                Next
            End If
            
            
            For Each OComp In myComp
                If OComp.AssComp = I Then
                    oComponent.TipSpeed = OComp.TipS
                    oComponent.RotatingSpeed = OComp.RSpeed
                    oComponent.Diameter = OComp.Diameter
                    oComponent.NumberStage = OComp.NumberStages
                End If
            Next

        ElseIf Datas(0, I) = "Gas Turbine" Then
            Set Turb = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
             oComponent.Pin = Turb.FeedPressure
            oComponent.Pout = Turb.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Turb.FeedTemperature + 273.15
            oComponent.Tout = Turb.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Turb.FeedStream.MassFlow
            oComponent.Fout = Turb.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = Turb.EnergyValue
            oComponent.Efficiency = Turb.ExpAdiabaticEff
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = Turb.ProductPressure / Turb.FeedPressure
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            
            For Each OTurb In myTurb
                If OTurb.AssTurb = I Then
                    oComponent.TipSpeed = OTurb.Ut
                    oComponent.RotatingSpeed = OTurb.RotatingSpeed
                    oComponent.Diameter = OTurb.Dm
                    oComponent.NumberStage = OTurb.NumberStage
                End If
            Next
            
            If Sheets(cycleName).Range("J33") <> "" Then
                ligne = Sheets(cycleName).Range("J32").End(xlDown).Row
                For j = 33 To ligne
                    If oComponent.cycleName = Sheets(cycleName).Cells(j, 10) And oComponent.CompName = Sheets(cycleName).Cells(j, 11) And oComponent.CompType = Sheets(cycleName).Cells(j, 12) Then
                        oComponent.LastComponent = True
                    End If
                Next
            End If

        ElseIf Datas(0, I) = "Pump" Then
            Set Pump = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
             oComponent.Pin = Pump.FeedPressure
            oComponent.Pout = Pump.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Pump.FeedTemperature + 273.15
            oComponent.Tout = Pump.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Pump.FeedStream.MassFlow
            oComponent.Fout = Pump.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = -Pump.EnergyStream.HeatFlow
            oComponent.Efficiency = Pump.AdiabaticEfficiency
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = Pump.ProductPressure / Pump.FeedPressure
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            oComponent.Fmin = Pump.FeedStream.IdealLiquidVolumeFlow * 3600
            
        ElseIf Datas(0, I) = "Cooler" Then
            Set Cool = hyFlwSht.Operations.Item(Datas(1, I))
            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
             oComponent.Pin = Cool.FeedPressure
            oComponent.Pout = Cool.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Cool.FeedTemperature + 273.15
            oComponent.Tout = Cool.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Cool.FeedStream.MassFlow
            oComponent.Fout = Cool.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = Cool.Duty
            oComponent.Efficiency = 0
            oComponent.DeltaT = Cool.DeltaT
            oComponent.DeltaP1 = Cool.PressureDrop * 100 / Cool.FeedPressure
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False


        ElseIf Datas(0, I) = "Heater" Then
            Set Heat = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
             oComponent.Pin = Heat.FeedPressure
            oComponent.Pout = Heat.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Heat.FeedTemperature + 273.15
            oComponent.Tout = Heat.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Heat.FeedStream.MassFlow
            oComponent.Fout = Heat.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = Heat.Duty
            oComponent.Efficiency = 0
            oComponent.DeltaT = Heat.DeltaT
            oComponent.DeltaP1 = Heat.PressureDrop * 100 / Heat.FeedPressure
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False


        ElseIf Datas(0, I) = "Heat Exchanger" Then
            Set HX = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = HX.TubeSideFeed.Pressure
            oComponent.Pout = HX.TubeSideProduct.Pressure
            oComponent.Pin2 = HX.ShellSideFeed.Pressure
            oComponent.Pout2 = HX.ShellSideProduct.Pressure
            oComponent.Tin = HX.TubeSideFeed.Temperature + 273.15
            oComponent.Tout = HX.TubeSideProduct.Temperature + 273.15
            oComponent.Tin2 = HX.ShellSideFeed.Temperature + 273.15
            oComponent.Tout2 = HX.TubeSideFeed.Temperature + 273.15
            oComponent.Fin = HX.TubeSideFeed.MassFlow
            oComponent.Fout = HX.TubeSideProduct.MassFlow
            oComponent.Fin2 = HX.ShellSideFeed.MassFlow
            oComponent.Fout2 = HX.TubeSideProduct.MassFlow
            oComponent.power = HX.TubeSideProduct.HeatFlow - HX.TubeSideFeed.HeatFlow
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = HX.TubeSidePressureDrop * 100 / HX.TubeSideFeed.Pressure
            oComponent.DeltaP2 = HX.ShellSidePressureDrop * 100 / HX.ShellSideFeed.Pressure
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            oComponent.hIn = HX.TubeSideFeed.MassEnthalpy
            oComponent.hout = HX.TubeSideProduct.MassEnthalpy
            oComponent.hout2 = HX.ShellSideProduct.MassEnthalpy
            colonne3 = Sheets(cycleName).Range("B73").End(xlToRight).column
            
            For g = 0 To colonne3 - 3
                If Sheets(cycleName).Cells(73, g + 3) = oComponent.CompName Then
                    oComponent.HXType = Sheets(cycleName).Cells(74, g + 3)
                End If
            Next
            
            If Sheets(cycleName).Range("J33") <> "" Then
                ligne = Sheets(cycleName).Range("J32").End(xlDown).Row
                For j = 33 To ligne
                    If oComponent.cycleName = Sheets(cycleName).Cells(j, 10) And oComponent.CompName = Sheets(cycleName).Cells(j, 11) And oComponent.CompType = Sheets(cycleName).Cells(j, 12) Then
                        oComponent.LastComponent = True
                    End If
                Next
            End If

        ElseIf Datas(0, I) = "Combustion Chamber" Then
            Set CC = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = CC.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = CC.VapourProduct.Pressure
            oComponent.Pin2 = CC.AttachedFeeds.Item(1).Pressure
            oComponent.Pout2 = 0
            oComponent.Tin = CC.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = CC.VapourProduct.Temperature + 273.15
            oComponent.Tin2 = CC.AttachedFeeds.Item(1).Temperature + 273.15
            oComponent.Tout2 = 0
            oComponent.Fin = CC.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = CC.VapourProduct.MassFlow
            oComponent.Fin2 = CC.AttachedFeeds.Item(1).MassFlow
            oComponent.Fout2 = 0
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = CC.PressureDrop * 100 / CC.AttachedFeeds.Item(0).Pressure
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = CC.ReactionSet
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            oComponent.Cp = CC.VapourProduct.MassHeatCapacity
            oComponent.HHV = CC.Feeds.Item(1).MassHigherHeatValue
            
            
            If Sheets(cycleName).Range("J33") <> "" Then
                ligne = Sheets(cycleName).Range("J32").End(xlDown).Row
                For j = 33 To ligne
                    If oComponent.cycleName = Sheets(cycleName).Cells(j, 10) And oComponent.CompName = Sheets(cycleName).Cells(j, 11) And oComponent.CompType = Sheets(cycleName).Cells(j, 12) Then
                        oComponent.LastComponent = True
                    End If
                Next
            End If
            



        ElseIf Datas(0, I) = "Splitter" Then
             Set Split = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = Split.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = Split.Products.Item(0).Pressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = Split.Products.Item(0).Pressure
            oComponent.Tin = Split.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = Split.Products.Item(0).Temperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = Split.Products.Item(0).Temperature + 273.15
            oComponent.Fin = Split.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = Split.Products.Item(0).MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = Split.Products.Item(0).MassFlow
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = Datas(9, I)
            oComponent.ExtraPercentage2 = Datas(10, I)
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False


        ElseIf Datas(0, I) = "Mixer" Then
            Set Mixer = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = Mixer.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = Mixer.Product.Pressure
            oComponent.Pin2 = Mixer.AttachedFeeds.Item(1).Pressure
            oComponent.Pout2 = 0
            oComponent.Tin = Mixer.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = Mixer.Product.Temperature + 273.15
            oComponent.Tin2 = Mixer.AttachedFeeds.Item(1).Temperature + 273.15
            oComponent.Tout2 = 0
            oComponent.Fin = Mixer.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = Mixer.Product.MassFlow
            oComponent.Fin2 = Mixer.AttachedFeeds.Item(1).MassFlow
            oComponent.Fout2 = 0
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False



        ElseIf Datas(0, I) = "Steam Turbine" Then
            Set Turb = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
             oComponent.Pin = Turb.FeedPressure
            oComponent.Pout = Turb.ProductPressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Turb.FeedTemperature + 273.15
            oComponent.Tout = Turb.ProductTemperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Turb.FeedStream.MassFlow
            oComponent.Fout = Turb.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = Turb.EnergyValue
            oComponent.Efficiency = Turb.ExpAdiabaticEff
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = Turb.ProductPressure / Turb.FeedPressure
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False

        ElseIf Datas(0, I) = "Solar Heater" Then
             Set Solar = hyFlwSht.Operations.Item(Datas(1, I))

            Set TankCold = hyFlwSht.Operations.Item("CT" & Datas(1, I))

            Set TankHot = hyFlwSht.Operations.Item("HT" & Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            
            col = Sheets(cycleName).Range("B64").End(xlToRight).column
            For colo = 4 To col
                If Datas(1, I) = Sheets(cycleName).Cells(64, colo) Then
                    bonnecol = colo
                End If
            Next
            
            oComponent.index = NumberComponent
            oComponent.Pin = TankCold.AttachedFeeds.Item(0).Pressure
            oComponent.Pout = TankHot.AttachedProducts.Item(Datas(4, I)).Pressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = TankCold.AttachedFeeds.Item(0).Temperature + 273.15
            oComponent.Tout = TankHot.AttachedProducts.Item(Datas(4, I)).Temperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = TankCold.AttachedFeeds.Item(0).MassFlow
            oComponent.Fout = TankHot.AttachedProducts.Item(Datas(4, I)).MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = Solar.Duty
            oComponent.Efficiency = Sheets(cycleName).Cells(65, bonnecol)
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = Sheets(cycleName).Cells(66, bonnecol)
            oComponent.SolarRadiation = Sheets(cycleName).Cells(67, bonnecol)
            oComponent.ExtraPercentage = 0
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False


        ElseIf Datas(0, I) = "Stream Saturator" Then
            Set Sat = hyFlwSht.Operations.Item(Datas(1, I))

            Set oComponent = New cComponent
            CreateCompObject.Add Item:=oComponent
            NumberComponent = NumberComponent + 1
            
            oComponent.index = NumberComponent
            oComponent.Pin = Sat.FeedStream.Pressure
            oComponent.Pout = Sat.ProductStream.Pressure
            oComponent.Pin2 = 0
            oComponent.Pout2 = 0
            oComponent.Tin = Sat.FeedStream.Temperature + 273.15
            oComponent.Tout = Sat.ProductStream.Temperature + 273.15
            oComponent.Tin2 = 0
            oComponent.Tout2 = 0
            oComponent.Fin = Sat.FeedStream.MassFlow
            oComponent.Fout = Sat.ProductStream.MassFlow
            oComponent.Fin2 = 0
            oComponent.Fout2 = 0
            oComponent.power = 0
            oComponent.Efficiency = 0
            oComponent.DeltaT = 0
            oComponent.DeltaP1 = 0
            oComponent.DeltaP2 = 0
            oComponent.PressureRatio = 0
            oComponent.Reaction = 0
            oComponent.CollectorSize = 0
            oComponent.SolarRadiation = 0
            oComponent.ExtraPercentage = Sat.RelativeHumidity
            oComponent.ExtraPercentage2 = 0
            oComponent.cycleName = Datas(15, I)
            oComponent.CompName = Datas(1, I)
            oComponent.CompType = Datas(0, I)
            oComponent.LastComponent = False
            
        End If
        

    Next

End Function
