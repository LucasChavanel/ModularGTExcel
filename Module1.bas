Attribute VB_Name = "Module1"
Option Explicit

Public simCase As SimulationCase
Public hyBasis As HYSYS.BasisManager
Public hyFluidPkg As HYSYS.FluidPackage

Public hyPRFluidPkg As HYSYS.PengRobinsonPropPkg
Public hyFlwSht As HYSYS.Flowsheet
Public hyStrm As HYSYS.ProcessStream

Public Reaction As ConversionReaction
Public SetReaction As ReactionSet
Public hyReac As Reactant

Public Reaction2 As ConversionReaction
Public SetReaction2 As ReactionSet
Public hyReac2 As Reactant

Public strcase As String
Public varComps As Variant
Public I As Integer
Public j As Integer

Public COmp As CompressOp
Public Turb As ExpandOp
Public Cool As CoolerOp
Public HX As HeatExchanger
Public CC As ConversionReactor
Public Heat As HeaterOp
Public Pump As PumpOp
Public Split As TeeOp
Public Mixer As MixerOp
Public Solar As HeaterOp
Public Tank As Separator
Public TankCold As Separator
Public TankHot As Separator
Public Sat As SaturatorOp
Public Fired As FiredHeater
Public Air As ProcessStream
Public Flare As ProcessStream
Public Stream As ProcessStream
Public Out As ProcessStream
Public Valve As Valve

Public NombreCC As Integer

Public NombreResults As Integer




Function Creation_Turbine(Datas() As Variant, cycleName As String) As Collection
Dim LastResults As Integer
LastResults = ActiveWorkbook.Sheets.count - 4


    

    Dim IndexResults As Integer
    Dim name As String
    IndexResults = LastResults + 1
    name = "Results" & IndexResults
    Set hyApp = CreateObject("HYSYS.Application")    'Connect itself at Aspen and create a new instance
    hyApp.Visible = True
    strcase = ThisWorkbook.Path & "\" & name & ".hsc"
    'Create new file with the name "ResultsX" (X is the index of the results : 1 for the first cycle, 2 for the second one, etc...)
    Dim valid As Boolean
    valid = False
    Dim Item As SimulationCase
    Dim I As Integer
    
   hyApp.SimulationCases.Close
    
    
    strcase = ThisWorkbook.Path & "\" & name & ".hsc"
    

    
    Set simCase = hyApp.SimulationCases.Add(strcase)
    simCase.Visible = True



    Dim Aye As String
Dim col1 As Integer, col2 As Integer, NombreReac As Integer, ligne As Integer
Dim TheSet As String
Dim X As Integer
Dim Y As Integer
X = UBound(Datas, 2) - LBound(Datas, 2) + 1
Y = UBound(Datas, 1) - LBound(Datas, 1) + 1
     
    Dim FuelFlow As Double
Dim ratio As Variant
Dim coucou As Integer
Dim Water As ProcessStream
Dim col As Integer, colo As Integer, bonnecol As Integer
Dim CCL As ProcessStream
Dim CCL2 As ProcessStream
Dim CCL3 As ProcessStream
Dim last_row As Integer, last_column As Integer
Dim Plo As String






For I = 0 To X - 2
    Datas(16, I) = "False"
Next

    NombreCC = 0
        



    Set hyBasis = simCase.BasisManager
    'As we just started, we begin in the Basis environment, we can check if this is ok
    If hyBasis.IsChangingBasis = False Then
        'Stop the solveur as Hysys wont calculate at each input
        simCase.Solver.CanSolve = False
        hyBasis.StartBasisChange
    End If
    
    
    Set hyFluidPkg = hyBasis.FluidPackages.Add("New FP")
        'We create the Fluid package that we will use
        hyFluidPkg.PropertyPackageName = "PengRob"
        'We use Peng-Robinson method
        Set hyPRFluidPkg = hyFluidPkg.PropertyPackage
        Dim hydropresent As Boolean
    hydropresent = False
    Dim LigneFluide As Integer
    LigneFluide = Sheets(cycleName).Range("C85").End(xlDown).Row
    For I = 86 To LigneFluide
            If Sheets(cycleName).Cells(I, 3) = "H2O" Or Sheets(cycleName).Cells(I, 3) = "Water" Then
                hydropresent = True
            End If
        'We add the fluids selected in the Fluid sheet
        With hyFluidPkg.Components
            Plo = Sheets(cycleName).Cells(I, 3)
            .Add Sheets(cycleName).Cells(I, 3)
        End With
    Next
    
    If hydropresent = False Then
        hyFluidPkg.Components.Add "H2O"
    End If
    

    'We create the reactions used in Combustion Chamber
            TheSet = "Set" & I
       
            'We create the reaction set (can englobe multiple reactions)
    Set SetReaction = hyBasis.ReactionPackageManager.ReactionSets.Add(TheSet)
    SetReaction.AssociateFluidPackage hyFluidPkg
    col1 = Sheets(cycleName).Range("B82").End(xlToRight).column + 3
    col2 = Sheets(cycleName).Cells(82, col1).End(xlToRight).column
    NombreReac = Sheets("GT Specs").Range("P9")
    For I = 1 To NombreReac
    

        
        'we create the inputed reaction : here a conversion reaction
        Set Reaction = SetReaction.ActiveReactions.Add("Rxn" & I, "conversionrxn")
        
        Reaction.Reactants.RemoveAll 'We delete former reactif and products just in case
        
        ligne = 8 + 7 * (I - 1) + 2
        Dim ligneReac As Integer
        ligneReac = Sheets("GT Specs").Cells(ligne, 14).End(xlDown).Row
    
        
        For j = ligne To ligneReac
            'We add differents reactifs, products, and stochiometric components
            Set hyReac = Reaction.Reactants.Add(Sheets("GT Specs").Cells(j, 14))
            hyReac.StoichiometricCoefficientValue = Sheets("GT Specs").Cells(j, 15)

        Next
        
    
    'We associate ReactionSet and FluidPackage
 
        'We parameter the other parameter
    If j = ligne + 1 Then
    Reaction.BaseComponent = hyReac
    End If
    Reaction.ReactionPhase = ptVapourPhase
    Reaction.ConversionCoefficientsValue = Array(100, 0, 0)
    Next

       
    'we change to the simulation environment
    If hyBasis.CanEndBasisChange = True Then
         hyBasis.EndBasisChange
     Else
         MsgBox "Couldn't finish Basis Change"
        Exit Function
     End If




    'Next we add the components, their parameters and connect them to the streams
    Aye = StreamCreation(Datas(), strcase, cycleName)
    Aye = ComponentCreation(Datas(), strcase, cycleName)



        '/////////////////////Compressor and Turbine Stage Design//////////////////////////////////
    'We  redimensionate Compressor and Turbine with the stage design method, to get different properties as number of stages,
    'tip speed, etc.... Code available in Module 5

    Dim myComp As Collection
    Set myComp = CompDesign(Datas(), strcase, cycleName)
    
    Dim myTurb As Collection
    Set myTurb = TurbDesign(Datas(), strcase, cycleName)
    
    
    
    
    '///////////////////////End of stage design////////////////////////////////////
    
    'As for the Compressor and turbine, a change in the stream input do not impact the output of cooler and heater because of how I coded the Delta. Therefore
    'I need to recalibrate every component when changing a parameter.
    
    Aye = HXRecalibration(Datas(), strcase, cycleName)
    

    '//////////////////////////End of the GT Creation//////////////////////////////
    
    'Once the turbine is deisgned and corrected, we will change some parameters to target the wanted power
   
    
    'Aye = HXRecalibration(Datas(), strcase)
    
        'We recalculate the Stage design because parameters have changed
    'Set myComp = CompDesign(Datas(), strcase)
    'Set myTurb = TurbDesign(Datas(), strcase)
    
    
    
    'We extract the results into a sheet "ResulsX" (Code available in module 4)
    
    Set Creation_Turbine = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
    
    

End Function
