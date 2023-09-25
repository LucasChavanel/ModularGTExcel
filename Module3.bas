Attribute VB_Name = "Module3"
Public hyApp As HYSYS.Application
Public simCase As SimulationCase
Public hyBasis As HYSYS.BasisManager
Public hyFluidPkg As HYSYS.FluidPackage
  Public cycleName As String


'this Module Contains the subs who will launch the process



Sub LaunchRun()
Sheets("Results").Range("A5:Z400").Clear

Dim StartTime As Double
Dim SecondsElapsed As Double
Dim a As Integer, b As Integer
Dim colonne As Integer
StartTime = Timer

'When the GT is designed in Excel we click on the "Launch Run" button wich calls this sub

ligne = Sheets("GT Specs").Range("F18").End(xlDown).Row
For p = 19 To ligne
    cycleName = Sheets("GT Specs").Cells(p, 6)
    If cycleName <> "ORC Rankine" Then
        Sheets(cycleName).Range("B91:D120").Clear
    Else
        Sheets(cycleName).Range("B92:D120").Clear
    End If
      Sheets(cycleName).Range("G81:H100").Clear
    Sheets(cycleName).Range("C33") = Sheets("GT Specs").Range("D9")
    Sheets(cycleName).Range("C34") = Sheets("GT Specs").Range("D10")
    Sheets(cycleName).Range("C41") = Sheets("GT Specs").Range("G10")
    Sheets(cycleName).Range("C42") = Sheets("GT Specs").Range("G9")
    Sheets(cycleName).Range("C35") = Sheets("GT Specs").Range("D11")
    Sheets(cycleName).Range("C36") = Sheets("GT Specs").Range("D12")
    
    Sheets(cycleName).Range("C43") = Sheets("GT Specs").Range("G14")
    Sheets(cycleName).Range("E40") = Sheets("GT Specs").Range("G11")
    Sheets(cycleName).Range("E41") = Sheets("GT Specs").Range("G12")
    Sheets(cycleName).Range("E42") = Sheets("GT Specs").Range("G13")
    Sheets(cycleName).Range("E43") = Sheets("GT Specs").Range("G15")
    
    If Sheets("GT Specs").Range("Z10") <> "" Then 'Solar Panel
        Sheets(cycleName).Range("D65") = Sheets("GT Specs").Range("G9")
        Sheets(cycleName).Range("D66") = Sheets("GT Specs").Range("G10")
        Sheets(cycleName).Range("D67") = Sheets("GT Specs").Range("G11")
        Sheets(cycleName).Range("D68") = Sheets("GT Specs").Range("G12")
        Sheets(cycleName).Range("D69") = Sheets("GT Specs").Range("G13")
        Sheets(cycleName).Range("D70") = Sheets("GT Specs").Range("G14")
    End If
    
    
    ligne2 = Sheets("GT Specs").Range("J9").End(xlDown).Row
    col = Sheets(cycleName).Range("B82").End(xlToRight).column
    ligne3 = Sheets(cycleName).Range("C82").End(xlDown).Row
    For j = 13 To ligne2
        Sheets(cycleName).Cells(ligne3 + 1, 3) = Sheets("GT Specs").Cells(j, 10)
        
        For k = 4 To col
            Sheets(cycleName).Cells(ligne3 + 1, k) = 0
        Next
        ligne3 = Sheets(cycleName).Range("C82").End(xlDown).Row
    Next
    
   
    With Sheets("GT Specs")
        .Range(.Cells(8, 10), .Cells(12, 12)).Copy Sheets(cycleName).Cells(ligne3 + 3, 3)
    End With
    Sheets(cycleName).Cells(ligne3 + 8, 3) = "Oxygen"
    Sheets(cycleName).Cells(ligne3 + 9, 3) = "Nitrogen"
    Sheets(cycleName).Cells(ligne3 + 10, 3) = "H2O"
    Sheets(cycleName).Cells(ligne3 + 11, 3) = "CO2"
    Sheets(cycleName).Cells(ligne3 + 12, 3) = "CO"
    If cycleName = "ORC Rankine" Then
        Sheets(cycleName).Cells(ligne3 + 13, 3) = "i-C5"
        Sheets(cycleName).Cells(ligne3 + 13, 4) = 0
        Sheets(cycleName).Cells(ligne3 + 13, 5) = 0
    End If
    Sheets(cycleName).Cells(ligne3 + 8, 4) = 0 '.0021
    Sheets(cycleName).Cells(ligne3 + 9, 4) = 0 '.0079
    Sheets(cycleName).Cells(ligne3 + 10, 4) = 0
    Sheets(cycleName).Cells(ligne3 + 11, 4) = 0 '.35 '0.044
    Sheets(cycleName).Cells(ligne3 + 12, 4) = 0
    Sheets(cycleName).Cells(ligne3 + 8, 5) = 0
    Sheets(cycleName).Cells(ligne3 + 9, 5) = 0
    Sheets(cycleName).Cells(ligne3 + 10, 5) = 0
    Sheets(cycleName).Cells(ligne3 + 11, 5) = 0
    Sheets(cycleName).Cells(ligne3 + 12, 5) = 0
    
    'Sheets(cycleName).Range(Cells(ligne3 + 8, 2), Cells(ligne3 + 13, 4)).Borders.Weight = xlThin
    
     With Sheets("GT Specs")
      If cycleName = "ORC Rankine" Then
        .Range(.Cells(13, 10), .Cells(ligne2, 12)).Copy Sheets(cycleName).Cells(ligne3 + 14, 3)
      Else
        .Range(.Cells(13, 10), .Cells(ligne2, 12)).Copy Sheets(cycleName).Cells(ligne3 + 13, 3)
     End If
    End With
    
    ligne5 = Sheets("GT Specs").Range("N8").End(xlDown).Row
     With Sheets("GT Specs")
        .Range(.Cells(8, 14), .Cells(ligne5, 15)).Copy Sheets(cycleName).Cells(81, col + 3)
    End With


    Sheets(cycleName).Range("C38") = Sheets("GT Specs").Range("C14")
    Sheets(cycleName).Range("C40") = Sheets("GT Specs").Range("C13")
    Sheets(cycleName).Range("D83") = Sheets("GT Specs").Range("D11")
    Sheets(cycleName).Range("D68") = Sheets("GT Specs").Range("R12")
    Sheets(cycleName).Range("D69") = Sheets("GT Specs").Range("R13")

    

    
Dim LastResults As Integer
LastResults = ActiveWorkbook.Sheets.count - 1
Sheets("Results").Activate


    
 
    


  
  
    
    last_row = Sheets(cycleName).Range("C10").End(xlDown).Row
    If Sheets(cycleName).Range("D10") = 0 Then
        last_column = 3
    Else
        last_column = Sheets(cycleName).Range("C10").End(xlToRight).column
    End If

'We create a database with the Components table, we will use this database to input components into Aspen
    Dim Datas()
    ReDim Datas(16, last_column - 2)
    For j = 0 To last_column - 2
        For I = 0 To 15
            Datas(I, j) = Sheets(cycleName).Cells(I + 10, j + 3)
         Next
            Datas(16, j) = "False"
    Next
    'If everything is parametered properly, we can launch the creation of the cycle (code in module 1)

    Dim myCompoCollec As Collection
    Set myCompoCollec = Creation_Turbine(Datas(), cycleName)
   
    Dim myCycleCollec As Collection
    Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
    Dim oCycle As cCycle
    Dim val As String

    If Sheets(cycleName).Range("C38") = "Design One Cycle" Then
    
    If cycleName <> "Fired Rankine" Then
       Aye = HXRecalibration(Datas(), strcase, cycleName)
    End If
        If Sheets(cycleName).Range("C40") = "Fixed Power" Then
            Aye = ApproxGTPower(myCompoCollec, myCycleCollec, Datas(), strcase, cycleName)
        
'            If cycleName <> "Fired Rankine" Then
'               Aye = HXRecalibration(Datas(), strcase, cycleName)
'            End If
         End If
            Dim myComp As Collection
            Set myComp = CompDesign(Datas(), strcase, cycleName)
    
            Dim myTurb As Collection
            Set myTurb = TurbDesign(Datas(), strcase, cycleName)
       
        Set myCompoCollec = CreateCompObject(Datas(), strcase, myComp, myTurb, cycleName)
        Set myCycleCollec = DefCycle(Datas(), strcase, myCompoCollec, cycleName)
        
        val = ExtractResults(myCompoCollec, strcase, myCycleCollec, Datas())
        
    ElseIf Sheets(cycleName).Range("C38") = "Find Optimum Point" Then
   
        val = ChangePressureRatio(myCompoCollec, strcase, myCycleCollec, Datas(), cycleName)
        
    End If


Next
    SecondsElapsed = Round(Timer - StartTime, 2)
        
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds"











End Sub

'Sub who will open the Gas parameter Userform when clicked on the "Define GT parameters" button
Sub GTParametersForm()
    
    GTParameters.Show

End Sub


