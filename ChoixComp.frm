VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChoixComp 
   Caption         =   "Components Choice"
   ClientHeight    =   10884
   ClientLeft      =   72
   ClientTop       =   312
   ClientWidth     =   8748.001
   OleObjectBlob   =   "ChoixComp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChoixComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COmp As String
Public Stream As String
Public CycleA As String
Public colonne As Integer



Private Sub ModifStream()
    
    ListBox_Input.Clear
    ListBox_Input2.Clear
    ListBox_Output.Clear
    ListBox_Output2.Clear
    Dim I As Integer
    Dim nbLignes As Integer
    
    nbLignes = Sheets("ListCompStream").Range("C1").End(xlDown).Row
    
    For I = 1 To nbLignes
        'On met � jour les 3 listes de streams
        Stream = Sheets("ListCompStream").Cells(I + 1, 3)
        ListBox_Input.AddItem Stream
        ListBox_Input2.AddItem Stream
        ListBox_Output.AddItem Stream
        ListBox_Output2.AddItem Stream
    Next

End Sub


Private Sub ModifCycle()
    
    ListBox_InputCycle.Clear

    Dim I As Integer
    Dim nbLignes As Integer
    
    nbLignes = Sheets("ListCompStream").Range("L1").End(xlDown).Row
    
    For I = 1 To nbLignes
        'On met � jour les 3 listes de streams
        CycleA = Sheets("ListCompStream").Cells(I + 1, 12)
        ListBox_InputCycle.AddItem CycleA
    Next

End Sub

Private Sub ModifQStream()
    ListBox_InputQ.Clear
    
    Dim I As Integer
    Dim nbLignes As Integer
    
    If Sheets("ListCompStream").Cells(2, 4) = "" Then
        'Meme manip que pour BoutonQStream
        Sheets("ListCompStream").Cells(2, 4) = "QTest"
        nbLignes = Sheets("ListCompStream").Range("D1").End(xlDown).Row - 1
        Sheets("ListCompStream").Cells(2, 4) = ""
    Else
        nbLignes = Sheets("ListCompStream").Range("D1").End(xlDown).Row
    End If
    
    For I = 1 To nbLignes
        'On ajoute � la liste visible les composants ajout�s
        QStream = Sheets("ListCompStream").Cells(I + 1, 4)
        ListBox_InputQ.AddItem QStream
    Next

End Sub



Private Sub Bouton_AjoutStream_Click()
    
    'Apres avoir rentr� le nom du stream, on mets � jour la liste dans la Sheet cach�e puis on met � jour la liste du formulaire
    Dim nblignes2 As Integer
    nblignes2 = Sheets("ListCompStream").Range("C1").End(xlDown).Row
    
    'Pas de s�curit� comme pour QStream car forcement au minimum 1 Input Stream et 1 Gas
    Sheets("ListCompStream").Cells(nblignes2 + 1, 3) = TextBox_Stream.Text
    ModifStream
    TextBox_Stream.Text = ""

End Sub

Private Sub BoutonCycle_Click()
    
    'Apres avoir rentr� le nom du stream, on mets � jour la liste dans la Sheet cach�e puis on met � jour la liste du formulaire
    'On met � jour liste dans feuille cach�e
    If Sheets("ListCompStream").Cells(2, 12) = "" Then
        'si aucun QStream ajout�, combine pour pas que �a bug
        Sheets("ListCompStream").Cells(2, 12) = "QTest"
        nbLignes = Sheets("ListCompStream").Range("L1").End(xlDown).Row - 1
        Sheets("ListCompStream").Cells(2, 12) = ""
    Else
        nbLignes = Sheets("ListCompStream").Range("L1").End(xlDown).Row
    End If
    
    'Pas de s�curit� comme pour QStream car forcement au minimum 1 Input Stream et 1 Gas
    Sheets("ListCompStream").Cells(nbLignes + 1, 12) = TextBoxCycle.Text
    ModifCycle
    TextBoxCycle.Text = ""

End Sub



Private Sub BoutonAjoutQStream_Click()

    'Lorsque que l'on clique sur le bouton Rajout QStream
    Dim nbLignes As Integer
    
    'On met � jour liste dans feuille cach�e
    If Sheets("ListCompStream").Cells(2, 4) = "" Then
        'si aucun QStream ajout�, combine pour pas que �a bug
        Sheets("ListCompStream").Cells(2, 4) = "QTest"
        nbLignes = Sheets("ListCompStream").Range("D1").End(xlDown).Row - 1
        Sheets("ListCompStream").Cells(2, 4) = ""
    Else
        nbLignes = Sheets("ListCompStream").Range("D1").End(xlDown).Row
    End If
    'On ajoute et on met � jour
    Sheets("ListCompStream").Cells(nbLignes + 1, 4) = TextBox_QStream.Text
    ModifQStream
    TextBox_QStream.Text = ""

End Sub

Private Sub BoutonNewComp_Click()

    'Lorsque tous les champs du composants sont compl�t�s mais que ce n'est pas le dernier du cycle, on enregistre ces donn�es dans le tableau et on recharge le formulaire
    
    'Securit� sur les champs
    If (NomComp.Visible = True And NomComp.Text = "") Or (ListBox_Input.Visible = True And ListBox_Input.Value = "") Or (NomComp.Visible = True And NomComp.Text = "") Or (NomComp.Visible = True And NomComp.Text = "") Or (NomComp.Visible = True And NomComp.Text = "") Or (NomComp.Visible = True And NomComp.Text = "") Or (NomComp.Visible = True And NomComp.Text = "") Or (NomComp.Visible = True And NomComp.Text = "") Then
        MsgBox "A field is empty"
    ElseIf (Parameter1.Visible = True And IsNumeric(Parameter1.Text) = False) Or (Parameter2.Visible = True And IsNumeric(Parameter2.Text) = False) Or (Parameter3.Visible = True And IsNumeric(Parameter3.Text) = False) Then
        MsgBox "One of the parameters is not a number"
    ElseIf composant_id = 2 And Sheets("ListCompStream").Range("J2") = "" Then
        MsgBox "Turbine Efficiency not validated"
    Else
        Enregistrer_comp
        Unload ChoixComp
        ChoixComp.Show
        Sheets("ListCompStream").Range("J2:J6").Delete
    End If
End Sub

Private Sub BoutonTerminer_Click()

    'Lorsque tous les champs sont remplis et que c'est le dernier composant, on enregistre dans le tableau et on ferme le formulaire
    
    'Securit� sur les champs
    If (NomComp.Visible = True And NomComp.Text = "") Or TurbSpecButton.Visible = True Or (ListBox_Input.Visible = True And ListBox_Input.ListIndex = -1) Or (ListBox_Input2.Visible = True And ListBox_Input2.ListIndex = -1) Or (ListBox_Output.Visible = True And ListBox_Output.ListIndex = -1) Or (ListBox_InputQ.Visible = True And ListBox_InputQ.ListIndex = -1) Or (ListBox_Output2.Visible = True And ListBox_Output2.ListIndex = -1) Or (Parameter1.Visible = True And Parameter1.Text = "") Or (Parameter2.Visible = True And Parameter2.Text = "") Or (Parameter3.Visible = True And Parameter3.Text = "") Or (Parameter4.Visible = True And Parameter4.Text = "") Or (ComboBoxReac.Visible = True And ComboBoxReac.SelText = "") Or (ComboBox1.Visible = True And ComboBox1.SelText = "") Or (ListBox_InputCycle.Text = "") Then
        MsgBox "A field is empty"
    ElseIf (Parameter1.Visible = True And IsNumeric(Parameter1.Text) = False) Or (Parameter2.Visible = True And IsNumeric(Parameter2.Text) = False) Or (Parameter3.Visible = True And IsNumeric(Parameter3.Text) = False) Then
        MsgBox "One of the parameters is not a number"
    ElseIf composant_id = 2 And Sheets("ListCompStream").Range("J2") = "" Then
        MsgBox "Turbine Efficiency not validated"
    ElseIf composant_id = 1 And Sheets("ListCompStream").Range("I2") = "" Then
        MsgBox "Compressor Efficiency not validated"
    ElseIf composant_id = 9 And ((Parameter1.Text + Parameter2.Text) <> "1") Then
        MsgBox "Sum of Split must be equal to one"
    Else
        Enregistrer_comp
        Unload ChoixComp
        Sheets("ListCompStream").Range("M2:M60").Clear
        CyclesChoose.Show
        
    End If


End Sub

Private Sub ComboBox_Comp_Change()

    'Lorsque l'on choisi un composant,
    Dim composant_id As Integer
    composant_id = ComboBox_Comp.ListIndex + 1
    
    'On rend les cases communes � tous les composants visibles
    LabelInput.Visible = True
    ListBox_Input.Visible = True
    LabelOutput.Visible = True
    ListBox_Output.Visible = True
    TextBox_Stream.Visible = True
    LabelChoixStream.Visible = True
    Bouton_AjoutStream.Visible = True
    
    LabelNomComp.Visible = True
    NomComp.Visible = True
    
    LabelCycle.Visible = True
    TextBoxCycle.Visible = True
    BoutonCycle.Visible = True

    
    Label2.Visible = True
    ListBox_InputCycle.Visible = True
    
    
    'Rendre les 3 choix de Streams Invisble
    If ListBox_InputQ.Visible = True Then
        ListBox_InputQ.Visible = False
        LabelQStream.Visible = False
    End If
    
    If ListBox_Input2.Visible = True Then
        ListBox_Input2.Visible = False
        LabelInput2.Visible = False
    End If
    
    If ListBox_Output2.Visible = True Then
        ListBox_Output2.Visible = False
        LabelOutput2.Visible = False
    End If
    
    'Rendre les boutons d'ajout de Qstream Invisible
    If TextBox_QStream.Visible = True Then
        TextBox_QStream.Visible = False
        LabelChoixQStream.Visible = False
        BoutonAjoutQStream.Visible = False
    End If
    
    If LabelReac.Visible = True Then
        LabelReac.Visible = False
        ComboBoxReac.Visible = False
    End If
    
    'Rendre les box d'entr�e de param�tres invisibles
    If LabelP1.Visible = True Then
        LabelP1.Visible = False
        Parameter1.Visible = False
    End If
    
    
    If LabelP2.Visible = True Then
        LabelP2.Visible = False
        Parameter2.Visible = False
    End If
    
    If LabelP3.Visible = True Then
        LabelP3.Visible = False
        Parameter3.Visible = False
    End If
    
    If LabelP4.Visible = True Then
        LabelP4.Visible = False
        Parameter4.Visible = False
    End If
    
    If Label3.Visible = True Then
        Label3.Visible = False
        ComboBox1.Visible = False
    End If
    
    '////////////////////////////////////Ajout des param�tres pour les composants/////////////////////////////
    'On met � jour les Streams
    ModifStream
    ModifCycle
    
    NomComp.Text = ""
    'Si on choisi un compresseur ou une turbine
    If composant_id = 1 Then
            'On rend les listes correspondantes visibles
        LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        LabelChoixQStream.Visible = True
        BoutonAjoutQStream.Visible = True
        
        'On rend les choix de param�tres visible et on change le label
        LabelP1.Visible = True
        LabelP1.Caption = "Pressure Ratio"
        Parameter1.Visible = True
        LabelP2.Visible = True
        LabelP2.Caption = "Rotating Speed (RPM)"
        Parameter2.Visible = True
        
        LabelP4.Caption = "Isentropic Efficiency (%)"
        TurbSpecButton.Visible = True
        TurbSpecButton.Caption = "Parameter Efficiency"
        
        ComboBoxReac.Clear
        LabelReac.Visible = True
        LabelReac.Caption = "Is this the last of a Compressor Chain?"
        ComboBoxReac.Visible = True
        ComboBoxReac.AddItem "Yes"
        ComboBoxReac.AddItem "No"
        
        
        
        'On met � jour la liste de QSTream
        ModifQStream
    End If
    
    
    If composant_id = 2 Then
    'On rend les listes correspondantes visibles
        LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        LabelChoixQStream.Visible = True
        BoutonAjoutQStream.Visible = True
        
        'On rend les choix de param�tres visible et on change le label
        LabelP1.Visible = True
        LabelP1.Caption = "Pressure Ratio"
        Parameter1.Visible = True
        LabelP2.Visible = True
        LabelP2.Caption = "Number Of Stage"
        Parameter2.Visible = True
        LabelP3.Visible = True
        LabelP3.Caption = "Rotating Speed (RPM)"
        Parameter3.Visible = True
        LabelP4.Visible = True
        LabelP4.Caption = "Isentropic Efficiency (%)"
        TurbSpecButton.Visible = True
        TurbSpecButton.Caption = "Parameter Efficiency"
        
        ComboBoxReac.Clear
        LabelReac.Visible = True
        LabelReac.Caption = "Is this the last of an Expander Chain?"
        ComboBoxReac.Visible = True
        ComboBoxReac.AddItem "Yes"
        ComboBoxReac.AddItem "No"
        
        
        'On met � jour la liste de QSTream
        ModifQStream
            
    End If
    
    If composant_id = 7 Then
        'On rend les listes correspondantes visibles
        LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        LabelChoixQStream.Visible = True
        BoutonAjoutQStream.Visible = True
        
        'On rend les choix de param�tres visible et on change le label
        LabelP1.Visible = True
        LabelP1.Caption = "Pressure Ratio"
        Parameter1.Visible = True
        
        'On met � jour la liste de QSTream
        ModifQStream
    End If
    
    'Comme au dessus
    If composant_id = 4 Or composant_id = 6 Then
    LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        BoutonAjoutQStream.Visible = True
        LabelChoixQStream.Visible = True
        
        LabelP1.Visible = True
        LabelP1.Caption = "DP (%)"
        Parameter1.Visible = True
        
        LabelP2.Visible = True
        LabelP2.Caption = "Outlet Temperature (K)"
        Parameter2.Visible = True
        
        ModifQStream
    End If
        
    'Comme au dessus
    If composant_id = 3 Then
        LabelInput2.Visible = True
        LabelInput2.Caption = "Choice of the Gas Stream"
        ListBox_Input2.Visible = True
        
        LabelP1.Visible = True
        LabelP1.Caption = "DP (%)"
        Parameter1.Visible = True
    
        Label3.Visible = True
        Label3.Caption = "Reaction Associated"
        ComboBox1.Visible = True
        ComboBox1.Clear
        col1 = Sheets("Fluids").Range("A7").End(xlToRight).column + 3
        col2 = Sheets("Fluids").Cells(7, col1).End(xlToRight).column
        NombreReac = (col2 - col1 + 1) / 2
        For I = 0 To NombreReac - 1
         ComboBox1.AddItem Sheets("Fluids").Cells(6, col1 + (I * 2 + 1) - 1)
        Next
        
        ComboBoxReac.Clear
        LabelReac.Visible = True
        LabelReac.Caption = "Is this the last of an Heating Process?"
        ComboBoxReac.Visible = True
        ComboBoxReac.AddItem "Yes"
        ComboBoxReac.AddItem "No"
        
        LabelP3.Visible = True
        LabelP3.Caption = "T Out (K)"
        Parameter3.Visible = True
        
    End If
    
    'Comme au dessus
    If composant_id = 5 Then
        
        LabelInput2.Visible = True
        LabelInput2.Caption = "Choice of the Hot Input Stream"
        ListBox_Input2.Visible = True
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
        LabelOutput2.Visible = True
        ListBox_Output2.Visible = True
        
        LabelP1.Visible = True
        LabelP1.Caption = "DP Shell Side (kPa)"
        Parameter1.Visible = True
        LabelP2.Visible = True
        LabelP2.Caption = "DP Tube Side (kPa)"
        Parameter2.Visible = True
        LabelP3.Visible = True
        LabelP3.Caption = "T Out (�K)"
        Parameter3.Visible = True
        
        LabelReac.Visible = True
        LabelReac.Caption = "Is this the last of a Regeneration HX?"
        ComboBoxReac.Clear
        ComboBoxReac.Visible = True
        ComboBoxReac.AddItem "Yes"
        ComboBoxReac.AddItem "No"
        
        Label3.Visible = True
        Label3.Caption = "What Kind of HX is this?"
        ComboBox1.Clear
        ComboBox1.Visible = True
        ComboBox1.AddItem "Regeneration"
        ComboBox1.AddItem "Boiler: Saturated Liquid"
        ComboBox1.AddItem "Boiler: Saturated Steam"
        ComboBox1.AddItem "Boiler: Superheated Steam"
        ComboBox1.AddItem "Reheat"
        ComboBox1.AddItem "Other HX"

    End If
    
    If composant_id = 8 Then
    
        LabelInput2.Visible = True
        LabelInput2.Caption = "Choice of the Second Input Stream"
        ListBox_Input2.Visible = True
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
    
    End If
    
    If composant_id = 9 Then

        LabelOutput.Visible = True
        ListBox_Output.Visible = True
            LabelOutput2.Visible = True
        ListBox_Output2.Visible = True
        LabelP1.Visible = True
        LabelP1.Caption = "Split Fraction of first stream"
        Parameter1.Visible = True
        LabelP2.Visible = True
        LabelP2.Caption = "Split Fraction of second stream"
        Parameter2.Visible = True
    
    End If
    
    If composant_id = 10 Then
    'On rend les listes correspondantes visibles
        LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        LabelChoixQStream.Visible = True
        BoutonAjoutQStream.Visible = True
        
        'On rend les choix de param�tres visible et on change le label
        LabelP1.Visible = True
        LabelP1.Caption = "Pressure Ratio"
        Parameter1.Visible = True
        LabelP2.Visible = True
        LabelP2.Caption = "Isentropic Efficiency"
        Parameter2.Visible = True
        
        'On met � jour la liste de QSTream
        ModifQStream
            
    End If
    
    If composant_id = 11 Then 'Solar Heater
        
        LabelInput.Visible = False
        ListBox_Input.Visible = False
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
        LabelQStream.Visible = True
        ListBox_InputQ.Visible = True
        TextBox_QStream.Visible = True
        BoutonAjoutQStream.Visible = True
        
        LabelP1.Visible = True
        LabelP1.Caption = "Efficiency Of the Collector (%)"
        Parameter1.Visible = True
        
        LabelP2.Visible = True
        LabelP2.Caption = "Size Of the collector (m2)"
        Parameter2.Visible = True
        
        LabelP3.Visible = True
        LabelP3.Caption = "Total Solar Radiation Intensity (W/m2)"
        Parameter3.Visible = True
        
        
        LabelP4.Visible = True
        LabelP4.Caption = "Input Stream"
        TurbSpecButton.Visible = True
        TurbSpecButton.Caption = "Parameter Stream"
        
        ModifQStream
    End If
    
     If composant_id = 12 Then 'Saturator
        
        LabelOutput.Visible = True
        ListBox_Output.Visible = True

        
        LabelP1.Visible = True
        LabelP1.Caption = "Humidity wanted (%)"
        Parameter1.Visible = True
        
        
        ModifQStream
    End If
    
    
    If composant_id = 13 Then 'Fired Heater
    
        LabelInput2.Visible = True
        LabelInput2.Caption = "Fuel Input"
        ListBox_Input2.Visible = True
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
        LabelOutput.Caption = "Feed Output"
        LabelOutput2.Visible = True
        ListBox_Output2.Visible = True
        LabelOutput2.Caption = "Flare Output"
        
        LabelP1.Visible = True
        LabelP1.Caption = "Efficiency (%)"
        Parameter1.Visible = True
        
        LabelP2.Visible = True
        LabelP2.Caption = "Excess Air (%)"
        Parameter2.Visible = True
        
        LabelP3.Visible = True
        LabelP3.Caption = "Feed Output Temperature"
        Parameter3.Visible = True
        
        ComboBoxReac.Clear
        LabelReac.Visible = True
        LabelReac.Caption = "Is this the last of an Heating Process?"
        ComboBoxReac.Visible = True
        ComboBoxReac.AddItem "Yes"
        ComboBoxReac.AddItem "No"
        
        ModifQStream
    End If
    
    If composant_id = 14 Then 'Tank
    
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
        LabelOutput.Caption = "Liquid Output"
        LabelOutput2.Visible = True
        ListBox_Output2.Visible = True
        LabelOutput2.Caption = "Vapour Output"
        
        LabelP1.Visible = True
        LabelP1.Caption = "Inlet Pressure Drop (%)"
        Parameter1.Visible = True
        
        ModifQStream
    End If
    
    If composant_id = 15 Then 'Flash
    
        LabelOutput.Visible = True
        ListBox_Output.Visible = True
        LabelOutput.Caption = "Liquid Output"
        LabelOutput2.Visible = True
        ListBox_Output2.Visible = True
        LabelOutput2.Caption = "Vapour Output"
        
        LabelP1.Visible = True
        LabelP1.Caption = "Inlet Pressure Drop (%)"
        Parameter1.Visible = True
        
        ModifQStream
    End If
    
    If composant_id = 16 Then
        
        LabelP1.Visible = True
        LabelP1.Caption = "DP (%)"
        Parameter1.Visible = True
    End If
    '///////////////////////////Fin de l'ajout des composants////////////////////////////

End Sub



Private Sub BoutonAnnuler_Click()

    Unload ChoixComp
    
    'Supprimer les streams et le tableau

End Sub






Private Sub ListBox_InputCycle_Click()

End Sub

Private Sub UserForm_Initialize()

    Dim I As Integer
    Dim nbLignes As Integer
    
    
'    If Sheets("Fluids").Range("E8") <> "" Then
'        Sheets("ListCompStream").Cells(4, 3) = Sheets("Fluids").Range("E8")
'        If Sheets("Fluids").Range("F8") <> "" Then
'            Sheets("ListCompStream").Cells(5, 3) = Sheets("Fluids").Range("F8")
'        End If
'    ElseIf Sheets("Fluids").Range("F8") <> "" Then
'            Sheets("ListCompStream").Cells(4, 3) = Sheets("Fluids").Range("F8")
'    End If
'
    nbLignes = Sheets("ListCompStream").Range("B1").End(xlDown).Row
    'On ajoute les composants � la comboBox, la liste de composant doit �tre modifi�e
    'A la main en cas de rajout
    For I = 1 To nbLignes - 1
        COmp = Sheets("ListCompStream").Cells(I + 1, 2)

        ComboBox_Comp.AddItem Sheets("ListCompStream").Cells(I + 1, 2)
    Next
    

    
    

End Sub

Private Sub Enregistrer_comp()

    'On enregistre les diff�rentes valeurs dans les cases appropri�es du tableau
    Dim nbLignes As Integer
    composant_id = ComboBox_Comp.ListIndex + 1
    colonne = Sheets("Fired Rankine").Range("A6").End(xlToRight).column
    Dim col As Integer
    '///////////////////////Remplissage en fonction des composants////////////////////////////:
    
    If composant_id = 1 Then 'Compressor
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Compressor"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter4.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
        
        
        If Sheets("ListCompStream").Range("P2") = "" Then
        'Meme manip que pour BoutonQStream
            Sheets("ListCompStream").Range("P2") = "QTest"
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row
            Sheets("ListCompStream").Range("P2") = ""
        Else
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row + 1
        End If
        
        If ComboBoxReac.Text = "Yes" Then
            Sheets("ListCompStream").Range("P" & nbLignes) = ListBox_InputCycle.Value
            Sheets("ListCompStream").Range("Q" & nbLignes) = NomComp.Text
            Sheets("ListCompStream").Range("R" & nbLignes) = "Compressor"
        End If
        
        col = Sheets("Constant Parameters").Range("A14").End(xlToRight).column + 1
        Sheets("Constant Parameters").Cells(14, col) = NomComp.Text
        Sheets("Constant Parameters").Cells(14, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(15, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(16, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(17, col).Borders.Weight = xlThin

        
    ElseIf composant_id = 2 Then 'Turbine

        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Gas Turbine"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter4.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = CDbl(Parameter3.Text)
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
       
 
        
       
        If Sheets("ListCompStream").Range("P2") = "" Then
        'Meme manip que pour BoutonQStream
            Sheets("ListCompStream").Range("P2") = "QTest"
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row
            Sheets("ListCompStream").Range("P2") = ""
        Else
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row + 1
        End If
        
        
        If ComboBoxReac.Text = "Yes" Then
            Sheets("ListCompStream").Range("P" & nbLignes) = ListBox_InputCycle.Value
            Sheets("ListCompStream").Range("Q" & nbLignes) = NomComp.Text
            Sheets("ListCompStream").Range("R" & nbLignes) = "Gas Turbine"
        End If
        
        
        col = Sheets("Constant Parameters").Range("A6").End(xlToRight).column + 1
        Sheets("Constant Parameters").Cells(6, col) = NomComp.Text
        Sheets("Constant Parameters").Cells(12, col) = CDbl(Parameter2.Text)
        Sheets("Constant Parameters").Cells(7, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(8, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(9, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(10, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(11, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(12, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(6, col).Borders.Weight = xlThin
        
        
    ElseIf composant_id = 3 Then 'Combustion Chamber
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Combustion Chamber"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = ListBox_Input2.Value
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = (ComboBox1.Value)
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
                If Sheets("ListCompStream").Range("P2") = "" Then
        'Meme manip que pour BoutonQStream
            Sheets("ListCompStream").Range("P2") = "QTest"
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row
            Sheets("ListCompStream").Range("P2") = ""
        Else
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row + 1
        End If
        
        If ComboBoxReac.Text = "Yes" Then
            Sheets("ListCompStream").Range("P" & nbLignes) = ListBox_InputCycle.Value
            Sheets("ListCompStream").Range("Q" & nbLignes) = NomComp.Text
            Sheets("ListCompStream").Range("R" & nbLignes) = "Combustion Chamber"
        End If
    
    
    ElseIf composant_id = 4 Then 'Cooler
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Cooler"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    
    ElseIf composant_id = 5 Then 'HX
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Heat Exchanger"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = ListBox_Input2.Value
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = CDbl(Parameter3.Text)
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
                If Sheets("ListCompStream").Range("P2") = "" Then
        'Meme manip que pour BoutonQStream
            Sheets("ListCompStream").Range("P2") = "QTest"
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row
            Sheets("ListCompStream").Range("P2") = ""
        Else
            nbLignes = Sheets("ListCompStream").Range("P1").End(xlDown).Row + 1
        End If
        
        If ComboBoxReac.Text = "Yes" Then
            Sheets("ListCompStream").Range("P" & nbLignes) = ListBox_InputCycle.Value
            Sheets("ListCompStream").Range("Q" & nbLignes) = NomComp.Text
            Sheets("ListCompStream").Range("R" & nbLignes) = "Heat Exchanger"
        End If
        
        col = Sheets("Constant Parameters").Range("A28").End(xlToRight).column + 1
        Sheets("Constant Parameters").Cells(28, col) = NomComp.Text
        Sheets("Constant Parameters").Cells(29, col) = ComboBox1.SelText

        
    ElseIf composant_id = 6 Then 'Heater
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Heater"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
        
    ElseIf composant_id = 7 Then 'Pump
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Pump"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    ElseIf composant_id = 8 Then 'Mixer
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Mixer"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = ListBox_Input2.Value
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
    
    ElseIf composant_id = 9 Then 'Splitter
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Splitter"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    ElseIf composant_id = 10 Then 'Steam Turbine
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Steam Turbine"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    ElseIf composant_id = 11 Then 'Solar Heater
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Solar Heater"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = NomComp.Text & "SI"
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = ListBox_InputQ.Value
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
        col = Sheets("Constant Parameters").Range("A19").End(xlToRight).column + 1
        Sheets("Constant Parameters").Cells(19, col) = NomComp.Text
        Sheets("Constant Parameters").Cells(20, col) = CDbl(Parameter1.Text)
        Sheets("Constant Parameters").Cells(21, col) = CDbl(Parameter2.Text)
        Sheets("Constant Parameters").Cells(22, col) = CDbl(Parameter3.Text)
    
    ElseIf composant_id = 12 Then 'Stream Saturator
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Stream Saturator"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = NomComp.Text & "SI"
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    ElseIf composant_id = 13 Then 'Fired Heater
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Fired Heater"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = ListBox_Input2.Value
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter2.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = CDbl(Parameter3.Text)
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
     ElseIf composant_id = 14 Then 'Tank
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Tank"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
        If ComboBoxReac.Text = "Yes" Then
            Sheets("ListCompStream").Range("P" & nbLignes) = ListBox_InputCycle.Value
            Sheets("ListCompStream").Range("Q" & nbLignes) = NomComp.Text
            Sheets("ListCompStream").Range("R" & nbLignes) = "Compressor"
        End If
        
        col = Sheets("Constant Parameters").Range("A14").End(xlToRight).column + 1
        Sheets("Constant Parameters").Cells(14, col) = NomComp.Text
        Sheets("Constant Parameters").Cells(14, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(15, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(16, col).Borders.Weight = xlThin
        Sheets("Constant Parameters").Cells(17, col).Borders.Weight = xlThin

    
    ElseIf composant_id = 15 Then 'Tank
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Flash"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(16, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    ElseIf composant_id = 15 Then 'Tank
        Sheets("Fired Rankine").Cells(6, colonne + 1) = "Valve"
        Sheets("Fired Rankine").Cells(7, colonne + 1) = NomComp.Text
        Sheets("Fired Rankine").Cells(8, colonne + 1) = ListBox_Input.Value
        Sheets("Fired Rankine").Cells(9, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(10, colonne + 1) = ListBox_Output.Value
        Sheets("Fired Rankine").Cells(11, colonne + 1) = ListBox_Output2.Value
        Sheets("Fired Rankine").Cells(12, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(14, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(15, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(16, colonne + 1) = CDbl(Parameter1.Text)
        Sheets("Fired Rankine").Cells(17, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(18, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(13, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(19, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(20, colonne + 1) = 0
        Sheets("Fired Rankine").Cells(21, colonne + 1) = ListBox_InputCycle.Value
        
    End If
    
    Sheets("Fired Rankine").Cells(6, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(7, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(8, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(9, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(10, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(11, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(12, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(13, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(14, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(15, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(16, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(17, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(18, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(19, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(20, colonne + 1).Borders.Weight = xlThin
    Sheets("Fired Rankine").Cells(21, colonne + 1).Borders.Weight = xlThin


    
    '//////////////////////////////////Fin des composants////////////////////////////
    
    With Sheets("Fired Rankine")
        .Columns(colonne + 1).AutoFit
    End With

End Sub
Private Sub TurbSpecButton_Click()
    Dim composant_id As Integer
    composant_id = ComboBox_Comp.ListIndex + 1
    If composant_id = 2 Then
        TurbSpec.Show
        Parameter4.Visible = True
    ElseIf composant_id = 1 Then
        CompSpec.Show
        Parameter4.Visible = True
    ElseIf composant_id = 11 Then
        SolarStream.Show
    End If

    TurbSpecButton.Visible = False
    
End Sub
