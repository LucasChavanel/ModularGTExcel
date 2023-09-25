VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cThermalEff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTComp As Double
Private pTCC As Double
Private pTTurb As Double
Private pTHX As Double
Private pTAir As Double
Private pcycleType As String
Private pCycleName As String
Private pCool As Double
Private pHeat As Double
Property Get ThermalEff() As Double

If pCycleName = "Brayton1" Or pCycleName = "Brayton2" Or pCycleName = "Brayton3" Then
    If pcycleType = "Classic" Then
        If pTComp = 0.0001 Or pTCC = 0.0001 Or pTTurb = 0.0001 Then
            ThermalEff = 0
        Else
            ThermalEff = ((pTCC - pTComp) - (pTTurb - pTAir)) / (pTCC - pTComp)
        End If
    ElseIf pcycleType = "HX" Then
        If pTComp = 0.0001 Or pTCC = 0.0001 Or pTTurb = 0.0001 Then
            ThermalEff = 0
        Else
            ThermalEff = ((pTCC - pTComp) - (pTTurb - pTAir)) / (pTCC - pTHX)
        End If
    End If
ElseIf pCycleName = "Rankine1" Or pCycleName = "Rankine2" Or pCycleName = "Rankine3" Then
    If pHeat = 0.0001 Or pCool = 0.0001 Then
        ThermalEff = 0
    Else
        ThermalEff = 1 - pHeat / pCool
    End If
End If
End Property


Property Get cycleName() As String
     cycleName = pCycleName
End Property

Property Let cycleName(ocycleName As String)
     pCycleName = ocycleName
End Property

Property Get cycleType() As String
     cycleType = pcycleType
End Property

Property Let cycleType(ocycleType As String)
     pcycleType = ocycleType
End Property

Property Get TComp() As Double
     TComp = pTComp
End Property

Property Let TComp(oTComp As Double)
     pTComp = oTComp
End Property

Property Get TAir() As Double
     TAir = pTAir
End Property

Property Let TAir(oTAir As Double)
     pTAir = oTAir
End Property

Property Get Heat() As Double
     Heat = pHeat
End Property

Property Let Cool(oCool As Double)
     pCool = oCool
End Property

Property Get Cool() As Double
     Cool = pCool
End Property

Property Let Heat(oHeat As Double)
     pTAir = oHeat
End Property

Property Get TTurb() As Double
     TTurb = pTTurb
End Property

Property Let TTurb(oTTurb As Double)
     pTTurb = oTTurb
End Property

Property Get TCC() As Double
     TCC = pTCC
End Property

Property Let TCC(oTCC As Double)
     pTCC = oTCC
End Property

Property Get THX() As Double
     THX = pTHX
End Property

Property Let THX(oTHX As Double)
     pTHX = oTHX
End Property
