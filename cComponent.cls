VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pPin As Double
Private pPout As Double
Private pPin2 As Double
Private pPout2 As Double

Private pTin As Double
Private pTout As Double
Private pTin2 As Double
Private pTout2 As Double

Private phin As Double
Private phout As Double
Private phin2 As Double
Private phout2 As Double

Private pFin As Double
Private pFout As Double
Private pFin2 As Double
Private pFout2 As Double
Private pFmin As Double

Private pCp As Double
Private pHHV As Double
Private ph As Double

Private pPower As Double
Private pEfficiency As Double
Private pDeltaT As Double
Private pDeltaP1 As Double
Private pDeltaP2 As Double

Private pPressureRatio As Double
Private pReaction As String
Private pCollectorSize As Double
Private pSolarRadiation As Double
Private pExtraPercentage As Double
Private pExtraPercentage2 As Double

Private pCycleName As String
Private pCompName As String
Private pCompType As String
Private pHXType As String
Private pCycleSheet As String

Private pTipSpeed As Double
Private pRotatingSpeed As Double
Private pNumberStage As Double
Private pDiameter As Double

Private iI As Integer
Private iNy As Integer
Private pPhiM As Double
Private iOH As Integer

Private pLastComponent As Boolean

Private iindex As Integer

Property Get PEC() As Double
    If pCompType = "Compressor" Then
        If pPressureRatio > 0 Then

            PEC = 71.1 * pFin / (0.9 - (pEfficiency / 100)) * (pPressureRatio) * WorksheetFunction.Ln(pPressureRatio)
        Else
            PEC = 0
        End If
    ElseIf pCompType = "Gas Turbine" Then
        If pPin <> 0 And pPout <> 0 And pPin / pPout > 0 Then
            PEC = (479.34 * pFin) / (0.95 - pEfficiency / 100) * WorksheetFunction.Ln(pPin / pPout) * (1 + Exp(0.036 * (pTin) - 54.4))
        Else
            PEC = 0
        End If
    ElseIf pCompType = "Combustion Chamber" Or pCompType = "Fired Heater" Then
        If pPin <> 0 And pPin / pPout > 0 Then
            If pPout / pPin = 1 Then
                PEC = ((46.08 * pFin) / (0.995 - 0.96)) * (1 + Exp(0.018 * (pTout) - 26.4))
            Else
                PEC = ((46.08 * pFin) / (0.995 - pPout / pPin)) * (1 + Exp(0.018 * (pTout) - 26.4))
            End If
           
        Else
            PEC = 0
        End If
    ElseIf pCompType = "Heat Exchanger" Then
        If pTin2 - pTout2 = 0 Or pTout - pTin = 0 Then
            PEC = 0
        ElseIf (pFin * (pPin2 - pPout2) / (18 * ((pTin2 - pTout2) / (pTout - pTin))) < 0) Then
             PEC = 0
        Else
            PEC = 4122 * (pFin * (pPin2 - pPout2) / (18 * ((pTin2 - pTout2) / (pTout - pTin)))) ^ 0.6
        End If
    ElseIf pCompType = "Pump" Then
            PEC = 0.00193 * (pFmin) ^ 2 + 15.697 * pFmin + 3569.8
    ElseIf pCompType = "Steam Turbine" Then
            PEC = -0.00185 * (pPower) ^ 2 + 135.251 * pPower + 11974
    ElseIf pCompType = "Solar Heater" Then
        If Sheets("GT Specs").Range("R14") = "Yes" Then
            PEC = pPower * 1000 * Sheets("GT Specs").Range("R15")
        Else
            PEC = 0
        End If
    End If
End Property

Property Get Ck() As Double

Ck = (PEC - 0.1 * (1 + I) ^ (-Ny)) * (I * (1 - (1 + I) ^ (-Ny)))

End Property
Property Get Zk() As Double

Zk = pPhiM * Ck '/iOH


End Property


Property Get index() As Integer
     index = iindex
End Property

Property Let index(oindex As Integer)
     iindex = oindex
End Property

Property Get I() As Integer
     I = iI
End Property

Property Let I(oI As Integer)
     iI = oI
End Property

Property Get Ny() As Integer
     Ny = iNy
End Property

Property Let Ny(oNy As Integer)
     iNy = oNy
End Property

Property Get OH() As Integer
     OH = iOH
End Property

Property Let OH(oOH As Integer)
     iOH = oOH
End Property

Property Get PhiM() As Double
     PhiM = pPhiM
End Property

Property Let PhiM(dPhiM As Double)
     pPhiM = dPhiM
End Property

Property Get HHV() As Double
     HHV = pHHV
End Property

Property Let HHV(dHHV As Double)
     pHHV = dHHV
End Property


Property Get h() As Double
     h = ph
End Property

Property Let h(dh As Double)
     ph = dh
End Property

Property Get Cp() As Double
     Cp = pCp
End Property

Property Let Cp(dCp As Double)
     pCp = dCp
End Property

Property Get LastComponent() As Boolean
     LastComponent = pLastComponent
End Property

Property Let LastComponent(dLastComponent As Boolean)
     pLastComponent = dLastComponent
End Property

Property Get TipSpeed() As Double
     TipSpeed = pTipSpeed
End Property

Property Let TipSpeed(dTipSpeed As Double)
     pTipSpeed = dTipSpeed
End Property

Property Get RotatingSpeed() As Double
     RotatingSpeed = pRotatingSpeed
End Property

Property Let RotatingSpeed(dRotatingSpeed As Double)
     pRotatingSpeed = dRotatingSpeed
End Property

Property Get NumberStage() As Double
     NumberStage = pNumberStage
End Property

Property Let NumberStage(dNumberStage As Double)
     pNumberStage = dNumberStage
End Property

Property Get Diameter() As Double
     Diameter = pDiameter
End Property

Property Let Diameter(dDiameter As Double)
     pDiameter = dDiameter
End Property

Property Get Pin() As Double
     Pin = pPin
End Property

Property Let Pin(dPin As Double)
     pPin = dPin
End Property

Property Get Pin2() As Double
     Pin2 = pPin2
End Property

Property Let Pin2(dPin2 As Double)
     pPin2 = dPin2
End Property

Property Get Pout() As Double
     Pout = pPout
End Property

Property Let Pout(dPout As Double)
     pPout = dPout
End Property

Property Get Pout2() As Double
     Pout2 = pPout2
End Property

Property Let Pout2(dPout2 As Double)
     pPout2 = dPout2
End Property

Property Get Tin() As Double
     Tin = pTin
End Property

Property Let Tin(dTin As Double)
     pTin = dTin
End Property

Property Get Tin2() As Double
     Tin2 = pTin2
End Property

Property Let Tin2(dTin2 As Double)
     pTin2 = dTin2
End Property

Property Get Tout() As Double
     Tout = pTout
End Property

Property Let Tout(dTout As Double)
     pTout = dTout
End Property

Property Get Tout2() As Double
     Tout2 = pTout2
End Property

Property Let Tout2(dTout2 As Double)
     pTout2 = dTout2
End Property
Property Get hIn() As Double
     hIn = phin
End Property

Property Let hIn(dhIn As Double)
     phin = dhIn
End Property

Property Get hIn2() As Double
     hIn2 = phin2
End Property

Property Let hIn2(dhIn2 As Double)
     phin2 = dhIn2
End Property

Property Get hout() As Double
     hout = phout
End Property

Property Let hout(dhout As Double)
     phout = dhout
End Property

Property Get hout2() As Double
     hout2 = phout2
End Property

Property Let hout2(dhout2 As Double)
     phout2 = dhout2
End Property

Property Let Fin(dFin As Double)
     pFin = dFin
End Property

Property Get Fin() As Double
     Fin = pFin
End Property

Property Let Fmin(dFmin As Double)
     pFmin = dFmin
End Property

Property Get Fmin() As Double
     Fmin = pFmin
End Property

Property Let Fin2(dFin2 As Double)
     pFin2 = dFin2
End Property

Property Get Fin2() As Double
     Fin2 = pFin2
End Property

Property Let Fout(dFout As Double)
     pFout = dFout
End Property

Property Get Fout() As Double
     Fout = pFout
End Property

Property Let Fout2(dFout2 As Double)
     pFout2 = dFout2
End Property

Property Get Fout2() As Double
     Fout2 = pFout2
End Property

Property Let power(dPower As Double)
     pPower = dPower
End Property

Property Get power() As Double
     power = pPower
End Property

Property Let Efficiency(dEfficiency As Double)
     pEfficiency = dEfficiency
End Property

Property Get Efficiency() As Double
     Efficiency = pEfficiency
End Property

Property Let DeltaT(dDeltaT As Double)
     pDeltaT = dDeltaT
End Property

Property Get DeltaT() As Double
     DeltaT = pDeltaT
End Property

Property Let DeltaP1(dDeltaP1 As Double)
     pDeltaP1 = dDeltaP1
End Property

Property Get DeltaP1() As Double
     DeltaP1 = pDeltaP1
End Property

Property Let DeltaP2(dDeltaP2 As Double)
     pDeltaP2 = dDeltaP2
End Property

Property Get DeltaP2() As Double
     DeltaP2 = pDeltaP2
End Property

Property Let PressureRatio(dPressureRatio As Double)
     pPressureRatio = dPressureRatio
End Property

Property Get PressureRatio() As Double
     PressureRatio = pPressureRatio
End Property

Property Let Reaction(dReaction As String)
     pReaction = dReaction
End Property

Property Get Reaction() As String
     Reaction = pReaction
End Property

Property Let CollectorSize(dCollectorSize As Double)
     pCollectorSize = dCollectorSize
End Property

Property Get CollectorSize() As Double
     CollectorSize = pCollectorSize
End Property

Property Let SolarRadiation(dSolarRadiation As Double)
     pSolarRadiation = dSolarRadiation
End Property

Property Get SolarRadiation() As Double
     SolarRadiation = pSolarRadiation
End Property

Property Let ExtraPercentage(dExtraPercentage As Double)
     pExtraPercentage = dExtraPercentage
End Property

Property Get ExtraPercentage() As Double
     ExtraPercentage = pExtraPercentage
End Property

Property Let ExtraPercentage2(dExtraPercentage2 As Double)
     pExtraPercentage2 = dExtraPercentage2
End Property

Property Get ExtraPercentage2() As Double
     ExtraPercentage2 = pExtraPercentage2
End Property

Property Let cycleName(dcycleName As String)
     pCycleName = dcycleName
End Property

Property Get cycleName() As String
     cycleName = pCycleName
End Property

Property Let cycleSheet(dcycleSheet As String)
     pCycleSheet = dcycleSheet
End Property

Property Get cycleSheet() As String
     cycleSheet = pCycleSheet
End Property

Property Let CompName(iCompName As String)
    pCompName = iCompName
End Property

Property Get CompName() As String
    CompName = pCompName
End Property

Property Let CompType(iCompType As String)
    pCompType = iCompType
End Property

Property Get CompType() As String
    CompType = pCompType
End Property


Property Let HXType(iHXType As String)
    pHXType = iHXType
End Property

Property Get HXType() As String
    HXType = pHXType
End Property

