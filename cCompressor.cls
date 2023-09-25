VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCompressor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPressureRatio As Double
Private pTemperatureRatio As Double
Private pIsenEfficiency As Double
Private pPhiAve As Double
Private pPhi As Double

Private pZ As Double 'Gas Compressibility
Private pGasC As Double 'Gas Constant : R/Mmol




Private pMaxTipSpeed As Double
Private pRSpeed As Double

Private pCompName As String

Private pTempo As Double '°R
Private pTempd As Double
Private pMassFlow As Double 'cfm

Private pPress As Double 'psia

Private cStages As Collection

Private pPressureRatioFinal As Double
Private pTemperatureRatioFinal As Double
Private pIsenEfficiencyFinal As Double
Private iindex As Integer

Private iAssComp As Integer
Private iNumberStages As Integer



Property Let CompName(iCompName As String)
    pCompName = iCompName
End Property

Property Get CompName() As String
    CompName = pCompName
End Property

'------------------------------Formulas---------------------------------
Property Get PsiAve() As Double

If 0.036 < pPhiAve Then
    PsiAve = -977.75 * (pPhiAve) ^ 3 + 195.56 * (pPhiAve) ^ 2 - 13.166 * (pPhiAve) + 0.8492
ElseIf 0.025 < pPhiAve Then
    PsiAve = (-9057.1 * (pPhiAve) ^ 3 + 234.68 * (pPhiAve) ^ 2 - 0.2528 * (pPhiAve) + 0.5679)
ElseIf 0.0165 < pPhiAve Then
    PsiAve = -799.79 * (pPhiAve) ^ 2 + 19.929 * (pPhiAve) + 0.4579
ElseIf 0 < pPhiAve Then
    PsiAve = -1698.3 * (pPhiAve) ^ 2 + 32.444 * (pPhiAve) + 0.3973
Else
    'MsgBox "La valeur de phi pour le stage n'est pas compris entre 0 et 0.036"
End If

End Property

Property Get Flow() As Double

Flow = pMassFlow * 60 * pZ * pGasC * pTempo / 144 / pPress / 0.145 'conversion in cubic feet/min

End Property

Property Get Hp() As Double
Dim nsurn As Double
nsurn = 0.286 / pIsenEfficiency

Hp = pZ * pGasC * pTempo / nsurn * (((pPressureRatio) ^ (nsurn)) - 1)

End Property

Property Get HeadPerStage() As Double

HeadPerStage = (PsiAve) * ((pMaxTipSpeed) ^ 2) / 32.2

End Property


Property Get TipS() As Double
     TipS = (Hp / iNumberStages * 9.81 * 3.28084 / PsiAve) ^ (1 / 2)
End Property


Property Get Diameter() As Double
    Diameter = 229 * TipS / pRSpeed
End Property





Property Get PEC() As Double

     PEC = 71.1 * pMassFlow / (0.9 - (pIsenEfficiency / 100)) * (pPressureRatio) * WorksheetFunction.Ln(pPressureRatio)
     
End Property

'-----------------------------End of formulas---------------------------

Property Get AssComp() As Integer
     AssComp = iAssComp
End Property

Property Let AssComp(oAssComp As Integer)
     iAssComp = oAssComp
End Property

Property Get NumberStages() As Integer
     NumberStages = iNumberStages
End Property

Property Let NumberStages(oNumberStages As Integer)
     iNumberStages = oNumberStages
End Property

Property Get index() As Integer
     index = iindex
End Property

Property Let index(oindex As Integer)
     iindex = oindex
End Property

Property Get MassFlow() As Double
     MassFlow = pMassFlow
End Property

Property Let MassFlow(dMassFlow As Double)
     pMassFlow = dMassFlow
End Property

Property Get MaxTipSpeed() As Double
     MaxTipSpeed = pMaxTipSpeed
End Property

Property Let MaxTipSpeed(dMaxTipSpeed As Double)
     pMaxTipSpeed = dMaxTipSpeed
End Property

Property Get GasC() As Double
     GasC = pGasC
End Property

Property Let GasC(dGasC As Double)
     pGasC = dGasC
End Property

Property Get RSpeed() As Double

 RSpeed = pRSpeed

End Property

Property Let RSpeed(dRSpeed As Double)
     pRSpeed = dRSpeed
End Property


Property Get Stages() As Collection
    Set Stages = cStagecomp
End Property

Property Set Stages(stagesi As Collection)
    Set cStages = stagesi
End Property

Property Get Tempd() As Double
    Tempd = pTempd
End Property


Property Let Tempd(iTempd As Double)
    pTempd = iTempd
End Property

Property Let Press(iPress As Double)
    pPress = iPress
End Property

Property Get Press() As Double
    Press = pPress
End Property

Property Get Tempo() As Double
    Tempo = pTempo
End Property

Property Let Tempo(iTempo As Double)
    pTempo = iTempo
End Property

Property Get PressureRatio() As Double
    PressureRatio = pPressureRatio
End Property

Property Let PressureRatio(iPressureRatio As Double)
    pPressureRatio = iPressureRatio
End Property

Property Get TemperatureRatio() As Double
    TemperatureRatio = pTemperatureRatio
End Property

Property Let TemperatureRatio(iTemperatureRatio As Double)
    pTemperatureRatio = iTemperatureRatio
End Property


Property Get IsenEfficiency() As Double
    IsenEfficiency = pIsenEfficiency
End Property

Property Let IsenEfficiency(iIsenEfficiency As Double)
    pIsenEfficiency = iIsenEfficiency
End Property

Property Let Z(iZ As Double)
    pZ = iZ
End Property

Property Get Z() As Double
    Z = pZ
End Property

Property Get PhiAve() As Double
    PhiAve = pPhiAve
End Property

Property Let PhiAve(iPhiAve As Double)
    pPhiAve = iPhiAve
End Property





Property Get PressureRatioFinal() As Double
    PressureRatioFinal = pPressureRatioFinal
End Property

Property Let PressureRatioFinal(iPressureRatioFinal As Double)
    pPressureRatioFinal = iPressureRatioFinal
End Property

Property Get TemperatureRatioFinal() As Double
    TemperatureRatioFinal = pTemperatureRatioFinal
End Property

Property Let TemperatureRatioFinal(iTemperatureRatio As Double)
    pTemperatureRatioFinal = iTemperatureRatioFinal
End Property

Property Get IsenEfficiencyFinal() As Double
    IsenEfficiencyFinal = pIsenEfficiencyFinal
End Property

Property Let IsenEfficiencyFinal(iIsenEfficiencyFinal As Double)
    pIsenEfficiencyFinal = iIsenEfficiencyFinal
End Property


