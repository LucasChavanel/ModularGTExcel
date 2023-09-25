VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTurbine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iindex As Integer
Private iAssTurb As Integer

'Inputs
Private iNumberStage As Integer
Private pRotatingSpeed As Double
Private pReactionDegree As Double
Private pHTRatio As Double
Private pFlowCoeff As Double
Private pLoadingCoeff As Double
Private pPressureRatio As Double
Private pEfficiency As Double

Private pPin As Double
Private pPout As Double
Private pTin As Double
Private pTout As Double
Private pFin As Double
Private pFout As Double
Private pPower As Double
Private pCycleName As String
Private pCompName As String

'Aspen Data
Private pMassFlow As Double
Private pPressOut As Double
Private pPressIn As Double
Private pTempIn As Double

Private pGamma As Double
Private pCp As Double
Private pMaxTipSpeed As Double


Property Get Pin() As Double
     Pin = pPin
End Property

Property Let Pin(dPin As Double)
     pPin = dPin
End Property

Property Get Pout() As Double
     Pout = pPout
End Property

Property Let Pout(dPout As Double)
     pPout = dPout
End Property

Property Get Tin() As Double
     Tin = pTin
End Property

Property Let Tin(dTin As Double)
     pTin = dTin
End Property

Property Get Tout() As Double
     Tout = pTout
End Property

Property Let Tout(dTout As Double)
     pTout = dTout
End Property

Property Let Fin(dFin As Double)
     pFin = dFin
End Property

Property Get Fin() As Double
     Fin = pFin
End Property

Property Let Fout(dFout As Double)
     pFout = dFout
End Property

Property Get Fout() As Double
     Fout = pFout
End Property

Property Let power(dPower As Double)
     pPower = dPower
End Property

Property Get power() As Double
     power = pPower
End Property

Property Let cycleName(dcycleName As String)
     pCycleName = dcycleName
End Property

Property Get cycleName() As String
     cycleName = pCycleName
End Property

Property Let CompName(iCompName As String)
    pCompName = iCompName
End Property

Property Get CompName() As String
    CompName = pCompName
End Property
'//////////////////////////////Formulas/////////////////////////////////
Property Get PEC() As Double

     PEC = (479.34 * pMassFlow) / (0.95 - pEfficiency / 100) * WorksheetFunction.Ln(pPressIn / pPressOut) * (1 + Exp(0.036 * (pTempIn) - 54.4))
     
End Property



Property Get Um() As Double

    Um = (pPower * 1000 / iNumberStage / pMassFlow / pLoadingCoeff) ^ (1 / 2)
End Property


Property Get Ut() As Double

    Ut = 2 / (1 + pHTRatio) * Um
End Property


Property Get Alpha2() As Double

    Alpha2 = Atn((pLoadingCoeff / 2 + pReactionDegree - 1) / pFlowCoeff)
    
End Property

Property Get V2() As Double

    V2 = pFlowCoeff * Um / Cos(Alpha2)
    
End Property

Property Get VR() As Double

    VR = Ut / V2
    
End Property

Property Get Dm() As Double

    Dm = 120 * Um / 3.14159205 / pRotatingSpeed
    
End Property

Property Get PolytropicEff() As Double

    PolytropicEff = pGamma / (1 - pGamma) * WorksheetFunction.Ln(1 - pEfficiency / 100 * (1 - pPressureRatio ^ ((1 - pGamma) / pGamma))) / WorksheetFunction.Ln(pPressureRatio)
    
End Property


Property Get DTstage() As Double
    
    DTstage = (1 / iNumberStage) * pTempIn * (1 - (pPressureRatio) ^ ((1 - pGamma) / pGamma * PolytropicEff))
    
End Property



'//////////////////////////////End of Formulas////////////////////////////

'Variable Definition
Property Get gamma() As Double
     gamma = pGamma
End Property

Property Let gamma(oGamma As Double)
     pGamma = oGamma
End Property

Property Get Cp() As Double
     Cp = pCp
End Property

Property Let Cp(oCp As Double)
     pCp = oCp
End Property

Property Get MaxTipSpeed() As Double
     MaxTipSpeed = pMaxTipSpeed
End Property

Property Let MaxTipSpeed(oMaxTipSpeed As Double)
     pMaxTipSpeed = oMaxTipSpeed
End Property


Property Get PressureRatio() As Double
     PressureRatio = pPressureRatio
End Property

Property Let PressureRatio(oPressureRatio As Double)
     pPressureRatio = oPressureRatio
End Property

Property Get LoadingCoeff() As Double
     LoadingCoeff = pLoadingCoeff
End Property

Property Let LoadingCoeff(oLoadingCoeff As Double)
     pLoadingCoeff = oLoadingCoeff
End Property

Property Get FlowCoeff() As Double
     FlowCoeff = pFlowCoeff
End Property

Property Let FlowCoeff(oFlowCoeff As Double)
     pFlowCoeff = oFlowCoeff
End Property

Property Get HTRatio() As Double
     HTRatio = pHTRatio
End Property

Property Let HTRatio(oHTRatio As Double)
     pHTRatio = oHTRatio
End Property

Property Get ReactionDegree() As Double
     ReactionDegree = pReactionDegree
End Property

Property Let ReactionDegree(oReactionDegree As Double)
     pReactionDegree = oReactionDegree
End Property

Property Get RotatingSpeed() As Double
     RotatingSpeed = pRotatingSpeed
End Property

Property Let RotatingSpeed(oRotatingSpeed As Double)
     pRotatingSpeed = oRotatingSpeed
End Property
Property Get NumberStage() As Integer
     NumberStage = iNumberStage
End Property

Property Let NumberStage(oNumberStage As Integer)
     iNumberStage = oNumberStage
End Property

Property Get Efficiency() As Double
     Efficiency = pEfficiency
End Property

Property Let Efficiency(oEfficiency As Double)
     pEfficiency = oEfficiency
End Property

Property Get AssTurb() As Integer
     AssTurb = iAssTurb
End Property

Property Let AssTurb(oAssTurb As Integer)
     iAssTurb = oAssTurb
End Property

Property Get MassFlow() As Double
     MassFlow = pMassFlow
End Property

Property Let MassFlow(dMassFlow As Double)
     pMassFlow = dMassFlow
End Property

Property Get PressOut() As Double
     PressOut = pPressOut
End Property

Property Let PressOut(dPressOut As Double)
     pPressOut = dPressOut
End Property

Property Get PressIn() As Double
     PressIn = pPressIn
End Property

Property Let PressIn(dPressIn As Double)
     pPressIn = dPressIn
End Property

Property Get TempIn() As Double
     TempIn = pTempIn
End Property

Property Let TempIn(dTempIn As Double)
     pTempIn = dTempIn
End Property


Property Get index() As Integer
     index = iindex
End Property

Property Let index(oindex As Integer)
     iindex = oindex
End Property
