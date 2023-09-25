VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pFlowSt As Double
Private pZo As Double
Private pTh As Double
Private pPh As Double
Private pNumberStage As Integer
Private pCompName As String
Private iindex As Integer
Private pRSpeed As Double
Private pDiameter As Double
Private pPhiAve As Double



Property Get index() As Integer
     index = iindex
End Property

Property Let index(oindex As Integer)
     iindex = oindex
End Property

Property Let RSpeed(dRSpeed As Double)
 pRSpeed = dRSpeed
End Property

Property Get RSpeed() As Double
 RSpeed = pRSpeed
End Property



Property Let Diameter(dDiameter As Double)
 pDiameter = dDiameter
End Property

Property Get Diameter() As Double
 Diameter = pDiameter
End Property
Property Let CompName(dFlowSt As String)
 pCompName = sCompName
End Property

Property Get CompName() As String
 CompName = pCompName
End Property

Property Let FlowSt(dFlowSt As Double)
 pFlowSt = dFlowSt
End Property

Property Get FlowSt() As Double
 FlowSt = pFlowSt
End Property

Property Let Zo(dZo As Double)
 pZo = dZo
End Property

Property Get Zo() As Double
 Zo = pZo
End Property

Property Let Th(dTh As Double)
 pTh = dTh
End Property

Property Get Th() As Double
 Th = pTh
End Property

Property Let ph(dPh As Double)
 pPh = dPh
End Property

Property Get ph() As Double
 ph = pPh
End Property

Property Let NumberStage(iNumberStage As Integer)
 pNumberStage = iNumberStage
End Property

Property Get NumberStage() As Integer
 NumberStage = pNumberStage
End Property

Property Get PhiAve() As Double

PhiAve = pFlowSt * 700 / pRSpeed / ((pDiameter) ^ 3)
 
End Property

Property Get ImpellerType() As String

If 0.037 < PhiAve Then
    ImpellerType = "Typical"

ElseIf 0.028 < PhiAve Then
    ImpellerType = "Erad"
ElseIf 0.022 < PhiAve Then
    ImpellerType = "Ep"
ElseIf 0.02 < PhiAve Then
    ImpellerType = "Fct"
ElseIf 0.018 < PhiAve Then
    ImpellerType = "Frad"
ElseIf 0.0165 < PhiAve Then
    ImpellerType = "Fp"
ElseIf 0.014 < PhiAve Then
    ImpellerType = "Gct"
ElseIf 0.012 < PhiAve Then
    ImpellerType = "Grad"
ElseIf 0 < PhiAve Then
    ImpellerType = "Gp"
Else
    'MsgBox "La valeur de phi pour le stage n'est pas compris entre 0 et 0.036"
End If

End Property


Property Get Etah() As Double

If ImpellerType = "Typical" Then
    Etah = -1631.2 * (PhiAve) ^ 3 + 335.23 * (PhiAve) ^ 2 - 20.469 * (PhiAve) + 1.1862
    
ElseIf ImpellerType = "Erad" Then
    Etah = -18763 * (PhiAve) ^ 3 + 759.14 * (PhiAve) ^ 2 - 1.2546 * (PhiAve) + 0.717
    
ElseIf ImpellerType = "Ep" Then
    Etah = -2335241.55 * (PhiAve) ^ 4 + 198568 * (PhiAve) ^ 3 - 6833.9 * (PhiAve) ^ 2 + 115.03 * (PhiAve) + 0.0717
    
ElseIf ImpellerType = "Fct" Then
    Etah = -3984766.09 * (PhiAve) ^ 4 + 295808 * (PhiAve) ^ 3 - 8990.5 * (PhiAve) ^ 2 + 133.22 * (PhiAve) + 0.0499
    
ElseIf ImpellerType = "Frad" Then
    Etah = -3784920.62 * (PhiAve) ^ 4 + 237995 * (PhiAve) ^ 3 - 6347.2 * (PhiAve) ^ 2 + 87.785 * (PhiAve) + 0.3445

ElseIf ImpellerType = "Fp" Then
    Etah = -4419820.8 * (PhiAve) ^ 4 + 263946 * (PhiAve) ^ 3 - 6881.2 * (PhiAve) ^ 2 + 94.47 * (PhiAve) + 0.3052

ElseIf ImpellerType = "Gct" Then
    Etah = -39236574.18 * (PhiAve) ^ 4 + 2528667.17 * (PhiAve) ^ 3 - 62850 * (PhiAve) ^ 2 + 709.69 * (PhiAve) - 2.2347
    
ElseIf ImpellerType = "Grad" Then
    Etah = -17165080.11 * (PhiAve) ^ 4 + 831391 * (PhiAve) ^ 3 - 16638 * (PhiAve) ^ 2 + 166.13 * (PhiAve) + 0.1465
    
ElseIf ImpellerType = "Gp" Then
    Etah = -55927417.21 * (PhiAve) ^ 4 + 3079406.48 * (PhiAve) ^ 3 - 65600.62 * (PhiAve) ^ 2 + 637.42 * (PhiAve) - 1.55
    
End If

End Property

Property Get nsurn() As Double

nsurn = 0.286 / Etah

End Property

Property Get PsiEtah() As Double

If ImpellerType = "Typical" Then
    PsiEtah = -2.3421 * PhiAve + 0.8233

ElseIf ImpellerType = "Erad" Then
    PsiEtah = -7.8078 * PhiAve + 0.8466
    
ElseIf ImpellerType = "Ep" Then
    PsiEtah = -6.0839 * PhiAve + 0.8101
    
ElseIf ImpellerType = "Fct" Then
    PsiEtah = -15.32 * PhiAve + 0.9363
    
ElseIf ImpellerType = "Frad" Then
    PsiEtah = -12.096 * PhiAve + 0.8654
    
ElseIf ImpellerType = "Fp" Then
    PsiEtah = -8.8361 * PhiAve + 0.8143
    
ElseIf ImpellerType = "Gct" Then
    PsiEtah = -25.463 * PhiAve + 0.9944
    
ElseIf ImpellerType = "Grad" Then
    PsiEtah = -19.167 * PhiAve + 0.8896
    
ElseIf ImpellerType = "Gp" Then
    PsiEtah = -15.696 * PhiAve + 0.8533

End If


End Property

Property Get DeltaT() As Double

DeltaT = 63.02 * PsiEtah

End Property

Property Get TsurTst() As Double

TsurTst = 1 + DeltaT / Th

End Property
Property Get PsurPst() As Double

PsurPst = (TsurTst) ^ (1 / nsurn)

End Property

Property Get DeltaP() As Double

DeltaP = ph * (PsurPst - 1)

End Property


