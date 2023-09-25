VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pname As String
Private pCType As String
Private pStreamPilot As String
Private iindex As Integer
Private pEfficiency As Double
Private pHeatRate As Double
Private pFiringTemp As Double
Private pPower As Double
Private pHeatPower As Double
Private pPressureRatio As Double
Private pNumberCompressor As Integer
Private pNumberTurbine As Integer
Private pNumberPump As Integer
Private pFeedFlow As Double
Private pFuelFlow As Double
Private pFuelCost As Double
Private pCost As Double
Private pCostKWH As Double

Property Get NumberCompressor() As Integer
     NumberCompressor = pNumberCompressor
End Property

Property Let NumberCompressor(oNumberCompressor As Integer)
     pNumberCompressor = oNumberCompressor
End Property

Property Get NumberPump() As Integer
     NumberPump = pNumberPump
End Property

Property Let NumberPump(oNumberPump As Integer)
     pNumberPump = oNumberPump
End Property

Property Get NumberTurbine() As Integer
     NumberTurbine = pNumberTurbine
End Property

Property Let NumberTurbine(oNumberTurbine As Integer)
     pNumberTurbine = oNumberTurbine
End Property

Property Get index() As Integer
     index = iindex
End Property

Property Let index(oindex As Integer)
     iindex = oindex
End Property

Property Get name() As String
     name = pname
End Property
Property Let name(oname As String)
     pname = oname
End Property

Property Get CType() As String
     CType = pCType
End Property

Property Let CType(oCType As String)
     pCType = oCType
End Property

Property Get StreamPilot() As String
     StreamPilot = pStreamPilot
End Property

Property Let StreamPilot(oStreamPilot As String)
     pStreamPilot = oStreamPilot
End Property

Property Get Efficiency() As Double
     Efficiency = pEfficiency
End Property

Property Let Efficiency(oEfficiency As Double)
     pEfficiency = oEfficiency
End Property

Property Get HeatRate() As Double
     HeatRate = pHeatRate
End Property

Property Let HeatRate(oHeatRate As Double)
     pHeatRate = oHeatRate
End Property

Property Get power() As Double
     power = pPower
End Property

Property Let power(oPower As Double)
     pPower = oPower
End Property

Property Get HeatPower() As Double
     HeatPower = pHeatPower
End Property

Property Let HeatPower(oHeatPower As Double)
     pHeatPower = oHeatPower
End Property

Property Get FuelCost() As Double
     FuelCost = pFuelCost
End Property

Property Let FuelCost(oFuelCost As Double)
     pFuelCost = oFuelCost
End Property

Property Get FeedFlow() As Double
     FeedFlow = pFeedFlow
End Property

Property Let FeedFlow(oFeedFlow As Double)
     pFeedFlow = oFeedFlow
End Property


Property Get FiringTemp() As Double
     FiringTemp = pFiringTemp
End Property

Property Let FiringTemp(oFiringTemp As Double)
     pFiringTemp = oFiringTemp
End Property

Property Get Cost() As Double
     Cost = pCost
End Property

Property Let Cost(oCost As Double)
     pCost = oCost
End Property


Property Get costKWH() As Double
     costKWH = pCostKWH
End Property

Property Let costKWH(ocostKWH As Double)
     pCostKWH = ocostKWH
End Property

Property Get FuelFlow() As Double
     FuelFlow = pFuelFlow
End Property

Property Let FuelFlow(oFuelFlow As Double)
     pFuelFlow = oFuelFlow
End Property

Property Get PressureRatio() As Double
     PressureRatio = pPressureRatio
End Property

Property Let PressureRatio(oPressureRatio As Double)
     pPressureRatio = oPressureRatio
End Property
