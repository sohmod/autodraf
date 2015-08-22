VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'''''''''''
Private ONEtoTWO, TWOtoTHREE, THREEtoFOUR, FOURtoFIVE
Private EtoD, DtoC, CtoB, BtoA

'''''''''''

Public Function Span_ONEtoTWO()
   Span_ONEtoTWO = ONEtoTWO
End Function
Public Property Let Value1(Setting)
   ONEtoTWO = Setting
End Property

Public Function Span_TWOtoTHREE()
   Span_TWOtoTHREE = TWOtoTHREE
End Function
Public Property Let Value2(Setting)
   TWOtoTHREE = Setting
End Property

Public Function Span_THREEtoFOUR()
  Span_THREEtoFOUR = THREEtoFOUR
End Function
Public Property Let Value3(Setting)
   THREEtoFOUR = Setting
End Property

Public Function Span_FOURtoFIVE()
  Span_FOURtoFIVE = FOURtoFIVE
End Function
Public Property Let Value4(Setting)
   FOURtoFIVE = Setting
End Property
''''''''''''''''''''''''''''''''''''
Public Function Span_EtoD()
   Span_EtoD = EtoD
End Function
Public Property Let Value5(Setting)
   EtoD = Setting
End Property

Public Function Span_DtoC()
   Span_DtoC = DtoC
End Function
Public Property Let Value6(Setting)
  DtoC = Setting
End Property

Public Function Span_CtoB()
  Span_CtoB = CtoB
End Function
Public Property Let Value7(Setting)
  CtoB = Setting
End Property

Public Function Span_BtoA()
   Span_BtoA = BtoA
End Function
Public Property Let Value8(Setting)
   BtoA = Setting
End Property

