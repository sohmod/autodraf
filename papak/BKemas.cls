VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BKemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PONEbK, PTWObK, PTHREEbK, PFOURbK, PFIVEbK

Private PSIXbK, PSEVENbK, PEIGHTbK, PNINEbK, PTENbK

Private PELEVENbK

Public Function BKemas_ONE_D()
   BKemas_ONE_D = PONEbK
End Function
Public Property Let Value9(Setting)
   PONEbK = Setting
End Property

Public Function BKemas_TWO_D()
   BKemas_TWO_D = PTWObK
End Function
Public Property Let Value10(Setting)
   PTWObK = Setting
End Property

Public Function BKemas_THREE_D()
   BKemas_THREE_D = PTHREEbK
End Function
Public Property Let Value11(Setting)
   PTHREEbK = Setting
End Property

Public Function BKemas_FOUR_D()
   BKemas_FOUR_D = PFOURbK
End Function
Public Property Let Value12(Setting)
   PFOURbK = Setting
End Property

Public Function BKemas_ONE_C()
   BKemas_ONE_C = PFIVEbK
End Function
Public Property Let Value13(Setting)
   PFIVEbK = Setting
End Property

Public Function BKemas_TWO_C()
   BKemas_TWO_C = PSIXbK
End Function
Public Property Let Value14(Setting)
  PSIXbK = Setting
End Property

Public Function BKemas_THREE_C()
   BKemas_THREE_C = PSEVENbK
End Function
Public Property Let Value15(Setting)
  PSEVENbK = Setting
End Property

Public Function BKemas_ONE_B()
   BKemas_ONE_B = PEIGHTbK
End Function
Public Property Let Value16(Setting)
   PEIGHTbK = Setting
End Property

Public Function BKemas_TWO_B()
   BKemas_TWO_B = PNINEbK
End Function
Public Property Let Value17(Setting)
   PNINEbK = Setting
End Property

Public Function BKemas_ONE_A()
   BKemas_ONE_A = PTENbK
End Function
Public Property Let Value18(Setting)
   PTENbK = Setting
End Property

Public Function BKemas_FOUR_A()
  BKemas_FOUR_A = PELEVENbK
End Function
Public Property Let Value19(Setting)
   PELEVENbK = Setting
End Property

