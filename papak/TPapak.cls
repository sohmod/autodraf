VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPapak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PONEbH, PTWObH, PTHREEbH, PFOURbH, PFIVEbH

Private PSIXbH, PSEVENbH, PEIGHTbH, PNINEbH, PTENbH

Private PELEVENbH

Public Function TPapak_ONE_D()
   TPapak_ONE_D = PONEbH
End Function
Public Property Let Value9(Setting)
   PONEbH = Setting
End Property

Public Function TPapak_TWO_D()
   TPapak_TWO_D = PTWObH
End Function
Public Property Let Value10(Setting)
   PTWObH = Setting
End Property

Public Function TPapak_THREE_D()
   TPapak_THREE_D = PTHREEbH
End Function
Public Property Let Value11(Setting)
   PTHREEbH = Setting
End Property

Public Function TPapak_FOUR_D()
   TPapak_FOUR_D = PFOURbH
End Function
Public Property Let Value12(Setting)
   PFOURbH = Setting
End Property

Public Function TPapak_ONE_C()
   TPapak_ONE_C = PFIVEbH
End Function
Public Property Let Value13(Setting)
   PFIVEbH = Setting
End Property

Public Function TPapak_TWO_C()
   TPapak_TWO_C = PSIXbH
End Function
Public Property Let Value14(Setting)
  PSIXbH = Setting
End Property

Public Function TPapak_THREE_C()
   TPapak_THREE_C = PSEVENbH
End Function
Public Property Let Value15(Setting)
  PSEVENbH = Setting
End Property

Public Function TPapak_ONE_B()
   TPapak_ONE_B = PEIGHTbH
End Function
Public Property Let Value16(Setting)
   PEIGHTbH = Setting
End Property

Public Function TPapak_TWO_B()
   TPapak_TWO_B = PNINEbH
End Function
Public Property Let Value17(Setting)
   PNINEbH = Setting
End Property

Public Function TPapak_ONE_A()
   TPapak_ONE_A = PTENbH
End Function
Public Property Let Value18(Setting)
   PTENbH = Setting
End Property

Public Function TPapak_FOUR_A()
  TPapak_FOUR_A = PELEVENbH
End Function
Public Property Let Value19(Setting)
   PELEVENbH = Setting
End Property

