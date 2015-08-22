VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DsgPerimeters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PONEfcu, PTWOfy, PTHREEdia, PFOURcvr, PFOURbmark


Public Function DSG_fcu()
   DSG_fcu = PONEfcu
End Function
Public Property Let Value9(Setting)
   PONEfcu = Setting
End Property

Public Function DSG_fy()
   DSG_fy = PTWOfy
End Function
Public Property Let Value10(Setting)
   PTWOfy = Setting
End Property

Public Function DSG_bardia()
   DSG_bardia = PTHREEdia
End Function
Public Property Let Value11(Setting)
   PTHREEdia = Setting
End Property

Public Function DSG_cover()
   DSG_cover = PFOURcvr
End Function
Public Property Let Value12(Setting)
   PFOURcvr = Setting
End Property

Public Function DSG_barmark()
   DSG_barmark = PFOURbmark
End Function
Public Property Let Value13(Setting)
   PFOURbmark = Setting
End Property
