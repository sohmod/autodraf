VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppliedStress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Xbase, Xbrace
Private Mxx2, Myy2
Private AxiaL
Private Mxx1, Myy1
Private Ybase, Ybrace


Public Function XbaseCondition()
   XbaseCondition = Xbase
End Function
Public Property Let Value4(Setting)
   Xbase = Setting
End Property


Public Function MomentX22()
If Val(Mxx2) = 0 Then
    Mxx2 = 0
       End If
   MomentX22 = Mxx2
End Function
Public Property Let Value20(Setting)
   Mxx2 = Setting
End Property


Public Function MomentY22()
If Val(Myy2) = 0 Then
    Myy2 = 0
       End If
   MomentY22 = Myy2
End Function
Public Property Let Value21(Setting)
   Myy2 = Setting
End Property


Public Function AxiaLoaD()
If Val(AxiaL) <= 0 Then
    AxiaL = 0
       End If
   AxiaLoaD = AxiaL
End Function
Public Property Let Value40(Setting)
   AxiaL = Setting
End Property


Public Function MomentX11()
If Val(Mxx1) = 0 Then
    Mxx1 = 0
       End If
   MomentX11 = Mxx1
End Function
Public Property Let Value22(Setting)
   Mxx1 = Setting
End Property


Public Function MomentY11()
If Val(Myy1) = 0 Then
    Myy1 = 0
       End If
   MomentY11 = Myy1
End Function
Public Property Let Value23(Setting)
   Myy1 = Setting
End Property


Public Function YbaseCondition()
   YbaseCondition = Ybase
End Function
Public Property Let Value39(Setting)
   Ybase = Setting
End Property


Public Function XBracedFrame()
   XBracedFrame = Xbrace
End Function
Public Property Let Value61(Setting)
   Xbrace = Setting
End Property


Public Function YBracedFrame()
   YBracedFrame = Ybrace
End Function
Public Property Let Value62(Setting)
   Ybrace = Setting
End Property
