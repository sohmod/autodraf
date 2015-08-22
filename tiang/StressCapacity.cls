VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StressCapacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Xin, Yin
Private fcuCM, fyCM, fyvCM
Private Shr, Crp
Private LnkD, LnkS
Private BarM


Public Function XinsertPoint()
If Val(Xin) = 0 Then
    Xin = 0
       End If
   XinsertPoint = Xin
End Function
Public Property Let Value51(Setting)
   Xin = Setting
End Property


Public Function YinsertPoint()
If Val(Yin) = 0 Then
    Yin = 0
       End If
   YinsertPoint = Yin
End Function
Public Property Let Value52(Setting)
   Yin = Setting
End Property
''''''

Public Function fcuNEW()
If Val(fcuCM) <= 0 Then
    fcuCM = 25
       End If
If Val(fcuCM) > 60 Then
    fcuCM = 60
       End If
   fcuNEW = fcuCM
End Function
Public Property Let Value53(Setting)
   fcuCM = Setting
End Property


Public Function fyNEW()
If Val(fyCM) <= 0 Then
    fyCM = 410
       End If
If Val(fyCM) > 485 Then
    fyCM = 485
       End If
   fyNEW = fyCM
End Function
Public Property Let Value54(Setting)
   fyCM = Setting
End Property


Public Function fyvNEW()
If Val(fyvCM) <= 0 Then
    fyvCM = 250
       End If
If Val(fyvCM) > 485 Then
    fyvCM = 485
       End If
   fyvNEW = fyvCM
End Function
Public Property Let Value55(Setting)
   fyvCM = Setting
End Property
''''''


Public Function ShrinkNEW()
If Val(Shr) <= 0 Then
    Shr = 0.00005
       End If
If Val(Shr) > 0.0004 Then
    Shr = 0.0004
       End If
   ShrinkNEW = Shr
End Function
Public Property Let Value56(Setting)
   Shr = Setting
End Property


Public Function CreepNEW()
If Val(Crp) <= 0 Then
    Crp = 0.5
       End If
If Val(Crp) > 4 Then
    Crp = 4
       End If
   CreepNEW = Crp
End Function
Public Property Let Value57(Setting)
   Crp = Setting
End Property


Public Function LinkDiameter()
If Val(LnkD) <= 0 Then
    LnkD = 10
       End If
If Val(LnkD) > 40 Then
    LnkD = 40
       End If
   LinkDiameter = LnkD
End Function
Public Property Let Value58(Setting)
   LnkD = Setting
End Property


Public Function LinkSpacing()
If Val(LnkS) <= 0 Then
    LnkS = 100
       End If
   LinkSpacing = LnkS
End Function
Public Property Let Value59(Setting)
   LnkS = Setting
End Property


Public Function BarMarkNEW()
If Val(BarM) <= 0 Then
    BarM = 1
       End If
   BarMarkNEW = Int(BarM)
End Function
Public Property Let Value60(Setting)
   BarM = Setting
End Property
''''''
