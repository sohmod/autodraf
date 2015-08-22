VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Xins, Yins, Nrsk
Private fcuT, fyT, fyvT, Shrk, Crp
Private cvrT, Slab, LnkD, Bmark, FontSize

Public Function XinsertPt()
If Val(Xins) = 0 Then
   Xins = 0
     End If
   XinsertPt = Xins
End Function
Public Property Let Value35(Setting)
   Xins = Setting
End Property


Public Function YinsertPt()
If Val(Yins) = 0 Then
   Yins = 0
     End If
   YinsertPt = Yins
End Function
Public Property Let Value36(Setting)
   Yins = Setting
End Property


Public Function SetFontSize()
FontSize = Trim(FontSize)
If Val(FontSize) <= 0 Then
   FontSize = 40
      End If
   SetFontSize = FontSize
End Function
Public Property Let Value57(Setting)
   FontSize = Setting
End Property





Public Function NamaRasuk()
   NamaRasuk = Nrsk
End Function
Public Property Let Value37(Setting)
   Nrsk = Setting
End Property


Public Function fcu()
   If Val(fcuT) <= 0 Then
      fcuT = 20
         End If
   If Val(fcuT) >= 60 Then
      fcuT = 60
         End If
   fcu = fcuT
End Function
Public Property Let Value38(Setting)
   fcuT = Setting
End Property


Public Function fy()
    If Val(fyT) < 250 Then
      fyT = 250
         End If
   If Val(fyT) > 480 Then
      fyT = 480
         End If
   fy = fyT
End Function
Public Property Let Value39(Setting)
   fyT = Setting
End Property


Public Function fyv()
    If Val(fyvT) < 250 Then
      fyvT = 250
         End If
   If Val(fyvT) > 460 Then
      fyvT = 460
         End If
   fyv = fyvT
End Function
Public Property Let Value40(Setting)
   fyvT = Setting
End Property


Public Function Shrink()
    If Val(Shrk) < 0.0007 Then
      Shrk = 0.0007
         End If
   If Val(Shrk) > 0.0003 Then
      Shrk = 0.0003
         End If
   Shrink = Shrk
End Function
Public Property Let Value41(Setting)
   Shrk = Setting
End Property


Public Function Creep()
    If Val(Crp) < 0.5 Then
      Crp = 0.5
         End If
   If Val(Crp) > 4# Then
      Crp = 4#
         End If
   Creep = Crp
End Function
Public Property Let Value42(Setting)
   Crp = Setting
End Property


Public Function Cover()
    If Val(cvrT) <= 0 Then
      cvrT = 1
         End If
   If Val(cvrT) >= 100 Then
      cvrT = 100
         End If
   Cover = cvrT
End Function
Public Property Let Value43(Setting)
   cvrT = Setting
End Property


Public Function SlabThk()
   If Val(Slab) <= 0 Then
     Slab = 0
        End If
    If Val(Slab) > beamH(NoOfSpan) Then
      Slab = Int(beamH(NoOfSpan) / 3)
         End If
   SlabThk = Slab
End Function
Public Property Let Value44(Setting)
   Slab = Setting
End Property


Public Function LinkD()
    If Val(LnkD) <= 0 Then
      LnkD = 6
         End If
   If Val(LnkD) > 16 Then
      LnkD = 16
         End If
   LinkD = LnkD
End Function
Public Property Let Value45(Setting)
   LnkD = Setting
End Property

Public Function BarMark()
If Val(Bmark) < 0 Then
   Bmark = 1
     End If
   BarMark = Bmark
End Function
Public Property Let Value46(Setting)
   Bmark = Setting
End Property

