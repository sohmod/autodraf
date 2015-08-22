VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dimensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private LeGrid
Private LeCol, LeB, LeH
Private L, b, h, SD, BU, sdT, buT, Fslab, Bslab
Private RiCol, RiB, RiH
Private RiGrid

Public Function LeftGrid()
   LeftGrid = LeGrid
End Function
Public Property Let Value1(Setting)
   LeGrid = Setting
End Property


Public Function LeftColumnSize()
If Val(LeCol) < 0 Then
   LeCol = 0
     End If
        LeftColumnSize = LeCol
End Function
Public Property Let Value2(Setting)
   LeCol = Setting
End Property


Public Function LeftXbeamB()
If Val(LeB) < 0 Then
   LeB = 0
     End If
        LeftXbeamB = LeB
End Function
Public Property Let Value3(Setting)
   LeB = Setting
End Property


Public Function LeftXbeamH()
If Val(LeH) < 0 Then
   LeH = 0
     End If
   LeftXbeamH = LeH
End Function
Public Property Let Value4(Setting)
   LeH = Setting
End Property
'''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
Public Function BeamLength()
   If Val(L) <= 0 Then
      MsgBox "L must be greater than zero! Default value = 10 x H akan digunakan.", , "Careful sikit:"
      L = 10 * beamH(NoOfSpan)
   End If
   BeamLength = L
End Function
Public Property Let Value5(Setting)
   L = Setting
End Property


Public Function BeamBreadth()
   If Val(b) <= 0 Then
      MsgBox "B must be greater than zero! Default value = H/3 atau 50mm akan digunakan.", , "Careful sikit:"
       b = Int(beamH(NoOfSpan) / 3)
         End If
   BeamBreadth = b
End Function
Public Property Let Value6(Setting)
   b = Setting
End Property


Public Function BeamTopDrop()
''If Val(SD) < 0 Then
''   SD = 0
''     End If
   BeamTopDrop = Val(SD)
End Function
Public Property Let Value8(Setting)
    SD = Setting
End Property


Public Function BeamSoffitDrop()
''If Val(BU) < 0 Then
''   BU = 0
''     End If
   BeamSoffitDrop = Val(BU)
End Function
Public Property Let Value9(Setting)
   BU = Setting
End Property


Public Function BeamHeight()
   If Val(h) <= 0 Then
      MsgBox "H must be greater than zero! Default value = 3 x B atau 100mm akan digunakan.", , "Careful sikit:"
       h = Int(3 * beamB(NoOfSpan))
        End If
   ''' h = hRef - BU - SD
   BeamHeight = h
End Function
Public Property Let Value7(Setting)
   h = Setting
End Property


Public Function FrontSlabLv()
   If Val(Fslab) <= 0 Then
      Fslab = 0
        End If
   FrontSlabLv = Fslab
End Function
Public Property Let Value47(Setting)
   Fslab = Setting
End Property


Public Function BackSlabLv()
   If Val(Bslab) <= 0 Then
      Bslab = 0
        End If
   BackSlabLv = Bslab
End Function
Public Property Let Value48(Setting)
   Bslab = Setting
End Property

'''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
Public Function RightXbeamb()
If Val(RiB) <= 0 Then
   RiB = 0
     End If
   RightXbeamb = RiB
End Function
Public Property Let Value10(Setting)
   RiB = Setting
End Property


Public Function RightXbeamh()
If Val(RiH) <= 0 Then
   RiH = 0
     End If
   RightXbeamh = RiH
End Function
Public Property Let Value11(Setting)
   RiH = Setting
End Property


Public Function RightColumnSize()
If Val(RiCol) <= 0 Then
   RiCol = 0
     End If
   RightColumnSize = RiCol
End Function
Public Property Let Value12(Setting)
   RiCol = Setting
End Property


Public Function RightGrid()
   RightGrid = RiGrid
End Function
Public Property Let Value13(Setting)
   RiGrid = Setting
End Property
