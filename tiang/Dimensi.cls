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

Private NamaT, CoverT
Private UTngB, UtngH, UTngHgt
Private UWBeamB, UWBeamH, UWBeamL
Private UEBeamB, UEBeamH, UEBeamL
Private USBeamB, USBeamH, USBeamL
Private UNBeamB, UNBeamH, UNBeamL
Private TiangB, TiangH, TiangHgt
Private LNBeamB, LNBeamH, LNBeamL
Private LSBeamB, LSBeamH, LSBeamL
Private LEBeamB, LEBeamH, LEBeamL
Private LWBeamB, LWBeamH, LWBeamL
Private LTngB, LtngH, LTngHgt

Public Function NamaTiang()
   NamaTiang = NamaT
End Function
Public Property Let Value49(Setting)
   NamaT = Setting
End Property


Public Function Cover()
If Val(CoverT) <= 0 Then
    CoverT = 0
       End If
   Cover = CoverT
End Function
Public Property Let Value50(Setting)
   CoverT = Setting
End Property
''''''''''''


Public Function BtngAtas()
If Val(UTngB) <= 0 Then
    UTngB = 0
       End If
   BtngAtas = UTngB
End Function
Public Property Let Value1(Setting)
   UTngB = Setting
End Property


Public Function HtngAtas()
If Val(UtngH) <= 0 Then
    UtngH = 0
       End If
   HtngAtas = UtngH
End Function
Public Property Let Value2(Setting)
   UtngH = Setting
End Property


Public Function TtngAtas()
If Val(UTngHgt) <= 0 Then
    UTngHgt = 0
       End If
   TtngAtas = UTngHgt
End Function
Public Property Let Value3(Setting)
   UTngHgt = Setting
End Property
''''''''''''


Public Function BrskWAtas()
If Val(UWBeamB) <= 0 Then
    UWBeamB = 0
       End If
   BrskWAtas = UWBeamB
End Function
Public Property Let Value5(Setting)
   UWBeamB = Setting
End Property


Public Function HrskWAtas()
If Val(UWBeamH) <= 0 Then
    UWBeamH = 0
       End If
   HrskWAtas = UWBeamH
End Function
Public Property Let Value6(Setting)
   UWBeamH = Setting
End Property


Public Function PrskWAtas()
If Val(UWBeamL) <= 0 Then
    UWBeamL = 0
       End If
   PrskWAtas = UWBeamL
End Function
Public Property Let Value7(Setting)
   UWBeamL = Setting
End Property
''''''''''''


Public Function BrskEAtas()
If Val(UEBeamB) <= 0 Then
    UEBeamB = 0
       End If
   BrskEAtas = UEBeamB
End Function
Public Property Let Value8(Setting)
   UEBeamB = Setting
End Property


Public Function HrskEAtas()
If Val(UEBeamH) <= 0 Then
    UEBeamH = 0
       End If
   HrskEAtas = UEBeamH
End Function
Public Property Let Value9(Setting)
   UEBeamH = Setting
End Property


Public Function PrskEAtas()
If Val(UEBeamL) <= 0 Then
    UEBeamL = 0
       End If
   PrskEAtas = UEBeamL
End Function
Public Property Let Value10(Setting)
   UEBeamL = Setting
End Property
''''''''''''


Public Function BrskSAtas()
If Val(USBeamB) <= 0 Then
    USBeamB = 0
       End If
   BrskSAtas = USBeamB
End Function
Public Property Let Value11(Setting)
   USBeamB = Setting
End Property


Public Function HrskSAtas()
If Val(USBeamH) <= 0 Then
    USBeamH = 0
       End If
   HrskSAtas = USBeamH
End Function
Public Property Let Value12(Setting)
   USBeamH = Setting
End Property


Public Function PrskSAtas()
If Val(USBeamL) <= 0 Then
    USBeamL = 0
       End If
   PrskSAtas = USBeamL
End Function
Public Property Let Value13(Setting)
   USBeamL = Setting
End Property
''''''''''''


Public Function BrskNAtas()
If Val(UNBeamB) <= 0 Then
    UNBeamB = 0
       End If
   BrskNAtas = UNBeamB
End Function
Public Property Let Value14(Setting)
   UNBeamB = Setting
End Property


Public Function HrskNAtas()
If Val(UNBeamH) <= 0 Then
    UNBeamH = 0
       End If
   HrskNAtas = UNBeamH
End Function
Public Property Let Value15(Setting)
   UNBeamH = Setting
End Property


Public Function PrskNAtas()
If Val(UNBeamL) <= 0 Then
    UNBeamL = 0
       End If
   PrskNAtas = UNBeamL
End Function
Public Property Let Value16(Setting)
   UNBeamL = Setting
End Property
''''''''''''


Public Function BtianG()
If Val(TiangB) <= 0 Then
    TiangB = 50
       End If
   BtianG = TiangB
End Function
Public Property Let Value17(Setting)
   TiangB = Setting
End Property


Public Function HtianG()
If Val(TiangH) <= 0 Then
    TiangH = 50
       End If
   HtianG = TiangH
End Function
Public Property Let Value18(Setting)
   TiangH = Setting
End Property


Public Function TtianG()
If Val(TiangHgt) <= 0 Then
    TiangHgt = 500
       End If
   TtianG = TiangHgt
End Function
Public Property Let Value19(Setting)
   TiangHgt = Setting
End Property
''''''''''''


Public Function BrskNBawah()
If Val(LNBeamB) <= 0 Then
   LNBeamB = 0
       End If
   BrskNBawah = LNBeamB
End Function
Public Property Let Value24(Setting)
   LNBeamB = Setting
End Property


Public Function HrskNBawah()
If Val(LNBeamH) <= 0 Then
    LNBeamH = 0
       End If
   HrskNBawah = LNBeamH
End Function
Public Property Let Value25(Setting)
   LNBeamH = Setting
End Property


Public Function PrskNBawah()
If Val(LNBeamL) <= 0 Then
    LNBeamL = 0
       End If
   PrskNBawah = LNBeamL
End Function
Public Property Let Value26(Setting)
   LNBeamL = Setting
End Property
''''''''''''


Public Function BrskSBawah()
If Val(LSBeamB) <= 0 Then
    LSBeamB = 0
       End If
   BrskSBawah = LSBeamB
End Function
Public Property Let Value27(Setting)
   LSBeamB = Setting
End Property


Public Function HrskSBawah()
If Val(LSBeamH) <= 0 Then
    LSBeamH = 0
       End If
   HrskSBawah = LSBeamH
End Function
Public Property Let Value28(Setting)
   LSBeamH = Setting
End Property


Public Function PrskSBawah()
If Val(LSBeamL) <= 0 Then
    LSBeamL = 0
       End If
   PrskSBawah = LSBeamL
End Function
Public Property Let Value29(Setting)
   LSBeamL = Setting
End Property
''''''''''''


Public Function BrskEBawah()
If Val(LEBeamB) <= 0 Then
    LEBeamB = 0
       End If
   BrskEBawah = LEBeamB
End Function
Public Property Let Value30(Setting)
   LEBeamB = Setting
End Property


Public Function HrskEBawah()
If Val(LEBeamH) <= 0 Then
    LEBeamH = 0
       End If
   HrskEBawah = LEBeamH
End Function
Public Property Let Value31(Setting)
   LEBeamH = Setting
End Property


Public Function PrskEBawah()
If Val(LEBeamL) <= 0 Then
    LEBeamL = 0
       End If
   PrskEBawah = LEBeamL
End Function
Public Property Let Value32(Setting)
   LEBeamL = Setting
End Property
''''''''


Public Function BrskWBawah()
If Val(LWBeamB) <= 0 Then
    LWBeamB = 0
       End If
   BrskWBawah = LWBeamB
End Function
Public Property Let Value33(Setting)
   LWBeamB = Setting
End Property


Public Function HrskWBawah()
If Val(LWBeamH) <= 0 Then
    LWBeamH = 0
       End If
   HrskWBawah = LWBeamH
End Function
Public Property Let Value34(Setting)
   LWBeamH = Setting
End Property


Public Function PrskWBawah()
If Val(LWBeamL) <= 0 Then
    LWBeamL = 0
       End If
   PrskWBawah = LWBeamL
End Function
Public Property Let Value35(Setting)
   LWBeamL = Setting
End Property
''''''''


Public Function BtngBawah()
If Val(LTngB) <= 0 Then
    LTngB = 0
       End If
   BtngBawah = LTngB
End Function
Public Property Let Value36(Setting)
   LTngB = Setting
End Property


Public Function HtngBawah()
If Val(LtngH) <= 0 Then
    LtngH = 0
       End If
   HtngBawah = LtngH
End Function
Public Property Let Value37(Setting)
   LtngH = Setting
End Property


Public Function TtngBawah()
If Val(LTngHgt) <= 0 Then
    LTngHgt = 0
       End If
   TtngBawah = LTngHgt
End Function
Public Property Let Value38(Setting)
   LTngHgt = Setting
End Property
''''''''
