VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tetulang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FLTBno, FLTBdia, FLTBcurt, SLTBno, SLTBdia, SLTBcurt, LinkLHS
Private LBCBno, LBCBdia, LBCBcurE
Private FMBBno, FMBBdia, SMBBcurS, SMBBno, SMBBdia, SMBBcurE, LinkMID
Private RBCBno, RBCBdia, RBCBcurS
Private FRTBno, FRTBdia, FRTBcurt, SRTBno, SRTBdia, SRTBcurt, LinkRHS
Private LINKCBno, LINKCBdia


Public Function FirstLTBno()
        If FLTBno < 2 Or Val(FLTBno) = 0 Then
          MsgBox "LHS bar no. Default 2.", , "Amaran :"
          FLTBno = 2
        End If
   FirstLTBno = FLTBno
End Function
Public Property Let Value14(Setting)
   FLTBno = Setting
End Property


Public Function FirstLTBdia()
         If FLTBdia <= 0 Or Val(FLTBdia) = 0 Then
          MsgBox "LHS bar dia. Default min 12mm.", , "Amaran:"
          FLTBdia = 12
        End If
        If FLTBdia > 40 Then
          MsgBox "LHS bar dia. Default max 40mm.", , "Amaran:"
          FLTBdia = 40
        End If
   FirstLTBdia = FLTBdia
End Function
Public Property Let Value15(Setting)
   FLTBdia = Setting
End Property


Public Function FirstLTBcurt()
    FirstLTBcurt = FLTBcurt
End Function
Public Property Let Value16(Setting)
   FLTBcurt = Setting
End Property


Public Function SecondLTBno()
      If SLTBno <= 0 Or Val(SLTBno) = 0 Then
          SLTBno = 0
          SLTBdia = 0
          SLTBcurt = 0
             End If
             
      SecondLTBno = Int(SLTBno)
End Function
Public Property Let Value17(Setting)
   SLTBno = Setting
End Property


Public Function SecondLTBdia()
    SecondLTBdia = SLTBdia
End Function
Public Property Let Value18(Setting)
   SLTBdia = Setting
End Property


Public Function SecondLTBcurt()
    SecondLTBcurt = SLTBcurt
End Function
Public Property Let Value19(Setting)
   SLTBcurt = Setting
End Property


Public Function LinkSpacingLHS()
    LinkSpacingLHS = LinkLHS
End Function
Public Property Let Value20(Setting)
   LinkLHS = Setting
End Property


'''''''''''''''''''''''''''''''''''''

Public Function FirstMBBno()
   FirstMBBno = FMBBno
End Function
Public Property Let Value21(Setting)
   FMBBno = Setting
End Property



Public Function FirstMBBdia()
    If FMBBdia > 40 Then
         MsgBox "MIDSPAN bar dia. Default max 40mm.", , "Amaran :"
           FMBBdia = 40
             End If
   FirstMBBdia = FMBBdia
End Function
Public Property Let Value22(Setting)
   FMBBdia = Setting
End Property


Public Function SecondMBBcurtS()
    SecondMBBcurtS = SMBBcurS
End Function
Public Property Let Value23(Setting)
   SMBBcurS = Setting
End Property


Public Function SecondMBBno()
   SecondMBBno = Int(SMBBno)
End Function
Public Property Let Value24(Setting)
   SMBBno = Setting
End Property


Public Function SecondMBBdia()
   If SMBBdia > 40 Then
           SMBBdia = 40
              End If
   SecondMBBdia = SMBBdia
End Function
Public Property Let Value25(Setting)
   SMBBdia = Setting
End Property



Public Function SecondMBBcurtE()
    SecondMBBcurtE = Int(SMBBcurE)
End Function
Public Property Let Value26(Setting)
   SMBBcurE = Setting
End Property


Public Function LinkSpacingMID()
    LinkSpacingMID = LinkMID
End Function
Public Property Let Value27(Setting)
   LinkMID = Setting
End Property

''''''''''''''''''''''''''''''''''''

Public Function FirstRTBno()
   FirstRTBno = FRTBno
End Function
Public Property Let Value28(Setting)
   FRTBno = Setting
End Property


Public Function FirstRTBdia()
    If FRTBdia > 40 Then
          FRTBdia = 40
        End If
   FirstRTBdia = FRTBdia
End Function
Public Property Let Value29(Setting)
   FRTBdia = Setting
End Property


Public Function FirstRTBcurt()
    FirstRTBcurt = FRTBcurt
End Function
Public Property Let Value30(Setting)
   FRTBcurt = Setting
End Property


Public Function SecondRTBno()
   SecondRTBno = SRTBno
End Function
Public Property Let Value31(Setting)
   SRTBno = Setting
End Property


Public Function SecondRTBdia()
   If SRTBdia > 40 Then
          SRTBdia = 40
             End If
   SecondRTBdia = SRTBdia
End Function
Public Property Let Value32(Setting)
   SRTBdia = Setting
End Property



Public Function SecondRTBcurt()
   SecondRTBcurt = SRTBcurt
End Function
Public Property Let Value33(Setting)
   SRTBcurt = Setting
End Property



Public Function LinkSpacingRHS()
   LinkSpacingRHS = LinkRHS
End Function
Public Property Let Value34(Setting)
   LinkRHS = Setting
End Property

'''''''''''''''***********************'''''''''''''''''''''

Public Function LinkCarrierNo()
     LinkCarrierNo = Int(LINKCBno)
End Function
Public Property Let Value49(Setting)
   LINKCBno = Setting
End Property


Public Function LinkCarrierDia()
    If LINKCBdia > 40 Then
           LINKCBdia = 40
               End If
   LinkCarrierDia = Int(LINKCBdia)
End Function
Public Property Let Value50(Setting)
   LINKCBdia = Setting
End Property


Public Function LeftConnBarNo()
    LeftConnBarNo = Int(LBCBno)
End Function
Public Property Let Value51(Setting)
   LBCBno = Setting
End Property



Public Function LeftConnBarDia()
   If LBCBdia > 40 Then
            LBCBdia = 40
               End If
   LeftConnBarDia = Int(LBCBdia)
End Function
Public Property Let Value52(Setting)
   LBCBdia = Setting
End Property


Public Function LeftConnBarCurtE()
    LeftConnBarCurtE = Int(LBCBcurE)
End Function
Public Property Let Value53(Setting)
   LBCBcurE = Setting
End Property


Public Function RightConnBarNo()
     RightConnBarNo = Int(RBCBno)
End Function
Public Property Let Value54(Setting)
   RBCBno = Setting
End Property



Public Function RightConnBarDia()
   If RBCBdia > 40 Then
             RBCBdia = 40
               End If
   RightConnBarDia = Int(RBCBdia)
End Function
Public Property Let Value55(Setting)
   RBCBdia = Setting
End Property


Public Function RightConnBarCurtS()
    RightConnBarCurtS = Int(RBCBcurS)
End Function
Public Property Let Value56(Setting)
   RBCBcurS = Setting
End Property

