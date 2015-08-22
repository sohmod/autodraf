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


Private BarXBn, BarXBd
Private BarXTn, BarXTd
Private BarYEn, BarYEd
Private BarYWn, BarYWd

Public Function BarOnXSno()
If BarXBn <= 0 Or Val(BarXBn) = 0 Then
   BarXBn = 0
     End If
   BarOnXSno = Int(BarXBn)
End Function
Public Property Let Value47(Setting)
   BarXBn = Setting
End Property

Public Function BarOnXSdia()
If Val(BarXBd) <= 0 Then
   BarXBd = 0
     End If
If Val(BarXBd) > 40 Then
   BarXBd = 40
     End If
   BarOnXSdia = Int(BarXBd)
End Function
Public Property Let Value48(Setting)
   BarXBd = Setting
End Property
''''''
Public Function BarOnXNno()
If BarXTn <= 0 Or Val(BarXTn) = 0 Then
   BarXTn = 0
     End If
   BarOnXNno = Int(BarXTn)
End Function
Public Property Let Value45(Setting)
   BarXTn = Setting
End Property

Public Function BarOnXNdia()
If Val(BarXTd) <= 0 Then
   BarXTd = 0
     End If
If Val(BarXTd) > 40 Then
   BarXTd = 40
     End If
   BarOnXNdia = Int(BarXTd)
End Function
Public Property Let Value46(Setting)
   BarXTd = Setting
End Property
''''''
Public Function BarOnYWno()
If BarYWn <= 0 Or Val(BarYWn) = 0 Then
   BarYWn = 0
     End If
   BarOnYWno = Int(BarYWn)
End Function
Public Property Let Value41(Setting)
   BarYWn = Setting
End Property

Public Function BarOnYWdia()
If Val(BarYWd) <= 0 Then
   BarYWd = 0
     End If
If Val(BarYWd) > 40 Then
   BarYWd = 40
     End If
   BarOnYWdia = Int(BarYWd)
End Function
Public Property Let Value42(Setting)
   BarYWd = Setting
End Property
''''''
Public Function BarOnYEno()
If BarYEn <= 0 Or Val(BarYEn) = 0 Then
   BarYEn = 0
     End If
   BarOnYEno = Int(BarYEn)
End Function
Public Property Let Value43(Setting)
   BarYEn = Setting
End Property

Public Function BarOnYEdia()
If Val(BarYEd) <= 0 Then
   BarYEd = 0
     End If
If Val(BarYEd) > 40 Then
   BarYEd = 40
     End If
   BarOnYEdia = Int(BarYEd)
End Function
Public Property Let Value44(Setting)
   BarYEd = Setting
End Property
''''''


