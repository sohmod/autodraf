Attribute VB_Name = "Module2"
Option Explicit

Public Function FourEContinSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
   FourEContinSag = 0.024 + 0.004 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   FourEContinSag = 0.028 + 0.004 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   FourEContinSag = 0.032 + 0.003 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
   FourEContinSag = 0.035 + 0.002 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   FourEContinSag = 0.037 + 0.003 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   FourEContinSag = 0.04 + 0.004 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
   FourEContinSag = 0.044 + 0.004 * (X - 1)
   End If
   
End Function

