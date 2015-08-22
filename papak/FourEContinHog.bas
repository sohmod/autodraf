Attribute VB_Name = "Module1"
Option Explicit

Public Function FourEContinHog(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    FourEContinHog = 0.031 + 0.06 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    FourEContinHog = 0.037 + 0.05 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    FourEContinHog = 0.042 + 0.04 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    FourEContinHog = 0.046 + 0.04 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    FourEContinHog = 0.05 + 0.03 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    FourEContinHog = 0.053 + 0.06 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    FourEContinHog = 0.059 + 0.04 * (X - 1)
   End If
   
End Function
