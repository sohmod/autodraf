Attribute VB_Name = "Module8"
Option Explicit

Public Function TwoAdjEDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    TwoAdjEDiscoSag = 0.036 + 0.006 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   TwoAdjEDiscoSag = 0.042 + 0.005 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   TwoAdjEDiscoSag = 0.047 + 0.004 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    TwoAdjEDiscoSag = 0.051 + 0.004 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    TwoAdjEDiscoSag = 0.055 + 0.004 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    TwoAdjEDiscoSag = 0.059 + 0.006 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    TwoAdjEDiscoSag = 0.065 + 0.005 * (X - 1)
   End If
   
End Function
