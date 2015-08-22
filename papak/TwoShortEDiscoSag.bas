Attribute VB_Name = "Module10"
Option Explicit

Public Function TwoShortEDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    TwoShortEDiscoSag = 0.034 + 0.004 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    TwoShortEDiscoSag = 0.038 + 0.002 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    TwoShortEDiscoSag = 0.04 + 0.003 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    TwoShortEDiscoSag = 0.043 + 0.002 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    TwoShortEDiscoSag = 0.045 + 0.002 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    TwoShortEDiscoSag = 0.047 + 0.003 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    TwoShortEDiscoSag = 0.05 + 0.003 * (X - 1)
   End If
   
End Function
