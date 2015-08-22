Attribute VB_Name = "Module9"
Option Explicit

Public Function TwoShortEDiscoHog(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    TwoShortEDiscoHog = 0.046 + 0.004 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    TwoShortEDiscoHog = 0.05 + 0.004 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    TwoShortEDiscoHog = 0.054 + 0.003 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    TwoShortEDiscoHog = 0.057 + 0.003 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    TwoShortEDiscoHog = 0.06 + 0.002 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    TwoShortEDiscoHog = 0.062 + 0.005 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    TwoShortEDiscoHog = 0.067 + 0.003 * (X - 1)
   End If
   
End Function
