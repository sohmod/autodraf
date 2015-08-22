Attribute VB_Name = "Module3"
Option Explicit

Public Function OneShortEDiscoHog(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
   OneShortEDiscoHog = 0.039 + 0.005 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   OneShortEDiscoHog = 0.044 + 0.004 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   OneShortEDiscoHog = 0.048 + 0.004 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
   OneShortEDiscoHog = 0.052 + 0.003 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   OneShortEDiscoHog = 0.055 + 0.003 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   OneShortEDiscoHog = 0.058 + 0.005 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
   OneShortEDiscoHog = 0.063 + 0.004 * (X - 1)
   End If
   
End Function
