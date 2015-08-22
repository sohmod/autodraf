Attribute VB_Name = "Module5"
Option Explicit

Public Function OneLongEDiscoHog(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    OneLongEDiscoHog = 0.039 + 0.01 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   OneLongEDiscoHog = 0.049 + 0.007 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    OneLongEDiscoHog = 0.056 + 0.006 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    OneLongEDiscoHog = 0.062 + 0.006 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   OneLongEDiscoHog = 0.068 + 0.005 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   OneLongEDiscoHog = 0.073 + 0.009 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
  OneLongEDiscoHog = 0.082 + 0.007 * (X - 1)
   End If
   
End Function
