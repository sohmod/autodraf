Attribute VB_Name = "Module7"
Option Explicit

Public Function TwoAdjEDiscoHog(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
   TwoAdjEDiscoHog = 0.047 + 0.009 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   TwoAdjEDiscoHog = 0.056 + 0.007 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   TwoAdjEDiscoHog = 0.063 + 0.006 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
   TwoAdjEDiscoHog = 0.069 + 0.005 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   TwoAdjEDiscoHog = 0.074 + 0.004 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    TwoAdjEDiscoHog = 0.078 + 0.009 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
   TwoAdjEDiscoHog = 0.087 + 0.006 * (X - 1)
   End If
   
End Function
