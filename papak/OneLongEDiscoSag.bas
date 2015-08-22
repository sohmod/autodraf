Attribute VB_Name = "Module6"
Option Explicit

Public Function OneLongEDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
   OneLongEDiscoSag = 0.03 + 0.006 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   OneLongEDiscoSag = 0.036 + 0.006 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   OneLongEDiscoSag = 0.042 + 0.005 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
   OneLongEDiscoSag = 0.047 + 0.004 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    OneLongEDiscoSag = 0.051 + 0.04 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   OneLongEDiscoSag = 0.055 + 0.007 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
   OneLongEDiscoSag = 0.062 + 0.005 * (X - 1)
   End If
   
End Function
