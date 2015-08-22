Attribute VB_Name = "Module17"
Option Explicit

Public Function FourEdgesDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    FourEdgesDiscoSag = 0.055 + 0.01 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    FourEdgesDiscoSag = 0.065 + 0.009 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    FourEdgesDiscoSag = 0.074 + 0.007 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    FourEdgesDiscoSag = 0.081 + 0.006 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    FourEdgesDiscoSag = 0.087 + 0.005 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    FourEdgesDiscoSag = 0.092 + 0.011 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    FourEdgesDiscoSag = 0.103 + 0.008 * (X - 1)
   End If
   
End Function
