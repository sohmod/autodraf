Attribute VB_Name = "Module14"
Option Explicit

Public Function ThreeELongDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    ThreeELongDiscoSag = 0.043 + 0.005 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    ThreeELongDiscoSag = 0.048 + 0.005 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    ThreeELongDiscoSag = 0.053 + 0.004 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    ThreeELongDiscoSag = 0.057 + 0.003 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    ThreeELongDiscoSag = 0.06 + 0.003 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    ThreeELongDiscoSag = 0.063 + 0.006 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    ThreeELongDiscoSag = 0.069 + 0.005 * (X - 1)
   End If
   
End Function
