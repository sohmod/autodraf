Attribute VB_Name = "Module16"
Option Explicit

Public Function ThreeEShortDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    ThreeEShortDiscoSag = 0.042 + 0.012 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    ThreeEShortDiscoSag = 0.054 + 0.009 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    ThreeEShortDiscoSag = 0.063 + 0.008 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    ThreeEShortDiscoSag = 0.071 + 0.007 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   ThreeEShortDiscoSag = 0.078 + 0.006 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    ThreeEShortDiscoSag = 0.084 + 0.012 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    ThreeEShortDiscoSag = 0.096 + 0.009 * (X - 1)
   End If
   
End Function
