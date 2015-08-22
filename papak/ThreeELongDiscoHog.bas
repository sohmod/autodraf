Attribute VB_Name = "Module13"
Option Explicit

Public Function ThreeELongDiscoHog(ByVal X As Double)
''ok
If X >= 1 And X < 1.1 Then
    ThreeELongDiscoHog = 0.057 + 0.008 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    ThreeELongDiscoHog = 0.065 + 0.006 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    ThreeELongDiscoHog = 0.071 + 0.005 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    ThreeELongDiscoHog = 0.076 + 0.005 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    ThreeELongDiscoHog = 0.081 + 0.003 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
    ThreeELongDiscoHog = 0.084 + 0.008 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    ThreeELongDiscoHog = 0.092 + 0.006 * (X - 1)
   End If
   
End Function
