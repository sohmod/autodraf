Attribute VB_Name = "Module4"
Option Explicit

Public Function OneShortEDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
   OneShortEDiscoSag = 0.029 + 0.004 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
   OneShortEDiscoSag = 0.033 + 0.003 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
   OneShortEDiscoSag = 0.036 + 0.003 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
   OneShortEDiscoSag = 0.039 + 0.002 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
   OneShortEDiscoSag = 0.041 + 0.002 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   OneShortEDiscoSag = 0.043 + 0.004 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
   OneShortEDiscoSag = 0.047 + 0.003 * (X - 1)
   End If
   
End Function
