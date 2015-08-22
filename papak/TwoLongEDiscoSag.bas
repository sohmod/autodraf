Attribute VB_Name = "Module12"
Option Explicit

Public Function TwoLongEDiscoSag(ByVal X As Double)
'''ok
If X >= 1 And X < 1.1 Then
    TwoLongEDiscoSag = 0.034 + 0.012 * (X - 1)
   End If
If X >= 1.1 And X < 1.2 Then
    TwoLongEDiscoSag = 0.046 + 0.01 * (X - 1)
   End If
If X >= 1.2 And X < 1.3 Then
    TwoLongEDiscoSag = 0.056 + 0.009 * (X - 1)
   End If
If X >= 1.3 And X < 1.4 Then
    TwoLongEDiscoSag = 0.065 + 0.007 * (X - 1)
   End If
If X >= 1.4 And X < 1.5 Then
    TwoLongEDiscoSag = 0.072 + 0.006 * (X - 1)
   End If
 If X >= 1.5 And X < 1.75 Then
   TwoLongEDiscoSag = 0.078 + 0.013 * (X - 1)
   End If
 If X >= 1.75 And X <= 2 Then
    TwoLongEDiscoSag = 0.091 + 0.009 * (X - 1)
   End If
   
End Function
