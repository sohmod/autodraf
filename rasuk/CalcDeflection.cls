VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcDeflection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CalcDeflection( _
ByVal L As Double, _
ByVal c4, c5, c6, c7, c8, c9 As Double) As Double

1390 Dim c As Integer
1392 Dim bt, dM, dn, dft As Double
1400 Rem total deflection
1402 '''MsgBox "shape parabolic="
1403 c = 1
1404 '''input #1, " 1=para  2=tria";c
1406 bt = (c6 + c8) / c4
1408 dM = L ^ 2 * c4
1409 dn = L ^ 2 * (c5 + c7 + c9) / 3 / 8
1410 '''On c GoTo 1412, 1414
1412 dft = (1 + bt / 10) * dM / 9.6 + dn  '''': GoTo 1418
1414 '''dft = (1 + bt / 4) * dm / 12 + dn: GoTo 1418
1418 '''Return
1419 CalcDeflection = Int(dft * 100) / 100
''''''''''

End Function

