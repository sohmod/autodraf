VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACurvature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ACurvature( _
ByVal CurveNo As Integer, _
ByVal fcu As Double, _
ByVal Shrink As Double, _
ByVal Creep As Double, _
ByVal Mtt As Double, _
ByVal Mpm As Double, _
ByVal b As Double, _
ByVal h As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal ast As Double, _
ByVal asc As Double, _
ByVal SLTerm As String) As Double

Dim ec, es, el, rd, rt, rc, rs, rl As Double
Dim i1, i2, i3, i4, i5, i6, i7 As Double
Dim x, x111, x222, x333, s As Double
Dim c1, c2, c3, c4, c5 As Double
Dim fct, xsl, isl, ae, mfct, mtn, mpn, fcc, fst, fsc As Double

1330 Rem deflection
1331 ec = 24 + (fcu - 20) / 5
1332 es = 200: rd = dp / d: rt = ast / b / d
1334 rc = asc / b / d: el = ec / (1 + Creep)
1336 rs = es / ec: rl = es / el
1338 i1 = -1 * rs * rt - (rs - 1) * rc
1340 i2 = rs * rt * (rs * rt + 2) + 2 * (rs - 1) * rc * (rs * rt + rd)
1342 i3 = i1 + Sqr(i2)
1344 i4 = i3 ^ 3 / 3 + rs * rt * (1 - i3) ^ 2 + (rs - 1) * rc * (i3 - rd) ^ 2
1346 i5 = i4 * b * d ^ 3
1348 x111 = -1 * rl * rt - (rl - 1) * rc
1350 x222 = rl * rt * (rl * rt + 2) + 2 * (rl - 1) * rc * (rl * rt + rd)
1352 x333 = x111 + Sqr(x222)
1354 x = x333 * d
1356 i6 = x333 ^ 3 / 3 + rl * rt * (1 - x333) ^ 2 + (rl - 1) * rc * (x333 - rd) ^ 2
1358 i7 = i6 * b * d ^ 3
1360 s = ast * (d - x) - asc * (x - dp)
1361 GoTo 1470
1362 c1 = (mtn - mpn) * 1000# / i5 / ec
1364 c2 = mpn * 1000# / i7 / el
1366 c3 = Shrink * rl * s / i7
1372 c4 = c1 + c2     ''c6  ''c8
1373 '''If sltd$ = "s" Then c4 = mtn * 1000# / i5 / ec
1374 c5 = c3          ''c7  ''c9
     If CurveNo = 1 Then
        ACurvature = c4 '''If sltd$ = "s" Then c5 = 0
           End If
     If CurveNo = 2 Then
        ACurvature = c5
           End If
1376 GoTo 1510
''''''''''''''''''

1470 Rem calc mnet
1472 fct = 0.55: xsl = x: isl = i7: ae = rl
1474 '''If sltd$ = "s" Then fct = 1
1475 '''If sltd$ = "s" Then isl = i5
1476 '''If sltd$ = "s" Then xsl = i3 * d
1478 mfct = fct * b * (h - xsl) ^ 3 / (d - xsl) / 3 / 1000000#
1480 mtn = Mtt - mfct
1482 mpn = Mpm - mfct
1484 '''If sltd$ = "s" Then ae = rs
1486 fcc = mtn * 1000000# * xsl / isl
1488 fst = ae * mtn * 1000000# * (d - xsl) / isl
1490 fsc = ae * mtn * 1000000# * (xsl - dp) / isl
1500 GoTo 1362
1510 Rem
End Function

