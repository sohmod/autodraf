VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcMoment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private fc, fy, bf, bw, hf, d, dp, t, c, pu, ap, dt, k
''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function CalcMoment( _
ByVal fc As Double, _
ByVal fy As Double, _
ByVal bf As Double, _
ByVal bw As Double, _
ByVal hf As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal t As Double, _
ByVal c As Double, _
ByVal pu As Double, _
ByVal ap As Double, _
ByVal dt As Double, _
ByVal k As Double) As Integer

Dim g, u As Integer
Dim k8, k9, er, e1, e2, e3, e4, e5, e6 As Double
Dim w, w9, r, r9, f4, f5, f6, x, A, mu As Double

''MsgBox Str(t) & "   " & Str(c) & "   " & Str(d), , " Ast   Asc    d"

103: If hf > d Or hf = 0 Then hf = d
104: If hf = 0 Then bw = bf
105: k8 = 0.45
110: k9 = 0.9
115: er = (1 - k / 1000) * 0.75 * pu / 200000#
120: If t = 0 Then d = 0
130: e1 = fy / 1.15 / 200000#
135: e2 = 4 * pu / 1000000# / 1.15
140: e3 = 0.005 + 5 * pu / 1000000# / 1.15
145: u = 100: w = 1: r = 0: g = 1
150: If pu = 0 Or ap = 0 Then dt = d
155: x = dp + 1
170: x = x - dt / u
175: x = x + dt / u
180: w9 = w: r9 = r
185: f4 = 0: f6 = 0
190: If t <> 0 Then GoSub 340
195: If ap <> 0 Then GoSub 370
200: If c <> 0 Then GoSub 400
205: A = bf * k9 * x
210: If k9 * x > hf Then A = bf * hf + bw * (k9 * x - hf)
215: r = k8 * fc * A + c * f6
220: w = t * f4 + ap * f5
225: If w - r > 0 And w9 - r9 > 0 Then GoTo 175
230: If g = 2 Then GoTo 245
235: x = x - dt / u: u = 2000: g = 2
240: w = 1: r = 0: GoTo 170
245: If k9 * x > hf Then GoTo 260
250: mu = k8 * fc * bf * k9 * x * (dt - k9 * 0.5 * x) - _
          t * f4 * (dt - d) + c * f6 * (dt - dp)
255: GoTo 265
260: mu = k8 * fc * bf * hf * (dt - hf / 2) + k8 * fc * _
          bw * (k9 * x - hf) * (dt - k9 * x / 2 - hf / 2) - _
          t * f4 * (dt - d) + c * f6 * (dt - dp)
265: 'MsgBox "Mult=", , mu
266: If d = 0 Or t = 0 Then d = dt
270: 'MsgBox "x/d=", , x / d
275: CalcMoment = Int(mu / 1000000)
280: GoTo 430
'''''''''''''''''
340: Rem stress in tension steel
345: e4 = 0.0035 * (d - x) / x
350: If Abs(e4) >= e1 Then f4 = fy / 1.15 * Sgn(e4)
355: If Abs(e4) < e1 Then f4 = e4 * 200000# * Sgn(e4)
360: Return
''''''''''''''''
370: Rem stress in prestressed steel
375: e5 = 0.0035 * (dt - x) / x + er
380: If e5 >= e3 Then f5 = pu / 1.15
385: If e5 < e3 And e5 > e2 Then f5 = 0.2 * pu * (0.02 + e5) / _
     (0.00575 + pu / 1000000#)
390: If e5 <= e2 Then f5 = 200000# * e5
395: Return
''''''''''''''''''
400: Rem stress in compression steel
403: If x < dp Then x = dp
405: e6 = 0.0035 * (x - dp) / x
410: If Abs(e6) >= e1 Then f6 = fy / 1.15 * Sgn(e6)
415: If Abs(e6) < e1 Then f6 = e6 * 200000 * Sgn(e6)
420: Return
'''''''''''''''''''
'''''''''''''''''''
430 Rem
End Function

