Attribute VB_Name = "Module18"
Option Explicit

Public Function RequiredAst( _
ByVal fc As Double, _
ByVal fy As Double, _
ByVal bardia As Double, _
ByVal cvr As Double, _
ByVal h As Double, _
ByVal barmark As Integer, _
ByVal md As Double, _
ByVal m As Double) As Double

Dim b, bf, bw, hf As Double
Dim k, kp, z, d, X, mf, mw, kf, f2, dp, c, t As Double
b = 1000
bf = 0
bw = 1000
hf = 0
d = h - cvr - bardia / 2
dp = cvr + bardia / 2

225 k = 0.402 * (md - 0.4) - 0.18 * (md - 0.4) ^ 2 ''' md=1 >1% ; md=0.9 >10% ; md=0.7 >30%
230 b = bf
235 If bf = 0 Or bf <= bw Then b = bw
240 If bf = 0 Then hf = 0
245 If hf = 0 Then bf = 0
250 kp = m * 10 ^ 6 / b / d ^ 2 / fc
255 If kp > k Then GoTo 335
260 ''''singly reinforced'''''
265 z = d * (0.5 + (0.25 - kp / 0.9) ^ 0.5)
270 If z > 0.95 * d Then z = 0.95 * d
275 X = (d - z) / 0.45
280 If bf = 0 Or hf = 0 Then GoTo 320
285 If 0.9 * X <= hf Then GoTo 320
290 mf = 0.45 * fc * (b - bw) * hf * (d - hf / 2)
295 kf = (m * 10 ^ 6 - mf) / (fc * bw * d ^ 2)
300 If kf > k Then GoTo 335
305 ''' print singly T 0.9x>hf
310 z = d * (0.5 + (0.25 - kf / 0.9) ^ 0.5): X = (d - z) / 0.45
315 t = mf / (0.87 * fy * (d - hf / 2)) + (m * 10 ^ 6 - mf) / (0.87 * fy * z): c = 0.0024 * bw * h: GoTo 330
320 ''if bw>=bf then print "singly rc rect." else " singly rc t 0.9x<hf"
325 t = m * 10 ^ 6 / (0.87 * fy * z): c = 0.0024 * bw * h
330 GoTo 440
335 '''''doubly reinforced'''''
340 z = d * (0.5 + (0.25 - k / 0.9) ^ 0.5)
345 X = (d - z) / 0.45
350 f2 = 0.87 * fy
355 If dp / X > 1 - fy / 800 Then f2 = 700 * (1 - dp / X)
360 If bf = 0 Or hf = 0 Or bf = bw Then GoTo 425
365 mf = 0.45 * fc * (b - bw) * bf * (d - hf / 2)
370 mw = k * fc * bw * d ^ 2
375 kf = (m * 10 ^ 6 - mf) / (fc * bw * d ^ 2)
380 If kf >= k Then GoTo 400
385 '''print singly rc t''''
390 t = mf / (0.87 * fy * (d - hf / 2)) + (m * 10 ^ 6 - mf) / (0.87 * fy * z): c = 0.0024 * bw * h
395 GoTo 440
400 ''''print doubly rc t  x>hf'''
405 c = (m * 10 ^ 6 - mf - mw) / (0.87 * f2 * (d - dp))
410 z = d * (0.5 + (0.25 - k / 0.9) ^ 0.5): X = (d - z) / 0.45
415 t = (0.45 * fc * (b - bw) * hf + 0.45 * fc * bw * 0.9 * X + 0.87 * f2 * c) / (0.87 * fy)
420 GoTo 440
425 '''print doubly rc rect.'''
430 c = (kp - k) * fc * b * d ^ 2 / (f2 * (d - dp))
435 t = k * fc * b * d ^ 2 / (0.87 * fy * z) + c
440 'GoSub 700   '''''>>>>>>>>>>min steel
445 'GoSub 660   '''''>>>>>>>>>>>print result
RequiredAst = t




End Function


