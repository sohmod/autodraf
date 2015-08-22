VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcCrackWidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CalcCrackWidth( _
ByVal Location As String, _
ByVal ms As Double, _
ByVal fc As Double, _
ByVal fy As Double, _
ByVal rb As Double, _
ByVal N As Double, _
ByVal ac As Double, _
ByVal at As Double, _
ByVal b As Double, _
ByVal h As Double, _
ByVal cs As Double, _
ByVal cb As Double) As Double


Dim CrkBotUnderBar, CrkBotMiddleBar As Double
Dim CrkCorner1, CrkCorner2, CrkCorner As Double
Dim CrkSideBesideBar, CrkSideAt125Bar As Double
Dim CrkSideAt250Bar, CrkSideAtH2Bar As Double

Dim Pi, mMe, Ro, Rp, d, dp, x, xD, Ic, Sp As Double
Dim kb, kc, ks1, ks2, ks3, ks4 As Double
Dim ec, CrackMax, CrackMin As Double
Dim Pbagi As Integer

Pi = 3.141592654

''''''110 at = Int(Pi * rb ^ 2 * N / 4)
''''''112 ac = ac
114 d = h - cb - rb / 2    ''''estimately
116 dp = cb + rb / 2       ''''
118 Ro = at / b / d
120 Rp = ac / b / d
122 ec = 24 + (fc - 20) / 5

124 mMe = 200 / ec * 2
126 xD = (mMe * Ro) ^ 2 + 2 * Ro * (mMe - 1) * Rp + _
         ((mMe - 1) * Rp) ^ 2 + 2 * mMe * Ro + 2 * dp * (mMe - 1) _
         * Rp / d
128 xD = Sqr(xD) - mMe * Ro - (mMe - 1) * Rp
130 x = xD * d
132 Ic = xD ^ 3 / 3 + mMe * Ro * (1 - xD) ^ 2 + (mMe - 1) _
        * Rp * (xD - dp / d) ^ 2
134 Ic = Ic * b * d ^ 3
If N = 1 Then
   Pbagi = 1
   Else
   Pbagi = N - 1
   End If

136 Sp = (b - 2 * cs - N * rb) / Pbagi
138 kb = Sqr((Sp + rb) ^ 2 / 4 + (cb + rb / 2) ^ 2) - rb / 2
140 kc = Sqr((cs + rb / 2) ^ 2 + (cb + rb / 2) ^ 2) - rb / 2
142 ks1 = Sqr((cs + rb / 2) ^ 2 + (d / 2 - x / 2) ^ 2) - rb / 2
144 ks2 = Sqr((cs + rb / 2) ^ 2 + (d - x - 250) ^ 2) - rb / 2
145 ks3 = Sqr((cs + rb / 2) ^ 2 + (d - x - 125) ^ 2) - rb / 2
146 ks4 = cs
 


255 ''ac = cb: a1 = h: c = cb: x1 = h - X: p$ = "BotUnderBar":
CrkBotUnderBar = CrackWidth(ms, cb, h, cb, h - x, "BotUnderBar", _
b, h, x, at, ec, Ic)
CalcCrackWidth = CrkBotUnderBar

257 ''ac = kb: a1 = h: c = cb: x1 = h - X: p$ = "BotMiddleBar":
CrkBotMiddleBar = CrackWidth(ms, kb, h, cb, h - x, "BotMiddleBar", _
b, h, x, at, ec, Ic)
If CalcCrackWidth < CrkBotMiddleBar Then
    CalcCrackWidth = CrkBotMiddleBar
         End If
     
260 '' ac = kc: a1 = h: c = cb: x1 = h - X: p$ = "Corner"
CrkCorner1 = CrackWidth(ms, kc, h, cb, h - x, "Corner", _
b, h, x, at, ec, Ic)
CrkCorner2 = CrackWidth(ms, kc, h, cs, h - x, "Corner", _
b, h, x, at, ec, Ic)
 If CrkCorner1 >= CrkCorner2 Then
    CrkCorner = CrkCorner1
       Else
         CrkCorner = CrkCorner2
           End If
 If CalcCrackWidth < CrkCorner Then
     CalcCrackWidth = CrkCorner
        End If
280 ''ac = ks4: a1 = D: c = cs: x1 = D - X: p$ = "SideBesideBar":
CrkSideBesideBar = CrackWidth(ms, ks4, d, cs, d - x, "SideBesideBar", _
b, h, x, at, ec, Ic)
  If CalcCrackWidth < CrkSideBesideBar Then
      CalcCrackWidth = CrkSideBesideBar
         End If
290 ''ac = ks3: a1 = X + 125: c = cs: x1 = 125: p$ = "SideAt125Bar":
CrkSideAt125Bar = CrackWidth(ms, ks3, x + 125, cs, 125, "SideAt125Bar", _
b, h, x, at, ec, Ic)
  If CalcCrackWidth < CrkSideAt125Bar Then
       CalcCrackWidth = CrkSideAt125Bar
           End If
300 ''ac = ks2: a1 = X + 250: c = cs: x1 = 250: p$ = "SideAt250Bar":
CrkSideAt250Bar = CrackWidth(ms, ks2, x + 250, cs, 250, "SideAt250Bar", _
b, h, x, at, ec, Ic)
  If CalcCrackWidth < CrkSideAt250Bar Then
       CalcCrackWidth = CrkSideAt250Bar
           End If
310 ''ac = ks1: a1 = (D + X) / 2: c = cs: x1 = (D - X) / 2: _
    p$ = "SideAtH2Bar": Return
CrkSideAtH2Bar = CrackWidth(ms, ks1, (d + x) / 2, cs, (d - x) / 2, _
"SideAtH2Bar", b, h, x, at, ec, Ic)
  If CalcCrackWidth < CrkSideAtH2Bar Then
       CalcCrackWidth = CrkSideAtH2Bar
           End If
           
If Location = "Maximum" Then
     CalcCrackWidth = CalcCrackWidth
        End If
If Location = "BotUnderBar" Then
     CalcCrackWidth = CrkBotUnderBar
        End If
If Location = "BotMiddleBar" Then
     CalcCrackWidth = CrkBotMiddleBar
        End If
If Location = "Corner" Then
     CalcCrackWidth = CrkCorner
        End If
If Location = "SideBesideBar" Then
     CalcCrackWidth = CrkSideBesideBar
        End If
If Location = "SideAt125Bar" Then
     CalcCrackWidth = CrkSideAt125Bar
        End If
If Location = "SideAt250Bar" Then
     CalcCrackWidth = CrkSideAt250Bar
        End If
If Location = "SideAtH2Bar" Then
     CalcCrackWidth = CrkSideAtH2Bar
        End If
        
 End Function

       Private Function CrackWidth(ByVal Mntser As Double, _
         ByVal ac As Double, ByVal a1 As Double, _
         ByVal c As Double, ByVal x1 As Double, _
         pLoc As String, _
         ByVal b As Double, _
         ByVal h As Double, _
         ByVal x As Double, _
         ByVal at As Double, _
         ByVal ec As Double, _
         ByVal Ic As Double) As Double
     Dim cw, e1, em As Double
     
152   ''  On z GoSub 250, 260, 270
154   ''    If a1 > h Then
      ''      msg "  not critical", , " "
      ''        End If
156         e1 = Mntser * 1000# * x1 * 2 / ec / Ic
158         em = b * (h - x) * (a1 - x) / (3 * 200 * 1000# * at) / (d - x)
160         em = e1 - em
162           cw = 1 + 2 * (ac - c) / (h - x)
164           cw = 3 * ac * em / cw
166           cw = Int(cw * 10000#) / 10000#
      CrackWidth = cw
        End Function


 
