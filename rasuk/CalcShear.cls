VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcShear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CalcShear( _
ByVal Link As Double, _
ByVal Cover As Double, _
ByVal sv As Double, _
ByVal asv As Double, _
ByVal asvt As Double, _
ByVal atr As Double, _
ByVal fcu As Double, _
ByVal fy As Double, _
ByVal fyv As Double, _
ByVal bw As Double, _
ByVal ht As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal ast As Double, _
ByVal asc As Double) As Integer

750 Rem shear capacity and torsion moment
753 Dim x11, y11, j1, j2, j3, trc, vs, v1, v2, vc, vct As Double
754 If sv = 0 Or asv = 0 Then GoTo 805  ''' for nil link rebar ie pad
755 x11 = bw - 2 * Cover - Link
760 y11 = ht - 2 * Cover - Link
762 atr = atr + asc
765 j1 = atr * fy / (fyv * (x11 + y11))
770 j2 = asvt / sv
775 ''rem If Abs(j1 - j2) < 0.5 Then MsgBox " balanced torsion linkage"
780 j3 = j1
785 If j1 >= j2 Then j3 = j2
790 trc = j3 * 0.8 * x11 * y11 * 0.87 * fyv
795 ''If Abs(j1 - j2) < 0.5 Then
    ''MsgBox " Balanced linkage: Mtor = " & _
    ''Str(j3) & "  " & Str(x11) & "  " & Str(y11) & "  " & Str(fyv), , "Torsional Capacity."
    ''End If
800 vs = asv * fyv * d / sv
805 v1 = 100 * ast / bw / d
810 If v1 > 3 Then v1 = 3
815 v2 = 400 / d
820 If v2 < 1 Then v2 = 1
825 vc = 0.79 * v1 ^ (1 / 3) * v2 ^ (1 / 4) / 1.25
826 If fcu > 40 Then fcu = 40
828 vc = vc * (fcu / 25) ^ (1 / 3)
830 vct = vs + vc * bw * d
835 'MsgBox vc & " -- " & vs, , "vc" & "vs"
840 CalcShear = Int(vct / 1000)

End Function

