Attribute VB_Name = "module1"
Option Explicit
Public BarMark As Integer
Public cVr As Double
Public sysVarName As String
Public sysVarData As Variant
Public sdiMode As Integer
Public osMode As Integer
Public Bar1BM, Bar2BM, Bar3BM, Bar4BM, Bar5BM, Bar6BM As Integer
Public Bar6No, Bar6Dia As Double
Public SWdthLft, SWdthRght As Double
Public Cv1, Cv2, Cv3, Cv4, Cv5 As Double
Public DetailType As String
Public FontSz As Integer
Public TxtCounter As Integer

Option Base 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''FOR RASUK''''
Public NoOfSpan As Integer
''Public barMark As Integer
Public NamaRasuk As String
Public fcu, fy, fyv, Shrink, Creep As Double
Public GridNameL(1 To 10), GridNameR(1 To 10) As String

Public scL(1 To 10), sbL(1 To 10), shL(1 To 10) As Double
Public beamL(1 To 10), beamB(1 To 10), beamH(1 To 10), _
    slabDrop(1 To 10), beamUplift(1 To 10) As Double
Public shR(1 To 10), sbR(1 To 10), scR(1 To 10) As Double
Public FrontSlabLvl(1 To 10), BackSlabLvl(1 To 10) As Double


Public RbarTL1no(1 To 10), RbarTL1dia(1 To 10), RbarTL1curE(1 To 10) As Double
Public RbarTL2no(1 To 10), RbarTL2dia(1 To 10), RbarTL2curE(1 To 10) As Double
Public RbarCLfno(1 To 10), RbarCLfdia(1 To 10), RbarCLfcurE(1 To 10) As Double

Public RbarMS1no(1 To 10), RbarMS1dia(1 To 10) As Double
Public RbarMS2curS(1 To 10), RbarMS2no(1 To 10), RbarMS2dia(1 To 10), RbarMS2curE(1 To 10) As Double

Public RbarTR1no(1 To 10), RbarTR1dia(1 To 10), RbarTR1curS(1 To 10) As Double
Public RbarTR2no(1 To 10), RbarTR2dia(1 To 10), RbarTR2curS(1 To 10) As Double
Public RbarCRtno(1 To 10), RbarCRtdia(1 To 10), RbarCRtcurS(1 To 10) As Double

Public RbarLCno(1 To 10), RbarLCdia(1 To 10) As Double

Public LinkDia(1 To 10), LinkLSpace(1 To 10) As Double
Public LinkMSpace(1 To 10), LinkRSpace(1 To 10) As Double

Public i As Integer
Public pbx(0 To 8 * 10) As Double
Public pbt(0 To 8 * 10) As Double
Public pbs(0 To 8 * 10) As Double
Public pbb(0 To 8 * 10) As Double

Public slabThick, gAp, stirupD As Double    'cVr
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Xinsertion, Yinsertion As Double
Public Components As Integer
Public DataFile As String
Public dwgName As String
Public acadApp As Object
Public acadDoc As Object
Public moSpace As Object
Public paSpace As Object


'''''''''''FOR RASUK''''
Public Function DrawLink(ByVal locx As Double, ByVal locy _
As Double, ByVal b As Double, ByVal h As Double, _
ByVal LinkDiam As Double, _
ByVal LLeft As Double, ByVal LeftSpace As Double, _
ByVal LMid As Double, ByVal MidSpace As Double, _
ByVal LRight As Double, ByVal RightSpace As Double) As Object


Dim link0(0 To 3) As Double
BarMark = BarMark + 1
link0(0) = locx + 2 * cVr  ''''locx + 2 * cVr
link0(1) = locy
link0(2) = link0(0)
link0(3) = locy + h - LinkDiam - 2 * cVr
Dim polyLink01 As Object
Set polyLink01 = moSpace.AddLightWeightPolyline(link0)
polyLink01.layer = "RebarLink"
polyLink01.Update
Dim link1(0 To 3) As Double
link1(0) = locx + LLeft '''link0(0) + LLeft - 3 * cVr    ''''3 * cVr
link1(1) = link0(1)
link1(2) = link1(0)
link1(3) = link0(3)
Set polyLink01 = moSpace.AddLightWeightPolyline(link1)
polyLink01.layer = "RebarLink"
polyLink01.Update

Dim link2(0 To 3) As Double
link2(0) = locx + LLeft + LeftSpace  ''link1(0) + LeftSpace + 1 * cVr
link2(1) = link1(1)
link2(2) = link2(0)
link2(3) = link1(3)
Set polyLink01 = moSpace.AddLightWeightPolyline(link2)
polyLink01.layer = "RebarLink"
polyLink01.Update
Dim link3(0 To 3) As Double
link3(0) = locx + LLeft + LMid - RightSpace '''link2(0) + LMid - LeftSpace - RightSpace + 1 * cVr  ''4
link3(1) = link2(1)
link3(2) = link3(0)
link3(3) = link2(3)
Set polyLink01 = moSpace.AddLightWeightPolyline(link3)
polyLink01.layer = "RebarLink"
polyLink01.Update

Dim link4(0 To 3) As Double
link4(0) = locx + LLeft + LMid   '''link3(0) + RightSpace + 1 * cVr
link4(1) = link3(1)
link4(2) = link4(0)
link4(3) = link3(3)
Set polyLink01 = moSpace.AddLightWeightPolyline(link4)
polyLink01.layer = "RebarLink"
polyLink01.Update
Dim link5(0 To 3) As Double
link5(0) = locx + LLeft + LMid + LRight - 2 * cVr '''link4(0) + LRight - 6 * cVr
link5(1) = link4(1)
link5(2) = link5(0)
link5(3) = link4(3)
Set polyLink01 = moSpace.AddLightWeightPolyline(link5)
polyLink01.layer = "RebarLink"
polyLink01.Update

Dim XL0, YL0, LL0, XM1, YM1, LM1, XR2, YR2, LR2 As Double
XL0 = link0(0)
YL0 = locy + h / 2.8
LL0 = link1(0) - link0(0)
XM1 = link2(0)
YM1 = locy + h / 2.8
LM1 = link3(0) - link2(0)
XR2 = link4(0)
YR2 = locy + h / 2.8
LR2 = link5(0) - link4(0)
Call ArrowHorizontal(XL0, YL0, LL0, 60, 20, "LabelRebarLink")
Call ArrowHorizontal(XM1, YM1, LM1, 60, 20, "LabelRebarLink")
Call ArrowHorizontal(XR2, YR2, LR2, 60, 20, "LabelRebarLink")
XL0 = XL0 + LL0 / 2 - 5 * FontSz
YL0 = YL0 + 2 * LinkDiam
XM1 = XM1 + LM1 / 2 - 7 * FontSz
YM1 = YM1 + 2 * LinkDiam
XR2 = XR2 + LR2 / 2 - 6 * FontSz
YR2 = YR2 + 2 * LinkDiam
Call LabelRbar(XL0, YL0, Int(LL0 / LeftSpace + 1), LinkDiam, BarMark, LeftSpace, 0, FontSz, "R", "LabelRebarLink")
Call LabelRbar(XM1, YM1, Int(LM1 / MidSpace + 1), LinkDiam, BarMark, MidSpace, 0, FontSz, "R", "LabelRebarLink")
Call LabelRbar(XR2, YR2, Int(LR2 / RightSpace + 1), LinkDiam, BarMark, RightSpace, 0, FontSz, "R", "LabelRebarLink")

End Function

