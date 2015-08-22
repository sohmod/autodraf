Attribute VB_Name = "Module13"
Option Explicit


''''FOR RASUK''''
Public Function DrawSectionDua(ByVal locx As Double, ByVal _
locy As Double, ByVal LinkDia As Double, _
ByVal b As Double, ByVal h As Double, ByVal SlabT As Double, _
ByVal SlabDropFront As Double, ByVal SlabDropBack As Double, _
ByVal Bar1No As Integer, ByVal Bar1Dia As Double, ByVal Bar1BM _
As Integer, ByVal Bar2No As Integer, ByVal Bar2Dia As Double, _
ByVal Bar2BM As Integer, ByVal Bar3No As Integer, ByVal Bar3Dia _
As Double, ByVal Bar3BM As Integer, ByVal Bar4No As Integer, _
ByVal Bar4Dia As Double, ByVal Bar4BM As Integer, ByVal Bar5No _
As Integer, ByVal Bar5Dia As Double, ByVal Bar5BM As Integer, _
ByVal Bar6No As Integer, ByVal Bar6Dia As Double, ByVal Bar6BM _
As Integer) As Object


Dim acadObj As Object

Dim PolyPt As Object

Dim pt(0 To 45) As Double
pt(0) = locx
pt(1) = locy - SlabDropFront
pt(2) = pt(0) + 1 * SlabT '''+ b
pt(3) = pt(1)
pt(4) = pt(2)
pt(5) = pt(3) + SlabDropFront
pt(6) = pt(4) + b
pt(7) = pt(5)
pt(8) = pt(6)
pt(9) = pt(7) - SlabDropBack

pt(10) = pt(8) + 1 * SlabT
pt(11) = pt(9)
pt(12) = pt(10)
pt(13) = pt(11) - SlabT / 3
pt(14) = pt(12) - SlabT / 3
pt(15) = pt(13)
pt(16) = pt(14) + 2 * SlabT / 3
pt(17) = pt(15) - SlabT / 3
pt(18) = pt(16) - SlabT / 3
pt(19) = pt(17)
pt(20) = pt(18)
pt(21) = pt(19) - SlabT / 3
pt(22) = pt(20) - 1 * SlabT
pt(23) = pt(21)
pt(24) = pt(22)
pt(25) = pt(23) - h + SlabDropBack + SlabT
pt(26) = pt(24) - b
pt(27) = pt(25)
pt(28) = pt(26)
pt(29) = pt(27) + h - SlabDropFront - SlabT
pt(30) = pt(28) - 1 * SlabT
pt(31) = pt(29)

pt(32) = pt(30)
pt(33) = pt(31) + SlabT / 3

pt(34) = pt(32) - SlabT / 3
pt(35) = pt(33)

pt(36) = pt(34) + 2 * SlabT / 3
pt(37) = pt(35) + SlabT / 3

pt(38) = pt(36) - SlabT / 3
pt(39) = pt(37)

pt(40) = pt(38)
pt(41) = pt(39) + SlabT / 3

pt(42) = pt(40) + 1 * SlabT
pt(43) = pt(41)

pt(44) = pt(42)
pt(45) = pt(43) + SlabDropFront



Set PolyPt = moSpace.AddLightWeightPolyline(pt)
'''PolyPt.Color = 120
PolyPt.layer = "BeamSection"
PolyPt.Update

'''''label section & location
Dim pt1(0 To 7) As Double
pt1(0) = locx + 200
pt1(1) = locy + 500
pt1(2) = pt1(0)
pt1(3) = pt1(1) - 7 * cVr
pt1(4) = pt1(2) - 3 * cVr
pt1(5) = pt1(3) + 2 * cVr
pt1(6) = pt1(4) + 3 * cVr
pt1(7) = pt1(5)
Set PolyPt = moSpace.AddLightWeightPolyline(pt1)
PolyPt.layer = "BeamSection"
PolyPt.Update


Dim corner1(0 To 2) As Double
Dim theight, Rotate As Double
Dim text As String
corner1(0) = pt1(4) - 7 * cVr
corner1(1) = pt1(5) - 2 * cVr
corner1(2) = 0
theight = 100
If i = 1 Then
       text = "A "
            End If
If i = 2 Then
       text = "B "
            End If
If i = 3 Then
       text = "C "
            End If
If i = 4 Then
       text = "D "
            End If
If i = 5 Then
       text = "E "
            End If
If i = 6 Then
       text = "F "
            End If
If i = 7 Then
       text = "G "
            End If
If i = 8 Then
       text = "H "
            End If
If i = 9 Then
       text = "I "
            End If



Set acadObj = moSpace.AddText(text, corner1, theight)
acadObj.Color = 2
acadObj.layer = "BeamSection"
acadObj.Update

pt1(0) = locx + 200
pt1(1) = locy + 800 + beamH(i) + 300
pt1(2) = pt1(0)
pt1(3) = pt1(1) + 7 * cVr
pt1(4) = pt1(2) - 3 * cVr
pt1(5) = pt1(3) - 2 * cVr
pt1(6) = pt1(4) + 3 * cVr
pt1(7) = pt1(5)
Set PolyPt = moSpace.AddLightWeightPolyline(pt1)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
PolyPt.layer = "BeamSection"
PolyPt.Update
corner1(0) = pt1(4) - 7 * cVr
corner1(1) = pt1(5) - 2 * cVr
corner1(2) = 0
theight = 100
Set acadObj = moSpace.AddText(text, corner1, theight)
acadObj.Color = 2
acadObj.layer = "BeamSection"
acadObj.Update

''''section dim.
corner1(0) = pt(0) - 6.5 * cVr
corner1(1) = pt(5) - 0.65 * h
corner1(2) = 0
theight = 50 ''fix
text = Str(h)
  Set acadObj = moSpace.AddText(text, corner1, theight)
  Rotate = 1.57  '' 90 degrees
  Call acadObj.Rotate(corner1, Rotate)
  acadObj.layer = "BeamDimension"
  acadObj.Update
Call ArrowVertical(pt(0) - 5 * cVr, pt(5), -h, -30, -15, "BeamDimension")
Call ArrowHorizontal(pt(0) - 6 * cVr, pt(5), 2 * cVr, 0, 0, "BeamDimension")
Call ArrowHorizontal(pt(0) - 6 * cVr, pt(5) - h, 2 * cVr, 0, 0, "BeamDimension")

corner1(0) = pt(26) + b / 2 - Val(Len(Str(b))) * FontSz / 2
corner1(1) = pt(27) - 16 * cVr
corner1(2) = 0
theight = 50  ''fix
text = Str(b)
Set acadObj = moSpace.AddText(text, corner1, theight)
  acadObj.layer = "BeamDimension"
  acadObj.Update
Call ArrowHorizontal(pt(26), pt(27) - 13 * cVr, b, 30, 15, "BeamDimension")

Call ArrowVertical(pt(26), pt(27) - 12 * cVr, -2 * cVr, 0, 0, "BeamDimension")

Call ArrowVertical(pt(26) + b, pt(27) - 12 * cVr, -2 * cVr, 0, 0, "BeamDimension")

'''''detail link
'''Call DrawLinkSectionDua(locx + SlabT + cVr + 2.5 * LinkDia, _
locy - cVr - LinkDia / 2, LinkDia, b, h, SlabT, _
Bar1No, Bar1Dia, Bar1BM, _
Bar2No, Bar2Dia, Bar2BM, _
Bar3No, Bar3Dia, Bar3BM, _
Bar4No, Bar4Dia, Bar4BM, _
Bar5No, Bar5Dia, Bar5BM, _
Bar6No, Bar6Dia, Bar6BM)

End Function




