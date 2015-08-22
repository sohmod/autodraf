Attribute VB_Name = "Module8"
Option Explicit


''''FOR RASUK & COLUMN''''
Public Function DrawLinkSection(ByVal locx As Double, ByVal _
locy As Double, ByVal LinkDia As Double, _
ByVal b As Double, ByVal h As Double, ByVal SlabT As Double, _
ByVal Bar1No As Integer, ByVal Bar1Dia As Double, ByVal Bar1BM _
As Integer, ByVal Bar2No As Integer, ByVal Bar2Dia As Double, _
ByVal Bar2BM As Integer, ByVal Bar3No As Integer, ByVal Bar3Dia _
As Double, ByVal Bar3BM As Integer, ByVal Bar4No As Integer, _
ByVal Bar4Dia As Double, ByVal Bar4BM As Integer, ByVal Bar5No _
As Integer, ByVal Bar5Dia As Double, ByVal Bar5BM As Integer, _
ByVal Bar6No As Integer, ByVal Bar6Dia As Double, ByVal Bar6BM _
As Integer) As Object


Dim PolyPt As Object

Dim pt0(0 To 65) As Double
Dim CoefR As Double

CoefR = 2 * LinkDia
pt0(0) = locx
pt0(1) = locy
pt0(2) = pt0(0) + b - 2 * cVr - LinkDia - 2 * CoefR ''<<<<
pt0(3) = pt0(1)

  pt0(4) = pt0(2) + CoefR * Cv1
  pt0(5) = pt0(3) - CoefR * Cv5
  pt0(6) = pt0(4) + CoefR * Cv2
  pt0(7) = pt0(5) - CoefR * Cv4
  pt0(8) = pt0(6) + CoefR * Cv3
  pt0(9) = pt0(7) - CoefR * Cv3
  pt0(10) = pt0(8) + CoefR * Cv4
  pt0(11) = pt0(9) - CoefR * Cv2
  pt0(12) = pt0(10) + CoefR * Cv5
  pt0(13) = pt0(11) - CoefR * Cv1
  pt0(14) = pt0(12)
  pt0(15) = pt0(13) - h + 2 * cVr + LinkDia + 2 * CoefR ''<<<<
    
  pt0(16) = pt0(14) - CoefR * Cv5
  pt0(17) = pt0(15) - CoefR * Cv1
  pt0(18) = pt0(16) - CoefR * Cv4
  pt0(19) = pt0(17) - CoefR * Cv2
  pt0(20) = pt0(18) - CoefR * Cv3
  pt0(21) = pt0(19) - CoefR * Cv3
  pt0(22) = pt0(20) - CoefR * Cv2
  pt0(23) = pt0(21) - CoefR * Cv4
  pt0(24) = pt0(22) - CoefR * Cv1
  pt0(25) = pt0(23) - CoefR * Cv5
  pt0(26) = pt0(24) - b + 2 * cVr + LinkDia + 2 * CoefR ''<<<<
  pt0(27) = pt0(25)
   
  pt0(28) = pt0(26) - CoefR * Cv1
  pt0(29) = pt0(27) + CoefR * Cv5
  pt0(30) = pt0(28) - CoefR * Cv2
  pt0(31) = pt0(29) + CoefR * Cv4
  pt0(32) = pt0(30) - CoefR * Cv3
  pt0(33) = pt0(31) + CoefR * Cv3
  pt0(34) = pt0(32) - CoefR * Cv4
  pt0(35) = pt0(33) + CoefR * Cv2
  pt0(36) = pt0(34) - CoefR * Cv5
  pt0(37) = pt0(35) + CoefR * Cv1
  pt0(38) = pt0(36)
  pt0(39) = pt0(37) + h - 2 * cVr - LinkDia - 2 * CoefR ''<<<<
    
  pt0(40) = pt0(38) + CoefR * Cv5
  pt0(41) = pt0(39) + CoefR * Cv1
  pt0(42) = pt0(40) + CoefR * Cv4
  pt0(43) = pt0(41) + CoefR * Cv2
  pt0(44) = pt0(42) + CoefR * Cv3
  pt0(45) = pt0(43) + CoefR * Cv3
  pt0(46) = pt0(44) + CoefR * Cv2
  pt0(47) = pt0(45) + CoefR * Cv4
  pt0(48) = pt0(46) + CoefR * Cv1
  pt0(49) = pt0(47) + CoefR * Cv5
  pt0(50) = pt0(48) + 3 * LinkDia
  pt0(51) = pt0(49) - 0.25 * LinkDia
''''''''''''''''''''
''''''''''''''''''''
  pt0(52) = pt0(48)
  pt0(53) = pt0(49)
  pt0(54) = pt0(46)
  pt0(55) = pt0(47)
  pt0(56) = pt0(44)
  pt0(57) = pt0(45)
  pt0(58) = pt0(42)
  pt0(59) = pt0(43)
  pt0(60) = pt0(40)
  pt0(61) = pt0(41)
  pt0(62) = pt0(38)
  pt0(63) = pt0(39)
  pt0(64) = pt0(62) + 0.25 * LinkDia
  pt0(65) = pt0(63) - 3 * LinkDia
  


Set PolyPt = moSpace.AddLightWeightPolyline(pt0)
PolyPt.Color = 2
PolyPt.layer = "BeamSection"
PolyPt.Update

'''Call CircleBarMarkDua(locx, locy, LinkDia, b, h, SlabT, _
Bar1No, Bar1Dia, Bar1BM, _
Bar2No, Bar2Dia, Bar2BM, _
Bar3No, Bar3Dia, Bar3BM, _
Bar4No, Bar4Dia, Bar4BM, _
Bar5No, Bar5Dia, Bar5BM, _
Bar6No, Bar6Dia, Bar6BM)

End Function





