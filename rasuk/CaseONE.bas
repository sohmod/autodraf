Attribute VB_Name = "Module11"
Option Explicit
Public Function CaseONE()
Dim xStat, yStat As Double
Dim caseI As Integer
Dim pt0(0 To 31) As Double
Dim pt0a(0 To 15) As Double
Dim pt1(0 To 7) As Double
Dim pt2(0 To 11) As Double
Dim TempX As Double
Dim polypt0 As Object

Dim polypt0a As Object

Dim Polypt1 As Object

Dim Polypt2 As Object



           If RbarTL1no(i) <> 0 Or RbarTL1dia(i) <> 0 Or RbarTL1curE(i) <> 0 Then
           pt2(0) = pbx(8 * i - 12) - RbarTR1curS(i - 1) + sbR(i - 1)
           pt2(1) = pbt(8 * i - 11) - cVr - LinkDia(i - 1) - RbarLCdia(i - 1) _
                    - RbarTL1dia(i) / 2 - 10
           pt2(2) = pt2(0) + 40
           pt2(3) = pt2(1) + 10
           pt2(4) = pt2(2) + RbarTR1curS(i - 1) - SWdthLft - 40
           pt2(5) = pt2(3)
           pt2(6) = pbx(8 * i - 6) + SWdthLft - sbL(i) ''pt2(4) + SWdthLft - sbR(i)
           pt2(7) = pbt(8 * i - 5) - cVr - LinkDia(i) - RbarLCdia(i) - _
                    RbarTL1dia(i) / 2
           pt2(8) = pt2(6) + RbarTL1curE(i) - SWdthLft - 40  ''pbx(8 * i - 6) + SWdthLft - sbL(i)
           pt2(9) = pt2(7)
           pt2(10) = pt2(8) + 40
           pt2(11) = pt2(9) - 10
             Set Polypt2 = moSpace.AddLightWeightPolyline(pt2)
             Polypt2.layer = "RebarSupt"
             Polypt2.thickness = RbarTR1dia(i) / 2
             Polypt2.Update
           BarMark = BarMark + 1
           Bar1BM = BarMark
           Call ArrowTail(pt2(0) + RbarTR1curS(i - 1) / 2, pt2(3), 400, 30, _
                          15, 150, "LabelRebarSupt")
           Call LabelRbar(pt2(0) + RbarTR1curS(i - 1) / 2, pt2(3) + _
           430 + FontSz / 2, RbarTL1no(i), RbarTL1dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSupt")
           Call LabelDimension(pt2(0) - 80, pt2(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt2(10) - 20, pt2(11) - 30, BarMark, _
                          30, "Curtailment")
           End If
           ''''''''''''''''''''''''''''
           '''' Bar2BM = barMark
           If RbarTL2no(i) <> 0 Or RbarTL2dia(i) <> 0 Or RbarTL2curE(i) <> 0 Then
           pt2(0) = pbx(8 * i - 12) - RbarTR2curS(i - 1) + sbR(i - 1)
           pt2(1) = pbt(8 * i - 11) - cVr - LinkDia(i - 1) - RbarLCdia(i - 1) - _
                    RbarTL1dia(i) - 20 - RbarTL2dia(i) / 2 - 10
           pt2(2) = pt2(0) + 40
           pt2(3) = pt2(1) + 10
           pt2(4) = pt2(2) + RbarTR2curS(i - 1) - SWdthLft - 40
           pt2(5) = pt2(3)
           pt2(6) = pbx(8 * i - 6) + SWdthLft - sbL(i)
           pt2(7) = pbt(8 * i - 5) - cVr - LinkDia(i) - RbarLCdia(i) - _
                    RbarTL1dia(i) - 20 - RbarTL2dia(i) / 2
           pt2(8) = pt2(6) + RbarTL2curE(i) - SWdthLft - 40
           pt2(9) = pt2(7)
           pt2(10) = pt2(8) + 40
           pt2(11) = pt2(9) - 10
             Set Polypt2 = moSpace.AddLightWeightPolyline(pt2)
             Polypt2.layer = "RebarSupt"
             Polypt2.thickness = RbarTL2dia(i) / 2
             Polypt2.Update
           BarMark = BarMark + 1
           Bar2BM = BarMark
           Call ArrowTail(pt2(10) - RbarTL2curE(i) / 2, pt2(9), 200, 30, _
                          15, 150, "LabelRebarSupt")
           Call LabelRbar(pt2(10) - RbarTL2curE(i) / 2, pt2(9) + 230 + FontSz / 2, _
                          RbarTL2no(i), RbarTL2dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSupt")
           Call LabelDimension(pt2(0) - 80, pt2(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt2(10) - 20, pt2(11) - 30, BarMark, _
                          30, "Curtailment")
           End If
           ''''''''''''''''''''''''''''
           ''LHS BOTTOM SUPT. BAR --DEFAULT VALUE.
           pt2(0) = pbx(8 * i - 12) - RbarCRtcurS(i - 1) + sbR(i - 1)
           pt2(1) = pbb(8 * i - 11) + cVr + LinkDia(i - 1) + RbarMS1dia(i - 1) + _
                    RbarCRtdia(i - 1) / 2 + 10
           pt2(2) = pt2(0) + 40
           pt2(3) = pt2(1) - 10
           pt2(4) = pbx(8 * i - 12) - SWdthLft
           pt2(5) = pt2(3)
           pt2(6) = pbx(8 * i - 6) + SWdthLft - sbL(i)
           pt2(7) = pbb(8 * i - 5) + cVr + LinkDia(i) + RbarMS1dia(i) + _
                    RbarCLfdia(i) / 2
           pt2(8) = pt2(6) + RbarCLfcurE(i) - 40
           pt2(9) = pt2(7)
           pt2(10) = pt2(8) + 40
           pt2(11) = pt2(9) + 10
             Set Polypt2 = moSpace.AddLightWeightPolyline(pt2)
             Polypt2.layer = "RebarSupt"
             Polypt2.thickness = RbarCLfdia(i) / 2
             Polypt2.Update
           BarMark = BarMark + 1
           Bar6BM = BarMark
           Call ArrowTail(pt2(2) + RbarCRtcurS(i - 1) / 2, pt2(3), -400, -30, -15, _
                           150, "LabelRebarSupt")
           Call LabelRbar(pt2(2) + RbarCRtcurS(i - 1) / 2, pt2(3) - 430 - 2 * FontSz, _
                          RbarCLfno(i), RbarCLfdia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSupt")
           Call LabelDimension(pt2(0) - 80, pt2(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt2(10) - 20, pt2(11), BarMark, _
                          30, "Curtailment")
           '''''''''''''''''''''''''''
           ''LINK CARRIER -- DEFAULT VALUE.
           pt1(0) = pbx(8 * i - 6) + cVr + SWdthLft - sbL(i)
           pt1(1) = pbt(8 * i - 5) - cVr - LinkDia(i) - RbarLCdia(i) / 2 - 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) + 10
           pt1(4) = pbx(8 * i - 4) - cVr - SWdthRght + sbR(i) - 40
           pt1(5) = pt1(3)
           pt1(6) = pt1(4) + 40
           pt1(7) = pt1(5) - 10
             Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
             Polypt1.layer = "RebarSpan"
             Polypt1.thickness = RbarLCdia(i) / 2
             Polypt1.Update
           BarMark = BarMark + 1
           Bar3BM = BarMark
           Call ArrowTail(pt1(0) + 0.45 * beamL(i), pt1(3), 250, 30, _
                          15, 150, "LabelRebarSpan")
           Call LabelRbar(pt1(0) + 0.45 * beamL(i), pt1(3) + 280 + FontSz / 2, _
                          RbarLCno(i), RbarLCdia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
           Call LabelDimension(pt1(0) - 80, pt1(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7) - 30, BarMark, _
                          30, "Curtailment")
           ''''''''''''''''''''''''''''''
           ''BOTTOM MAIN BAR.
           If RbarMS1no(i) <> 0 Or RbarMS1dia(i) <> 0 Then
           pt1(0) = pbx(8 * i - 6) + cVr + SWdthLft - sbL(i)
           pt1(1) = pbb(8 * i - 5) + cVr + LinkDia(i) + RbarMS1dia(i) / 2 + 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) - 10
           pt1(4) = pbx(8 * i - 4) - cVr - SWdthRght + sbR(i) - 40
           pt1(5) = pt1(3)
           pt1(6) = pt1(4) + 40
           pt1(7) = pt1(5) + 10
             Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
             Polypt1.layer = "RebarSpan"
             Polypt1.thickness = RbarMS1dia(i) / 2
             Polypt1.Update
           BarMark = BarMark + 1
           Bar4BM = BarMark
           Call ArrowTail(pt1(0) + RbarTL1curE(i), pt1(3), -250, -30, _
                          -15, 150, "LabelRebarSpan")
           Call LabelRbar(pt1(0) + RbarTL1curE(i), pt1(3) - 280 - 2 * FontSz, _
                          RbarMS1no(i), RbarMS1dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
           Call LabelDimension(pt1(0) - 80, pt1(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7), BarMark, _
                          30, "Curtailment")
           End If
           ''''''''''''''''''''''''''''
           If RbarMS2no(i) <> 0 Or RbarMS2dia(i) <> 0 Then
           pt1(0) = pbx(8 * i - 6) + RbarMS2curS(i) + SWdthLft - sbL(i)
           pt1(1) = pbb(8 * i - 5) + cVr + LinkDia(i) + RbarMS1dia(i) + _
                    20 + RbarMS2dia(i) / 2 + 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) - 10
           pt1(4) = pbx(8 * i - 4) - RbarMS2curE(i) - SWdthRght + sbR(i) - 40
           pt1(5) = pt1(3)
           pt1(6) = pt1(4) + 40
           pt1(7) = pt1(5) + 10
             Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
             Polypt1.layer = "RebarSpan"
             Polypt1.thickness = RbarMS2dia(i) / 2
             Polypt1.Update
           BarMark = BarMark + 1
           Bar5BM = BarMark
           TempX = pbx(8 * i - 4) - RbarTR1curS(i) - cVr - sbR(i) - SWdthRght
           Call ArrowTail(TempX, pt1(3), -350, -30, -15, 150, "LabelRebarSpan")
           Call LabelRbar(TempX, pt1(3) - 380 - 2 * FontSz, _
                          RbarMS2no(i), RbarMS2dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
           Call LabelDimension(pt1(0) - 80, pt1(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7), BarMark, _
                          30, "Curtailment")
           End If

End Function
