Attribute VB_Name = "Module12"
Option Explicit
Public Function CaseTWO()
Dim xStat, yStat As Double
Dim caseI As Integer
Dim pt0(0 To 31) As Double
Dim pt0a(0 To 17) As Double
Dim pt1(0 To 7) As Double
Dim pt2(0 To 11) As Double
Dim TempX As Double
Dim polypt0 As Object

Dim polypt0a As Object

Dim Polypt1 As Object

Dim Polypt2 As Object



           If RbarTL1no(i) <> 0 Or RbarTL1dia(i) <> 0 Or RbarTL1curE(i) <> 0 Then
           pt2(0) = pbx(8 * i - 12) - RbarTR1curS(i - 1) + sbR(i - 1)
           pt2(1) = pbt(8 * i - 11) - cVr - LinkDia(i - 1) - RbarLCdia(i - 1) - _
                    RbarTL1dia(i) / 2 - 10
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
             Polypt2.thickness = RbarTL1dia(i) / 2
             Polypt2.Update
           BarMark = BarMark + 1
           Bar1BM = BarMark
           Call ArrowTail(pt2(0) + RbarTR1curS(i - 1) / 2, pt2(3), 400, 30, _
                          15, 150, "LabelRebarSupt")
           Call LabelRbar(pt2(0) + RbarTR1curS(i - 1) / 2, pt2(3) + 430 + FontSz / 2, _
                          RbarTL1no(i), RbarTL1dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSupt")
           Call LabelDimension(pt2(0) - 80, pt2(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt2(10) - 20, pt2(11) - 30, BarMark, _
                          30, "Curtailment")
           End If
           ''''''''''''''''''''''''''''
           '''' Bar2BM = barMark'''''''
           If RbarTL2no(i) <> 0 Or RbarTL2dia(i) <> 0 Or RbarTL2curE(i) <> 0 Then
           pt2(0) = pbx(8 * i - 12) - RbarTR2curS(i - 1) + sbR(i - 1)
           pt2(1) = pbt(8 * i - 11) - cVr - LinkDia(i) - RbarLCdia(i - 1) - _
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
           ''''''''''''''''''''''''''''''''''''''''
           ''LHS BOTTOM SUPT. BAR --DEFAULT VALUE.'
           pt2(0) = pbx(8 * i - 12) - RbarCRtcurS(i - 1) + sbR(i - 1)
           pt2(1) = pbb(8 * i - 11) + cVr + LinkDia(i - 1) + _
                    RbarMS1dia(i - 1) + RbarCRtdia(i - 1) / 2 + 10
           pt2(2) = pt2(0) + 40
           pt2(3) = pt2(1) - 10
           pt2(4) = pt2(2) + RbarCRtcurS(i - 1) - SWdthLft - 40 ''' pbx(8 * i - 12) - SWdthLft + sbR(i - 1)
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
           Call ArrowTail(pt2(2) + RbarCRtcurS(i) / 2, pt2(3), -400, -30, -15, _
                           150, "LabelRebarSupt")
           Call LabelRbar(pt2(2) + RbarCRtcurS(i) / 2, pt2(3) - 430 - 2 * FontSz, _
                          RbarCLfno(i), RbarCLfdia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSupt")
           Call LabelDimension(pt2(0) - 80, pt2(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt2(10) - 20, pt2(11), BarMark, _
                          30, "Curtailment")
           '''''''''''''''''''''''''''''''''
           ''LINK CARRIER -- DEFAULT VALUE.'
           pt1(0) = pbx(8 * i - 6) + cVr - sbL(i) + SWdthLft
           pt1(1) = pbt(8 * i - 5) - cVr - LinkDia(i) - RbarLCdia(i) / 2 - 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) + 10
           pt1(4) = pbx(8 * i - 4) - cVr + sbR(i) - SWdthRght - 40
           pt1(5) = pt1(3)
           pt1(6) = pt1(4) + 40
           pt1(7) = pt1(5) - 10
             Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
             Polypt1.layer = "RebarSpan"
             Polypt1.thickness = RbarLCdia(i) / 2
             Polypt1.Update
           BarMark = BarMark + 1
           Bar3BM = BarMark
           Call ArrowTail(pt1(0) + 0.45 * beamL(i), pt1(3), 250, 30, 15, _
                          150, "LabelRebarSpan")
           Call LabelRbar(pt1(0) + 0.45 * beamL(i), pt1(3) + 280 + FontSz / 2, _
                          RbarLCno(i), RbarLCdia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
           Call LabelDimension(pt1(0) - 80, pt1(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7) - 30, BarMark, _
                          30, "Curtailment")
           '''''''''''''''''''''
           ''BOTTOM MAIN BAR.'''
           If RbarMS1no(i) <> 0 Or RbarMS1dia(i) <> 0 Then
           pt1(0) = pbx(8 * i - 6) + cVr - sbL(i) + SWdthLft
           pt1(1) = pbb(8 * i - 5) + cVr + LinkDia(i) + RbarMS1dia(i) / 2 + 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) - 10
           pt1(4) = pbx(8 * i - 4) - cVr + sbR(i) - SWdthRght - 40
           pt1(5) = pt1(3)
           pt1(6) = pt1(4) + 40
           pt1(7) = pt1(5) + 10
             Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
             Polypt1.layer = "RebarSpan"
             Polypt1.thickness = RbarMS1dia(i) / 2
             Polypt1.Update
           BarMark = BarMark + 1
           Bar4BM = BarMark
           Call ArrowTail(pt1(0) + RbarTL1curE(i) / 2, pt1(3), -250, -30, _
                          -15, 150, "LabelRebarSpan")
           Call LabelRbar(pt1(0) + RbarTL1curE(i) / 2, pt1(3) - 250 - 2 * FontSz, _
                          RbarMS1no(i), RbarMS1dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
          Call LabelDimension(pt1(0) - 80, pt1(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7), BarMark, _
                          30, "Curtailment")
           End If
           ''''''''''''''''''''''''''''
           If RbarMS2no(i) <> 0 Or RbarMS2dia(i) <> 0 Then
           pt1(0) = pbx(8 * i - 6) + RbarMS2curS(i) - sbL(i) + SWdthLft
           pt1(1) = pbb(8 * i - 5) + cVr + LinkDia(i) + RbarMS1dia(i) + _
                    20 + RbarMS2dia(i) / 2 + 10
           pt1(2) = pt1(0) + 40
           pt1(3) = pt1(1) - 10
           pt1(4) = pbx(8 * i - 4) - RbarMS2curE(i) + sbR(i) - SWdthRght - 40
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
           Call LabelRbar(TempX, pt1(3) - 350 - 2 * FontSz, _
                          RbarMS2no(i), RbarMS2dia(i), BarMark, 0, 0, FontSz, _
                          "T", "LabelRebarSpan")
           Call LabelDimension(pt1(0) - 80, pt1(1), BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt1(6) - 20, pt1(7), BarMark, _
                          30, "Curtailment")
           End If
          ''''''''''''''''''''''''''''''''''
          ''''''''''''''''''''''''''''''''''
       If i = NoOfSpan Then
      
          If RbarTR1no(i) <> 0 Or RbarTR1dia(i) <> 0 Or RbarTR1curS(i) <> 0 Then
               If i = NoOfSpan Then
                If i = 1 Then
                     xStat = pbx(4)
                     yStat = pbt(5)
                Else
           xStat = pbx(8 * i - 2)
           yStat = pbt(8 * i - 1) '''''- slabDrop(i)
                End If
           pt0(0) = xStat - RbarTR1curS(i) + sbR(i)
           pt0(1) = yStat - cVr - LinkDia(i) - RbarLCdia(i) - _
                    RbarTR1dia(i) / 2 - 10
           pt0(2) = pt0(0) + 40
           pt0(3) = pt0(1) + 10
           pt0(4) = pt0(2) + RbarTR1curS(i) + SWdthRght - _
                    2 * cVr - LinkDia(i) - 2 * RbarTR1dia(i) - 40
           pt0(5) = pt0(3)
           pt0(6) = pt0(4) + 2 * Cv1 * RbarTR1dia(i)
           pt0(7) = pt0(5) - 2 * Cv5 * RbarTR1dia(i)
           pt0(8) = pt0(6) + 2 * Cv2 * RbarTR1dia(i)
           pt0(9) = pt0(7) - 2 * Cv4 * RbarTR1dia(i)
           pt0(10) = pt0(8) + 2 * Cv3 * RbarTR1dia(i)
           pt0(11) = pt0(9) - 2 * Cv3 * RbarTR1dia(i)
           pt0(12) = pt0(10) + 2 * Cv4 * RbarTR1dia(i)
           pt0(13) = pt0(11) - 2 * Cv2 * RbarTR1dia(i)
           pt0(14) = pt0(12) + 2 * Cv5 * RbarTR1dia(i)
           pt0(15) = pt0(13) - 2 * Cv1 * RbarTR1dia(i)
           pt0(16) = pt0(14)
               If i = 1 Then
                 pt0(17) = pbb(7) + cVr + LinkDia(i) + RbarMS1dia(i) + _
                           2.5 * RbarTR1dia(i)
               Else
           pt0(17) = pbb(8 * i - 1) + cVr + LinkDia(i) + RbarMS1dia(i) + _
                     2.5 * RbarTR1dia(i)
           
               End If
           pt0(18) = pt0(16) - 2 * Cv5 * RbarTR1dia(i)
           pt0(19) = pt0(17) - 2 * Cv1 * RbarTR1dia(i)
           pt0(20) = pt0(18) - 2 * Cv4 * RbarTR1dia(i)
           pt0(21) = pt0(19) - 2 * Cv2 * RbarTR1dia(i)
           pt0(22) = pt0(20) - 2 * Cv3 * RbarTR1dia(i)
           pt0(23) = pt0(21) - 2 * Cv3 * RbarTR1dia(i)
           pt0(24) = pt0(22) - 2 * Cv2 * RbarTR1dia(i)
           pt0(25) = pt0(23) - 2 * Cv4 * RbarTR1dia(i)
           pt0(26) = pt0(24) - 2 * Cv1 * RbarTR1dia(i)
           pt0(27) = pt0(25) - 2 * Cv5 * RbarTR1dia(i)
           pt0(28) = pt0(26) - RbarCLfcurE(i) - 2 * SWdthRght + 40
           pt0(29) = pt0(27)
           pt0(30) = pt0(28) - 40
           pt0(31) = pt0(29) + 10
             Set polypt0 = moSpace.AddLightWeightPolyline(pt0)
             polypt0.layer = "RebarSupt"
             polypt0.thickness = RbarTR1dia(i) / 2
             polypt0.Update
             polypt0.thickness = 5
           BarMark = BarMark + 1
           Call ArrowTail(pt0(30) + RbarCRtcurS(i) / 2, pt0(29), -400, -30, _
                          -15, 150, "LabelRebarSupt")
           Call LabelRbar(pt0(30) + RbarCRtcurS(i) / 2, pt0(29) - 400 - 2 * FontSz, RbarTR1no(i), _
                          RbarTR1dia(i), BarMark, 0, 0, FontSz, "T", "LabelRebarSupt")
           Call LabelDimension(pt0(0) - 80, pt0(1) - 30, BarMark, _
                          30, "Curtailment")
           Call LabelDimension(pt0(30) - 80, pt0(31), BarMark, _
                          30, "Curtailment")
          End If
          
          ''''''''''''''''''''''''
          ''''''''''''''''''''''''
          If RbarTR2no(i) <> 0 Or RbarTR2dia(i) <> 0 Or RbarTR2curS(i) <> 0 Then
                      
           xStat = pbx(8 * i - 2)
           yStat = pbt(8 * i - 1) ''''''- slabDrop(i)
           
           pt0a(0) = xStat - RbarTR2curS(i) + sbR(i)
           pt0a(1) = yStat - cVr - LinkDia(i) - RbarLCdia(i) - _
                     RbarTR1dia(i) - 20 - RbarTR2dia(i) / 2 - 10
           pt0a(2) = pt0a(0) + 40
           pt0a(3) = pt0a(1) + 10
           pt0a(4) = pt0a(2) + RbarTR2curS(i) + SWdthRght - _
                     2 * cVr - LinkDia(i) - RbarTR1dia(i) - 20 - 2 * RbarTR2dia(i) - 40
           pt0a(5) = pt0a(3)
           pt0a(6) = pt0a(4) + 2 * Cv1 * RbarTR2dia(i)
           pt0a(7) = pt0a(5) - 2 * Cv5 * RbarTR2dia(i)
           pt0a(8) = pt0a(6) + 2 * Cv2 * RbarTR2dia(i)
           pt0a(9) = pt0a(7) - 2 * Cv4 * RbarTR2dia(i)
           pt0a(10) = pt0a(8) + 2 * Cv3 * RbarTR2dia(i)
           pt0a(11) = pt0a(9) - 2 * Cv3 * RbarTR2dia(i)
           pt0a(12) = pt0a(10) + 2 * Cv4 * RbarTR2dia(i)
           pt0a(13) = pt0a(11) - 2 * Cv2 * RbarTR2dia(i)
           pt0a(14) = pt0a(12) + 2 * Cv5 * RbarTR2dia(i)
           pt0a(15) = pt0a(13) - 2 * Cv1 * RbarTR2dia(i)
           pt0a(16) = pt0a(14)
           pt0a(17) = pt0a(15) - beamH(i) + 2 * cVr + LinkDia(i) + _
                      2 * RbarTR1dia(i) + 2 * RbarTR2dia(i)  ''tukar anchor lgth
             Set polypt0a = moSpace.AddLightWeightPolyline(pt0a)
             polypt0a.layer = "RebarSupt"
             polypt0a.thickness = RbarTR2dia(i) / 2
             polypt0a.Update
           BarMark = BarMark + 1
           Call ArrowTail(pt0a(0) + RbarTR2curS(i) / 2, pt0a(3), 400, 30, _
                          15, 150, "LabelRebarSupt")
           Call LabelRbar(pt0a(0) + RbarTR2curS(i) / 2, pt0a(3) + _
           430 + FontSz / 2, RbarTR2no(i), RbarTR2dia(i), BarMark, 0, 0, FontSz, _
           "T", "LabelRebarSupt")
           
          End If
     End If
End If
End Function


