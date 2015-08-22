Attribute VB_Name = "Module16"
Option Explicit

''''GENERAL''''
Public Function ArrowHorizontal(ByVal locx As Double, ByVal _
locy As Double, ByVal Mlength As Double, ByVal ArrowWidth _
As Double, ByVal ArrowHeight As Double, _
LayerName As String) As Object


Dim PolyPt As Object

Dim pt(0 To 19) As Double
pt(0) = locx + ArrowWidth
pt(1) = locy
pt(2) = pt(0)
pt(3) = pt(1) - ArrowHeight / 2
pt(4) = pt(2) - ArrowWidth
pt(5) = pt(3) + ArrowHeight / 2
pt(6) = pt(4) + ArrowWidth
pt(7) = pt(5) + ArrowHeight / 2
pt(8) = pt(6)
pt(9) = pt(7) - ArrowHeight / 2
pt(10) = pt(8) + Mlength - 2 * ArrowWidth
pt(11) = pt(9)
pt(12) = pt(10)
pt(13) = pt(11) + ArrowHeight / 2
pt(14) = pt(12) + ArrowWidth
pt(15) = pt(13) - ArrowHeight / 2
pt(16) = pt(14) - ArrowWidth
pt(17) = pt(15) - ArrowHeight / 2
pt(18) = pt(16)
pt(19) = pt(17) + ArrowHeight / 2


Set PolyPt = moSpace.AddLightWeightPolyline(pt)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
'''PolyPt.Color = 254
PolyPt.layer = LayerName
PolyPt.Update

End Function




