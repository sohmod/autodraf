Attribute VB_Name = "Module18"
Option Explicit



''''GENERAL''''
Public Function ArrowTail(ByVal locx As Double, ByVal _
locy As Double, ByVal Mlength As Double, ByVal ArrowWidth _
As Double, ByVal ArrowHeight As Double, ByVal Mwidth As Double, _
LayerName As String) As Object


Dim PolyPt As Object


Dim pt(0 To 13) As Double
pt(0) = locx
pt(1) = locy + ArrowWidth
pt(2) = pt(0) + ArrowHeight / 2
pt(3) = pt(1)
pt(4) = pt(2) - ArrowHeight / 2
pt(5) = pt(3) - ArrowWidth
pt(6) = pt(4) - ArrowHeight / 2
pt(7) = pt(5) + ArrowWidth
pt(8) = pt(6) + ArrowHeight / 2
pt(9) = pt(7)
pt(10) = pt(8)
pt(11) = pt(9) + Mlength
pt(12) = pt(10) + Mwidth
pt(13) = pt(11)


Set PolyPt = moSpace.AddLightWeightPolyline(pt)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
'''PolyPt.Color = 254
PolyPt.layer = LayerName
PolyPt.Update

locx = pt(12) + ArrowWidth
locy = pt(13) - ArrowWidth / 2

End Function



