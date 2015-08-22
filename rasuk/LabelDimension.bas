Attribute VB_Name = "Module2"
Option Explicit

''''GENERAL''''
Public Function LabelDimension(ByVal locx As Double, ByVal _
locy As Double, ByVal DimOrNum As Double, _
ByVal txtHeight As Integer, laPisan As String) As Object
Dim acadObj As Object


Dim corner1(0 To 2) As Double
Dim theight As Double
Dim text As String

corner1(0) = locx
corner1(1) = locy
corner1(2) = 0
If txtHeight < 20 Then
txtHeight = 20
End If
text = Str(DimOrNum)
Set acadObj = moSpace.AddText(text, corner1, txtHeight)
acadObj.layer = laPisan
acadObj.Update
End Function



