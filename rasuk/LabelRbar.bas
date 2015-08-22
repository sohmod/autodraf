Attribute VB_Name = "Module3"
Option Explicit


''''GENERAL''''
Public Function LabelRbar(ByVal locx As Double, ByVal _
locy As Double, ByVal rbarNo As Integer, ByVal rbarDia _
As Integer, ByVal rbarMark As Integer, ByVal RbarSpacing _
As Double, ByVal Rotate As Double, ByVal txtHeight As Integer, _
barType As String, laPisan As String) As Object
 

Dim acadObj As Object

Dim corner1(0 To 2) As Double
Dim text As String

corner1(0) = locx
corner1(1) = locy
corner1(2) = 0
If txtHeight < 20 Then
txtHeight = 20
End If
If RbarSpacing = 0 Then
  text = Trim(Str(rbarNo)) & barType & Trim(Str(rbarDia)) & "-" & _
       Trim(Str(rbarMark))
Else
  text = Trim(Str(rbarNo)) & barType & Trim(Str(rbarDia)) & "-" & _
       Trim(Str(rbarMark)) & "-" & Trim(Str(RbarSpacing))
End If

Set acadObj = moSpace.AddText(text, corner1, txtHeight)
Call acadObj.Rotate(corner1, Rotate)
acadObj.layer = laPisan
'''acadObj.Color = 0
acadObj.Update

End Function


