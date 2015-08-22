Attribute VB_Name = "Module15"
Option Explicit

''''FOR RASUK''''
Public Function CircleBarMarkDua(ByVal locx As Double, _
ByVal locy As Double, ByVal LinkDia As Double, _
ByVal b As Double, ByVal h As Double, ByVal SlabT As Double, _
ByVal Bar1No As Integer, ByVal Bar1Dia As Double, ByVal Bar1BM _
As Integer, ByVal Bar2No As Integer, ByVal Bar2Dia As Double, _
ByVal Bar2BM As Integer, ByVal Bar3No As Integer, ByVal Bar3Dia _
As Double, ByVal Bar3BM As Integer, ByVal Bar4No As Integer, _
ByVal Bar4Dia As Double, ByVal Bar4BM As Integer, ByVal Bar5No _
As Integer, ByVal Bar5Dia As Double, ByVal Bar5BM As Integer, _
ByVal Bar6No As Integer, ByVal Bar6Dia As Double, ByVal Bar6BM _
As Integer) As Object
  

Dim circleObj As Object

Dim barmarkObj As Object
Dim center(0 To 2) As Double
Dim Radius As Double
Dim insPnt(0 To 2) As Double
Dim textHgt As Double
Dim textStr As String
Dim Rotate As Double
Dim j As Integer
Dim deltaX, deltaY As Double

textHgt = 25

If Bar1No <> 0 Then
  center(0) = locx - 1.5 * LinkDia + Bar1Dia / 2
  center(1) = locy - LinkDia / 2 - Bar3Dia - Bar1Dia / 2
  If Bar1No = 1 Then
  deltaX = 0
  Else
  deltaX = (b - 2 * cVr - 2 * LinkDia - Bar1Dia) / (Bar1No - 1)
  End If
  
  For j = 1 To Bar1No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar1Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
  
  
  insPnt(1) = locy + (Bar1Dia + Bar3Dia + LinkDia + 2.5 * cVr) + textHgt / 2
  insPnt(2) = 0
 
  textStr = Trim(Str(Bar1BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If

If Bar2No <> 0 Then
  center(0) = locx - 1.5 * LinkDia + Bar2Dia / 2
  center(1) = locy - LinkDia / 2 - Bar3Dia - Bar1Dia - 20 - Bar2Dia / 2
  If Bar2No = 1 Then
  deltaX = 0
  Else
  deltaX = (b - 2 * cVr - 2 * LinkDia - Bar2Dia) / (Bar2No - 1)
  End If
  
  For j = 1 To Bar2No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar2Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
  
  
  insPnt(1) = locy + (Bar2Dia + 20 + Bar1Dia + Bar3Dia + LinkDia + _
              3.5 * cVr) + textHgt
  insPnt(2) = 0
 
  textStr = Trim(Str(Bar2BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If

'' Link carrier
If Bar3No <> 0 Then
  center(0) = locx
  center(1) = locy - LinkDia / 2 - Bar3Dia / 2
  If Bar3No = 1 Then
  deltaX = 0
  Else
  deltaX = (b - 2 * cVr - 5 * LinkDia) / (Bar3No - 1)
  End If
  
  For j = 1 To Bar3No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar3Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
    
  
  insPnt(1) = locy + (Bar3Dia + LinkDia + 1 * cVr) '''+ textHgt
  insPnt(2) = 0
 
  textStr = Trim(Str(Bar3BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If

''bottom main bar
If Bar4No <> 0 Then
  center(0) = locx
  center(1) = locy - h + 2 * cVr + 1.5 * LinkDia + Bar4Dia / 2
  If Bar4No = 1 Then
  deltaX = 0
  Else
  deltaX = (b - 2 * cVr - 5 * LinkDia) / (Bar4No - 1)
  End If
  
  For j = 1 To Bar4No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar4Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
    
  
  insPnt(1) = locy - h - (Bar4Dia + 1.5 * cVr) '''- textHgt
  insPnt(2) = 0
 
  textStr = Trim(Str(Bar4BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If

If Bar5No <> 0 Then
  center(0) = locx - 1.5 * LinkDia + Bar5Dia / 2
  center(1) = locy - h + 2 * cVr + 1.5 * LinkDia + 2 * Bar4Dia + Bar5Dia / 2
  If Bar5No = 1 Then
  deltaX = 0
  Else
  deltaX = (b - 2 * cVr - 2 * LinkDia - Bar5Dia) / (Bar5No - 1)
  End If
  
  For j = 1 To Bar5No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar5Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
    
  
  insPnt(1) = locy - h - (Bar4Dia + 7.5 * cVr) - textHgt
  insPnt(2) = 0

  textStr = Trim(Str(Bar5BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If

If Bar6No <> 0 Then
  center(0) = locx - 1.5 * LinkDia + Bar6Dia / 2
  center(1) = locy - h + 2 * cVr + 1.5 * LinkDia + Bar4Dia + Bar6Dia / 2
  If Bar6No = 1 Then
  deltaX = 0
  Else
   deltaX = (b - 2 * cVr - 2 * LinkDia - Bar6Dia) / (Bar6No - 1)
  End If
  
  For j = 1 To Bar6No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar6Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
    
  
  insPnt(1) = locy - h - (Bar4Dia + 4.25 * cVr) - textHgt / 2
  insPnt(2) = 0
 
  textStr = Trim(Str(Bar6BM))
  insPnt(0) = center(0) + 2 * textHgt / 3
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
 Rotate = 1.57  '' 90 degrees
 Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.layer = "BeamSection"
  barmarkObj.Color = 7
  barmarkObj.Update
  Next
End If


''''''''''''''anti crack''''''''''''''''

If beamH(i) >= 750 Then
  center(0) = locx - 1.5 * LinkDia + 16 / 2
  center(1) = locy - h + 2 * cVr + 250
    
  For j = 1 To 2
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = 16 / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
  center(1) = center(1) + 200
  Next
  
  center(0) = locx + b - 3.5 * LinkDia - 16 / 2 - 2 * cVr
  center(1) = locy - h + 2 * cVr + 250
    
  For j = 1 To 2
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = 16 / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "BeamSection"
  circleObj.Color = 30
  circleObj.Update
  center(1) = center(1) + 200
  Next
  
End If
End Function




