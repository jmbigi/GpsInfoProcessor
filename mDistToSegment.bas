Attribute VB_Name = "mDistToSegment"
Option Compare Database
Option Explicit

Public Function DistancePointLine(ByVal px As Double, ByVal py As Double, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  Dim dx As Double
  Dim dy As Double
  Dim t As Double
  Dim near_x As Double
  Dim near_y As Double
  '
  dx = X2 - X1
  dy = Y2 - Y1
  If dx = 0 And dy = 0 Then
      dx = px - X1
      dy = py - Y1
      near_x = X1
      near_y = Y1
  Else
    t = ((px - X1) * dx + (py - Y1) * dy) / (dx * dx + dy * dy)
    If t < 0 Then
        dx = px - X1
        dy = py - Y1
        near_x = X1
        near_y = Y1
    ElseIf t > 1 Then
        dx = px - X2
        dy = py - Y2
        near_x = X2
        near_y = Y2
    Else
        near_x = X1 + t * dx
        near_y = Y1 + t * dy
        dx = px - near_x
        dy = py - near_y
    End If
  End If
  DistancePointLine = distancia(near_x, near_y, X1, Y1)
  '
End Function

