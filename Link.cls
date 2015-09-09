VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Shape As New Collection
Public Segments As New Collection

Public IdNodeA As Long
Public IdNodeB As Long

Public Function Key() As String
  Key = IdNodeA & "-" & IdNodeB
End Function

Public Function minDistToPoint(ByVal x As Double, ByVal y As Double) As Double
  Dim min As Double
  Dim aux As Double
  Dim segm As segment
  Dim minX1 As Double
  Dim minY1 As Double
  Dim minX2 As Double
  Dim minY2 As Double
  
  min = 999999999
  
  For Each segm In Segments
    aux = DistancePointLine(x, y, segm.pointA.x, segm.pointA.y, segm.pointB.x, segm.pointA.y)
    If aux < min Then
      min = aux
      minX1 = segm.pointA.x
      minY1 = segm.pointA.y
      minX2 = segm.pointB.x
      minY2 = segm.pointB.y
      Debug.Print minX1 / 1000000, vbTab, minY1 / 1000000, " - ", minX2 / 1000000, vbTab, minY2 / 1000000, vbTab, distancia(x, y, minX2, minY2)
    End If
  Next
  minDistToPoint = min
  
End Function

Public Function loadSegments()
  Dim db As Database
  Dim rs As Recordset
  Dim point As point
  Dim segment As segment
  
  Set Shape = New Collection
  Set Segments = New Collection
  
  Set db = CurrentDb
  
  Set rs = db.OpenRecordset("select * from [Shapes] where nodo_a=" & IdNodeA & " and nodo_b=" & IdNodeB & " order by correlativo", dbOpenDynaset)
  Do While Not rs.EOF
    If Not segment Is Nothing Then
      Set segment.pointA = point
      Set point = New point
      point.x = rs.Fields("longitude").Value
      point.y = rs.Fields("latitude").Value
      Set segment.pointB = point
      Segments.Add segment
      Set segment = New segment
    Else
      Set point = New point
      point.x = rs.Fields("longitude").Value
      point.y = rs.Fields("latitude").Value
      Set segment = New segment
    End If
    Shape.Add point
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  
End Function

