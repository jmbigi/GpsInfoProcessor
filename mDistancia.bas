Attribute VB_Name = "mDistancia"
Option Explicit

Public Const PI As Double = 3.14159265358979
Public Const RadiusKm As Double = 6378.137

Function distancia(ByVal Lon1 As Double, ByVal Lat1 As Double, ByVal Lon2 As Double, ByVal Lat2 As Double, Optional unit As String = "M") As Double
    Dim dlon As Double
    Dim dist As Double
    
    Lon1 = Lon1 / 1000000#
    Lat1 = Lat1 / 1000000#
    Lon2 = Lon2 / 1000000#
    Lat2 = Lat2 / 1000000#
    
    Lat1 = Lat1 / 180# * PI
    Lon1 = Lon1 / 180# * PI
    Lat2 = Lat2 / 180# * PI
    Lon2 = Lon2 / 180# * PI
    
    dist = 2# * ArcSin(Sqr((Sin((Lat1 - Lat2) / 2#) ^ 2#) + Cos(Lat1) * Cos(Lat2) * (Sin((Lon1 - Lon2) / 2#) ^ 2#))) * RadiusKm
    
    Select Case UCase(unit)
    Case "K"
        distancia = dist
    Case "M"
        distancia = dist * 1000#
    Case "L"
        distancia = dist * 0.621371192
    Case "N"
        distancia = dist * 0.539956803
    Case Else
        Err.Raise 1001, , "Error en la unidad"
    End Select
    
End Function

Function acos(ByVal rad As Double) As Double
  
  If Abs(rad) <> 1# Then
    acos = PI / 2# - Atn(rad / Sqr(1# - rad * rad))
  ElseIf rad = -1# Then
    acos = PI
  End If

End Function

Function ArcSin(ByVal x As Double) As Double
    ArcSin = Atn(x / Sqr(-x * x + 1#))
End Function

Private Sub test()
    Dim unit As String
    Const Lat1 As Double = -41472849#
    Const Lon1 As Double = -72934593
    Const Lat2 As Double = -41473822#
    Const Lon2 As Double = -72935243
       
    unit = "M"
    Debug.Print unit
    Debug.Print distancia(Lon1, Lat1, Lon2, Lat2, unit) & " " & unit
    Debug.Print ""
    
    
End Sub

