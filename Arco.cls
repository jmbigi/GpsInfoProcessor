VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ID As Long
Public trackId As Long
Public trackName As String
Public SegmentID As Long
Public HoraInicial As Variant
Public HoraFinal As Variant
'
Public LengthMts As Long
Public Segundos As Long
'
Public AltitudeMts As Long    'multiplicado por distancia
Public IndPendAbs As Long    'multiplicado por distancia
Public IndiceGiro As Long    'multiplicado por distancia
'
Public DistPesoAltitudeMts As Long
Public DistPesoIndPendAbs As Long
Public DistPesoIndiceGiro As Long
'
Public NroPuntos As Integer
Public Coordenadas As New Collection
'

Public Function KmsHr() As Variant
  If IsNull(Segundos) Or IsNull(LengthMts) Then
    KmsHr = Null
  ElseIf Segundos = 0 Then
    KmsHr = Null
  Else
    KmsHr = Round((LengthMts / 1000) / (Segundos / 3600), 0)
  End If
End Function

Public Function Hora() As Variant
  Dim aux As Date
  Dim dif As Variant
  '
  If IsNull(HoraInicial) Or IsNull(HoraFinal) Then
    Hora = Null
  Else
    aux = HoraInicial
    dif = HoraFinal - HoraInicial
    aux = aux + dif / 2
    Hora = aux
  End If
  '
End Function


