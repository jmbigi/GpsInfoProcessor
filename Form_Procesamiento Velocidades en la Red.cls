VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Procesamiento Velocidades en la Red"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ChkJuntarTracks_Click()
  controlarOpciones
End Sub

Private Sub chkSoloGeo_Click()
  controlarOpciones
End Sub

Private Sub Command0_Click()
  EjecutarBatch
End Sub

Private Sub Form_Load()
  cargarParametrosPredeterminados
End Sub

Private Sub cargarParametrosPredeterminados()

End Sub

Private Sub controlarOpciones()
  Rem Me.chkUtlizarNodos.Enabled = (chkSoloGeo.Value = True)
  Me.chkSegmentar.Enabled = (chkSoloGeo.Value = True)
End Sub

Private Sub EjecutarBatch()
  Dim nombreVentana As String
  '
  nombreVentana = Me.Name
  If ProcesarBatch(Me, False, (Me.ChkJuntarTracks.Value = True), (Me.chkSoloGeo.Value = True), (Me.chkUtlizarNodos.Value = True), (Me.chkSegmentar.Value = True), Me.cboTipoSalida.ListCount, Me.chkTPoint_6_15_11.Value = True, Me.txtDistMaxBusq.Value, Me.txtRelMaxDistAlArcoLongArco.Value, Me.txtRelMaxDistMedidLongArco, Me.txtVirajeMaxUnMetro.Value, Me.txtDistMaxSep) Then
    Screen.MousePointer = 0
    MsgBox "Terminó de Procesar", vbInformation, nombreVentana & " - " & CurrentProject.Name
  End If
  Screen.MousePointer = 0
  '
End Sub
