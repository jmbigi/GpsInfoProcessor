VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Opciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cmdGenArchGeogr_Click()
On Error GoTo Err_cmdGenArchGeogr_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Generación de Archivos Geográficos"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog

Exit_cmdGenArchGeogr_Click:
    Exit Sub

Err_cmdGenArchGeogr_Click:
    MsgBox Err.Description
    Resume Exit_cmdGenArchGeogr_Click
    
End Sub
Private Sub cmdProcVelEnRed_Click()
On Error GoTo Err_cmdProcVelEnRed_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Procesamiento Velocidades en la Red"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog

Exit_cmdProcVelEnRed_Click:
    Exit Sub

Err_cmdProcVelEnRed_Click:
    MsgBox Err.Description
    Resume Exit_cmdProcVelEnRed_Click
    
End Sub
