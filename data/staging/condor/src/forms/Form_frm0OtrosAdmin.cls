Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frm0OtrosAdmin (Versión Simplificada)
' RESPONSABILIDAD: Menú de opciones de consulta y auditoría para Administradores.
' ==========================================================================

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Apartado adminstracion"
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdBuscarExpedientes_Click()
    On Error GoTo Errores
    
    DoCmd.OpenForm "frmBuscarExpediente"
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdBuscarExpedientes_Click", Me)
    Resume LimpiarYSalir
End Sub



Private Sub cmdVerLogCambios_Click()
    On Error GoTo Errores
    
    DoCmd.OpenForm "frmVerLogCambios"
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVerLogCambios_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdVerLogErrores_Click()
    On Error GoTo Errores
    
    DoCmd.OpenForm "frmVerLogErrores"
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVerLogErrores_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdVolver_Click()
    On Error GoTo Errores
    DoCmd.Close acForm, Me.name
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVolver_Click", Me)
    Resume LimpiarYSalir
End Sub