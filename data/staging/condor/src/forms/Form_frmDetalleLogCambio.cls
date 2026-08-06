
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmDetalleLogCambio (CORREGIDO - Manejo de errores en cmdCerrar_Click)
' ==========================================================================

Private Sub Form_Load()
    Dim logServ As New LogServicio
    Dim objLog As LogCambio
    Dim idLog As Long
    
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Detalle del Registro de Cambio"
    idLog = CLng(Nz(Me.OpenArgs, 0))
    
    If idLog <= 0 Then
        Err.Raise 513, , "No se ha especificado un registro de log para mostrar."
    End If
    
    Set objLog = logServ.getLogCambioPorID(idLog)
    If objLog Is Nothing Then
        Err.Raise 513, , "No se pudo cargar el registro del log con ID " & idLog
    End If
    
    Me.txtfechaHora = objLog.fechaHora
    Me.txtusuario = objLog.usuario
    Me.txttabla = objLog.tabla
    Me.txtRegistroID = objLog.registro
    Me.txtcampo = objLog.campo
    Me.txtvalorAnterior = objLog.valorAnterior
    Me.txtvalorNuevo = objLog.valorNuevo
    Me.txtOperacion = objLog.tipoOperacion
    
LimpiarYSalir:
    On Error Resume Next
    Set logServ = Nothing
    Set objLog = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

' --- INICIO DE LA CORRECCIÓN ---
Private Sub cmdCerrar_Click()
    On Error GoTo Errores
    DoCmd.Close acForm, Me.name
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCerrar_Click", Me)
    Resume LimpiarYSalir
End Sub
' --- FIN DE LA CORRECCIÓN ---