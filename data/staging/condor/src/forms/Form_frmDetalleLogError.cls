
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmDetalleLogError (CORREGIDO - Manejo de errores en cmdCerrar_Click)
' ==========================================================================

Private Sub Form_Load()
    Dim logServ As New LogServicio
    Dim objError As LogError
    Dim idLog As Long
    
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Detalle del Registro de Error"
    idLog = CLng(Nz(Me.OpenArgs, 0))
    
    If idLog <= 0 Then
        Err.Raise 513, , "No se ha especificado un registro de error para mostrar."
    End If
    
    Set objError = logServ.getLogErrorPorID(idLog)
    If objError Is Nothing Then
        Err.Raise 513, , "No se pudo cargar el registro del log de error con ID " & idLog
    End If
    
    Me.txtfechaHora = objError.fechaHora
    Me.txtusuario = objError.usuario
    Me.txtsuplantadoPor = objError.suplantadoPor
    Me.txtModulo = objError.modulo
    Me.txtprocedimiento = objError.procedimiento
    Me.txtnumeroError = objError.numeroError
    Me.txtdescripcionError = objError.descripcionError
    Me.txtcontexto = objError.contexto
    
LimpiarYSalir:
    On Error Resume Next
    Set logServ = Nothing
    Set objError = Nothing
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