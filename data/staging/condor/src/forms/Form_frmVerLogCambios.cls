
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmVerLogCambios (Refactorizado con ListBox)
' RESPONSABILIDAD: Visualizar y filtrar la tabla de logs de cambios.
' ==========================================================================

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Visor de Log de Cambios"
    
    Me.cboFiltroUsuario.RowSourceType = "Table/Query"
    Me.cboFiltroUsuario.rowSource = "SELECT '(Todos)' AS U FROM MSysObjects UNION SELECT DISTINCT usuario FROM tbLogCambios ORDER BY U;"
    Me.cboFiltroUsuario.value = "(Todos)"
    
    Me.cboFiltroTabla.RowSourceType = "Table/Query"
    Me.cboFiltroTabla.rowSource = "SELECT '(Todos)' AS T FROM MSysObjects UNION SELECT DISTINCT tabla FROM tbLogCambios ORDER BY T;"
    Me.cboFiltroTabla.value = "(Todos)"

    Call AplicarFiltros
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub AplicarFiltros()
    Dim logServ As New LogServicio
    Dim colLogs As Scripting.Dictionary
    Dim unLog As LogCambio
    Dim filtroUsuario As String, filtroTabla As String, filtroRegistro As String
    Dim rowSourceString As String
    Dim key As Variant
    
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    Me.lstResultados.rowSource = ""
    
    filtroUsuario = Nz(Me.cboFiltroUsuario.value, "")
    filtroTabla = Nz(Me.cboFiltroTabla.value, "")
    filtroRegistro = Nz(Me.txtFiltroRegistro.value, "")
    
    If filtroUsuario = "(Todos)" Then filtroUsuario = ""
    If filtroTabla = "(Todos)" Then filtroTabla = ""
    
    ' Llamada simplificada sin el parámetro de error
    Set colLogs = logServ.getLogs(filtroUsuario, filtroTabla, filtroRegistro)
    
    rowSourceString = "ID;Fecha y Hora;Usuario;Tabla;ID Registro;Campo;Valor Anterior;Valor Nuevo;Operación"
    
    If Not colLogs Is Nothing Then
        If colLogs.count > 0 Then
            For Each key In colLogs.Keys
                Set unLog = colLogs(key)
                rowSourceString = rowSourceString & ";" & _
                    unLog.idLogCambio & ";" & _
                    Format(unLog.fechaHora, "dd/mm/yyyy hh:nn:ss") & ";" & _
                    unLog.usuario & ";" & _
                    unLog.tabla & ";" & _
                    unLog.registro & ";" & _
                    unLog.campo & ";" & _
                    unLog.valorAnterior & ";" & _
                    unLog.valorNuevo & ";" & _
                    unLog.tipoOperacion
            Next key
        End If
    End If
    
    Me.lstResultados.rowSource = rowSourceString
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set logServ = Nothing
    Set colLogs = Nothing
    Set unLog = Nothing
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("AplicarFiltros", Me)
    Resume LimpiarYSalir
End Sub

' ----- AÑADIR ESTAS SUBRUTINAS NUEVAS -----

Private Sub lstResultados_DblClick(Cancel As Integer)
    Call AbrirDetalle
End Sub

Private Sub cmdVerDetalle_Click()
    Call AbrirDetalle
End Sub

Private Sub AbrirDetalle()
    Dim idLog As Long
    
    On Error GoTo Errores
    
    ' La columna 0 es ahora nuestro ID oculto
    idLog = CLng(Nz(Me.lstResultados.Column(0), 0))
    
    If idLog <= 0 Then
        MsgBox "Por favor, seleccione un registro de la lista para ver su detalle.", vbInformation
        Exit Sub
    End If
    If FormularioAbierto("frmDetalleLogCambio") Then
        If Forms("frmDetalleLogCambio").OpenArgs() <> CStr(idLog) Then
            DoCmd.Close acForm, "frmDetalleLogCambio", acSaveNo
        Else
            Forms("frmDetalleLogCambio").SetFocus
            GoTo LimpiarYSalir
        End If
    End If
    
    ' Abrimos el formulario de detalle en modo diálogo, pasándole el ID
    DoCmd.OpenForm "frmDetalleLogCambio", , , , , , CStr(idLog)
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("AbrirDetalle", Me)
    Resume LimpiarYSalir
End Sub

' --- Eventos de controles ---

Private Sub cboFiltroUsuario_AfterUpdate()
    Call AplicarFiltros
End Sub

Private Sub cboFiltroTabla_AfterUpdate()
    Call AplicarFiltros
End Sub

Private Sub txtFiltroRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errores
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call AplicarFiltros
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("txtFiltroRegistro_KeyDown", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdLimpiarFiltros_Click()
    On Error GoTo Errores
    
    Me.cboFiltroUsuario.value = "(Todos)"
    Me.cboFiltroTabla.value = "(Todos)"
    Me.txtFiltroRegistro.value = Null
    
    Call AplicarFiltros
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdLimpiarFiltros_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo Errores
    DoCmd.Close acForm, Me.name
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdSalir_Click", Me)
    Resume LimpiarYSalir
End Sub
