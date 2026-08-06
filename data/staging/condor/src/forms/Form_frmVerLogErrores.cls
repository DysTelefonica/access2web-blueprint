
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmVerLogErrores (Refactorizado con ListBox)
' RESPONSABILIDAD: Visualizar y filtrar la tabla de logs de errores.
' ==========================================================================

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Visor de Log de Errores"
    
    ' Poblar combo de Usuarios
    Me.cboFiltroUsuario.RowSourceType = "Table/Query"
    Me.cboFiltroUsuario.rowSource = "SELECT '(Todos)' AS U FROM MSysObjects UNION SELECT DISTINCT usuario FROM tbLogErrores ORDER BY U;"
    Me.cboFiltroUsuario.value = "(Todos)"
    
    ' Poblar combo de Módulos
    Me.cboFiltroModulo.RowSourceType = "Table/Query"
    Me.cboFiltroModulo.rowSource = "SELECT '(Todos)' AS M FROM MSysObjects UNION SELECT DISTINCT modulo FROM tbLogErrores ORDER BY M;"
    Me.cboFiltroModulo.value = "(Todos)"

    ' Carga inicial de datos
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
    Dim colErrores As Scripting.Dictionary
    Dim unError As LogError
    Dim filtroUsuario As String, filtroModulo As String, filtroProcedimiento As String
    Dim rowSourceString As String
    Dim key As Variant
    
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    Me.lstResultados.rowSource = ""
    
    filtroUsuario = Nz(Me.cboFiltroUsuario.value, "")
    filtroModulo = Nz(Me.cboFiltroModulo.value, "")
    filtroProcedimiento = Nz(Me.txtFiltroProcedimiento.value, "")
    
    If filtroUsuario = "(Todos)" Then filtroUsuario = ""
    If filtroModulo = "(Todos)" Then filtroModulo = ""
    
    ' Llamada simplificada sin el parámetro de error
    Set colErrores = logServ.getErrores(filtroUsuario, filtroModulo, filtroProcedimiento)
    
    rowSourceString = "ID;Fecha y Hora;Usuario;Módulo;Nº Error;Descripción"
    
    If Not colErrores Is Nothing Then
        If colErrores.count > 0 Then
            For Each key In colErrores.Keys
                Set unError = colErrores(key)
                rowSourceString = rowSourceString & ";" & _
                    unError.idLogError & ";" & _
                    Format(unError.fechaHora, "dd/mm/yyyy hh:nn:ss") & ";" & _
                    unError.usuario & ";" & _
                    unError.modulo & ";" & _
                    unError.numeroError & ";" & _
                    unError.descripcionError
            Next key
        End If
    End If
    
    Me.lstResultados.rowSource = rowSourceString
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set logServ = Nothing
    Set colErrores = Nothing
    Set unError = Nothing
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("AplicarFiltros", Me)
    Resume LimpiarYSalir
End Sub

' ----- AÑADIR ESTAS SUBRUTINAS NUEVAS AL FINAL DEL MÓDULO -----

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
    If FormularioAbierto("frmDetalleLogError") Then
        If Forms("frmDetalleLogError").OpenArgs() <> CStr(idLog) Then
            DoCmd.Close acForm, "frmDetalleLogError", acSaveNo
        Else
            Forms("frmDetalleLogError").SetFocus
            GoTo LimpiarYSalir
        End If
    End If
    ' Abrimos el formulario de detalle en modo diálogo, pasándole el ID
    DoCmd.OpenForm "frmDetalleLogError", , , , , , CStr(idLog)
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("AbrirDetalle", Me)
    GoTo LimpiarYSalir
End Sub

' --- Eventos de controles ---

Private Sub cboFiltroUsuario_AfterUpdate()
    Call AplicarFiltros
End Sub

Private Sub cboFiltroModulo_AfterUpdate()
    Call AplicarFiltros
End Sub

Private Sub txtFiltroProcedimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errores
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call AplicarFiltros
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("txtFiltroProcedimiento_KeyDown", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdLimpiarFiltros_Click()
    On Error GoTo Errores
    
    Me.cboFiltroUsuario.value = "(Todos)"
    Me.cboFiltroModulo.value = "(Todos)"
    Me.txtFiltroProcedimiento.value = Null
    
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
