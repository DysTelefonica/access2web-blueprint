
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmBuscarExpediente (CORREGIDO - Manejo de errores en cmdSalir_Click)
' ==========================================================================

Public Event ExpedienteElegido(idExpedienteSeleccionado As Long)

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Búsqueda de Expedientes Registrados con PECAL"
    
    With Me.lstResultados
        .RowSourceType = "Value List"
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "0cm;2cm;6cm;3cm;3cm;5,5cm;5cm"
    End With
    
    If Not IsNull(Me.OpenArgs) Then
        Me.txtPalabraClave.value = Me.OpenArgs
    End If
    
    Call PoblarCombosDeFiltro
    Call RealizarBusqueda
    
LimpiarYSalir:
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub cboResponsableCalidad_AfterUpdate()
    Call RealizarBusqueda
End Sub

Private Sub cboResponsableTecnico_AfterUpdate()
    Call RealizarBusqueda
End Sub

Private Sub txtPalabraClave_AfterUpdate()
    Call RealizarBusqueda
End Sub

Private Sub cmdElegir_Click()
    Call SeleccionarItemActual
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)
    Call SeleccionarItemActual
End Sub

Private Sub cmdLimpiarRespCalidad_Click()
    Me.cboResponsableCalidad.value = 0
    Call RealizarBusqueda
End Sub

Private Sub cmdLimpiarRespTecnico_Click()
    Me.cboResponsableTecnico.value = 0
    Call RealizarBusqueda
End Sub

Private Sub cmdLimpiarPalabraClave_Click()
    Me.txtPalabraClave.value = Null
    Call RealizarBusqueda
End Sub

' --- INICIO DE LA CORRECCIÓN ---
Private Sub cmdSalir_Click()
    On Error GoTo Errores
    
    Me.Tag = "" ' Nos aseguramos de que el Tag esté vacío si se cancela
    DoCmd.Close acForm, Me.name
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdSalir_Click", Me)
    Resume LimpiarYSalir
End Sub
' --- FIN DE LA CORRECCIÓN ---

Private Sub PoblarCombosDeFiltro()
    On Error GoTo Errores
    
    Dim colTecnicos As Scripting.Dictionary
    Dim colCalidad As Scripting.Dictionary
    Set colTecnicos = m_ObjEntorno.ResponsablesTecnicos
    Set colCalidad = m_ObjEntorno.ResponsablesCalidad
    
    Call LlenarComboDesdeDiccionario(Me.cboResponsableTecnico, colTecnicos, "Id", "Nombre")
    Call LlenarComboDesdeDiccionario(Me.cboResponsableCalidad, colCalidad, "Id", "Nombre")
    
    Exit Sub
Errores:
     Call ManejadorErroresFormulario("PoblarCombosDeFiltro", Me)
End Sub

' ==========================================================================
' MÓDULO: Form_frmBuscarExpediente.cls
' ACCIÓN: Corregir la visualización de fechas nulas en la lista de resultados.
' ==========================================================================

Private Sub RealizarBusqueda()
    Dim gestorExpedientes As New ExpedienteServicio
    Dim colExpedientes As Scripting.Dictionary
    Dim unExpediente As Expediente
    Dim idKey As Variant
    Dim palabraClaveFiltro As String, idRespTecnicoFiltro As String, idRespCalidadFiltro As String
    Dim rowSourceString As String
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    Me.lstResultados.rowSource = ""
    
    palabraClaveFiltro = Nz(Me.txtPalabraClave.value, "")
    idRespTecnicoFiltro = CStr(Nz(Me.cboResponsableTecnico.value, 0))
    idRespCalidadFiltro = CStr(Nz(Me.cboResponsableCalidad.value, 0))
    
    Set colExpedientes = gestorExpedientes.getExpedientes(palabraClaveFiltro, idRespCalidadFiltro, idRespTecnicoFiltro)
    
    rowSourceString = "IDExpediente;Cod. Exp;Nemotécnico;F. Inicio;F. Fin;Resp. Técnico;RAC Asignado"
    
    If colExpedientes.count > 0 Then
        For Each idKey In colExpedientes.Keys
            Set unExpediente = colExpedientes(idKey)
            
            ' --- INICIO DE LA CORRECCIÓN ---
            ' Usamos una expresión IIf para formatear la fecha solo si es válida (> 0),
            ' de lo contrario, devolvemos una cadena vacía.
            rowSourceString = rowSourceString & ";" & unExpediente.idExpediente _
                & ";" & unExpediente.CodExp _
                & ";" & unExpediente.Nemotecnico _
                & ";" & IIf(unExpediente.FechaInicioContrato > 0, Format(unExpediente.FechaInicioContrato, "dd/mm/yyyy"), "") _
                & ";" & IIf(unExpediente.FechaFinContrato > 0, Format(unExpediente.FechaFinContrato, "dd/mm/yyyy"), "") _
                & ";" & Nz(unExpediente.ResponsableTecnico, "N/A") _
                & ";" & Nz(unExpediente.RacAsignado, "N/A")
            ' --- FIN DE LA CORRECCIÓN ---
        Next idKey
    End If
    
    Me.lstResultados.rowSource = rowSourceString

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub

Errores:
    Call ManejadorErroresFormulario("RealizarBusqueda", Me)
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmBuscarExpediente.cls (ACTUALIZADO)
' ACCIÓN: Devolver el ID seleccionado a través de la propiedad Tag.
' ==========================================================================

Private Sub SeleccionarItemActual()
    Dim selectedExpedienteID As Long
    On Error GoTo Errores
    
    ' Si no hay nada seleccionado, no hacemos nada.
    If Me.lstResultados.ListIndex < 1 Then Exit Sub
    
    selectedExpedienteID = CLng(Nz(Me.lstResultados.Column(0, Me.lstResultados.ListIndex), 0))
    
    If selectedExpedienteID = 0 Then Exit Sub
    
    ' --- INICIO DE LA MODIFICACIÓN ---
    ' Guardamos el ID seleccionado en la propiedad Tag del formulario.
    ' El formulario llamador (en modo diálogo) podrá leer este valor.
    Me.Tag = CStr(selectedExpedienteID)
    
    ' Ocultamos el formulario en lugar de cerrarlo.
    ' El código llamador (en modo acDialog) leerá el Tag y luego cerrará el formulario.
    Me.Visible = False
    ' --- FIN DE LA MODIFICACIÓN ---
    
LimpiarYSalir:
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("SeleccionarItemActual", Me)
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmBuscarExpediente.cls
' ACCIÓN: CORREGIR cmdVerDetalle_Click para abrir el formulario de detalle en modo Diálogo.
' ==========================================================================
Private Sub cmdVerDetalle_Click()
    Dim selectedExpedienteID As String
    
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Abriendo detalle del expediente..."
    DoCmd.Hourglass True
    DoEvents
    
    If Me.lstResultados.ListIndex < 1 Then
        MsgBox "Por favor, seleccione un expediente de la lista para ver su detalle.", vbInformation, "Condor"
        GoTo LimpiarYSalir
    End If
    
    selectedExpedienteID = Nz(Me.lstResultados.Column(0, Me.lstResultados.ListIndex), "")
    
    If selectedExpedienteID = "" Then
        MsgBox "No se pudo obtener el identificador del expediente seleccionado.", vbExclamation, "Condor"
        GoTo LimpiarYSalir
    End If
    
    ' --- INICIO DE LA CORRECCIÓN ---
    ' Añadimos el argumento acDialog para forzar la apertura en modo modal.
    ' Esto asegura que el formulario aparezca siempre en primer plano.
    DoCmd.OpenForm "frmDetalleExpediente", acNormal, , , , acDialog, selectedExpedienteID
    ' --- FIN DE LA CORRECCIÓN ---
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub

Errores:
    Call ManejadorErroresFormulario("cmdVerDetalle_Click", Me)
    Resume LimpiarYSalir
End Sub
