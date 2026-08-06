Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmFiltrosAvanzadosSolicitudes (Versión con Estado)
' ==========================================================================

Private m_Filtros As FiltrosSolicitud

Public Property Get filtros() As FiltrosSolicitud
    Set filtros = m_Filtros
End Property

Private Sub ComandoLimpiarResponsableCalidad_Click()
    ' Resetea el combo al valor que representa "(Todos)", que es 0.
    Me.cboResponsableCalidad.value = 0
End Sub

Private Sub ComandoLimpiarResponsableTecnico_Click()
    ' Resetea el combo al valor que representa "(Todos)", que es 0.
    Me.cboResponsableTecnico.value = 0
End Sub

Private Sub cmdLimpiarSuministrador_Click()
    ' Resetea el combo al valor que representa "(Todos)", que es 0.
    Me.cboSuministrador.value = 0
End Sub

Private Sub cmdLimpiarTipoSolicitud_Click()
    ' Resetea el combo al valor que representa "(Todos)", que es una cadena vacía.
    Me.cboTipoSolicitud.value = ""
End Sub

Private Sub cmdLimpiarEstado_Click()
    ' Resetea el combo al valor que representa "(Todos)", que es 0.
    Me.cboEstado.value = 0
End Sub


Private Sub Form_Load()
    Dim frmBusqueda As Form_frmBuscarSolicitudes
    On Error GoTo Errores
    
    
    DoCmd.Hourglass True
    DoEvents
    
    Ajustar Me
    Me.Caption = "Filtros Avanzados de Búsqueda"
    
    ' --- INICIO DE LA MODIFICACIÓN ---
    ' El formulario de filtros ya no se preocupa por el estado del checkbox.
    ' Simplemente carga el estado actual del objeto de filtros que le pasan.
    If FormularioAbierto("frmBuscarSolicitudes") Then
        Set frmBusqueda = Forms("frmBuscarSolicitudes")
        Set m_Filtros = frmBusqueda.FiltrosActivos
    Else
        Set m_Filtros = New FiltrosSolicitud
    End If
    ' --- FIN DE LA MODIFICACIÓN ---
    
    Call PoblarCombos
    Call PoblarControlesDesdeFiltros
    
    Me.cmdAplicar.Enabled = True
    Me.cmdCancelar.Enabled = True
    
LimpiarYSalir:
    On Error Resume Next
   
    DoCmd.Hourglass False
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    Resume LimpiarYSalir
End Sub
Private Sub PoblarCombos()
    ' Rellena los desplegables con datos de la caché del Entorno o de la BBDD
    
    ' Usamos nuestra función pública y genérica para los combos basados en diccionarios
    Call LlenarComboDesdeDiccionario(Me.cboResponsableTecnico, m_ObjEntorno.ResponsablesTecnicos, "Id", "Nombre")
    Call LlenarComboDesdeDiccionario(Me.cboResponsableCalidad, m_ObjEntorno.ResponsablesCalidad, "Id", "Nombre")
    Call LlenarComboDesdeDiccionario(Me.cboSuministrador, m_ObjEntorno.Suministradores, "IDSuministrador", "NombreParaMostrar")
    
    ' Para Tipo Solicitud, usamos una lista de valores fija
    Me.cboTipoSolicitud.RowSourceType = "Value List"
    Me.cboTipoSolicitud.ColumnCount = 2
    Me.cboTipoSolicitud.BoundColumn = 1
    Me.cboTipoSolicitud.ColumnWidths = "0cm;8cm"
    
    ' --- LÍNEA CORREGIDA ---
    ' El valor para "(Todos)" es una cadena vacía (''), y el texto visible es '(Todos)'.
    ' El resto de las opciones siguen el patrón 'VALOR';'TEXTO VISIBLE'.
    Me.cboTipoSolicitud.rowSource = "'''';(Todos);'PC';'PC - Propuesta de Cambio';'CD_CA';'CD - Concesión desviación';'CD_CA_SUB';'CD_sub - Concesión desviación subsiministrador'"
    
    ' Esta línea ahora funcionará correctamente para seleccionar "(Todos)" por defecto
    Me.cboTipoSolicitud.value = ""
    
    ' Para el combo de Estados, usamos la caché del Entorno
    Call LlenarComboDesdeDiccionario(Me.cboEstado, m_ObjEntorno.estados, "idEstado", "nombreEstado")
End Sub
' --- NUEVA FUNCIÓN ---
Private Sub PoblarControlesDesdeFiltros()
    ' Responsabilidad: Asigna los valores del objeto m_Filtros a los controles.
    If m_Filtros Is Nothing Then Exit Sub
    
    Me.cboResponsableTecnico.value = m_Filtros.idResponsableTecnico
    Me.cboResponsableCalidad.value = m_Filtros.idResponsableCalidad
    Me.cboSuministrador.value = m_Filtros.idSuministrador
    Me.cboTipoSolicitud.value = m_Filtros.tipoSolicitud
    Me.cboEstado.value = m_Filtros.idEstado
End Sub

Private Sub cmdAplicar_Click()
    ' (Esta función no cambia)
    With m_Filtros
        .idResponsableTecnico = CLng(Nz(Me.cboResponsableTecnico.value, 0))
        .idResponsableCalidad = CLng(Nz(Me.cboResponsableCalidad.value, 0))
        .idSuministrador = CLng(Nz(Me.cboSuministrador.value, 0))
        .tipoSolicitud = Nz(Me.cboTipoSolicitud.value, "")
        If .tipoSolicitud = "'" Then .tipoSolicitud = ""
        .idEstado = CStr(Nz(Me.cboEstado.value, "0"))
    End With
    
    Me.Tag = "OK"
    Me.Visible = False
End Sub

Private Sub cmdCancelar_Click()
    ' (Esta función no cambia)
    Me.Tag = "Cancel"
    Me.Visible = False
End Sub