
Option Compare Database
Option Explicit


Private WithEvents m_frmGestion As Form_frmGestionSolicitud

Private m_FiltrosAvanzados As FiltrosSolicitud


Public Property Set FiltroInicial(value As FiltrosSolicitud)
    Set m_FiltrosAvanzados = value
End Property

Private Sub cmdLimpiarPalabraClave_Click()
    Me.txtPalabraClave = Null
    Call RealizarBusqueda
End Sub

Private Sub Form_Load()
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    Ajustar Me
    Me.Caption = "Búsqueda de Solicitudes"
    
    With Me.lstResultados
        .RowSourceType = "Value List"
        .ColumnCount = 9 ' idSolicitud, idExpediente + 4 visibles
        .ColumnHeads = True
        .ColumnWidths = "0cm;0cm;4,5cm;5,5cm;3cm;4cm;0cm;5cm;3cm" ' Ocultamos las 2 primeras
    End With
    
    If m_FiltrosAvanzados Is Nothing Then
        Set m_FiltrosAvanzados = New FiltrosSolicitud
        If rolUsuario = rol.Tecnico Then
            Me.cmdFiltrosAvanzados.Visible = False
            Me.cmdQuitarFiltros.Visible = False
            Me.chkSoloMisSolicitudes.value = True
            Me.chkSoloMisSolicitudes.Enabled = False
        Else
            Me.cmdFiltrosAvanzados.Visible = True
            Me.chkSoloMisSolicitudes.Enabled = True
        
        End If
    End If
    ' --- NUEVA LÓGICA PARA TAREAS PENDIENTES ---
    ' Si el formulario se abre con el argumento "PENDIENTES", marcamos la casilla.
    If Nz(Me.OpenArgs, "") = "PENDIENTES" Then
        Me.chkSoloPendientes.value = True
    End If
    Call RealizarBusqueda

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    Resume LimpiarYSalir
End Sub

Public Sub RealizarBusqueda()
    Dim solServicio As New SolicitudServicio
    Dim colSolicitudes As Scripting.Dictionary
    Dim unViewModel As SolicitudBusquedaViewModel
    Dim palabraClave As String
    Dim soloMis As Boolean, soloPend As Boolean
    Dim rowSourceString As String
    Dim m_ID As Variant
    
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    palabraClave = Nz(Me.txtPalabraClave.value, "")
    soloMis = Nz(Me.chkSoloMisSolicitudes.value, False)
    soloPend = Nz(Me.chkSoloPendientes.value, False)
    Set colSolicitudes = solServicio.getSolicitudesViewModel(palabraClave, m_FiltrosAvanzados, soloMis, soloPend)
    
    Me.lstResultados.rowSource = ""
    ' --- CABECERA MEJORADA: Añadimos "Estado" ---
    rowSourceString = "idSolicitud;idExpediente;Estado;Tipo;Cód. Solicitud;Nemotécnico;Cód. Exp;Resp. Técnico;Resp. Calidad"
    
    If Not colSolicitudes Is Nothing Then
        If colSolicitudes.count > 0 Then
            For Each m_ID In colSolicitudes.Keys
                Set unViewModel = colSolicitudes(m_ID)
                rowSourceString = rowSourceString & ";" & unViewModel.idSolicitud _
                    & ";" & unViewModel.idExpediente _
                    & ";" & unViewModel.EstadoNombre _
                    & ";" & unViewModel.tipoSolicitud _
                    & ";" & unViewModel.codigoSolicitud _
                    & ";" & Nz(unViewModel.Nemotecnico, "") _
                    & ";" & Nz(unViewModel.CodExp, "") _
                    & ";" & Nz(unViewModel.ResponsableTecnico, "N/A") _
                    & ";" & Nz(unViewModel.ResponsableCalidad, "N/A")
            Next
        End If
    End If
    
    Me.lstResultados.rowSource = rowSourceString
    Call ActualizarEstadoFiltrosVisual

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set solServicio = Nothing
    Set colSolicitudes = Nothing
    Set unViewModel = Nothing
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("RealizarBusqueda", Me)
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmBuscarSolicitudes.cls
' ACCIÓN: Añadir lógica de limpieza de filtro al evento Click.
' ==========================================================================

Private Sub chkSoloPendientes_Click()
    ' --- INICIO DE LA MODIFICACIÓN ---
    
    ' Si el usuario marca esta casilla, tiene prioridad sobre cualquier
    ' filtro de estado manual que se haya establecido previamente.
    If Me.chkSoloPendientes.value = True Then
        If Not m_FiltrosAvanzados Is Nothing Then
            m_FiltrosAvanzados.idEstado = "0" ' Reseteamos el filtro de estado
        End If
    End If
    
    ' --- FIN DE LA MODIFICACIÓN ---

    ' En cualquier caso, relanzamos la búsqueda.
    Call RealizarBusqueda
End Sub

Private Sub ActualizarEstadoFiltrosVisual()
    Dim hayFiltros As Boolean
    Dim hayFiltrosAvanzados As Boolean
    
    On Error GoTo Errores
    
    hayFiltrosAvanzados = False
    If Not m_FiltrosAvanzados Is Nothing Then
        hayFiltrosAvanzados = m_FiltrosAvanzados.HayFiltrosActivos()
    End If
    
    hayFiltros = (Nz(Me.txtPalabraClave.value, "") <> "" Or _
                  Nz(Me.chkSoloMisSolicitudes.value, False) = True Or _
                  hayFiltrosAvanzados)

    ' --- INICIO DE LA MODIFICACIÓN ---
    If hayFiltrosAvanzados Then
        ' REGLA DE ESTILO: Si hay filtros avanzados, el botón cambia a color Amarillo Corporativo.
        ' HEX: #E4C35C, Long de VBA: 6079460
        Me.cmdFiltrosAvanzados.BackColor = 6079460
        Me.cmdFiltrosAvanzados.ForeColor = 4666133 ' Yellow Very Dark (#473515) para el texto
    Else
        ' Si no hay filtros, vuelve a su color azul por defecto.
        ' HEX: #0066FF, Long de VBA: 16737792
        Me.cmdFiltrosAvanzados.BackColor = 16737792
        Me.cmdFiltrosAvanzados.ForeColor = 16777215 ' Blanco (#FFFFFF)
    End If
    
    ' La visibilidad del botón de quitar filtros no cambia.
    If hayFiltros Then
        If rolUsuario <> rol.Tecnico Then
            Me.cmdQuitarFiltros.Visible = True
        End If
    Else
        Me.cmdBuscar.SetFocus
        Me.cmdQuitarFiltros.Visible = False
    End If
    ' --- FIN DE LA MODIFICACIÓN ---
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("ActualizarEstadoFiltrosVisual", Me)
    Resume LimpiarYSalir
End Sub

' --- EVENTOS DE CONTROLES ---

Private Sub cmdBuscar_Click()
    Call RealizarBusqueda
End Sub



Private Sub txtPalabraClave_AfterUpdate()
    Call RealizarBusqueda
End Sub

Private Sub chkSoloMisSolicitudes_Click()
    Call RealizarBusqueda
End Sub

Private Sub cmdQuitarFiltros_Click()
    On Error GoTo Errores
    
    ' Limpiamos los filtros simples de la UI
    Me.txtPalabraClave.value = Null
    If Me.chkSoloMisSolicitudes.Enabled Then
        Me.chkSoloMisSolicitudes.value = False
    End If
    Me.chkSoloPendientes.value = False ' <-- AÑADIDO: También limpiar este filtro
    
    ' --- INICIO DE LA MODIFICACIÓN ---
    ' En lugar de crear un objeto nuevo, simplemente reseteamos las
    ' propiedades del objeto de filtros que ya existe en memoria.
    If Not m_FiltrosAvanzados Is Nothing Then
        m_FiltrosAvanzados.idResponsableTecnico = 0
        m_FiltrosAvanzados.idResponsableCalidad = 0
        m_FiltrosAvanzados.idSuministrador = 0
        m_FiltrosAvanzados.tipoSolicitud = ""
        m_FiltrosAvanzados.idEstado = "0"
    Else
        ' Si por alguna razón el objeto no existiera, lo creamos.
        ' Esto hace el código a prueba de fallos.
        Set m_FiltrosAvanzados = New FiltrosSolicitud
    End If
    ' --- FIN DE LA MODIFICACIÓN ---
    
    ' Volvemos a llamar a la búsqueda. Ahora usará el objeto de filtros reseteado.
    Call RealizarBusqueda
    
    ' La llamada a ActualizarEstadoFiltrosVisual ya está dentro de RealizarBusqueda,
    ' por lo que no es necesario llamarla aquí de nuevo.

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdQuitarFiltros_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdFiltrosAvanzados_Click()
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    DoCmd.OpenForm "frmFiltrosAvanzadosSolicitudes", , , , , acDialog
    
    On Error Resume Next
    If Forms("frmFiltrosAvanzadosSolicitudes").Tag = "OK" Then
        Set m_FiltrosAvanzados = Forms("frmFiltrosAvanzadosSolicitudes").filtros
        Call RealizarBusqueda
    End If
    On Error GoTo Errores
    
LimpiarYSalir:
    On Error Resume Next
    If FormularioAbierto("frmFiltrosAvanzadosSolicitudes") Then
        DoCmd.Close acForm, "frmFiltrosAvanzadosSolicitudes"
    End If
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("cmdFiltrosAvanzados_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)
    Call AbrirDetalleSeleccion
End Sub

Private Sub cmdVerDetalleSolicitud_Click()
    Call AbrirDetalleSeleccion
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub AbrirDetalleSeleccion()
    Dim idSol As String
    
    On Error GoTo Errores
    
    idSol = Nz(Me.lstResultados.Column(0), "")
    If idSol = "" Then
        MsgBox "Por favor, seleccione una solicitud de la lista.", vbInformation
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    If FormularioAbierto("frmGestionSolicitud") Then
        If Forms("frmGestionSolicitud").OpenArgs() <> idSol Then
            DoCmd.Close acForm, "frmGestionSolicitud", acSaveNo
        Else
            Forms("frmGestionSolicitud").SetFocus
            GoTo LimpiarYSalir
        End If
    End If
    DoCmd.OpenForm "frmGestionSolicitud", , , , , acWindowNormal, idSol
    Set m_frmGestion = Forms("frmGestionSolicitud")
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub

Errores:
    Call ManejadorErroresFormulario("AbrirDetalleSeleccion", Me)
    Resume LimpiarYSalir
End Sub

Public Property Get FiltrosActivos() As FiltrosSolicitud
    Set FiltrosActivos = m_FiltrosAvanzados
End Property



Private Sub m_frmGestion_SolicitudEliminada()
    ' Este procedimiento se ejecutará automáticamente cuando frmGestionSolicitud
    ' dispare el evento "SolicitudEliminada".
    On Error GoTo Errores
    
    Debug.Print "Evento 'SolicitudEliminada' capturado. Refrescando la búsqueda..."
    
    ' Simplemente volvemos a ejecutar la búsqueda para actualizar la lista.
    Call RealizarBusqueda
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("m_frmGestion_SolicitudEliminada", Me)
    Resume LimpiarYSalir
End Sub

Private Sub m_frmGestion_SolicitudModificada()
    ' Este procedimiento se ejecutará automáticamente cuando frmGestionSolicitud
    ' dispare el evento "SolicitudEliminada".
    On Error GoTo Errores
    
    Debug.Print "Evento 'SolicitudModificada' capturado. Refrescando la búsqueda..."
    
    ' Simplemente volvemos a ejecutar la búsqueda para actualizar la lista.
    Call RealizarBusqueda
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("m_frmGestion_SolicitudModificada", Me)
    Resume LimpiarYSalir
End Sub
