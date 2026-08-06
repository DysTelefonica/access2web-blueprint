Option Compare Database
Option Explicit


Public m_SolicitudSeleccionada As solicitud
Public m_ColSolicitudesTSeleccionadas As Scripting.Dictionary

Public Event SolicitudModificada()
Private m_TramitadorUltimo As String
Private m_EnumEstadoUltimo As EnumEstado
Private m_Error As String
Private Sub cmdReprogramarCorreos_Click()
    Dim m_SolOp As SolicitudOperaciones
    
    On Error GoTo errores
    
    ' 1. Confirmación de seguridad
    If MsgBox("Esta acción programará los correos automáticos pendientes para esta solicitud, " & _
              "siempre que la fecha de envío calculada sea HOY o en el FUTURO." & vbNewLine & vbNewLine & _
              "¿Desea continuar?", vbQuestion + vbYesNo, "Reprogramar Correos") = vbNo Then Exit Sub
              
    VBA.DoEvents
    DoCmd.Hourglass True
    m_Error = ""
    
    ' 2. Inicializar operación
    Set m_SolOp = New SolicitudOperaciones
    Set m_SolOp.solicitud = m_SolicitudSeleccionada ' Usamos la seleccionada en la lista
    
    ' 3. Ejecutar reprogramación
    m_SolOp.ReprogramarAvisosFuturos m_Error
    
    DoCmd.Hourglass False
    
    If m_Error <> "" Then
        MsgBox "Hubo un error al reprogramar: " & m_Error, vbExclamation
    Else
        MsgBox "Proceso finalizado. Se han programado los correos que cumplen los criterios de fecha.", vbInformation
    End If
    
    Exit Sub
    
errores:
    DoCmd.Hourglass False
    MsgBox "Error inesperado: " & Err.Description, vbCritical
End Sub
Public Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_ID As Variant
    Dim m_Expediente As Expediente
    Dim m_Suministrador As Suministrador
    Dim m_Nombre As String
    Dim m_Usuario As Usuario
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    
    Me.TIPO = "Todos"
    Set cmb = Me.IDEmpresaUsuario
    cmb.RowSource = ""
    cmb.AddItem "0;Todos"
    
    Set m_Col = m_ObjEntorno.Suministradores
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
             Set m_Suministrador = m_Col(m_ID)
            m_Nombre = m_Suministrador.Nemotecnico
            If m_Nombre = "" Then
                m_Nombre = m_Suministrador.Nombre
            End If
            cmb.AddItem m_Suministrador.IDSuministrador & ";" & m_Nombre
            
            Set m_Suministrador = Nothing
        Next
    End If
    
    cmb.value = "0"
    Set cmb = Me.IDExpediente
    cmb.RowSource = ""
    cmb.AddItem "0;Todos"
    Set m_Col = m_ObjEntorno.Expedientes
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Expediente = m_Col(m_ID)
            cmb.AddItem m_Expediente.IDExpediente & ";" & m_Expediente.TextoExpediente
            Set m_Expediente = Nothing
        Next
    End If
    cmb.value = "0"
    
    
    Set cmb = Me.Gestor
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.Coltramitadores
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    cmb.AddItem "0;Todos"
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Usuario = m_Col(m_ID)
            cmb.AddItem m_Usuario.id & ";" & m_Usuario.Nombre
            Set m_Usuario = Nothing
        Next
        If m_Col.Exists(CStr(m_ObjUsuarioConectado.id)) Then
            cmb.value = m_ObjUsuarioConectado.id
        Else
            cmb.value = "0"
        End If
    Else
        cmb.value = "0"
    End If
    
    Set cmb = Me.Estado
    cmb.RowSource = ""
    
    Set m_Col = m_ObjEntorno.ColEstadosParaTitulo
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    cmb.AddItem "0;Todos"
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            cmb.AddItem m_ID & ";" & m_Col(m_ID)
        Next
    End If
    cmb.value = CStr(EnumEstado.EnTramite)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerCombos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function Filtrar(Optional ByRef p_Error As String) As String

    Dim m_ColFiltrados As Scripting.Dictionary
    Dim m_Solicitud As solicitud
    Dim m_ID As Variant
    Dim m_Estado As EnumEstado
    Dim m_Tramitador As String
    Dim m_NemotecnicoBusqueda As String ' Variable auxiliar para búsqueda
    
    Dim lst As ListBox
    Dim i As Integer

    On Error GoTo errores

    p_Error = ""
    Set m_SolicitudSeleccionada = Nothing
    Set m_ColSolicitudesTSeleccionadas = Nothing
    Set m_ColFiltrados = Nothing
    Me.ComandoExportarExcel.Enabled = False
    If Me.ListaFiltrados.ListCount > 1 Then
        For i = 1 To Me.ListaFiltrados.ListCount - 1
            If Me.ListaFiltrados.Selected(i) = True Then
                Me.ListaFiltrados.Selected(i) = False
            End If
        Next
    End If
    
    Set lst = Me.ListaFiltrados
    lst.RowSource = "IDSolicitud;Tipo;DNI;Nombre;Estado"
    Me.lblListaDeResultados.Caption = "Registros filtrados (0)"

    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoVerLogs.Enabled = False
    
    If Nz(Me.Estado.Column(0), "") <> "" Then
        m_Estado = CLng(Me.Estado.Column(0))
    End If
    m_Tramitador = Nz(Me.Gestor.Column(1), "")
    With m_ObjEntorno
        If m_Estado <> 0 Then
            ' ... (Lógica de filtrado por estado sin cambios) ...
            If m_Estado = EnumEstado.PendienteEnvioExcel Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendientesEnvioExcel = Nothing
                End If
                Set .ColPendientesEnvioExcel = Nothing
                Set m_ColFiltrados = .ColPendientesEnvioExcel
                
            ElseIf m_Estado = EnumEstado.PendienteRecordatorioExcel1 Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendientesRecordatorioExcel1 = Nothing
                End If
                Set .ColPendientesRecordatorioExcel1 = Nothing
                Set m_ColFiltrados = .ColPendientesRecordatorioExcel1
            ElseIf m_Estado = EnumEstado.PendienteRecordatorioExcel2 Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendientesRecordatorioExcel2 = Nothing
                End If
                Set .ColPendientesRecordatorioExcel2 = Nothing
                Set m_ColFiltrados = .ColPendientesRecordatorioExcel2
            ElseIf m_Estado = EnumEstado.PendienteExcelRelleno Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendientesEnvioExcel = Nothing
                End If
                Set .ColPendientesEnvioExcel = Nothing
                Set m_ColFiltrados = .ColPendientesEnvioExcel
            ElseIf m_Estado = EnumEstado.PendienteRegistroMarga Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteRegistroMarga = Nothing
                End If
                Set .ColPendienteRegistroMarga = Nothing
                Set m_ColFiltrados = .ColPendienteRegistroMarga
            ElseIf m_Estado = EnumEstado.PendienteRecordatorioEnvioDPS1 Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteRecordatorioEnvioDPS1 = Nothing
                End If
                Set m_ColFiltrados = .ColPendienteRecordatorioEnvioDPS1
            ElseIf m_Estado = EnumEstado.PendienteRecordatorioEnvioDPS2 Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteRecordatorioEnvioDPS2 = Nothing
                End If
                Set .ColPendienteRecordatorioEnvioDPS2 = Nothing
                Set m_ColFiltrados = .ColPendienteRecordatorioEnvioDPS2
            ElseIf m_Estado = EnumEstado.PendienteEnvioDPS Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteEnvioDPS = Nothing
                End If
                Set .ColPendienteEnvioDPS = Nothing
                Set m_ColFiltrados = .ColPendienteEnvioDPS
            ElseIf m_Estado = EnumEstado.PendienteRegistroONS Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteRegistroONS = Nothing
                End If
                Set .ColPendienteRegistroONS = Nothing
                Set m_ColFiltrados = .ColPendienteRegistroONS
           
            ElseIf m_Estado = EnumEstado.Tramitado Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColTramitados = Nothing
                End If
                Set .ColTramitados = Nothing
                Set m_ColFiltrados = .ColTramitados
            ElseIf m_Estado = EnumEstado.Desestimado Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColDesestimados = Nothing
                End If
                Set .ColDesestimados = Nothing
                Set m_ColFiltrados = .ColDesestimados
            ElseIf m_Estado = EnumEstado.Cancelado Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColCancelados = Nothing
                End If
                Set .ColCancelados = Nothing
                Set m_ColFiltrados = .ColCancelados
            ElseIf m_Estado = EnumEstado.PendienteEnvioTraspasoONS Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColPendienteEnvioTraspasoONS = Nothing
                End If
                Set .ColPendienteEnvioTraspasoONS = Nothing
                Set m_ColFiltrados = .ColPendienteEnvioTraspasoONS
            ElseIf m_Estado = EnumEstado.EnTramite Then
                If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                    Set .ColEnTramite = Nothing
                End If
                Set .ColEnTramite = Nothing
                Set m_ColFiltrados = .ColEnTramite

            Else
               p_Error = "El tipo no está codificado"
               Err.Raise 1000
            End If
        Else
             If m_EnumEstadoUltimo <> m_Estado Or m_TramitadorUltimo <> m_Tramitador Then
                Set .ColSolicitudes = Nothing
            End If
            Set m_ColFiltrados = .ColSolicitudes
        End If
       
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_EnumEstadoUltimo = m_Estado
        m_TramitadorUltimo = Nz(Me.Gestor.Column(1), "")
    End With
    If m_ColFiltrados Is Nothing Then
        Exit Function
    End If
    Me.ComandoExportarExcel.Enabled = True
    
    For Each m_ID In m_ColFiltrados
        Set m_Solicitud = m_ColFiltrados(m_ID)
        With m_Solicitud
            If Nz(Me.TIPO, "") <> "" And Nz(Me.TIPO, "") <> "Todos" Then
                If .TIPO <> Nz(Me.TIPO, "") Then
                    GoTo siguiente
                End If
            End If
            If Nz(Me.Gestor.Column(1), "") <> "" And Nz(Me.Gestor.Column(1), "") <> "Todos" Then
                If Me.Gestor.Column(1) <> .Gestor Then
                    GoTo siguiente
                End If
            End If
            
            ' Filtro por Nombre (Case Insensitive)
            If Nz(Me.Nombre, "") <> "" Then
                If InStr(1, .NombreCompleto, Me.Nombre, vbTextCompare) = 0 Then
                    GoTo siguiente
                End If
            End If
            
            If IsNumeric(Me.IDEmpresaUsuario.Column(0)) Then
                If Me.IDEmpresaUsuario.Column(0) <> 0 Then
                    If .IDEmpresaUsuario <> Me.IDEmpresaUsuario.Column(0) Then
                        GoTo siguiente
                    End If
                End If
            End If
            
            If IsNumeric(Me.IDExpediente.Column(0)) Then
                If Me.IDExpediente.Column(0) <> 0 Then
                    If .IDExpediente <> Me.IDExpediente.Column(0) Then
                        GoTo siguiente
                    End If
                End If
            End If
            
            ' --- CORRECCIÓN AQUÍ ---
            ' Filtro por Nemotécnico
            If Nz(Me.Nemotecnico, "") <> "" Then
                m_NemotecnicoBusqueda = .Nemotecnico
                ' Si el campo directo está vacío, intentamos buscar en el objeto Expediente
                If m_NemotecnicoBusqueda = "" And Not .Expediente Is Nothing Then
                    m_NemotecnicoBusqueda = .Expediente.Nemotecnico
                End If
                ' Corregido: Buscar "Me.Nemotecnico" DENTRO de "m_NemotecnicoBusqueda"
                If InStr(1, m_NemotecnicoBusqueda, Me.Nemotecnico, vbTextCompare) = 0 Then
                    GoTo siguiente
                End If
            End If
            ' -----------------------

            If Nz(Me.DNI, "") <> "" Then
                If InStr(1, .DNI, Me.DNI) = 0 Then
                    GoTo siguiente
                End If
            End If
            
            lst.AddItem m_ID & ";" & .TIPO & ";" & .DNI & ";" & .NombreCompletoCapitalizado & ";" & .EstadoTitulo
        End With
siguiente:
        Set m_Solicitud = Nothing
    Next
    Me.lblListaDeResultados.Caption = "Registros filtrados (" & lst.ListCount - 1 & ")"
    If lst.ListCount = 1 Then
        Exit Function
    End If
    lst.Selected(1) = True
    ListaFiltrados_Click


    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub ComandoActualizar_Click()

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.TIPO = "Todos"
    Me.Gestor = m_ObjUsuarioConectado.id
    Me.IDEmpresaUsuario = 0
    Me.IDExpediente = 0
    Me.Estado = 0
    Me.Nombre = Null
    With m_ObjEntorno
        Set .ColPendientesEnvioExcel = Nothing
        Set .ColPendientesRecordatorioExcel1 = Nothing
        Set .ColPendientesRecordatorioExcel2 = Nothing
        Set .ColPendientesExcelRelleno = Nothing
        Set .ColPendienteRegistroMarga = Nothing
        Set .ColPendienteRecordatorioEnvioDPS1 = Nothing
        Set .ColPendienteRecordatorioEnvioDPS2 = Nothing
        Set .ColPendienteEnvioDPS = Nothing
        Set .ColPendienteRegistroONS = Nothing
        
        Set .ColTramitados = Nothing
        Set .ColDesestimados = Nothing
        Set .ColCancelados = Nothing
        Set .ColPendienteEnvioTraspasoONS = Nothing
        Set .ColEnTramite = Nothing
    End With
    
    Filtrar
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoActualizarEstado_Click()
     On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    ActualizarEstados
    Filtrar
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarEstado_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoAlta_Click()
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    m_Error = ""
   
    
    If FormularioAbierto("FormSolicitudAlta") Then
        DoCmd.Close acForm, "FormSolicitudAlta", acSaveNo
    End If
    
    DoCmd.OpenForm "FormSolicitudAlta"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEditar_Click()
    
    
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_SolicitudSeleccionada Is Nothing Then
        m_Error = "Seleccione una Solicitud de la lista"
        Err.Raise 1000
    End If
    Set m_ObjSolicitudActiva = m_SolicitudSeleccionada
    With m_SolicitudSeleccionada
        If .TIPO = "Alta" Then
            If FormularioAbierto("FormSolicitudAltaTraspaso") Then
                DoCmd.Close acForm, "FormSolicitudAltaTraspaso", acSaveNo
            End If
            DoCmd.OpenForm "FormSolicitudAltaTraspaso", OpenArgs:="Alta"
        ElseIf .TIPO = "Traspaso" Then
            If FormularioAbierto("FormSolicitudAltaTraspaso") Then
                DoCmd.Close acForm, "FormSolicitudAltaTraspaso", acSaveNo
            End If
            DoCmd.OpenForm "FormSolicitudAltaTraspaso", OpenArgs:="Traspaso"
        Else
            If .UsuarioHPS Is Nothing Then
                m_Error = "No se puede identificar el usuario HPS al ser una renovación"
                Err.Raise 1000
            End If
            If FormularioAbierto("FormSolicitudRenovacion") Then
                DoCmd.Close acForm, "FormSolicitudRenovacion", acSaveNo
            End If
            Set m_ObjUsuarioHPSActivo = .UsuarioHPS
            DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=.TIPO
        End If
    End With
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminar_Click()
     
    Dim m_SolOp As SolicitudOperaciones
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
   
    If m_SolicitudSeleccionada Is Nothing Then
        m_Error = "Seleccione una Solicitud de la lista"
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox("¿Desea realmente eliminar definitivamente la Solicitud Seleccionada?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Solicitud")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_SolOp = New SolicitudOperaciones
    With m_SolOp
        Set .solicitud = m_SolicitudSeleccionada
        .Eliminar p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    ComandoActualizar_Click
    Filtrar p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcel_Click()
    Dim m_ID As Variant
    Dim m_SOl As solicitud
    Dim lst As ListBox
    Dim i As Long
    Dim m_ColFiltrados As Scripting.Dictionary
    Dim m_URL As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Set lst = Me.ListaFiltrados
    If lst.ListCount = 1 Then
        Me.AllowEdits = True
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    For i = 1 To lst.ListCount - 1
        Set m_SOl = constructor.getSolicitud(p_IDSolicitud:=lst.Column(0, i), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_SOl Is Nothing Then
            If m_ColFiltrados Is Nothing Then
                Set m_ColFiltrados = New Scripting.Dictionary
            End If
            If Not m_ColFiltrados.Exists(CStr(lst.Column(0, i))) Then
                m_ColFiltrados.Add CStr(lst.Column(0, i)), m_SOl
            End If
        End If
    Next
    Avance "Obteniendo lista para informe...."
    m_URL = GenerarConsultaSolicitudes( _
                                            p_Col:=m_ColFiltrados, _
                                            p_Error:=m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    If Not fso.FileExists(m_URL) Then
        m_Error = "No se ha podido obtener la consulta"
        Err.Raise 1000
    End If
    Ejecutar Me.hwnd, "open", m_URL, "", "", 1
    
    Me.AllowEdits = True
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcel_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerLogs_Click()
        
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_SolicitudSeleccionada Is Nothing Then
        m_Error = "Seleccione una Solicitud de la lista"
        Err.Raise 1000
    End If
    InformeDelRegistro p_IDSolicitud:=m_SolicitudSeleccionada.IDSolicitud, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerLogs_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub DNI_AfterUpdate()
    Filtrar
End Sub

Private Sub Estado_AfterUpdate()
    Filtrar
End Sub


Private Sub Gestor_AfterUpdate()
    Filtrar
End Sub

Private Sub IDEmpresaUsuario_AfterUpdate()
    Filtrar
End Sub


Private Sub IDExpediente_AfterUpdate()
    Filtrar
End Sub



Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Me.AllowEdits = False
    If Me.ComandoEditar.Enabled = True Then
        ComandoEditar_Click
    End If

    Me.AllowEdits = True
End Sub




Private Sub Form_Open(Cancel As Integer)

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    Ajustar Me
    Me.Caption = m_TituloFormulario
    Me.AllowEdits = False
    
    Me.Caption = Me.lblTitulo.Caption
    EstablecerCombos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If EsAdministrador = EnumSiNo.Sí Then
        If m_ObjEntorno.UltimoCambio Is Nothing Then
            Me.lblUltimaModificacion.Visible = False
            Me.lblUltimaModificacion.Caption = "Aún sin registrar"
        Else
            Me.lblUltimaModificacion.Visible = True
            Me.lblUltimaModificacion.Caption = m_ObjEntorno.UltimoCambio.texto
        End If
    Else
        Me.lblUltimaModificacion.Visible = False
    End If
    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    'On Error Resume Next
    Me.AllowEdits = True
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaFiltrados_Click()

    Dim m_SOl As solicitud
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""


    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoVerLogs.Enabled = False
   
    Me.AllowEdits = False
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_SolicitudSeleccionada = Nothing
        Set m_ColSolicitudesTSeleccionadas = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_SolicitudSeleccionada Is Nothing Then
        Set m_SolicitudSeleccionada = constructor.getSolicitud(p_IDSolicitud:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If Me.ListaFiltrados.Column(0) <> m_SolicitudSeleccionada.IDSolicitud Then
            Set m_SolicitudSeleccionada = constructor.getSolicitud(p_IDSolicitud:=Me.ListaFiltrados.Column(0), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If

        End If
    End If
    
    
    If m_SolicitudSeleccionada Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
    Me.ComandoVerLogs.Enabled = True
    ' --- NUEVA LÓGICA PARA cmdReprogramarCorreos ---
    Dim blnActivarReprog As Boolean
    blnActivarReprog = False
    
    ' Solo si hay solicitud seleccionada
    If Not m_SolicitudSeleccionada Is Nothing Then
        ' 1. Variable Global debe ser SÍ
        If m_ObjEntorno.Configuracion.CorreosAutomaticos = "Sí" Then
            ' 2. Variable de la Solicitud debe ser SÍ
            If m_SolicitudSeleccionada.CorreosAutomaticos = "Sí" Then
                blnActivarReprog = True
            End If
        End If
    End If
    
    ' Aplicar estado al botón (asegúrate que el botón existe y se llama así)
    Me.cmdReprogramarCorreos.Enabled = blnActivarReprog
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
   
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub


Private Sub Nemotecnico_AfterUpdate()
    Filtrar
End Sub

Private Sub Nombre_AfterUpdate()
    Filtrar
End Sub

Private Sub TIPO_AfterUpdate()
    Filtrar
End Sub

