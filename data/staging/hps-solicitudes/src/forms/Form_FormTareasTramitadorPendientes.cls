Option Compare Database
Option Explicit


Public Event SolicitudModificada()
Public m_SolicitudSeleccionada As solicitud
Public m_TipoTareaSeleccionada As EnumEstado
Private m_Error As String

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
'
Public Function Filtrar(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_Solicitud As solicitud
    Dim m_ID As Variant
    Dim m_TipoTarea As EnumEstado
    Dim m_NombreCompleto As String
    Dim lst As ListBox


    On Error GoTo errores

    p_Error = ""
    Set m_SolicitudSeleccionada = Nothing
    Set lst = Me.ListaFiltrados
    lst.RowSource = "ID;Tipo;DNI;Nombre;Estado"
    Me.lblListaDeResultados.Caption = "Registros filtrados (0)"

    Me.ComandoEditar.Enabled = False

    If Not IsNumeric(Nz(Me.ListaTipoTareas.Column(0), "")) Then
        Exit Function
    End If
    m_TipoTarea = CLng(Me.ListaTipoTareas.Column(0))

    Avance "Obteniendo Listas"
    If m_TipoTarea = EnumEstado.PendienteEnvioExcel Then
        Set m_Col = constructor.getSolicitudesPendientesEnvioExcel(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
        
    ElseIf m_TipoTarea = EnumEstado.PendienteRecordatorioExcel1 Then
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel1(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteRecordatorioExcel2 Then
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel2(p_Tramitador:=Me.Gestor, p_Error:=p_Error)

    ElseIf m_TipoTarea = EnumEstado.PendienteExcelRelleno Then
        Set m_Col = constructor.getSolicitudesPendientesRellenoExcel(p_Tramitador:=Me.Gestor, p_Error:=p_Error)

    ElseIf m_TipoTarea = EnumEstado.PendienteRegistroMarga Then
        Set m_Col = constructor.getSolicitudesPendientesAltaMarga(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteRecordatorioEnvioDPS1 Then
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS1(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteRecordatorioEnvioDPS2 Then
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS2(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteEnvioDPS Then
        Set m_Col = constructor.getSolicitudesPendientesEnvioDPS(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteRegistroONS Then
        Set m_Col = constructor.getSolicitudesPendientesRegistroONS(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
   
    ElseIf m_TipoTarea = EnumEstado.PendienteEnvioTraspasoONS Then
        Set m_Col = constructor.getSolicitudesPendientesEnvioTraspasoONS(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    ElseIf m_TipoTarea = EnumEstado.PendienteRegistroEnHPS Then
        Set m_Col = constructor.getSolicitudesPendientesRegistroEnHPS(p_Tramitador:=Me.Gestor, p_Error:=p_Error)
    
    Else
        Exit Function
    End If
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If

    'lst.RowSource = "ID;Tipo;DNI;Nombre;Estado"
    For Each m_ID In m_Col
        If Nz(m_ID, "") = "" Then GoTo siguiente


        Set m_Solicitud = m_Col(m_ID)
        With m_Solicitud
            
            lst.AddItem m_ID & ";" & .TIPO & ";" & .DNI & ";" & .NombreCompletoCapitalizado & ";" & .EstadoCalculadoTitulo
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

    PintarIndicadores
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





Private Sub Gestor_BeforeUpdate(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    EstablecerDatos p_Error:=m_Error
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
        m_Error = "Al Gestor_AfterUpdate se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Me.AllowEdits = False
    If Me.ComandoEditar.Enabled = True Then
        ComandoEditar_Click
    End If

    Me.AllowEdits = True
End Sub

Public Sub ListaTipoTareas_Click()

    Dim lst As ListBox
    On Error Resume Next

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Application.Echo False
    Set lst = Me.ListaTipoTareas
    
    If Not IsNumeric(Nz(lst.Column(0), "")) Then
        m_TipoTareaSeleccionada = Empty
        Me.Explicacion = Null
        Me.ListaFiltrados.RowSource = "ID;Tipo;DNI;Nombre;Estado"
        Me.ComandoEditar.Enabled = False
        DoCmd.Hourglass False
        Application.Echo True
        Exit Sub
    End If
    m_TipoTareaSeleccionada = CLng(lst.Column(0))
    
    Me.Explicacion = m_ObjEntorno.colExplicacionTipoSeleccionTarea(CStr(m_TipoTareaSeleccionada))

    Filtrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Application.Echo True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaTipoTareas_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Application.Echo True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Function RellenarListas(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    Dim m_Tramitador As String
    Dim m_EsTodos As Boolean
    
    ' Variables para contadores
    Dim c_EnvioExcel As Long
    Dim c_Rec1Excel As Long
    Dim c_Rec2Excel As Long
    Dim c_RellenoExcel As Long
    Dim c_RegMarga As Long
    Dim c_Rec1DPS As Long
    Dim c_Rec2DPS As Long
    Dim c_EnvioDPS As Long
    Dim c_RegONS As Long
    Dim c_TraspasoONS As Long
    Dim c_FinHPS As Long
    
    On Error GoTo errores
    
    Set lst = Me.ListaTipoTareas
    lst.RowSource = ""
    
    m_Tramitador = Nz(Me.Gestor, "")
    If m_Tramitador = "" Or m_Tramitador = "Todos" Or m_Tramitador = "0" Then
        m_EsTodos = True
    Else
        m_EsTodos = False
    End If
    
    ' ==========================================================================================
    ' ESTRATEGIA DE CARGA:
    ' Si filtramos por "Todos", usamos las colecciones del Entorno (Caché en RAM = Rápido).
    ' Si filtramos por un Gestor, usamos el Constructor (Consulta BD = Necesario).
    ' ==========================================================================================
    
    ' 1. Pendiente Envio Excel
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendientesEnvioExcel Is Nothing Then c_EnvioExcel = m_ObjEntorno.ColPendientesEnvioExcel.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesEnvioExcel(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_EnvioExcel = m_Col.Count
    End If
    
    ' 2. Recordatorio Excel 1
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendientesRecordatorioExcel1 Is Nothing Then c_Rec1Excel = m_ObjEntorno.ColPendientesRecordatorioExcel1.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel1(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_Rec1Excel = m_Col.Count
    End If
    
    ' 3. Recordatorio Excel 2
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendientesRecordatorioExcel2 Is Nothing Then c_Rec2Excel = m_ObjEntorno.ColPendientesRecordatorioExcel2.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioExcel2(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_Rec2Excel = m_Col.Count
    End If

    ' 4. Pendiente Excel Relleno (sin caché en Entorno, consulta directa)
    Set m_Col = constructor.getSolicitudesPendientesRellenoExcel(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
    If Not m_Col Is Nothing Then c_RellenoExcel = m_Col.Count

    ' 5. Pendiente Registro Marga
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteRegistroMarga Is Nothing Then c_RegMarga = m_ObjEntorno.ColPendienteRegistroMarga.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesAltaMarga(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_RegMarga = m_Col.Count
    End If

    ' 6. Recordatorio Envio DPS 1
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteRecordatorioEnvioDPS1 Is Nothing Then c_Rec1DPS = m_ObjEntorno.ColPendienteRecordatorioEnvioDPS1.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS1(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_Rec1DPS = m_Col.Count
    End If
    
    ' 6. Recordatorio Envio DPS 2
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteRecordatorioEnvioDPS2 Is Nothing Then c_Rec2DPS = m_ObjEntorno.ColPendienteRecordatorioEnvioDPS2.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRecordatorioEnvioDPS2(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_Rec2DPS = m_Col.Count
    End If
    
    ' 7. Pendiente Envio DPS
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteEnvioDPS Is Nothing Then c_EnvioDPS = m_ObjEntorno.ColPendienteEnvioDPS.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesEnvioDPS(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_EnvioDPS = m_Col.Count
    End If
    
    ' 8. Pendiente Envio Traspaso ONS (Este no suele tener caché global directa, consultamos siempre o añadimos propiedad al entorno si es crítico)
    ' Asumimos consulta directa para evitar errores si falta la propiedad en Entorno
    Set m_Col = constructor.getSolicitudesPendientesEnvioTraspasoONS(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
    If Not m_Col Is Nothing Then c_TraspasoONS = m_Col.Count
    
    ' 9. Pendiente Registro ONS
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteRegistroONS Is Nothing Then c_RegONS = m_ObjEntorno.ColPendienteRegistroONS.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRegistroONS(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_RegONS = m_Col.Count
    End If
    
    ' 10. Pendiente Registro En HPS (Fin)
    If m_EsTodos Then
        If Not m_ObjEntorno.ColPendienteRegistroEnHPS Is Nothing Then c_FinHPS = m_ObjEntorno.ColPendienteRegistroEnHPS.Count
    Else
        Set m_Col = constructor.getSolicitudesPendientesRegistroEnHPS(p_Tramitador:=m_Tramitador, p_Error:=p_Error)
        If Not m_Col Is Nothing Then c_FinHPS = m_Col.Count
    End If
    
    ' RELLENAR EL LISTBOX
    With lst
        .AddItem EnumEstado.PendienteEnvioExcel & ";" & c_EnvioExcel & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteEnvioExcel)) & " (" & c_EnvioExcel & ")"
            
        .AddItem EnumEstado.PendienteRecordatorioExcel1 & ";" & c_Rec1Excel & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRecordatorioExcel1)) & " (" & c_Rec1Excel & ")"
            
        .AddItem EnumEstado.PendienteRecordatorioExcel2 & ";" & c_Rec2Excel & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRecordatorioExcel2)) & " (" & c_Rec2Excel & ")"

        .AddItem EnumEstado.PendienteExcelRelleno & ";" & c_RellenoExcel & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteExcelRelleno)) & " (" & c_RellenoExcel & ")"

        .AddItem EnumEstado.PendienteRegistroMarga & ";" & c_RegMarga & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRegistroMarga)) & " (" & c_RegMarga & ")"
            
        .AddItem EnumEstado.PendienteRecordatorioEnvioDPS1 & ";" & c_Rec1DPS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRecordatorioEnvioDPS1)) & " (" & c_Rec1DPS & ")"
            
        .AddItem EnumEstado.PendienteRecordatorioEnvioDPS2 & ";" & c_Rec2DPS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRecordatorioEnvioDPS2)) & " (" & c_Rec2DPS & ")"
            
        .AddItem EnumEstado.PendienteEnvioDPS & ";" & c_EnvioDPS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteEnvioDPS)) & " (" & c_EnvioDPS & ")"
            
        .AddItem EnumEstado.PendienteEnvioTraspasoONS & ";" & c_TraspasoONS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteEnvioTraspasoONS)) & " (" & c_TraspasoONS & ")"
        
        .AddItem EnumEstado.PendienteRegistroONS & ";" & c_RegONS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRegistroONS)) & " (" & c_RegONS & ")"
            
        .AddItem EnumEstado.PendienteRegistroEnHPS & ";" & c_FinHPS & ";" & _
            m_ObjEntorno.ColEstadosParaTitulo(CStr(EnumEstado.PendienteRegistroEnHPS)) & " (" & c_FinHPS & ")"
    End With
    
    ' Refrescar la selección del listado de tareas
    'Form_FormTareasTramitadorPendientes.m_TipoTareaSeleccionada = Empty
    'Form_FormTareasTramitadorPendientes.ListaTipoTareas_Click
    
    RellenarListas = "OK"
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    
    Dim m_NPendientesEnvioExcel As Integer
    Dim m_NPendientesRegistroMarga As Integer
    
    
    On Error GoTo errores
    RellenarListas p_Error:=p_Error
    If p_Error <> "" Then
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
    AvanceCerrar
    Set lst = Me.ListaTipoTareas
    If IsNumeric(Nz(lst.Column(1, 0), "")) Then
        m_NPendientesEnvioExcel = CInt(lst.Column(1, 0))
    End If
    If IsNumeric(Nz(lst.Column(1, 1), "")) Then
        m_NPendientesRegistroMarga = CInt(lst.Column(1, 1))
    End If
    
    
    
    
    If m_NPendientesEnvioExcel > 0 Then
        Me.ListaTipoTareas.Selected(0) = True
        ListaTipoTareas_Click
    ElseIf m_NPendientesRegistroMarga > 0 Then
        Me.ListaTipoTareas.Selected(1) = True
        ListaTipoTareas_Click
    
        
    End If
    
    AvanceCerrar
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

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
    RellenarComboTramitadores Me.Gestor, EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.Coltramitadores Is Nothing Then
        If m_ObjEntorno.Coltramitadores.Exists(CStr(m_ObjUsuarioConectado.id)) Then
            Me.Gestor = m_ObjUsuarioConectado.Nombre
        Else
            Me.Gestor = "Todos"
        End If
    Else
        Me.Gestor = "Todos"
    End If
    EstablecerDatos p_Error:=m_Error
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


    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""


    Me.ComandoEditar.Enabled = False

    Me.AllowEdits = False
    If Nz(Me.ListaFiltrados.Column(0), "") = "" Then
        Set m_SolicitudSeleccionada = Nothing
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

Private Sub m_Form_Alta()
    RaiseEvent SolicitudModificada
End Sub

Private Sub m_Form_Editado()
    RaiseEvent SolicitudModificada
End Sub