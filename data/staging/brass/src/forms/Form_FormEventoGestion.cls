Option Compare Database
Option Explicit
Private m_Error As String
Dim strIDEventoAnterior As String
Dim strIDEventoSeleccionado As String, strIDActividadSeleccionada As String, strIDMaterialSeleccionado As String, strEquipoSeleccionado As String, _
    strSeleccionadoFranqueado As String
Dim strURLParteEvento As String, strURLAnexoEvento As String, strURLAnexoActividad As String, strURLAnexoMaterial As String
Dim strBUIAnterior As String, strSubSistemaAnterior As String, strIDEquipoAnterior As String, _
    strFechaInicialAnterior As String, strFechaFinalAnterior As String, strTipoEventoAnterior As String, _
    strFranqueoAnterior As String, strFormularioQueAbre As String
Private m_ObjEventoSeleccionado As Evento
Private fso As New Scripting.FileSystemObject
Private WithEvents m_FormFranqueo As Form_FormEventoFranqueo
Private WithEvents m_FormEventoCopia As Form_FormEventoCopia
Public Sub LimpiarUltimos()
    strBUIAnterior = ""
    strIDEventoSeleccionado = ""
    strIDActividadSeleccionada = ""
    strSubSistemaAnterior = ""
    strIDEquipoAnterior = ""
    strFechaInicialAnterior = ""
    strFechaFinalAnterior = ""
    strSeleccionadoFranqueado = ""
    strTipoEventoAnterior = ""
    strFranqueoAnterior = ""
End Sub
Public Sub ComandoActualizarTodo_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarIDEvento_Click
    ComandoLimpiarFechas_Click
    ComandoLimpiarBUI_Click
    ComandoLimpiarEquipo_Click
    ComandoLimpiarSubSistema_Click
    ComandoLimpiarTipoEvento_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar actualizar las listas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub CargarEquipo()
    Dim strValor As String, p_BUI As String, strIDSubSistema As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    strValor = Nz(Me.ComboEquipoBusqueda, "")
    Me.ComboEquipoBusqueda = Null
    p_BUI = Nz(Me.ComboBUIBusqueda, "")
    strIDSubSistema = Nz(Me.ComboSubsistemaBusqueda, "")
    If p_BUI = "" Then
        m_SQL = "SELECT distinct Equipo FROM TbEquipos ORDER BY TbEquipos.Equipo; "
    Else
        If strIDSubSistema = "" Then
            m_SQL = "SELECT distinct Equipo FROM TbEquipos ORDER BY TbEquipos.Equipo; "
        Else
            m_SQL = "SELECT distinct Equipo FROM TbEquipos " & _
                    "WHERE BUI='" & p_BUI & "' AND SUBSISTEMA='" & strIDSubSistema & "' ORDER BY TbEquipos.Equipo;"
        End If
    End If
  
    Me.ComboEquipoBusqueda.RowSource = m_SQL
    Me.ComboEquipoBusqueda.Requery
    If strValor <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaEnCombo = Sí/No
        '       EstaEnCombo = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = EstaEnCombo(Me.ComboEquipoBusqueda, strValor)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método EstaEnCombo ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag = "Sí" Then
            Me.ComboEquipoBusqueda = strValor
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir cargar el combo del Equipo se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub CargarSubSistema()
    Dim strValor As String, p_BUI As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    Me.ComboSubsistemaBusqueda = Null
    Me.ComboEquipoBusqueda = Null
    p_BUI = Nz(Me.ComboBUIBusqueda, "")
    If p_BUI = "" Then
        m_SQL = "SELECT distinct SUBSISTEMA FROM TbSubsistemaBui ORDER BY SUBSISTEMA;"
    Else
        m_SQL = "SELECT distinct SUBSISTEMA FROM TbSubsistemaBui " & _
                "WHERE BUI='" & p_BUI & "' ORDER BY SUBSISTEMA;"
    End If
  
    Me.ComboSubsistemaBusqueda.RowSource = m_SQL
    Me.ComboSubsistemaBusqueda.Requery
    If strValor <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaEnCombo = Sí/No
        '       EstaEnCombo = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = EstaEnCombo(Me.ComboSubsistemaBusqueda, strValor)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método EstaEnCombo ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag = "Sí" Then
            Me.ComboSubsistemaBusqueda = strValor
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir cargar el combo del subsistema se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAltaActividadesPorLotes_Click()
    
    On Error GoTo errores
    m_Error = ""
    If m_ObjEventoSeleccionado Is Nothing Then
        If Nz(Me.ListaEventos.Column(1), "") = "" Then
             m_Error = "Ha de seleccionar el Evento"
             Err.Raise 1000
        End If
    End If
    Set m_ObjEventoActivo = m_ObjEventoSeleccionado
    If FormularioAbierto("FormActividadDeEventoPorLotes") Then
        DoCmd.Close acForm, "FormActividadDeEventoPorLotes", acSaveNo
    End If
    DoCmd.OpenForm "FormActividadDeEventoPorLotes"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaActividadesPorLotes_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAltaMasivaPorNS_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
    If FormularioAbierto("FormEventosAltaMasivaPorNS") Then
        DoCmd.Close acForm, "FormEventosAltaMasivaPorNS", acSaveNo
    End If
    DoCmd.OpenForm "FormEventosAltaMasivaPorNS"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaMasivaPorNS_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarMaterial_Click()
    Dim strIDMaterialSeleccionado As String, strFranqueado As String, strEstaFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDMaterialSeleccionado = Nz(Me.ListaMateriales.Column(0), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDMaterialSeleccionado = "" Then
        m_Error = "Seleccione un material de la lista"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
       
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Facturado"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar el material seleccionado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Material")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       EliminarMaterialEnActividad = Descriptivo
    '       EliminarMaterialEnActividad = "#ERR" & "|" & m_Error
    '-----------------------------------------------
    flag = EliminarMaterialEnActividad(strIDMaterialSeleccionado, "Sí", , strFranqueado, strEstaFacturado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método EliminarMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.cmdSalir.SetFocus
    ActualizarMaterial
    strIDMaterialSeleccionado = ""
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar eliminar el material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoLimpiarBUI_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    LimpiarUltimos
    If Nz(Me.ComboBUIBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboBUIBusqueda = Null
    CargarSubSistema
    CargarEquipo
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarEquipo_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.ComboEquipoBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboEquipoBusqueda = Null
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarFechas_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.FechaInicialBusqueda, "") = "" And Nz(Me.FechaFinalBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.FechaInicialBusqueda = Null
    Me.FechaFinalBusqueda = Null
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarFranqueado_Click()
     On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.ComboFranqueadoBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboFranqueadoBusqueda = Null
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarIDEvento_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.ComboIDEventoBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboIDEventoBusqueda = Null
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarSubSistema_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.ComboSubsistemaBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboSubsistemaBusqueda = Null
    CargarEquipo
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarTipoEvento_Click()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    LimpiarUltimos
    If Nz(Me.ComboTipoEventoBusqueda, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComboTipoEventoBusqueda = Null
    ActualizarEventos
    ActualizarActividades
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoParaInforme_Click()
    
    
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    m_Error = ""
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    If strIDEventoSeleccionado = "" Then
        m_Error = "Se ha de seleccionar un evento de la lista superior"
        Err.Raise 1000
    End If
    If Not FormularioAbierto("FormInformesDesdeLista") Then
        m_Error = "No aparece abierto el formulario para generar el informe para el RAC"
        Err.Raise 1000
    End If
    If m_ColEventosParaInformeDeRAC Is Nothing Then
        Set m_ColEventosParaInformeDeRAC = New Scripting.Dictionary
    Else
        If m_ColEventosParaInformeDeRAC.Exists(strIDEventoSeleccionado) Then
            m_Error = "El Evento " & strIDEventoSeleccionado & " ya está incluido en la lista para informe"
            Err.Raise 1000
        End If
    End If
   
    Set m_ObjEvento = Constructor.getEvento(strIDEventoSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ColEventosParaInformeDeRAC.Add strIDEventoSeleccionado, m_ObjEvento
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Form_FormInformesDesdeLista.RellenarLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Forms("FormInformesDesdeLista").SetFocus
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al llevarme el Evento para informe se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Public Sub ComboBUIBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboBUIBusqueda, "")
    If strValor <> strBUIAnterior Then
        ActualizarEventos
        ActualizarActividades
        ActualizarMaterial
        CargarSubSistema
    End If
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboBUIBusqueda_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComboFranqueadoBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboFranqueadoBusqueda, "")
    If strValor <> strFranqueoAnterior Then
       ActualizarEventos
    End If
    strFranqueoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboFranqueadoBusqueda_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ComboEquipoBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipoBusqueda, "")
    If strValor <> strIDEquipoAnterior Then
       ActualizarEventos
    End If
    m_Error = ""
    strIDEquipoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Public Sub ComboIDEventoBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboIDEventoBusqueda, "")
    If strValor <> strIDEventoAnterior Then
        ActualizarEventos
    End If
    strIDEventoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboIDEventoBusqueda_Enter()
    On Error Resume Next
    Me.ComboIDEventoBusqueda.Requery
End Sub



Private Sub ComboSubsistemaBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaAnterior Then
        ActualizarEventos
        CargarEquipo
    End If
    strSubSistemaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComboTipoEventoBusqueda_AfterUpdate()
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.ComboTipoEventoBusqueda, "")
    If strValor <> strTipoEventoAnterior Then
       ActualizarEventos
    End If
    strTipoEventoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub FechaFinalBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.FechaFinalBusqueda, "")
    If strValor <> strFechaFinalAnterior Then
        ActualizarEventos
    End If
    strFechaFinalAnterior = strValor
     VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub FechaInicialBusqueda_Exit(Cancel As Integer)
    Dim strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strValor = Nz(Me.FechaInicialBusqueda, "")
    If strValor <> strFechaInicialAnterior Then
        ActualizarEventos
    End If
    strFechaInicialAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ImagenAnexoActividad_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strURLAnexoActividad = "" Then
        m_Error = "En estos momentos el anexo seleccionado no es alcanzable por la herramienta. Revise que ubicación de la documentación remota esté bien establecida"
        Err.Raise 1000
    End If
    If InStr(1, strURLAnexoActividad, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexoActividad, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        ComandoAdjuntosActividades_Click
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ImagenAnexoEvento_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strURLAnexoEvento = "" Then
        m_Error = "En estos momentos el anexo seleccionado no es alcanzable por la herramienta. Revise que ubicación de la documentación remota esté bien establecida"
        Err.Raise 1000
    End If
    If InStr(1, strURLAnexoEvento, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexoEvento, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        ComandoAdjuntosEventos_Click
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ImagenAnexoMaterial_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strURLAnexoMaterial = "" Then
        m_Error = "En estos momentos el anexo seleccionado no es alcanzable por la herramienta. Revise que ubicación de la documentación remota esté bien establecida"
        Err.Raise 1000
    End If
    If InStr(1, strURLAnexoMaterial, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & m_Error
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexoMaterial, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        ComandoAdjuntosMateriales_Click
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaEventos_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoEditarEvento.Enabled Then
        
        ComandoEditarEvento_Click
        
    End If
    Application.Echo True
End Sub

Public Sub ListaMateriales_Click()
    Dim m_MaterialSeleccionado As Material
    Dim strEstaFacturado As String, strFranqueado As String, strMaterialSeleccionado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    strIDMaterialSeleccionado = Nz(Me.ListaMateriales.Column(0), "")
    strMaterialSeleccionado = Nz(Me.ListaMateriales.Column(1), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDActividadSeleccionada <> "" Then
        Me.ComandoAltaMaterial.Enabled = True
    Else
        Me.ComandoAltaMaterial.Enabled = False
    End If
    Me.ComandoDetalleMaterial.Enabled = False
    Me.ImagenAnexoMaterial.Visible = False
    Me.ComandoAdjuntosMateriales.Enabled = False
    
    strURLAnexoMaterial = ""
    
    
    If strIDMaterialSeleccionado = "" Then
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Set m_MaterialSeleccionado = Constructor.getMaterial(strIDMaterialSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Then
        Me.ComandoEliminarMaterial.Enabled = False
    Else
        Me.ComandoEliminarMaterial.Enabled = True
    End If
    
    Me.ComandoDetalleMaterial.Enabled = True
    Me.ComandoAdjuntosMateriales.Enabled = True
    
    strURLAnexoMaterial = m_MaterialSeleccionado.URLAnexos
    
    If InStr(1, strURLAnexoMaterial, ";") <> 0 Then
        Me.ImagenAnexoMaterial.Visible = True
        Me.ImagenAnexoMaterial.BorderWidth = 0
        Me.ImagenAnexoMaterial.BorderColor = RGB(255, 255, 255)
        
    Else
        If strURLAnexoMaterial = "" Then
            Me.ImagenAnexoMaterial.Visible = False
            
        Else
            Me.ImagenAnexoMaterial.Visible = True
            If fso.FileExists(strURLAnexoMaterial) Then
                Me.ImagenAnexoMaterial.BorderStyle = 0
                Me.ImagenAnexoMaterial.BorderWidth = 0
                
            Else
                Me.ImagenAnexoMaterial.BorderStyle = 1
                Me.ImagenAnexoMaterial.BorderWidth = 2
                Me.ImagenAnexoMaterial.BorderColor = RGB(255, 0, 0)
               
            End If
        End If
    End If
    Me.ComandoAdjuntosMateriales.Visible = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar hacer click sobre un material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntosActividades_Click()

    Dim m_Actividad As Actividad
    Dim strEscritura As String, strEstaFacturado As String, strFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strIDActividadSeleccionada = "" Then
        strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    End If
    If strIDActividadSeleccionada = "" Then
        m_Error = "Se ha de Seleccionar una actividad primero"
        Err.Raise 1000
    End If
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        strFranqueado = "Sí"
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        strEstaFacturado = "Sí"
    End If
    If strFranqueado = "Sí" Or strEstaFacturado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    Set m_Actividad = Constructor.getActividad(strIDActividadSeleccionada, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    
    strFormularioQueAbre = ""
    Set m_ObjEntidadParaAnexoActiva = m_Actividad
    
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntosEventos_Click()
    Dim m_Evento As Evento
    Dim strEscritura As String, strFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strIDEventoSeleccionado = "" Then
        strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Se ha de SeleccionarEvento el evento primero"
        Err.Raise 1000
    End If
    
    strFacturado = Nz(Me.ListaEventos.Column(0), "")
    If strFacturado = "*" Then
        strFacturado = "Sí"
    Else
        strFacturado = "No"
    End If
    If strFacturado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    Set m_Evento = Constructor.getEvento(strIDEventoSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = ""
    Set m_ObjEntidadParaAnexoActiva = m_Evento
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntosMateriales_Click()
    Dim m_Material As Material
    Dim strIDMaterial As String, strEstaFranqueado As String, strFacturado As String, strEscritura As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDMaterialSeleccionado = Nz(Me.ListaMateriales.Column(0), "")
    strFacturado = Nz(Me.ListaEventos.Column(0), "")
    strEstaFranqueado = Nz(Me.ListaEventos.Column(7), "")
    If strFacturado = "*" Then
        strFacturado = "Sí"
    Else
        strFacturado = "No"
    End If
    If strFacturado = "Sí" Or strEstaFranqueado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    If strIDMaterialSeleccionado = "" Then
        m_Error = "Se ha de Seleccionar un material primero"
        Err.Raise 1000
    End If
    Set m_Material = Constructor.getMaterial(strIDMaterialSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    strFormularioQueAbre = ""
    Set m_ObjEntidadParaAnexoActiva = m_Material
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ComandoAltaActividad_Click()
    Dim strFranqueado As String, strEstaFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    
    If strFranqueado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Franqueado"
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Facturado"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormActividadAlta") Then
        DoCmd.Close acForm, "FormActividadAlta", acSaveNo
    End If
    'strArgumento=strIDEvento & "|" & strIDActividadOriginal & "|" & strEventoFranqueado
    'strArgumento=""
    DoCmd.OpenForm "FormActividadAlta", , , , , , strIDEventoSeleccionado & "|" & "|" & strEstaFacturado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar dar de alta una actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAltaActividadSinEvento_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormActividadAlta") Then
        DoCmd.Close acForm, "FormActividadAlta", acSaveNo
    End If
    'strArgumento=strIDEvento & "|" & strIDActividadOriginal & "|" & strEventoFranqueado
    'strArgumento=""
    DoCmd.OpenForm "FormActividadAlta"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar dar de alta una actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoAltaEvento_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormEventoAlta"
End Sub
Private Sub ComandoAltaMaterial_Click()
    Dim p_BUI As String, strSUBSISTEMA As String, strEquipo As String, strIDEquipo As String, strIDActividad As String, strIDEvento As String, _
        strEstaFacturado As String, strTipoAlta As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDEvento = Nz(Me.ListaEventos.Column(1), "")
    strSeleccionadoFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strSeleccionadoFranqueado <> "Sí" And strSeleccionadoFranqueado <> "No" Then
        m_Error = "No se ha podido determinar si el evento está o no franqueado"
        Err.Raise 1000
    End If
    If strSeleccionadoFranqueado = "Sí" Then
        m_Error = "El evento del que dependería el material ya está franqueado, ha de quitar el franqueo para introducirla"
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento del que dependería el material ya está facturado."
        Err.Raise 1000
    End If
    If strIDEvento = "" Then
        m_Error = "No se sabe de qué evento se deriva"
        Err.Raise 1000
    End If
    strIDActividad = Nz(Me.ListaActividades.Column(1), "")
    If strIDActividad = "" Then
        m_Error = "No se sabe de qué actividad se deriva"
        Err.Raise 1000
    End If
    p_BUI = Nz(Me.ListaEventos.Column(2), "")
    If p_BUI = "" Then
        m_Error = "No se sabe de qué BUI se deriva"
        Err.Raise 1000
    End If
    strSUBSISTEMA = Nz(Me.ListaEventos.Column(3), "")
    If strSUBSISTEMA = "" Then
        m_Error = "No se sabe de qué SubSistema se deriva"
        Err.Raise 1000
    End If
    strEquipo = Nz(Me.ListaEventos.Column(4), "")
    If strEquipo = "" Then
        m_Error = "No se sabe de qué Equipo se deriva"
        Err.Raise 1000
    End If
    strIDEquipo = Nz(Me.ListaEventos.Column(8), "")
    If strIDEquipo = "" Then
        m_Error = "No se sabe de qué Equipo se deriva"
        Err.Raise 1000
    End If
'    Repara Tel.;Repara Otro;Repuesto
'    strTipoAlta = Nz(ComboMaterialTipoAlta, "")
'    If strTipoAlta = "" Then
'        m_Error = "Se ha de seleccionar el tipo de alta de material"
'        Err.Raise 1000
'    End If
'
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    If FormularioAbierto("FormMaterialEleccion") Then
        DoCmd.Close acForm, "FormMaterialEleccion", acSaveNo
    End If
    DoCmd.OpenForm "FormMaterialEleccion", , , , , , strIDEvento & "|" & strIDActividad
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar dar de alta un material de la actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Function FiltrarEventos( _
                                Optional strIDEvento As String, _
                                Optional p_BUI As String, _
                                Optional strSUBSISTEMA As String, _
                                Optional strEquipo As String, _
                                Optional strFechaInicial As String, _
                                Optional strFechaFinal As String, _
                                Optional strTipoEvento As String, _
                                Optional strFranqueado As String, _
                                Optional ByRef p_Error As String _
                                ) As String
   
    Dim strWhereIDEvento As String, strWhereFechas As String, strWhereBUI As String, strWhereSubSistema As String, strWhereEquipo As String, _
        strWhereTipoEvento As String, strWhereFranqueo As String, strSQLInicial As String, strSQLFinal As String, strParteWhere As String
    Dim m_ObjGestorInforme As GestorInforme, m_ColCampos As Scripting.Dictionary
    On Error GoTo errores
    
    m_Error = ""
    p_Error = ""
    Set m_ObjGestorInforme = New GestorInforme
    strSQLInicial = "SELECT IIf(IsNull([IDFactura]),'','*') AS F, TbEventos.IDEvento, TbEventos.BUI, TbEventos.SUBSISTEMA, " & _
                "TbEquipos.Equipo, TbEventos.FECHAALTAEVENTO AS FALTA, TbEventos.DESCRIPCION AS DESCRIPCIÓN, " & _
                "IIf([Franqueado],'Sí','No') AS Franqueo, TbEventos.IDEquipo " & _
                "FROM (TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) " & _
                "LEFT JOIN TbFacturaEventosInvolucrados ON TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento "
    

    
    If strIDEvento = "" And p_BUI = "" And strSUBSISTEMA = "" And strEquipo = "" And strFechaInicial = "" And _
        strFechaFinal = "" And strTipoEvento = "" And strFranqueado = "" Then
        strParteWhere = "WHERE (((TbEventos.IDEvento)='-1'));"
        strSQLFinal = strSQLInicial & strParteWhere
        FiltrarEventos = strSQLFinal
        Set m_ObjGestorInforme = Nothing
        Exit Function
    End If
    If strIDEvento = "" Then
        strWhereIDEvento = "((TbEventos.IDEvento) Like '*' Or (TbEventos.IDEvento) Is Null) "
    Else
        strWhereIDEvento = "((TbEventos.IDEvento)='" & strIDEvento & "')"
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("IDEvento") Then
            m_ColCampos.Add "IDEvento", "IDEvento"
        End If
    End If
    If p_BUI = "" Then
        strWhereBUI = "((TbEventos.BUI) Like '*' Or (TbEventos.BUI) Is Null) "
    Else
        strWhereBUI = "((TbEventos.BUI)='" & p_BUI & "')"
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("BUI") Then
            m_ColCampos.Add "BUI", "BUI"
        End If
    End If
    If strSUBSISTEMA = "" Then
        strWhereSubSistema = "((TbEventos.SUBSISTEMA) Like '*' Or (TbEventos.SUBSISTEMA) Is Null) "
    Else
        strWhereSubSistema = "((TbEventos.SUBSISTEMA)='" & strSUBSISTEMA & "')"
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("SUBSISTEMA") Then
            m_ColCampos.Add "SUBSISTEMA", "SUBSISTEMA"
        End If
    End If
    If strEquipo = "" Then
        strWhereEquipo = "((TbEquipos.Equipo) Like '*' Or (TbEquipos.Equipo) Is Null) "
    Else
        strWhereEquipo = "((TbEquipos.Equipo)='" & strEquipo & "')"
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("Equipo") Then
            m_ColCampos.Add "Equipo", "Equipo"
        End If
    End If
    If strTipoEvento = "" Then
        strWhereTipoEvento = "((TbEventos.TIPOEVENTO) Like '*' Or (TbEventos.TIPOEVENTO) Is Null) "
    Else
        strWhereTipoEvento = "((TbEventos.TIPOEVENTO)='" & strTipoEvento & "')"
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("TIPOEVENTO") Then
            m_ColCampos.Add "TIPOEVENTO", "TIPOEVENTO"
        End If
    End If
    If strFranqueado = "" Then
        strWhereFranqueo = "((TbEventos.Franqueado)=True Or (TbEventos.Franqueado)=False) "
    Else
        If strFranqueado = "Sí" Then
           strWhereFranqueo = "((TbEventos.Franqueado)=True)"
        Else
           strWhereFranqueo = "((TbEventos.Franqueado)=False)"
        End If
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("Franqueado") Then
            m_ColCampos.Add "Franqueado", "Franqueado"
        End If
                
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicial = "01/01/1900"
    Else
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("FECHAALTAEVENTO") Then
            m_ColCampos.Add "FECHAALTAEVENTO", "FECHAALTAEVENTO"
        End If
    End If
    If Not IsDate(strFechaFinal) Then
        strFechaFinal = "01/01/2900"
    Else
        If m_ColCampos Is Nothing Then
            Set m_ColCampos = New Scripting.Dictionary
            m_ColCampos.CompareMode = TextCompare
        End If
        If Not m_ColCampos.Exists("FECHAALTAEVENTO") Then
            m_ColCampos.Add "FECHAALTAEVENTO", "FECHAALTAEVENTO"
        End If
    End If
    m_ObjGestorInforme.Registrar m_ColCampos
    strWhereFechas = "((TbEventos.FECHAALTAEVENTO) Between #" & Format(strFechaInicial, "mm/dd/yyyy") & "# And #" & Format(strFechaFinal, "mm/dd/yyyy") & "#)"
    strParteWhere = "WHERE (" & _
                    strWhereIDEvento & " AND " & _
                    strWhereBUI & " AND " & _
                    strWhereSubSistema & " AND " & _
                    strWhereEquipo & " AND " & _
                    strWhereTipoEvento & " AND " & _
                    strWhereFranqueo & " AND " & _
                    strWhereFechas & ") " & _
                    "ORDER BY TbEventos.FECHAALTAEVENTO DESC;"

    strSQLFinal = strSQLInicial & strParteWhere
    FiltrarEventos = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarEventos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function FiltrarActividades( _
                                    Optional strIDEvento As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim strWhereIDEvento As String, strSQLInicial As String, strSQLFinal As String, strParteWhere As String
    On Error GoTo errores
    
    m_Error = ""
    p_Error = ""
    strSQLInicial = "SELECT DISTINCT IIf(IsNull([TbMaterial].[IDMaterial]),'','*') AS M, TbActividades.IDActividad AS ID, " & _
                    "TbActividades.ALIASTECNICO AS TÉCNICO, TbActividades.FechaAlta AS FALTA, " & _
                    "TbActividades.HorasLaborables AS HLab, TbActividades.DESCRIPCION AS DESCRIPCIÓN, " & _
                    "TbFacturaActividadesInvolucradas.IDFactura " & _
                    "FROM (TbActividades LEFT JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad) " & _
                    "LEFT JOIN TbFacturaActividadesInvolucradas ON " & _
                    "TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad "
    If strIDEvento = "" Then
        strParteWhere = "WHERE (((TbActividades.IDEvento)='-1'));"
        strSQLFinal = strSQLInicial & strParteWhere
        FiltrarActividades = strSQLFinal
        Exit Function
    End If
    strWhereIDEvento = "((TbActividades.IDEvento)='" & strIDEvento & "')"
    strParteWhere = "WHERE (" & _
                    strWhereIDEvento & ") " & _
                    "ORDER BY TbActividades.FechaAlta DESC;"

    strSQLFinal = strSQLInicial & strParteWhere
    FiltrarActividades = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarActividades ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function FiltrarMateriales( _
                                    Optional strIDActividad As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim strWhereIDActividad As String, strSQLInicial As String, strSQLFinal As String, strParteWhere As String
    On Error GoTo errores
    m_Error = ""
    p_Error = ""
    
    strSQLInicial = "SELECT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.PN, TbMaterial.NS, " & _
                    "TbMaterial.EsReparacion, TbFacturaMaterialesInvolucrados.IDFactura " & _
                    "FROM TbMaterial LEFT JOIN TbFacturaMaterialesInvolucrados ON " & _
                    "TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial "
    If strIDActividad = "" Then
        strParteWhere = "WHERE (((TbMaterial.IDActividad)='-1'));"
        strSQLFinal = strSQLInicial & strParteWhere
        FiltrarMateriales = strSQLFinal
        Exit Function
    End If
    strWhereIDActividad = "((TbMaterial.IDActividad)='" & strIDActividad & "')"
    strParteWhere = "WHERE (" & _
                    strWhereIDActividad & ");"
    strSQLFinal = strSQLInicial & strParteWhere
    FiltrarMateriales = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarMateriales ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub ComandoCopiarActividad_Click()
    Dim strIDActividadSeleccionada As String, strFranqueado As String, strEstaFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    If strIDActividadSeleccionada = "" Then
        m_Error = "Seleccione la actividad que pretende copiar"
        Err.Raise 1000
    End If
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
   
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    
    If strFranqueado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Franqueado"
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Facturado"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '       FormularioAbierto = true
    '-------------------------------------------------------------------
    If FormularioAbierto("FormActividadAlta") Then
        DoCmd.Close acForm, "FormActividadAlta", acSaveNo
    End If
    'strArgumento=strIDEvento & "|" & strIDActividadOriginal & "|" & strEventoFranqueado
    'strArgumento=""
    DoCmd.OpenForm "FormActividadAlta", , , , , , strIDEventoSeleccionado & "|" & strIDActividadSeleccionada & "|" & strEstaFacturado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar copiar la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoCopiarEvento_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(Nz(Me.ListaEventos.Column(1), ""))
        If m_ObjEventoSeleccionado Is Nothing Then
            m_Error = "Seleccione un evento de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjEventoActivo = m_ObjEventoSeleccionado
    If FormularioAbierto("FormEventoCopia") Then
        DoCmd.Close acForm, "FormEventoCopia", acSaveNo
    End If
    DoCmd.OpenForm "FormEventoCopia"
    If FormularioAbierto("FormEventoCopia") Then
        Set m_FormEventoCopia = Forms("FormEventoCopia")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCopiarEvento_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoEditarActividad_Click()
    Dim strIDActividadSeleccionada As String, strEstaFacturado As String, strFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    strEstaFacturado = Nz(Me.ListaActividades.Column(7), "")
    If strEstaFacturado <> "" Then
         strEstaFacturado = "Sí"
    Else
         strEstaFacturado = "No"
    End If
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '       FormularioAbierto = true
    '-------------------------------------------------------------------
    If FormularioAbierto("FormActividadEdicion") Then
        DoCmd.Close acForm, "FormActividadEdicion", acSaveNo
    End If
    'strArgumento=strIDActividad & "|"  & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexoActividad
    DoCmd.OpenForm "FormActividadEdicion", , , , , , strIDActividadSeleccionada & "|" & strFranqueado & "|" & _
                    strEstaFacturado & "|" & strURLAnexoActividad
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar editar la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoEditarEvento_Click()
    Dim strEstaFacturado As String, strFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(Nz(Me.ListaEventos.Column(1), ""))
        If m_ObjEventoSeleccionado Is Nothing Then
            m_Error = "Seleccione un evento de la lista"
            Err.Raise 1000
        End If
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '       FormularioAbierto = true
    '-----------------------------------------------------------------
    If FormularioAbierto("FormEventoEdicion") Then
        DoCmd.Close acForm, "FormEventoEdicion", acSaveNo
    End If
    Set m_ObjEventoActivo = m_ObjEventoSeleccionado
    'strArgumento=strIDEvento & "|" &  strEstaFacturado & "|" & strFranqueado
    DoCmd.OpenForm "FormEventoEdicion", , , , , , strIDEventoSeleccionado & "|" & strEstaFacturado & "|" & strFranqueado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar editar un evento se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoDetalleMaterial_Click()
    Dim strEsReparacion As String
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDMaterialSeleccionado = Nz(Me.ListaMateriales.Column(0), "")
    strEsReparacion = Nz(Me.ListaMateriales.Column(4), "")
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
        
    If strIDMaterialSeleccionado = "" Then
        m_Error = "Seleccione el material que pretende editar"
        Err.Raise 1000
    End If
    
    If strIDActividadSeleccionada = "" Then
        m_Error = "No se sabe de qué actividad se deriva"
        Err.Raise 1000
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        m_Error = "No se ha podido determinar si es un material para reparar o no."
        Err.Raise 1000
    End If
    If FormularioAbierto("FormMaterialesGestion") Then
        DoCmd.Close acForm, "FormMaterialesGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormMaterialesGestion"
    Forms("FormMaterialesGestion").Controls("ComboIDMaterialBusqueda") = strIDMaterialSeleccionado
    Form_FormMaterialesGestion.ActualizarLista

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoDetalleMaterial_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoEliminarActividad_Click()
    Dim strFranqueado As String, strEstaFacturado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaActividades.Column(7), "")
    
    If strEstaFacturado <> "" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    
    
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento al que pertenecen esta actividad está Facturado"
        Err.Raise 1000
    End If
    If strFranqueado = "Sí" Then
        m_Error = "El evento al que pertenecen esta actividad está Franqueado"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar la actividad con Descripción: " & Nz(Me.ListaActividades.Column(5), "") & " ?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar actividad")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       ActividadEliminar = cstr(intNumeroMaterialesEliminados) & ";" & strCadenaMaterialesNoEliminados
    '       ActividadEliminar = "#ERR" & "|" & m_Error
    '-----------------------------------------------
    flag = ActividadEliminar(strIDActividadSeleccionada, "Sí", "No", strEstaFacturado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método ActividadEliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDMaterialSeleccionado = ""
    ActualizarActividades
    ActualizarMaterial
    strIDActividadSeleccionada = ""
    Me.ComandoAltaMaterial.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar eliminar la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoEliminarEvento_Click()
    Dim strFranqueado As String, strEstaFacturado As String
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    If strFranqueado <> "Sí" And strFranqueado <> "No" Then
        m_Error = "Seleccione un evento de la lista"
        Err.Raise 1000
    End If
    
    If strFranqueado = "Sí" Then
        m_Error = "El evento está Franqueado"
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Then
        m_Error = "El evento al que pertenecen estas actividades está Facturado"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el evento: " & Left(Nz(Me.ListaEventos.Column(6), ""), 20) & _
                        " con todas las actividades y materiales que puedieran ser dependientes de él incluyendo el parte que pudiera estar generado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de Evento")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       Eliminar = CStr(intNumeroActividadesEliminadas) & ";" & CStr(intNumeroActividadesNoEliminadas) & _
                        ";" & CStr(intNumeroMaterialesEliminados) & ";" & CStr(intNumeroMaterialesNoEliminados)
    '       Eliminar="#ERR" & "|" & m_Error
    '---------------------------------------------------------------------------------------------------------------
    flag = Eliminar(strIDEventoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaEventos.Requery
    strIDEventoSeleccionado = ""
    strIDActividadSeleccionada = ""
    strIDMaterialSeleccionado = ""
    strEquipoSeleccionado = ""
    ActualizarEventos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar eliminar el evento se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcelActividades_Click()
    Dim strSQLInicial As String, strParteWhere As String, strSQLTotal As String, strSQLLista As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strIDEventoSeleccionado = "" Then
        m_Error = "Se ha de SeleccionarEvento el evento primero"
        Err.Raise 1000
    End If
    strSQLLista = Me.ListaActividades.RowSource
    If InStr(1, strSQLLista, "Where") = 0 Then
        strParteWhere = ";"
    Else
        dato = Split(strSQLLista, "Where")
        strParteWhere = "Where" & dato(1)
    End If
    strSQLInicial = "SELECT TbActividades.IDActividad AS ID, TbActividades.IDEvento, TbActividades.ALIASTECNICO AS TÉCNICO, " & _
                    "TbActividades.FechaAlta AS FALTA,  TbActividades.HorasLaborables, TbActividades.HorasExtras, " & _
                    "TbActividades.Ubicacion, TbActividades.DESCRIPCION AS DESCRIPCIÓN, TbActividades.IDEVENTOGENERADO, " & _
                    "TbMaterial.Material, " & _
                    "TbMaterial.PN, TbMaterial.NS, TbMaterial.COSTE, TbMaterial.TipoAccion, TbMaterial.ReparadoPor, " & _
                    "TbMaterial.FechaEntrega, TbMaterial.Descripcion  " & _
                    "FROM (TbActividades LEFT JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento "
    strSQLTotal = strSQLInicial & strParteWhere
    GenerarConsultasSQL strSQLTotal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    '------------------------------------------
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ComandoExportarExcelEventos_Click()
    
    Dim m_Col As Scripting.Dictionary
    Dim m_EventoBusqueda As New EventoBusqueda
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    With m_EventoBusqueda
        .BUI = Nz(Me.ComboBUIBusqueda, "")
        .Equipo = Nz(Me.ComboEquipoBusqueda, "")
        .FechaFinal = Nz(Me.FechaFinalBusqueda, "")
        .FechaInicial = Nz(Me.FechaInicialBusqueda, "")
        .Franqueado = Nz(Me.ComboFranqueadoBusqueda, "")
        .IDEVENTO = Nz(Me.ComboIDEventoBusqueda, "")
        .SubSistema = Nz(Me.ComboSubsistemaBusqueda, "")
        .TipoEvento = Nz(Me.ComboTipoEventoBusqueda, "")
        
    End With
    Set m_Col = getEventosFiltrados(p_EventoBusqueda:=m_EventoBusqueda, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_Error = "No se ha encontrado ningún Evento"
        Err.Raise 1000
    End If
    GenerarConsultaEventos p_Col:=m_Col, p_Error:=m_Error
    If m_Error <> "" Then
       Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcelEventos_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcelMateriales_Click()
    Dim strSQLInicial As String, strParteWhere As String, strSQLTotal As String, strSQLLista As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strIDActividadSeleccionada = "" Then
        m_Error = "Se ha de seleccionar una IDActividad primero"
        Err.Raise 1000
    End If
    strSQLLista = Me.ListaMateriales.RowSource
    strSQLInicial = "SELECT TbMaterial.Material, TbMaterial.TipoAccion, TbMaterial.PN, TbMaterial.NS, TbMaterial.COSTE, TbMaterial.ReparadoPor, " & _
                    "TbMaterial.FechaEntrega, TbActividades.IDEvento, TbEventos.NODO, TbEventos.BUI, TbEventos.SUBSISTEMA, " & _
                    "TbEquipos.Equipo " & _
                    "FROM (((TbMaterial LEFT JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                    "LEFT JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                    "LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) "
    If InStr(1, strSQLLista, "Where") > 0 Then
        dato = Split(strSQLLista, "Where")
        strParteWhere = "Where" & dato(1)
        strSQLTotal = strSQLInicial & strParteWhere
    Else
        strSQLTotal = strSQLInicial & ";"
    End If
    GenerarConsultasSQL strSQLTotal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoFranquear_Click()
    
    Dim m_ObjEvento As Evento
    Dim m_CausaTecnico As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If strIDEventoSeleccionado = "" Then
        strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(0), "")
        If strIDEventoSeleccionado = "" Then
            m_Error = "No se sabe de qué evento se trata"
            Err.Raise 1000
        End If
    End If
    Set m_ObjEvento = Constructor.getEvento(strIDEventoSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEvento Is Nothing Then
        m_Error = "No se ha encontrado el registro " & strIDEventoSeleccionado
        Err.Raise 1000
    End If
    
    If Me.ComandoFranquear.Caption = "Franquear" Then
        If m_ObjEvento.Franqueado = True Then
            m_Error = "Este evento ya está franqueado"
            Err.Raise 1000
        End If
        If FormularioAbierto("FormEventoFranqueo") Then
            DoCmd.Close acForm, "FormEventoFranqueo", acSaveNo
        End If
        Set m_ObjEventoActivo = m_ObjEvento
        DoCmd.OpenForm "FormEventoFranqueo"
        Set m_FormFranqueo = Forms("FormEventoFranqueo")
    ElseIf Me.ComandoFranquear.Caption = "Quitar Franqueo" Then
        If m_ObjEvento.Franqueado = False Then
            m_Error = "Este evento no está franqueado"
            Err.Raise 1000
        End If
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("¿Desea realmente quitar el franqueo del evento: " & strIDEventoSeleccionado & " ?", vbExclamation + vbYesNo + vbDefaultButton2, "Quitar Franqueo de Evento")
        If pregunta <> vbYes Then
            Exit Sub
        End If
        VBA.DoEvents
        DoCmd.Hourglass True
        VBA.DoEvents
        m_ObjEvento.QuitarFranqueo m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        ListaEventos_Click
        
    Else
        m_Error = "Este botón ha de tener el título Franquear ó Quitar Franqueo"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar franquear el evento se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoGestionPartes_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormParteGestion"
End Sub
Private Sub Form_Open(Canel As Integer)
    Dim lst As ListBox
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    

    Me.Caption = Me.lblTitulo1.Caption
    Me.ComandoAltaEvento.Enabled = True
    Me.ComandoGestionPartes.Enabled = True
    Me.ComandoEditarEvento.Enabled = False
    Me.ComandoEliminarEvento.Enabled = False
    Me.ComandoCopiarEvento.Enabled = False
    Me.ComandoFranquear.Caption = "Franquear"
    Me.ComandoFranquear.Enabled = False
    
    
    Me.ComandoExportarExcelEventos.Enabled = False
    Me.ComandoAltaActividad.Enabled = False
    Me.ComandoAltaActividadesPorLotes.Enabled = False
    Me.ComandoEditarActividad.Enabled = False
    Me.ComandoCopiarActividad.Enabled = False
    Me.ComandoEliminarActividad.Enabled = False
    
    
    
    Me.ComandoAltaActividadSinEvento.Enabled = True
    Me.ComandoExportarExcelActividades.Enabled = False
    Me.ComandoAltaMaterial.Enabled = False
    Me.ComandoDetalleMaterial.Enabled = False
    Me.ComandoExportarExcelMateriales.Enabled = False
    
    Me.ComandoAdjuntosEventos.Enabled = False
    
    Me.ImagenAnexoEvento.Visible = False
    Me.lblAnexo.Visible = False
    Me.ImagenParteAnexado.Visible = False
    Me.lblParte.Visible = False
    Me.ComandoAdjuntosActividades.Enabled = False
    
    Me.ImagenAnexoActividad.Visible = False
    
    Me.ComandoAdjuntosMateriales.Enabled = False
   
    If FormularioAbierto("FormInformesDesdeLista") Then
        Me.ComandoParaInforme.Visible = True
    Else
        Me.ComandoParaInforme.Visible = False
    End If
    CargarSubSistema
    CargarEquipo
    Set lst = Me.ListaEventos
    
    m_SQL = FiltrarEventos(Nz(Me.ComboIDEventoBusqueda, ""), Nz(Me.ComboBUIBusqueda, ""), Nz(Me.ComboSubsistemaBusqueda, ""), _
                        Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.FechaInicialBusqueda, ""), _
                        Nz(Me.FechaFinalBusqueda, ""), Nz(Me.ComboTipoEventoBusqueda, ""), _
                        Nz(Me.ComboFranqueadoBusqueda, ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    lst.RowSource = m_SQL
    lst.Requery
    Set lst = Me.ListaActividades
    
    m_SQL = FiltrarActividades(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    lst.RowSource = m_SQL
    lst.Requery
    Set lst = Me.ListaMateriales
    
    m_SQL = FiltrarMateriales(, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    lst.RowSource = m_SQL
    lst.Requery
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImagenParteAnexado_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AbrirEnLocal = strURLAnexoLocal
    '       AbrirEnLocal = "#ERR" & "|" & m_Error
    '-------------------------------------------------------------------
    flag = AbrirEnLocal(strURLParteEvento, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ListaActividades_Click()
    Dim m_ActividadSeleccionada As Actividad
    Dim strEstaFacturado As String, strFranqueado As String, strActividadSeleccionada As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = Nz(Me.ListaActividades.Column(1), "")
    strActividadSeleccionada = Nz(Me.ListaActividades.Column(5), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    If strIDEventoSeleccionado = "" Then
        Me.ComandoAltaActividad.Enabled = False
        Me.ComandoAltaActividadesPorLotes.Enabled = False
    Else
        Me.ComandoAltaActividad.Enabled = True
        Me.ComandoAltaActividadesPorLotes.Enabled = True
    End If
    Me.ComandoEditarActividad.Enabled = False
    Me.ComandoCopiarActividad.Enabled = False
    Me.ComandoEliminarActividad.Enabled = False
    
    Me.ComandoAltaActividadSinEvento.Enabled = False
    
   
    Me.ImagenAnexoActividad.Visible = False
    
    Me.ComandoAdjuntosActividades.Visible = True
    
    strURLAnexoActividad = ""
    
    If strIDActividadSeleccionada = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        
        Exit Sub
    End If
    
    If strFranqueado = "Sí" Or strEstaFacturado = "Sí" Then
        Me.ComandoAltaMaterial.Enabled = False
        Me.ComandoCopiarActividad.Enabled = False
    Else
        Me.ComandoAltaMaterial.Enabled = True
        Me.ComandoCopiarActividad.Enabled = True
    End If
    If strEstaFacturado = "Sí" Then
        Me.ComandoEliminarActividad.Enabled = False
    Else
        Me.ComandoEliminarActividad.Enabled = True
    End If
    Me.ComandoEditarActividad.Enabled = True
    
    Me.ComandoAltaActividadSinEvento.Enabled = True
    Me.ComandoAdjuntosActividades.Enabled = True
   
    
    
    Set m_ActividadSeleccionada = Constructor.getActividad(strIDActividadSeleccionada, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    strURLAnexoActividad = m_ActividadSeleccionada.URLAnexos
    If InStr(1, strURLAnexoActividad, ";") <> 0 Then
        Me.ImagenAnexoActividad.Visible = True
        Me.ImagenAnexoActividad.BorderWidth = 0
        Me.ImagenAnexoActividad.BorderColor = RGB(255, 255, 255)
       
    Else
        If strURLAnexoActividad = "" Then
            Me.ImagenAnexoMaterial.Visible = False
            
        Else
            Me.ImagenAnexoActividad.Visible = True
            If fso.FileExists(strURLAnexoActividad) Then
                Me.ImagenAnexoActividad.BorderStyle = 0
                Me.ImagenAnexoActividad.BorderWidth = 0
                
            Else
                Me.ImagenAnexoActividad.BorderStyle = 1
                Me.ImagenAnexoActividad.BorderWidth = 2
                Me.ImagenAnexoActividad.BorderColor = RGB(255, 0, 0)
                
            End If
        End If
    End If
    Me.ComandoAdjuntosActividades.Visible = True
    ActualizarMaterial
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar hacer click sobre la lista de actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ListaEventos_Click()
    
    
    Dim strEstaFacturado As String, strFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDEventoSeleccionado = Nz(Me.ListaEventos.Column(1), "")
    strFranqueado = Nz(Me.ListaEventos.Column(7), "")
    strEstaFacturado = Nz(Me.ListaEventos.Column(0), "")
    strEquipoSeleccionado = Nz(Me.ListaEventos.Column(4), "")
    If strEstaFacturado = "*" Then
        strEstaFacturado = "Sí"
    Else
        strEstaFacturado = "No"
    End If
    
    Me.ComandoEditarEvento.Enabled = False
    Me.ComandoEliminarEvento.Enabled = False
    Me.ComandoCopiarEvento.Enabled = False
    Me.ComandoFranquear.Enabled = False
    Me.ComandoFranquear.Caption = "Franquear"
    
    
    Me.ImagenParteAnexado.Visible = False
    Me.lblParte.Visible = False
    Me.ImagenAnexoEvento.Visible = False
    Me.lblAnexo.Visible = False
    Me.ComandoAdjuntosEventos.Visible = True
    Me.ComandoAdjuntosEventos.Enabled = False
    
    strURLParteEvento = ""
    strSeleccionadoFranqueado = ""
    
    
    If strIDEventoSeleccionado = "" Then
        Set m_ObjEventoSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ObjEventoSeleccionado Is Nothing Then
        Set m_ObjEventoSeleccionado = Constructor.getEvento(strIDEventoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If strIDEventoSeleccionado <> m_ObjEventoSeleccionado.IDEVENTO Then
            Set m_ObjEventoSeleccionado = Constructor.getEvento(strIDEventoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If strFranqueado = "Sí" Or strEstaFacturado = "Sí" Then
        Me.ComandoAltaActividad.Enabled = False
        Me.ComandoAltaActividadesPorLotes.Enabled = False
        Me.ComandoEliminarEvento.Enabled = False
        
    Else
        Me.ComandoAltaActividad.Enabled = True
        Me.ComandoAltaActividadesPorLotes.Enabled = True
        Me.ComandoEliminarEvento.Enabled = True
        
    End If
    Me.ComandoAltaActividadSinEvento.Enabled = True
    Me.ComandoEditarEvento.Enabled = True
    Me.ComandoFranquear.Enabled = True
    Me.ComandoCopiarEvento.Enabled = True
    Me.ComandoAdjuntosEventos.Enabled = True
    If strFranqueado = "Sí" Then
        Me.ComandoFranquear.Caption = "Quitar Franqueo"
    Else
        Me.ComandoFranquear.Caption = "Franquear"
    End If
    
    
    strURLParteEvento = m_ObjEventoSeleccionado.URLAnexoParte
    m_Error = m_ObjEventoSeleccionado.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If strURLParteEvento = "" Then
        Me.ImagenParteAnexado.Visible = False
        Me.lblParte.Visible = False
    Else
        Me.ImagenParteAnexado.Visible = True
        Me.lblParte.Visible = True
        If fso.FileExists(strURLParteEvento) Then
            Me.ImagenParteAnexado.BorderStyle = 0
            
        Else
            Me.ImagenParteAnexado.BorderStyle = 2
            Me.ImagenParteAnexado.BorderColor = RGB(255, 0, 0)
            
        End If
    End If
    
    strURLAnexoEvento = m_ObjEventoSeleccionado.URLAnexos
    If strURLAnexoEvento = "" Then
        Me.ImagenAnexoEvento.Visible = False
        Me.lblAnexo.Visible = False
    Else
        Me.ImagenAnexoEvento.Visible = True
        Me.lblAnexo.Visible = True
        If InStr(1, strURLAnexoEvento, ";") <> 0 Then
            Me.ImagenAnexoEvento.BorderStyle = 0
            
        Else
            If fso.FileExists(strURLAnexoEvento) Then
                Me.ImagenAnexoEvento.BorderStyle = 0
                
            Else
                Me.ImagenAnexoEvento.BorderStyle = 1
                Me.ImagenAnexoEvento.BorderWidth = 2
                Me.ImagenAnexoEvento.BorderColor = RGB(255, 0, 0)
                
            End If
        End If
    End If
    Me.ComandoAdjuntosEventos.Visible = True
    ActualizarActividades
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar hacer click sobre la lista, se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub ActualizarEventos()
    Dim strTitulo As String, intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant
    Dim strIDEventoSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDEventoSeleccionado = ""
    strIDActividadSeleccionada = ""
    strIDMaterialSeleccionado = ""
    strEquipoSeleccionado = ""
    strSeleccionadoFranqueado = ""
    Me.ComandoEditarEvento.Enabled = False
    Me.ComandoEliminarEvento.Enabled = False
    Me.ComandoCopiarEvento.Enabled = False
    Me.ComandoFranquear.Enabled = False
    
    
    
    Me.ComandoAdjuntosEventos.Enabled = False
    
    Me.ImagenAnexoEvento.Visible = False
    Me.lblAnexo.Visible = False
    Me.ImagenParteAnexado.Visible = False
    Me.lblParte.Visible = False
    
    Me.ComandoAdjuntosActividades.Enabled = False
    
    Me.ImagenAnexoActividad.Visible = False
    
    
    Me.ComandoAdjuntosMateriales.Enabled = False
   
    Me.ImagenAnexoMaterial.Visible = False
    
    Me.ComandoExportarExcelEventos.Enabled = False
    Me.ComandoEliminarMaterial.Enabled = False
    Set lst = Me.ListaActividades
    m_SQL = "SELECT TbActividades.IDActividad AS ID, TbActividades.ALIASTECNICO AS TÉCNICO, TbActividades.FechaAlta AS FALTA, " & _
                "TbActividades.HorasLaborables AS HLab, TbActividades.DESCRIPCION AS DESCRIPCIÓN " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.IDActividad)='-1')) " & _
                "ORDER BY TbActividades.FechaAlta DESC;"
    lst.RowSource = m_SQL
    lst.Requery
    Me.ComandoAltaActividad.Enabled = False
    Me.ComandoAltaActividadesPorLotes.Enabled = False
    Me.ComandoEditarActividad.Enabled = False
    Me.ComandoEliminarActividad.Enabled = False
    
    Me.ComandoCopiarActividad.Enabled = False
    Set lst = Me.ListaMateriales
    m_SQL = "SELECT TbMaterial.IDMaterial, TbMaterial.Material " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='-1'));"
    lst.RowSource = m_SQL
    lst.Requery
    Set lst = Me.ListaEventos
    strIDEventoSeleccionadoActualmente = Nz(lst.Column(1), "")
    
    m_SQL = FiltrarEventos(Nz(Me.ComboIDEventoBusqueda, ""), Nz(Me.ComboBUIBusqueda, ""), Nz(Me.ComboSubsistemaBusqueda, ""), _
                        Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.FechaInicialBusqueda, ""), _
                        Nz(Me.FechaFinalBusqueda, ""), Nz(Me.ComboTipoEventoBusqueda, ""), _
                        Nz(Me.ComboFranqueadoBusqueda, ""), m_Error)
    If m_Error <> "" Then
        
        Err.Raise 1000
    End If
   
    lst.RowSource = m_SQL
    lst.Requery
    Me.ComandoAltaMaterial.Enabled = False
    Me.ComandoDetalleMaterial.Enabled = False
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcelEventos.Enabled = False
    ElseIf intElementos = 1 Then
        Me.ComandoExportarExcelEventos.Enabled = True
        strIDEventoSeleccionadoActualmente = lst.Column(1, 1)
        lst.Selected(1) = True
        ListaEventos_Click
    Else
        Me.ComandoExportarExcelEventos.Enabled = True
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDEventoSeleccionadoActualmente = lst.Column(1, i)
                ListaEventos_Click
                Exit For
            End If
        Next
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar actualizar la lista de eventos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ActualizarActividades()
    Dim intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant
    Dim strIDActividadSeleccionadaActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = ""
    Me.ComandoEditarActividad.Enabled = False
    Me.ComandoCopiarActividad.Enabled = False
    Me.ComandoEliminarActividad.Enabled = False
   
    Me.ComandoAdjuntosActividades.Enabled = False
    Me.ImagenAnexoActividad.Visible = False
    
    Me.ComandoExportarExcelActividades.Enabled = False
    Set lst = Me.ListaActividades
    strIDActividadSeleccionadaActualmente = Nz(lst.Column(1), "")
    
    m_SQL = FiltrarActividades(strIDEventoSeleccionado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcelActividades.Enabled = False
    ElseIf intElementos = 1 Then
        Me.ComandoExportarExcelActividades.Enabled = True
        strIDActividadSeleccionadaActualmente = lst.Column(1, 1)
        lst.Selected(1) = True
        ListaActividades_Click
    Else
        Me.ComandoExportarExcelActividades.Enabled = True
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDActividadSeleccionadaActualmente = lst.Column(1, i)
                ListaActividades_Click
                Exit For
            End If
        Next
    End If
    ActualizarMaterial
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar actualizar la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Sub ActualizarMaterial()
    Dim intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant
    Dim strIDMaterialSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strIDActividadSeleccionada = ""
    strIDMaterialSeleccionado = ""
    strEquipoSeleccionado = ""
    Me.ComandoDetalleMaterial.Enabled = False
    Me.ImagenAnexoMaterial.Visible = False
    Me.ComandoAdjuntosMateriales.Enabled = False
    
    
    Me.ComandoExportarExcelMateriales.Enabled = False
    Me.ComandoEliminarMaterial.Enabled = False
    Set lst = Me.ListaMateriales
    strIDMaterialSeleccionadoActualmente = Nz(lst.Column(0), "")
    
    m_SQL = FiltrarMateriales(Nz(Me.ListaActividades.Column(1), ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcelMateriales.Enabled = False
    ElseIf intElementos = 1 Then
        Me.ComandoExportarExcelMateriales.Enabled = True
        strIDMaterialSeleccionadoActualmente = lst.Column(0, 1)
        lst.Selected(1) = True
        ListaMateriales_Click
    Else
        Me.ComandoExportarExcelMateriales.Enabled = True
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDMaterialSeleccionadoActualmente = lst.Column(0, i)
                ListaMateriales_Click
                Exit For
            End If
        Next
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar actualizar la lista de Materiales se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Public Function DameURLParteDeEvento( _
                                        strIDEvento As String _
                                        ) As String
                                        
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/03/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a la tabla TbAnexos y ve el nombre del documento anexo
    '
    '   -llamada por:
    '       -Form_FormEventoGestion.ListaEventos_Click
    
    '   -Devuelve:
    '       DameURLParteDeEvento = strURLAnexo
    '       DameURLParteDeEvento = "#ERR" & "|" & m_Error
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset
    Dim strNombreArchivo As String, strURLAnexo As String
    On Error GoTo errores
    If strIDEvento = "" Then
        m_Error = "No se ha indicado el IDEvento"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT TbPartesPpal.NombreArchivoAdjunto " & _
            "FROM TbPartesPpal INNER JOIN TbPartesDetalle ON TbPartesPpal.IDParte = TbPartesDetalle.IDParte " & _
            "WHERE (((TbPartesDetalle.IDEvento)='" & strIDEvento & "'));"
    
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strNombreArchivo = Nz(.Fields("NombreArchivoAdjunto"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If strNombreArchivo <> "" Then
        strURLAnexo = m_ObjEntorno.URLDirectorioPartesFirmados & strNombreArchivo
    End If
    
    DameURLParteDeEvento = strURLAnexo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método DameURLParteDeEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameURLParteDeEvento = "#ERR" & "|" & m_Error
End Function

Public Function Eliminar( _
                          strIDEvento As String, _
                          Optional strEliminando As String, _
                          Optional strRespetandoElFranqueo As String _
                        ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '        -Tabla TbActividades --->Elimina todas las que TbActividades.IDEvento=strIDEvento
    '        -Tabla TbMaterialUtilizado --->Elimina todas las que TbMaterialUtilizado.IDEvento=strIDEvento
    '        -Tabla TbEventoAnexos--->Elimina todos los registros que TbEventoAnexos.IDEvento=strIDEvento
    '        -Tabla TbEventoAnexos--->Elimina todos los registros que TbPartesDetalle.IDEvento=strIDEvento
    '        -Elimina los anexos físicos también del disco
    
    '       -1) Verificación de datos
    '           -strIDEvento  existente
    '       -2) verifica que no haya formado parte de una certificación
    
    '   -Llamada desde
    '       -Form_FormEventoGestion.ComandoEliminarEvento_Click
    '   -Devuelve:
    '       Eliminar = CStr(intNumeroActividadesEliminadas) & ";" & CStr(intNumeroActividadesNoEliminadas) & _
                        ";" & CStr(intNumeroMaterialesEliminados) & ";" & CStr(intNumeroMaterialesNoEliminados)
    '       Eliminar="#ERR" & "|" & m_Error
    '---------------------------------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strCadena As String, strURLAnexo As String, strNombreAnexo As String, _
        strCadenaAnexosAEliminar As String, strCadenaAnexosEliminados As String, dato1
    Dim strIDEventoEnParte As String, strIDActividad As String
    Dim strResultadoEliminarActividad As String, strResultadoMaterialesNoEliminados As String, strNumeroMaterialesEliminados As String
    Dim intNumeroActividadesEliminadas As Integer, intNumeroActividadesNoEliminadas As Integer, intNumeroMaterialesEliminados As Integer, _
        intNumeroMaterialesNoEliminados As Integer
    On Error GoTo errores
    
    If strIDEvento = "" Then
        m_Error = "No se puede dejar en blanco el IDEvento"
        Err.Raise 1000
    End If
    If strEliminando <> "Sí" And strEliminando <> "No" Then
        strEliminando = "Sí"
    End If
    If m_ObjEventoActivo Is Nothing Then
        Set m_ObjEventoActivo = Constructor.getEvento(strIDEvento)
    Else
        If m_ObjEventoActivo.IDEVENTO <> strIDEvento Then
            Set m_ObjEventoActivo = Constructor.getEvento(strIDEvento)
        End If
    End If
    'hemos de ver si ha formado parte de algún parte
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & m_Error
    '--------------------------------------------
    flag = Dame("TbPartesDetalle", "IDEvento", "IDEvento", strIDEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEventoEnParte = flag
    If strIDEventoEnParte <> "" Then
        m_Error = "No se puede dar de Eliminar un evento que ha formado parte de un parte de averías"
        Err.Raise 1000
    End If
    Dim m_Evento As Evento
    Set m_Evento = Constructor.getEvento(strIDEvento, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If strEliminando <> "Sí" Then
        If Not m_Evento.ColAnexos Is Nothing Then
            If m_Evento.AnexosCerrados <> EnumSino.Sí Then
                m_Error = "Al menos hay un anexo abierto"
                Err.Raise 1000
            End If
        End If
    Else
        m_Evento.EliminarAnexos m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_ObjEventoActivo.EliminarEquiposMedida
        
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
           .MoveFirst
           Do While Not .EOF
                strIDActividad = Nz(.Fields("IDActividad"), "")
                '-----------------------------------------------
                '   -Devuelve:
                '       ActividadEliminar = cstr(intNumeroMaterialesEliminados) & "[" & strCadenaMaterialesNoEliminados
                '       ActividadEliminar = "#ERR" & "|" & m_Error
                '-----------------------------------------------
                flag = ActividadEliminar(strIDActividad, strEliminando, strRespetandoElFranqueo)
                If InStr(1, flag, "|") <> 0 Then
                    intNumeroActividadesNoEliminadas = intNumeroActividadesNoEliminadas + 1
                Else
                    strResultadoEliminarActividad = flag
                    If InStr(1, flag, "[") <> 0 Then
                        dato = Split(strResultadoEliminarActividad, "[")
                        strNumeroMaterialesEliminados = Nz(dato(0), "")
                        strResultadoMaterialesNoEliminados = Nz(dato(1), "")
                        If IsNumeric(strNumeroMaterialesEliminados) Then
                            intNumeroMaterialesEliminados = intNumeroMaterialesEliminados + CLng(strNumeroMaterialesEliminados)
                        End If
                        If strResultadoMaterialesNoEliminados <> "" Then
                            dato1 = Split(strResultadoMaterialesNoEliminados, ";")
                            intNumeroMaterialesNoEliminados = intNumeroMaterialesNoEliminados + UBound(dato1) + 1
                        End If
                    End If
                End If
            .MoveNext
           Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "DELETE * " & _
            "FROM TbEventosEquipoMedida " & _
            "WHERE IDEvento='" & strIDEvento & "';"
    getdb.Execute m_SQL
    '-----------------
    ' Eliminar el evento
    '-----------------
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If strEliminando = "Sí" Then
                .Delete
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Eliminar = CStr(intNumeroActividadesEliminadas) & ";" & CStr(intNumeroActividadesNoEliminadas) & _
            ";" & CStr(intNumeroMaterialesEliminados) & ";" & CStr(intNumeroMaterialesNoEliminados)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Evento.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Eliminar = "#ERR" & "|" & m_Error
End Function




Private Sub m_FormEventoCopia_AltaEvento(m_Evento As Evento)
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.cmdSalir.SetFocus
    ComandoActualizarTodo_Click
    Me.ComboIDEventoBusqueda = m_Evento.IDEVENTO
    ComboIDEventoBusqueda_AfterUpdate
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Se ha generado el evento con ID: " & m_Evento.IDEVENTO, vbInformation, "Evento Generado")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormEventoCopia_AltaEvento se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormFranqueo_Franqueado()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    strSeleccionadoFranqueado = "Sí"
    
    ListaEventos_Click
    'ComandoEditarEvento_Click

    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormFranqueo_Franqueado se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub