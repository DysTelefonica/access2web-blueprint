Option Compare Database
Option Explicit
Dim strIDEquipoSeleccionado As String
Dim strBUIAnterior As String, strSubSistemaAnterior As String, strIDEquipoAnterior As String, strFiltroAnterior
Public Sub LimpiarUltimos()
    strBUIAnterior = ""
    strSubSistemaAnterior = ""
    strIDEquipoAnterior = ""
    strFiltroAnterior = ""
End Sub
Public Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarBUI_Click
    ComandoLimpiarSubSistema_Click
    ComandoLimpiarIDEquipo_Click
    ComandoLimpiarFiltro_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar Actualizar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEquipoAlta_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    DoCmd.OpenForm "FormIDEquiposGestion"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un equipo se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                            Optional strBUI As String, _
                            Optional strSUBSISTEMA As String, _
                            Optional strEquipo As String, _
                            Optional strCriterioFiltro As String _
                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    '       -me.ActualizarLista
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strSQLParteInicial As String, strSQLFinal As String
    Dim strParteWhereBUI As String, strParteWhereSubSistema As String, strParteWhereEquipo As String, strParteWhereCriterio As String
    Dim strTextoError As String
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbEquipos.IDEquipo, TbEquipos.Equipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                    "FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.IDEquipo "
    If strBUI = "" And strSUBSISTEMA = "" And strEquipo = "" And strCriterioFiltro = "" Then
        strParteWhere = "WHERE (((TbEquipos.BUI)='-1'));"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If strBUI = "" Then
        strParteWhereBUI = "((TbEquipos.BUI) Like '*' Or (TbEquipos.BUI) Is Null)"
    Else
        strParteWhereBUI = "((TbEquipos.BUI) = '" & strBUI & "')"
    End If
    If strSUBSISTEMA = "" Then
        strParteWhereSubSistema = "((TbEquipos.SUBSISTEMA) Like '*' Or (TbEquipos.SUBSISTEMA) Is Null)"
    Else
        strParteWhereSubSistema = "((TbEquipos.SUBSISTEMA) = '" & strSUBSISTEMA & "')"
    End If
    If strEquipo = "" Then
        strParteWhereEquipo = "((TbEquipos.Equipo) Like '*')"
    Else
        strParteWhereEquipo = "((TbEquipos.Equipo) Like '*" & strEquipo & "*')"
    End If
    If strCriterioFiltro = "Con Planificaciones pendientes" Then
        strParteWhereCriterio = "((NPlanificaciones(0,0,[TbEquipos].[IDEquipo],'No Realizadas'))>0)"
    ElseIf strCriterioFiltro = "Sin Planificaciones Pendientes" Then
        strParteWhereCriterio = "((NPlanificaciones(0,0,[TbEquipos].[IDEquipo],'No Realizadas'))=0)"
    ElseIf strCriterioFiltro = "Sin Planificaciones" Then
        strParteWhereCriterio = "((EquipoSinPlanificaciones([TbEquipos].[IDEquipo],'No'))=1)"
    Else
        strCriterioFiltro = ""
   
    End If
    If strCriterioFiltro = "" Then
        strParteWhere = "WHERE (" & _
                                strParteWhereBUI & " AND " & _
                                strParteWhereSubSistema & " AND " & _
                                strParteWhereEquipo & ") ORDER BY TbEquipos.Equipo;"
    Else
        strParteWhere = "WHERE (" & _
                                strParteWhereBUI & " AND " & _
                                strParteWhereSubSistema & " AND " & _
                                strParteWhereEquipo & " AND " & strParteWhereCriterio & ") ORDER BY TbEquipos.Equipo;"
    End If
    strSQLFinal = strSQLParteInicial & strParteWhere
    Filtrar = strSQLFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Filtrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Filtrar = "#ERR" & "|" & strTextoError
End Function
Public Sub ActualizarLista()
    Dim intElementos As Integer, lst As ListBox, VarItem As Variant
    Dim strIDEquipoSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = ""
    Set lst = Me.ListaFiltrados
    strIDEquipoSeleccionadoActualmente = Nz(lst.Column(0), "")
    Me.ComandoEquipoEdicion.Enabled = False
    Me.ComandoMantenimientosPlanificados.Enabled = False
    Me.ComandoMantenimientosPlanificados.Caption = "Mtos. Planificados"
    Me.ComandoSacarMeterEnPlanificacion.Enabled = False
    Me.ComandoPlanificarNuevoMantenimiento.Enabled = False
    Me.lblNotaAlPie.Visible = False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = Filtrar( _
                    Nz(Me.ComboBui, ""), Nz(Me.ComboSubsistema, ""), _
                    Nz(Me.ComboEquipos, ""), Nz(Me.ComboFiltro, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Filtrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = flag
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcel.Enabled = False
    Else
        If intElementos = 1 Then
            strIDEquipoSeleccionadoActualmente = lst.Column(0, 1)
            lst.Selected(1) = True
            ListaFiltrados_Click
        Else
            For i = 1 To lst.ListCount - 1
                If lst.Selected(i) = True Then
                    strIDEquipoSeleccionadoActualmente = lst.Column(0, i)
                    ListaFiltrados_Click
                    Exit For
                End If
            Next
        End If
        Me.ComandoExportarExcel.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEquipoEdicion_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Seleccione un equipo de la lista"
        Err.Raise 1000
    End If
    '-------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormPlanificacionEdicionEquipo") Then
        DoCmd.Close acForm, "FormPlanificacionEdicionEquipo", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionEdicionEquipo", , , , , , strIDEquipoSeleccionado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar el equipo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoLimpiarBUI_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboBui, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboBui = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarIDEquipo_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEquipos, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipos = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarSubSistema_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboSubsistema, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboSubsistema = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoSacarMeterEnPlanificacion_Click()
    Dim strTituloBoton As String, strTexto As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Ha de seleccionar un equipo de la lista"
        Err.Raise 1000
    End If
    strTituloBoton = Me.ComandoSacarMeterEnPlanificacion.Caption
    If strTituloBoton = "Sacar equipo de Preventivos" Then
        strTexto = "¿Desea sacar al equipo " & Nz(Me.ListaFiltrados.Column(1), "") & " del mantenimiento?"
    ElseIf strTituloBoton = "Meter equipo en Preventivos" Then
         strTexto = "¿Desea Readmitir al equipo " & Nz(Me.ListaFiltrados.Column(1), "") & " a la lista de mantenimiento?"
    Else
        strTextoError = "El título del botón no está entre las opciones disponibles"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox(strTexto, vbExclamation + vbYesNo + vbDefaultButton1, "Meter/Sacar equipo en Mantenimiento")
    If pregunta <> vbYes Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strTituloBoton = "Sacar equipo de Preventivos" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       SacarEquipoDePlanificacion = strDescriptivo
        '       SacarEquipoDePlanificacion = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = SacarEquipoDePlanificacion(strIDEquipoSeleccionado)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método SacarEquipoDePlanificacion ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       MeterEquipoEnPlanificacion = strDescriptivo
        '       MeterEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MeterEquipoEnPlanificacion(strIDEquipoSeleccionado)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MeterEquipoEnPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    ListaFiltrados_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar sacar/meter un equipo del mantenimiento se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = Me.ListaFiltrados.RowSource
    GenerarConsultasSQL m_SQL, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar exportar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoLimpiarFiltro_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboFiltro, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboFiltro = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub CargarIDEquipo()
    Dim strBUI As String, strSUBSISTEMA As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipos = Null
    strBUI = Nz(Me.ComboBui, "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    If strBUI = "" Or strSUBSISTEMA = "" Then
        m_SQL = "SELECT TbEquipos.Equipo " & _
                "FROM TbEquipos " & _
                "ORDER BY TbEquipos.Equipo;"
    Else
        m_SQL = "SELECT TbEquipos.Equipo " & _
                "FROM TbEquipos " & _
                "WHERE (((TbEquipos.BUI)='" & strBUI & "') AND ((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & "')) " & _
                "ORDER BY TbEquipos.Equipo;"
    End If
    Me.ComboEquipos.RowSource = m_SQL
    Me.ComboEquipos.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir cargar el combo del subsistema se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub CargarSubSistema()
    Dim strTextoError As String, strBUI As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipos = Null
    Me.ComboSubsistema = Null
    strBUI = Nz(Me.ComboBui, "-1")
    m_SQL = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
            "FROM TbSubsistemaBui INNER JOIN TbEquipos ON (TbSubsistemaBui.SUBSISTEMA = TbEquipos.SUBSISTEMA) AND " & _
            "(TbSubsistemaBui.BUI = TbEquipos.BUI) " & _
            "WHERE (((TbEquipos.BUI)='" & strBUI & "')) " & _
            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    Me.ComboSubsistema.RowSource = m_SQL
    Me.ComboSubsistema.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir cargar el combo del subsistema se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoPlanificarNuevoMantenimiento_Click()
    Dim strArgumento As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "No se se puede determinar el equipo"
        Err.Raise 1000
    End If
    strBUI = Nz(Me.ListaFiltrados.Column(2), "")
    strSUBSISTEMA = Nz(Me.ListaFiltrados.Column(3), "")
    strEquipo = Nz(Me.ListaFiltrados.Column(1), "")
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipoSeleccionado & "|" & strEquipo
    strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipoSeleccionado & "|" & strEquipo
    '-------------------------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormPlanificacionAlta") Then
        DoCmd.Close acForm, "FormPlanificacionAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionAlta", , , , , , strArgumento
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un mantenimiento programado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoMantenimientosPlanificados_Click()
    Dim strArgumento As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "No se se puede determinar el equipo"
        Err.Raise 1000
    End If
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strEquipo
    strBUI = Nz(Me.ListaFiltrados.Column(2), "")
    strSUBSISTEMA = Nz(Me.ListaFiltrados.Column(3), "")
    strEquipo = Nz(Me.ListaFiltrados.Column(1), "")
    strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strEquipo
    '---------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------------
    If FormularioAbierto("FormPlanificacionesBusqueda") Then
        DoCmd.Close acForm, "FormPlanificacionesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionesBusqueda", , , , , , strArgumento
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los mantenimientos planificados se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboBui_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBui, "")
    If strValor <> strBUIAnterior Then
        CargarSubSistema
        CargarIDEquipo
        ActualizarLista
    End If
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboBui_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBui, "")
    If strValor <> strBUIAnterior Then
        CargarSubSistema
        CargarIDEquipo
        ActualizarLista
    End If
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboSubsistema_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistema, "")
    If strValor <> strBUIAnterior Then
        CargarIDEquipo
        ActualizarLista
    End If
    strSubSistemaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboSubSistema_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistema, "")
    If strValor <> strBUIAnterior Then
        CargarIDEquipo
        ActualizarLista
    End If
    strSubSistemaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Canel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    ActualizarLista
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Sub ListaFiltrados_Click()
    Dim strFechaFueraMto As String, intNumeroPlanificaciones As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    Me.ComandoEquipoEdicion.Enabled = False
    Me.ComandoMantenimientosPlanificados.Enabled = False
    Me.ComandoMantenimientosPlanificados.Caption = "Mtos. Planificados"
    Me.ComandoSacarMeterEnPlanificacion.Enabled = False
    Me.ComandoPlanificarNuevoMantenimiento.Enabled = False
    Me.lblNotaAlPie.Visible = False
    If strIDEquipoSeleccionado = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEquipoEdicion.Enabled = True
    Me.ComandoSacarMeterEnPlanificacion.Enabled = True
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionEquipos", "FechabajaParaPlanificacion", "IDEquipo", strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaFueraMto = flag
    If IsDate(strFechaFueraMto) Then
        Me.lblNotaAlPie.Visible = True
        Me.lblNotaAlPie.Caption = "Equipo Fuera de mantenimiento desde : " & strFechaFueraMto
        Me.ComandoPlanificarNuevoMantenimiento.Enabled = False
        Me.ComandoSacarMeterEnPlanificacion.Caption = "Meter equipo en Preventivos"
    Else
        Me.lblNotaAlPie.Visible = False
        Me.ComandoPlanificarNuevoMantenimiento.Enabled = True
        Me.ComandoSacarMeterEnPlanificacion.Caption = "Sacar equipo de Preventivos"
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       NPlanificaciones = clng(strNumeroPlanificaciones)
    '       NPlanificaciones =-1--->Si Error
    '-------------------------------------------------------------------
    intNumeroPlanificaciones = NPlanificaciones(, CInt(strIDEquipoSeleccionado), "No Realizadas")
    If intNumeroPlanificaciones = -1 Then
        strTextoError = "El método NPlanificaciones ha devuelto el error"
        Err.Raise 1000
    End If
    If intNumeroPlanificaciones = 0 Then
        Me.ComandoSacarMeterEnPlanificacion.Enabled = False
    Else
        Me.ComandoSacarMeterEnPlanificacion.Enabled = True
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       NPlanificaciones = clng(strNumeroPlanificaciones)
    '       NPlanificaciones =-1--->Si Error
    '-------------------------------------------------------------------
    intNumeroPlanificaciones = NPlanificaciones(, CInt(strIDEquipoSeleccionado), "Realizadas")
    If intNumeroPlanificaciones = -1 Then
        strTextoError = "El método NPlanificaciones ha devuelto el error"
        Err.Raise 1000
    End If
    If intNumeroPlanificaciones = 0 Then
        Me.ComandoMantenimientosPlanificados.Caption = "Mtos. Planificados (0)"
        Me.ComandoMantenimientosPlanificados.Enabled = False
    Else
        Me.ComandoMantenimientosPlanificados.Caption = "Mtos. Planificados (" & intNumeroPlanificaciones & ")"
        Me.ComandoMantenimientosPlanificados.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click sobre la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function SacarEquipoDePlanificacion( _
                                            strIDEquipo As String, _
                                            Optional strFechaBaja As String, _
                                            Optional strEquipoConPlanificacionesPendientes As String _
                                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va eliminar un equipo de la lista TbPlanificacionEquipos siempre que no tenga ningún programado ya grabado
    '
   
    '   -llamada por:
    '       -Form_FormPlanificacionEquipos.ComandoEquipoBaja_Click
    '       -Form_FormIDEquiposGestion.ComandoMarcarParaMtoProg_Click
    '   -Devuelve:
    '       SacarEquipoDePlanificacion =Descriptivo
    '       SacarEquipoDePlanificacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intNumeroPlanificaciones As Integer, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El Equipo lo identifica un identificador numérico"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionEquipos", "FechabajaParaPlanificacion", "IDEquipo", strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If IsDate(flag) Then
        strTextoError = "Este equipo causó baja de mantenimientos programados el día " & flag
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EquipoPlanificable = strEquipoPlanificable
    '       EquipoPlanificable = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------
    flag = EquipoPlanificable(strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EquipoPlanificable ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "No" Then
        strTextoError = "No se puede dar de baja a un equipo que ya no forma parte de los planificables"
        Err.Raise 1000
    End If
    If strEquipoConPlanificacionesPendientes <> "Sí" And strEquipoConPlanificacionesPendientes <> "No" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       NPlanificaciones = clng(strNumeroPlanificaciones)
        '       NPlanificaciones =-1--->Si Error
        '-------------------------------------------------------------------
        intNumeroPlanificaciones = NPlanificaciones(, CInt(strIDEquipo), "No Realizadas")
        If intNumeroPlanificaciones = -1 Then
            strTextoError = "El método NPlanificaciones ha devuelto el error"
            Err.Raise 1000
        End If
        If intNumeroPlanificaciones = 0 Then
            strEquipoConPlanificacionesPendientes = "No"
        Else
            strEquipoConPlanificacionesPendientes = "Sí"
        End If
    End If
    If strEquipoConPlanificacionesPendientes = "Sí" Then
        strTextoError = "Este equipo tiene al menos una planificación registrada no finalizada"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionEquipos.* " & _
            "FROM TbPlanificacionEquipos " & _
            "WHERE (((TbPlanificacionEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechabajaParaPlanificacion") = strFechaBaja
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    SacarEquipoDePlanificacion = "Equipo dado de baja para las planificaciones"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método SacarEquipoDePlanificacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    SacarEquipoDePlanificacion = "#ERR" & "|" & strTextoError
End Function
Public Function MeterEquipoEnPlanificacion(strIDEquipo As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a establecer a nulo el campo TbPlanificacionEquipos.FechabajaParaPlanificacion
    '       -si no está en la tabla TbPlanificacionEquipos uso el método AltaEquipo
   
    '   -llamada por:
    '       -Form_FormPlanificacionEquipos.ComandoSacarMeterEnPlanificacion_Click
    '   -Devuelve:
    '       MeterEquipoEnPlanificacion = Descriptivo
    '       MeterEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strIDEquipoEnBD As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El Equipo lo identifica un identificador numérico"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionEquipos", "IDEquipo", "IDEquipo", strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEquipoEnBD = flag
    If Not IsNumeric(strIDEquipoEnBD) Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AltaEquipoEnPlanificacion = Descriptivo
        '       AltaEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = AltaEquipoEnPlanificacion(strIDEquipo, , , 12)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método AltaEquipoEnPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        m_SQL = "UPDATE TbPlanificacionEquipos SET TbPlanificacionEquipos.FechabajaParaPlanificacion = Null " & _
                "WHERE (((TbPlanificacionEquipos.IDEquipo)=" & strIDEquipo & _
                "));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
    End If
    MeterEquipoEnPlanificacion = "Equipo restaurado para planificaciones"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MeterEquipoEnPlanificacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    MeterEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
End Function

