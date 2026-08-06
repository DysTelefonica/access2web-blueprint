Option Compare Database
Option Explicit
Dim strIDPlanificacionSeleccionada As String
Dim strCadenaURLAnexos As String, strURLAnexoEvento As String, strIDEventoLoCierra As String
Dim strBUIBusquedaAnterior As String, strSubSistemaBusquedaAnterior As String, strIDBusquedaAnterior As String, _
    strIDEquipoAnterior  As String, strCriterioPlanificacionBusquedaAnterior As String, strFechaFinalBusquedaAnterior As String, strFormularioQueAbre As String
Public Sub ActualizarLista()
    Dim intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant
    Dim strIDPlanificacionSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDPlanificacionSeleccionada = ""
    Me.ComandoAdjuntos.Enabled = False
    Me.ComandoCambiarEventoDeCierre.Enabled = False
    Me.ComandoInformeEnExcel.Enabled = False
    Me.ComandoMarcarComoRealizada.Enabled = False
    Me.ComandoReprogramar.Enabled = False
    Me.ComandoTrabajoQueLaCierra.Enabled = False
    Me.ComandoVerDetalle.Enabled = False
    strIDPlanificacionSeleccionadoActualmente = Nz(Me.ListaFiltrados.Column(0), "")
    '-----------------------------------------------
    '   -Devuelve:
    '       Eliminar = strDescriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Filtrar(Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.ComboSubsistemaBusqueda, ""), _
                            Nz(Me.ComboBUIBusqueda, ""), Nz(Me.FechaInicialBusqueda, ""), Nz(Me.FechaFinalBusqueda, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Filtrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = flag
    Set lst = Me.ListaFiltrados
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 1 Then
        strIDPlanificacionSeleccionadoActualmente = lst.Column(0, 1)
        lst.Selected(1) = True
        ListaFiltrados_Click
    Else
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDPlanificacionSeleccionadoActualmente = lst.Column(0, i)
                ListaFiltrados_Click
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
        strTextoError = "Al intentar actualizar la lista  se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Sub LimpiarUltimos()
    strBUIBusquedaAnterior = ""
    strSubSistemaBusquedaAnterior = ""
    strIDBusquedaAnterior = ""
    strIDEquipoAnterior = ""
    strCriterioPlanificacionBusquedaAnterior = ""
    strFechaFinalBusquedaAnterior = ""
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoAltaPlanificacion_Click()
    Dim strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strEquipo As String, strArgumento As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Desde aquí sólo se puede generar una planificación a partir de otra existente"
        Err.Raise 1000
    End If
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipoSeleccionado & "|" & strEquipo
    strBUI = Nz(Me.ListaFiltrados.Column(7), "")
    strSUBSISTEMA = Nz(Me.ListaFiltrados.Column(8), "")
    strIDEquipo = Nz(Me.ListaFiltrados.Column(9), "")
    strEquipo = Nz(Me.ListaFiltrados.Column(1), "")
    strArgumento = strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipo & "|" & strEquipo
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
        strTextoError = "Al intentar crear una nueva planificación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntos_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "No se puede conocer la planificación seleccionada"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormPlanificacionAnexos") Then
        DoCmd.Close acForm, "FormPlanificacionAnexos", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionAnexos", , , , , , strIDPlanificacionSeleccionada
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar anexar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                            Optional strID As String, _
                            Optional strIDEquipo As String, _
                            Optional strSUBSISTEMA As String, _
                            Optional strBUI As String, _
                            Optional strFechaInicial As String, _
                            Optional strFechaFinal As String, _
                            Optional strEstadoPlanificacion As String _
                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 30/09/2019
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strFechaInicialParaConsulta As String, strFechaFinalParaConsulta As String
    Dim strSQLParteInicial As String, strParteWhere As String, strMotivoCierre As String
    Dim strParteWhereFechas As String, strParteWhereBUI As String, strParteWhereSubSistema As String, strParteWhereEquipo As String, _
        strParteWhereMotivoCierre As String, strParteWhereDifFechas As String, strParteWhereID As String, strSQLFinal As String, _
        strTextoError As String
    On Error GoTo errores
    If strEstadoPlanificacion = "" Then
        strEstadoPlanificacion = "Todas"
    End If
    
   If IsDate(strFechaInicial) Then
        strFechaInicialParaConsulta = Format(strFechaInicial, "mm/dd/yyyy")
    Else
        strFechaInicialParaConsulta = "01/01/1900"
    End If
    If IsDate(strFechaFinal) Then
        strFechaFinalParaConsulta = Format(strFechaFinal, "mm/dd/yyyy")
    Else
        strFechaFinalParaConsulta = "01/01/2200"
    End If
    strParteWhereFechas = "((TbPlanificacion.FechaPrevistaCierre) Between #" & strFechaInicialParaConsulta & "# And #" & strFechaFinalParaConsulta & "#)"
    
    If strBUI <> "" Then
        strParteWhereBUI = "((TbEquipos.BUI)='" & strBUI & "')"
    Else
        strParteWhereBUI = " ((TbEquipos.BUI) Like '*' Or (TbEquipos.BUI) Is Null)"
    End If
    If strSUBSISTEMA <> "" Then
        strParteWhereSubSistema = "((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & "')"
    Else
        strParteWhereSubSistema = " ((TbEquipos.SUBSISTEMA) Like '*' Or (TbEquipos.SUBSISTEMA) Is Null)"
    End If
    If strIDEquipo <> "" Then
        strParteWhereEquipo = "((TbEquipos.Equipo)='" & strIDEquipo & "')"
    Else
        strParteWhereEquipo = " ((TbEquipos.Equipo) Like '*' Or (TbEquipos.Equipo) Is Null)"
    End If
    If strID <> "" Then
        strParteWhereID = "((TbPlanificacion.IDPlanificacion)=" & strID & ")"
    Else
        strParteWhereID = " ((TbPlanificacion.IDPlanificacion) Like '*' Or (TbPlanificacion.IDPlanificacion) Is Null)"
    End If
    If strEstadoPlanificacion = "Realizadas" Then
        If strMotivoCierre <> "" Then
            strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre)='" & strMotivoCierre & "')"
        Else
            strParteWhereMotivoCierre = "(Not (TbPlanificacion.MotivoCierre) Is Null)"
        End If
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0)) Like '*')"
    ElseIf strEstadoPlanificacion = "No Realizadas" Then
        strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0)) Like '*')"
    ElseIf strEstadoPlanificacion = "Fuera de Plazo" Then
        strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0))<0)"
    ElseIf strEstadoPlanificacion = "Cumplen Próximamente" Then
        strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0)) Between 0 And " & CInt(m_ObjEntorno.DiasParaAvisoDePlanificaciones) & ")"
        
    ElseIf strEstadoPlanificacion = "En Plazo" Then
        strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0))>0)"
    ElseIf strEstadoPlanificacion = "Totales" Then
        strParteWhereMotivoCierre = "((TbPlanificacion.MotivoCierre) Like '*' Or (TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0)) Like '*')"
        
    End If
    strSQLParteInicial = "SELECT TbPlanificacion.IDPlanificacion AS ID, TbEquipos.Equipo, " & _
                        "TbPlanificacion.FechaPrevistaCierre AS [Fecha Prevista], TbPlanificacion.FechaCierre AS FCierre, " & _
                        "TbPlanificacion.MotivoCierre, TbPlanificacion.IDNuevaPlanificacion AS IDNueva, " & _
                        "IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0) AS [Días Restan], " & _
                        "TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbPlanificacion.IDEquipo, TbPlanificacion.IDEvento " & _
                        "FROM (TbPlanificacion INNER JOIN TbPlanificacionEquipos ON TbPlanificacion.IDEquipo = TbPlanificacionEquipos.IDEquipo) " & _
                        "INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo "
    strParteWhere = "WHERE (" & _
                                strParteWhereFechas & " AND " & _
                                strParteWhereBUI & " AND " & _
                                strParteWhereSubSistema & " AND " & _
                                strParteWhereEquipo & " AND " & _
                                strParteWhereMotivoCierre & " AND " & _
                                strParteWhereID & " AND " & _
                                strParteWhereDifFechas & ");"
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


Private Sub ComandoCambiarEventoDeCierre_Click()
    Dim strIDEquipo As String, strFechaPrevistaCierre As String
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipo = Nz(Me.ListaFiltrados.Column(9), "")
    strFechaPrevistaCierre = Nz(Me.ListaFiltrados.Column(2), "")
    '------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------------
    If FormularioAbierto("formPlanificacionCambioIDEvento") Then
        DoCmd.Close acForm, "formPlanificacionCambioIDEvento", acSaveNo
    End If
    DoCmd.OpenForm "formPlanificacionCambioIDEvento", , , , , , strIDPlanificacionSeleccionada & "|" & strIDEventoLoCierra & "|" & strIDEquipo & "|" & strFechaPrevistaCierre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoInformeEnExcel_Click()
    Dim strSQLFinal As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-----------------------------------------------
    '   -Devuelve:
    '       GenerarConsultaEnExcel = strDescriptivo
    '       GenerarConsultaEnExcel = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = GenerarConsultaEnExcel
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método GenerarConsultaEnExcel ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strSQLFinal = flag
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el informe  se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub ComandoLimpiarBUI_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboBUIBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboBUIBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarCriterios_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboCriterioPlanificacion, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboCriterioPlanificacion = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarEquipo_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEquipoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarFechas_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.FechaInicialBusqueda, "") = "" And Nz(Me.FechaFinalBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.FechaInicialBusqueda = Null
    Me.FechaFinalBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarID_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEquipoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarSubSistema_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboSubsistemaBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboSubsistemaBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoMarcarComoRealizada_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Se ha de seleccionar una planificación"
        Err.Raise 1000
    End If
   
    If FormularioAbierto("FormPlanificacionCierre") Then
        DoCmd.Close acForm, "FormPlanificacionCierre", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionCierre", , , , , , strIDPlanificacionSeleccionada
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar marcar como realizada se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoReprogramar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Se ha de seleccionar una planificación"
        Err.Raise 1000
    End If
   
    If FormularioAbierto("FormPlanificacionReprogramacion") Then
        DoCmd.Close acForm, "FormPlanificacionReprogramacion", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionReprogramacion", , , , , , strIDPlanificacionSeleccionada
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar marcar para reprogramar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoTrabajoQueLaCierra_Click()
    Dim strIDEvento As String, strIDActividad As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDPlanificacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Se ha de seleccionar un elemento de la lista"
        Err.Raise 1000
    End If
     '-------------------------------------------------------------------
     '   -Devuelve:
    '      EstadoPlanificacion=strEstado
    '      EstadoPlanificacion="-1" Si error
    '       strEstado="Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '-------------------------------------------------------------------
    flag = EstadoPlanificacion(strIDPlanificacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstadoPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "Cerrada por trabajo realizado" Then
        strTextoError = "Esta planificación no está cerrada por trabajo realizado"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbPlanificacion", "IDEvento", "IDPlanificacion", strIDPlanificacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEvento = flag
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbPlanificacion", "IDNoEvento", "IDPlanificacion", strIDPlanificacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDActividad = flag
    If strIDEvento = "" And strIDActividad = "" Then
        strTextoError = "No se ha podido determinar qué evento o qué actividad no de envento, cierra esta planificación"
        Err.Raise 1000
    End If
    If strIDEvento <> "" Then
        '-------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '-------------------------------------------------
        If FormularioAbierto("FormEventoGestion") Then
            DoCmd.Close acForm, "FormEventoGestion", acSaveNo
        End If
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        '-------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '-------------------------------------------------
        If FormularioAbierto("FormActividadGestion") Then
            DoCmd.Close acForm, "FormActividadGestion", acSaveNo
        End If
        DoCmd.OpenForm "FormActividadGestion"
        Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value = strIDActividad
        Form_FormActividadGestion.ComboIDActividadBusqueda_Change
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el trabajo que cierra una planificación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    Me.ComboEquipoBusqueda = Null
    Me.ComboSubsistemaBusqueda = Null
    strBUI = Nz(Me.ComboBUIBusqueda, "-1")
    m_SQL = "SELECT DISTINCT TbEquipos.SUBSISTEMA " & _
            "FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.IDEquipo " & _
            "WHERE (((TbEquipos.BUI)='" & strBUI & "')) ORDER BY SUBSISTEMA;"
    
    Me.ComboSubsistemaBusqueda.RowSource = m_SQL
    Me.ComboSubsistemaBusqueda.Requery
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
Private Sub CargarIDEquipo()
    Dim strTextoError As String, strBUI As String, strSUBSISTEMA As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipoBusqueda = Null
    strBUI = Nz(Me.ComboBUIBusqueda, "")
    strSUBSISTEMA = Nz(Me.ComboSubsistemaBusqueda, "")
    If strBUI = "" Or strSUBSISTEMA = "" Then
        m_SQL = "SELECT TbEquipos.Equipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                "FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.IDEquipo " & _
                "WHERE (((TbPlanificacionEquipos.FechabajaParaPlanificacion) Is Null)) " & _
                "ORDER BY TbEquipos.Equipo;"
    Else
        m_SQL = "SELECT TbEquipos.Equipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                "FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.IDEquipo " & _
                "WHERE (((TbPlanificacionEquipos.FechabajaParaPlanificacion) Is Null) AND ((TbEquipos.BUI)='" & strBUI & _
                "') AND ((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & "')) " & _
                "ORDER BY TbEquipos.Equipo;"
    End If
    Me.ComboEquipoBusqueda.RowSource = m_SQL
    Me.ComboEquipoBusqueda.Requery
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

Private Sub ComandoVerDetalle_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Se ha de seleccionar una planificación"
        Err.Raise 1000
    End If
    '---------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------------
    If FormularioAbierto("FormPlanificacionCierre") Then
        DoCmd.Close acForm, "FormPlanificacionCierre", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionCierre", , , , , , strIDPlanificacionSeleccionada
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar marcar como realizada se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboBUIBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBUIBusqueda, "")
    If strValor <> strBUIBusquedaAnterior Then
        ActualizarLista
        CargarSubSistema
        CargarIDEquipo
    End If
    strBUIBusquedaAnterior = strValor
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
Private Sub ComboCriterioPlanificacion_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboCriterioPlanificacion, "")
    If strValor <> strCriterioPlanificacionBusquedaAnterior Then
        ActualizarLista
    End If
    strCriterioPlanificacionBusquedaAnterior = strValor
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
Private Sub ComboEquipoBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipoBusqueda, "")
    If strValor <> strIDEquipoAnterior Then
        ActualizarLista
    End If
    strIDEquipoAnterior = strValor
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
Private Sub ComboIDMantenimiento_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIDMantenimiento, "")
    If strValor <> strIDBusquedaAnterior Then
        ActualizarLista
    End If
    strIDBusquedaAnterior = strValor
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
Private Sub ComboSubsistemaBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaBusquedaAnterior Then
        ActualizarLista
        CargarIDEquipo
    End If
    strSubSistemaBusquedaAnterior = strValor
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
Private Sub FechaFinalBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.FechaFinalBusqueda, "")
    If strValor <> strFechaFinalBusquedaAnterior Then
        ActualizarLista
    End If
    strFechaFinalBusquedaAnterior = strValor
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
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") Then
        dato = Split(strArgumento, "|")
        If UBound(dato) <> 2 Then
            strTextoError = "Se ha abierto el formulario con datos insuficientes"
            Err.Raise 1000
        End If
        strBUI = Nz(dato(0), "")
        strSUBSISTEMA = Nz(dato(1), "")
        strIDEquipo = Nz(dato(2), "")
    End If
    If strBUI <> "" Then
        Me.ComboBUIBusqueda = strBUI
        ComboBUIBusqueda_Change
    End If
    If strSUBSISTEMA <> "" Then
        Me.ComboSubsistemaBusqueda = strSUBSISTEMA
        ComboSubsistemaBusqueda_Change
    End If
    If strIDEquipo <> "" Then
        Me.ComboEquipoBusqueda = strIDEquipo
        ComboEquipoBusqueda_Change
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar entrar en el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImagenAnexoEvento_Click()
    Dim m_Evento As Evento
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If InStr(1, strURLAnexoEvento, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexoEvento, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        If strIDEventoLoCierra <> "" Then
            Set m_Evento = Constructor.getEvento(strIDEventoLoCierra, strTextoError)
            If strTextoError <> "" Then
                Err.Raise 1000
            End If
            
            If FormularioAbierto("FormAnexos") Then
                DoCmd.Close acForm, "FormAnexos", acSaveNo
            End If
            strFormularioQueAbre = ""
            Set m_ObjEntidadParaAnexoActiva = m_Evento
            'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ImagenDocAnexado_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strCadenaURLAnexos = "" Then
        strTextoError = "No hay ningún documento que abrir"
        Err.Raise 1000
    End If
    If InStr(1, strCadenaURLAnexos, ";") = 0 Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AbrirEnLocal = strURLAnexoLocal
        '       AbrirEnLocal = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = AbrirEnLocal(strURLAnexoEvento, Me.hWnd)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        If Not IsNumeric(strIDPlanificacionSeleccionada) Then
            strTextoError = "No se puede determinar la planificación seleccionada"
            Err.Raise 1000
        End If
        '-------------------------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '---------------------------------------------------
        If FormularioAbierto("FormPlanificacionAnexos") Then
            DoCmd.Close acForm, "FormPlanificacionAnexos", acSaveNo
        End If
        DoCmd.OpenForm "FormPlanificacionAnexos", , , , , , strIDPlanificacionSeleccionada
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ListaFiltrados_Click()
    Dim m_Evento As Evento
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComandoAdjuntos.Enabled = False
    Me.ComandoCambiarEventoDeCierre.Enabled = False
    Me.ComandoInformeEnExcel.Enabled = False
    Me.ComandoMarcarComoRealizada.Enabled = False
    Me.ComandoReprogramar.Enabled = False
    Me.ComandoTrabajoQueLaCierra.Enabled = False
    Me.ComandoVerDetalle.Enabled = False
    strIDPlanificacionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEventoLoCierra = Nz(Me.ListaFiltrados.Column(10), "")
    If Not IsNumeric(strIDPlanificacionSeleccionada) Then
        strTextoError = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    Me.ComandoAdjuntos.Enabled = True
    Me.ComandoVerDetalle.Enabled = True
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       strEstado = "Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '       EstadoPlanificacion=strEstado
    '       EstadoPlanificacion="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EstadoPlanificacion(strIDPlanificacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstadoPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Abierta" Then
        Me.ComandoTrabajoQueLaCierra.Enabled = False
        Me.ComandoMarcarComoRealizada.Enabled = True
        Me.ComandoReprogramar.Enabled = True
    ElseIf flag = "Cerrada por trabajo realizado" Then
        Me.ComandoTrabajoQueLaCierra.Enabled = True
        Me.ComandoMarcarComoRealizada.Enabled = False
        Me.ComandoReprogramar.Enabled = False
    ElseIf flag = "Cerrada por reprogramación" Then
        Me.ComandoTrabajoQueLaCierra.Enabled = True
        Me.ComandoMarcarComoRealizada.Enabled = False
        Me.ComandoReprogramar.Enabled = False
    Else
        strTextoError = "El estado de la planificación no está tipificado"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       strURLCadenaAnexos=strURLAnexo1 & ";" & strURLAnexo2 & ";" & .... & strURLAnexoN
    '       DameCadenaAnexos = strURLCadenaAnexos
    '       DameCadenaAnexos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameCadenaAnexos(strIDPlanificacionSeleccionada)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameCadenaAnexos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strCadenaURLAnexos = flag
    If strCadenaURLAnexos = "" Then
        Me.ImagenDocAnexado.Visible = False
    Else
        Me.ImagenDocAnexado.Visible = True
    End If
    If strIDEventoLoCierra = "" Then
        Me.ImagenAnexoEvento.Visible = False
        Me.ComandoCambiarEventoDeCierre.Enabled = False
    Else
        Me.ComandoCambiarEventoDeCierre.Enabled = True
        Set m_Evento = Constructor.getEvento(strIDEventoLoCierra, strTextoError)
        If strTextoError <> "" Then
            Err.Raise 1000
        End If
        
        strURLAnexoEvento = m_Evento.URLAnexos
        If strURLAnexoEvento = "" Then
            Me.ImagenAnexoEvento.Visible = False
        Else
            Me.ImagenAnexoEvento.Visible = True
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click en la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function GenerarConsultaEnExcel() As String
    Dim strWhereFiltro As String, strSQLInicial As String
    Dim strSQLCompleto As String, strSQLFiltro As String, strTextoError As String
    On Error GoTo errores
    strSQLFiltro = Me.ListaFiltrados.RowSource
    strSQLInicial = "SELECT TbPlanificacion.IDPlanificacion AS ID, TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo, " & _
                    "TbPlanificacion.FechaPrevistaCierre AS [Fecha Prevista], TbPlanificacion.FechaCierre AS FCierre, " & _
                    "TbPlanificacion.MotivoCierre, TbPlanificacion.IDNuevaPlanificacion AS IDNueva, " & _
                    "IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0) AS [Días Restan], " & _
                    "TbPlanificacion.IDEvento AS [Evento Cierra] " & _
                    "FROM (TbPlanificacion INNER JOIN TbPlanificacionEquipos ON TbPlanificacion.IDEquipo = TbPlanificacionEquipos.IDEquipo) " & _
                    "INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo "
    If InStr(1, strSQLFiltro, "Where") <> 0 Then
        dato = Split(strSQLFiltro, "Where")
        strWhereFiltro = "Where " & dato(1)
    Else
        strWhereFiltro = ";"
    End If
    strSQLCompleto = strSQLInicial & strWhereFiltro
    GenerarConsultasSQL strSQLCompleto, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    GenerarConsultaEnExcel = strSQLCompleto
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método GenerarConsultaEnExcel ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    GenerarConsultaEnExcel = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarBUI_Click
    ComandoLimpiarSubSistema_Click
    ComandoLimpiarID_Click
    ComandoLimpiarFechas_Click
    ComandoLimpiarEquipo_Click
    ComandoLimpiarCriterios_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar las listas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


