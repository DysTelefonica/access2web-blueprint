Option Compare Database
Option Explicit
Dim strIDEquipoSeleccionado As String
Dim strBUIAnterior As String, strSubSistemaAnterior As String, strIDEquipoAnterior As String, strEnMantenimientoAnterior As String
Public Sub LimpiarUltimos()
    strBUIAnterior = ""
    strSubSistemaAnterior = ""
    strIDEquipoAnterior = ""
    strEnMantenimientoAnterior = ""
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
    ComandoLimpiarEnMantenimiento_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAltaEquipo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEquipoAlta") Then
        DoCmd.Close acForm, "FormEquipoAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormEquipoAlta"
    If Me.ListaFiltrados.ListCount > 1 Then
        Forms("FormEquipoAlta").Controls("ComboNodo") = Nz(Me.ListaFiltrados.Column(5, 1), "")
        Form_FormEquipoAlta.ComboNodo_BeforeUpdate False
        Forms("FormEquipoAlta").Controls("ComboBUI") = Nz(Me.ListaFiltrados.Column(1, 1), "")
        Form_FormEquipoAlta.ComboBui_BeforeUpdate False
        Forms("FormEquipoAlta").Controls("ComboSubsistema") = Nz(Me.ListaFiltrados.Column(2, 1), "")
        Forms("FormEquipoAlta").Controls("Equipo").SetFocus
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta un gasto se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub ComandoEditar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Seleccione el gasto que pretende editar"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEquipoEdicion") Then
        DoCmd.Close acForm, "FormEquipoEdicion", acSaveNo
    End If
    'strArgumento=strIDEquipo
    DoCmd.OpenForm "FormEquipoEdicion", , , , , , strIDEquipoSeleccionado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDEquipoSeleccionado = "" Then
        strTextoError = "Se ha de seleccionar el IDEquipo que desea eliminar"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEventos", "IDEvento", "IDEquipo", strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Al menos forma parte de un Evento"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacion", "IDPlanificacion", "IDEquipo", strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Al menos forma parte de una Planificación"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbMaterial", "IDMaterial", "IDEquipo", strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Al menos hay registrado un material de este equipo"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar el equipo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Equipo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = "SELECT TbEquipos.* " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipoSeleccionado & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe Equipo con IDEquipo: " & strIDEquipoSeleccionado
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el IDEquipo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEstablecerParaPlanificacion_Click()
    Dim strEquipoSeleccionado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(3), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Se ha de elegir un equipo de la lista"
        Err.Raise 1000
    End If
    '-------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormPlanificacionAltaEquipo") Then
        DoCmd.Close acForm, "FormPlanificacionAltaEquipo", acSaveNo
    End If
    DoCmd.OpenForm "FormPlanificacionAltaEquipo", , , , , , strIDEquipoSeleccionado & "|" & strEquipoSeleccionado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar Establecer para Planificación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoHistoricoDeEquipo_Click()
    Dim strEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strEquipo = Nz(Me.ListaFiltrados.Column(3), "")
    If strEquipo = "" Then
        strTextoError = "Seleccione un equipo"
        Err.Raise 1000
    End If
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Form_FormEventoGestion.ComandoActualizarTodo_Click
    Else
        DoCmd.OpenForm "FormEventoGestion"
    End If
    Forms("FormEventoGestion").Controls("ComboEquipoBusqueda").Value = strEquipo
    Form_FormEventoGestion.ComboEquipoBusqueda_AfterUpdate
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los eventos a los que ha formado parte se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    EstablecerCombo Me.ComboSubsistemaBusqueda
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarEnMantenimiento_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEnMantenimientoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEnMantenimientoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarIDEquipo_Click()
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
    EstablecerCombo Me.ComboEquipoBusqueda
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoMantenimientosRealizados_Click()
    Dim strEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strEquipo = Nz(Me.ListaFiltrados.Column(3), "")
    If strEquipo = "" Then
        strTextoError = "Seleccione un equipo"
        Err.Raise 1000
    End If
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Form_FormEventoGestion.ComandoActualizarTodo_Click
    Else
        DoCmd.OpenForm "FormEventoGestion"
    End If
    Forms("FormEventoGestion").Controls("ComboEquipoBusqueda").Value = strEquipo
    Forms("FormEventoGestion").Controls("ComboTipoEventoBusqueda").Value = "MANT. PROGRAMADO"
    Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los eventos a los que ha formado parte se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    If strValor <> strBUIAnterior Then
        EstablecerCombo Me.ComboSubsistemaBusqueda
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

Private Sub ComboBUIBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBUIBusqueda, "")
    If strValor <> strBUIAnterior Then
        EstablecerCombo Me.ComboSubsistemaBusqueda
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

Private Sub ComboEnMantenimientoBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEnMantenimientoBusqueda, "")
    If strValor <> strEnMantenimientoAnterior Then
        ActualizarLista
    End If
    strEnMantenimientoAnterior = strValor
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

Private Sub ComboEnMantenimientoBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEnMantenimientoBusqueda, "")
    If strValor <> strEnMantenimientoAnterior Then
        ActualizarLista
    End If
    strEnMantenimientoAnterior = strValor
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

Private Sub ComboIDEquipoBusqueda_Change()
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

Private Sub ComboEquipoBusqueda_Exit(Cancel As Integer)
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

Private Sub ComboSubsistemaBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaAnterior Then
        EstablecerCombo Me.ComboEquipoBusqueda
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

Private Sub ComboSubsistemaBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaAnterior Then
        EstablecerCombo Me.ComboEquipoBusqueda
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
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    m_SQL = "SELECT TbEquipos.Equipo " & _
                "FROM TbEquipos " & _
                "WHERE (((TbEquipos.SUBSISTEMA)='-1')) " & _
                "ORDER BY TbEquipos.Equipo;"
    Me.ComboEquipoBusqueda.RowSource = m_SQL
    Me.ComboEquipoBusqueda.Requery
    m_SQL = "SELECT TbSubsistemaBui.SUBSISTEMA " & _
            "FROM TbSubsistemaBui " & _
            "WHERE (((TbSubsistemaBui.BUI)='-1')) " & _
            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    Me.ComboSubsistemaBusqueda.RowSource = m_SQL
    Me.ComboSubsistemaBusqueda.Requery
    ActualizarLista
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Sub ListaFiltrados_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDEquipoSeleccionado = "" Then
        Me.ComandoEditar.Enabled = False
        Me.ComandoEliminar.Enabled = False
        Me.ComandoHistoricoDeEquipo.Enabled = False
        Me.ComandoEstablecerParaPlanificacion.Enabled = False
        Me.ComandoMantenimientosRealizados.Enabled = False
        Me.ComandoExportarExcel.Enabled = False
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEditar.Enabled = True
    Me.ComandoEliminar.Enabled = True
    Me.ComandoHistoricoDeEquipo.Enabled = True
    Me.ComandoEstablecerParaPlanificacion.Enabled = True
    Me.ComandoMantenimientosRealizados.Enabled = True
    
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
Public Function Filtrar( _
                            Optional strBUI As String, _
                            Optional strSUBSISTEMA As String, _
                            Optional strEquipo As String, _
                            Optional strEnMantenimiento As String _
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
    Dim strParteWhereBUI As String, strParteWhereSubSistema As String, strParteWhereEquipo As String, _
        strParteWhereEnMantenimiento As String, strTextoError As String
    On Error GoTo errores
    strSQLParteInicial = "SELECT TbEquipos.IDEquipo,  TbEquipos.BUI, " & _
                        "TbEquipos.SUBSISTEMA, TbEquipos.Equipo AS Equipo, TbPlanificacionEquipos.IDEquipo, TbEquipos.NODO " & _
                        "FROM TbEquipos LEFT JOIN TbPlanificacionEquipos ON " & _
                        "TbEquipos.IDEquipo = TbPlanificacionEquipos.IDEquipo "
    If strBUI = "" And strSUBSISTEMA = "" And strEquipo = "" And strEnMantenimiento = "" Then
        strParteWhere = "WHERE (((TbEquipos.IDEquipo)=-1));"
        strSQLFinal = strSQLParteInicial & strParteWhere
        Filtrar = strSQLFinal
        Exit Function
    End If
    If strSUBSISTEMA = "" Then
        strParteWhereSubSistema = "((TbEquipos.SUBSISTEMA) Like '*' Or (TbEquipos.SUBSISTEMA) Is Null)"
    Else
        strParteWhereSubSistema = "((TbEquipos.SUBSISTEMA) = '" & strSUBSISTEMA & "')"
    End If
    If strBUI = "" Then
        strParteWhereBUI = "((TbEquipos.BUI) Like '*' Or (TbEquipos.BUI) Is Null)"
    Else
        strParteWhereBUI = "((TbEquipos.BUI) = '" & strBUI & "')"
    End If
    If strEquipo = "" Then
        strParteWhereEquipo = "((TbEquipos.Equipo) Like '*' Or (TbEquipos.Equipo) Is Null)"
    Else
        strParteWhereEquipo = "((TbEquipos.Equipo) Like '*" & strEquipo & "*')"
    End If
    If strEnMantenimiento <> "" Then
        strParteWhereEnMantenimiento = "((TbPlanificacionEquipos.IDEquipo) Like '*' Or (TbPlanificacionEquipos.IDEquipo) Is Null)"
    Else
        If strEnMantenimiento = "Sí" Then
            strParteWhereEnMantenimiento = "(Not(TbPlanificacionEquipos.IDEquipo) Is Null)"
        Else
            strParteWhereEnMantenimiento = "((TbPlanificacionEquipos.IDEquipo) Is Null)"
        End If
    End If
    strParteWhere = "WHERE (" & _
                            strParteWhereBUI & _
                            " AND " & strParteWhereSubSistema & _
                            " AND " & strParteWhereEquipo & _
                            " AND " & strParteWhereEnMantenimiento & _
                            ");"
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
    Me.ComandoEditar.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Me.ComandoHistoricoDeEquipo.Enabled = False
    Me.ComandoEstablecerParaPlanificacion.Enabled = False
    Me.ComandoMantenimientosRealizados.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
   '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = Filtrar( _
                    Nz(Me.ComboBUIBusqueda, ""), Nz(Me.ComboSubsistemaBusqueda, ""), _
                    Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.ComboEnMantenimientoBusqueda, ""))
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
Public Sub EstablecerCombo(cbo As ComboBox)
    
    Dim strSQLCombo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If cbo.Name = Me.ComboSubsistemaBusqueda.Name Then
        strBUI = Nz(Me.ComboBUIBusqueda, "")
        If strBUI = "" Then
            strBUI = "-1"
        End If
        cbo.Value = Null
        Me.ComboSubsistemaBusqueda = Null
        Me.ComboEquipoBusqueda = Null
        strSQLCombo = "SELECT TbSubsistemaBui.SUBSISTEMA " & _
                    "FROM TbSubsistemaBui " & _
                    "WHERE (((TbSubsistemaBui.BUI)='" & strBUI & "')) " & _
                    "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
        cbo.RowSource = strSQLCombo
        cbo.Requery
        If cbo.ListCount = 1 Then
            cbo.Value = cbo.Column(0, 0)
        End If
    ElseIf cbo.Name = Me.ComboEquipoBusqueda.Name Then
        strBUI = Nz(Me.ComboBUIBusqueda, "")
        strSUBSISTEMA = Nz(Me.ComboSubsistemaBusqueda, "")
        If strBUI = "" Then
            strBUI = "-1"
        End If
        If strSUBSISTEMA = "" Then
            strSUBSISTEMA = "-1"
        End If
        cbo.Value = Null
        Me.ComboEquipoBusqueda = Null
        strSQLCombo = "SELECT TbEquipos.Equipo " & _
                        "FROM TbEquipos " & _
                        "WHERE (((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & _
                        "') AND ((TbEquipos.BUI)='" & strBUI & "')) " & _
                        "ORDER BY TbEquipos.Equipo;"
        cbo.RowSource = strSQLCombo
        cbo.Requery
        If cbo.ListCount = 1 Then
            cbo.Value = cbo.Column(0, 0)
        End If
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

