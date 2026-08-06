Option Compare Database
Option Explicit
Dim strIDEquipoSeleccionado As String
Dim strBUIAnterior As String, strEquipoAnterior As String, strCalibracionAnterior As String, strAmbitoAnterior As String
Public Sub LimpiarUltimos()
    strBUIAnterior = ""
    strEquipoAnterior = ""
    strCalibracionAnterior = ""
    strAmbitoAnterior = ""
End Sub
Public Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarBUI_Click
    ComandoLimpiarEquipo_Click
    ComandoLimpiarCalibracion_Click
    ComandoLimpiarAmbito_Click
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
Public Sub ActualizarLista()
    Dim strTextoError As String, intElementos As Integer, lst As ListBox, VarItem As Variant
    Dim strIDEquipoSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = ""
    Set lst = Me.ListaFiltrados
    strIDEquipoSeleccionadoActualmente = Nz(lst.Column(0), "")
    Me.ComandoEditarEquipo.Enabled = False
    Me.ComandoEliminarEquipo.Enabled = False
    Me.ComandoCalibraciones.Enabled = False
    Me.ComandoExportarExcel.Enabled = False
    Me.lblCalibraciones.Visible = False
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Filtrar=strSQLFinal
    '       Filtrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    flag = Filtrar( _
                    Nz(Me.ComboBui, ""), Nz(Me.ComboEquipos, ""), _
                    Nz(Me.ComboCalibracion, ""), Nz(Me.ComboAmbito, ""))
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
Private Sub ComandoCalibraciones_Click()
    Dim strTextoCalibraciones As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "No se ha seleccionado ningún equipo antes"
        Err.Raise 1000
    End If
    '----------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '----------------------------------------------
    If FormularioAbierto("FormCalibracionGestion") Then
        DoCmd.Close acForm, "FormCalibracionGestion", acSaveNo
    End If
    If Me.lblCalibraciones.Visible = True Then
        strTextoCalibraciones = Nz(Me.lblCalibraciones.Caption, "")
    End If
    DoCmd.OpenForm "FormCalibracionGestion", , , , , , strIDEquipoSeleccionado & "|" & strTextoCalibraciones
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver las calibraciones se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAltaEquipo_Click()
    On Error Resume Next
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto = true or false
    '
    '           strEstado="1"---->true or false
    '           strEstado="-1"--->false
    '-------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionEquipoAlta") Then
        DoCmd.Close acForm, "FormCalibracionEquipoAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormCalibracionEquipoAlta"
End Sub
Private Sub ComandoEditarEquipo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Ha de seleccionar una equipo primero"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto =True or false
    '           strEstado="1"---->True or false
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionEquipoEdicion") Then
        DoCmd.Close acForm, "FormCalibracionEquipoEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormCalibracionEquipoEdicion", , , , , , strIDEquipoSeleccionado
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
Private Sub ComandoEliminarEquipo_Click()
    Dim strNumeroFechasDeCalibracionAsociadas As String, strPregunta As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDEquipoSeleccionado) Then
        strTextoError = "Se ha de elegir un equipo que queramos borrar"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrablesFechas.IDEquipoCalibrable " & _
        "FROM TbEquiposCalibrablesFechas " & _
        "WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable)=" & strIDEquipoSeleccionado & "));"
     '-------------------------------------------------------------
    '   -Devuelve:
    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameNumeroRegistrosPorSQL(m_SQL)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameNumeroRegistrosPorSQL ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strNumeroFechasDeCalibracionAsociadas = flag
    If Not IsNumeric(strNumeroFechasDeCalibracionAsociadas) Then
        strTextoError = "No se ha podido determinar el número de calibraciones asociadas hay que borrar también"
        Err.Raise 1000
    End If
    If strNumeroFechasDeCalibracionAsociadas = "0" Then
        strPregunta = "¿Desea realmente borrar el equipo calibrable seleccionado?"
    ElseIf strNumeroFechasDeCalibracionAsociadas = "1" Then
        strPregunta = "Este equipo tiene asociada una calibración" & vbCrLf & "¿Desea borrar al equipo junto con su calibración?"
    Else
        strPregunta = "Este equipo tiene asociada " & strNumeroFechasDeCalibracionAsociadas & " calibraciones" & vbCrLf & "¿Desea borrar al equipo junto con sus calibraciones?"
    End If
    pregunta = MsgBox(strPregunta, vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Equipo calibrable")
    If pregunta <> vbYes Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EliminarEquipoACalibrar = Descriptivo
    '       EliminarEquipoACalibrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EliminarEquipoACalibrar(strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EliminarEquipoACalibrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.cmdSalir.SetFocus
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar borrar un equipo de calibración se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
Private Sub ComandoLimpiarAmbito_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboAmbito, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboAmbito = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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
Private Sub ComandoLimpiarCalibracion_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboCalibracion, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboCalibracion = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarEquipo_Click()
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
Private Sub ComboAmbito_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboAmbito, "")
    If strValor <> strAmbitoAnterior Then
        ActualizarLista
    End If
    strAmbitoAnterior = strValor
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

Private Sub ComboAmbito_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboAmbito, "")
    If strValor <> strAmbitoAnterior Then
        ActualizarLista
    End If
    strAmbitoAnterior = strValor
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

Private Sub ComboBui_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBui, "")
    If strValor <> strBUIAnterior Then
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

Private Sub ComboCalibracion_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboCalibracion, "")
    If strValor <> strCalibracionAnterior Then
        ActualizarLista
    End If
    strCalibracionAnterior = strValor
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

Private Sub ComboCalibracion_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboCalibracion, "")
    If strValor <> strCalibracionAnterior Then
        ActualizarLista
    End If
    strCalibracionAnterior = strValor
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

Private Sub ComboEquipos_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipos, "")
    If strValor <> strEquipoAnterior Then
        ActualizarLista
    End If
    strEquipoAnterior = strValor
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

Private Sub ComboEquipos_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipos, "")
    If strValor <> strEquipoAnterior Then
        ActualizarLista
    End If
    strEquipoAnterior = strValor
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
    Dim strTextoCalibracion As String, strNumeroCalibraciones As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEquipoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    Me.ComandoEditarEquipo.Enabled = False
    Me.ComandoEliminarEquipo.Enabled = False
    Me.ComandoCalibraciones.Enabled = False
    Me.lblCalibraciones.Visible = False
    Me.ComandoCalibraciones.Caption = "Calibraciones"
    If strIDEquipoSeleccionado = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoEditarEquipo.Enabled = True
    Me.ComandoEliminarEquipo.Enabled = True
    Me.ComandoCalibraciones.Enabled = True
    '-----------------------------------------------
    '   -Devuelve:
    '       strTextoCalibracion = strTexto
    '       strTextoCalibracion = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = DameTextoCalibracion(strIDEquipoSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameTextoCalibracion ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strTextoCalibracion = flag
    If strTextoCalibracion <> "" Then
        Me.lblCalibraciones.Visible = True
        Me.lblCalibraciones.Caption = strTextoCalibracion
    Else
        Me.lblCalibraciones.Visible = False
    End If
    m_SQL = "SELECT TbEquiposCalibrablesFechas.IDEquipoCalibrable " & _
            "FROM TbEquiposCalibrablesFechas " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable)=" & strIDEquipoSeleccionado & "));"
    '-------------------------------------------------------------
    '   -Devuelve:
    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameNumeroRegistrosPorSQL(m_SQL)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameNumeroRegistrosPorSQL ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strNumeroCalibraciones = flag
    If Not IsNumeric(strNumeroCalibraciones) Then
        strNumeroCalibraciones = "#ERR"
    End If
    Me.ComandoCalibraciones.Caption = "Calibraciones (" & strNumeroCalibraciones & ")"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al hacer click sobre la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Filtrar( _
                        Optional strBUI As String, _
                        Optional strEquipo As String, _
                        Optional strCriteriocalibracion As String, _
                        Optional strAmbito As String _
                        ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionEquipoGestion.ActualizarLista
    '   -Devuelve:
    '       Filtrar = strSQLFiltro
    '       Filtrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim strSQLFiltro As String, strParteWhereBUI As String, strParteWhereEquipo As String, strParteWhereCalibracion As String, strParteWhereAmbito As String
    Dim strParteWhere As String, strSQLInicial As String, strTextoError As String
    On Error GoTo errores
    strSQLInicial = "SELECT DISTINCT TbEquiposCalibrables.IDEquipoCalibrable, TbEquiposCalibrables.BUI, TbEquiposCalibrables.Equipo, " & _
                    "TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS, " & _
                    "DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]) AS Vigencia " & _
                    "FROM TbEquiposCalibrables  LEFT JOIN TbEquiposCalibrablesFechas ON " & _
                    "TbEquiposCalibrables.IDEquipoCalibrable = TbEquiposCalibrablesFechas.IDEquipoCalibrable "
    If strBUI = "" Then
        strParteWhereBUI = "((TbEquiposCalibrables.BUI) Like '*' Or (TbEquiposCalibrables.BUI) Is Null)"
    Else
        strParteWhereBUI = "((TbEquiposCalibrables.BUI)='" & strBUI & "')"
    End If
    
     If strEquipo = "" Then
        strParteWhereEquipo = "((TbEquiposCalibrables.Equipo) Like '*' Or (TbEquiposCalibrables.Equipo) Is Null)"
    Else
        strParteWhereEquipo = "((TbEquiposCalibrables.Equipo) Like '*" & strEquipo & "*')"
    End If
    If strAmbito = "Todos" Or strAmbito = "" Then
        strParteWhereAmbito = "((TbEquiposCalibrables.FechaBajaParaCalibracion) Like '*' Or (TbEquiposCalibrables.FechaBajaParaCalibracion) Is Null)"
    ElseIf strAmbito = "Vigentes" Then
        strParteWhereAmbito = "((TbEquiposCalibrables.FechaBajaParaCalibracion) Is Null)"
    ElseIf strAmbito = "Obsoletos" Then
        strParteWhereAmbito = "(Not (TbEquiposCalibrables.FechaBajaParaCalibracion) Is Null)"
    End If
    If strCriteriocalibracion = "" Then
        strParteWhereCalibracion = "((DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable])) Like '*')"
    Else
        If strCriteriocalibracion = "Con Calibración" Or strCriteriocalibracion = "Sin Calibración" Or _
            strCriteriocalibracion = "Con Calibración no vigente" Or strCriteriocalibracion = "Con Calibración vigente" Then
            If strCriteriocalibracion = "Con Calibración" Then
                strParteWhereCalibracion = "((DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]))<>'SC')"
            ElseIf strCriteriocalibracion = "Sin Calibración" Then
                 strParteWhereCalibracion = "((DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]))='SC')"
            ElseIf strCriteriocalibracion = "Con Calibración no vigente" Then
                strParteWhereCalibracion = " ((Left(DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]),1))='-')"
            Else
                strParteWhereCalibracion = "((Left(DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]),1))<>'s' " & _
                                        "And (Left(DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable]),1))<>'-')"
            End If
        Else
             strParteWhereCalibracion = "((DameDiasVigenciaCalibracion([TbEquiposCalibrables].[IDEquipoCalibrable])) Like '*')"
        End If
    End If
    strParteWhere = "WHERE (" & _
                                strParteWhereBUI & " AND " & _
                                strParteWhereCalibracion & " AND " & _
                                strParteWhereEquipo & " AND " & _
                                strParteWhereAmbito & ");"
    strSQLFiltro = strSQLInicial & strParteWhere
    Filtrar = strSQLFiltro
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Filtrar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Filtrar = "#ERR" & "|" & strTextoError
End Function
Public Function DameTextoCalibracion( _
                                        strIDEquipo As String _
                                        ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Si no tiene calibraciones.-strTexto="Equipo sin ninguna calibración registrada"
    '       -Si tiene calibraciones.
    '           -Si strUltimaFechaFinCalibracion<strHoy-->strTexto="Equipo con calibración " &  cstr(intDias) & " fuera de calibración"
    '           -Si strUltimaFechaFinCalibracion>strHoy-->strTexto="Equipo con calibración " &   cstr(intDias) & " hasta que finalice el periodo de calibración"
    '   -Llamada desde
    '       -Form_FormCalibracionGestion.SeleccionarLista
    '   -Devuelve:
    '       DameTextoCalibracion = strTexto
    '       DameTextoCalibracion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim strTexto As String, rcdDatos As DAO.Recordset, strUltimaFechaFinCalibracion As String, intDias As Integer, strFechaBajaParaCalibracion As String
    Dim strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "Se ha de indicar el IDEquipo"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEquiposCalibrables", "FechaBajaParaCalibracion", "IDEquipoCalibrable", strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaBajaParaCalibracion = flag
    If IsDate(strFechaBajaParaCalibracion) Then
        strTexto = "Equipo dado de baja de calibración desde " & strFechaBajaParaCalibracion
        DameTextoCalibracion = strTexto
        Exit Function
    End If
    m_SQL = "SELECT Last(TbEquiposCalibrablesFechas.FechaFinCalibrado) AS ÚltimoDeFechaFinCalibrado " & _
            "FROM TbEquiposCalibrablesFechas " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strUltimaFechaFinCalibracion = Nz(.Fields("ÚltimoDeFechaFinCalibrado"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not IsDate(strUltimaFechaFinCalibracion) Then
        strTexto = "Equipo sin ninguna calibración registrada"
    Else
        intDias = DateDiff("d", Now(), strUltimaFechaFinCalibracion)
        If intDias < 0 Then
            strTexto = "Equipo cuya calibración lleva " & CStr(Abs(intDias)) & " días fuera de vigencia"
        ElseIf intDias = 0 Then
            strTexto = "Equipo con calibración vigente hasta hoy mismo"
        Else
            strTexto = "Equipo con calibración vigente hasta dentro de " & CStr(intDias) & " días"
        End If
    End If
    DameTextoCalibracion = strTexto
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameTextoCalibracion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameTextoCalibracion = "#ERR" & "|" & strTextoError
End Function
Public Function EliminarEquipoACalibrar( _
                                        strIDEquipoCalibrar As String, _
                                        Optional strNumeroFechasDeCalibracionAsociadas As String _
                                        ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionEquipoGestion.ComandoEliminarEquipo_Click
    '   -Devuelve:
    '       EliminarEquipoACalibrar = Descriptivo
    '       EliminarEquipoACalibrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDEquipoCalibrar) Then
        strTextoError = "Se ha de indicar qué equipo de calibración se pretende eliminar"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrables.* " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibrar & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado el equipo que se pretende eliminar"
            Err.Raise 1000
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EliminarEquipoACalibrar = "Equipo eliminado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Calibracion.EliminarEquipoACalibrar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EliminarEquipoACalibrar = "#ERR" & "|" & strTextoError
End Function