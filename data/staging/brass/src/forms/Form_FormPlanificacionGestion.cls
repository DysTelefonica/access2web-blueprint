Option Compare Database
Option Explicit
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoActualizar_Click()
    Dim strAño As String
    Dim intPreventivoEquipos As Integer, intPreventivosEquiposProgramados As Integer, intPreventivosProgramados As Integer, intPreventivosEquiposNoProgramados As Integer, _
        intPreventivosRealizados As Integer, intPreventivosRealizadosEnPlazo As Integer, intPreventivosRealizadosFueraDePlazo As Integer, _
        intPreventivosNoRealizados As Integer, intPreventivosNoRealizadosVencidos As Integer, intPreventivosNoRealizadosEnPlazo As Integer, _
        intPreventivosNoRealizadosProximos As Integer, ctl As Control, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    For Each ctl In Me.Controls
        If ctl.Tag = "CONTADOR" Then ctl.Value = Null
    Next
    strAño = Nz(Me.AÑO, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       CuadroDeMando = lngNumeroRegistros
    '       CuadroDeMando = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = CuadroDeMando( _
                            strAño, colPreventivoEquipos, colPreventivosEquiposProgramados, colPreventivosProgramados, colPreventivosEquiposNoProgramados, _
                            colPreventivosRealizados, colPreventivosRealizadosEnPlazo, colPreventivosRealizadosFueraDePlazo, _
                            colPreventivosNoRealizados, colPreventivosNoRealizadosVencidos, colPreventivosNoRealizadosEnPlazo, _
                            colPreventivosNoRealizadosProximos _
                            )
                                            
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "0" Then
        strTextoError = "No hay registros para ese año"
        Err.Raise 1000
    End If
    '------------------
    ' CONTADORES
    '----------------
        Me.PreventivoEquipos = colPreventivoEquipos.count
        Me.PreventivosEquiposProgramados = colPreventivosEquiposProgramados.count
        Me.PreventivosEquiposNoProgramados = colPreventivosEquiposNoProgramados.count
        Me.PreventivosProgramados = colPreventivosProgramados.count
        Me.PreventivosRealizados = colPreventivosRealizados.count
        Me.PreventivosNoRealizados = colPreventivosNoRealizados.count
        Me.PreventivosRealizadosFueraDePlazo = colPreventivosRealizadosFueraDePlazo.count
        Me.PreventivosNoRealizadosProximos = colPreventivosNoRealizadosProximos.count
        Me.PreventivosRealizadosEnPlazo = colPreventivosRealizadosEnPlazo.count
        Me.PreventivosNoRealizadosVencidos = colPreventivosNoRealizadosVencidos.count
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoBuscarEquiposPlanificables_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormPlanificacionEquipos"
End Sub
Private Sub ComandoBuscarPlanificaciones_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormPlanificacionesBusqueda"
End Sub
Private Sub ComandoInforme_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameListadoPlanificacionesPendientesConIDEventosPosibles = Descriptivo
    '       DameListadoPlanificacionesPendientesConIDEventosPosibles = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameListadoPlanificacionesPendientesConIDEventosPosibles()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameListadoPlanificacionesPendientesConIDEventosPosibles ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = "Select * from TbAuxPlanificacion;"
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
        strTextoError = "Al intentar obtener la consulta se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo.Caption
    Me.AÑO = Year(Now())
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
Private Sub PreventivoEquipos_Click()
    Dim strTextoError As String, strMensaje As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivoEquipos, "Sí", "EQUIPOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosEquiposNoProgramados_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosEquiposNoProgramados, "Sí", "EQUIPOS SIN PROGRAMACIÓN")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    
End Sub
Private Sub PreventivosEquiposProgramados_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosProgramados, "Sí", "EQUIPOS CON PROGRAMACIÓN")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosNoRealizados_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosNoRealizados, "Sí", "NO REALIZADOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosNoRealizadosProximos_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosNoRealizadosProximos, "No", "NO REALIZADOS PRÓXIMOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosNoRealizadosVencidos_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosNoRealizadosVencidos, "No", "NO REALIZADOS VENCIDOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosProgramados_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosRealizados, "No", "PROGRAMADOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub PreventivosRealizados_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosRealizados, "No", "REALIZADOS")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosRealizadosEnPlazo_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosRealizadosEnPlazo, "No", "REALIZADOS EN PLAZO")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub PreventivosRealizadosFueraDePlazo_Click()
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameHTML(colPreventivosRealizadosFueraDePlazo, "No", "REALIZADOS FUERA DE PLAZO")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTML ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strMensaje = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el HTML se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function DameListadoPlanificacionesPendientesConIDEventosPosibles() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar la tabla strNombreTabla y nos va a dar un entero más que tenga el máximo dentro del campo llamado strNombreCampoID
    '   -Llamada por:
    
    '   -Devuelve:
    '       DameListadoPlanificacionesPendientesConIDEventosPosibles = Descriptivo
    '       DameListadoPlanificacionesPendientesConIDEventosPosibles = "#ERR" & "|" & strTextoError
    '---------------------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, colEquiposSinMto As New Collection, strEquipo As String, VarItem As Variant, rcdDestino As DAO.Recordset
    Dim strIDPlanificacion As String, strIDEquipo As String, strFechaPrevista As String, strDias As String, strBUI As String, strSUBSISTEMA As String, _
        strCadenaIDEventosPosibles As String, strCadenaIDEventosUsados As String, strTextoError
    On Error GoTo errores
    m_SQL = "SELECT DISTINCT TbEquipos.IDEquipo " & _
            "FROM (TbPlanificacion INNER JOIN TbPlanificacionEquipos ON TbPlanificacion.IDEquipo = TbPlanificacionEquipos.IDEquipo) " & _
            "INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo " & _
            "WHERE (((TbPlanificacion.FechaCierre) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strEquipo = Nz(.Fields("IDEquipo"), "")
                colEquiposSinMto.Add strEquipo
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If colEquiposSinMto.count = 0 Then
        DameListadoPlanificacionesPendientesConIDEventosPosibles = "Sin Planificaciones"
        Exit Function
    End If
    m_SQL = "delete * from TbAuxPlanificacion;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    m_SQL = "TbAuxPlanificacion"
    Set rcdDestino = CurrentDb().OpenRecordset(m_SQL)
    For Each VarItem In colEquiposSinMto
        strEquipo = CStr(VarItem)
        m_SQL = "SELECT TbPlanificacion.IDPlanificacion , TbEquipos.Equipo,TbEquipos.IDEquipo, TbPlanificacion.FechaPrevistaCierre , " & _
                "DateDiff('d',Now(),[fechaprevistacierre]) AS Días, TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                "FROM (TbPlanificacion INNER JOIN TbPlanificacionEquipos ON TbPlanificacion.IDEquipo = TbPlanificacionEquipos.IDEquipo) " & _
                "INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbPlanificacion.FechaCierre) Is Null) AND ((TbPlanificacion.IDEquipo)=" & strEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        If Not rcdDatos.EOF Then
            rcdDatos.MoveFirst
            Do While Not rcdDatos.EOF
                strIDPlanificacion = Nz(rcdDatos.Fields("IDPlanificacion"), "")
                strEquipo = Nz(rcdDatos.Fields("Equipo"), "")
                strIDEquipo = Nz(rcdDatos.Fields("IDEquipo"), "")
                strFechaPrevista = Nz(rcdDatos.Fields("FechaPrevistaCierre"), "")
                strDias = Nz(rcdDatos.Fields("Días"), "")
                strBUI = Nz(rcdDatos.Fields("BUI"), "")
                strSUBSISTEMA = Nz(rcdDatos.Fields("SUBSISTEMA"), "")
                '---------------------------------------------------------------------------------------------------
                '   -Devuelve:
                '       strCadenaIDEventosPosibles=strCadenaIDEventosPosibles & "|" & strIDEvento & "(" & Format(strFechaFinEvento, "dd/mm/yyyy")
                '       strCadenaIDEventosUsados = strCadenaIDEventosUsados & "|" & strIDEvento & "(" & strIDMto & ")"
                '       DameCadenaIDEventos = strCadenaIDEventosPosibles & "#" & strCadenaIDEventosUsados
                '       DameCadenaIDEventos = "#ERR" & "[" & strTextoError
                '---------------------------------------------------------------------------------------------------
                flag = DameCadenaIDEventos(strIDEquipo)
                If InStr(1, flag, "[") <> 0 Then
                    dato = Split(flag, "[")
                    strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
                    Err.Raise 1000
                End If
                If InStr(1, flag, "#") <> 0 Then
                    dato = Split(flag, "#")
                    strCadenaIDEventosPosibles = dato(0)
                    strCadenaIDEventosUsados = dato(1)
                    rcdDestino.AddNew
                        rcdDestino.Fields("IDPlanificacion") = rcdDatos.Fields("IDPlanificacion")
                        rcdDestino.Fields("Equipo") = rcdDatos.Fields("Equipo")
                        rcdDestino.Fields("FechaPrevista") = rcdDatos.Fields("FechaPrevistaCierre")
                        rcdDestino.Fields("DiasRestan") = rcdDatos.Fields("Días")
                        rcdDestino.Fields("BUI") = rcdDatos.Fields("BUI")
                        rcdDestino.Fields("SUBSISTEMA") = rcdDatos.Fields("SUBSISTEMA")
                        rcdDestino.Fields("CadenaIDEventosPosibles") = strCadenaIDEventosPosibles
                        rcdDestino.Fields("CadenaIDEventosUsados") = strCadenaIDEventosUsados
                    rcdDestino.Update
                End If
                rcdDatos.MoveNext
            Loop
        End If
        rcdDatos.Close
        Set rcdDatos = Nothing
    Next
    DameListadoPlanificacionesPendientesConIDEventosPosibles = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameListadoPlanificacionesPendientesConIDEventosPosibles ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Not rcdDestino Is Nothing Then
        rcdDestino.Close
        Set rcdDestino = Nothing
    End If
    DoCmd.SetWarnings True
    DameListadoPlanificacionesPendientesConIDEventosPosibles = "#ERR" & "|" & strTextoError
End Function
Private Function DameCadenaIDEventos(strIDEquipo As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar la tabla strNombreTabla y nos va a dar un entero más que tenga el máximo dentro del campo llamado strNombreCampoID
    '   -Llamada por:
    
    '   -Devuelve:
    '       strCadenaIDEventosPosibles=strCadenaIDEventosPosibles & "|" & strIDEvento & "(" & Format(strFechaFinEvento, "dd/mm/yyyy")
    '       strCadenaIDEventosUsados = strCadenaIDEventosUsados & "|" & strIDEvento & "(" & strIDMto & ")"
    '       DameCadenaIDEventos = strCadenaIDEventosPosibles & "#" & strCadenaIDEventosUsados
    '       DameCadenaIDEventos = "#ERR" & "[" & strTextoError
    '---------------------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEvento As String, strCadenaIDEventosPosibles As String, strFechaFinEvento As String, strCadenaIDEventosUsados As String, _
        strIDMto As String
    Dim colUsados As New Collection, colPosibles As New Collection, strTextoError As String
    On Error GoTo errores
    m_SQL = "SELECT TbEventos.IDEvento " & _
            "FROM TbPlanificacion RIGHT JOIN TbEventos ON TbPlanificacion.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbEventos.IDEquipo)=" & strIDEquipo & _
            ") AND ((TbPlanificacion.IDPlanificacion) Is Null) AND ((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND (Not (TbEventos.HORAFINALEVENTO) Is Null));"

    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDEvento = .Fields("IDEvento")
                '-------------------------------------------------------------------
                '   -Devuelve:
                '       DameFechaFinEvento = strFechaUltimaActividad
                '       DameFechaFinEvento = #ERR
                '-------------------------------------------------------------------
                flag = DameFechaFinEvento(strIDEvento)
                If InStr(1, flag, "|") = 0 Then
                    strFechaFinEvento = flag
                    If IsDate(strFechaFinEvento) Then
                        If strCadenaIDEventosPosibles = "" Then
                            strCadenaIDEventosPosibles = strIDEvento & "(" & Format(strFechaFinEvento, "dd/mm/yyyy") & ")"
                        Else
                            strCadenaIDEventosPosibles = strCadenaIDEventosPosibles & "|" & strIDEvento & "(" & Format(strFechaFinEvento, "dd/mm/yyyy") & ")"
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbEventos.IDEvento, TbPlanificacion.IDPlanificacion " & _
            "FROM TbPlanificacion RIGHT JOIN TbEventos ON TbPlanificacion.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbEventos.IDEquipo)=" & strIDEquipo & ") AND (Not (TbPlanificacion.IDPlanificacion) Is Null) AND " & _
            "((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND (Not (TbEventos.HORAFINALEVENTO) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDEvento = .Fields("IDEvento")
                strIDMto = .Fields("IDPlanificacion")
                If strCadenaIDEventosUsados = "" Then
                    strCadenaIDEventosUsados = strIDEvento & "(" & strIDMto & ")"
                Else
                    strCadenaIDEventosUsados = strCadenaIDEventosUsados & "|" & strIDEvento & "(" & strIDMto & ")"
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameCadenaIDEventos = strCadenaIDEventosPosibles & "#" & strCadenaIDEventosUsados
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameCadenaIDEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameCadenaIDEventos = "#ERR" & "[" & strTextoError
End Function
Public Function CuadroDeMando( _
                                strAño As String, ByRef colPreventivoEquipos As Collection, ByRef colPreventivosEquiposProgramados As Collection, ByRef colPreventivosProgramados As Collection, _
                                ByRef colPreventivosEquiposNoProgramados As Collection, ByRef colPreventivosRealizados As Collection, ByRef colPreventivosRealizadosEnPlazo As Collection, _
                                ByRef colPreventivosRealizadosFueraDePlazo As Collection, ByRef colPreventivosNoRealizados As Collection, ByRef colPreventivosNoRealizadosVencidos As Collection, _
                                ByRef colPreventivosNoRealizadosEnPlazo As Collection, ByRef colPreventivosNoRealizadosProximos As Collection _
                            ) As String

    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 19/05/2016
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -1) rellena equipos para mantenimiento del año
    '       -2) Recorre los equipos y para cada uno de ellos se van rellenando las colecciones que entran
    '
   
    '   -llamada por:
    
    '   -Devuelve:
    '       CuadroDeMando = lngNumeroRegistros
    '       CuadroDeMando = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim VarItem As Variant, rcdDatos As DAO.Recordset, intDias As Integer, strFechaInicial As String, strFechaFinal As String, strResultado As String, _
        strFechaInicialDeConsulta As String, strFechaFinalDeConsulta As String
    Dim strIDEquipo As String, strFechabajaParaPlanificacion As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, _
        strPeriodicidadEnMesesRecomendada As String
    Dim strMotivoCierre As String, strIDEvento As String, strIDNOEvento As String, strFechaPrevistaCierre As String, strFechaCierre As String, _
        strMotivacionReprogramacion As String, strIDNuevaPlanificacion As String, strIDPlanificacion As String, lngNumeroRegistros As Long, strTextoError As String
    On Error GoTo errores
    '--------------------------
    ' COMPROBAR DATOS INICIALES
    '----------------------------
        If Not IsNumeric(strAño) Or strAño = "0" Then
            strAño = Year(Now())
        End If
        strFechaInicial = "01/01/" & strAño
        strFechaFinal = "31/12/" & strAño
        strFechaInicialDeConsulta = strFechaInicial
        strFechaFinalDeConsulta = "12/31/" & strAño
        Set colPreventivoEquipos = New Collection
        Set colPreventivosEquiposProgramados = New Collection
        Set colPreventivosProgramados = New Collection
        Set colPreventivosEquiposNoProgramados = New Collection
        Set colPreventivosRealizados = New Collection
        Set colPreventivosRealizadosEnPlazo = New Collection
        Set colPreventivosRealizadosFueraDePlazo = New Collection
        Set colPreventivosNoRealizados = New Collection
        Set colPreventivosNoRealizadosVencidos = New Collection
        Set colPreventivosNoRealizadosEnPlazo = New Collection
        Set colPreventivosNoRealizadosProximos = New Collection
    '--------------------------
    ' RELLENAR colPreventivoEquipos
    '----------------------------
        m_SQL = "SELECT TbEquipos.Equipo,TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo, TbPlanificacionEquipos.PeriodicidadEnMesesRecomendada, " & _
                "TbPlanificacionEquipos.FechabajaParaPlanificacion " & _
                "FROM TbPlanificacionEquipos INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo;"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                .MoveFirst
                Do While Not .EOF
                    strEquipo = Nz(.Fields("Equipo"), "")
                    strBUI = Nz(.Fields("BUI"), "")
                    strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                    strPeriodicidadEnMesesRecomendada = Nz(.Fields("PeriodicidadEnMesesRecomendada"), "")
                    strFechabajaParaPlanificacion = Nz(.Fields("FechabajaParaPlanificacion"), "")
                    strResultado = strEquipo & "|" & strBUI & "|" & strSUBSISTEMA & "|" & strEquipo & "|" & strPeriodicidadEnMesesRecomendada
                    If Not IsDate(strFechabajaParaPlanificacion) Then
                        colPreventivoEquipos.Add strResultado
                    Else
                        '--------------------------------------------------------------------------------------------
                        '  SI EL EQUIPO ESTÁ DADO DE BAJA DURANTE EL AÑO NO FORMA PARTE DE LOS PREVENTIVOS DE ESE AÑO
                        '--------------------------------------------------------------------------------------------
                        If CDate(strFechabajaParaPlanificacion) > CDate(strFechaFinal) Then
                            colPreventivoEquipos.Add strResultado
                        End If
                        
                    End If
                    .MoveNext
                Loop
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    If colPreventivoEquipos.count = 0 Then
        CuadroDeMando = "0"
        Exit Function
    End If
    '--------------------------
    ' RELLENAR colPreventivosEquiposProgramados,colPreventivosEquiposNoProgramados PARTIENDO DE colPreventivoEquipos (tiene que tener al menos una programación en ese año)
    '----------------------------
        'strResultado = strEquipo & "|" & strBUI & "|" & strSUBSISTEMA & "|" & strEquipo & "|" & strPeriodicidadEnMesesRecomendada
        For Each VarItem In colPreventivoEquipos
            dato = Split(VarItem, "|")
            strEquipo = CStr(dato(0))
            m_SQL = "SELECT TbPlanificacion.FechaPrevistaCierre " & _
                    "FROM TbPlanificacion " & _
                    "WHERE (((TbPlanificacion.IDEquipo)=" & strIDEquipo & _
                    ") AND ((TbPlanificacion.FechaPrevistaCierre) Between #" & strFechaInicialDeConsulta & "# And #" & strFechaFinalDeConsulta & "#));"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If Not .EOF Then
                    colPreventivosEquiposProgramados.Add VarItem
                Else
                    colPreventivosEquiposNoProgramados.Add VarItem
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        Next
    '--------------------------
    ' RELLENAR RESTO DE COLECCIONES PARTIENDO DE colPreventivosEquiposProgramados
    '----------------------------
        'strResultado = strEquipo & "|" & strBUI & "|" & strSUBSISTEMA & "|" & strEquipo & "|" & strPeriodicidadEnMesesRecomendada
        For Each VarItem In colPreventivosEquiposProgramados
            dato = Split(VarItem, "|")
            strEquipo = CStr(dato(0))
            strBUI = CStr(dato(1))
            strSUBSISTEMA = CStr(dato(2))
            strEquipo = CStr(dato(3))
            strPeriodicidadEnMesesRecomendada = CStr(dato(4))
            m_SQL = "SELECT TbPlanificacion.* " & _
                        "FROM TbPlanificacion " & _
                        "WHERE (((TbPlanificacion.IDEquipo)=" & strIDEquipo & "));"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If Not .EOF Then
                    .MoveFirst
                    Do While Not .EOF
                        strIDPlanificacion = Nz(.Fields("IDPlanificacion"), "")
                        strFechaPrevistaCierre = Nz(.Fields("FechaPrevistaCierre"), "")
                        strFechaCierre = Nz(.Fields("FechaCierre"), "")
                        strIDEvento = Nz(.Fields("IDEVENTO"), "")
                        strIDNOEvento = Nz(.Fields("IDNOEvento"), "")
                        strMotivacionReprogramacion = Nz(.Fields("MotivacionReprogramacion"), "")
                        strIDNuevaPlanificacion = Nz(.Fields("IDNuevaPlanificacion"), "")
                        strMotivoCierre = Nz(.Fields("MotivoCierre"), "")
                        strResultado = CStr(VarItem) & "|" & strIDPlanificacion & "|" & strFechaPrevistaCierre & "|" & strFechaCierre & "|" & _
                                        strMotivoCierre & "|" & strIDEvento & "|" & strIDNOEvento & "|" & strMotivacionReprogramacion & "|" & strIDNuevaPlanificacion
                        
                        If CDate(strFechaPrevistaCierre) >= CDate(strFechaInicial) And CDate(strFechaPrevistaCierre) <= CDate(strFechaFinal) Then
                            colPreventivosProgramados.Add strResultado
                            '---------------------------------------
                            ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                            '-------------------------------------------
                            If IsDate(strFechaCierre) Then
                                '----------------------------------------
                                ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                ' Realizados
                                '-----------------------------------------
                                colPreventivosRealizados.Add strResultado
                                If CDate(strFechaCierre) > CDate(strFechaPrevistaCierre) Then
                                    '----------------------------------------
                                    ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                    ' Realizados fuera de plazo
                                    '-----------------------------------------
                                    colPreventivosRealizadosFueraDePlazo.Add strResultado
                                Else
                                    '----------------------------------------
                                    ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                    ' Realizados en plazo
                                    '-----------------------------------------
                                    colPreventivosRealizadosEnPlazo.Add strResultado
                                    
                                End If
                            Else
                                '----------------------------------------
                                ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                ' NO Realizados
                                '-----------------------------------------
                                colPreventivosNoRealizados.Add strResultado
                                intDias = DateDiff("d", Now(), strFechaPrevistaCierre)
                                If intDias >= 0 And intDias <= 30 Then
                                    '----------------------------------------
                                    ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                    ' NO Realizados Próximos
                                    '-----------------------------------------
                                    colPreventivosNoRealizadosProximos.Add strResultado
                                ElseIf intDias > 30 Then
                                    '----------------------------------------
                                    ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                    ' NO Realizados No Vencidos y no próximos
                                    '-----------------------------------------
                                    
                                Else
                                    '----------------------------------------
                                    ' LOS PREVENTIVOS SON DEL AÑO SOLICITADO
                                    ' NO Realizados Vencidos
                                    '-----------------------------------------
                                    colPreventivosNoRealizadosVencidos.Add strResultado
                                End If
                            End If
                        
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        Next
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosEquiposNoProgramados.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosProgramados.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosProgramados.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosRealizadosFueraDePlazo.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosRealizadosEnPlazo.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosNoRealizados.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosNoRealizadosProximos.count
        lngNumeroRegistros = lngNumeroRegistros + colPreventivosNoRealizadosVencidos.count
    CuadroDeMando = CStr(lngNumeroRegistros)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método CuadroDeMando ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    CuadroDeMando = "#ERR" & "|" & strTextoError
End Function

