Option Compare Database
Option Explicit
Dim strIDPlanificacion As String
Private Sub ComandoCierre_Click()
    Dim strIDEvento As String, strIDNOEvento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se puede determinar la planificación que se pretende cerrar"
        Err.Raise 1000
    End If
    strIDEvento = Nz(Me.ComboIDEvento, "")
    strIDNOEvento = Nz(Me.ComboIDNoEvento, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       CerrarPlanificacionConTrabajos = Descriptivo
    '       CerrarPlanificacionConTrabajos ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = CerrarPlanificacionConTrabajos(strIDPlanificacion, strIDEvento, strIDNOEvento)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método CerrarPlanificacionConTrabajos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar cerrar una planifiación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strIDEvento As String, strEquipo As String, strBUI As String, strSUBSISTEMA As String, _
        strIDEquipo As String, strIDNOEvento As String, strMes As String, strSemana As String, strAño As String, strFechaCierre As String, _
        strFechaPrevistaCierre As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.BUI = strBUI
    Me.SubSistema = strSUBSISTEMA
    Me.Equipo = strEquipo
    m_SQL = "SELECT  TbPlanificacion.Anio, TbPlanificacion.Mes, " & _
            "TbPlanificacion.Semana, TbPlanificacion.FechaCierre, TbPlanificacion.FechaPrevistaCierre, TbPlanificacion.IDEvento, " & _
            "TbPlanificacion.IDNoEvento, TbPlanificacion.IDEquipo " & _
            "FROM TbPlanificacion INNER JOIN TbEquipos ON TbPlanificacion.IDEquipo = TbEquipos.IDEquipo " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
            Err.Raise 1000
        End If
        strAño = Nz(.Fields("Anio"), "")
        strMes = Nz(.Fields("Mes"), "")
        strSemana = Nz(.Fields("Semana"), "")
        strIDEvento = Nz(.Fields("IDEvento"), "")
        strIDNOEvento = Nz(.Fields("IDNoEvento"), "")
        strFechaCierre = Nz(.Fields("FechaCierre"), "")
        strFechaPrevistaCierre = Nz(.Fields("FechaPrevistaCierre"), "")
        strIDEquipo = Nz(.Fields("IDEquipo"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not IsNumeric(strAño) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.AÑO = strAño
    If Not IsNumeric(strMes) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.Mes = strMes
    If Not IsNumeric(strSemana) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.Semana = strSemana
    If IsDate(strFechaPrevistaCierre) Then
        Me.FechaPrevistaCierre = Format(strFechaPrevistaCierre, "dd/mm/yyyy")
    End If
    If strIDEvento <> "" Then
        Me.ComboIDEvento = strIDEvento
    End If
    If strIDNOEvento <> "" Then
        Me.ComboIDNoEvento = strIDNOEvento
    End If
    If IsDate(strFechaCierre) Then
        Me.lblFechaCierre.Visible = True
        Me.FechaCierre.Visible = True
        Me.FechaCierre = Format(strFechaCierre, "dd/mm/yyyy")
        Me.ComandoCierre.Enabled = False
    Else
        Me.lblFechaCierre.Visible = False
        Me.FechaCierre.Visible = False
        Me.ComandoCierre.Enabled = True
    End If
    m_SQL = "SELECT TbEventos.IDEvento, TbEventos.TIPOEVENTO, DameFechaFinEvento([TbEventos].[IDEvento]) as FECHA_EVENTO " & _
                "FROM (TbEventos LEFT JOIN TbPlanificacion ON TbEventos.IDEvento = TbPlanificacion.IDEvento) " & _
                "LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND ((TbEventos.BUI)='" & strBUI & _
                "') AND ((TbEventos.SUBSISTEMA)='" & strSUBSISTEMA & _
                "') AND ((TbEquipos.Equipo)='" & strEquipo & _
                "') AND ((TbPlanificacion.IDEvento) Is Null));"
    Me.ComboIDEvento.RowSource = m_SQL
    Me.ComboIDEvento.Requery
    m_SQL = "SELECT TbPlanificacionRegistrada.ID, TbPlanificacionRegistrada.FechaMantoRealizado " & _
                "FROM TbPlanificacionRegistrada LEFT JOIN TbPlanificacion ON TbPlanificacionRegistrada.ID = TbPlanificacion.IDNoEvento " & _
                "WHERE (((TbPlanificacionRegistrada.IDEquipo)=" & strIDEquipo & "));"
    Me.ComboIDNoEvento.RowSource = m_SQL
    Me.ComboIDNoEvento.Requery
    Me.ComandoCierre.Enabled = False
    m_SQL = "SELECT TbEventos.IDEvento, TbPlanificacion.IDPlanificacion " & _
            "FROM TbPlanificacion RIGHT JOIN TbEventos ON TbPlanificacion.IDEvento = TbEventos.IDEvento " & _
            "WHERE (((TbEventos.IDEquipo)=" & strIDEquipo & ") AND (Not (TbPlanificacion.IDPlanificacion) Is Null) AND " & _
            "((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND (Not (TbEventos.HORAFINALEVENTO) Is Null));"
    Me.ComboIDEventosUsados.RowSource = m_SQL
    Me.ComboIDEventosUsados.Requery
    EstablecerDatos = "Datos OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoVerProgramacion_Click()
    Dim strIDProgramacion As String, strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    strIDProgramacion = Nz(Me.ComboIDEventosUsados.Column(1), "")
    If IsNumeric(strIDProgramacion) Then
        DoCmd.Close acForm, strNombreFormulario, acSaveNo
        DoCmd.OpenForm "FormPlanificacionCierre", , , , , , strIDProgramacion
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver la programación se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboIDEvento_Exit(Cancel As Integer)
    On Error Resume Next
    If Nz(Me.ComboIDEvento, "") = "" And Nz(Me.ComboIDNoEvento, "") = "" Then
        Me.ComandoCierre.Enabled = False
    Else
        Me.ComandoCierre.Enabled = True
    End If
End Sub
Private Sub ComboIDEventosUsados_Change()
    Dim strIDProgramacion As String
    On Error Resume Next
    strIDProgramacion = Nz(Me.ComboIDEventosUsados.Column(1), "")
    If IsNumeric(strIDProgramacion) Then
        Me.ComandoVerProgramacion.Enabled = True
    Else
        Me.ComandoVerProgramacion.Enabled = False
    End If
End Sub
Private Sub ComboIDNoEvento_Exit(Cancel As Integer)
    On Error Resume Next
    If Nz(Me.ComboIDEvento, "") = "" And Nz(Me.ComboIDNoEvento, "") = "" Then
        Me.ComandoCierre.Enabled = False
    Else
        Me.ComandoCierre.Enabled = True
    End If
End Sub
Private Sub Form_Load()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    strIDPlanificacion = Nz(Me.OpenArgs(), "")
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    '---------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = Descriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '---------------------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
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








