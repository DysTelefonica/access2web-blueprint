Option Compare Database
Option Explicit
Dim strIDPlanificacion As String
Private Sub EstablecerBotonera()
    Dim ctl As Control
    On Error GoTo errores
    For Each ctl In Me.Controls
        If ctl.Tag = "OBLIGATORIO" Then
            If Nz(ctl.Value, "") = "" Then
                Me.ComandoReplanificar.Enabled = False
                Exit Sub
            End If
        End If
    Next
    Me.ComandoReplanificar.Enabled = True
    Exit Sub
errores:
    Me.ComandoReplanificar.Enabled = False
    
End Sub
Public Sub EstablecerListaSemana()
    Dim strAño As String, strMes As String, strCadenaSemana As String
    Dim strTextoError As String
    On Error GoTo errores
    strAño = Nz(Me.AñoNuevo, "")
    strMes = Nz(Me.ComboMesNuevo, "")
    If Not IsNumeric(strAño) Or Not IsNumeric(strMes) Then
        strCadenaSemana = "#ERR"
    Else
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       FechaViernes =dteFechaViernesSemanax para intSemana=x <>0
        '       FechaViernes =strCadenaFechas para intSemana=0
        '       strCadenaFechas=01 & dteFechaViernesSemana1 & vbcrlf & _
                                02 & dteFechaViernesSemana2 & vbcrlf & _
        '                       ...
        '                       intNumeroSemanas & dteFechaViernesSemanax
        '       FechaViernes="#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = FechaViernes(CInt(strMes), CLng(strAño))
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método FechaViernes ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strCadenaSemana = flag
        If strCadenaSemana = "-1" Then
            strTextoError = "No se ha podido cartar el combo de las semanas del mes"
            Err.Raise 1000
        End If
        strCadenaSemana = Replace(strCadenaSemana, vbCrLf, ";")
    End If
    Me.ComboSemanaNueva.RowSource = strCadenaSemana
    Me.ComboSemanaNueva.Requery
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar cargar el combo de las semanas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
   
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub AñoNuevo_Exit(Cancel As Integer)
     On Error Resume Next
     EstablecerListaSemana
     EstablecerBotonera
End Sub
Private Sub ComandoReplanificar_Click()
    Dim strMotivo As String, strAño As String, strMes As String, strSemana As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se puede determinar la planificación que se pretende cerrar"
        Err.Raise 1000
    End If
    strMotivo = Nz(Me.MotivoReplanificacion, "")
    strAño = Nz(Me.AñoNuevo, "")
    strMes = Nz(Me.ComboMesNuevo, "")
    strSemana = Nz(Me.ComboSemanaNueva, "")
    If Len(strSemana) < 3 Then
         strTextoError = "No se ha introducido una válida"
        Err.Raise 1000
    End If
    If Not IsNumeric(Left(strSemana, 2)) Then
        strTextoError = "No se ha introducido una válida"
        Err.Raise 1000
    End If
    strSemana = Left(strSemana, 2)
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       CerrarPlanificacionConTrabajos = Descriptivo
    '       CerrarPlanificacionConTrabajos ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = CerrarPlanificacionSinTrabajos(strIDPlanificacion, strMotivo, CInt(strSemana), CInt(strMes), CLng(strAño))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método CerrarPlanificacionConTrabajos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormPlanificacionesBusqueda") Then
        Forms("FormPlanificacionesBusqueda").Controls("ListaFiltrados").Requery
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
    
    Dim rcdDatos As DAO.Recordset, strEquipo As String, strBUI As String, strSUBSISTEMA As String
    Dim strMes As String, strSemana As String, strAño As String, strFechaCierre As String, strFechaPrevistaCierre As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.BUI = strBUI
    Me.SubSistema = strSUBSISTEMA
    Me.Equipo = strEquipo
    m_SQL = "SELECT TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo, TbPlanificacion.Anio, TbPlanificacion.Mes, " & _
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
        strBUI = Nz(.Fields("BUI"), "")
        strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
        strEquipo = Nz(.Fields("Equipo"), "")
        strAño = Nz(.Fields("Anio"), "")
        strMes = Nz(.Fields("Mes"), "")
        strSemana = Nz(.Fields("Semana"), "")
        strFechaCierre = Nz(.Fields("FechaCierre"), "")
        strFechaPrevistaCierre = Nz(.Fields("FechaPrevistaCierre"), "")
        strEquipo = Nz(.Fields("Equipo"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
     If strBUI = "" Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.BUI = strBUI
    If strSUBSISTEMA = "" Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.SubSistema = strSUBSISTEMA
    If strEquipo = "" Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.Equipo = strEquipo
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
    Me.ComandoReplanificar.Enabled = False
    EstablecerDatos = "Datos establecidos correctamente"
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
Private Sub ComboMesNuevo_Change()
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerListaSemana
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComboMesNuevo_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerListaSemana
    EstablecerBotonera
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComboSemanaNueva_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerBotonera
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub Form_Load()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
   AjustarTamaño Me
    Me.Caption = Me.lblTitulo.Caption
    'strArgumento=strIDPlanificacion
    strIDPlanificacion = Nz(Me.OpenArgs(), "")
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    EstablecerDatos
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub MotivoReplanificacion_Exit(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    EstablecerBotonera
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
