Option Compare Database
Option Explicit
Dim strIDPlanificacion As String, strIDEventoCierre As String, strIDEquipo As String, strFechaPrevistaCierre As String
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoRegistrar_Click()
    Dim strIDEventoNuevo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDEventoNuevo = Nz(Me.ComboIDEventoNuevo, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       CerrarPlanificacionCambioTrabajo = strDescriptivo
    '       CerrarPlanificacionCambioTrabajo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = CerrarPlanificacionCambioTrabajo(strIDPlanificacion, strIDEventoNuevo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método CerrarPlanificacionCambioTrabajo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '--------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
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
        strTextoError = "Al intentar cambiar el trabajo de cierre se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")

End Sub
Private Sub ComandoVerEventoCierre_Click()
    Dim strIDEvento As String, strTextoError As String
    On Error GoTo errores
    strIDEvento = Nz(Me.IDEVENTO, "")
    If strIDEvento = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> strIDEvento Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el Evento se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoVerNuevoEventoCierre_Click()
    Dim strIDEvento As String, strTextoError As String
    On Error GoTo errores
    strIDEvento = Nz(Me.ComboIDEventoNuevo, "")
    If strIDEvento = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> strIDEvento Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el Evento se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Load()
    Dim strTextoError As String, strArgumento As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    'strArgumento=strIDPlanificacion & "|" & strIDEventoCierre & "|" & strIDEquipo & "|" & strFechaPrevistaCierre
    Me.Caption = Me.lblTitulo1.Caption
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 3 Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    strIDPlanificacion = dato(0)
    strIDEventoCierre = dato(1)
    strIDEquipo = dato(2)
    strFechaPrevistaCierre = dato(3)
    If Not IsNumeric(strIDPlanificacion) Or strIDEventoCierre = "" Or Not IsNumeric(strIDEquipo) Or Not IsDate(strFechaPrevistaCierre) Then
        strTextoError = "No se ha abierto el formulario con los parámetros adecuados"
        Err.Raise 1000
    End If
    Me.IDEVENTO = strIDEventoCierre
    Me.FechaPrevistaCierre = Format(strFechaPrevistaCierre, "dd/mm/yyyy")
    m_SQL = "SELECT TbEventos.IDEvento, TbEventos.TIPOEVENTO, TbEventos.FECHAALTAEVENTO " & _
            "FROM TbEventos LEFT JOIN TbPlanificacion ON TbEventos.IDEvento = TbPlanificacion.IDEvento " & _
            "WHERE (((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND ((TbEventos.IDEquipo)=" & strIDEquipo & _
            ") AND ((TbPlanificacion.IDEvento) Is Null));"
    Me.ComboIDEventoNuevo.RowSource = m_SQL
    Me.ComboIDEventoNuevo.Requery
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
Public Function CerrarPlanificacionCambioTrabajo( _
                                                strIDPlanificacion As String, _
                                                strIDEventoNuevo As String _
                                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    
    
    '       2) ACTUACIÓN
    
   
    '   -llamada por:
    
    '   -Devuelve:
    '       CerrarPlanificacionCambioTrabajo = Descriptivo
    '       CerrarPlanificacionCambioTrabajo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaCierre As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "Se ha de introducir un valor correcto de la planificación a cerrar"
        Err.Raise 1000
    End If
    If strIDEventoNuevo = "" Then
        strTextoError = "Ha de introducir el evento que va a ser permutado"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbPlanificacion", "IDEvento", "IDEvento", strIDEventoNuevo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Ese evento ya ha sido usado para cerrar otra planificación anteriormente."
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEventos", "TIPOEVENTO", "IDEvento", strIDEventoNuevo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "MANT. PROGRAMADO" Then
        strTextoError = "El evento ha de ser del tipo MANT. PROGRAMADO"
        Err.Raise 1000
    End If
    '----------------------------------------------------------------
    '   -Devuelve:
    '       DameFechaFinEvento = strFechaUltimaActividad
    '       DameFechaFinEvento = #ERR
    '----------------------------------------------------------------
    flag = DameFechaFinEvento(strIDEventoNuevo)
    If flag = "#ERR" Then
        strTextoError = "No se ha podido determinar la fecha fin del Evento"
        Err.Raise 1000
    End If
    If Not IsDate(flag) Then
        strTextoError = "El evento ha de estar cerrado"
        Err.Raise 1000
    End If
    strFechaCierre = flag
    m_SQL = "SELECT TbPlanificacion.* " & _
            "FROM TbPlanificacion " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "La Planificacion con ID: " & strIDPlanificacion & " no ha sido grabada en el Sismeta"
            Err.Raise 1000
        End If
        .Edit
            .Fields("FechaCierre") = Format(strFechaCierre, "dd/mm/yyyy")
            .Fields("IDEvento") = strIDEventoNuevo
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    CerrarPlanificacionCambioTrabajo = "Cambio de evento que cierra la planificación a Fecha " & strFechaCierre
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método CerrarPlanificacionCambioTrabajo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    CerrarPlanificacionCambioTrabajo = "#ERR" & "|" & strTextoError
End Function







