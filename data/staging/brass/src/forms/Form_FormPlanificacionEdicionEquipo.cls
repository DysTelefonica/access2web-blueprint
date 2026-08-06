Option Compare Database
Option Explicit
Dim strIDEquipo As String, strAliasAlInicio As String, strPeriodicidadEnMesesRecomendadaAlInicio As String
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    If Not HaHabidoCambios() Then
        strTextoError = "Ningún cambio que grabar"
        Err.Raise 1000
    End If
     '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strDescriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar  se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strAlias As String, strPeriodicidadEnMesesRecomendada As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "Se ha abierto el formulario con datos insuficientes"
        Err.Raise 1000
    End If
    strAlias = Nz(Me.Alias, "#")
    strPeriodicidadEnMesesRecomendada = Nz(Me.PERIODICIDADMESES, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EditarEquipo = Descriptivo
    '       EditarEquipo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EditarEquipo(strIDEquipo, strAlias, CInt(strPeriodicidadEnMesesRecomendada))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EditarEquipo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Function

Private Sub cmdSalir_Click()
   Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    If HaHabidoCambios() Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = Descriptivo
            '       Registrar = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Load()
    Dim rcdDatos As DAO.Recordset, strNombreEquipo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    strIDEquipo = Nz(Me.OpenArgs(), "")
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
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
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 19/5/2017
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-------------------------------------------------------------------
    Dim strAlias As String, strPeriodicidadEnMesesRecomendada As String
    On Error GoTo errores
    strAlias = Nz(Me.Alias, "")
    strPeriodicidadEnMesesRecomendada = Nz(Me.PERIODICIDADMESES, "")
    If strAlias <> strAliasAlInicio Or strPeriodicidadEnMesesRecomendada <> strPeriodicidadEnMesesRecomendadaAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Public Function EditarEquipo( _
                                strIDEquipo As String, _
                                Optional strAlias As String, _
                                Optional intNumeroMesesMto As Integer _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 19/05/2017
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va marcar un equipo como planificable metiéndolo en la tabla TbPlanificacionEquipos
    '
    
    '   -llamada por:
    
    '   -Devuelve:
    '       EditarEquipo = Descriptivo
    '       EditarEquipo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El Equipo lo identifica un identificador numérico"
        Err.Raise 1000
    End If
    If intNumeroMesesMto < 1 Then
        strTextoError = "La peridicidad es en meses"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionEquipos.* " & _
            "FROM TbPlanificacionEquipos " & _
            "WHERE (((TbPlanificacionEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "Equipo planificable no registrado"
            Err.Raise 1000
        End If
        .Edit
            If strAlias <> "" And strAlias <> "#" Then
                .Fields("Alias") = strAlias
            Else
                If strAlias = "#" Then
                    .Fields("Alias") = Null
                End If
            End If
            .Fields("PeriodicidadEnMesesRecomendada") = intNumeroMesesMto
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EditarEquipo = "Equipo planificable editado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EditarEquipo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EditarEquipo = "#ERR" & "|" & strTextoError
End Function
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strNombreEquipo As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionEquipos.Alias, TbPlanificacionEquipos.PeriodicidadEnMesesRecomendada, TbEquipos.Equipo " & _
            "FROM TbPlanificacionEquipos INNER JOIN TbEquipos ON TbPlanificacionEquipos.IDEquipo = TbEquipos.IDEquipo " & _
            "WHERE (((TbPlanificacionEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "Equipo planificable no registrado"
            Err.Raise 1000
        End If
        strNombreEquipo = Nz(.Fields("Equipo"), "")
        strAliasAlInicio = Nz(.Fields("Alias"), "")
        strPeriodicidadEnMesesRecomendadaAlInicio = Nz(.Fields("PeriodicidadEnMesesRecomendada"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Me.ComboEquipoPlanificable = strNombreEquipo
    Me.Alias = strAliasAlInicio
    Me.PERIODICIDADMESES = strPeriodicidadEnMesesRecomendadaAlInicio
    EstablecerDatos = "Datos del IDEquipo establecidos correctamente"
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


