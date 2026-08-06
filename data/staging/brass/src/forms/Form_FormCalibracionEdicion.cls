Option Compare Database
Option Explicit
Dim strIDCalibracion As String
Dim strFechaInicialAlInicio As String, strFechaFinalAlInicio As String, strFechaComunicacionClienteAlInicio As String
Private Function EstablecerDatos() As String
   
    Dim rcdDatos As DAO.Recordset, strEquipo As String, strModelo As String, strNS As String, strBUI As String, _
        strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDCalibracion) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrablesFechas.*, TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS, " & _
                "TbEquiposCalibrables.BUI " & _
                "FROM TbEquiposCalibrablesFechas INNER JOIN TbEquiposCalibrables ON " & _
                "TbEquiposCalibrablesFechas.IDEquipoCalibrable = TbEquiposCalibrables.IDEquipoCalibrable " & _
                "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strIDCalibracion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strFechaInicialAlInicio = Nz(.Fields("FechaCalibrado"), "")
        strFechaFinalAlInicio = Nz(.Fields("FechaFinCalibrado"), "")
        strFechaComunicacionClienteAlInicio = Nz(.Fields("FechaComunicacionCliente"), "")
        strEquipo = Nz(.Fields("Equipo"), "")
        strModelo = Nz(.Fields("Modelo"), "")
        strNS = Nz(.Fields("NS"), "")
        strBUI = Nz(.Fields("BUI"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If IsDate(strFechaInicialAlInicio) Then
        Me.FechaCalibracion = Format(strFechaInicialAlInicio, "dd/mm/yyyy")
    End If
    If IsDate(strFechaFinalAlInicio) Then
        Me.FechaFinCalibracion = Format(strFechaFinalAlInicio, "dd/mm/yyyy")
    End If
    If IsDate(strFechaComunicacionClienteAlInicio) Then
        Me.FechaComunicacionCliente = Format(strFechaComunicacionClienteAlInicio, "dd/mm/yyyy")
    End If
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
     If strModelo <> "" Then
        Me.Modelo = strModelo
    End If
     If strNS <> "" Then
        Me.NS = strNS
    End If
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    EstablecerDatos = "Datos del equipo establecidos correctamente"
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
Private Function Registrar() As String
    Dim strFechaInicial As String, strFechaFinal As String, strFechaComunicacionCliente As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDCalibracion) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strFechaInicial = Nz(Me.FechaCalibracion, "#")
    strFechaFinal = Nz(Me.FechaFinCalibracion, "#")
    strFechaComunicacionCliente = Nz(Me.FechaComunicacionCliente, "#")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EdicionCalibracion = Descriptivo
    '       EdicionCalibracion = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = EdicionCalibracion(strIDCalibracion, strFechaInicial, strFechaFinal, strFechaComunicacionCliente)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EdicionCalibracion ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '------------------------------------------------
    '       HaHabidoCambios = true or false
    '       HaHabidoCambios =False-->Error
    '------------------------------------------------
    If Not HaHabidoCambios() Then
        strTextoError = "No hay cambios que registrar"
        Err.Raise 1000
    End If
    '----------------------------------------------------------------
    '   -Devuelve:
    '       Registrar = strURLAnexo
    '       Registrar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------------
    If FormularioAbierto("FormCalibracionGestion") Then
        Forms("FormCalibracionGestion").Controls("ListaFiltrados").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
    Dim strTextoError As String, strNombreFormulario As String
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
            '-------------------------------------------------------------------
            '       FormularioAbierto = true or false
            '       FormularioAbierto =False-->Error
            '--------------------------------------------
            If FormularioAbierto("FormCalibracionGestion") Then
                Forms("FormCalibracionGestion").Controls("ListaFiltrados").Requery
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
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    strIDCalibracion = Nz(Me.OpenArgs, "")
    Me.Caption = Me.lblTitulo1.Caption
    If Not IsNumeric(strIDCalibracion) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & dato(1)
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
    ' Función creada por Andrés Román del Peral el día  11/06/2012
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '
    '           strEstado="1"----> True Or False
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    Dim strFechaInicial As String, strFechaFinal As String, strFechaComunicacionCliente As String
    On Error GoTo errores
    strFechaInicial = Nz(Me.FechaCalibracion, "")
    strFechaFinal = Nz(Me.FechaFinCalibracion, "")
    strFechaComunicacionCliente = Nz(Me.FechaComunicacionCliente, "")
    If IsDate(strFechaInicial) Then
        strFechaInicial = Format(strFechaInicial, "dd/mm/yyyy")
    End If
    If IsDate(strFechaFinal) Then
        strFechaFinal = Format(strFechaFinal, "dd/mm/yyyy")
    End If
    If IsDate(strFechaComunicacionCliente) Then
        strFechaComunicacionCliente = Format(strFechaComunicacionCliente, "dd/mm/yyyy")
    End If
    If strFechaInicial <> strFechaInicialAlInicio Or strFechaFinal <> strFechaFinalAlInicio Or strFechaComunicacionCliente <> strFechaComunicacionClienteAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Public Function EdicionCalibracion( _
                                    strIDCalibracion As String, _
                                    strFechaCalibrado As String, _
                                    strFechaFinCalibrado As String, _
                                    Optional strFechaComunicacionCliente As String _
                                    ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    '       -12/05/2016: Se incluye el campo strFechaComunicacionCliente
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionEdicion.editar
    '   -Devuelve:
    '       EdicionCalibracion =Descriptivo
    '       EdicionCalibracion ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDCalibracion) Then
        strTextoError = "No se ha indicado la calibración a editar"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaCalibrado) Then
        strTextoError = "Es obligatoria la fecha de calibrado"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaFinCalibrado) Then
        strTextoError = "Es obligatoria la fecha de fin de calibrado"
        Err.Raise 1000
    End If
    If strFechaComunicacionCliente <> "" And strFechaComunicacionCliente <> "#" Then
        If Not IsDate(strFechaComunicacionCliente) Then
            strTextoError = "La fecha de comunicación al cliente no es válida"
            Err.Raise 1000
        End If
    End If
    If CDate(strFechaFinCalibrado) < CDate(strFechaCalibrado) Then
        strTextoError = "La fecha de fin de calibrado ha de ser igual o mayor que la de calibrado"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrablesFechas.* " & _
            "FROM TbEquiposCalibrablesFechas " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strIDCalibracion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado la calibración que se desea editar"
            Err.Raise 1000
        End If
        .Edit
            .Fields("FechaCalibrado") = strFechaCalibrado
            .Fields("FechaFinCalibrado") = strFechaFinCalibrado
            If IsDate(strFechaComunicacionCliente) Then
                .Fields("FechaComunicacionCliente") = strFechaComunicacionCliente
            Else
                If strFechaComunicacionCliente = "#" Then
                    .Fields("FechaComunicacionCliente") = Null
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EdicionCalibracion = "Editado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EdicionCalibracion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EdicionCalibracion = "#ERR" & "|" & strTextoError
End Function





