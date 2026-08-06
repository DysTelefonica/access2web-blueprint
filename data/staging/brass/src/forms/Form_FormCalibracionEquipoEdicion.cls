Option Compare Database
Option Explicit
Dim strIDEquipoCalibrable As String
Dim strEquipoAlInicio As String, strModeloAlInicio As String, strNSAlinicio As String, strBUIAlInicio As String, strObservacionesAlInicio As String, _
    strFechaBajaParaCalibracionAlInicio As String
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipoCalibrable) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrables.* " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibrable & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strEquipoAlInicio = Nz(.Fields("Equipo"), "")
        strModeloAlInicio = Nz(.Fields("Modelo"), "")
        strNSAlinicio = Nz(.Fields("NS"), "")
        strBUIAlInicio = Nz(.Fields("BUI"), "")
        strObservacionesAlInicio = Nz(.Fields("Observaciones"), "")
        strFechaBajaParaCalibracionAlInicio = Nz(.Fields("FechaBajaParaCalibracion"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEquipoAlInicio <> "" Then
        Me.Equipo = strEquipoAlInicio
    End If
    If strModeloAlInicio <> "" Then
        Me.Modelo = strModeloAlInicio
    End If
    If strNSAlinicio <> "" Then
        Me.NS = strNSAlinicio
    End If
    If strBUIAlInicio <> "" Then
        Me.ComboBui = strBUIAlInicio
    End If
     If strObservacionesAlInicio <> "" Then
        Me.Observaciones = strObservacionesAlInicio
    End If
    If IsDate(strFechaBajaParaCalibracionAlInicio) Then
        Me.FechaBajaParaCalibracion = Format(strFechaBajaParaCalibracionAlInicio, "dd/mm/yyyy")
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
    Dim strFechaBajaParaCalibracion As String, strEquipo As String, strModelo As String, _
        strNS As String, strBUI As String, strObservaciones As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipoCalibrable) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strEquipo = Nz(Me.Equipo, "#")
    strModelo = Nz(Me.Modelo, "#")
    strNS = Nz(Me.NS, "#")
    strBUI = Nz(Me.ComboBui, "#")
    strFechaBajaParaCalibracion = Nz(Me.FechaBajaParaCalibracion, "#")
    strObservaciones = Nz(Me.Observaciones, "#")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EdicionEquipoACalibrar = strDescriptivo
    '       EdicionEquipoACalibrar = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------
    flag = EdicionEquipoACalibrar(strIDEquipoCalibrable, strBUI, strEquipo, strModelo, strNS, _
                                strObservaciones, strFechaBajaParaCalibracion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EdicionEquipoACalibrar ha devuelto el error: " & vbNewLine & dato(1)
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
Private Sub ComandoCalibraciones_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDEquipoCalibrable) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    '------------------------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionGestion") Then
        DoCmd.Close acForm, "FormCalibracionGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormCalibracionGestion", , , , , , strIDEquipoCalibrable
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
    If FormularioAbierto("FormCalibracionEquipoGestion") Then
        Form_FormCalibracionEquipoGestion.ActualizarLista
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
            If FormularioAbierto("FormCalibracionEquipoGestion") Then
                Form_FormCalibracionEquipoGestion.ActualizarLista
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
    Me.Caption = Me.lblTitulo1.Caption
    strIDEquipoCalibrable = Nz(Me.OpenArgs, "")
    If Not IsNumeric(strIDEquipoCalibrable) Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       URLCompletaArchivoAyuda = strURLAnexo
    '       URLCompletaArchivoAyuda = "#ERR" & "|" & strTextoError
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
    Dim strEquipo As String, strModelo As String, strNS As String, strBUI As String, strObservaciones As String, _
        strFechaBajaParaCalibracion As String
    On Error GoTo errores
    strEquipo = Nz(Me.Equipo, "")
    strModelo = Nz(Me.Modelo, "")
    strNS = Nz(Me.NS, "")
    strBUI = Nz(Me.ComboBui, "")
    strObservaciones = Nz(Me.Observaciones, "")
    strFechaBajaParaCalibracion = Nz(Me.FechaBajaParaCalibracion, "")
    If IsDate(strFechaBajaParaCalibracion) Then
        strFechaBajaParaCalibracion = Format(strFechaBajaParaCalibracion, "dd/mm/yyyy")
    End If
    If strEquipo <> strEquipoAlInicio Or strModelo <> strModeloAlInicio Or _
        strNS <> strNSAlinicio Or strBUI <> strBUIAlInicio Or strObservaciones <> strObservacionesAlInicio Or _
        strFechaBajaParaCalibracion <> strFechaBajaParaCalibracionAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Public Function EdicionEquipoACalibrar( _
                                        strIDEquipoCalibrar As String, _
                                        Optional strBUI As String, _
                                        Optional strEquipo As String, _
                                        Optional strModelo As String, _
                                        Optional strNS As String, _
                                        Optional strObservaciones As String, _
                                        Optional strFechaBajaParaCalibracion As String _
                                        ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionEquipoEdicion.editar
    '   -Devuelve:
    '       EdicionEquipoACalibrar = Descriptivo
    '       EdicionEquipoACalibrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDEquipoCalibrar) Then
        strTextoError = "Se ha de indicar qué equipo de calibración se pretende hacer algún cambio"
        Err.Raise 1000
    End If
    If strBUI = "" Or strBUI = "#" Then
        strTextoError = "Es obligatorio indicar el BUI"
        Err.Raise 1000
    End If
    If strEquipo = "" Or strEquipo = "#" Then
        strTextoError = "Es obligatorio indicar el Equipo"
        Err.Raise 1000
    End If
    If strFechaBajaParaCalibracion <> "" And strFechaBajaParaCalibracion <> "#" Then
        If Not IsDate(strFechaBajaParaCalibracion) Then
            strTextoError = "La fecha de baja de calibración no es válida"
            Err.Raise 1000
        End If
    End If
   m_SQL = "SELECT TbEquiposCalibrables.IDEquipoCalibrable " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE ((Not (TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibrar & _
            ") AND ((TbEquiposCalibrables.BUI)='" & strBUI & _
            "') AND ((TbEquiposCalibrables.Equipo)='" & strEquipo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Para " & strBUI & " el equipo " & strEquipo & " ya ha sido registrado previamente"
            Err.Raise 1000
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbEquiposCalibrables.* " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibrar & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado el equipo que se pretende editar"
            Err.Raise 1000
        End If
        .Edit
            .Fields("Equipo") = strEquipo
            .Fields("BUI") = strBUI
            If strModelo <> "" And strModelo <> "#" Then
                .Fields("Modelo") = strModelo
            Else
                If strModelo = "#" Then
                     .Fields("Modelo") = Null
                End If
            End If
            If strNS <> "" And strNS <> "#" Then
                .Fields("NS") = strNS
            Else
                If strNS = "#" Then
                    .Fields("NS") = Null
                End If
            End If
            If strObservaciones <> "" And strObservaciones <> "#" Then
                .Fields("Observaciones") = strObservaciones
            Else
                If strObservaciones = "#" Then
                     .Fields("Observaciones") = Null
                End If
            End If
            If IsDate(strFechaBajaParaCalibracion) Then
                .Fields("FechaBajaParaCalibracion") = strFechaBajaParaCalibracion
            Else
                If strFechaBajaParaCalibracion = "#" Then
                    .Fields("FechaBajaParaCalibracion") = Null
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EdicionEquipoACalibrar = "Equipo editado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Calibracion.EdicionEquipoACalibrar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EdicionEquipoACalibrar = "#ERR" & "|" & strTextoError
End Function



