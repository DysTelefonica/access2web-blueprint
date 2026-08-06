Option Compare Database
Option Explicit
Dim strCodActividadAlInicio As String, strTipoActividadAlInicio As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  29/01/13
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
    Dim strTipoActividad As String, strCodActividad As String
    On Error GoTo errores
    strTipoActividad = Nz(Me.TIPOACTIVIDAD, "")
    strCodActividad = Nz(Me.CodTipo, "")
    If strTipoActividad <> strTipoActividadAlInicio Or strCodActividad <> strCodActividadAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
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
            If FormularioAbierto("FormActividadTipoGestion") Then
                Forms("FormActividadTipoGestion").Controls("ListaTiposDeActividad").Requery
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
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    '--------------------------------------------------
    '   -Devuelve
    '       HaHabidoCambios = true or false
    '       HaHabidoCambios =true-->Error
    '--------------------------------------------------
    If Not HaHabidoCambios() Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("No ha habido cambios que guardar", vbCritical, "Sin cambios")
        Exit Sub
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strEstaFacturado
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '--------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------------
    If FormularioAbierto("FormActividadTipoGestion") Then
        Forms("FormActividadTipoGestion").Controls("ListaTiposDeActividad").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    '--------------------------------------------------------
    Dim strTipoActividad As String, strTextoError As String
    On Error GoTo errores
    If strCodActividadAlInicio = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strTipoActividad = Nz(Me.TIPOACTIVIDAD, "")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Editar = strDescriptivo
    '       Editar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = editar(strTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Editar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    strCodActividadAlInicio = Nz(Me.OpenArgs, "")
    If strCodActividadAlInicio = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    Me.CodTipo = strCodActividadAlInicio
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbCodActividad", "ACTIVIDADNOPROGRAMADA", "CODIGO", strCodActividadAlInicio)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strTipoActividadAlInicio = flag
    If strTipoActividadAlInicio = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    Me.TIPOACTIVIDAD = strTipoActividadAlInicio
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
Public Function editar( _
                        strTipoActividad As String _
                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/2/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -sólo se puede cambiar el tipo de actividad
    '       -strCodActividadAlInicio <>""  no existente
    '       -strTipoActividad <>"" y no existente
    '       - No ha formado parte de ningún registro de actividad con Tipo de actividad strTipoActividadActual
    '   -llamada por:
    
    '   -Devuelve:
    '       Editar = Descriptivo
    '       Editar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    
    If strCodActividadAlInicio = "" Then
        strTextoError = "No se ha indicado el código de la actividad"
        Err.Raise 1000
    End If
    If strTipoActividad = "" Then
        strTextoError = "No se ha indicado el tipo de actividad"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbCodActividad", "ACTIVIDADNOPROGRAMADA", "CODIGO", strCodActividadAlInicio)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "" Then
        strTextoError = "No existe un código de actividad como ese"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbCodActividad", "ACTIVIDADNOPROGRAMADA", "ACTIVIDADNOPROGRAMADA", strTipoActividad)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "Ya existe un tipo de actividad como esa"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbActividades", "Actividad", "Actividad", strCodActividadAlInicio)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "No se puede cambiar ya que forma parte de al menos una actividad ya registrada"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbCodActividad.ACTIVIDADNOPROGRAMADA " & _
                "FROM TbCodActividad " & _
                "WHERE (((TbCodActividad.CODIGO)='" & strCodActividadAlInicio & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe un código de actividad como ese"
            Err.Raise 1000
        End If
        .Edit
            .Fields("ACTIVIDADNOPROGRAMADA") = strTipoActividad
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    editar = "Se ha editado el Tipo de Actividad correctamente "
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Editar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    editar = "#ERR" & "|" & strTextoError
End Function











