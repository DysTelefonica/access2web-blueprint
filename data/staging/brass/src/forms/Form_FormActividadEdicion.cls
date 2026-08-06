Option Compare Database
Option Explicit
Dim strIDActividad As String, strEscritura As String, strURLAnexo As String
Dim strIDEventoAlInicio As String, strAliasTecnicoAlInicio As String, strUbicacionAlInicio As String, strFECHAALTAAlInicio As String, strHorasLaborablesAlInicio As String, _
        strHorasExtrasAlInicio As String, strActividadAlInicio As String, strIDEventoGeneradoAlInicio As String, strDescripcionAlInicio As String
Public strEstaFacturado As String, strFranqueado As String, strFormularioQueAbre As String

Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  22/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '-------------------------------------------------------------------
    Dim strAliasTecnico As String, strUbicacion As String, strFechaAlta As String, strHorasLaborables As String, _
        strHORASEXTRAS As String, strActividad As String, strIDEventoGenerado As String, strDESCRIPCION As String, _
        strCAUSAFIN As String
    On Error GoTo errores
    strAliasTecnico = Nz(Me.ComboTecnico.Column(0), "")
    strUbicacion = Nz(Me.ComboUbicacion, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strHorasLaborables = Nz(Me.HorasLaborables, "")
    strHORASEXTRAS = Nz(Me.HorasExtras, "")
    strActividad = Nz(Me.ComboActividad, "")
    strIDEventoGenerado = Nz(Me.ComboIDEventoGenerado, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    If IsDate(strFechaAlta) Then
        strFechaAlta = Format(strFechaAlta, "dd/mm/yyyy")
    End If
    
    If strAliasTecnico <> strAliasTecnicoAlInicio Or strUbicacion <> strUbicacionAlInicio Or strFechaAlta <> strFECHAALTAAlInicio Or _
        strHorasLaborables <> strHorasLaborablesAlInicio Or strHORASEXTRAS <> strHorasExtrasAlInicio Or strActividad <> strActividadAlInicio Or strIDEventoGenerado <> strIDEventoGeneradoAlInicio Or _
        strDESCRIPCION <> strDescripcionAlInicio Then
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
    '---------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '---------------------------------------------
    If HaHabidoCambios() And strEscritura = "Sí" Then
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
            If FormularioAbierto("FormEventoGestion") Then
                Forms("FormEventoGestion").Controls("ListaActividades").Requery
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
Private Sub ComandoAdjuntar_Click()
    Dim m_Actividad As Actividad
    Dim strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" Then
        strTextoError = "No se ha podido determinar el IDActividad"
        Err.Raise 1000
    End If
    '---------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '---------------------------------------------
    If HaHabidoCambios() Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       Registrar = descriptivo
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
        If FormularioAbierto("FormEventoGestion") Then
            Forms("FormEventoGestion").Controls("ListaActividades").Requery
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        '-------------------------------------------------------------------
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '--------------------------------------------
        If FormularioAbierto("FormAnexos") Then
            DoCmd.Close acForm, "FormAnexos", acSaveNo
        End If
    End If
    Set m_Actividad = Constructor.getActividad(strIDActividad, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    strFormularioQueAbre = Me.Name
    Set m_ObjEntidadParaAnexoActiva = m_Actividad
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    Dim blnDistintoEquipoMedida As Boolean
    Dim strAliasTecnico As String, strUbicacion As String, strFechaAlta As String, strHorasLaborables As String, _
        strHORASEXTRAS As String, strActividad As String, strIDEventoGenerado As String, strDESCRIPCION As String, strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    strAliasTecnico = Nz(Me.ComboTecnico.Column(0), "")
    strUbicacion = Nz(Me.ComboUbicacion, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strHorasLaborables = Nz(Me.HorasLaborables, "")
    strHORASEXTRAS = Nz(Me.HorasExtras, "")
    strActividad = Nz(Me.ComboActividad, "")
    strIDEventoGenerado = Nz(Me.ComboIDEventoGenerado, "#")
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    '-----------------------------------------------
    '   -Devuelve:
    '       Edicion = Descriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Edicion( _
                    strIDActividad, strAliasTecnico, strDESCRIPCION, strFechaAlta, strUbicacion, _
                    strHorasLaborables, strHORASEXTRAS, strActividad, strIDEventoAlInicio, strIDEventoGenerado, , strEstaFacturado)
    
    dato = Split(flag, "|")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Edicion ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If strIDEventoGenerado = "#" Then strIDEventoGenerado = ""
    If strDESCRIPCION = "#" Then strDESCRIPCION = ""
    strAliasTecnicoAlInicio = strAliasTecnico
    strUbicacionAlInicio = strUbicacion
    strFECHAALTAAlInicio = strFechaAlta
    strHorasLaborablesAlInicio = strHorasLaborables
    strHorasExtrasAlInicio = strHORASEXTRAS
    strActividadAlInicio = strActividad
    strIDEventoGeneradoAlInicio = strIDEventoGenerado
    strDescripcionAlInicio = strDESCRIPCION
    
    Registrar = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoRegistrar_Click()
    Dim strAliasTecnico As String, strFechaAlta As String, strHorasLaborables As String, strTextoError As String
    Dim strHorasYaHechasEseDia As String, strHorasTotales As String, strHorasEscritasAhora As String
    Dim strDESCRIPCION As String, strUbicacion As String, strHORASEXTRAS As String, strActividad As String, strIDEventoGenerado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDActividad = "" Then
        strTextoError = "El formulario no se ha abierto con los datos suficientes"
        Err.Raise 1000
    End If
    '---------------------------------------------
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '       HaHabidoCambios = True -->Error
    '---------------------------------------------
    If Not HaHabidoCambios() Then
        strTextoError = "No ha habido cambios que guarda"
        Err.Raise 1000
    End If
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    strUbicacion = Nz(Me.ComboUbicacion, "#")
    strHORASEXTRAS = Nz(Me.HorasExtras, "#")
    strActividad = Nz(Me.ComboActividad, "#")
    strIDEventoGenerado = Nz(Me.ComboIDEventoGenerado, "#")
    strAliasTecnico = Nz(Me.ComboTecnico.Column(0), "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strHorasLaborables = Nz(Me.HorasLaborables, "")
    If strHorasLaborables <> strHorasLaborablesAlInicio And strAliasTecnico <> "" And strFechaAlta <> "" Then
        If IsNumeric(strHorasLaborables) And IsNumeric(strHorasLaborablesAlInicio) Then
            If CDbl(strHorasLaborablesAlInicio) < CDbl(strHorasLaborables) Then
                '-------------------------------------------------------------------
                '   -Devuelve:
                '       DameHorasTecnico = cstr(dblHoras)
                '       DameHorasTecnico = "#ERR" & "|" & strTextoError
                '-------------------------------------------------------------------
                flag = DameHorasTecnico(strAliasTecnico, strFechaAlta, strIDActividad)
                If InStr(1, flag, "|") <> 0 Then
                    flag = ""
                End If
                strHorasYaHechasEseDia = flag
                If IsNumeric(strHorasYaHechasEseDia) Then
                    strHorasTotales = CStr(CDbl(strHorasYaHechasEseDia) + CDbl(strHorasLaborables))
                    If CDbl(strHorasTotales) > 7.5 Then
                        strTextoError = "Ese día " & strAliasTecnico & " ya llevaba " & strHorasYaHechasEseDia & _
                                        " horas y si se suman las " & strHorasEscritasAhora & " haría " & strHorasTotales & " horas"
                        Me.HorasExtras.SetFocus
                        Err.Raise 1000
                    End If
                End If
            End If
        End If
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
    '-------------------------------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Forms("FormEventoGestion").Controls("ListaActividades").Requery
    End If
    DoCmd.Hourglass False
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar la Actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoVerEvento_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If strIDEventoAlInicio = "" Then
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
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEventoAlInicio
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> strIDEventoAlInicio Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEventoAlInicio
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el Evento origen  se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboActividad_Enter()
     On Error Resume Next
     Me.ComboActividad.Requery
End Sub
Private Sub ComboIDEventoGenerado_Enter()
    On Error Resume Next
     Me.ComboIDEventoGenerado.Requery
End Sub
Private Sub ComboUbicacion_Enter()
    On Error Resume Next
    Me.ComboUbicacion.Requery
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    'strArgumento=strIDActividad  & "|" & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexo
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 3 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    strIDActividad = Nz(dato(0), "")
    strFranqueado = Nz(dato(1), "")
    strEstaFacturado = Nz(dato(2), "")
    
    
    If strIDActividad = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If strEstaFacturado = "Sí" Or strFranqueado = "Sí" Then
        strEscritura = "No"
    Else
        strEscritura = "Sí"
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       EstablecerDatos = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí, , strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = Me.lblTitulo1.Caption
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
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
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, ctl As Control, strTextoError As String
    On Error GoTo errores
    
    
   
    
    m_SQL = "SELECT TbActividades.* " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se han cargado todos los datos"
            Err.Raise 1000
        End If
        strIDEventoAlInicio = Nz(.Fields("IDEvento"), "")
        strAliasTecnicoAlInicio = Nz(.Fields("ALIASTECNICO"), "")
        strUbicacionAlInicio = Nz(.Fields("Ubicacion"), "")
        strFECHAALTAAlInicio = Nz(.Fields("FechaAlta"), "")
        strHorasLaborablesAlInicio = Nz(.Fields("HorasLaborables"), "")
        strHorasExtrasAlInicio = Nz(.Fields("HorasExtras"), "")
        strActividadAlInicio = Nz(.Fields("Actividad"), "")
        strIDEventoGeneradoAlInicio = Nz(.Fields("IDEVENTOGENERADO"), "")
        strDescripcionAlInicio = Nz(.Fields("DESCRIPCION"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Me.IDActividad = strIDActividad
    If IsDate(strFECHAALTAAlInicio) Then
        strFECHAALTAAlInicio = Format(strFECHAALTAAlInicio, "dd/mm/yyyy")
    End If
    If strIDEventoAlInicio <> "" Then
        Me.ComandoVerEvento.Caption = "Evento " & strIDEventoAlInicio
        Me.ComandoVerEvento.Visible = True
    Else
        Me.ComandoVerEvento.Visible = False
    End If
    If strAliasTecnicoAlInicio <> "" Then
        Me.ComboTecnico = strAliasTecnicoAlInicio
    End If
    If strUbicacionAlInicio <> "" Then
        Me.ComboUbicacion = strUbicacionAlInicio
    End If
    If IsDate(strFECHAALTAAlInicio) Then
        Me.FechaAlta = strFECHAALTAAlInicio
    End If
    If IsNumeric(strHorasLaborablesAlInicio) Then
        Me.HorasLaborables = CDbl(strHorasLaborablesAlInicio)
    End If
    If IsNumeric(strHorasExtrasAlInicio) Then
        Me.HorasExtras = CDbl(strHorasExtrasAlInicio)
    End If
    If strActividadAlInicio <> "" Then
        Me.ComboActividad = strActividadAlInicio
    End If
    If strIDEventoGeneradoAlInicio <> "" Then
        Me.ComboIDEventoGenerado = strIDEventoGeneradoAlInicio
    End If
    If strDescripcionAlInicio <> "" Then
        Me.Descripcion = strDescripcionAlInicio
    End If
    If strEscritura = "Sí" Then
        If strIDEventoAlInicio <> "" Then
            Me.lblTitulo1.Caption = "EDICIÓN DE ACTIVIDAD " & strIDActividad & " del Evento: " & strIDEventoAlInicio
        Else
            Me.lblTitulo1.Caption = "EDICIÓN DE ACTIVIDAD " & strIDActividad & " Sin Evento"
        End If
        Me.AllowEdits = True
'        For Each ctl In Me.Controls
'            If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox And ctl.Name <> "Tecnico" Then
'                ctl.BackColor = RGB(255, 255, 255)
'                ctl.Enabled = True
'                ctl.Locked = False
'            End If
'        Next
        Me.ComandoRegistrar.Enabled = True
    Else
        If strIDEventoAlInicio <> "" Then
            Me.lblTitulo1.Caption = "EDICIÓN DE ACTIVIDAD " & strIDActividad & " de evento ya franqueado"
        Else
            Me.lblTitulo1.Caption = "EDICIÓN DE ACTIVIDAD " & strIDActividad & " Sin Evento (Sólo lectura)"
        End If
        Me.AllowEdits = False
'        For Each ctl In Me.Controls
'            If ctl.ControlType = acComboBox Or ctl.ControlType = acTextBox Then
'                ctl.BackColor = lngColorCampoNoEditable
'                ctl.Enabled = True
'                ctl.Locked = True
'            End If
'        Next
        Me.ComandoRegistrar.Enabled = False
    End If
    
    m_SQL = "SELECT TbEventos.IDEvento, TbEventos.DESCRIPCION " & _
            "FROM TbEventos " & _
            "WHERE ((Not (TbEventos.IDEvento)='" & strIDEventoAlInicio & "')) " & _
            "ORDER BY TbEventos.IDEvento;"
    Me.ComboIDEventoGenerado.RowSource = m_SQL
    Me.ComboIDEventoGenerado.Requery
    If strURLAnexo = "" Then
        Me.ImagenAnexo.Visible = False
    Else
        Me.ImagenAnexo.Visible = True
    End If
    
    
    If strEstaFacturado = "Sí" Then
        Me.lblFacturado.Visible = True
    Else
        Me.lblFacturado.Visible = False
    End If
    If strFranqueado = "Sí" Then
        Me.lblFranqueado.Visible = True
    Else
        Me.lblFranqueado.Visible = False
    End If
    EstablecerDatos = "Datos de la actividad establecidos correctamente"
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
Private Sub ImagenAnexo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If InStr(1, strURLAnexo, ";") <> 0 Then
        If Me.ComandoAdjuntar.Enabled = True Then
            ComandoAdjuntar_Click
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AbrirEnLocal = strURLAnexoLocal
    '       AbrirEnLocal = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AbrirEnLocal(strURLAnexo, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function Edicion( _
                        strIDActividad As String, _
                        strAliasTecnico As String, _
                        strDESCRIPCION As String, _
                        strFechaAlta As String, _
                        strUbicacion As String, _
                        Optional strHorasLaborables As String, _
                        Optional strHORASEXTRAS As String, _
                        Optional strTipoActividad As String, _
                        Optional strIDEvento As String, _
                        Optional strIDEventoGenerado As String, _
                        Optional strAvisoMaximasHoras As String, _
                        Optional strEstaFacturado As String _
                        ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a tomar todos los datos que entran y los va a meter TbActividades--->
    '       -una actividad, puede depender de un evento o no
    '       -una actividad sólo puede ser en una jornada, no hay fecha final, si se acaba en el día se empieza otra en el día siguiente
    '       -1) Verificación de datos
    '           -strIDEvento Existente si está relleno
    '           -strIDEventoGenerado Existente si está relleno
    '           -strDescripcion obligatorio no puede ser #
    '           -strFechaAlta Fecha, si depende de un strIDEvento ha de ser igual o mayor que ésta
    '           -strHorasLaborables numérico <>0 si strHorasExtras no es un número o es =0
    '           -strTipoActividad obligatorio no puede ser #
    '   -Llamada desde
    '       -Form_FormActividadEdicion.Editar
    '   -Devuelve:
    '       Edicion = Descriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFECHAALTAEVENTO As String, strIDEventoAlInicio As String, strIDEventoGeneradoAlInicio As String
    Dim blnCambioEnCamposSensibles As Boolean
    Dim strFechaAlInicio As String, strHorasLabAlInicio As String, strHorasExtrasAlInicio As String, strTecnicoAlInicio As String
    Dim strHorasYaParaElDia As String, dblHorasYa As Double, dblHorasTotales As Double, strTextoError As String
    On Error GoTo errores
    
    If strAvisoMaximasHoras <> "Sí" And strAvisoMaximasHoras <> "No" Then
        strAvisoMaximasHoras = "No"
    End If
    If strIDEvento <> "" And strIDEventoGenerado <> "" And strIDEvento <> "#" And strIDEventoGenerado <> "#" Then
        If strIDEventoGenerado = strIDEvento Then
            strTextoError = "No puede ser igual el IDEvento que el IDEvento generado"
            Err.Raise 1000
        End If
    End If
    If strDESCRIPCION = "" Or strDESCRIPCION = "#" Then
        strTextoError = "No se puede dejar vacía la descripción de la actividad "
        Err.Raise 1000
    End If
    If Not IsDate(strFechaAlta) Then
        strTextoError = "La fecha de Edicion ha de ser una fecha válida "
        Err.Raise 1000
    End If
    If Not IsNumeric(strHorasLaborables) Then
        strHorasLaborables = "0"
    End If
    If Not IsNumeric(strHORASEXTRAS) Then
        strHORASEXTRAS = "0"
    End If
    If strAvisoMaximasHoras = "Sí" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       DameHorasTecnico = cstr(dblHoras)
        '       DameHorasTecnico = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = DameHorasTecnico(strAliasTecnico, strFechaAlta, strIDActividad)
        If InStr(1, flag, "|") <> 0 Then
            flag = ""
        End If
        strHorasYaParaElDia = flag
        If IsNumeric(strHorasYaParaElDia) Then
            dblHorasYa = CDbl(strHorasYaParaElDia)
            dblHorasTotales = dblHorasYa + CDbl(strHorasLaborables)
                
            If dblHorasTotales > 7.5 Then
                strTextoError = "Para el " & strFechaAlta & " " & strAliasTecnico & " ya llevaba " & CStr(dblHorasYa) & " y con " & strHorasLaborables & " horas sumaría " & CStr(dblHorasTotales) & " que supera el tope de 7,5"
                Err.Raise 1000
            End If
        End If
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaFacturado = strEstaFacturado
        '       EstaFacturado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFacturado("AC", strIDActividad)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFacturado ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEstaFacturado = flag
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        strTextoError = "El método Factura.EstaFacturado ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    If strTipoActividad = "" Or strTipoActividad = "#" Then
        strTextoError = "Se ha de indicar el tipo de actividad "
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe la actividad con código: " & strIDActividad
            Err.Raise 1000
        End If
        strIDEventoAlInicio = Nz(.Fields("IDEvento"), "")
        strIDEventoGeneradoAlInicio = Nz(.Fields("IDEVENTOGENERADO"), "")
        If strIDEvento <> "" And strIDEvento <> "#" Then
            If strIDEventoAlInicio <> strIDEvento Then
                '--------
                ' compruebo si el nuevo IDEvento existe
                '--------------
                '------------------------------------------
                '   -Devuelve:
                '       Dame = strValorObtenido
                '       Dame = "#ERR" & "|" & strTextoError
                '--------------------------------------------
                flag = Dame("TbEventos", "FECHAALTAEVENTO", "IDEvento", strIDEvento)
                If InStr(1, flag, "|") <> 0 Then
                    dato = Split(flag, "|")
                    strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
                    Err.Raise 1000
                End If
                strFECHAALTAEVENTO = flag
                If Not IsDate(strFECHAALTAEVENTO) Then
                    strTextoError = "El ID del evento ha de existir previamente "
                    Err.Raise 1000
                End If
                If CDate(strFechaAlta) < CDate(strFECHAALTAEVENTO) Then
                    strTextoError = "La fecha de Alta de la actividad no puede ser inferior de la de apertura del evento a la que pertenece "
                    Err.Raise 1000
                End If
            End If
        End If
        .Edit
            If strEstaFacturado = "Sí" Then
                strTecnicoAlInicio = Nz(.Fields("ALIASTECNICO"), "")
                strFechaAlInicio = Nz(.Fields("FechaAlta"), "")
                strHorasLabAlInicio = Nz(.Fields("HorasLaborables"), "")
                strHorasExtrasAlInicio = Nz(.Fields("HorasExtras"), "")
                If strFechaAlInicio <> strFechaAlta Or strHorasLabAlInicio <> strHorasLaborables Or strHorasExtrasAlInicio <> strHORASEXTRAS Or _
                    strTecnicoAlInicio <> strAliasTecnico Then
                    strTextoError = "Esta actividad ya ha sido facturada y no puede cambiar: " & vbCrLf & _
                                    vbTab & "- Técnico." & vbCrLf & _
                                    vbTab & "- Horas Laborables." & vbCrLf & _
                                    vbTab & "- Horas Extras." & vbCrLf & _
                                    vbTab & "- Fecha Alta."
                    
                    Err.Raise 1000
                End If
            End If
            If strIDEvento <> "" And strIDEvento <> "#" Then
                .Fields("IDEvento") = strIDEvento
            Else
                If strIDEvento = "#" Then
                    .Fields("IDEvento") = Null
                End If
            End If
            If strAliasTecnico <> "" Then
                .Fields("ALIASTECNICO") = strAliasTecnico
            End If
            
            If IsDate(strFechaAlta) Then
                .Fields("FechaAlta") = strFechaAlta
            End If
            .Fields("HorasLaborables") = CDbl(strHorasLaborables)
            .Fields("HorasExtras") = CDbl(strHORASEXTRAS)
            If strTipoActividad <> "" Then
                .Fields("Actividad") = strTipoActividad
            End If

            If strUbicacion <> "" Then
                .Fields("Ubicacion") = strUbicacion
            End If
            If strDESCRIPCION <> "" Then
                .Fields("DESCRIPCION") = strDESCRIPCION
            End If
            If strIDEventoGenerado <> "" And strIDEventoGenerado <> "#" Then
                .Fields("IDEVENTOGENERADO") = strIDEventoGenerado
            Else
                If strIDEventoGenerado = "#" Then
                    .Fields("IDEVENTOGENERADO") = Null
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Edicion = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.Edicion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Edicion = "#ERR" & "|" & strTextoError
End Function




