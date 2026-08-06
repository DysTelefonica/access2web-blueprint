Option Compare Database
Option Explicit
Dim strIDEvento As String, strIDActividadOriginal As String, strEventoFacturado As String



Private Sub ComandoAlta_Click()
    Dim strHorasEscritasAhora As String, strHorasYaHechasEseDia As String, strHorasTotales As String
    Dim strAliasTecnico As String, strUbicacion As String, strFechaAlta As String, strHorasLaborables As String, _
        strHORASEXTRAS As String, strActividad As String, strIDEventoGenerado As String, strDESCRIPCION As String, _
        m_IDActividad As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strAliasTecnico = Nz(Me.ComboTecnico, "")
    strUbicacion = Nz(Me.ComboUbicacion, "")
    strFechaAlta = Nz(Me.FechaAlta, "")
    strHorasLaborables = Nz(Me.HorasLaborables, "")
    strHORASEXTRAS = Nz(Me.HorasExtras, "")
    strActividad = Nz(Me.ComboActividad, "")
    strIDEventoGenerado = Nz(Me.IDEventoGenerado, "")
    strDESCRIPCION = Nz(Me.Descripcion, "")
    
    If Not IsNumeric(strHorasLaborables) Then
        strHorasLaborables = "0"
    End If
    If Not IsNumeric(strHORASEXTRAS) Then
        strHORASEXTRAS = "0"
    End If
    strHorasEscritasAhora = CStr(CDbl(strHorasLaborables) + CDbl(strHORASEXTRAS))
    If strAliasTecnico <> "" And strFechaAlta <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       DameHorasTecnico = cstr(dblHoras)
        '       DameHorasTecnico = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = DameHorasTecnico(strAliasTecnico, strFechaAlta)
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
    
    '-----------------------------------------------
    '   -Devuelve:
    '       ActividadAlta = strIDActividad
    '       ActividadAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    m_IDActividad = ActividadAlta( _
                        strAliasTecnico, strDESCRIPCION, strFechaAlta, strUbicacion, strHorasLaborables, strHORASEXTRAS, _
                        strActividad, strIDEvento, strIDEventoGenerado, , strEventoFacturado)
    If InStr(1, m_IDActividad, "|") <> 0 Then
        dato = Split(m_IDActividad, "|")
        strTextoError = "El método ActividadAlta ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormEventoGestion") Then
        Form_FormEventoGestion.ListaEventos_Click
    End If
    '---------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormActividadGestion") Then
        Forms("FormActividadGestion").Controls("ListaFiltrados").Requery
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta la actividad se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboActividad_Enter()
    On Error Resume Next
     Me.ComboActividad.Requery
End Sub
Private Sub IDEventoGenerado_Enter()
    On Error Resume Next
     Me.IDEventoGenerado.Requery
End Sub
Private Sub ComboUbicacion_Enter()
     On Error Resume Next
     Me.ComboUbicacion.Requery
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, cbo As ComboBox, strTextoError As String
    Dim strAliasTecnico As String, strUbicacion As String, strFechaAlta As String, strHorasLaborables As String, _
        strHORASEXTRAS As String, strActividad As String, strIDEventoGenerado As String, strDESCRIPCION As String, _
        strActividadFinalizaEvento As String
    On Error GoTo errores
    m_SQL = "SELECT TbCodActividad.CODIGO, TbCodActividad.ACTIVIDADNOPROGRAMADA FROM TbCodActividad ORDER BY TbCodActividad.CODIGO;"
    Set cbo = Me.ComboActividad
    RellenaCombo cbo, m_SQL
    m_SQL = "SELECT TbUbicacion.UBICACION FROM TbUbicacion ORDER BY TbUbicacion.UBICACION; "
    Set cbo = Me.ComboUbicacion
    RellenaCombo cbo, m_SQL
    If strIDActividadOriginal <> "" Then
        m_SQL = "SELECT TbActividades.* " & _
                    "FROM TbActividades " & _
                    "WHERE (((TbActividades.IDActividad)='" & strIDActividadOriginal & "'));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "No se han cargado todos los datos"
                Err.Raise 1000
            End If
            strIDEvento = Nz(.Fields("IDEvento"), "")
            strAliasTecnico = Nz(.Fields("ALIASTECNICO"), "")
            strUbicacion = Nz(.Fields("Ubicacion"), "")
            strFechaAlta = Nz(.Fields("FechaAlta"), "")
            strHorasLaborables = Nz(.Fields("HorasLaborables"), "")
            strHORASEXTRAS = Nz(.Fields("HorasExtras"), "")
            strActividad = Nz(.Fields("Actividad"), "")
            strIDEventoGenerado = Nz(.Fields("IDEVENTOGENERADO"), "")
            strDESCRIPCION = Nz(.Fields("DESCRIPCION"), "")
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Me.ComboTecnico.SetFocus
        If strAliasTecnico <> "" Then
            Me.ComboTecnico = strAliasTecnico
        End If
        If strUbicacion <> "" Then
            Me.ComboUbicacion = strUbicacion
        End If
        If IsDate(strFechaAlta) Then
            Me.FechaAlta = strFechaAlta
        End If
        If IsNumeric(strHorasLaborables) Then
            Me.HorasLaborables = CDbl(strHorasLaborables)
        End If
        If IsNumeric(strHORASEXTRAS) Then
            Me.HorasExtras = CDbl(strHORASEXTRAS)
        End If
        If strActividad <> "" Then
            Me.ComboActividad = strActividad
        End If
        If strIDEventoGenerado <> "" Then
            Me.IDEventoGenerado = strIDEventoGenerado
        End If
        If strDESCRIPCION <> "" Then
            Me.Descripcion = strDESCRIPCION
        End If
    End If
    If strIDEvento <> "" Then
        Me.lblTitulo1.Caption = "ALTA DE ACTIVIDAD DE: " & strIDEvento
    Else
        Me.lblTitulo1.Caption = "ALTA DE ACTIVIDAD NO DEPENDIENTE DE EVENTO"
        Me.FechaAlta = Format(Now(), "dd/mm/yyyy")
    End If
    Me.IDEVENTO = strIDEvento
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
Private Sub Form_Load()
    Dim strTextoError As String, strArgumento As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    'strArgumento=strIDEvento & "|" & strIDActividadOriginal & "|" & strEventoFacturado
    'strArgumento=""
    strArgumento = Nz(Me.OpenArgs(), "")
    If strArgumento <> "" Then
        If InStr(1, strArgumento, "|") = 0 Then
            strTextoError = "No se ha abierto el formulario con datos suficientes"
            Err.Raise 1000
        End If
        dato = Split(strArgumento, "|")
        If UBound(dato) <> 2 Then
            strTextoError = "No se ha abierto el formulario con datos suficientes"
            Err.Raise 1000
        End If
        strIDEvento = Nz(dato(0), "")
        strIDActividadOriginal = Nz(dato(1), "")
        strEventoFacturado = Nz(dato(2), "")
        If strIDEvento = "" Then
            strTextoError = "No se ha abierto el formulario con datos suficientes"
            Err.Raise 1000
        End If
        If strEventoFacturado <> "Sí" And strEventoFacturado <> "No" Then
            strTextoError = "No se ha abierto el formulario con datos suficientes"
            Err.Raise 1000
        End If
    End If
    
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    RellenarComboTecnicos Me.ComboTecnico, EnumSino.No, EnumSino.Sí, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

