Option Compare Database
Option Explicit

Private m_Form As Form

Private m_Error As String

Private Sub ComandoEnvioCancelacionMarga_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.CancelacionMarga
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioCancelacionMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioCancelacionPreMarga_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.CancelacionPreMarga
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioCancelacionPreMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioExcel_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
   
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.EnvioExcel
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioExcel_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioRecordatorioExcel1_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.RecordatorioExcel1
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioRecordatorioExcel1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioRecordatorioExcel2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.RecordatorioExcel2
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioRecordatorioExcel2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioRecordatorioMarga1_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.RecordatorioMarga1
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioRecordatorioMarga1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnvioRecordatorioMarga2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=EnumTipoEnvioCorreo.RecordatorioMarga2
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnvioRecordatorioMarga2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoRegistroEnvioONS_Click()
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    Dim m_Resultado As String
    On Error GoTo errores
    
    m_Error = ""
   
    If Not IsDate(m_ObjSolicitudActiva.SolicitudFechas.FechaTramitacionAltaMarga) Then
        m_Error = "Primero hay que registrar el alta en Marga"
        Err.Raise 1000
    End If
    m_Fecha = Nz(InputBox("Fecha de Registro ONS", "Registro ONS", Date), "")
    If m_Fecha = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_SolicitudOp = New SolicitudOperaciones
    Set m_SolicitudOp.solicitud = m_ObjSolicitudActiva
    With m_SolicitudOp
        
        .RegistroAltaONS m_Fecha, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarAltaMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistroHPS_Click()
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    Dim m_MotivoNoOKRegistrarUsuarioHPSEnHPS As String
    Dim m_Resultado As String
    Dim m_EstadoAlInicio As String
    Dim m_EstadoAlFinal As String
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsDate(m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS) Then
        m_Error = "Primero hay que registrar la fecha de envío ONS"
        Err.Raise 1000
    End If
    m_EstadoAlInicio = m_ObjSolicitudActiva.Estado
    Set m_SolicitudOp = New SolicitudOperaciones
    Set m_SolicitudOp.solicitud = m_ObjSolicitudActiva
    m_MotivoNoOKRegistrarUsuarioHPSEnHPS = m_SolicitudOp.MotivoNoOKRegistrarUsuarioHPSEnHPS(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_Error = m_MotivoNoOKRegistrarUsuarioHPSEnHPS
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    With m_SolicitudOp
        
        m_Resultado = .RegistrarUsuarioHPSEnHPS(m_Error)
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        If m_Resultado = "0" Then
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
             pregunta = MsgBox("Se ha generado un UsuarioHPS en el aplicativo de HPS. La tramitación de la solicitud ya ha terminado", _
                        vbExclamation, "Fin de tramitación")
        ElseIf m_Resultado = "1" Then
            pregunta = MsgBox("Se ha actualizado un UsuarioHPS ya existente en el aplicativo de HPS. La tramitación de la solicitud ya ha terminado", _
                        vbExclamation, "Fin de tramitación")
        End If
    End With
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_EstadoAlFinal = m_ObjSolicitudActiva.Estado
    If m_EstadoAlFinal <> m_EstadoAlInicio Then
       If FormularioAbierto("FormSolicitudesGestion") Then
            Form_FormSolicitudesGestion.Filtrar
        End If
    End If
    
    PintarIndicadores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistroHPS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerCancelacionMarga_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLAutocancelacionMARGA, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerCancelacionMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerCancelacionPreMarga_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLAutocancelacionPREMARGA, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerCancelacionPreMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerHTMLEnvioExcel_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLPlantillasHTMLExcelSolicitud, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerHTMLEnvioExcel_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerRecordatorioExcel1_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLRecordatorioExcel1, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRecordatorioExcel1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerRecordatorioExcel2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLRecordatorioExcel2, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRecordatorioExcel2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerRecordatorioMarga1_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLRecordatorioMarga1, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRecordatorioMarga1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerRecordatorioMarga2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    ObtenerHTMLEnLocal m_ObjEntorno.URLHTMLRecordatorioMarga2, , m_ObjSolicitudActiva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRecordatorioMarga2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_SolFechas As SolicitudFechas
    On Error GoTo errores
    

    If m_ObjSolicitudActiva Is Nothing Then
    
       Exit Function
    End If
    
    Set m_SolFechas = m_ObjSolicitudActiva.SolicitudFechas
    If m_SolFechas Is Nothing Then
        Exit Function
    End If
    
    With m_SolFechas
        If IsDate(.FechaEnvioExcel) Then
            Me.FechaEnvioExcel = .FechaEnvioExcel
        End If
        If IsDate(.FechaCorreoRecordatorioExcel1) Then
            Me.FechaCorreoRecordatorioExcel1 = .FechaCorreoRecordatorioExcel1
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel1) Then
            Me.FechaPrevistaCorreoRecordatorioExcel1 = .FechaPrevistaCorreoRecordatorioExcel1
        End If
        If IsDate(.FechaCorreoRecordatorioExcel2) Then
            Me.FechaCorreoRecordatorioExcel2 = .FechaCorreoRecordatorioExcel2
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel2) Then
            Me.FechaPrevistaCorreoRecordatorioExcel2 = .FechaPrevistaCorreoRecordatorioExcel2
        End If
        If IsDate(.FechaPrevistaAutocancelacionPreMarga) Then
            Me.FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMarga
        End If
        If IsDate(.FechaRecepcionExcel) Then
            Me.FechaRecepcionExcel = .FechaRecepcionExcel
        End If
        If IsDate(.FechaTramitacionAltaMarga) Then
            Me.FechaTramitacionAltaMarga = .FechaTramitacionAltaMarga
        End If
        If IsDate(.FechaCorreoRecordatorioRellenoMarga1) Then
            Me.FechaCorreoRecordatorioRellenoMarga1 = .FechaCorreoRecordatorioRellenoMarga1
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioRellenoMarga1) Then
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = .FechaPrevistaCorreoRecordatorioRellenoMarga1
        End If
        If IsDate(.FechaCorreoRecordatorioRellenoMarga2) Then
            Me.FechaCorreoRecordatorioRellenoMarga2 = .FechaCorreoRecordatorioRellenoMarga2
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioRellenoMarga2) Then
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2
        End If
        If IsDate(.FechaPrevistaAutocancelacionMarga) Then
            Me.FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMarga
        End If
        If IsDate(.FechaEnvioONS) Then
            Me.FechaEnvioONS = .FechaEnvioONS
        End If
        If IsDate(.FechaRegistroEnHPS) Then
            Me.FechaRegistroEnHPS = .FechaRegistroEnHPS
        End If
        If m_EnOficina = EnumSiNo.Sí Then
            Me.ComandoRegistroEnvioONS.Enabled = True
        Else
            Me.ComandoRegistroEnvioONS.Enabled = False
        End If
        If IsDate(.FechaDesestimado) Then
            Me.FechaDesestimado = .FechaDesestimado
        End If
        If IsDate(.FechaAutocancelacion) Then
            Me.FechaAutocancelacion = .FechaAutocancelacion
        End If
        
        
        If IsDate(Me.FechaPrevistaAutocancelacionMarga) Then
            Me.DiasParaCancelacion = DateDiff("d", CDate(Date), CDate(Me.FechaPrevistaAutocancelacionMarga))
        Else
            If IsDate(Me.FechaPrevistaAutocancelacionPreMarga) Then
                Me.DiasParaCancelacion = DateDiff("d", CDate(Date), CDate(Me.FechaPrevistaAutocancelacionPreMarga))
            End If
        End If
    End With
    
    If IsDate(Me.FechaRegistroEnHPS) Then
        Me.ComandoRegistroHPS.Caption = "Quitar registro en HPS (Fin tramitación solicitud)"
        Me.ComandoRegistroHPS.Enabled = True
        Me.ComandoEnvioExcel.Enabled = False
        Me.ComandoVerRecordatorioExcel1.Enabled = False
        Me.ComandoEnvioRecordatorioExcel1.Enabled = False
        Me.ComandoVerRecordatorioExcel2.Enabled = False
        Me.ComandoEnvioRecordatorioExcel2.Enabled = False
        Me.ComandoEnvioCancelacionPreMarga.Enabled = False
        Me.ComandoAdjuntarExcelRecibido.Enabled = False
        Me.ComandoEnvioRecordatorioMarga1.Enabled = False
        Me.ComandoVerRecordatorioMarga1.Enabled = False
        Me.ComandoEnvioRecordatorioMarga2.Enabled = False
        Me.ComandoVerRecordatorioMarga2.Enabled = False
        Me.ComandoEnvioCancelacionMarga.Enabled = False
        Me.ComandoVerCancelacionMarga.Enabled = False
    Else
        Me.ComandoRegistroHPS.Caption = "Registrar en HPS (Fin tramitación solicitud)"
        Me.ComandoRegistroHPS.Enabled = True
        If IsDate(Me.FechaRecepcionExcel) Then
            Me.FechaEnvioExcel.Enabled = False
            Me.ComandoVerRecordatorioExcel1.Enabled = False
            Me.ComandoEnvioRecordatorioExcel1.Enabled = False
            Me.ComandoVerRecordatorioExcel2.Enabled = False
            Me.ComandoEnvioRecordatorioExcel2.Enabled = False
            Me.ComandoEnvioCancelacionPreMarga.Enabled = False
            
            
        Else
            Me.FechaEnvioExcel.Enabled = True
            If IsDate(Me.FechaEnvioExcel) Then
                
            Else
                Me.ComandoVerRecordatorioExcel1.Enabled = False
                Me.ComandoEnvioRecordatorioExcel1.Enabled = False
                Me.ComandoVerRecordatorioExcel2.Enabled = False
                Me.ComandoEnvioRecordatorioExcel2.Enabled = False
                Me.ComandoEnvioCancelacionPreMarga.Enabled = False
            End If
            
            
        End If
    End If
    
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
