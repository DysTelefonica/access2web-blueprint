Option Compare Database
Option Explicit


Private WithEvents m_FormAnexo As Form_FormAdjuntarExcelSolicitante
Private WithEvents m_FormAnexoTraspaso As Form_FormAdjuntaTraspasoONS

Private m_SolFechas As SolicitudFechas
Private m_Error As String

Public Function Eliminar( _
                        p_TipoEnvioCorreo As EnumTipoEnvioCorreo, _
                        Optional ByRef p_Error As String) As String
    
    Dim m_SolicitudOp As SolicitudOperaciones
    
    On Error GoTo errores
    
    
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .EliminarFase p_TipoEnvioCorreo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    
    
    EstablecerDatos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub ComandoActualizarRecordatoriosExcelPreMarga_Click()

    Dim m_SolicitudfechasOp As SolicitudFechasOperaciones
    On Error GoTo errores
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If m_SolFechas Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    With m_SolFechas
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel1Calculado) Then
            .FechaPrevistaCorreoRecordatorioExcel1 = .FechaPrevistaCorreoRecordatorioExcel1Calculado
             Me.FechaPrevistaCorreoRecordatorioExcel1 = .FechaPrevistaCorreoRecordatorioExcel1
        Else
            .FechaPrevistaCorreoRecordatorioExcel1 = ""
            Me.FechaPrevistaCorreoRecordatorioExcel1 = Null
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel2Calculado) Then
            .FechaPrevistaCorreoRecordatorioExcel2 = .FechaPrevistaCorreoRecordatorioExcel2Calculado
             Me.FechaPrevistaCorreoRecordatorioExcel2 = .FechaPrevistaCorreoRecordatorioExcel2
        Else
            .FechaPrevistaCorreoRecordatorioExcel2 = ""
            Me.FechaPrevistaCorreoRecordatorioExcel2 = Null
        End If
        If IsDate(.FechaPrevistaAutocancelacionPreMargaCalculado) Then
            .FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMargaCalculado
             Me.FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMarga
        Else
            .FechaPrevistaAutocancelacionPreMarga = ""
            Me.FechaPrevistaAutocancelacionPreMarga = Null
        End If
    End With
    Set m_SolicitudfechasOp = New SolicitudFechasOperaciones
    With m_SolicitudfechasOp
        Set .SolicitudFechas = m_SolFechas
        .Registrar EnumSiNo.no, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarRecordatoriosExcelPreMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarCancelacion_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Cancelación?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.cancelacion, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    PintarIndicadores
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarCancelacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoEliminarEnvioTraspasoONS_Click()
    Dim m_SolOp As SolicitudOperaciones
    On Error GoTo errores
    m_Error = ""
    
    
    pregunta = MsgBox("¿Desea borrar la fase de Envío del Formulario de traspaso a la ONS?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fase")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    Set m_SolOp = New SolicitudOperaciones
    With m_SolOp
        Set .solicitud = m_ObjSolicitudActiva
        .EliminarEnvioTraspasoONS m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarEnvioTraspasoONS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoRegistraCancelacion_Click()
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Error = ""
    m_Fecha = Nz(InputBox("Fecha de Cancelación", "Registro Cancelación", Date), "")
    If m_Fecha = "" Then
        Exit Sub
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistroCancelado m_Fecha, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    PintarIndicadores
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistraCancelacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrarDesestimado_Click()
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Error = ""
    m_Fecha = Nz(InputBox("Fecha de Desestimación", "Registro Desectivación", Date), "")
    If m_Fecha = "" Then
        Exit Sub
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistroDesestimado m_Fecha, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    PintarIndicadores
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarDesestimado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub




Private Sub ComandoRegistrarEnvioTraspasoONS_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If FormularioAbierto("FormAdjuntaTraspasoONS") Then
        DoCmd.Close acForm, "FormAdjuntaTraspasoONS"
    End If
    
    DoCmd.OpenForm "FormAdjuntaTraspasoONS"
    If FormularioAbierto("FormAdjuntaTraspasoONS") Then
        Set m_FormAnexoTraspaso = Forms("FormAdjuntaTraspasoONS")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarEnvioTraspasoONS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Private Sub ComandoRegistrarExcelRecibido_Click()
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If FormularioAbierto("FormAdjuntarExcelSolicitante") Then
        DoCmd.Close acForm, "FormAdjuntarExcelSolicitante"
    End If
    
    DoCmd.OpenForm "FormAdjuntarExcelSolicitante"
    If FormularioAbierto("FormAdjuntarExcelSolicitante") Then
        Set m_FormAnexo = Forms("FormAdjuntarExcelSolicitante")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarExcelRecibido_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoEliminarDesestimado_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de desestimado?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.Desestimado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    EstablecerDatos
    PintarIndicadores
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarDesestimado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarExcelEnviado_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Envío del Excel1?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.EnvioExcel, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
   
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarExcelEnviado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoEliminarExcelRecibido_Click()
    Dim m_SolOp As SolicitudOperaciones
    On Error GoTo errores
    m_Error = ""
    
    
    pregunta = MsgBox("¿Desea borrar el archivo de excel que mandó el usuario?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    Set m_SolOp = New SolicitudOperaciones
    With m_SolOp
        Set .solicitud = m_ObjSolicitudActiva
        .EliminarExcelDevuelto m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarExcelEnviado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarRecordatorioExcel1_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Correo Recordatorio Excel1?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.RecordatorioExcel1, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarRecordatorioExcel1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarRecordatorioExcel2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Correo Recordatorio Excel2?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.RecordatorioExcel2, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarRecordatorioExcel2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

Private Sub ComandoRegistrarExcelEnviado_Click()
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
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarExcelEnviado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    m_EstadoAlInicio = m_ObjSolicitudActiva.Estado
    If Not IsDate(m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS) Then
        m_Error = "Primero hay que registrar la fecha de envío ONS"
        Err.Raise 1000
    End If
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
    m_EstadoAlFinal = m_ObjSolicitudActiva.Estado
    If m_EstadoAlFinal <> m_EstadoAlInicio Then
       If FormularioAbierto("FormSolicitudesGestion") Then
            Form_FormSolicitudesGestion.Filtrar
        End If
    End If
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
        m_Error = "Al ComandoRegistroHPS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoURLDocumentoTraspasoEnviado_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If m_ObjSolicitudActiva.URLAdjuntoEnvioONSCalculado = "" Then
        m_Error = "No hay ningún Adjunto de envío del formulario de traspaso a la ONS"
        Err.Raise 1000
    End If
    AbrirEnLocal m_ObjSolicitudActiva.URLAdjuntoEnvioONSCalculado, m_Error
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
        m_Error = "Al ComandoURLDocumentoTraspasoEnviado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoURLExcelRespuesta_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    If m_ObjSolicitudActiva.URLAdjuntoCalculado = "" Then
        m_Error = "No hay ningún Adjunto del Solicitante"
        Err.Raise 1000
    End If
    AbrirEnLocal m_ObjSolicitudActiva.URLAdjuntoCalculado, m_Error
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
        m_Error = "Al ComandoURLExcelRespuesta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoURLExcelSolicitud_Click()
    
    Dim m_URL As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    With m_ObjSolicitudActiva
        If .TipoCalculado = EnumTipo.Alta Or .TipoCalculado = EnumTipo.Renovacion Then
            m_URL = m_ObjEntorno.URLPlantillaAlta
        ElseIf .TipoCalculado = EnumTipo.Traspaso Then
            m_URL = m_ObjEntorno.URLPlantillaTraspaso
        Else
            m_Error = "Tipo deconocido"
            Err.Raise 1000
        End If
    End With
    AbrirEnLocal m_URL, m_Error
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
        m_Error = "Al ComandoURLExcelSolicitud_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

Private Sub m_FormAnexo_Adjuntado()
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
        m_Error = "Al m_FormAnexo_Adjuntado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    

    If m_ObjSolicitudActiva Is Nothing Then
    
       Exit Function
    End If
    Set m_SolFechas = m_ObjSolicitudActiva.SolicitudFechas
    If m_SolFechas Is Nothing Then
        Exit Function
    End If
    Me.ComandoRegistrarExcelEnviado.SetFocus
    
    If IsDate(m_SolFechas.FechaEnvioExcel) Then
        Me.FechaEnvioExcel = m_SolFechas.FechaEnvioExcel
        Me.ComandoEliminarExcelEnviado.Visible = True
    Else
        Me.FechaEnvioExcel = Null
        Me.ComandoEliminarExcelEnviado.Visible = False
    End If
    If IsDate(m_SolFechas.FechaCorreoRecordatorioExcel1) Then
        Me.FechaCorreoRecordatorioExcel1 = m_SolFechas.FechaCorreoRecordatorioExcel1
        Me.ComandoEliminarRecordatorioExcel1.Visible = True
    Else
        Me.FechaCorreoRecordatorioExcel1 = Null
        Me.ComandoEliminarRecordatorioExcel1.Visible = False
    End If
    If IsDate(m_SolFechas.FechaPrevistaCorreoRecordatorioExcel1) Then
        Me.FechaPrevistaCorreoRecordatorioExcel1 = m_SolFechas.FechaPrevistaCorreoRecordatorioExcel1
    Else
        Me.FechaPrevistaCorreoRecordatorioExcel1 = Null
    End If
    If IsDate(m_SolFechas.FechaCorreoRecordatorioExcel2) Then
        Me.FechaCorreoRecordatorioExcel2 = m_SolFechas.FechaCorreoRecordatorioExcel2
        Me.ComandoEliminarRecordatorioExcel2.Visible = True
    Else
        Me.FechaCorreoRecordatorioExcel2 = Null
        Me.ComandoEliminarRecordatorioExcel2.Visible = False
    End If
    If IsDate(m_SolFechas.FechaPrevistaCorreoRecordatorioExcel2) Then
        Me.FechaPrevistaCorreoRecordatorioExcel2 = m_SolFechas.FechaPrevistaCorreoRecordatorioExcel2
    Else
        Me.FechaPrevistaCorreoRecordatorioExcel2 = Null
    End If
    If IsDate(m_SolFechas.FechaPrevistaAutocancelacionPreMarga) Then
        Me.FechaPrevistaAutocancelacionPreMarga = m_SolFechas.FechaPrevistaAutocancelacionPreMarga
    Else
        Me.FechaPrevistaAutocancelacionPreMarga = Null
    End If
    If IsDate(m_SolFechas.FechaRecepcionExcel) Then
        Me.FechaRecepcionExcel = m_SolFechas.FechaRecepcionExcel
        Me.ComandoEliminarExcelRecibido.Visible = True
        Me.ComandoURLExcelRespuesta.Visible = True
    Else
        Me.FechaRecepcionExcel = Null
        Me.ComandoEliminarExcelRecibido.Visible = False
        Me.ComandoURLExcelRespuesta.Visible = False
    End If
    If IsDate(m_SolFechas.FechaEnvioONS) Then
        Me.FechaEnvioONS = m_SolFechas.FechaEnvioONS
        Me.ComandoURLDocumentoTraspasoEnviado.Visible = True
        Me.ComandoRegistroHPS.Enabled = True
        If IsDate(m_SolFechas.FechaRegistroEnHPS) Then
            Me.ComandoRegistrarEnvioTraspasoONS.Enabled = False
            Me.ComandoEliminarEnvioTraspasoONS.Visible = False
            
        Else
            Me.ComandoRegistrarEnvioTraspasoONS.Enabled = True
            Me.ComandoEliminarEnvioTraspasoONS.Visible = True
           
        End If
        
    Else
        Me.FechaEnvioONS = Null
        Me.ComandoURLDocumentoTraspasoEnviado.Visible = False
        Me.ComandoEliminarEnvioTraspasoONS.Visible = False
        Me.ComandoRegistroHPS.Enabled = False
    End If
    If IsDate(m_SolFechas.FechaRegistroEnHPS) Then
        Me.FechaRegistroEnHPS = m_SolFechas.FechaRegistroEnHPS
        
    Else
        Me.FechaRegistroEnHPS = Null
        
    End If
    If IsDate(m_SolFechas.FechaDesestimado) Then
        Me.FechaDesestimado = m_SolFechas.FechaDesestimado
        Me.ComandoEliminarDesestimado.Visible = True
    Else
        Me.FechaDesestimado = Null
        Me.ComandoEliminarDesestimado.Visible = False
    End If
    If IsDate(m_SolFechas.FechaCancelado) Then
        Me.FechaCancelado = m_SolFechas.FechaCancelado
        Me.ComandoEliminarCancelacion.Visible = True
    Else
        Me.FechaCancelado = Null
        Me.ComandoEliminarCancelacion.Visible = False
    End If
    
    If IsDate(Me.FechaPrevistaAutocancelacionPreMarga) Then
        Me.DiasParaCancelacion = DateDiff("d", CDate(Date), CDate(Me.FechaPrevistaAutocancelacionPreMarga))
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub m_FormAnexoTraspaso_AdjuntadoEnvio()
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
        m_Error = "Al m_FormAnexoTraspaso_AdjuntadoEnvio se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
End Sub
