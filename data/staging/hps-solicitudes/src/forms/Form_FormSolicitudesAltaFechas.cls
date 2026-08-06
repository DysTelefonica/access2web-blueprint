Option Compare Database
Option Explicit


Private WithEvents m_FormAnexo As Form_FormAdjuntarExcelSolicitante
Private m_Form As Form
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
Private Sub ComandoActualizarRecordatoriosExcelMarga_Click()
    
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
        If IsDate(.FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado) Then
            .FechaPrevistaCorreoRecordatorioRellenoMarga1 = .FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado
             Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = .FechaPrevistaCorreoRecordatorioRellenoMarga1
        Else
            .FechaPrevistaCorreoRecordatorioRellenoMarga1 = ""
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = Null
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel2Calculado) Then
            .FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
             Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2
        Else
            .FechaPrevistaCorreoRecordatorioRellenoMarga2 = ""
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = Null
        End If
        If IsDate(.FechaPrevistaAutocancelacionMargaCalculado) Then
            .FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMargaCalculado
             Me.FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMarga
        Else
            .FechaPrevistaAutocancelacionMarga = ""
            Me.FechaPrevistaAutocancelacionMarga = Null
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
        m_Error = "Al ComandoActualizarRecordatoriosExcelMarga_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

Private Sub ComandoEliminarDPS_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Envío DPS?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.EnvioDPS, m_Error
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
        m_Error = "Al ComandoEliminarDPS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoEliminarONS_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de Registro ONS?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.EnvioONS, m_Error
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
        m_Error = "Al ComandoEliminarONS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistraCancelacion_Click()
    
    Dim m_Tipo As EnumTipoEnvioCorreo
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    
    If FormularioAbierto("FormEnvioCorreo") Then
        DoCmd.Close acForm, "FormEnvioCorreo"
    End If
    If m_ObjSolicitudActiva.SolicitudFechas.FechaTramitacionAltaMarga <> "" Then
        m_Tipo = EnumTipoEnvioCorreo.CancelacionMarga
    Else
        m_Tipo = EnumTipoEnvioCorreo.CancelacionPreMarga
    End If
    DoCmd.OpenForm "FormEnvioCorreo", OpenArgs:=m_Tipo
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
    
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

Private Sub ComandoRegistrarDPS_Click()
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Error = ""
    If Not IsDate(m_SolFechas.FechaTramitacionAltaMarga) Then
        m_Error = "Primero hay que registrar el alta de la ficha en Marga"
        Err.Raise 1000
    End If
    m_Fecha = Nz(InputBox("Fecha de Registro DPS", "Registro DPS", Date), "")
    If m_Fecha = "" Then
        Exit Sub
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistroAltaDPS m_Fecha, m_Error
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
        m_Error = "Al ComandoRegistrarDPS_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrarEnvioONS_Click()
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

Private Sub ComandoRegistrarAltaMarga_Click()
    
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Error = ""
    m_Fecha = Nz(InputBox("Fecha de Alta en Marga", "Alta en Marga", Date), "")
    If m_Fecha = "" Then
        Exit Sub
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistroAltaMarga m_Fecha, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PintarIndicadores
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



Private Sub ComandoEliminarAltaMarga_Click()
    
    Dim m_SolicitudOp As SolicitudOperaciones
    
    On Error GoTo errores
    
    m_Error = ""
    pregunta = MsgBox("¿Desea borrar la fecha de Registro en Marga?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .EliminarMarga m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
       
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Set m_ObjSolicitudActiva.SolicitudFechas = Nothing
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.AllowEdits = True
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
    
    
    pregunta = MsgBox("¿Desea borrar el archivo de excel que mandó el usuiario?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
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

Private Sub ComandoEliminarRecordatorioMarga1_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de correo Recordatorio Marga1?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.RecordatorioMarga1, m_Error
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
        m_Error = "Al ComandoEliminarRecordatorioMarga1_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarRecordatorioMarga2_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea borrar la fecha de correo Recordatorio Marga2?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Fecha")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Eliminar EnumTipoEnvioCorreo.RecordatorioMarga2, m_Error
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
        m_Error = "Al ComandoEliminarRecordatorioMarga2_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


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
    Set m_Form = Forms("FormEnvioCorreo")
    m_Form.Programado.Enabled = False
'    m_Form.FECHA.Enabled = False
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
    Set m_Form = Forms("FormEnvioCorreo")
    
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
    With m_SolFechas
        If IsDate(.FechaEnvioExcel) Then
            Me.FechaEnvioExcel = .FechaEnvioExcel
            Me.ComandoEliminarExcelEnviado.Visible = True
        Else
            Me.FechaEnvioExcel = Null
            Me.ComandoEliminarExcelEnviado.Visible = False
        End If
        If IsDate(.FechaCorreoRecordatorioExcel1) Then
            Me.FechaCorreoRecordatorioExcel1 = .FechaCorreoRecordatorioExcel1
            Me.ComandoEliminarRecordatorioExcel1.Visible = True
        Else
            Me.FechaCorreoRecordatorioExcel1 = Null
            Me.ComandoEliminarRecordatorioExcel1.Visible = False
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel1) Then
            Me.FechaPrevistaCorreoRecordatorioExcel1 = .FechaPrevistaCorreoRecordatorioExcel1
        Else
            Me.FechaPrevistaCorreoRecordatorioExcel1 = Null
        End If
        If IsDate(.FechaCorreoRecordatorioExcel2) Then
            Me.FechaCorreoRecordatorioExcel2 = .FechaCorreoRecordatorioExcel2
            Me.ComandoEliminarRecordatorioExcel2.Visible = True
        Else
            Me.FechaCorreoRecordatorioExcel2 = Null
            Me.ComandoEliminarRecordatorioExcel2.Visible = False
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioExcel2) Then
            Me.FechaPrevistaCorreoRecordatorioExcel2 = .FechaPrevistaCorreoRecordatorioExcel2
        Else
            Me.FechaPrevistaCorreoRecordatorioExcel2 = Null
        End If
        If IsDate(.FechaPrevistaAutocancelacionPreMarga) Then
            Me.FechaPrevistaAutocancelacionPreMarga = .FechaPrevistaAutocancelacionPreMarga
        Else
            Me.FechaPrevistaAutocancelacionPreMarga = Null
        End If
        If IsDate(.FechaRecepcionExcel) Then
            Me.FechaRecepcionExcel = .FechaRecepcionExcel
            Me.ComandoEliminarExcelRecibido.Visible = True
            Me.ComandoURLExcelRespuesta.Visible = True
        Else
            Me.FechaRecepcionExcel = Null
            Me.ComandoEliminarExcelRecibido.Visible = False
            Me.ComandoURLExcelRespuesta.Visible = False
        End If
        If IsDate(.FechaTramitacionAltaMarga) Then
            Me.FechaTramitacionAltaMarga = .FechaTramitacionAltaMarga
            Me.ComandoEliminarAltaMarga.Visible = True
        Else
            Me.FechaTramitacionAltaMarga = Null
            Me.ComandoEliminarAltaMarga.Visible = False
        End If
        If IsDate(.FechaCorreoRecordatorioRellenoMarga1) Then
            Me.FechaCorreoRecordatorioRellenoMarga1 = .FechaCorreoRecordatorioRellenoMarga1
            Me.ComandoEliminarRecordatorioMarga1.Visible = True
        Else
            Me.FechaCorreoRecordatorioRellenoMarga1 = Null
            Me.ComandoEliminarRecordatorioMarga1.Visible = False
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioRellenoMarga1) Then
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = .FechaPrevistaCorreoRecordatorioRellenoMarga1
        Else
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = Null
        End If
        If IsDate(.FechaCorreoRecordatorioRellenoMarga2) Then
            Me.FechaCorreoRecordatorioRellenoMarga2 = .FechaCorreoRecordatorioRellenoMarga2
            Me.ComandoEliminarRecordatorioMarga2.Visible = True
        Else
            Me.FechaCorreoRecordatorioRellenoMarga2 = Null
            Me.ComandoEliminarRecordatorioMarga2.Visible = False
        End If
        If IsDate(.FechaPrevistaCorreoRecordatorioRellenoMarga2) Then
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = .FechaPrevistaCorreoRecordatorioRellenoMarga2
        Else
            Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = Null
        End If
        If IsDate(.FechaPrevistaAutocancelacionMarga) Then
            Me.FechaPrevistaAutocancelacionMarga = .FechaPrevistaAutocancelacionMarga
        Else
            Me.FechaPrevistaAutocancelacionMarga = Null
        End If
        If IsDate(.FechaEnvioDPS) Then
            Me.FechaEnvioDPS = .FechaEnvioDPS
            Me.ComandoEliminarDPS.Visible = True
        Else
            Me.FechaEnvioDPS = Null
            Me.ComandoEliminarDPS.Visible = False
        End If
        If IsDate(.FechaEnvioONS) Then
            Me.FechaEnvioONS = .FechaEnvioONS
            If IsDate(Me.FechaRegistroEnHPS) Then
                Me.ComandoRegistrarEnvioONS.Enabled = False
                Me.ComandoRegistroHPS.Enabled = False
                Me.ComandoEliminarONS.Visible = False
            Else
                Me.ComandoRegistrarEnvioONS.Enabled = True
                Me.ComandoRegistroHPS.Enabled = True
                Me.ComandoEliminarONS.Visible = True
            End If
            
        Else
            Me.FechaEnvioONS = Null
            If Me.FechaEnvioDPS <> "" Then
                Me.ComandoRegistrarEnvioONS.Enabled = True
                
            Else
                Me.ComandoRegistrarEnvioONS.Enabled = False
               
                
            End If
            Me.ComandoEliminarONS.Visible = False
             Me.ComandoRegistroHPS.Enabled = False
            
        End If
        If IsDate(.FechaRegistroEnHPS) Then
            Me.FechaRegistroEnHPS = .FechaRegistroEnHPS
            
        Else
            Me.FechaRegistroEnHPS = Null
        End If
        
        If IsDate(.FechaDesestimado) Then
            Me.FechaDesestimado = .FechaDesestimado
            Me.ComandoEliminarDesestimado.Visible = True
        Else
            Me.FechaDesestimado = Null
            Me.ComandoEliminarDesestimado.Visible = False
        End If
        If IsDate(.FechaCancelado) Then
            Me.FechaCancelado = .FechaCancelado
            Me.ComandoEliminarCancelacion.Visible = True
        Else
            Me.FechaCancelado = Null
            Me.ComandoEliminarCancelacion.Visible = False
        End If
        
        
       
        
    End With
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function