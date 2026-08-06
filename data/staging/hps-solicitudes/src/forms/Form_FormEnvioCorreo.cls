 Option Compare Database
Option Explicit


Public Event RegistroFechaEnvioExcel()
Public Event RegistroRecordatorioExcel1()
Public Event RegistroRecordatorioExcel2()
Public Event RegistroRecordatorioMarga1()
Public Event RegistroRecordatorioMarga2()
Private m_TipoEnvioCorreo As EnumTipoEnvioCorreo
Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEnviar_Click()
    Dim msg As String
    
    On Error GoTo errores
    m_Error = ""
    
    ' 1. Mensaje de confirmación según modo
    If Me.chkActualizarFecha.value = True Then
        msg = "¿Desea REGISTRAR el hito con fecha " & Me.FECHA & " y ENVIAR el correo?"
    Else
        msg = "ATENCIÓN: Se va a proceder al REENVÍO del correo." & vbNewLine & vbNewLine & _
              "NO se modificará la fecha del hito en la base de datos ni se alterarán los plazos." & vbNewLine & _
              "¿Desea continuar?"
    End If
    
    ' 2. Confirmación
    pregunta = MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, "Confirmación de Operación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    ' 2.1 Validación de fecha si se va a actualizar
    If Me.chkActualizarFecha.value = True Then
        If Not ValidarFechaEdicion(m_Error) Then
            Err.Raise 1000
        End If
    End If
    
    ' 3. Bloqueo visual
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Procesando envío..."
    
    ' 4. Ejecutar
    Registrar EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 5. Refresco de la solicitud en memoria
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 6. Actualización UI formularios padre
    On Error Resume Next
    If FormularioAbierto("FormSolicitudAltaTraspaso") Then
        If Forms("FormSolicitudAltaTraspaso").formDetalle.SourceObject = "FormSolicitudesAltaFechas" Then
            Form_FormSolicitudesAltaFechas.EstablecerDatos
        Else
            Form_FormSolicitudesTraspasoFechas.EstablecerDatos
        End If
    ElseIf FormularioAbierto("FormSolicitudRenovacion") Then
        Form_FormSolicitudRenovacionFechas.EstablecerDatos
    End If
    On Error GoTo errores
    PintarIndicadores
    
    ' 7. Mensaje final y cierre
    Dim msgFinal As String
    If Me.chkActualizarFecha.value = True Then
        msgFinal = "Operación completada: Fecha registrada y correo enviado."
    Else
        msgFinal = "Operación completada: Correo reenviado correctamente."
    End If
    
    MsgBox msgFinal, vbInformation, "Éxito"
    
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    
    If Err.Number <> 1000 Then
        m_Error = "Error en ComandoEnviar_Click: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    MsgBox m_Error, vbCritical, "Error"
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim m_titulo As String
    Dim m_Asunto As String
    Dim m_Destinatarios As String
    Dim m_DestinatariosEnCopia As String
    Dim fechaYaRegistrada As String
    
    On Error GoTo errores
    
    If m_ObjSolicitudActiva Is Nothing Then
        p_Error = "No se ha podido saber el solicitante"
        Err.Raise 1000
    End If
    
    ' 1. Cargar textos
    With m_ObjSolicitudActiva
        m_titulo = .TituloParaFormulario(p_EnumTipoEnvioCorreo:=m_TipoEnvioCorreo, p_Error:=p_Error)
        m_Asunto = .AsuntoCalculado(p_EnumTipoEnvioCorreo:=m_TipoEnvioCorreo, p_Error:=p_Error)
        m_Destinatarios = .DestinatariosCalculados(p_Error:=p_Error)
        m_DestinatariosEnCopia = .DestinatariosConCopiaCalculados(p_EnumTipoEnvioCorreo:=m_TipoEnvioCorreo, p_Error:=p_Error)
    End With
    
    ' 2. Configurar controles básicos
    Me.lblTitulo.Caption = m_titulo
    Me.Asunto = m_Asunto
    Me.DESTINATARIOS = m_Destinatarios
    Me.DestinatariosConCopia = m_DestinatariosEnCopia
    Me.Programado = "No"
    
    ' 3. ¿Ya se realizó este hito?
    fechaYaRegistrada = ""
    If Not m_ObjSolicitudActiva.SolicitudFechas Is Nothing Then
        With m_ObjSolicitudActiva.SolicitudFechas
            Select Case m_TipoEnvioCorreo
                Case EnumTipoEnvioCorreo.EnvioExcel:         fechaYaRegistrada = .FechaEnvioExcel
                Case EnumTipoEnvioCorreo.RecordatorioExcel1:  fechaYaRegistrada = .FechaCorreoRecordatorioExcel1
                Case EnumTipoEnvioCorreo.RecordatorioExcel2:  fechaYaRegistrada = .FechaCorreoRecordatorioExcel2
                Case EnumTipoEnvioCorreo.CancelacionPreMarga: fechaYaRegistrada = .FechaCancelado
                Case EnumTipoEnvioCorreo.RecordatorioMarga1:  fechaYaRegistrada = .FechaCorreoRecordatorioRellenoMarga1
                Case EnumTipoEnvioCorreo.RecordatorioMarga2:  fechaYaRegistrada = .FechaCorreoRecordatorioRellenoMarga2
                Case EnumTipoEnvioCorreo.CancelacionMarga:    fechaYaRegistrada = .FechaCancelado
            End Select
        End With
    End If
    
    ' 4. Configurar el Checkbox y Botón según el estado histórico
    If IsDate(fechaYaRegistrada) Then
        ' Ya existe: modo reenvío por defecto
        Me.chkActualizarFecha.value = False
        Me.chkActualizarFecha.Enabled = True
        Me.ComandoEnviar.Caption = "Reenviar Correo"
        Me.FECHA = fechaYaRegistrada     ' Mostrar la fecha registrada para poder corregirla
    Else
        ' No existe: modo registro obligatorio
        Me.chkActualizarFecha.value = True
        Me.chkActualizarFecha.Enabled = False
        Me.ComandoEnviar.Caption = "Registrar y Enviar"
        Me.FECHA = Date
    End If
    
    ' 5. Habilitar edición de la fecha siempre (se validará antes de actualizar)
    Me.FECHA.Enabled = True
    Me.FECHA.Locked = False
    HabilitarEdicionFecha
    Me.ComandoEnviar.Enabled = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en EstablecerDatos: " & Err.Description
    End If
End Function


Public Function Registrar( _
                            Optional p_EnviandoCorreo As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String) As String
    
    Dim m_SolicitudOp As SolicitudOperaciones
    Dim m_IDCorreo As String
    Dim blnSoloEnvio As Boolean
    
    On Error GoTo errores
    
    ' Leemos el checkbox: Si está marcado, actualizamos fecha. Si NO, es solo envío.
    ' Nota: p_SoloEnvioCorreo es True cuando chkActualizarFecha es False
    If Me.chkActualizarFecha.value = True Then
        blnSoloEnvio = False
    Else
        blnSoloEnvio = True
    End If
    
    ' Si es "Grabar Solo Fecha" (botón secundario), forzamos actualización
    If p_EnviandoCorreo = EnumSiNo.no Then
        blnSoloEnvio = False
    End If
    
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        
        ' Pasamos el nuevo parámetro a todas las llamadas
        Select Case m_TipoEnvioCorreo
            Case EnumTipoEnvioCorreo.EnvioExcel
                m_IDCorreo = .RegistroEnvioExcel(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                                Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
                                
            Case EnumTipoEnvioCorreo.RecordatorioExcel1
                m_IDCorreo = .RegistroEnvioRecordatorioExcel1(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                            Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            
            Case EnumTipoEnvioCorreo.RecordatorioExcel2
                m_IDCorreo = .RegistroEnvioRecordatorioExcel2(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                            Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            
            Case EnumTipoEnvioCorreo.CancelacionPreMarga
                m_IDCorreo = .RegistroEnvioAutocancelaconPREMARGA(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                            Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            
            Case EnumTipoEnvioCorreo.RecordatorioMarga1
                m_IDCorreo = .RegistroEnvioRecordatorioMarga1(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                            Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            
            Case EnumTipoEnvioCorreo.RecordatorioMarga2
                m_IDCorreo = .RegistroEnvioRecordatorioMarga2(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                        Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            
            Case EnumTipoEnvioCorreo.CancelacionMarga
                m_IDCorreo = .RegistroEnvioAutocancelaconMARGA(p_EnviandoCorreo, Nz(Me.FECHA, ""), Nz(Me.DESTINATARIOS, ""), _
                            Nz(Me.DestinatariosConCopia, ""), Nz(Me.Asunto, ""), blnSoloEnvio, p_Error)
            Case Else
                p_Error = "Acción no codificada"
                Err.Raise 1000
        End Select
        
        If p_Error <> "" Then Err.Raise 1000
    End With
    
    ' Refrescos de UI
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    PintarIndicadores
    
    ' Actualizar formularios padre si están abiertos
    On Error Resume Next
    If m_ObjSolicitudActiva.TIPO = "Alta" Then
        Form_FormSolicitudesAltaFechas.EstablecerDatos
    ElseIf m_ObjSolicitudActiva.TIPO = "Traspaso" Then
        Form_FormSolicitudesTraspasoFechas.EstablecerDatos
    Else
        Form_FormSolicitudesAltaRenovacionFechas.EstablecerDatos
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en Form_FormEnvioCorreo.Registrar: " & Err.Description
    End If
End Function

Private Sub ComandoGrabarFecha_Click()
    On Error GoTo errores
    
    m_Error = ""
    pregunta = MsgBox("Va a grabar sólo la fecha, NO SE VA A MANDAR CORREO ELECTRÓNICO" & vbNewLine & _
                      "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Grabar Fecha")
    
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    ' Validación de fecha (no futura)
    If Not ValidarFechaEdicion(m_Error) Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Registrar EnumSiNo.no, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabarFecha_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    If Not IsNumeric(Nz(Me.OpenArgs(), "")) Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    m_TipoEnvioCorreo = CLng(Me.OpenArgs())
    
    Me.AllowEdits = False
    Me.Caption = m_TituloFormulario
       
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    HabilitarEdicionFecha
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Current()
    On Error Resume Next
    HabilitarEdicionFecha
End Sub
Private Function ValidarFechaEdicion(ByRef p_Error As String) As Boolean
    Dim d As Date
    p_Error = ""
    
    If Not IsDate(Nz(Me.FECHA, "")) Then
        p_Error = "Debe indicar una fecha válida."
        ValidarFechaEdicion = False
        Exit Function
    End If
    
    d = CDate(Me.FECHA)
    If d > Date Then
        p_Error = "No se permite una fecha posterior a la fecha de edición (hoy)."
        ValidarFechaEdicion = False
        Exit Function
    End If
    
    ValidarFechaEdicion = True
End Function
Private Sub HabilitarEdicionFecha()
    On Error Resume Next
    
    ' Asegurar edición a nivel de formulario
    Me.AllowEdits = True
    
    ' Forzar habilitado del control FECHA
    With Me.FECHA
        .Enabled = True
        .Locked = False
        .TabStop = True
    End With
    
    ' Si el control está enlazado a una expresión no editable, desenlazar
    ' (Opcional: solo si detectamos signos típicos de expresión)
    If Nz(Me.FECHA.ControlSource, "") <> "" Then
        Dim src As String
        src = Trim(Me.FECHA.ControlSource)
        If Left$(src, 1) = "=" Or InStr(src, "(") > 0 Or InStr(src, "+") > 0 Or InStr(src, "&") > 0 Then
            Me.FECHA.ControlSource = ""  ' Unbound para permitir edición
        End If
    End If
End Sub