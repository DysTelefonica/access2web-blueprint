
Option Compare Database
Option Explicit

'===========================================================
' Form_FormEventoSLA
' Spec-003: Formulario de edición completa de campos SLA
'===========================================================

Private m_IDEvento As String
Private m_Evento As Evento

'===========================================================
' LimpiarCamposSLAUI
' Limpia todos los controles SLA para evitar arrastre visual
'===========================================================
Private Sub LimpiarCamposSLAUI()
    On Error Resume Next

    Me.FechaRecepcionNotificacion = Null
    Me.FechaInicioContactoCliente = Null

    Me.cmbIncidenciaAveriaOReparacion = Null
    Me.FechaInicioTiempoAdquisicion = Null
    Me.FechaFinTiempoAdquisicion = Null
    Me.cmbTipoReparacion = Null
    Me.chkUrgente = Null

    Me.cmbEventoConServicioAfectado = Null
    Me.FechaRestablecimientoServicio = Null

    Me.chkTipoRepInsitu = Null
    Me.chkTipoRepNoSMT = Null
    Me.chkTipoRepValvulas = Null
End Sub

'===========================================================
' Form_Load
' Inicializa el formulario y carga datos del evento
'===========================================================
Private Sub Form_Load()
    Dim strIDEvento As String
    Dim strTextoError As String
    Dim strEstaFranqueado As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    strIDEvento = Nz(Me.OpenArgs(), "")
    
    If strIDEvento <> "" Then
        ' --- FRANQUEADO GUARD (UX) ---
        strEstaFranqueado = EstaFranqueado("EV", strIDEvento)
        If strEstaFranqueado = "Sí" Then
            Me.AllowEdits = False
            Me.cmdGuardar.Enabled = False
        End If
        ' --- END GUARD ---
        
        m_IDEvento = strIDEvento
        If strEstaFranqueado = "Sí" Then
            Me.lblTitulo1.Caption = "Datos SLA - Evento " & m_IDEvento & " (Solo lectura)"
        Else
            Me.lblTitulo1.Caption = "Datos SLA - Evento " & m_IDEvento
        End If
        Call CargarDatosEvento(m_IDEvento)
    Else
        Me.lblTitulo1.Caption = "Datos SLA - Sin evento"
    End If
    
    Me.Caption = Me.lblTitulo1.Caption
    
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al cargar el formulario: " & Err.Number & vbNewLine & Err.Description
    End If
    DoCmd.Hourglass False
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'===========================================================
' CargarDatosEvento
' Carga los datos SLA del evento desde la base de datos
'===========================================================
Private Sub CargarDatosEvento(p_IDEvento As String)
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim strTextoError As String
    
    On Error GoTo errores
    
    If p_IDEvento = "" Then
        Exit Sub
    End If

    ' Blindaje: evitar valores residuales al cambiar de evento
    Call LimpiarCamposSLAUI
    
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE TbEventos.IDEvento='" & p_IDEvento & "';"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If Not .EOF Then
            ' TRES - Tiempo de Respuesta
            If IsDate(.Fields("FechaRecepcionNotificacion")) Then
                Me.FechaRecepcionNotificacion = Format(CDate(.Fields("FechaRecepcionNotificacion")), "dd/mm/yyyy hh:nn")
            End If
            If IsDate(.Fields("FechaInicioContactoCliente")) Then
                Me.FechaInicioContactoCliente = Format(CDate(.Fields("FechaInicioContactoCliente")), "dd/mm/yyyy hh:nn")
            End If
            
            ' TRCM - Tiempo de Reparación
            If Not IsNull(.Fields("IncidenciaAveriaOReparacion")) Then
                If .Fields("IncidenciaAveriaOReparacion") = True Then
                    Me.cmbIncidenciaAveriaOReparacion = "Sí"
                Else
                    ' Si está a False por defecto y no hay datos TRCM asociados,
                    ' mostrar vacío para evitar "No" automático
                    If IsDate(.Fields("FechaInicioTiempoAdquisicion")) Or _
                       IsDate(.Fields("FechaFinTiempoAdquisicion")) Or _
                       Nz(.Fields("TipoReparacion"), "") <> "" Then
                        Me.cmbIncidenciaAveriaOReparacion = "No"
                    Else
                        Me.cmbIncidenciaAveriaOReparacion = Null
                    End If
                End If
            End If
            
            If IsDate(.Fields("FechaInicioTiempoAdquisicion")) Then
                Me.FechaInicioTiempoAdquisicion = .Fields("FechaInicioTiempoAdquisicion")
            End If
            
            If IsDate(.Fields("FechaFinTiempoAdquisicion")) Then
                Me.FechaFinTiempoAdquisicion = .Fields("FechaFinTiempoAdquisicion")
            End If
            
            If Not IsNull(.Fields("TipoReparacion")) And .Fields("TipoReparacion") <> "" Then
                Me.cmbTipoReparacion = .Fields("TipoReparacion")
            End If
            
            If Not IsNull(.Fields("Urgente")) Then
                Me.chkUrgente = (.Fields("Urgente") = True)
            End If
            
            ' TRSS - Restablecimiento Servicio
            If Not IsNull(.Fields("EventoConServicioAfectado")) Then
                If .Fields("EventoConServicioAfectado") = True Then
                    Me.cmbEventoConServicioAfectado = "Sí"
                Else
                    ' Si está a False por defecto y no hay fecha de restablecimiento,
                    ' mostrar vacío para evitar "No" automático
                    If IsDate(.Fields("FechaRestablecimientoServicio")) Then
                        Me.cmbEventoConServicioAfectado = "No"
                    Else
                        Me.cmbEventoConServicioAfectado = Null
                    End If
                End If
            End If
            
            If IsDate(.Fields("FechaRestablecimientoServicio")) Then
                Me.FechaRestablecimientoServicio = .Fields("FechaRestablecimientoServicio")
            End If
            
            ' SLA-4 - Reparación In-situ
            If Not IsNull(.Fields("TipoRepInsitu")) Then
                Me.chkTipoRepInsitu = (.Fields("TipoRepInsitu") = True)
            End If
            If Not IsNull(.Fields("TipoRepNoSMT")) Then
                Me.chkTipoRepNoSMT = (.Fields("TipoRepNoSMT") = True)
            End If
            If Not IsNull(.Fields("TipoRepValvulas")) Then
                Me.chkTipoRepValvulas = (.Fields("TipoRepValvulas") = True)
            End If
        End If
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al cargar datos del evento: " & Err.Number & vbNewLine & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'===========================================================
' cmdGuardar_Click
' Guarda los campos SLA del evento
'===========================================================
Private Sub cmdGuardar_Click()
    Dim strTextoError As String
    Dim m_ResultadoValidacion As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    ' Validar usando SLAValidator (edición normal: p_EnFranqueo=False)
    m_ResultadoValidacion = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:=Nz(Me.FechaRecepcionNotificacion, ""), _
        p_FechaInicioContacto:=Nz(Me.FechaInicioContactoCliente, ""), _
        p_IncidenciaAveria:=Nz(Me.cmbIncidenciaAveriaOReparacion, ""), _
        p_FechaInicioAdquisicion:=Nz(Me.FechaInicioTiempoAdquisicion, ""), _
        p_FechaFinAdquisicion:=Nz(Me.FechaFinTiempoAdquisicion, ""), _
        p_TipoReparacion:=Nz(Me.cmbTipoReparacion, ""), _
        p_Urgente:=IIf(Me.chkUrgente = True, True, IIf(Me.chkUrgente = False, False, Null)), _
        p_EventoConServicioAfectado:=Nz(Me.cmbEventoConServicioAfectado, ""), _
        p_FechaRestablecimiento:=Nz(Me.FechaRestablecimientoServicio, ""), _
        p_Mensaje:=strTextoError, _
        p_Error:=strTextoError)
    If m_ResultadoValidacion <> "" Then
        Err.Raise 1000
    End If
    
    ' Guardar en base de datos
    Call GuardarSLA
    
    MsgBox "Datos SLA guardados correctamente.", vbInformation, "Éxito"
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al guardar los datos SLA: " & Err.Number & vbNewLine & Err.Description
    End If
    DoCmd.Hourglass False
    VBA.DoEvents
    If strTextoError <> "" Then
        MsgBox strTextoError, vbCritical, "Error"
    End If
End Sub

'===========================================================
' cmdCerrar_Click
' Cierra el formulario sin guardar
'===========================================================
Private Sub cmdCerrar_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

'===========================================================
' GuardarSLA
' Guarda los campos SLA en la tabla TbEventos
'===========================================================
Private Sub GuardarSLA()
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim strTextoError As String
    Dim m_FechaRecepcion As String
    Dim m_FechaInicioContacto As String
    Dim m_IncidenciaAveria As Variant
    Dim m_FechaInicioAdquisicion As String
    Dim m_FechaFinAdquisicion As String
    Dim m_TipoReparacion As String
    Dim m_Urgente As Variant
    Dim m_EventoConServicio As Variant
    Dim m_FechaRestablecimiento As String
    Dim m_TipoRepInsitu As Variant
    Dim m_TipoRepNoSMT As Variant
    Dim m_TipoRepValvulas As Variant
    
    On Error GoTo errores
    
    ' --- FRANQUEADO GUARD (SERVICE LAYER) ---
    If m_IDEvento <> "" Then
        Dim strFlag As String
        strFlag = EstaFranqueado("EV", m_IDEvento)
        If InStr(1, strFlag, "|") = 0 Then  ' not an error
            If strFlag = "Sí" Then
                strTextoError = "No se puede guardar SLA de un evento franqueado."
                Err.Raise 1000
            End If
        End If
    End If
    ' --- END GUARD ---
    
    If m_IDEvento = "" Then
        strTextoError = "No hay evento seleccionado"
        Err.Raise 1000
    End If
    
    ' Recoger valores de los controles
    m_FechaRecepcion = Nz(Me.FechaRecepcionNotificacion, "")
    m_FechaInicioContacto = Nz(Me.FechaInicioContactoCliente, "")
    
    If Me.cmbIncidenciaAveriaOReparacion = "Sí" Then
        m_IncidenciaAveria = True
    ElseIf Me.cmbIncidenciaAveriaOReparacion = "No" Then
        m_IncidenciaAveria = False
    Else
        m_IncidenciaAveria = Null
    End If
    
    m_FechaInicioAdquisicion = Nz(Me.FechaInicioTiempoAdquisicion, "")
    m_FechaFinAdquisicion = Nz(Me.FechaFinTiempoAdquisicion, "")
    m_TipoReparacion = Nz(Me.cmbTipoReparacion, "")
    m_Urgente = IIf(Me.chkUrgente = True, True, IIf(Me.chkUrgente = False, False, Null))
    
    If Me.cmbEventoConServicioAfectado = "Sí" Then
        m_EventoConServicio = True
    ElseIf Me.cmbEventoConServicioAfectado = "No" Then
        m_EventoConServicio = False
    Else
        m_EventoConServicio = Null
    End If
    
    m_FechaRestablecimiento = Nz(Me.FechaRestablecimientoServicio, "")
    m_TipoRepInsitu = IIf(Me.chkTipoRepInsitu = True, True, IIf(Me.chkTipoRepInsitu = False, False, Null))
    m_TipoRepNoSMT = IIf(Me.chkTipoRepNoSMT = True, True, IIf(Me.chkTipoRepNoSMT = False, False, Null))
    m_TipoRepValvulas = IIf(Me.chkTipoRepValvulas = True, True, IIf(Me.chkTipoRepValvulas = False, False, Null))
    
    ' Abrir registro para edición
    m_SQL = "SELECT TbEventos.* " & _
            "FROM TbEventos " & _
            "WHERE TbEventos.IDEvento='" & m_IDEvento & "';"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If .EOF Then
            strTextoError = "No se encontró el evento"
            Err.Raise 1000
        End If
        
        .Edit
            ' TRES
            If IsDate(m_FechaRecepcion) Then
                .Fields("FechaRecepcionNotificacion") = CDate(m_FechaRecepcion)
            Else
                .Fields("FechaRecepcionNotificacion") = Null
            End If
            
            If IsDate(m_FechaInicioContacto) Then
                .Fields("FechaInicioContactoCliente") = CDate(m_FechaInicioContacto)
            Else
                .Fields("FechaInicioContactoCliente") = Null
            End If
            
            ' TRCM
            If Not IsNull(m_IncidenciaAveria) Then
                .Fields("IncidenciaAveriaOReparacion") = m_IncidenciaAveria
            Else
                .Fields("IncidenciaAveriaOReparacion") = Null
            End If
            
            If IsDate(m_FechaInicioAdquisicion) Then
                .Fields("FechaInicioTiempoAdquisicion") = CDate(m_FechaInicioAdquisicion)
            Else
                .Fields("FechaInicioTiempoAdquisicion") = Null
            End If
            
            If IsDate(m_FechaFinAdquisicion) Then
                .Fields("FechaFinTiempoAdquisicion") = CDate(m_FechaFinAdquisicion)
            Else
                .Fields("FechaFinTiempoAdquisicion") = Null
            End If
            
            If m_TipoReparacion <> "" Then
                .Fields("TipoReparacion") = m_TipoReparacion
            Else
                .Fields("TipoReparacion") = Null
            End If
            
            If Not IsNull(m_Urgente) Then
                .Fields("Urgente") = m_Urgente
            Else
                .Fields("Urgente") = Null
            End If
            
            ' TRSS
            If Not IsNull(m_EventoConServicio) Then
                .Fields("EventoConServicioAfectado") = m_EventoConServicio
            Else
                .Fields("EventoConServicioAfectado") = Null
            End If
            
            If IsDate(m_FechaRestablecimiento) Then
                .Fields("FechaRestablecimientoServicio") = CDate(m_FechaRestablecimiento)
            Else
                .Fields("FechaRestablecimientoServicio") = Null
            End If
            
            ' SLA-4
            If Not IsNull(m_TipoRepInsitu) Then
                .Fields("TipoRepInsitu") = m_TipoRepInsitu
            Else
                .Fields("TipoRepInsitu") = Null
            End If
            
            If Not IsNull(m_TipoRepNoSMT) Then
                .Fields("TipoRepNoSMT") = m_TipoRepNoSMT
            Else
                .Fields("TipoRepNoSMT") = Null
            End If
            
            If Not IsNull(m_TipoRepValvulas) Then
                .Fields("TipoRepValvulas") = m_TipoRepValvulas
            Else
                .Fields("TipoRepValvulas") = Null
            End If
            
            .Fields("FechaRegistroModificacion") = Now()
        .Update
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Error al guardar SLA: " & Err.Number & vbNewLine & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
End Sub

