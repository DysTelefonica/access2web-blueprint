Option Compare Database
Option Explicit
Private m_EdicionAlInicio As Edicion
Private m_Error As String
Public Event Motivado(p_Motivo As String)


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
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGrabar_Click()

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Nz(Me.PropuestaRechazadaPorCalidadMotivo, "") = "" Then
        m_Error = "No se puede quedar vacío el campo del motivo del rechazo"
        Err.Raise 1000
    End If
    If Me.PropuestaRechazadaPorCalidadMotivo = m_EdicionAlInicio.PropuestaRechazadaPorCalidadMotivo Then
        m_Error = "Ningún cambio por grabar"
        Err.Raise 1000
    End If
    RaiseEvent Motivado(Me.PropuestaRechazadaPorCalidadMotivo)
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoSalir_Click()
    
    On Error GoTo errores
    If Nz(Me.PropuestaRechazadaPorCalidadMotivo, "") = "" Then
        If m_EdicionAlInicio.PropuestaRechazadaPorCalidadMotivo <> "" Then
            pregunta = MsgBox("Si desea quitar el rechazo, basta con que publique", vbExclamation, "Propuesta")
            
        End If
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSalir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub Form_Load()

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    Set m_EdicionAlInicio = GetCachedEdicion(m_ObjEdicionActiva.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    
    On Error GoTo errores
    
    With m_EdicionAlInicio
        If Not IsDate(.PropuestaRechazadaPorCalidadFecha) And Not IsDate(.FechaPreparadaParaPublicar) Then
            p_Error = "La edición no aparece como propuesta para publicación"
            Err.Raise 1000
        End If
        If IsDate(.FechaPreparadaParaPublicar) Then
            Me.FechaPreparadaParaPublicar = Format(.FechaPreparadaParaPublicar, "dd/mm/yyyy")
        End If
        
        If IsDate(.PropuestaRechazadaPorCalidadFecha) Then
            Me.PropuestaRechazadaPorCalidadFecha = Format(.PropuestaRechazadaPorCalidadFecha, "dd/mm/yyyy")
        End If
        If .PropuestaRechazadaPorCalidadMotivo <> "" Then
            Me.PropuestaRechazadaPorCalidadMotivo = .PropuestaRechazadaPorCalidadMotivo
            Me.PropuestaRechazadaPorCalidadMotivo.Enabled = False
            Me.PropuestaRechazadaPorCalidadMotivo.Locked = False
            Me.ComandoGrabar.Enabled = False
        Else
            If EsTecnico <> EnumSiNo.Sí Then
                Me.PropuestaRechazadaPorCalidadMotivo.Enabled = True
                Me.PropuestaRechazadaPorCalidadMotivo.Locked = False
                Me.ComandoGrabar.Enabled = True
            Else
                Me.PropuestaRechazadaPorCalidadMotivo.Enabled = False
                Me.PropuestaRechazadaPorCalidadMotivo.Locked = False
                Me.ComandoGrabar.Enabled = False
            End If
        End If
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
