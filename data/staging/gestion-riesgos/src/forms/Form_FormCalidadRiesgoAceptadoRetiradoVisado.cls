Attribute VB_Name = "Form_FormCalidadRiesgoAceptadoRetiradoVisado"
Option Compare Database
Option Explicit
Private m_EsMitigacion As String
Private m_EstadoRiesgo As EnumRiesgoEstado
Private blnRiesgoActivo As Boolean
Public Event AceptacionAprobada()
Public Event AceptacionRechazada()
Public Event AceptacionAprobadaQuitado()
Public Event AceptacionRechazadaQuitada()


Public Event RetiroAprobado()
Public Event RetiroRechazado()
Public Event RetiroAprobadoQuitado()
Public Event RetiroRechazadoQuitada()
Private m_Error As String



Private Sub ComandoAceptar_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    With m_ObjRiesgoActivo
        If m_EsMitigacion = "Sí" Then
            If Me.ComandoAceptar.Caption = "Aceptar" Then
                .AceptacionAprobar Date, m_EstadoRiesgo, m_Error
            Else
                .AceptacionAprobarQuitar m_EstadoRiesgo, m_Error
            End If
            
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            If Me.ComandoAceptar.Caption = "Aceptar" Then
                .RetiroAprobar Date, m_EstadoRiesgo, m_Error
            Else
                .RetiroAprobacionQuitar m_EstadoRiesgo, m_Error
            End If
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End With
    
    If m_EsMitigacion = "Sí" Then
        If Me.ComandoAceptar.Caption = "Aceptar" Then
            RaiseEvent AceptacionAprobada
        Else
            RaiseEvent AceptacionAprobadaQuitado
        End If
        
    Else
        If Me.ComandoAceptar.Caption = "Aceptar" Then
            RaiseEvent RetiroAprobado
        Else
            RaiseEvent RetiroAprobadoQuitado
        End If
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAceptar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoRechazar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    With m_ObjRiesgoActivo
        If m_EsMitigacion = "Sí" Then
            If Me.ComandoRechazar.Caption = "Rechazar" Then
                .AceptacionRechazo m_EstadoRiesgo, Date, m_Error
            Else
                .AceptacionRechazoQuitar m_EstadoRiesgo, m_Error
            End If
            
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            If Me.ComandoRechazar.Caption = "Rechazar" Then
                .RetiroRechazo m_EstadoRiesgo, Date, m_Error
            Else
                .RetiroRechazoQuitar m_EstadoRiesgo, m_Error
            End If
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End With
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsMitigacion = "Sí" Then
        If Me.ComandoRechazar.Caption = "Rechazar" Then
            RaiseEvent AceptacionRechazada
        Else
            RaiseEvent AceptacionRechazadaQuitada
        End If
        
    Else
        If Me.ComandoRechazar.Caption = "Rechazar" Then
            RaiseEvent RetiroRechazado
        Else
            RaiseEvent RetiroRechazadoQuitada
        End If
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRechazar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    m_EsMitigacion = Nz(Me.openArgs(), "")
    If m_EsMitigacion <> "Sí" And m_EsMitigacion <> "No" Then
        m_Error = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo Is Nothing Then
        m_Error = "El formulario se ha abierto con datos insuficientes"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.EsActivo = EnumSiNo.Sí Then
        blnRiesgoActivo = True
    Else
        blnRiesgoActivo = False
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




Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
  
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_RiesgoEnAceptacion As EnumSiNo
    
    On Error GoTo errores
    
    
    
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    
    m_RiesgoEnRetirada = RiesgoEnRetirada(m_EstadoRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoEnRetirada = EnumSiNo.No Then
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_EstadoRiesgo, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_RiesgoEnAceptacion = EnumSiNo.No Then
            p_Error = "El riesgo no está ni en estado de aceptación ni de retirada"
            Err.Raise 1000
        End If
    End If
    With m_ObjRiesgoActivo
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            If .JustificacionAceptacionRiesgo <> "" Then
                Me.Justificacion = .JustificacionAceptacionRiesgo
            End If
            If IsDate(.FechaJustificacionAceptacionRiesgo) Then
                Me.FechaJustificacion = .FechaJustificacionAceptacionRiesgo
            End If
            If IsDate(.FechaAprobacionAceptacionPorCalidad) Then
                Me.FechaAprobacion = .FechaAprobacionAceptacionPorCalidad
            End If
            If IsDate(.FechaRechazoAceptacionPorCalidad) Then
                Me.FechaRechazo = .FechaRechazoAceptacionPorCalidad
            End If
            If m_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoAceptar.Enabled = False
                Me.ComandoRechazar.Enabled = False
                
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                    Me.ComandoRechazar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.Aceptado Then
                Me.ComandoAceptar.Caption = "Quitar Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoRechazar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Quitar Rechazar"
                Me.ComandoAceptar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoRechazar.Enabled = False
                End If
            End If
        Else
            If .JustificacionRetiroRiesgo <> "" Then
                Me.Justificacion = .JustificacionRetiroRiesgo
            End If
            If IsDate(.FechaJustificacionRetiroRiesgo) Then
                Me.FechaJustificacion = .FechaJustificacionRetiroRiesgo
            End If
            If IsDate(.FechaAprobacionRetiroPorCalidad) Then
                Me.FechaAprobacion = .FechaAprobacionRetiroPorCalidad
            End If
            If IsDate(.FechaRechazoRetiroPorCalidad) Then
                Me.FechaRechazo = .FechaRechazoRetiroPorCalidad
            End If
            If m_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoAceptar.Enabled = False
                Me.ComandoRechazar.Enabled = False
                
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                    Me.ComandoRechazar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.Retirado Then
                Me.ComandoAceptar.Caption = "Quitar Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoRechazar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Quitar Rechazar"
                Me.ComandoAceptar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoRechazar.Enabled = False
                End If
            End If
        End If
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
