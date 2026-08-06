Option Compare Database
Option Explicit
Private m_RiesgoAlInicio As riesgo
Private m_Estado As EnumRiesgoEstado
Private m_EsAceptacion As String
Public Event Justificado(m_Justificacion As String)

Private m_Error As String

Private Sub ComandoGrabar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_EsAceptacion = "Sí" Then
        If Nz(Me.Justificacion, "") = "" Then
            m_Error = "No se ha justificado"
            Err.Raise 1000
        End If
        If m_RiesgoAlInicio.JustificacionAceptacionRiesgo = Me.Justificacion Then
            m_Error = "No ha habido cambios"
            Err.Raise 1000
        End If
        If m_RiesgoAlInicio.FechaAprobacionAceptacionPorCalidad <> "" Then
            m_Error = "Ya ha sido aprobado por Calidad"
            Err.Raise 1000
        End If
'        m_ObjRiesgoActivo.Mitigacion = "Aceptar"
'        m_ObjRiesgoActivo.JustificacionAceptacionRiesgo = Me.Justificacion
'        m_ObjRiesgoActivo.FechaJustificacionAceptacionRiesgo = Date
'        m_ObjRiesgoActivo.FechaMitigacionAceptar = Date
        m_ObjRiesgoActivo.AceptacionRegistrar Me.Justificacion, Date, m_Estado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    Else
        If Nz(Me.Justificacion, "") = "" Then
            m_Error = "No se ha justificado"
            Err.Raise 1000
        End If
        If m_RiesgoAlInicio.JustificacionRetiroRiesgo = Me.Justificacion Then
            m_Error = "No ha habido cambios"
            Err.Raise 1000
        End If
        If m_RiesgoAlInicio.FechaAprobacionRetiroPorCalidad <> "" Then
            m_Error = "Ya ha sido aprobado por Calidad"
            Err.Raise 1000
        End If
        m_ObjRiesgoActivo.RetiroRegistrar Me.Justificacion, Date, m_Estado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    RaiseEvent Justificado(Me.Justificacion)
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

    m_EsAceptacion = Nz(Me.openArgs(), "")
    If m_EsAceptacion <> "Sí" And m_EsAceptacion <> "No" Then
        m_Error = "Ha de ser aceptación o retiro"
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
   
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion
    Dim m_Titulo As String
    On Error GoTo errores
    
    
    
    Set m_RiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoAlInicio Is Nothing Then
        p_Error = "No hay datos del Riesgo"
        Err.Raise 1000
    End If
    
    With m_ObjRiesgoActivo
        m_Edicion = .Edicion
        If m_Edicion.EsActivo = EnumSiNo.No Then
            blnPermitidoEditar = False
        Else
            If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                blnPermitidoEditar = True
            
            Else
                blnPermitidoEditar = True
            End If
        End If
        m_Estado = .EstadoEnum
        If m_EsAceptacion = "Sí" Then
            m_Titulo = "Justificación de aceptación del riesgo"
            If .JustificacionAceptacionRiesgo <> "" Then
                Me.Justificacion = .JustificacionAceptacionRiesgo
            End If
            If IsDate(.FechaJustificacionAceptacionRiesgo) Then
                Me.Fecha = .FechaJustificacionAceptacionRiesgo
            End If
            If IsDate(.FechaAprobacionAceptacionPorCalidad) Then
                Me.FechaAprobacion = .FechaAprobacionAceptacionPorCalidad
            End If
            If IsDate(.FechaRechazoAceptacionPorCalidad) Then
                Me.FechaRechazo = .FechaRechazoAceptacionPorCalidad
            End If
            If IsNumeric(.DiasRespuestaCalidadAceptacion) Then
                Me.DiasSinRespuesta = .DiasRespuestaCalidadAceptacion
            End If
        Else
             m_Titulo = "Justificación del retiro del riesgo"
            If .JustificacionRetiroRiesgo <> "" Then
                Me.Justificacion = .JustificacionRetiroRiesgo
            End If
            If IsDate(.FechaJustificacionRetiroRiesgo) Then
                Me.Fecha = .FechaJustificacionRetiroRiesgo
            End If
            If IsDate(.FechaAprobacionRetiroPorCalidad) Then
                Me.FechaAprobacion = .FechaAprobacionRetiroPorCalidad
            End If
            If IsDate(.FechaRechazoRetiroPorCalidad) Then
                Me.FechaRechazo = .FechaRechazoRetiroPorCalidad
            End If
            If IsNumeric(.DiasRespuestaCalidadRetiro) Then
                Me.DiasSinRespuesta = .DiasRespuestaCalidadRetiro
            End If
        End If
        If blnPermitidoEditar = False Then
            Me.Justificacion.Enabled = False
            Me.ComandoGrabar.Enabled = False
        Else
            If .FechaAprobacionAceptacionPorCalidad <> "" Or .FechaAprobacionRetiroPorCalidad <> "" Then
                Me.Justificacion.Enabled = False
                Me.ComandoGrabar.Enabled = False
            Else
                Me.Justificacion.Enabled = True
                Me.ComandoGrabar.Enabled = True
            End If
            
        End If
    End With
    Me.lblTitulo1.Caption = m_Titulo
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
