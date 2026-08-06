Option Compare Database
Option Explicit
Private m_Error As String
Private m_CorreoRechazo As CORREO

Private WithEvents m_FormAnexos As Form_FormAnexos
Private WithEvents m_FormPublicar As Form_FormPublicacionCalidadPublicarEjecutar
Private WithEvents m_FormMotivos As Form_FormPublicacionCalidadRechazoObservaciones

Private Sub ComandoEnviarCorreoTecnicoPropuestaPublicacion_Click()
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision
    
    On Error GoTo errores
    
    m_Error = ""
    pregunta = MsgBox("¿Desea enviar un correo a los responsables para que revisen la edición para la propuesta de publicación?", vbExclamation + vbYesNo + vbDefaultButton1, "Correo a responsables")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_EdicionCorreoRevision = New EdicionCorreoRevision
    With m_EdicionCorreoRevision
        .IDEdicion = m_ObjEdicionActiva.IDEdicion
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjEdicionActiva = GetCachedEdicion(p_IDEdicion:=m_ObjEdicionActiva.IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    ' HR3e-1 (issue #109): el handler NO envía correo en este thread — solo
    ' registra la EdicionCorreoRevision y dispara AvanceCerrar. El correo
    ' sale de forma asíncrona más tarde. Mensaje honesto para no mentir al
    ' usuario. Cuando se introduzca el helper EsPropuestaPublicacionNotificada()
    ' (similar a EsRechazoPropuestaNotificado en Funciones Generales.bas:534),
    ' este MsgBox se puede mover a la bifurcación If Notificado/Else.
    pregunta = MsgBox("Revisión de propuesta registrada. El correo a responsables se enviará al finalizar.", vbInformation, "Correo a responsables")

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnviarCorreoTecnicoPropuestaPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGenerarInformePublicabilidad_Click()
    Dim m_URL As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_URL = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=m_ObjEdicionActiva, p_hWnd:=Me.hWnd, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLHTMLActivo = m_URL
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInformePublicabilidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
   
End Sub

Private Sub ComandoHistorialPublicaciones_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    HTMLENTXT m_ObjEdicionActiva.HTMLHistorialPublicaciones, , EnumSiNo.No, m_Error
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoHistorialPublicaciones_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoVerCorreosEnviados_Click()
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    HTMLENTXT m_ObjEdicionActiva.HTMLHistorialcorreosResponsables, , EnumSiNo.No, m_Error
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerCorreosEnviados_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerRechazo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        DoCmd.Close acForm, "FormPublicacionCalidadRechazoObservaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadRechazoObservaciones"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRechazo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoPublicar_Click()
    
    
    
    On Error GoTo errores
    
    
    m_Error = ""
   
    If FormularioAbierto("FormPublicacionCalidadPublicarEjecutar") Then
        DoCmd.Close acForm, "FormPublicacionCalidadPublicarEjecutar", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadPublicarEjecutar"
    If FormularioAbierto("FormPublicacionCalidadPublicarEjecutar") Then
        Set m_FormPublicar = Forms("FormPublicacionCalidadPublicarEjecutar")
    End If
    
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPublicar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRechazar_Click()
    On Error GoTo errores
    
    
    m_Error = ""
    
    If m_ObjEdicionActiva.PublicacionRechazada = EnumSiNo.Sí Then
        pregunta = MsgBox("En Estos momentos ya está rechazada la propuesta de publicación" & vbNewLine & _
                        "¿Desea ver las observaciones registradaSí", vbExclamation + vbYesNo + vbDefaultButton1, "Rechazo")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    Else
        pregunta = MsgBox("Para poder rechazar la propuesta rellene, en el siguiente formulario, los motivos", vbExclamation + vbYesNo + vbDefaultButton1, "Rechazo")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        DoCmd.Close acForm, "FormPublicacionCalidadRechazoObservaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadRechazoObservaciones"
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        Set m_FormMotivos = Forms("FormPublicacionCalidadRechazoObservaciones")
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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


Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    

    
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay ninguna edición activa"
        Err.Raise 1000
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub


' HR4-sister (HR3d follow-up #9, HIGH-PRIORITY - 2026-07-01):
'   Era Public Function EstablecerDatos. Ahora Private.
'   Body es un thin UI adapter: delega al presenter
'   PublicacionCalidadStatusPresenter y aplica el view state via
'   AplicarEstadoPublicacionCalidad (Private Sub, no extraida).
'   Razon: el body ya es UI adapter + presenter pattern (no business logic
'   para extraer). La visibilidad Public era innecesaria (1 caller
'   interno en Form_Open). Hard rule 4: Public business methods son
'   anti-patron. Sin external callers -> Private.
Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_HaySum As Boolean
    Dim m_StatusPresenter As PublicacionCalidadStatusPresenter
    Dim m_StatusViewState As PublicacionCalidadStatusViewState

    On Error GoTo errores
    m_Error = ""

    With m_ObjEdicionActiva


        m_ObjEdicionActiva.Proyecto.ColSuministradoresVaciar
        If m_ObjEdicionActiva.Proyecto.ColSuministradores Is Nothing Then
            m_HaySum = False
        Else
            m_HaySum = True
        End If

        Set m_StatusPresenter = New PublicacionCalidadStatusPresenter
        Set m_StatusViewState = m_StatusPresenter.Build( _
                .Proyecto.EnUTE = "Sí", _
                .TieneAnexoEvidenciaUTE = EnumSiNo.Sí, _
                m_HaySum, _
                .EvidenciasSuministradoresCompletadas = EnumSiNo.Sí, _
                .EsActivo = EnumSiNo.Sí, _
                .FechaPreparadaParaPublicar, _
                .PropuestaRechazadaPorCalidadFecha, _
                .FechaPublicacion)

        AplicarEstadoPublicacionCalidad m_StatusViewState

    End With



    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Establecer datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub AplicarEstadoPublicacionCalidad( _
    ByVal p_State As PublicacionCalidadStatusViewState _
)
    With p_State
        Me.lbl1Requerido.Caption = .UteRequeridoCaption
        Me.lbl1Complementado.Caption = .UteComplementadoCaption
        Me.lbl1Complementado.ForeColor = ColorEstadoPublicacionCalidad(.UteComplementadoColorRole)

        Me.lbl2Requerido.Caption = .SuministradoresRequeridoCaption
        Me.lbl2Complementado.Caption = .SuministradoresComplementadoCaption
        Me.lbl2Complementado.ForeColor = ColorEstadoPublicacionCalidad(.SuministradoresComplementadoColorRole)

        Me.ComandoRechazar.Enabled = .ComandoRechazarEnabled
        Me.ComandoPublicar.Enabled = .ComandoPublicarEnabled
        Me.ComandoPublicar.Caption = .ComandoPublicarCaption
        Me.ComandoGenerarInformePublicabilidad.Enabled = .ComandoGenerarInformePublicabilidadEnabled
        Me.ComandoEnviarCorreoTecnicoPropuestaPublicacion.Enabled = .ComandoEnviarCorreoTecnicoPropuestaPublicacionEnabled
        If .HasComandoVerRechazoVisibleDecision Then
            Me.ComandoVerRechazo.Visible = .ComandoVerRechazoVisible
        End If
        Me.lblPublicacion.Visible = .PublicacionVisible
        If .PublicacionVisible Then
            Me.lblPublicacion.Caption = .PublicacionCaption
            Me.lblPublicacion.ForeColor = ColorEstadoPublicacionCalidad(.PublicacionColorRole)
        End If
    End With
End Sub

Private Function ColorEstadoPublicacionCalidad( _
    ByVal p_ColorRole As String _
) As Long
    If p_ColorRole = "no-complementado" Then
        ColorEstadoPublicacionCalidad = m_ColorNoComplementado
    Else
        ColorEstadoPublicacionCalidad = m_ColorComplementado
    End If
End Function


Private Sub m_FormMotivos_Motivado(p_Motivo As String)
    
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If p_Motivo = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    With m_ObjEdicionActiva
        .PropuestaRechazadaPorCalidadMotivo = p_Motivo
        Set m_CorreoRechazo = .RechazoPropuestaParaPublicacion(m_Error)
        If m_Error <> "" Then
           Err.Raise 1000
        End If
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Dim m_Notificado As Boolean
    m_Notificado = EsRechazoPropuestaNotificado(m_CorreoRechazo)
    If m_Notificado Then
        pregunta = MsgBox("Propuesta de publicación rechazada. Le ha llegado un correo a los responsables del proyecto", vbInformation, "Rechazo de propuesta")
    Else
        pregunta = MsgBox("Rechazo registrado. La notificación por correo no se envía para este proyecto.", vbInformation, "Rechazo de propuesta")
    End If
    If FormularioAbierto("FormPublicacionCalidad") Then
        DoCmd.Close acForm, "FormPublicacionCalidad", acSaveNo
    End If
    ' HR3d: cross-form Forms("FormRiesgosGestion") + FormDetalle.SourceObject
    '   navigation + subform.Form + call to global EstablecerlblRechazado
    '   -> Coord_EstablecerLblRechazadoEnDetalleRiesgos (thin proxy que
    '   encapsula la navegacion). El caller ya no toca Forms("X") ni
    '   mForm.FormDetalle.
    Call modFormCoordinationHelper.Coord_EstablecerLblRechazadoEnDetalleRiesgos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormMotivos_Motivado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub m_FormPublicar_Publicado()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    DoCmd.Close acForm, Me.Name
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormPublicar_Publicado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
