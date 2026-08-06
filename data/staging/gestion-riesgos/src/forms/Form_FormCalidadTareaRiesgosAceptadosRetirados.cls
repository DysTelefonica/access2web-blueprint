Attribute VB_Name = "Form_FormCalidadTareaRiesgosAceptadosRetirados"
Option Compare Database
Option Explicit

Private m_EsAceptacion As String  'Sí o No
Private m_Riesgo As riesgo
Private m_Error As String
Private WithEvents m_FormVisado As Form_FormCalidadRiesgoAceptadoRetiradoVisado
' ==========================================
' MSG BOX HELPER (per access-vba-tdd -0 - no interactive UI in tests)
' p_PromptResult: -1 (default) = show MsgBox; 0 = skip MsgBox (test mode)
' ==========================================
Private Function MostrarMsgBoxSiNoEsTest(ByVal p_Message As String, _
                                       ByVal p_Style As VbMsgBoxStyle, _
                                       ByVal p_Title As String, _
                                       Optional ByVal p_PromptResult As Long = -1) As Long
    If p_PromptResult = 0 Then
        MostrarMsgBoxSiNoEsTest = vbOK
    Else
        MostrarMsgBoxSiNoEsTest = MsgBox(p_Message, p_Style, p_Title)
    End If
End Function

Private Sub ComandoIrACalidad_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormCalidadRiesgoAceptadoRetiradoVisado") Then
        DoCmd.Close acForm, "FormCalidadRiesgoAceptadoRetiradoVisado", acSaveNo
    End If
    If m_Riesgo Is Nothing Then
        Set m_Riesgo = Form_FormCalidadTareas.m_RiesgoSeleccionado
        If m_Riesgo Is Nothing Then
            If Nz(Me.IDRiesgo, "") = "" Then
                m_Error = "No se sabe el riesgo seleccionado"
                Err.Raise 1000
            End If
            Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=Me.IDRiesgo, p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Riesgo Is Nothing Then
                m_Error = "No se sabe el riesgo seleccionado"
                Err.Raise 1000
            End If
        End If
    End If

    Set m_ObjRiesgoActivo = m_Riesgo
    DoCmd.OpenForm "FormCalidadRiesgoAceptadoRetiradoVisado", , , , , , m_EsAceptacion
    If FormularioAbierto("FormCalidadRiesgoAceptadoRetiradoVisado") Then
        Set m_FormVisado = Forms("FormCalidadRiesgoAceptadoRetiradoVisado")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrACalidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerRiesgo_Click()

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormRiesgo") Then
        DoCmd.Close acForm, "FormRiesgo", acSaveNo
    End If
    m_EsAlta = EnumSiNo.No
    If m_Riesgo Is Nothing Then
        Set m_Riesgo = Form_FormCalidadTareas.m_RiesgoSeleccionado
        If m_Riesgo Is Nothing Then
            If Nz(Me.IDRiesgo, "") = "" Then
                m_Error = "No se sabe el riesgo seleccionado"
                Err.Raise 1000
            End If
            Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=Me.IDRiesgo, p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_Riesgo Is Nothing Then
                m_Error = "No se sabe el riesgo seleccionado"
                Err.Raise 1000
            End If
        End If
    End If
    Set m_ObjRiesgoActivo = m_Riesgo
    DoCmd.OpenForm "FormRiesgo"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub Form_Open(Cancel As Integer)


    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    If Form_FormCalidadTareas.m_RiesgoSeleccionado Is Nothing Then
        Exit Sub
    End If
    Set m_Riesgo = Form_FormCalidadTareas.m_RiesgoSeleccionado

    If m_Riesgo Is Nothing Then
        m_Error = "No se sabe el riesgo activo"
        Err.Raise 1000
    End If
    Me.AllowEdits = False


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
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub m_FormVisado_AceptacionAprobada()

    On Error GoTo errores
    m_Error = ""
    FormOutcomeHandler_Run "AceptacionAprobada", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormVisado_AceptacionAprobada se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String

   
    Dim m_Estado As EnumRiesgoEstado
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    
   
    
    On Error GoTo errores
    
    
   
    Me.ComandoVerRiesgo.Enabled = False
    Me.ComandoIrACalidad.Enabled = False
    
    
    If m_Riesgo Is Nothing Then
        p_Error = "No se conoce el riesgo activo"
        Err.Raise 1000
    End If
    If EsAdministrador = EnumSiNo.Sí Or EsAdministradorConectadoInicialmente = EnumSiNo.Sí Then
        Me.lblIDExpediente.Visible = True
        Me.IDExpediente.Visible = True
        Me.lblIDProyecto.Visible = True
        Me.IDProyecto.Visible = True
        Me.lblIDEdicion.Visible = True
        Me.IDEdicion.Visible = True
        Me.lblIDRiesgo.Visible = True
        Me.IDRiesgo.Visible = True
    Else
        Me.lblIDExpediente.Visible = False
        Me.IDExpediente.Visible = False
        Me.lblIDProyecto.Visible = False
        Me.IDProyecto.Visible = False
        Me.lblIDEdicion.Visible = False
        Me.IDEdicion.Visible = False
        Me.lblIDRiesgo.Visible = False
        Me.IDRiesgo.Visible = False
    End If
    
    With m_Riesgo
        Me.IDExpediente = .Edicion.Proyecto.IDExpediente
        Me.IDProyecto = .Edicion.IDProyecto
        Me.IDEdicion = .IDEdicion
        Me.IDRiesgo = .IDRiesgo
        
        
        Me.Proyecto = .Edicion.Proyecto.Proyecto
        Me.NombreProyecto = .Edicion.Proyecto.NombreProyecto
        Me.Supervisor = .Edicion.Proyecto.NombreUsuarioCalidad
        Me.Descripcion = "( " & .CodigoRiesgo & ") " & .Descripcion
        Me.CausaRaiz = .CausaRaiz
        Me.Autorizados = Replace(.Edicion.Proyecto.CadenaNombreAutorizadosCalculados, ";", vbNewLine)
        
        m_Estado = .EstadoEnum
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_Estado)
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            m_EsAceptacion = "Sí"
        Else
            m_RiesgoEnRetirada = RiesgoEnRetirada(m_Estado)
            If m_RiesgoEnRetirada = EnumSiNo.Sí Then
                m_EsAceptacion = "No"
            Else
                p_Error = "El Riesgo no está camino de aceptarse ni de retirarse"
                Err.Raise 1000
            End If
        End If
        If EsTecnico <> EnumSiNo.Sí Then
            Me.ComandoIrACalidad.Enabled = True
        Else
            Me.ComandoIrACalidad.Enabled = False
        End If
        If m_EsAceptacion = "Sí" Then
            Me.lblAceptacionRetiro.Caption = "Mitigación"
            Me.Fecha = Format(.FechaJustificacionAceptacionRiesgo, "dd/mm/yyyy")
            Me.MitigacionRetiro = "Aceptar"
            If .JustificacionAceptacionRiesgo <> "" Then
                Me.Justificacion = .JustificacionAceptacionRiesgo
            End If
            If IsDate(.FechaAprobacionAceptacionPorCalidad) Then
                Me.FechaAprobacion = Format(.FechaAprobacionAceptacionPorCalidad, "dd/mm/yyyy")
            End If
            If IsDate(.FechaRechazoAceptacionPorCalidad) Then
                Me.FechaAprobacion = Format(.FechaRechazoAceptacionPorCalidad, "dd/mm/yyyy")
            End If
            If Not IsNumeric(.DiasSinRespuestaCalidadAceptacion) Then
                Me.DiasSinRespuesta = .DiasSinRespuestaCalidadAceptacionCalculado
            Else
                 Me.DiasSinRespuesta = .DiasSinRespuestaCalidadAceptacion
            End If
            
        Else
            Me.lblAceptacionRetiro.Caption = "F. Retirado"
            Me.Fecha = Format(.FechaRetirado, "dd/mm/yyyy")
            Me.MitigacionRetiro = Me.Fecha
            If .JustificacionRetiroRiesgo <> "" Then
                Me.Justificacion = .JustificacionRetiroRiesgo
            End If
            If IsDate(.FechaAprobacionRetiroPorCalidad) Then
                Me.FechaAprobacion = Format(.FechaAprobacionRetiroPorCalidad, "dd/mm/yyyy")
            End If
            If IsDate(.FechaRechazoRetiroPorCalidad) Then
                Me.FechaAprobacion = Format(.FechaRechazoRetiroPorCalidad, "dd/mm/yyyy")
            End If
            If Not IsNumeric(.DiasSinRespuestaCalidadRetiro) Then
                Me.DiasSinRespuesta = .DiasSinRespuestaCalidadRetiroCalculado
            Else
                 Me.DiasSinRespuesta = .DiasSinRespuestaCalidadRetiro
            End If
        End If
    End With
    If EsTecnico <> EnumSiNo.Sí Then
        Me.ComandoVerRiesgo.Enabled = True
    Else
        Me.ComandoVerRiesgo.Enabled = False
    End If
    
    Me.AllowEdits = False
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub m_FormVisado_AceptacionAprobadaQuitado()
    On Error GoTo errores
    m_Error = ""
    FormOutcomeHandler_Run "AceptacionAprobadaQuitado", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormVisado_AceptacionAprobadaQuitado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub m_FormVisado_AceptacionRechazada()
    On Error GoTo errores
    m_Error = ""
    FormOutcomeHandler_Run "AceptacionRechazada", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormVisado_AceptacionRechazada se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Sub m_FormVisado_AceptacionRechazadaQuitada()
    On Error GoTo errores
    m_Error = ""
    FormOutcomeHandler_Run "AceptacionRechazadaQuitada", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormVisado_AceptacionRechazadaQuitada se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbCritical, "Error")
    Else
        pregunta = MostrarMsgBoxSiNoEsTest(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
