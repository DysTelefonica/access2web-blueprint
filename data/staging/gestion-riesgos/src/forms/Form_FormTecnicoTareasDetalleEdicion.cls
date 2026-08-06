Option Compare Database
Option Explicit
Private m_Error As String
Private m_Edicion As Edicion

Private Sub ComandoInformePublicabilidad_Click()
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
     m_URL = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=m_Edicion, p_hWnd:=Me.hWnd, p_Error:=m_Error)
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
        m_Error = "Al ComandoInformePublicabilidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComandoIrAPublicar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjProyectoActivo = m_Edicion.Proyecto
    Set m_ObjEdicionActiva = m_Edicion
    
    If FormularioAbierto("FormPublicacionCalidad") Then
        DoCmd.Close acForm, "FormPublicacionCalidad", acSaveNo
    End If
    If FormularioAbierto("FormPublicacionTecnico") Then
        DoCmd.Close acForm, "FormPublicacionTecnico", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionTecnico"
    
    
       
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrAPublicar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Public Sub ComandoVerDetalle_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjProyectoActivo = m_Edicion.Proyecto
    Set m_ObjEdicionActiva = m_Edicion
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
     VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerDetalle_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    If Form_FormTecnicoTareas.m_EdicionSeleccionada Is Nothing Then
        m_Error = "No se conoce la edición activa"
        Err.Raise 1000
    End If
    Set m_Edicion = Form_FormTecnicoTareas.m_EdicionSeleccionada
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
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_FechaProxPublicacion As String
    Dim m_FechaPrevistaCierre As String
    On Error GoTo errores
    
    If EsAdministrador = EnumSiNo.Sí Or EsAdministradorConectadoInicialmente = EnumSiNo.Sí Then
        Me.lblIDExpediente.Visible = True
        Me.IDExpediente.Visible = True
        Me.lblIDProyecto.Visible = True
        Me.IDProyecto.Visible = True
        Me.lblIDEdicion.Visible = True
        Me.IDEdicion.Visible = True
    Else
        Me.lblIDExpediente.Visible = False
        Me.IDExpediente.Visible = False
        Me.lblIDProyecto.Visible = False
        Me.IDProyecto.Visible = False
        Me.lblIDEdicion.Visible = False
        Me.IDEdicion.Visible = False
    End If
    
    Me.ComandoVerDetalle.Enabled = False
   
    With m_Edicion.Proyecto
        Me.IDExpediente = .IDExpediente
        Me.IDProyecto = .IDProyecto
        Me.IDEdicion = m_Edicion.IDEdicion
        Me.Proyecto = .Proyecto
        Me.NombreProyecto = .NombreProyecto
        Me.Supervisor = .NombreUsuarioCalidad
        If Not .EdicionUltima Is Nothing Then
            Me.EdicionActual = .EdicionUltima.Edicion
            If .FechaCierre <> "" Then
                Me.EsEdicionFinal = "Sí"
            Else
                Me.EsEdicionFinal = "No"
            End If
            Me.FechaPreparadaParaPublicar = Format(.EdicionUltima.FechaPreparadaParaPublicar, "dd/mm/yyyy")
            If Form_FormTecnicoTareas.m_EdicionSeleccionada.NombreUsuarioProponePublicar <> "" Then
                Me.UsuarioProponePublicar = m_Edicion.NombreUsuarioProponePublicar
            End If
            m_FechaProxPublicacion = .EdicionUltima.FechaMaxProximaPublicacion
            p_Error = .EdicionUltima.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If IsDate(m_FechaProxPublicacion) Then
                Me.FechaMaxProximaPublicacion = Format(m_FechaProxPublicacion, "dd/mm/yyyy")
            End If
            If IsDate(.FechaCierre) Then
                Me.FechaCierre = Format(.FechaCierre, "dd/mm/yyyy")
            End If
            m_FechaPrevistaCierre = .FechaPrevistaCierreCalculada
            If IsDate(m_FechaPrevistaCierre) Then
                Me.FechaPrevistaCierre = Format(m_FechaPrevistaCierre, "dd/mm/yyyy")
            End If
            Me.Autorizados = Replace(.CadenaNombreAutorizadosCalculados, ";", vbNewLine)
        End If
    End With
    Me.DiasSinPublicar = m_Edicion.DiasSinPublicar
    Me.ComandoVerDetalle.Enabled = True
    Me.AllowEdits = False
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function