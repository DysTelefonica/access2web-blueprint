Option Compare Database
Option Explicit
Private WithEvents m_FormRiesgosBibliotecaGestion As Form_FormRiesgosBibliotecaGestion

Private m_Riesgo As riesgo
Private m_ObjRiesgoAlInicio As riesgo
Private m_Error As String


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
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




Private Function EstablecerDatos(Optional ByRef p_Error As String) As String

    On Error GoTo errores
    p_Error = ""

    Me.ComandoVerRiesgo.Enabled = False
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.AllowEdits = False
    Else
       Me.AllowEdits = True

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
    Set m_Riesgo = Form_FormCalidadTareas.m_RiesgoSeleccionado
    If m_Riesgo Is Nothing Then
        p_Error = "No se ha podido determinar el riesgo"
        Err.Raise 1000
    End If
    Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_Riesgo.IDRiesgo, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjRiesgoAlInicio Is Nothing Then
        p_Error = "No se ha podido determinar el riesgo"
        Err.Raise 1000
    End If
    
    
    With m_ObjRiesgoAlInicio
        Me.IDExpediente = .Edicion.Proyecto.IDExpediente
        Me.IDProyecto = .Edicion.IDProyecto
        Me.IDEdicion = .IDEdicion
        Me.IDRiesgo = .IDRiesgo
        Me.Proyecto = .Edicion.Proyecto.Proyecto
        Me.Supervisor = .Edicion.Proyecto.NombreUsuarioCalidad
        Me.CodigoRiesgo = .CodigoRiesgo
        Me.CodRiesgoBiblioteca = .CodRiesgoBiblioteca
        If .RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            Me.TextoPendienteRetipificacion.Visible = True
            Me.TextoPendienteRetipificacion.value = "Pendiente de Calidad"
        Else
            Me.TextoPendienteRetipificacion.Visible = False

        End If
        Me.Descripcion = .Descripcion
        Me.CausaRaiz = .CausaRaiz
        Me.FechaRiesgoParaRetipificar = .FechaRiesgoParaRetipificar
        Me.DiasSinRetipificar = .DiasSinRespuestaCalidadRetipificacionCalculado

    End With
    Me.ComandoVerRiesgo.Enabled = True
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub Form_Load()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""


    
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

Private Sub m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido(m_ObjRiesgoBiblioteca As RiesgoBiblioteca)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    With m_ObjRiesgoBiblioteca
        Me.CodRiesgoBiblioteca = .codigo
        m_Riesgo.CodRiesgoBiblioteca = Nz(Me.CodRiesgoBiblioteca, "")

        If .RiesgoParaRetipificar = EnumSiNo.No Then
            Me.Descripcion = .Descripcion
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = True
        Else
            Me.Descripcion = Null
            Me.Descripcion.Locked = False
            Me.Descripcion.Enabled = False
        End If
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub