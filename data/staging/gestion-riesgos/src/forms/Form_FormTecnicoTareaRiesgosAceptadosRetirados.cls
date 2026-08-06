Option Compare Database
Option Explicit
Private m_EsAceptacion As String  'Sí o No
Private m_Riesgo As riesgo
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


Private Sub Form_Open(Cancel As Integer)
   
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Dim m_RiesgoSeleccionado As Object
    Set m_RiesgoSeleccionado = modFormCoordinationHelper.Coord_GetRiesgoSeleccionadoDeTecnicoTareas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        m_Error = "No se sabe el riesgo seleccionado"
        Err.Raise 1000
    End If
    Set m_Riesgo = m_RiesgoSeleccionado
    
    If m_Riesgo Is Nothing Then
        m_Error = "No se sabe el riesgo activo"
        Err.Raise 1000
    End If
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
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

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    Dim lst As ListBox
    
    On Error GoTo errores
    
    
    Set lst = Me.ListaAutorizados
    lst.RowSource = ""
    
    
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
        
        lst.RowSource = .Edicion.Proyecto.CadenaNombreAutorizados
       
        If .Estado = "" Then
            .Estado = .ESTADOCalculadoTexto
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If InStr(1, .Estado, "Acepta") <> 0 Then
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
            m_EsAceptacion = "Sí"
            
        ElseIf InStr(1, .Estado, "Retira") <> 0 Then
            Me.lblAceptacionRetiro.Caption = "F. retirado"
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
            m_EsAceptacion = "No"
        Else
            Me.lblAceptacionRetiro.Caption = "#ERROR"
        End If
    End With
    Me.ComandoVerRiesgo.Enabled = True
    Me.AllowEdits = False
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function