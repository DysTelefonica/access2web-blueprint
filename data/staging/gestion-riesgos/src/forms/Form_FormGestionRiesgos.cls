Attribute VB_Name = "Form_FormGestionRiesgos"
Option Compare Database
Option Explicit
Public Event Alta()
Public Event Edicion()
Public blnPermitidoEditar As Boolean
Private blnAltaProyecto As Boolean
Private blnEdicionProyecto As Boolean
Private m_Error As String



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
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
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
    If Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales" Then
        modGestionRiesgosDatosGeneralesHelper.DatosGenerales_RellenarDatosAlObjeto _
            Me.FormDetalle.Form, m_ObjProyectoActivo, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    grabar EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_EsAlta = EnumSiNo.No
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    pregunta = MsgBox("Datos guardados con éxito", vbInformation, "Guardar")
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function grabar(Optional ByRef p_SinEco As EnumSiNo = EnumSiNo.No, Optional ByRef p_Error As String) As String
    
    Dim blnHaHabidocambios As Boolean
    Dim blnAvisarNoInforme As Boolean
    Dim saveDecisionService As GestionRiesgosSaveDecisionService
    Dim m_Usuario As Usuario
    Dim m_IDProyecto As String
    Dim m_Form As Form
    On Error GoTo errores
    If p_SinEco = Empty Then
        p_SinEco = EnumSiNo.No
    End If
    If m_ObjProyectoActivo Is Nothing Then
        Set m_ObjProyectoActivo = New Proyecto
        If Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales" Then
            With Me.FormDetalle
                m_ObjProyectoActivo.IDExpediente = .Controls("IDExpediente")
                m_ObjProyectoActivo.ParaInformeAvisos = .Controls("ParaInformeAvisos")
                m_ObjProyectoActivo.EnUTE = .Controls("EnUTE")
                m_ObjProyectoActivo.Elaborado = .Controls("Elaborado")
                m_ObjProyectoActivo.Revisado = .Controls("Revisado")
                m_ObjProyectoActivo.Aprobado = .Controls("Aprobado")
            End With
        End If
    End If
    blnHaHabidocambios = HaHabidoCambios(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not blnHaHabidocambios Then
        p_Error = "Ningún cambio que grabar"
        Err.Raise 1000
    End If
    
    Set saveDecisionService = New GestionRiesgosSaveDecisionService
    blnAvisarNoInforme = saveDecisionService.DebeAvisarNoInforme(p_SinEco, m_ObjProyectoAlInicio, m_ObjProyectoActivo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_SinEco = EnumSiNo.No Then
        If blnAvisarNoInforme = True Then
            
            pregunta = MsgBox("Al poner que no es para informes, implica que este proyecto no va a generar ningún correo electrónico" & vbNewLine & _
                                "¿Está de acuerdo con eso?", vbExclamation + vbYesNo + vbDefaultButton2, "Sin informes")
            If pregunta <> vbYes Then
                AvanceCerrar
                
                Exit Function
            End If
            
        End If
    End If
    
    With m_ObjProyectoActivo
        
        .Registrar m_ObjProyectoAlInicio, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_IDProyecto = .IDProyecto
        If m_ObjProyectoAlInicio Is Nothing Then
            
            'se ha de generar la primera edición
           Set m_ObjEdicionActiva = .EdicionActiva
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se ha generado la primera edición automáticamente"
                Err.Raise 1000
            End If
            If EsTecnico = EnumSiNo.Sí Then
                EstablecerTareasTecnico EnumSiNo.Sí, p_Error
            Else
                EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
            End If
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            RaiseEvent Alta
            
        Else
            RaiseEvent Edicion
                
        End If
        m_EsAlta = EnumSiNo.No
        Set m_ObjProyectoActivo = GetCachedProyecto(m_IDProyecto, p_Error)
        Set m_ObjProyectoAlInicio = GetCachedProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .EstablecerFechaMaxPublicacion (p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    AvanceCerrar
    
    Me.NavResponsables.Enabled = True
    Me.NavRiesgosOferta.Enabled = True
    Me.NavSuministradores.Enabled = True
    Me.lblPendienteGrabacion.Visible = False
    Me.ComandoRegistrarRiesgos.Visible = True
    Me.ComandoRegistrarRiesgos.Enabled = True
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Grabar ha devuelto el error: " & Err.description
    End If
End Function

Private Sub ComandoRegistrarRiesgos_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSalir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
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
    Me.InsideHeight = 9195
    Me.InsideWidth = 13440
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.NavDatosGenerales.NavigationTargetName = "FormGestionRiesgosDatosGenerales"
    Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales"
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim builder As GestionRiesgosViewStateBuilder
    Dim state As GestionRiesgosViewState
    Dim strNombreProyecto As String
    
    
    On Error GoTo errores
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    
    Set builder = New GestionRiesgosViewStateBuilder
    If m_EsAlta = Empty Then
        Set state = builder.Build(m_EsAlta, _
                                  EsTecnico, _
                                  vbNullString, _
                                  p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If

    If m_ObjProyectoActivo Is Nothing Then
        Set m_ObjProyectoActivo = New Proyecto
        Me.ComandoRegistrarRiesgos.Visible = False
    Else
        Me.ComandoRegistrarRiesgos.Visible = True
    End If

    If m_EsAlta <> Empty And m_EsAlta <> EnumSiNo.Sí Then
        If m_ObjProyectoAlInicio Is Nothing Then
            Set m_ObjProyectoAlInicio = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjProyectoAlInicio Is Nothing Then
                p_Error = "No se conocen los datos de la gestión de riesgos"
                Err.Raise 1000
            End If
        End If
        strNombreProyecto = m_ObjProyectoActivo.NombreProyecto
    End If

    Set state = builder.Build(m_EsAlta, _
                              EsTecnico, _
                              strNombreProyecto, _
                              p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Me.NavResponsables.Enabled = state.NavResponsablesEnabled
    Me.NavRiesgosOferta.Enabled = state.NavRiesgosOfertaEnabled
    Me.NavSuministradores.Enabled = state.NavSuministradoresEnabled
    Me.NavDatosGenerales.Enabled = state.NavDatosGeneralesEnabled
    Me.lblTitulo.Caption = state.Titulo
    blnPermitidoEditar = state.PermitidoEditar
    Me.ComandoGrabar.Enabled = state.ComandoGrabarEnabled
    'Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales"
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean

    Dim detector As GestionRiesgosChangeDetector
    On Error GoTo errores

    Set detector = New GestionRiesgosChangeDetector
    HaHabidoCambios = detector.HaHabidoCambios(m_ObjProyectoAlInicio, m_ObjProyectoActivo, p_Error)

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en el método HaHabidoCambios" & vbNewLine & Err.description
    End If
End Function
