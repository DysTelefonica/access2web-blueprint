Option Compare Database
Option Explicit
Private m_Error As String
Public Event Publicado()


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

Private Sub ComandoGenerarInforme_Click()
    Dim url As String
    Dim alcance As EnumControlCambiosAlcance
    Dim generarPDF As Boolean
    Dim valorAlcance As String
    Dim tipoDocumento As String
    
    On Error GoTo errores
    m_Error = ""
    
    If m_ObjEdicionActiva.TodosLosRiesgosPriorizados = EnumSiNo.No Then
        pregunta = MsgBox("Para poder generar el informe ha de tener todos sus riesgos priorizados", vbExclamation, "Error")
        Exit Sub
    End If
    
    valorAlcance = modFormInteractionHelper.FormInteraction_ObtenerTextoControl(Me, "cboControlCambiosAlcance", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    tipoDocumento = modFormInteractionHelper.FormInteraction_ObtenerTextoControl(Me, "cboTipoDocumento", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    alcance = ControlCambios_ResolverAlcanceDesdeSelector(valorAlcance, EnumControlCambiosAlcanceResumen3)
    generarPDF = InformePublicacionDebeGenerarPDFDesdeTipoDocumento(tipoDocumento)
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    url = GenerarInformePublicacion( _
            p_Edicion:=m_ObjEdicionActiva, _
            p_FechaCierre:="", _
            p_FechaPublicacion:="", _
            p_GenerarPDF:=generarPDF, _
            p_Error:=m_Error, _
            p_ControlCambiosAlcance:=alcance)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' NOTA: La función GenerarInformeEdicionHTML ya tiene internamente un 'Ejecutar p_hWnd, "open"...'
    ' Si notas que el documento se abre dos veces seguidas, simplemente comenta o borra la siguiente línea:
    Ejecutar Me.hWnd, "open", url, "", "", 1
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    AvanceCerrar
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoPublicar_Click()
    Dim m_fechaUltimaEdicion As String
    Dim m_FechaRef As String
    Dim m_ConCorreoAlRAC As EnumSiNo
    Dim alcance As EnumControlCambiosAlcance
    Dim valorAlcance As String
    
    On Error GoTo errores
    
    
    m_Error = ""
   
   
    
    pregunta = MsgBox("¿Desea enviar el correo de publicación también al RAC?", vbExclamation + vbYesNo + vbDefaultButton2, "Con envío al RAC")
    If pregunta = vbYes Then
        If Nz(Me.CorreoRAC, "") = "" Then
            pregunta = MsgBox("No está definido el correo del RAC en esta gestión de riesgos", vbExclamation, _
            "Con envío al RAC")
            Exit Sub
        End If
        m_ConCorreoAlRAC = EnumSiNo.Sí
    Else
        m_ConCorreoAlRAC = EnumSiNo.No
        
    End If
    If Nz(Me.CorreoRAC, "") <> m_ObjEdicionActiva.Proyecto.CorreoRAC Then
        If Nz(Me.CorreoRAC, "") <> "" Then
            m_ObjEdicionActiva.Proyecto.RegistrarCorreoRAC Me.CorreoRAC, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.ComboUltimaEdicion, "") = "Sí" Then
        m_fechaUltimaEdicion = Date
    End If
    valorAlcance = modFormInteractionHelper.FormInteraction_ObtenerTextoControl(Me, "cboControlCambiosAlcance", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    alcance = ControlCambios_ResolverAlcanceDesdeSelector(valorAlcance, EnumControlCambiosAlcanceResumen3)
    'm_FechaRef = "11/07/2025"
    m_URLInforme = Publicar(p_Edicion:=m_ObjEdicionActiva, _
                            p_FechaCierre:=m_fechaUltimaEdicion, _
                            p_RegenerandoCambiosDesdeElInicio:=EnumSiNo.No, _
                            p_FechaSiguientePublicacion:=Nz(Me.FechaMaxProximaPublicacion, ""), _
                            p_fechaRef:=m_FechaRef, _
                            p_ControlCambiosAlcance:=alcance, _
                            p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo = GetCachedProyecto(m_ObjProyectoActivo.IDProyecto, m_Error)
    RefrescarTareasCalidadDespuesPublicacion m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyectoActivo.ParaInformeAvisos = "Sí" Then
        EnvioCorreoNuevaPublicacion m_ObjProyectoActivo, m_URLInforme, m_ObjProyectoActivo.EdicionUltimaPublicada, _
                m_FechaRef, Nz(Me.ComboUltimaEdicion, ""), m_ConCorreoAlRAC, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    AvanceCerrar
    
    If fso.FileExists(m_URLInforme) Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("¿Desea ver el informe de la edición generada?", vbExclamation + vbYesNo + vbDefaultButton1, "Informe generado")
        If pregunta = vbYes Then
            m_URLHTMLActivo = m_URLInforme
'            If FormularioAbierto("FormWeb") Then
'                DoCmd.Close acForm, "FormWeb", acSaveNo
'            End If
'            DoCmd.OpenForm "FormWeb"
           Ejecutar Me.hWnd, "open", m_URLInforme, "", "", 1
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
'    If FormularioAbierto("FormRiesgosGestion") Then
'        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
'    End If
    RaiseEvent Publicado
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    If FormularioAbierto("FormPublicacionCalidad") Then
        DoCmd.Close acForm, "FormPublicacionCalidad", acSaveNo
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    
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

Private Sub RefrescarTareasCalidadDespuesPublicacion(ByRef p_Error As String)
    Dim m_Servicio As CalidadPublicacionService
    Dim m_Resultado As CalidadPublicacionResultado
    Dim m_IDRiesgo As String
    Dim m_Reseteando As EnumSiNo
    Dim m_Regularizando As EnumSiNo

    On Error GoTo errores

    p_Error = ""
    m_IDRiesgo = ResolverIDRiesgoRefrescoPublicacion(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If m_IDRiesgo = "" Then
        m_Reseteando = EnumSiNo.Sí
        m_Regularizando = EnumSiNo.No
        EstablecerTareasCalidad m_Reseteando, m_Regularizando, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Sub
    End If

    Set m_Servicio = New CalidadPublicacionService
    Set m_Resultado = m_Servicio.SaveAndRefreshTareasCalidad( _
                        CStr(m_ObjEdicionActiva.IDEdicion), _
                        m_IDRiesgo, _
                        "calidad", _
                        p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Resultado Is Nothing Then
        p_Error = "No se pudo refrescar las tareas de calidad tras publicar"
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If p_Error = "" Then
        p_Error = "El método RefrescarTareasCalidadDespuesPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Sub

Private Function ResolverIDRiesgoRefrescoPublicacion(Optional ByRef p_Error As String) As String
    ' Issue #39: precondition decision delegated to PublicacionCalidadExecutionService.
    ' Form keeps only the local variable plumbing.
    Dim m_Service As PublicacionCalidadExecutionService

    On Error GoTo errores

    Set m_Service = New PublicacionCalidadExecutionService
    ResolverIDRiesgoRefrescoPublicacion = m_Service.ResolveRefreshRiskID( _
        m_ObjRiesgoActivo, m_ObjEdicionActiva, p_Error)
    Exit Function

errores:
    p_Error = "El método ResolverIDRiesgoRefrescoPublicacion ha devuelto el error: " & vbNewLine & Err.Description
End Function

Private Sub ComandoSalir_Click()
    
    On Error GoTo errores
    
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

Private Sub EstablecerFechaMax()
    
    Dim m_FechaMax As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_FechaMax = CalcularFechaMaxProximaPublicacion(p_Edicion:=m_ObjEdicionActiva, _
                                                    p_FechaPublicacion:=Date, _
                                                    p_EsUltimaEdicion:=Me.ComboUltimaEdicion, _
                                                    p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaMax) Then
        Me.FechaMaxProximaPublicacion = m_FechaMax
    Else
        Me.FechaMaxProximaPublicacion = Null
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al EstablecerFechaMax se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComboUltimaEdicion_AfterUpdate()
    EstablecerFechaMax
End Sub

Private Sub Form_Load()
    Dim m_FechaPrevistaCierre As String
    Dim m_CorreoRAC As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    With m_ObjEdicionActiva.Proyecto
        m_FechaPrevistaCierre = .FechaPrevistaCierreCalculada
        If IsDate(m_FechaPrevistaCierre) Then
            If DateDiff("d", CDate(m_FechaPrevistaCierre), Date) <= 15 Then
                Me.ComboUltimaEdicion = "Sí"
            Else
                Me.ComboUltimaEdicion = "No"
            End If
        Else
            Me.ComboUltimaEdicion = "No"
        End If
        Me.ComboUltimaEdicion = "No"
        If IsDate(.Expediente.FechaInicioContrato) Then
            Me.FechaInicioContrato = .Expediente.FechaInicioContrato
        End If
        If IsDate(.Expediente.FechaFinContrato) Then
            Me.FechaFinContrato = .Expediente.FechaFinContrato
        End If
        m_CorreoRAC = .CorreoRACCalculado
        If m_CorreoRAC <> "" Then
            Me.CorreoRAC = m_CorreoRAC
        End If
    End With
    EstablecerFechaMax
   
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



