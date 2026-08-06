Attribute VB_Name = "Form_Form0BDOpcionesTecnico"
Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjUltimoProyecto As UltimoProyecto
Private m_ObjProyectoActivo As Proyecto
Private m_ObjEdicionActiva As Edicion

Private Sub CmdSalir_Click()
     
    On Error Resume Next
    pregunta = MsgBox("¿Desea realmente salir de la aplicación?", vbExclamation + vbYesNo, "Salir de la aplicación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    'CerrarAplicacion
    Application.Quit acQuitSaveNone
End Sub

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

Private Sub ComandoGestionesDeRiesgo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    ActualizarObjetosActivos EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormProyectosGestion") Then
        DoCmd.Close acForm, "FormProyectosGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormProyectosGestion"
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionesDeRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ComandoTareasPendientes_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    DoCmd.OpenForm "FormTecnicoTareas"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoTareasPendientes_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoUltimoProyecto_Click()
    
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    DoCmd.Hourglass True
    
    If m_ObjUltimoProyecto Is Nothing Then
       ComandoGestionesDeRiesgo_Click
       DoCmd.Hourglass False
       Exit Sub
    End If
    
    Set m_ObjProyectoActivo = m_ObjUltimoProyecto.Proyecto
    Set m_ObjEdicionActiva = m_ObjUltimoProyecto.Proyecto.EdicionUltima
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
        m_Error = "Al ComandoUltimoProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Sub ActivarBotonera()
    On Error Resume Next
    Me.ComandoTareasPendientes.Enabled = True
    Me.ComandoGestionesDeRiesgo.Enabled = True
    Me.ComandoUltimoProyecto.Enabled = True
    Me.ComandoVerControlCambios.Enabled = True
End Sub


Private Sub ComandoVerControlCambios_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    DoCmd.OpenForm "FormControlCambiosGestion"
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerControlCambios_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVolverAlPerfilInicial_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    CerrarFormulariosAbiertos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioConectadoInicialmente = Nothing
    
    EVE , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.OpenForm "Form0BDOpciones"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVolverAlPerfilInicial_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
    
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    AjustarTamaño Me
    If Nz(Me.openArgs(), "") = "" Then
        EVE , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
   
    
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.lblTitulo1.Caption = m_Titulo
    
    Me.lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion
    If m_ObjUltimoProyecto Is Nothing Then
        Me.lblUltimoProyecto.Caption = "Último Proyecto"
    Else
        Me.lblUltimoProyecto.Caption = m_ObjUltimoProyecto.Proyecto.NombreProyecto
    End If
    
    
    If Not m_ObjUsuarioConectadoInicialmente Is Nothing Then
        
        If m_ObjUsuarioConectadoInicialmente.EsTecnicoCalculado <> EnumSiNo.Sí Then
            Me.ComandoVolverAlPerfilInicial.Visible = True
            Me.ComandoVolverAlPerfilInicial.Enabled = True
        Else
            Me.ComandoVolverAlPerfilInicial.Visible = False
        End If
    Else
        Me.ComandoVolverAlPerfilInicial.Visible = False
    End If
    
    ActivarBotonera
    AvanceCerrar
    Me.lblEstado.Visible = False
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
   DoCmd.Quit acQuitSaveNone
End Sub

Private Sub lblVideo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not fso.FileExists(m_ObjEntorno.URLVideo) Then
        m_Error = "No es alcanzable el video en este momento"
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_ObjEntorno.URLVideo, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al lblVideo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub