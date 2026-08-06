Attribute VB_Name = "Form_Form0BDOpciones"
Option Compare Database
Option Explicit
Private WithEvents m_FormGestionRiesgos As Form_FormGestionRiesgos
Private m_Error As String
'-------------------------------------------
' Nombre: EstablecerModeloCadenaJerarquica
' Propósito: Permitir cambiar la bandera entre "antiguo" y "nuevo"
' Parámetros:
'   - sModelo As String: "antiguo" o "nuevo"
' Retorno: Boolean (True si OK)
'-------------------------------------------
Public Function EstablecerModeloCadenaJerarquica( _
                                                    sModelo As String _
                                                    ) As Boolean
    On Error GoTo ManejoError
    Dim sVal As String
    sVal = LCase$(Trim$(Nz(sModelo, "")))
    If sVal <> "antiguo" And sVal <> "nuevo" Then
        Err.Raise 1000, , "Valor inválido. Use 'antiguo' o 'nuevo'."
    End If
    
    Application.TempVars("CadenaJerarquicaModelo") = sVal
    
    EstablecerModeloCadenaJerarquica = True
    Exit Function
ManejoError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    EstablecerModeloCadenaJerarquica = False
End Function


    
Private Sub CmdSalir_Click()
     
    On Error Resume Next
    pregunta = MsgBox("¿Desea realmente salir de la aplicaciónº", vbExclamation + vbYesNo, "Salir de la aplicación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    'CerrarAplicacion
    Application.Quit acQuitSaveNone
    
End Sub

Private Sub ComandoEntrarComo_Click()

    Dim m_Usuario As Usuario
    Dim m_NombreUsuario As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_Usuario = Constructor.getUsuario(, , Nz(Me.EntrarComo, ""), , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Usuario Is Nothing Then
        m_Error = "Usuario no registrado"
        Err.Raise 1000
    End If
    If m_Usuario.EsTecnicoCalculado <> EnumSiNo.Sí Then
        m_Error = "El usuario seleccionado no tiene perfil Técnico"
        Err.Raise 1000
    End If
    CerrarFormulariosAbiertos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EVE m_Usuario.CorreoUsuario, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.OpenForm "Form0BDOpcionesTecnico", , , , , , "Para Otro Usuario"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEntrarComo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoInformeIndicador_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If EsTecnico = EnumSiNo.Sí Then
        m_Error = "Es un apartado reservado para usuarios de Calidad"
        Err.Raise 1000
    End If
    DoCmd.OpenForm "FormIndicador"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoInformeIndicador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
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
    t1 = Timer
    ActualizarObjetosActivos EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    t2 = Timer
    'Debug.Print "ActualizarObjetosActivos", t2 - t1
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
    If EsTecnico = EnumSiNo.Sí Then
        m_Error = "Sólo puede hacer esta operación un miembro de Calidad"
        Err.Raise 1000
    End If
    DoCmd.OpenForm "FormCalidadTareas"
    
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
Private Sub ComandoBibliotecaRiesgos_Click()
    
    Dim m_FormRiesgosBibliotecaGestion As Form
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        DoCmd.Close acForm, "FormRiesgosBibliotecaGestion", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosBibliotecaGestion"
    If FormularioAbierto("FormRiesgosBibliotecaGestion") Then
        Set m_FormRiesgosBibliotecaGestion = Forms("FormRiesgosBibliotecaGestion")
        ' HR3e (2026-07-01): reemplazar acceso directo a control de otro
        ' form .cls por helper modFormCoordinationHelper.
        Call modFormCoordinationHelper.Coord_SetComandoNoExisteRiesgoVisible( _
            p_Visible:=False, _
            p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBibliotecaRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoNuevaGestionDeRiesgo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If EsTecnico = EnumSiNo.Sí Then
        m_Error = "Sólo los miembros de Calidad pueden generar el alta de una gestión de riesgos" & vbNewLine & _
            "Mire a ver si ya la ha dado de alta"
        Err.Raise 1000
    End If
    ActualizarObjetosActivos EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    
    Set m_ObjProyectoActivo = Nothing
    Set m_ObjProyectoAlInicio = Nothing
    m_EsAlta = EnumSiNo.Sí
    
    
    
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    If FormularioAbierto("FormGestionRiesgos") Then
        Set m_FormGestionRiesgos = Forms("FormGestionRiesgos")
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
   
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevaGestionDeRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    
    Set m_ObjProyectoActivo = Constructor.getProyecto(p_IDProyecto:=m_ObjUltimoProyecto.IDProyecto, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEdicionActiva = m_ObjProyectoActivo.EdicionUltima
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
    Me.ComandoNuevaGestionDeRiesgo.Enabled = True
    Me.ComandoBibliotecaRiesgos.Enabled = True
    
    Me.ComandoInformeIndicador.Enabled = True
    Me.ComandoUltimoProyecto.Enabled = True
    Me.ComandoVerControlCambios.Enabled = True
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


Private Sub EntrarComo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ComandoEntrarComo.Enabled = False Then
            Exit Sub
        End If
        Me.ComandoEntrarComo.SetFocus
        Me.EntrarComo.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoEntrarComo_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub

Private Sub Form_Activate()
    AvanceCerrar
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    AjustarTamaño Me
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.lblTitulo1.Caption = m_Titulo
    
    Avance "Obteniendo las tareas"
    Me.EntrarComo.RowSource = m_ObjEntorno.ListaUsuarios
    If m_ObjUltimoProyecto Is Nothing Then
        Me.lblUltimoProyecto.Caption = "Último Proyecto"
    Else
        Me.lblUltimoProyecto.Caption = m_ObjUltimoProyecto.Proyecto.NombreProyecto
    End If
    Me.lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion
    
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
