Option Compare Database
Option Explicit

Private Sub cmdSalir_Click()
    Dim strTextoError As String
    On Error GoTo errores
    pregunta = MsgBox("¿Desea salir de la aplicación", vbCritical + vbYesNo + vbDefaultButton2, "Salir de Gestión de Proyectos")
    If pregunta = vbNo Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    'RegistroCierre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Quit acQuitSaveAll
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub CuadroFormulario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    DesactivaLabel Me
End Sub

Private Sub Detalle_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Application.Echo True
End Sub
Private Sub Detalle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    DesactivaLabel Me
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me
    Me.Caption = Me.lbTitulo.Caption
    
   
    
    EVE strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnDesarrollo") = "Sí" Then
        If m_ObjUsuarioConectado.EsAdministradorCalculado = EnumSino.Sí Then
            DoCmd.ShowToolbar "Ribbon", acToolbarYes
        Else
            DoCmd.ShowToolbar "Ribbon", acToolbarNo
        End If
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
    Me.Caption = "Usuario conectado: " & m_ObjUsuarioConectado.Nombre & " " & m_TextoWin64
    
    Me.lbVersion.Caption = m_ObjEntorno.VersionAplicacion
    'RegistroApertura
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Quit acQuitSaveNone
End Sub
Private Sub lblActividades_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormActividadGestion"
End Sub
Private Sub lblActividades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblActividades.ForeColor <> longColorEtiquetaMenuPulsado Then lblActividades.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGastosYDietas_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormGastoGestion"
End Sub
Private Sub lblGastosYDietas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGastosYDietas.ForeColor <> longColorEtiquetaMenuPulsado Then lblGastosYDietas.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblInformes_Click()
    On Error Resume Next
    DoCmd.OpenForm "Form0BDInformes"
End Sub
Private Sub lblInformes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblInformes.ForeColor <> longColorEtiquetaMenuPulsado Then lblInformes.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionTecnicos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormTecnicoGestion"
End Sub
Private Sub lblGestionTecnicos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionTecnicos.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionTecnicos.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblEventos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormEventoGestion"
End Sub
Private Sub lblEventos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblEventos.ForeColor <> longColorEtiquetaMenuPulsado Then lblEventos.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionDeLosMateriales_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormMaterialesGestion"
End Sub
Private Sub lblGestionDeLosMateriales_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionDeLosMateriales.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionDeLosMateriales.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionEquipos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormEquipoGestion"
End Sub
Private Sub lblGestionEquipos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionEquipos.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionEquipos.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionFacturacion_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormFacturacionGestion"
End Sub
Private Sub lblGestionFacturacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionFacturacion.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionFacturacion.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionSubContrataciones_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormSubContratacionGestion"
End Sub
Private Sub lblGestionSubContrataciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionSubContrataciones.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionSubContrataciones.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblOtrasOpciones_Click()
    On Error Resume Next
    DoCmd.OpenForm "Form0BDOpcionesOtras"
End Sub
Private Sub lblOtrasOpciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblOtrasOpciones.ForeColor <> longColorEtiquetaMenuPulsado Then lblOtrasOpciones.ForeColor = longColorEtiquetaMenuPulsado
End Sub








