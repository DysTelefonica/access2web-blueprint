Option Compare Database
Option Explicit
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Detalle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    DesactivaLabel Me
End Sub

Private Sub Detalle_DblClick(Cancel As Integer)
    Application.Echo False
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Application.Echo True
End Sub

Private Sub lblAnexosNoAlcanzables_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormAnexosNoAlcanzables"
End Sub

Private Sub lblEquiposMedida_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormEquiposMedidaGestion"
End Sub

Private Sub lblEquiposMedida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblEquiposMedida.ForeColor <> longColorEtiquetaMenuPulsado Then lblEquiposMedida.ForeColor = longColorEtiquetaMenuPulsado
End Sub

Private Sub lblMtoProg_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormPlanificacionGestion"
End Sub
Private Sub lblMtoProg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblMtoProg.ForeColor <> longColorEtiquetaMenuPulsado Then lblMtoProg.ForeColor = longColorEtiquetaMenuPulsado
End Sub

Private Sub lblCalibraciones_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormCalibracionEquipoGestion"
End Sub
Private Sub lblCalibraciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblCalibraciones.ForeColor <> longColorEtiquetaMenuPulsado Then lblCalibraciones.ForeColor = longColorEtiquetaMenuPulsado
End Sub


Private Sub Form_Open(Cancel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
'    Me.InsideHeight = 6480
'    Me.InsideWidth = 10260
    AjustarTamaño Me
    Me.Caption = Me.lbTitulo.Caption
    
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub lblGestionBUI_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormBUIGestion"
End Sub
Private Sub lblGestionBUI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionBUI.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionBUI.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionIDEquipos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormEquipoGestion"
End Sub
Private Sub lblGestionIDEquipos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionIDEquipos.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionIDEquipos.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionNodos_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormNodoGestion"
End Sub
Private Sub lblGestionNodos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionNodos.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionNodos.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionOriginadores_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormOriginadorGestion"
End Sub
Private Sub lblGestionOriginadores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionOriginadores.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionOriginadores.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionSubSistemas_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormSubSistemaGestion"
End Sub
Private Sub lblGestionSubSistemas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionSubSistemas.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionSubSistemas.ForeColor = longColorEtiquetaMenuPulsado
End Sub

Private Sub lblGestionTipoActividad_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormActividadTipoGestion"
End Sub
Private Sub lblGestionTipoActividad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionTipoActividad.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionTipoActividad.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblGestionTipoEventos_Click()
     On Error Resume Next
    DoCmd.OpenForm "FormEventoTipoGestion"
End Sub
Private Sub lblGestionTipoEventos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionTipoEventos.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionTipoEventos.ForeColor = longColorEtiquetaMenuPulsado
End Sub

Private Sub lblGestionUbicaciones_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormUbicacionGestion"
End Sub
Private Sub lblGestionUbicaciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionUbicaciones.ForeColor <> longColorEtiquetaMenuPulsado Then lblGestionUbicaciones.ForeColor = longColorEtiquetaMenuPulsado
End Sub
Private Sub lblRelacionNodoBUI_Click()
     On Error Resume Next
    DoCmd.OpenForm "FormNodoBUIGestion"
End Sub
Private Sub lblRelacionNodoBUI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     On Error Resume Next
    If lblRelacionNodoBUI.ForeColor <> longColorEtiquetaMenuPulsado Then lblRelacionNodoBUI.ForeColor = longColorEtiquetaMenuPulsado
End Sub


