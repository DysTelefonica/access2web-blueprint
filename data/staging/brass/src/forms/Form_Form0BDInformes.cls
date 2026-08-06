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



Private Sub Form_Open(Cancel As Integer)
    
    On Error Resume Next
    AjustarTamaño Me
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    Me.Caption = Me.lbTitulo.Caption
End Sub




Private Sub lblConsultarHorasTecnico_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormTecnicoGestion"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub lblConsultarHorasTecnico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblConsultarHorasTecnico.ForeColor <> 26367 Then lblConsultarHorasTecnico.ForeColor = 26367
    
End Sub


Private Sub lblGestionDePartes_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormParteGestion"
    DoCmd.Close acForm, Me.Name
End Sub



Private Sub lblGestionDePartes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblGestionDePartes.ForeColor <> 26367 Then lblGestionDePartes.ForeColor = 26367
    
End Sub

Private Sub PieDelFormulario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    DesactivaLabel Me
End Sub


Private Sub lblInformeParaRAC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblInformeParaRAC.ForeColor <> 26367 Then lblInformeParaRAC.ForeColor = 26367
    
End Sub
Private Sub lblInformeParaRAC_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormInformesParaRAC"
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub lblInformeParaRACPorFactura_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormInformesPorFactura"
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub lblInformeParaRACPorFactura_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblInformeParaRACPorFactura.ForeColor <> 26367 Then lblInformeParaRACPorFactura.ForeColor = 26367
    
End Sub
Private Sub lblInformeParaRACDesdeLista_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormInformesDesdeLista"
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub lblInformeParaRACDesdeLista_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblInformeParaRACDesdeLista.ForeColor <> 26367 Then lblInformeParaRACDesdeLista.ForeColor = 26367
    
End Sub

Private Sub lblInformeSLA_Click()
    On Error Resume Next
    DoCmd.OpenForm "FormInformeSLA"
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub lblInformeSLA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lblInformeSLA.ForeColor <> 26367 Then lblInformeSLA.ForeColor = 26367
    
End Sub
