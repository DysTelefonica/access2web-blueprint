Option Compare Database
Option Explicit


Private m_Error As String
Public Event Alta()

Private Sub ComandoRegistrar_Click()
    
    Dim m_ObjFiesta As Fiesta
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Set m_ObjFiesta = New Fiesta
    
    With m_ObjFiesta
        .FechaFiesta = Nz(Me.FechaFiesta, "")
        .Observaciones = Nz(Me.Observaciones, "")
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RaiseEvent Alta
    End With
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea introducir alguna fiesta más?", vbExclamation + vbYesNo + vbDefaultButton1, "Alta de Festivo")
    If pregunta <> vbYes Then
        DoCmd.Close acForm, Me.Name, acSaveNo
        Exit Sub
    End If
    Me.FechaFiesta = Null
    Me.Observaciones = Null
    Me.FechaFiesta.SetFocus
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub Form_Load()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
