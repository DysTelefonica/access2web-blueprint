Option Compare Database
Option Explicit


Private m_Error As String

Private Sub cmdSalir_Click()
    
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerLista(Optional ByRef p_Error As String) As String

    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores

    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set parsed = JsonConverter.ParseJson(mod0BDGestorEntidadesHelper.GestorEntidades_ListRowSource(m_ObjEntorno.OtrasOpciones, p_Error))
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set payload = parsed("payload")
    Me.ListaOtrasOpciones.RowSource = CStr(payload("rowSource"))
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ComandoAbrirFormulario_Click()
    
    Dim m_NombreFormulario As String
    Dim parsed As Object
    Dim payload As Object
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDGestorEntidadesHelper.GestorEntidades_OpenDecision(CStr(Nz(Me.ListaOtrasOpciones.Column(0), "")), m_ObjEntorno.OtrasOpcionesFormularios, m_Error))
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set payload = parsed("payload")
    m_NombreFormulario = CStr(payload("formName"))
    DoCmd.OpenForm m_NombreFormulario
    ListaOtrasOpciones_Click
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAbrirFormulario_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me
    
    If Application.TempVars("EnDesarrollo") <> "Sí" Then
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End If
    Me.Caption = m_TituloFormulario
    Me.ComandoAbrirFormulario.Enabled = True
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaOtrasOpciones_Click()
    Dim parsed As Object
    Dim payload As Object

    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDGestorEntidadesHelper.GestorEntidades_SelectionState(CStr(Nz(Me.ListaOtrasOpciones.Column(0), "")), m_Error))
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set payload = parsed("payload")
    Me.ComandoAbrirFormulario.Enabled = CBool(payload("enableOpen"))
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaOtrasOpciones_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaOtrasOpciones_DblClick(Cancel As Integer)
    Dim parsed As Object
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(mod0BDGestorEntidadesHelper.GestorEntidades_DoubleClickDecision(Me.ComandoAbrirFormulario.Enabled))
    If CStr(parsed("payload")("action")) = "open" Then
        ComandoAbrirFormulario_Click
    End If
End Sub
