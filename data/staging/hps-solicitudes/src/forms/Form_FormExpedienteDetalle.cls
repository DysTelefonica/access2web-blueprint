 Option Compare Database
Option Explicit


Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    Me.Caption = Me.lblTitulo.Caption
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

   
    On Error GoTo errores
    If m_ObjExpedienteActivo Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    With m_ObjExpedienteActivo
        Me.IDExpediente = .IDExpediente
        If .Nemotecnico <> "" Then
            Me.Nemotecnico = .Nemotecnico
        End If
        If .Titulo <> "" Then
            Me.Titulo = .Titulo
        End If
        If .CodExp <> "" Then
            Me.CodExp = .CodExp
        End If
        If .CodExpLargo <> "" Then
            Me.CodExpLargo = .CodExpLargo
        End If
        If .Ordinal <> "" Then
            Me.Ordinal = .Ordinal
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

