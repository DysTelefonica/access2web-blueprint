Option Compare Database
Option Explicit
Private m_Error As String


Private Sub ComandoRegistrar_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    With m_ObjEdicionActiva
        .NotasCalidadParaPublicar = Nz(Me.NotasCalidadParaPublicar, "")
        .RegistrarNotasParaCalidad p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If Nz(Me.NotasCalidadParaPublicar, "") = "" Then
        pregunta = MsgBox("Notas eliminadas correctamente", vbInformation, "Notas Calidad")
    Else
        pregunta = MsgBox("Notas registradas correctamente", vbInformation, "Notas Calidad")
    End If
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    EstablecerDatos m_Error
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
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    
    On Error GoTo errores
    If m_ObjEdicionActiva Is Nothing Then
        Exit Function
    End If
    If m_ObjEdicionActiva.NotasCalidadParaPublicar <> "" Then
        Me.NotasCalidadParaPublicar = m_ObjEdicionActiva.NotasCalidadParaPublicar
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        Me.ComandoRegistrar.Enabled = True
    Else
        Me.ComandoRegistrar.Enabled = False
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
