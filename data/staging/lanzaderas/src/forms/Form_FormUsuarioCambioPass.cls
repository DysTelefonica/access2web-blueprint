Option Compare Database
Option Explicit

Private m_Error As String
Private m_NombreFormulario As String

Private Sub ComandoRegistrar_Click()
    
    Dim m_PassNueva As String
    Dim m_PassNueva1 As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_ObjUsuarioConectadoLogin Is Nothing Then
        m_Error = "No hay usuario activo"
        Err.Raise 1000
    End If
    m_PassNueva = Nz(Me.PassNueva, "")
    m_PassNueva1 = Nz(Me.PassNueva1, "")
    If StrComp(m_PassNueva, m_PassNueva1, vbBinaryCompare) <> 0 Then
        Me.PassNueva1.SetFocus
        m_Error = "Las Passs no coinciden"
        Err.Raise 1000
    End If
    m_ObjUsuarioConectadoLogin.CambioPass m_PassNueva, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioConectadoLogin = Constructor.getUsuario(m_ObjUsuarioConectadoLogin.ID, , , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Pregunta = MsgBox("Se ha cambiado la Pass satisfactoriamente", vbInformation, "Lanzadera")
    If FormularioAbierto("FormLogin") Then
        DoCmd.Close acForm, "FormLogin"
    End If
    
    If FormularioAbierto("FormMenuPrincipalAdmin") Then
        DoCmd.Close acForm, "FormMenuPrincipalAdmin"
    End If
    If FormularioAbierto("FormMenuPrincipalUsuario") Then
        DoCmd.Close acForm, "FormMenuPrincipalUsuario"
    End If
    
    DoCmd.OpenForm "FormLogin"
    Forms("FormLogin").Controls("Pass").Value = Null
    DoCmd.Close acForm, m_NombreFormulario, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar cambiar la Pass se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComandoSalir_Click()
   On Error Resume Next
   DoCmd.Close acForm, Me.Name
End Sub
Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Ajustar Me
    m_Error = ""
    m_NombreFormulario = Me.Name
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    DoCmd.Restore

    If m_ObjUsuarioConectadoLogin Is Nothing Then
        m_Error = "No hay usuario activo"
        Err.Raise 1000
    End If
    Me.CorreoUsuario = m_ObjUsuarioConectadoLogin.CorreoUsuario
    Me.Nombre = m_ObjUsuarioConectadoLogin.Nombre
    
    Me.Caption = Me.lblTitulo1.Caption
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, m_NombreFormulario, acSaveNo
End Sub

Private Sub ComandoAyudaFormulario_Click()
    On Error Resume Next
    AbrirArchivoAyuda Me
End Sub



