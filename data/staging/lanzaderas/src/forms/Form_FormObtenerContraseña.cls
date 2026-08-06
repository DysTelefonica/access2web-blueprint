Option Compare Database
Option Explicit

Private m_Error As String
Private m_NombreFormulario As String

Private Sub ComandoObtener_Click()
    
    Dim m_CorreoUsuario As String
    
    Dim m_ObjUsuario As usuario
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_CorreoUsuario = Nz(Me.CorreoUsuario, "")
   
    FormatoCorreoCorrecto m_CorreoUsuario, m_Error
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuario = Constructor.getUsuario(, , , m_CorreoUsuario, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuario Is Nothing Then
        m_Error = "No hay registrado ningún usuario con ese correo. Hable con su administrador"
        Err.Raise 1000
    End If
    
    m_ObjUsuario.ResetearPass m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Pregunta = MsgBox("En un periodo máximo de 10 minutos recibirá un correo electrónico con una contraseña de un solo uso", vbInformation, "Lanzadera")
    
    DoCmd.OpenForm "Formlogin"
    DoCmd.Close acForm, "FormObtenerContraseña", acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener la contraseña se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    
    Dim strCorreoUsuario As String
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_NombreFormulario = Me.Name
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    DoCmd.Restore
    Application.Echo False

    strCorreoUsuario = Nz(Me.OpenArgs(), "")
    If strCorreoUsuario <> "" Then
        Me.CorreoUsuario = strCorreoUsuario
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, m_NombreFormulario, acSaveNo
End Sub




