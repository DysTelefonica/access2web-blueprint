Option Compare Database
Option Explicit
Private m_Error As String



Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
   
    
    m_Error = ""
    Me.UsuariosConectados = Nz(Me.OpenArgs(), "")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub


