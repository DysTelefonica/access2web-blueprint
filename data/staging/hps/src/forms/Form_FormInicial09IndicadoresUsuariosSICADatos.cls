Attribute VB_Name = "Form_FormInicial09IndicadoresUsuariosSICADatos"
Option Compare Database
Option Explicit
Private m_Error As String




Private Sub Form_Current()
    Dim m_ID As String
    On Error Resume Next
    m_ID = Nz(Me.ID, "")
    If Err.Number <> 0 Then
         Me.Parent.ComandoEditarUsuario.Enabled = False
    Else
        Me.Parent.ComandoEditarUsuario.Enabled = True
    End If
    EstablecerSemaforo m_ID
    Err.Clear
    
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    
    m_Error = ""
    CopiarDatosAIndicadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.RecordSource = "SELECT TbUsuariosSICALocalParaIndicadores.* " & _
                    "FROM TbUsuariosSICALocalParaIndicadores WHERE id='-1';"
    Me.Requery
    Me.Parent.ComandoEditarUsuario.Enabled = False
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El Form_Load ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
