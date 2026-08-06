Attribute VB_Name = "Form_FormInicial08UsuariosHistoricosDatos"
Option Compare Database
Option Explicit
Private WithEvents m_Indicadores As clsIndicadoresBus
Private m_Error As String

Private Sub Form_Current()
    Dim m_ID As String
    On Error GoTo errores

    m_Error = ""
    m_ID = Nz(Me.ID, "")

    Set m_ObjUsuarioHistoricoActivo = Constructor.getUsuarioHistorico(m_ID, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Current se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Load()
    Me.RecordSource = "SELECT * " & _
                    "FROM TbUsuariosHistoricosLocal;"

    Me.Requery
    Set m_Indicadores = modIndicadores.m_Indicadores
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.Requery
    On Error GoTo 0
End Sub

Private Sub m_Indicadores_LifecycleMoveCompleted(p_Direction As String, p_UserId As String)
    On Error Resume Next
    Me.Requery
    On Error GoTo 0
End Sub