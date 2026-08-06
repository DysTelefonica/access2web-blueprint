Attribute VB_Name = "Form_FormInicial03ConsultasDatos"
Option Compare Database
Option Explicit
Private m_Error As String
Public Event UsuarioSeleccionado()

Public Sub Form_Current()

    Dim m_ID As String
    On Error GoTo errores

    m_ID = Nz(Me.ID, "")
    EstablecerSemaforo m_ID
    Set m_ObjUsuarioActivo = Constructor.getUsuarioHPS(m_ID, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Current ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Private Sub Form_Load()
    Dim m_SQL As String
    m_SQL = "SELECT TbDatosLocal.* FROM TbDatosLocal;"

   
    Me.RecordSource = m_SQL
    Me.Requery
End Sub