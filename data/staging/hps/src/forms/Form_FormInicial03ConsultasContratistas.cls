Attribute VB_Name = "Form_FormInicial03ConsultasContratistas"
Option Compare Database
Option Explicit

Private m_Error As String

Private Sub ComandoLimpiarJuridicaContrato_Click()
    Me.CadenaContratistas = Null
    CadenaContratistas_AfterUpdate
End Sub
Private Sub ComandoLimpiarEstadoHPS_Click()
    Me.ComboEstadoHPS = Null
    ComboEstadoHPS_AfterUpdate
End Sub

Private Sub CadenaContratistas_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al CadenaContratistas_AfterUpdate ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComboEstadoHPS_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEstadoHPS ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub Filtrar()
    
    On Error GoTo errores
    If m_ObjColParaFiltros Is Nothing Then
        Set m_ObjColParaFiltros = New Scripting.Dictionary
        m_ObjColParaFiltros.CompareMode = TextCompare
    End If
    m_Error = ""
    m_ObjColParaFiltros("Estado") = Nz(Me.ComboEstadoHPS, "")
    m_ObjColParaFiltros("CadenaContratistas") = Nz(Me.CadenaContratistas, "")
    EstablecerFiltroUsuarios Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al filtrar ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub Form_Load()
    
    On Error GoTo errores
    
    m_Error = ""
    Set m_ObjColParaFiltros = Nothing
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarCadenaContratistas Me.CadenaContratistas, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    RellenarComboEstado Me.ComboEstadoHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
