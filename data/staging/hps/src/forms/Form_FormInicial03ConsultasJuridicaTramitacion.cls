Attribute VB_Name = "Form_FormInicial03ConsultasJuridicaTramitacion"
Option Compare Database
Option Explicit

Private m_Error As String



Private Sub Filtrar()
    Dim m_Suministrador As Suministrador
    On Error GoTo errores
    m_Error = ""
    If m_ObjColParaFiltros Is Nothing Then
        Set m_ObjColParaFiltros = New Scripting.Dictionary
        m_ObjColParaFiltros.CompareMode = TextCompare
    End If
    m_ObjColParaFiltros("Estado") = Nz(Me.ComboEstadoHPS, "")
    Set m_Suministrador = Constructor.getSuministrador(p_ID:=Nz(Me.IDEmpresaHPS.Column(0), ""))
    If Not m_Suministrador Is Nothing Then
        m_ObjColParaFiltros("EmpresaTramitadora") = m_Suministrador.Nombre
    End If
    
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

Private Sub ComandoLimpiarJuridicaQueTramitaHPS_Click()
    Me.IDEmpresaHPS = Null
    IDEmpresaHPS_AfterUpdate
End Sub
Private Sub ComandoLimpiarEstadoHPS_Click()
    Me.ComboEstadoHPS = Null
    ComboEstadoHPS_AfterUpdate
End Sub
Private Sub IDEmpresaHPS_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al IDEmpresaHPS ha devuelto el error:" & vbNewLine & Err.Description
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
Private Sub Form_Load()
    
    On Error GoTo errores
    
     m_Error = ""
    Set m_ObjColParaFiltros = Nothing
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboEmpresas Me.IDEmpresaHPS, EnumSiNo.Sí, , m_Error
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