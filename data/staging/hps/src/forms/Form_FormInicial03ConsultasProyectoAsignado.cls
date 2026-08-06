Attribute VB_Name = "Form_FormInicial03ConsultasProyectoAsignado"
Option Compare Database
Option Explicit

Private m_Error As String


Private Sub Filtrar()
    
    On Error GoTo errores
    If m_ObjColParaFiltros Is Nothing Then
        Set m_ObjColParaFiltros = New Scripting.Dictionary
        m_ObjColParaFiltros.CompareMode = TextCompare
    End If
    m_ObjColParaFiltros("Estado") = Nz(Me.ComboEstadoHPS, "")
    m_ObjColParaFiltros("IDExpediente") = Nz(Me.ComboProyectoAsignado, "")
    
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

Private Sub ComandoLimpiarProyectoAsignado_Click()
    Me.ComboProyectoAsignado = Null
    ComboProyectoAsignado_AfterUpdate
End Sub
Private Sub ComandoLimpiarEstadoHPS_Click()
    Me.ComboEstadoHPS = Null
    ComboEstadoHPS_AfterUpdate
End Sub
Private Sub ComboProyectoAsignado_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboProyectoAsignado ha devuelto el error:" & vbNewLine & Err.Description
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
        m_Error = "ComboEstadoHPS ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
   
    Set m_ObjColParaFiltros = Nothing
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarCombo Me.ComboProyectoAsignado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
'    RellenarComboExpediente Me.ComboProyectoAsignado, m_Error
'    If m_Error <> "" Then
'        Err.Raise 1000
'    End If
    RellenarComboEstado Me.ComboEstadoHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function RellenarCombo( _
                                p_Cbo As ComboBox, _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_sTextoExpediente As String
    Dim m_Nemotecnico As String
    Dim m_CodExp As String
    Dim m_IDExpediente As Variant
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    m_SQL = "SELECT  distinct TbExpedientes.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuarios ON TbExpedientes.IDExpediente = TbUsuarios.IDExpediente " & _
            "ORDER BY TbExpedientes.CodExp"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
                m_IDExpediente = .Fields("IDExpediente").value
                m_Nemotecnico = Nz(.Fields("Nemotecnico").value, "")
                m_CodExp = Nz(.Fields("CodExp").value, "")
                If m_Nemotecnico <> "" Then
                    m_sTextoExpediente = m_Nemotecnico
                Else
                    m_sTextoExpediente = m_CodExp
                End If
                If m_Col Is Nothing Then
                    Set m_Col = New Scripting.Dictionary
                    m_Col.CompareMode = TextCompare
                End If
                If Not m_Col.Exists(CStr(m_IDExpediente)) Then
                    m_Col.Add CStr(m_IDExpediente), m_IDExpediente & ";" & m_sTextoExpediente
                End If
                
                .MoveNext
            Loop
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT  distinct TbExpedientes.* " & _
            "FROM TbExpedientes INNER JOIN TbUsuariosHistoricos ON TbExpedientes.IDExpediente = TbUsuariosHistoricos.IDExpediente " & _
            "ORDER BY TbExpedientes.CodExp"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
                m_IDExpediente = .Fields("IDExpediente").value
                m_Nemotecnico = Nz(.Fields("Nemotecnico").value, "")
                m_CodExp = Nz(.Fields("CodExp").value, "")
                If m_Nemotecnico <> "" Then
                    m_sTextoExpediente = m_Nemotecnico
                Else
                    m_sTextoExpediente = m_CodExp
                End If
                If m_Col Is Nothing Then
                    Set m_Col = New Scripting.Dictionary
                    m_Col.CompareMode = TextCompare
                End If
                If Not m_Col.Exists(CStr(m_IDExpediente)) Then
                    m_Col.Add CStr(m_IDExpediente), m_IDExpediente & ";" & m_sTextoExpediente
                End If
                .MoveNext
            Loop
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_IDExpediente In m_Col
        
        p_Cbo.AddItem m_Col(m_IDExpediente)
        
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCombo ha devuelto el error: " & Err.Description
    End If
End Function
