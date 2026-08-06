Attribute VB_Name = "Form_FormInicial03ConsultasGrado"
Option Compare Database
Option Explicit
Private m_Error As String


Private Sub ComandoLimpiarEspecialidad_Click()
    Me.Especialidad = Null
    Especialidad_AfterUpdate
End Sub

Private Sub ComandoLimpiarEstadoHPS_Click()
    Me.ComboEstadoHPS = Null
    ComboEstadoHPS_AfterUpdate
End Sub
Private Sub ComandoLimpiarGradoESA_Click()
    Me.ComboGradoESA = Null
    ComboGradoESA_AfterUpdate
End Sub

Private Sub ComandoLimpiarGradoNacional_Click()
    Me.ComboGradoNacional = Null
    ComboGradoNacional_AfterUpdate
End Sub

Private Sub ComandoLimpiarGradoOTAN_Click()
    Me.ComboGradoOtan = Null
    ComboGradoOtan_AfterUpdate
End Sub

Private Sub ComandoLimpiarGradoUE_Click()
    Me.ComboGradoUE = Null
    ComboGradoUE_AfterUpdate
End Sub

Private Sub ComboGradoNacional_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboGradoNacional ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComboGradoOtan_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboGradoESA ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComboGradoUE_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboGradoESA ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ComboGradoESA_AfterUpdate()
    
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboGradoESA ha devuelto el error:" & vbNewLine & Err.Description
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
    m_ObjColParaFiltros("GradoNAC") = Nz(Me.ComboGradoNacional, "")
    m_ObjColParaFiltros("GradoOTAN") = Nz(Me.ComboGradoOtan, "")
    m_ObjColParaFiltros("GradoESA") = Nz(Me.ComboGradoESA, "")
    m_ObjColParaFiltros("GradoUE") = Nz(Me.ComboGradoUE, "")
    m_ObjColParaFiltros("Especialidad") = Nz(Me.Especialidad, "")
    m_ObjColParaFiltros("YO") = Nz(Me.ComboYO, "")
    
    m_ObjColParaFiltros("Estado") = Nz(Me.ComboEstadoHPS, "")
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
Private Sub ComboYO_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboYO ha devuelto el error:" & vbNewLine & Err.Description
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

Private Sub Especialidad_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Especialidad_AfterUpdate ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    Dim m_ObjColGrado As Scripting.Dictionary
    Dim m_Grado As Variant
    Dim cmb As ComboBox
    
    On Error GoTo errores
   
    Set m_ObjColParaFiltros = Nothing
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboEstado Me.ComboEstadoHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEntorno.colGradosHPS Is Nothing Then
        m_Error = m_ObjEntorno.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la lista de los distintos grados de HPS"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.colGradosHPS.Exists("Nacional") Then
        m_Error = "No se ha podido obtener la lista de los distintos grados de HPS Nacional"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.colGradosHPS.Exists("OTAN") Then
        m_Error = "No se ha podido obtener la lista de los distintos grados de HPS OTAN"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.colGradosHPS.Exists("UE") Then
        m_Error = "No se ha podido obtener la lista de los distintos grados de HPS UE"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.colGradosHPS.Exists("ESA") Then
        m_Error = "No se ha podido obtener la lista de los distintos grados de HPS ESA"
        Err.Raise 1000
    End If
    Set m_ObjColGrado = m_ObjEntorno.colGradosHPS("Nacional")
    Set cmb = Me.ComboGradoNacional
    cmb.RowSource = ""
    For Each m_Grado In m_ObjColGrado
        cmb.AddItem m_Grado
    Next
    Set m_ObjColGrado = m_ObjEntorno.colGradosHPS("OTAN")
    Set cmb = Me.ComboGradoOtan
    cmb.RowSource = ""
    For Each m_Grado In m_ObjColGrado
        cmb.AddItem m_Grado
    Next
    Set m_ObjColGrado = m_ObjEntorno.colGradosHPS("UE")
    Set cmb = Me.ComboGradoUE
    cmb.RowSource = ""
    For Each m_Grado In m_ObjColGrado
        cmb.AddItem m_Grado
    Next
    Set m_ObjColGrado = m_ObjEntorno.colGradosHPS("ESA")
    Set cmb = Me.ComboGradoESA
    cmb.RowSource = ""
    For Each m_Grado In m_ObjColGrado
        cmb.AddItem m_Grado
    Next
    Me.ComboYO = "O"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




