Attribute VB_Name = "Form_FormInicial07UsuariosSICA"
Option Compare Database
Option Explicit
Private m_Error As String





Public Sub ComandoEditarUsuarioSICA_Click()
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioSICAActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario SICA primero"
        Err.Raise 1000
        
    End If
    If FormularioAbierto("FormUsuario07SICAEdicion") Then
        DoCmd.Close acForm, "FormUsuario07SICAEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario07SICAEdicion"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El ComandoEditarUsuarioSICA_Click ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarUsuarioSICA_Click()
    Dim m_ID As String
    
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioSICAActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario SICA primero"
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioSICAActivo.ID
    pregunta = MsgBox("¿Desea eliminar definitivamente al usuario SICA seleccionado con todos sus Anexos asociados " & _
                    "?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Usuario SICA")
    
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjUsuarioSICAActivo.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    ActualizarUsuarioSICA m_ID, EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarUsuarioSICA_Clickse ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoLimpiarNombre_Click()
    Me.Nombre = Null
    Nombre_Exit False
    
End Sub

Private Sub Filtrar()
    
    On Error GoTo errores
    Set m_ObjColParaFiltros = New Scripting.Dictionary
    m_ObjColParaFiltros.CompareMode = TextCompare
    m_ObjColParaFiltros("Nombre") = Nz(Me.Nombre, "")
    
    EstablecerFiltroUsuarios m_Error
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

Public Function EstablecerFiltroUsuarios( _
                                            Optional ByRef p_Error As String _
                                        ) As String
    
    
    
    Dim m_Where As String
    Dim m_SQL As String
    Dim m_Nombre As String
    Dim frm As Form
    On Error GoTo errores
    
    p_Error = ""
    
    Set frm = Me.Controls("SubFormResultados").Form
    m_Nombre = Nz(Me.Nombre, "")
    If m_Nombre = "" Then
        m_Where = ";"
    Else
        m_Where = "WHERE ([TbUsuariosSICALocal].[Nombre] & ' ' & [TbUsuariosSICALocal].[Apellido_1] & ' ' & [TbUsuariosSICALocal].[Apellido_2]) Like '*" & m_Nombre & "*';"
    End If
    m_SQL = "SELECT TbUsuariosSICALocal.* " & _
                    "FROM TbUsuariosSICALocal " & _
                    m_Where
    
    frm.RecordSource = m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerFiltroUsuarios ha devuelto el error: " & Err.Description
    End If
    
End Function

Private Sub ComandoListaEnExcel_Click()
    Dim m_URLExcelGenerado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_URLExcelGenerado = ConsultaAExcel(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)
   
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_URLExcelGenerado, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Obtener Listado en Excel ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoNuevoUsuarioSICA_Click()
    On Error GoTo errores
    
    Set m_ObjUsuarioSICAActivo = Nothing
    If FormularioAbierto("FormUsuario07SICAAlta") Then
        DoCmd.Close acForm, "FormUsuario07SICAAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario07SICAAlta"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevoUsuarioSICA_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Load()
    On Error GoTo errores
    
    m_Error = ""
   
    CopiarEstiloBoton Me.ComandoNuevoUsuarioSICA, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoEditarUsuarioSICA, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    EstablecerFiltroUsuarios m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoNuevoUsuarioSICA.Enabled = False
        Me.ComandoEliminarUsuarioSICA.Enabled = False
    Else
        Me.ComandoNuevoUsuarioSICA.Enabled = True
        Me.ComandoEliminarUsuarioSICA.Enabled = True
    End If
    Me.ComandoEditarUsuarioSICA.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar abrir el formulario se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub Nombre_Exit(Cancel As Integer)
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Nombre_Exit ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
