Attribute VB_Name = "Form_FormInicial08UsuariosHistoricos"
Option Compare Database
Option Explicit
Private m_Error As String

Private m_SQLAlInicio As String



Public Sub ComandoVerUsuarioHistorico_Click()
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario primero"
        Err.Raise 1000
        
    End If
    If FormularioAbierto("FormUsuarioHistorico00Principal") Then
        DoCmd.Close acForm, "FormUsuarioHistorico00Principal", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuarioHistorico00Principal"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerUsuarioHistorico_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarUsuarioHistorico_Click()
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario Histórico primero"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar definitivamente al usuario Histórico seleccionado " & _
                    "?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Usuario Histórico")
    
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjUsuarioHistoricoActivo.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    
    If m_ObjEntorno Is Nothing Then
        EVE m_Error
        If m_Error <> "" Then Err.Raise 1000
    End If
    ' EVE ' <-- Si EVE no es estrictamente necesario para recargar variables globales aquí, puedes comentarlo para ganar velocidad
    InicializarTablasLocalesYEntidad_Optimizado p_ForzarRegeneracionTotal:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ActualizarListaDatosLocal
    EstablecerContadores
    
    
    Me.SubFormResultados.Form.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar ComandoEliminarUsuarioHistorico_Clickse ha dado el error: " & vbNewLine & Err.Description
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
    
    EstablecerFiltroUsuariosHistoricos m_Error
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

Public Function EstablecerFiltroUsuariosHistoricos( _
                                                        Optional ByRef p_Error As String _
                                                    ) As String
    
    
    
    Dim m_Where As String
    Dim m_SQL As String
    Dim m_SQLInicial As String
    Dim m_Nombre As String
    Dim frm As Form
    On Error GoTo errores
    
    p_Error = ""
    m_SQLInicial = "SELECT * " & _
                    "FROM TbUsuariosHistoricosLocal "
    
    Me.SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"
    Set frm = Me.Controls("SubFormResultados").Form
    m_Nombre = Nz(Me.Nombre, "")
    If m_Nombre = "" Then
        m_Where = ";"
    Else
        m_Where = "WHERE Nombre & ' ' & Apellido_1 & ' ' & Apellido_2 Like '*" & m_Nombre & "*';"
    End If
    
    m_SQL = m_SQLInicial & _
            m_Where
    
    frm.RecordSource = m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerFiltroUsuariosHistoricos ha devuelto el error: " & Err.Description
    End If
    
End Function

Private Sub ComandoListaEnExcel_Click()
    Dim m_URLExcelGenerado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"
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



Private Sub Form_Load()
    
    On Error GoTo errores
    
    m_Error = ""
   
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoVerUsuarioHistorico, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoEliminarUsuarioHistorico, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar abrir el formulario se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Public Sub Nombre_Exit(Cancel As Integer)
    On Error GoTo errores
    m_Error = ""
    EstablecerFiltroUsuariosHistoricos m_Error
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
