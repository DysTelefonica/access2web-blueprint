Attribute VB_Name = "Form_FormInicial02ConsultasPrincipal"
Option Compare Database
Option Explicit

Private m_Error As String


Private m_IDListaSeleccionada As String
Private m_IDListaSeleccionadaUltima As String
Private m_SQLSinWhere As String
Private WithEvents m_FormInicial00Principal As Form_FormInicial00Principal
Private WithEvents m_FormInicial03ConsultasDatos As Form_FormInicial03ConsultasDatos

Public Sub ComandoAnexos_Click()
    
    
    On Error GoTo errores
    m_Error = ""
    

    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "Ha de seleccionar un usuario de la lista"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormUsuario00Principal") Then
        DoCmd.Close acForm, "FormUsuario00Principal", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario00Principal", OpenArgs:="No"
    
    Form_FormUsuario01Lateral.lblAnexos_Click
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar ver el anexo ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Public Sub ComandoEditarUsuario_Click()
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario primero"
        Err.Raise 1000
        
    End If
    If FormularioAbierto("FormUsuario00Principal") Then
        DoCmd.Close acForm, "FormUsuario00Principal", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario00Principal"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar ver el detalle ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarUsuario_Click()
    
    
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "Se ha de seleccionar el usuario primero"
        Err.Raise 1000
    End If
    
    pregunta = MsgBox("¿Desea eliminar definitivamente al usuario seleccionado con todos sus restros asociados " & _
                    "(Observaciones,HPSs,Anexos)?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Usuario")
    
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjUsuarioActivo.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    If m_ObjEntorno Is Nothing Then
        EVE m_Error
        If m_Error <> "" Then Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarUsuario_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoListaEnExcel_Click()
    
    Dim m_URLExcelGenerado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_URLExcelGenerado = ConsultaAExcel(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)
    'm_URLExcelGenerado = GenerarConsultaUsuarioOrdenada(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)
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

Private Sub ComandoNuevoUsuario_Click()
    
    On Error GoTo errores
    
    Set m_ObjUsuarioActivo = Nothing
    If FormularioAbierto("FormUsuario02Alta") Then
        DoCmd.Close acForm, "FormUsuario02Alta", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario02Alta"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevoUsuario_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    
    On Error GoTo errores
    
    m_Error = ""
    Me.AllowEdits = False
    Me.SubFormParametrosConsulta.SourceObject = ""
    EstablecerTablasConsultas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoNuevoUsuario.Enabled = False
        Me.ComandoEditarUsuario.Caption = "Detalle Usuario"
        Me.ComandoEliminarUsuario.Enabled = False
    Else
        Me.ComandoNuevoUsuario.Enabled = True
        Me.ComandoEditarUsuario.Caption = "Editar Usuario"
        Me.ComandoEliminarUsuario.Enabled = True
    End If
    CopiarEstiloBoton Me.ComandoNuevoUsuario, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoEditarUsuario, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoAnexos, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    On Error Resume Next
    If FormularioAbierto("FormInicial00Principal") Then
        Set m_FormInicial00Principal = Forms("FormInicial00Principal")
    End If
    Err.Clear
    On Error GoTo errores
    
    Me.ListaConsultas.Selected(0) = True
    ListaConsultas_Click
    
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método abrir formulario ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Sub ListaConsultas_Click()

    Dim m_SubFormParametrosConsultaNombreForm As String
    On Error GoTo errores
    
    m_Error = ""
    m_IDListaSeleccionada = Nz(Me.ListaConsultas.Column(0), "")
    If Not IsNumeric(m_IDListaSeleccionada) Then
        m_IDListaSeleccionadaUltima = m_IDListaSeleccionada
        Me.SubFormParametrosConsulta.SourceObject = ""
        Me.SubFormResultados.SourceObject = ""
        If Not m_ObjUsuarioActivo Is Nothing Then
            Set m_ObjUsuarioActivo = Nothing
        End If
        Me.ComandoListaEnExcel.Enabled = False
        Me.ComandoAnexos.Enabled = False
        Me.ComandoEditarUsuario.Enabled = False
        Exit Sub
    End If
    If m_IDListaSeleccionada = m_IDListaSeleccionadaUltima Then
        Exit Sub
    End If
    Me.ComandoEditarUsuario.Enabled = True
    'reiniciamos el objeto usuario seleccionado
    If Not m_ObjUsuarioActivo Is Nothing Then
        Set m_ObjUsuarioActivo = Nothing
    End If
    Me.ComandoAnexos.Enabled = True
    m_IDListaSeleccionadaUltima = m_IDListaSeleccionada
    Me.SubFormResultados.SourceObject = "FormInicial03ConsultasDatos"
    Set m_FormInicial03ConsultasDatos = Me.SubFormResultados.Form
    
    If m_ObjEntorno.ColTiposConsultasFormulario Is Nothing Then
        m_Error = "No se han encontrado datos de la consulta nº " & m_IDListaSeleccionada & " en la colección Entorno.ColTiposConsultasFormulario"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno.ColTiposConsultasFormulario.Exists(CStr(CLng(m_IDListaSeleccionada))) Then
        m_Error = "No se han encontrado datos de la consulta nº " & m_IDListaSeleccionada & " en la colección Entorno.ColTiposConsultasFormulario"
        Err.Raise 1000
    End If
    m_SubFormParametrosConsultaNombreForm = m_ObjEntorno.ColTiposConsultasFormulario(CStr(CLng(m_IDListaSeleccionada)))
    m_Error = m_ObjEntorno.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SubFormParametrosConsultaNombreForm = "" Then
        m_Error = "No se han encontrado datos de la consulta nº " & m_IDListaSeleccionada & " en la colección Entorno.ColTiposConsultasFormulario"
        Err.Raise 1000
    End If
    Me.SubFormParametrosConsulta.SourceObject = m_SubFormParametrosConsultaNombreForm
    Me.ComandoListaEnExcel.Enabled = True
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método click ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    m_IDListaSeleccionadaUltima = 0
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Function EstablecerTablasConsultas(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_ObjConsulta As Consulta
    Dim lst As ListBox
    On Error GoTo errores
    p_Error = ""
    Set lst = Me.ListaConsultas
    lst.RowSource = ""
    If Not m_ObjEntorno.ColConsultas Is Nothing Then
        For Each m_ID In m_ObjEntorno.ColConsultas.Keys
            If Nz(m_ID, "") <> "" Then
                Set m_ObjConsulta = m_ObjEntorno.ColConsultas(m_ID)
                lst.AddItem Format(m_ID, "00") & ";" & m_ObjConsulta.Nombre
                Set m_ObjConsulta = Nothing
            End If
        Next
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Ver EstablecerTablasConsultas ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Private Sub m_FormInicial00Principal_ActualizadoUsuarios()
    Me.SubFormResultados.Form.Requery
    Me.ListaConsultas.Selected(0) = True
    ListaConsultas_Click
End Sub