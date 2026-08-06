Option Compare Database
Option Explicit
Private m_ObjUsuarioSeleccionado As usuario
Public Event UsuarioSeleccionado(usuario As usuario)
Public Event QuitarUsuarioSeleccionado()
Private m_Error As String



Private Sub ComandoLimpiarNombre_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.Nombre = Null
    Filtrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiarNombre_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    CorreoAlAdministrador m_Error
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function Filtrar( _
                            Optional p_SoloActivos As EnumSino = EnumSino.No, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim lst As ListBox
    Dim m_Nombre As String
    Dim m_ObjUsuario As usuario
    Dim m_ObjListaUsuarios As Scripting.Dictionary
    Dim m_UsuarioRed As Variant
    
    On Error GoTo errores
    
    If p_SoloActivos = 0 Then
        p_SoloActivos = EnumSino.No
    End If
    Me.ComandoResetearPass.Enabled = False
    Set m_ObjUsuarioActivo = Nothing
    Me.UsuarioSeleccionado = Null
    Set lst = Me.ListaUsuarios
    lst.RowSource = "ID;Correo;Nombre"
    m_Nombre = Nz(Me.Nombre, "")
    Set m_ObjListaUsuarios = Constructor.getListaUsuariosPorNombre(m_Nombre, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjListaUsuarios Is Nothing Then
        Exit Function
    End If
    For Each m_UsuarioRed In m_ObjListaUsuarios
        Set m_ObjUsuario = m_ObjListaUsuarios(m_UsuarioRed)
        If p_SoloActivos = EnumSino.Sí Then
            If m_ObjUsuario.Activado = False Then
                GoTo siguiente
            End If
        End If
        With m_ObjUsuario
            lst.AddItem .ID & ";" & .CorreoUsuario & ";" & .Nombre
        End With
        
        Set m_ObjUsuarioSeleccionado = m_ObjUsuario
siguiente:
    Next
    If lst.ListCount = 2 Then
        lst.Selected(1) = True
        ListaUsuarios_Click
        Set m_ObjUsuarioActivo = m_ObjUsuarioSeleccionado
        
        Me.UsuarioSeleccionado = m_ObjUsuarioActivo.Nombre
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub ComandoQuitarUsuarioActivo_Click()
    Me.UsuarioSeleccionado = Null
    Set m_ObjUsuarioActivo = Nothing
    RaiseEvent QuitarUsuarioSeleccionado
    Me.ComandoResetearPass.Enabled = False
End Sub

Private Sub ComandoResetearPass_Click()
    On Error GoTo errores
    
    If m_ObjUsuarioSeleccionado Is Nothing Then
        m_Error = "Seleccione un usuario de la lista"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ObjUsuarioSeleccionado.ResetearPass m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Pregunta = MsgBox("Se le ha mandado una clave de un solo uso al usuario por correo", vbInformation, "Resetear Pass")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoResetearPass_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    Me.ListaUsuarios.RowSource = "ID;Correo;Nombre"
    If Not m_ObjUsuarioActivo Is Nothing Then
        Me.UsuarioSeleccionado = m_ObjUsuarioActivo.Nombre
    End If
    Me.ComandoResetearPass.Enabled = False
End Sub

Private Sub ListaUsuarios_Click()
    Dim m_ID As String
    On Error GoTo errores
    m_Error = ""
    
    m_ID = Nz(Me.ListaUsuarios.Column(0), "")
    If Not IsNumeric(m_ID) Then
        Set m_ObjUsuarioSeleccionado = Nothing
        Set m_ObjUsuarioActivo = Nothing
        Exit Sub
    End If
    Set m_ObjUsuarioSeleccionado = Constructor.getUsuario(m_ID, , , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioActivo = m_ObjUsuarioSeleccionado
    Me.UsuarioSeleccionado = m_ObjUsuarioActivo.Nombre
    RaiseEvent UsuarioSeleccionado(m_ObjUsuarioActivo)
    Me.ComandoResetearPass.Enabled = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaUsuarios_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Nombre_AfterUpdate()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Filtrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Nombre_AfterUpdate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

