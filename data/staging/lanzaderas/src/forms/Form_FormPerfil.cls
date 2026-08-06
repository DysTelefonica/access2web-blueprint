Option Compare Database
Option Explicit
Private m_Error As String
Private m_NombreFormulario As String
Private m_ObjUsuario As usuario


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim lst As ListBox
    On Error GoTo errores
    
    
    If m_ObjUsuarioConectadoLogin Is Nothing Then
        p_Error = "No se sabe quién ha hecho login"
        Err.Raise 1000
    End If
    Set m_ObjUsuario = m_ObjUsuarioConectadoLogin
    With m_ObjUsuario
        If .Nombre <> "" Then
            Me.Nombre = .Nombre
        End If
        If .CorreoUsuario <> "" Then
            Me.CorreoUsuario = .CorreoUsuario
        End If
        If .UsuarioRed <> "" Then
            Me.UsuarioRed = .UsuarioRed
        End If
        If IsDate(.FechaAlta) Then
            Me.FechaAlta = Format(.FechaAlta, "dd/mm/yyyy")
        End If
        If .Matricula <> "" Then
            Me.Matricula = .Matricula
        End If
        If .Telefono <> "" Then
            Me.Telefono = .Telefono
        End If
        If .Movil <> "" Then
            Me.Movil = .Movil
        End If
        If .TextoTiempoParaExpirarPass <> "" Then
            Me.TextoTiempoParaExpirarPass = .TextoTiempoParaExpirarPass
        End If
        If .MantenerLanzaderaAbierta = True Then
            Me.MantenerLanzaderaAbierta = "Sí"
        Else
            Me.MantenerLanzaderaAbierta = "No"
        End If
        If IsDate(.FechaBaja) Then
            Me.FechaBaja = Format(.FechaBaja, "dd/mm/yyyy")
        End If
        If Not .JefeUsuarioObj Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Me.Responsable = .JefeUsuarioObj.Nombre
        End If
       
        Me.ID = .ID
        
    End With
    Me.AllowEdits = True
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            If InStr(1, ctl.Tag, "DATO") <> 0 Then
                ctl.Enabled = True
                ctl.Locked = True
            End If
        End If
    Next
    Set lst = Me.ListaUsuarios
    lst.Enabled = True
    EstablecerLista p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_ID As Variant
    Dim m_IDApp As String
    Dim m_ObjCol As Scripting.Dictionary
    Dim m_objAplicacion As Aplicacion
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    Dim Admin As String
    Dim Calidad As String
    Dim CalidadAvisos As String
    Dim Secretaria As String
    Dim Tecnico As String
    Dim Economia As String
    Dim PartePermisos As String
    On Error GoTo errores
    
    p_Error = ""
    Set lst = Me.ListaUsuarios
    
    lst.RowSource = "IDAplicacion;Aplicación;Admin;Calidad;Economia;Técnico;Secretaría;Cal Avisos"
    If m_ObjUsuario Is Nothing Then
        p_Error = "No se conoce el usuario"
        Err.Raise 1000
    End If
    
    Set m_ObjCol = m_ObjUsuario.ColAplicacionesPermisos
    p_Error = m_ObjUsuario.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjCol Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjCol
        Set m_ObjUsuarioAplicacionPermisos = m_ObjCol(m_ID)
        With m_ObjUsuarioAplicacionPermisos
            If .EsUsuarioAdministradorCalculado = EnumSino.Sí Then
                Admin = "Sí"
            Else
                Admin = "No"
            End If
            If .EsUsuarioTecnicoCalculado = EnumSino.Sí Then
                Tecnico = "Sí"
            Else
                Tecnico = "No"
            End If
            If .EsUsuarioCalidadAvisosCalculado = EnumSino.Sí Then
                CalidadAvisos = "Sí"
            Else
                CalidadAvisos = "No"
            End If
            If .EsUsuarioCalidadCalculado = EnumSino.Sí Then
                Calidad = "Sí"
            Else
                Calidad = "No"
            End If
            If .EsUsuarioEconomiaCalculado = EnumSino.Sí Then
                Economia = "Sí"
            Else
                Economia = "No"
            End If
            If .EsUsuarioSecretariaCalculado = EnumSino.Sí Then
                Secretaria = "Sí"
            Else
                Secretaria = "No"
            End If
            'lst.RowSource = "IDAplicacion;Aplicación;Admin;Calidad;Economia;Técnico;Secretaría;Cal Avisos"
            PartePermisos = Admin & ";" & Calidad & ";" & Economia & ";" & Tecnico & ";" & Secretaria & ";" & CalidadAvisos
            lst.AddItem .IDAplicacion & ";" & .Aplicacion.NombreAplicacion & ";" & PartePermisos
        End With
        Set m_ObjUsuarioAplicacionPermisos = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    m_NombreFormulario = Me.Name
    
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub
Private Sub ComandoCambiarPass_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormUsuarioCambioPass") Then
        DoCmd.Close acForm, "FormUsuarioCambioPass", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuarioCambioPass"
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCambiarPass_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

