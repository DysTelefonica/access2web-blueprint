Option Compare Database
Option Explicit
Private m_Error As String


Public Function EstablecerBotonera(Optional ByRef p_Error As String) As String
    
    
    Dim m_objAplicacion As Aplicacion
    Dim m_ObjUsuarioAplicacionPermiso As UsuarioAplicacionPermisos
    Dim m_ObjColAplicacionesUsuario As Scripting.Dictionary
    Dim m_IDAplicacion As String
    Dim m_ID As Variant
    Dim m_ObjBoton As CommandButton
    Dim m_URLAcceso As String
    Dim m_URLNoAcceso As String
   
    Dim m_AplicacionEnOficina As EnumSino
    
    Dim m_ColAplicaciones As Scripting.Dictionary
    On Error GoTo errores
    
   
    
    Set m_ObjColAplicacionesUsuario = m_ObjUsuarioConectadoLogin.ColAplicacionesPermisos
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Set m_ObjColAplicacionesUsuario = m_ObjEntorno.ColAplicacionesActivas
        
    Else
        Set m_ObjColAplicacionesUsuario = m_ObjUsuarioConectadoLogin.ColAplicacionesPermisos
        
    End If
    
    Set m_ColAplicaciones = m_ObjEntorno.ColAplicacionesActivas
    For Each m_ID In m_ColAplicaciones
        If Nz(m_ID, "") = "" Then
            GoTo siguiente
        End If
        'If Nz(m_ID, "") = "20" Then Stop
        Set m_objAplicacion = m_ColAplicaciones(m_ID)
        If m_objAplicacion.ConIconoEnLanzaderaCalculado = EnumSino.No Then
            GoTo siguiente
        End If
        m_AplicacionEnOficina = m_objAplicacion.EjecucionEnOficinaCalculado
        Set m_ObjBoton = getBoton(Me, CStr(m_ID), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjBoton Is Nothing Then
            GoTo siguiente
        End If
        m_URLAcceso = m_objAplicacion.URLIconoAcceso
        m_URLNoAcceso = m_objAplicacion.URLIconoNOAcceso
        If m_AplicacionEnOficina = EnumSino.Sí Then
            If m_EnOficina = EnumSino.No Then
                m_ObjBoton.Enabled = False
                m_ObjBoton.Picture = m_URLNoAcceso
                GoTo siguiente
            End If
        End If
        
        If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
            If Nz(m_ID, "") = "20" Then
                Me.lblSICA.Visible = True
                Me.lblSICAVerServidor.Visible = True
                Me.lblSICAVerUsuario.Visible = True
                Me.ComandoSICA.Visible = True
            End If
            m_ObjBoton.Enabled = True
            m_ObjBoton.Picture = m_URLAcceso
            
            GoTo siguiente
        Else
            If Nz(m_ID, "") = "20" Then
                Me.lblSICA.Visible = False
                Me.lblSICAVerServidor.Visible = False
                Me.lblSICAVerUsuario.Visible = False
                Me.ComandoSICA.Visible = False
            End If
        End If
        If m_ObjColAplicacionesUsuario Is Nothing Then
            m_ObjBoton.Enabled = False
            m_ObjBoton.Picture = m_URLNoAcceso
            GoTo siguiente
        End If
        If m_ObjColAplicacionesUsuario.Exists(m_ID) Then
            m_ObjBoton.Enabled = True
            m_ObjBoton.Picture = m_URLAcceso
        Else
            m_ObjBoton.Enabled = False
            m_ObjBoton.Picture = m_URLNoAcceso
        End If
        If m_objAplicacion.NombreAplicacion = "EXPEDIENTES" Then
            m_ObjBoton.Enabled = True
            m_ObjBoton.Picture = m_URLAcceso
        End If
siguiente:
        Set m_objAplicacion = Nothing
    Next
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        ActualizarAperturas p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    EstablecerBotonera = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerBotonera ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Private Sub ComandoActualizar_Click()
    Dim m_ObjColAplicacionesUsuario As Scripting.Dictionary
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    
    EstablecerBotonera m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Set m_ObjColAplicacionesUsuario = m_ObjEntorno.ColAplicacionesActivas
    Else
        Set m_ObjColAplicacionesUsuario = m_ObjUsuarioConectadoLogin.ColAplicacionesPermisos
    End If
    LimpiarCache m_ObjColAplicacionesUsuario, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    setVersiones Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoActualizarVersiones_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    setVersiones Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarVersiones_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoAGEDO_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoAGEDO, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAGEDO_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoGestionContratos_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoGestionContratos, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionContratos_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoHPS_Click()
    
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoHPS_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoSeguridad_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoSeguridad, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSeguridad_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoSICA_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoSICA, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSICA_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub Form_Load()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    If m_ObjUsuarioConectadoLogin Is Nothing Then
        m_Error = "No se puede determinar el usuario que ha abierto la aplicación"
        Err.Raise 1000
    End If
    EstablecerBotonera m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    setVersiones Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

