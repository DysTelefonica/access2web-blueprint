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
    Dim m_NombreControlUsuariosConectados As String
    Dim dato
    On Error GoTo errores
    
    
    
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Set m_ObjColAplicacionesUsuario = m_ObjEntorno.ColAplicacionesActivas
        
    Else
        Set m_ObjColAplicacionesUsuario = m_ObjUsuarioConectadoLogin.ColAplicacionesPermisos
        
    End If
    
    
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        'Debug.Print m_ID
        If Nz(m_ID, "") = "" Then
            GoTo siguiente
        End If
       'If CStr(m_ID) = "22" Then Stop
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
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
'        On Error Resume Next
'        dato=split(m_ObjBoton.Name,"Comando"
'        m_NombreControlUsuariosConectados=
'        me.Controls("UsuariosConectados"
        If Err.Number <> 0 Then
            Err.Raise 1000
        End If
        On Error GoTo errores
'
'        If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Or m_ObjUsuarioConectadoLogin.UsuarioRed = "mtr" Then
'            m_ObjBoton.Enabled = True
'            m_ObjBoton.Picture = m_URLAcceso
'            If m_objAplicacion.IDAplicacion = "51" Then
'                m_ObjBoton.Visible = True
'                Me.lblGestionRiesgosPrueba.Visible = True
'                Me.lblRiesgosPruebaVerServidor.Visible = True
'                Me.lblRiesgosPruebaVerUsuario.Visible = True
'
'            End If
'            GoTo siguiente
'        End If
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
    
    Me.AllowEdits = False
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
    
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoActualizarVersiones_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    setVersiones Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.AllowEdits = True
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
    Me.AllowEdits = True
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoExpediente_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoExpediente, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExpedientes_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoCondor_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoCondor, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCondor_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoSHPS_Click()
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoSHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNC_Click ha producido el error nº: " & Err.Number & _
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
    m_Error = ""
    
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

Private Sub ComandoAgedysTec_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoAgedysTec, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAgedysTec_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub


Private Sub ComandoAgedysEco_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoAgedysEco, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAgedysEco_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoRegistro_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoRegistro, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistro_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoGestionRiesgos_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoGestionRiesgos, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionRiesgos_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoNC_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoNC, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNC_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub ComandoBrass_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoBrass, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBrass_Click ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub



