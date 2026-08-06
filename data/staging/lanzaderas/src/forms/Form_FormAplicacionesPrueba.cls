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
        'If Nz(m_ID, "") = "231" Then Stop
        Set m_objAplicacion = m_ColAplicaciones(m_ID)
        If m_objAplicacion.ConIconoEnLanzaderaCalculado = EnumSino.No Then
            GoTo siguiente
        End If
        Set m_ObjBoton = getBoton(Me, CStr(m_ID), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjBoton Is Nothing Then
            GoTo siguiente
        End If
        m_URLAcceso = m_objAplicacion.URLIconoAcceso
        m_URLNoAcceso = m_objAplicacion.URLIconoNOAcceso
        If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
            m_ObjBoton.Enabled = True
            m_ObjBoton.Picture = m_URLAcceso
            GoTo siguiente
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
        
siguiente:
        Set m_objAplicacion = Nothing
    Next
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        ActualizarAperturas p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Me.ComandoCondorPrueba.Enabled = False
    
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
    setVersionAplicacionesEnPrueba m_Error
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

Private Sub ComandoCondorPrueba_Click()
'    On Error GoTo errores
'
'    m_Error = ""
'    VBA.DoEvents
'    DoCmd.Hourglass True
'    VBA.DoEvents
'    Lanzar Me.ComandoCondorPrueba, m_Error
'    If m_Error <> "" Then
'        Err.Raise 1000
'    End If
'
'    VBA.DoEvents
'    DoCmd.Hourglass False
'    VBA.DoEvents
'    Exit Sub
'errores:
'    If Err.Number <> 1000 Then
'        m_Error = "Al ComandoCondorPrueba_Click ha producido el error nº: " & Err.Number & _
'                    vbCrLf & "Detalle: " & Err.Description
'        CorreoAlAdministrador m_Error
'    End If
'    DoCmd.Hourglass False
'
'    Pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub


Private Sub ComandoNoConformidadesPrueba_Click()
    On Error GoTo errores

    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Lanzar Me.ComandoNoConformidadesPrueba, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNoConformidadesPrueba_Click ha producido el error nº: " & Err.Number & _
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
    setVersionAplicacionesEnPrueba m_Error
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

Public Function setVersionAplicacionesEnPrueba( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    
    Dim m_URLIniLocal As String
    Dim m_URLIniRemoto As String
    Dim m_VersionLocal As String
    Dim m_VersionRemota As String
    Dim m_URLDirectorioAplicacionUsuario As String
    Dim m_URLDirectorioRemoto  As String
    Dim m_TituloServidor As String
    Dim m_TituloLocal As String
    
    Dim m_Titulo As String
    
    
    On Error GoTo errores
  
    
    
    
   
    m_URLDirectorioAplicacionUsuario = Environ("APPDATA") & "\" & "Aplicaciones DYSN\CONDOR Prueba\"
    m_URLDirectorioRemoto = NormalizarRuta(m_URLRutaAplicacionesRemotas) & "CONDOR Prueba\recursos\"
    m_URLIniLocal = m_URLDirectorioAplicacionUsuario & "Condor.ini"
    m_URLIniRemoto = m_URLDirectorioRemoto & "Condor.ini"


    m_VersionLocal = getVersion(m_URLIniLocal, p_Error)
    m_VersionRemota = getVersion(m_URLIniRemoto, p_Error)

    Me.lblCondorPruebaVerServidor.Caption = "Ver. Servidor: " & m_VersionRemota
    Me.lblCodorPruebaVerUsuario.Caption = "Ver. Usuario: " & m_VersionLocal
    
    
    
    m_URLDirectorioAplicacionUsuario = Environ("APPDATA") & "\" & "Aplicaciones DYSN\No Conformidades Prueba\"
    m_URLDirectorioRemoto = NormalizarRuta(m_URLRutaAplicacionesRemotas) & "No Conformidades Prueba\recursos\"
    m_URLIniLocal = m_URLDirectorioAplicacionUsuario & "NoConformidades.ini"
    m_URLIniRemoto = m_URLDirectorioRemoto & "NoConformidades.ini"


    m_VersionLocal = getVersion(m_URLIniLocal, p_Error)
    m_VersionRemota = getVersion(m_URLIniRemoto, p_Error)

    Me.lblNoConformidadesPruebaVerServidor.Caption = "Ver. Servidor: " & m_VersionRemota
    Me.lblNoConformidadesPruebaVerUsuario.Caption = "Ver. Usuario: " & m_VersionLocal
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método setVersionAplicacionesEnPrueba ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
