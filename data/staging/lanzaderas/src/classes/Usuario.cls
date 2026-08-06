VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public CorreoUsuario As String
Public Password As String
Public UsuarioRed As String
Public Nombre As String
Public Matricula As String
Public FechaAlta As String
Public Activado As Boolean
Public FechaProximoCambioContrasenia As String
Public TieneQueCambiarLaPass As Boolean
Public Telefono As String
Public Movil As String
Public Observaciones As String
Public UsuarioImborrable As Boolean
Public EsAdministrador As String
Public PermisosAsignados As Boolean
Public FechaBaja As String
Public PasswordNuncaCaduca As Boolean
Public MantenerLanzaderaAbierta As Boolean
Public PassIncialPlana As String
Public UsuarioSSID As String
Public ID As String
Public JefeDelUsuario As String
Public PermisoPruebas As String
Public ParaTareasProgramadas As Boolean
Public FechaBloqueo As String
Public FechaUltimaConexion As String

Private m_sPassIncialPlanaCalculada As String
Private m_eEsAdministradorCalculado As EnumSino
Private m_eEsUsuarioCalidadCalculado As EnumSino
Private m_eEsUsuarioTecnicoCalculado As EnumSino
Private m_objJefeUsuarioObj As usuario
Private m_objColAplicacionesPermisos As Scripting.Dictionary
Private m_sFechaProximaPassCalculado As String
Private m_objColConexiones As Scripting.Dictionary
Private m_objUltimaConexion As Conexion
Private m_objUltimoCambioPass As HistoricoPass
Private m_objColListaCambiosPasss As Scripting.Dictionary
Private m_eCookieActiva As EnumSino
Private m_ePassCaducada As EnumSino
Private m_eTieneQueCambiarLaPassCalculado As EnumSino
Private m_sIDCalculado As String
Private m_sTextoTiempoParaExpirarPass As String
Private m_eParaAlta As EnumSino
Private m_objUltimaAplicacionAbierta As Aplicacion
Private m_sEtiquetaAplicacionesAPulsar As String
Private m_objConexionActual As Conexion
Private m_ObjPermisos As UsuarioAplicacionPermisos
Private m_sEquipo As String
Private m_objComprobacionAccesoVideos As ComprobacionAccesoVideos
Public Error As String
Private m_objColCampos As Collection


Public Property Get EsAdministradorCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EsAdministrador = "Sí" Then
        m_eEsAdministradorCalculado = EnumSino.Sí
    ElseIf Me.EsAdministrador = "No" Then
        m_eEsAdministradorCalculado = EnumSino.No
    End If
    EsAdministradorCalculado = m_eEsAdministradorCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.EsAdministradorCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    

End Property

Public Property Get JefeUsuarioObj() As usuario

    On Error GoTo errores
    Me.Error = ""
    If Not m_objJefeUsuarioObj Is Nothing Then
        Set JefeUsuarioObj = m_objJefeUsuarioObj
        Exit Property
    End If
    If Me.JefeDelUsuario = "" Then
        Exit Property
    End If
    Set m_objJefeUsuarioObj = Constructor.getUsuario(, Me.JefeDelUsuario, , , Me.Error)
    Set JefeUsuarioObj = m_objJefeUsuarioObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.JefeUsuarioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get ColAplicacionesPermisos() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColAplicacionesPermisos Is Nothing Then
        Set ColAplicacionesPermisos = m_objColAplicacionesPermisos
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objColAplicacionesPermisos = Constructor.getAplicacionesPermisos(Me.CorreoUsuario, Me.Error)
    
    Set ColAplicacionesPermisos = m_objColAplicacionesPermisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.colAplicacionesPermisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property
Public Property Set ColAplicacionesPermisos(ByVal Valor As Variant)

    Set m_objColAplicacionesPermisos = Valor
End Property
Public Property Get FechaProximaPassCalculado() As String
    
    Dim m_FechaUltimoCambio As String
    On Error GoTo errores
    Me.Error = ""
    If Me.PasswordNuncaCaduca = True Then
        Exit Property
    End If
    If Me.UltimoCambioPass Is Nothing Then
        Exit Property
    End If
    m_FechaUltimoCambio = Me.UltimoCambioPass.FechaPass
    
    If Not IsDate(m_FechaUltimoCambio) Then
        Exit Property
    End If
    m_sFechaProximaPassCalculado = DateAdd("d", 90, CDate(m_FechaUltimoCambio))
    
    FechaProximaPassCalculado = m_sFechaProximaPassCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.FechaProximaPassCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    

End Property

Public Property Get ColConexiones() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColConexiones Is Nothing Then
        Set ColConexiones = m_objColConexiones
        Exit Property
    End If
    If Me.UsuarioRed = "" Then
        Exit Property
    End If
    Set m_objColConexiones = Constructor.getConexiones(Me.UsuarioRed, Me.Error)
    Set ColConexiones = m_objColConexiones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.ColConexiones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    

End Property

Public Property Get UltimaConexion() As Conexion

    On Error GoTo errores
    Me.Error = ""
    If Not m_objUltimaConexion Is Nothing Then
        Set UltimaConexion = m_objUltimaConexion
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objUltimaConexion = Constructor.getUltimaConexion(Me.CorreoUsuario, Me.Error)
    Set UltimaConexion = m_objUltimaConexion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.UltimaConexion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get ConexionActual() As Conexion

    On Error GoTo errores
    Me.Error = ""
    If Not m_objConexionActual Is Nothing Then
        Set ConexionActual = m_objConexionActual
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objConexionActual = Constructor.getConexion(Me.CorreoUsuario, , Me.Error)
    Set ConexionActual = m_objConexionActual
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.ConexionActual ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get UltimoCambioPass() As HistoricoPass

    On Error GoTo errores
    Me.Error = ""
    If Not m_objUltimoCambioPass Is Nothing Then
        Set UltimoCambioPass = m_objUltimoCambioPass
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objUltimoCambioPass = Constructor.getUltimoCambioPass(Me.CorreoUsuario, Me.Error)
    Set UltimoCambioPass = m_objUltimoCambioPass
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.UltimoCambioPass ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get ColListaCambiosPasss() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColListaCambiosPasss Is Nothing Then
        Set ColListaCambiosPasss = m_objColListaCambiosPasss
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objColListaCambiosPasss = Constructor.getCambiosPasss(Me.CorreoUsuario, Me.Error)
    Set ColListaCambiosPasss = m_objColListaCambiosPasss
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.ColListaCambiosPasss ha devuelto el error: " & vbNewLine & Err.Description
    End If
    

End Property

Public Property Get TextoTiempoParaExpirarPass() As String

    On Error GoTo errores
    Me.Error = ""
    m_sTextoTiempoParaExpirarPass = getTextoTiempoParaExpirarPass(Me.Error)
    TextoTiempoParaExpirarPass = m_sTextoTiempoParaExpirarPass
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.TextoTiempoParaExpirarPass ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get IDCalculado() As String

    On Error GoTo errores
    Me.Error = ""
    m_sIDCalculado = DameID("TbUsuariosAplicaciones", "ID", getdb(), Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    IDCalculado = m_sIDCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.IDCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get CookieActiva() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    If m_eCookieActiva <> 0 Then
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    If Me.TieneQueCambiarLaPassCalculado = EnumSino.Sí Then
        m_eCookieActiva = EnumSino.No
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    If Me.UltimaConexion Is Nothing Then
        If Me.Error <> "" Then
            Exit Property
        End If
        m_eCookieActiva = EnumSino.No
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    If Year(Me.UltimaConexion.FechaConexion) <> Year(Date) Then
        m_eCookieActiva = EnumSino.No
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    If Month(Me.UltimaConexion.FechaConexion) <> Month(Date) Then
        m_eCookieActiva = EnumSino.No
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    If Day(Me.UltimaConexion.FechaConexion) <> Day(Date) Then
        m_eCookieActiva = EnumSino.No
        CookieActiva = m_eCookieActiva
        Exit Property
    End If
    
    m_eCookieActiva = EnumSino.Sí
    CookieActiva = m_eCookieActiva
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.CookieActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property
Public Property Get PassCaducada() As EnumSino
    
    On Error GoTo errores
    Me.Error = ""
    If m_ePassCaducada <> 0 Then
        PassCaducada = m_ePassCaducada
        Exit Property
    End If
    If Me.PasswordNuncaCaduca = True Then
         m_ePassCaducada = EnumSino.No
         PassCaducada = m_ePassCaducada
        Exit Property
    End If
    If Not IsDate(Me.FechaProximaPassCalculado) Then
        m_ePassCaducada = EnumSino.Sí
        PassCaducada = m_ePassCaducada
        Exit Property
    End If
    If CDate(Date) > CDate(Format(Me.FechaProximaPassCalculado, "dd/mm/yyyy")) Then
        m_ePassCaducada = EnumSino.Sí
    Else
        m_ePassCaducada = EnumSino.No
    End If
    
    PassCaducada = m_ePassCaducada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.PassCaducada ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get TieneQueCambiarLaPassCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.TieneQueCambiarLaPass Then
        m_eTieneQueCambiarLaPassCalculado = EnumSino.Sí
    Else
        m_eTieneQueCambiarLaPassCalculado = EnumSino.No
    End If
    TieneQueCambiarLaPassCalculado = m_eTieneQueCambiarLaPassCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.TieneQueCambiarLaPassCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

'Public Property Get ComprobacionAccesoVideos() As ComprobacionAccesoVideos
'
'    On Error GoTo errores
'    Me.Error = ""
'    If Not m_objComprobacionAccesoVideos Is Nothing Then
'        Set ComprobacionAccesoVideos = m_objComprobacionAccesoVideos
'        Exit Property
'    End If
'    If Me.UsuarioRed = "" Then
'        Exit Property
'    End If
'    Set m_objComprobacionAccesoVideos = Constructor.getComprobacionAccesoVideos(Me.UsuarioRed, Me.Error)
'    Set ComprobacionAccesoVideos = m_objComprobacionAccesoVideos
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El método Usuario.ComprobacionAccesoVideos ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'
'End Property



Public Function PassNoRepetida( _
                                Optional p_pass As String, _
                                Optional p_PassEncriptada As String, _
                                Optional p_NumeroVeces As Long, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino

    Dim m_ID As Variant
    Dim i As Long
    Dim m_ObjHistorico As HistoricoPass
    
    On Error GoTo errores
    If p_pass = "" And p_PassEncriptada = "" Then
        p_Error = "Se ha de indicar la pass"
        Err.Raise 1000
    End If
    If p_PassEncriptada = "" Then
        p_PassEncriptada = SHA256(p_pass)
    End If
    If p_PassEncriptada = "" Then
        p_Error = "No se ha podido obtener la Password encriptada"
        Err.Raise 1000
    End If
    If Me.ColListaCambiosPasss Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        PassNoRepetida = EnumSino.No
        Exit Function
    End If
    If p_NumeroVeces = 0 Then
        p_NumeroVeces = 10
    End If
    i = 1
    For Each m_ID In Me.ColListaCambiosPasss
        Set m_ObjHistorico = Me.ColListaCambiosPasss(m_ID)
        If m_ObjHistorico.PassAntigua = p_PassEncriptada Then
            PassNoRepetida = EnumSino.No
            Set m_ObjHistorico = Nothing
            Exit Function
        End If
        Set m_ObjHistorico = Nothing
        i = i + 1
        If i = p_NumeroVeces Then
            Exit For
        End If
    Next
    PassNoRepetida = EnumSino.No
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PassNoRepetida ha devuelto el error: " & Err.Description
    End If
End Function


Public Property Get PassIncialPlanaCalculada() As String

    On Error GoTo errores
    Me.Error = ""
    
    m_sPassIncialPlanaCalculada = GenerarPassValida(8, Me.Error)
    
    PassIncialPlanaCalculada = m_sPassIncialPlanaCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.PassIncialPlanaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get UltimaAplicacionAbierta() As Aplicacion

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objUltimaAplicacionAbierta Is Nothing Then
        Set UltimaAplicacionAbierta = m_objUltimaAplicacionAbierta
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objUltimaAplicacionAbierta = Constructor.getUltimaAplicacionAbierta(Me.CorreoUsuario, Me.Error)
    Set UltimaAplicacionAbierta = m_objUltimaAplicacionAbierta
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.UltimaAplicacionAbierta ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get Permisos() As UsuarioAplicacionPermisos

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_ObjPermisos Is Nothing Then
        Set Permisos = m_ObjPermisos
        Exit Property
    End If
    If Me.ColAplicacionesPermisos Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If IDAplicacion = "" Then
        Exit Property
    End If
    If Not Me.ColAplicacionesPermisos.Exists(CStr(IDAplicacion)) Then
        Exit Property
    End If
    Set m_ObjPermisos = Me.ColAplicacionesPermisos(CStr(IDAplicacion))
    Set Permisos = m_ObjPermisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.Permisos ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Property Get EsUsuarioCalidadCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    If m_eEsUsuarioCalidadCalculado <> 0 Then
        EsUsuarioCalidadCalculado = m_eEsUsuarioCalidadCalculado
        Exit Property
    End If
    If Me.Permisos Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        m_eEsUsuarioCalidadCalculado = EnumSino.No
        EsUsuarioCalidadCalculado = m_eEsUsuarioCalidadCalculado
        Exit Property
    End If
    m_eEsUsuarioCalidadCalculado = Me.Permisos.EsUsuarioCalidadCalculado
    Me.Error = Me.Permisos.Error
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    EsUsuarioCalidadCalculado = m_eEsUsuarioCalidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.EsUsuarioCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get EsUsuarioTecnicoCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    If m_eEsUsuarioTecnicoCalculado <> 0 Then
        EsUsuarioTecnicoCalculado = m_eEsUsuarioTecnicoCalculado
        Exit Property
    End If
    If Me.Permisos Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        m_eEsUsuarioTecnicoCalculado = EnumSino.No
        EsUsuarioTecnicoCalculado = m_eEsUsuarioTecnicoCalculado
        Exit Property
    End If
    m_eEsUsuarioTecnicoCalculado = Me.Permisos.EsUsuarioTecnicoCalculado
    Me.Error = Me.Permisos.Error
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    EsUsuarioTecnicoCalculado = m_eEsUsuarioTecnicoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.EsUsuarioTecnicoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Property Get Equipo() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sEquipo <> "" Then
        Equipo = m_sEquipo
        Exit Property
    End If
    If Me.Permisos Is Nothing Then
        If Me.Error <> "" Then
            
        End If
        m_sEquipo = "Técnico"
        Equipo = m_sEquipo
        Exit Property
    End If
    If Me.Permisos.EsUsuarioAdministradorCalculado = EnumSino.Sí Then
        m_sEquipo = "Administradores"
    ElseIf Me.Permisos.EsUsuarioCalidadCalculado = EnumSino.Sí Then
        m_sEquipo = "Calidad"
    Else
        m_sEquipo = "Técnico"
    End If
    
    
    Equipo = m_sEquipo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Usuario.Equipo ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get ParaAlta() As EnumSino

    ParaAlta = m_eParaAlta

End Property

Public Property Let ParaAlta(ByVal eNewValue As EnumSino)

    m_eParaAlta = eNewValue

End Property

Public Property Get EtiquetaAplicacionesAPulsar() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Me.UltimaAplicacionAbierta Is Nothing Then
        m_sEtiquetaAplicacionesAPulsar = "1"
    Else
        
        If Me.UltimaAplicacionAbierta.EjecucionEnOficinaCalculado = EnumSino.No Then
            m_sEtiquetaAplicacionesAPulsar = "1"
        ElseIf Me.UltimaAplicacionAbierta.EjecucionEnOficinaCalculado = EnumSino.Sí Then
            m_sEtiquetaAplicacionesAPulsar = "2"
        
            
        End If
    End If
    
    EtiquetaAplicacionesAPulsar = m_sEtiquetaAplicacionesAPulsar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.EtiquetaAplicacionesAPulsar ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Function Bloquear( _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim m_Fecha As String
    On Error GoTo errores
    
    m_Fecha = Now()
    m_SQL = "UPDATE TbUsuariosAplicaciones SET TbUsuariosAplicaciones.FechaBloqueo = #" & CDate(m_Fecha) & "#, " & _
            "TbUsuariosAplicaciones.Activado=false " & _
            "WHERE (((TbUsuariosAplicaciones.UsuarioRed)='" & Me.UsuarioRed & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    Me.FechaBloqueo = m_Fecha
    
    Exit Function
    
errores:
    DoCmd.SetWarnings True
    If Err.Number <> 1000 Then
        p_Error = "El método Bloquear ha devuelto el error: " & Err.Description
    End If
End Function

Public Function DesBloquear( _
                            Optional ByRef p_Error As String _
                            ) As String

        
    On Error GoTo errores
    
    m_SQL = "UPDATE TbUsuariosAplicaciones SET TbUsuariosAplicaciones.FechaBloqueo =  Null, " & _
            "TbUsuariosAplicaciones.Activado=true " & _
            "WHERE (((TbUsuariosAplicaciones.UsuarioRed)='" & Me.UsuarioRed & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    Me.FechaBloqueo = ""
    
    Exit Function
    
errores:
    DoCmd.SetWarnings True
    If Err.Number <> 1000 Then
        p_Error = "El método DesBloquear ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ResetearPass( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_PassAleatoria As String
    On Error GoTo errores
    If Me.CorreoUsuario = "" Then
        p_Error = "No se conoce el correo del usuario"
        Err.Raise 1000
    End If
    m_PassAleatoria = GenerarPassValida(8, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_PassAleatoria = "" Then
        p_Error = "Se ha de indicar el password de un sólo uso"
        Err.Raise 1000
    End If
    m_SQL = "SELECT  TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE (((TbUsuariosAplicaciones.CorreoUsuario)='" & Me.CorreoUsuario & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "El correo del usuario no está registrado"
            Err.Raise 1000
        End If
        .Edit
            .Fields("Password") = SHA256(m_PassAleatoria)
            Me.Password = .Fields("Password")
            .Fields("TieneQueCambiarLaContrasenia") = True
            Me.TieneQueCambiarLaPass = .Fields("TieneQueCambiarLaContrasenia")
            
            .Fields("PassIncialPlana") = m_PassAleatoria
            Me.PassIncialPlana = m_PassAleatoria
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    CorreoAlUsuarioRestablecimientoPass p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    ResetearPass = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResetearPass ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Public Function CorreoAlUsuarioRestablecimientoPass( _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    
    Dim m_HTMLCabecera As String
    Dim m_HTMLTexto As String
    Dim m_HTMLFin As String
    Dim m_mensaje As String
    Dim m_Destinatarios As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
    
    On Error GoTo errores
    
    
    m_Asunto = "Envio de Pass (LANZADERA)"
    m_Destinatarios = Me.CorreoUsuario
    m_HTMLCabecera = DameHTMLCabecera("Envio de Pass " & Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    m_HTMLTexto = "Se ha reseteado la contraseña de su cuenta de la lanzadera. Si no ha sido Usted, hable con su administrador "
    m_HTMLTexto = m_HTMLTexto & vbNewLine & "Pass de un sólo uso: <br>" & Me.PassIncialPlana
    m_HTMLFin = DameFinHTML(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = m_HTMLCabecera & vbNewLine
   
    m_mensaje = m_mensaje & m_HTMLTexto & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLFin & vbNewLine
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = Me.CorreoUsuario
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlUsuarioRestablecimientoPass ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Private Function getTextoTiempoParaExpirarPass(Optional ByRef p_Error As String) As String
    
    Dim dias As Long
    Dim m_FechaProximoCambioContrasenia As String
    On Error GoTo errores
    
    If Me.Activado = False Then
        Exit Function
    End If
    If Me.PasswordNuncaCaduca = True Then
        getTextoTiempoParaExpirarPass = "Su Pass nunca caduca"
        Exit Function
    End If
    m_FechaProximoCambioContrasenia = Me.FechaProximaPassCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaProximoCambioContrasenia) Then
        dias = DateDiff("d", CDate(Date), CDate(m_FechaProximoCambioContrasenia))
        
        If dias = 1 Then
            getTextoTiempoParaExpirarPass = "Falta " & dias & " día para que su Pass expire"
        ElseIf dias = 0 Then
            getTextoTiempoParaExpirarPass = "Su Pass expira hoy"
        ElseIf dias > 0 Then
            getTextoTiempoParaExpirarPass = "Faltan " & dias & " días para que su Pass expire"
        ElseIf dias < 0 Then
            getTextoTiempoParaExpirarPass = "Su Pass expiró hace " & -dias & " dias"
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTextoTiempoParaExpirarPass ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function CambioPass( _
                                    p_PassNueva As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Usuario As String
    Dim m_PassNuevaEncriptada As String
    Dim m_FormatoCorreoCorrecto As EnumSino
    Dim m_PassDFormatoCorrecto As EnumSino
    Dim m_PassRepetida As EnumSino
    Dim m_ObjHistorico As HistoricoPass
    On Error GoTo errores
    
   
    
    m_PassDFormatoCorrecto = PassFormatoValido(p_PassNueva, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_PassDFormatoCorrecto = EnumSino.No Then
        p_Error = "La Pass debe tener al entre 8 y 16 caracteres, al menos un dígito, al menos una minúscula y al menos una mayúscula."
        Err.Raise 1000
    End If
    m_PassNuevaEncriptada = SHA256(p_PassNueva)
    m_PassRepetida = PassRepetida(m_PassNuevaEncriptada, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_PassRepetida = EnumSino.Sí Then
        p_Error = "La Pass ha sido ya usada las últimas 10 veces"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE (((TbUsuariosAplicaciones.CorreoUsuario)='" & Me.CorreoUsuario & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "El usuario no está registrado"
            Err.Raise 1000
        End If
        .Edit
            .Fields("Password") = m_PassNuevaEncriptada
            If Me.PasswordNuncaCaduca = False Then
                .Fields("FechaProximoCambioContrasenia") = DateAdd("m", 3, Date)
            Else
                .Fields("FechaProximoCambioContrasenia") = Null
            End If
            
            .Fields("TieneQueCambiarLaContrasenia") = False
            .Fields("PassIncialPlana") = Null
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set m_ObjHistorico = New HistoricoPass
    With m_ObjHistorico
        .usuario = Me.CorreoUsuario
        .FechaPass = Now()
        .PassAntigua = m_PassNuevaEncriptada
        .Save p_Error
        If p_Error <> "" Then
            p_Error = "La Pass se guardó pero no quedó registrada en el histórico"
            Err.Raise 1000
        End If
    End With
    Set m_ObjHistorico = Nothing
    CambioPass = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambioPass ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function PassRepetida( _
                                p_PassEncriptada As String, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    Dim rcdDatos As DAO.Recordset
    
    
    
    On Error GoTo errores
    If p_PassEncriptada = "" Then
        p_Error = "No se ha introducido el Password"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbUsuariosHistoricoContrasenias.* " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE (((TbUsuariosHistoricoContrasenias.Usuario)='" & Me.CorreoUsuario & _
            "') AND ((TbUsuariosHistoricoContrasenias.PassAntigua)='" & p_PassEncriptada & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            PassRepetida = EnumSino.No
        Else
            PassRepetida = EnumSino.Sí
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PassRepetida ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Public Function DesactivarActivarCuenta( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    On Error GoTo errores
    
    
    m_SQL = "SELECT  TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE (((TbUsuariosAplicaciones.Usuario)='" & Me.UsuarioRed & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "Usuario no registrado"
            Err.Raise 1000
        End If
        .Edit
            If .Fields("Activado") = False Then
                .Fields("Activado") = True
                Me.Activado = True
            Else
                .Fields("Activado") = False
                Me.Activado = False
            End If
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DesactivarActivarCuenta = False
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DesactivarActivarCuenta ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function



Function Eliminar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    On Error GoTo errores
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.No Then
        p_Error = "Acceso denegado a esta operación."
        Err.Raise 1000
    End If
    If Me.UsuarioImborrable = True Then
        p_Error = "Está marcado como imborrable"
        Err.Raise 1000
    End If
    
    m_SQL = "DELETE TbUsuariosHistoricoContrasenias.Usuario " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE (((TbUsuariosHistoricoContrasenias.Usuario)='" & Me.CorreoUsuario & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    m_SQL = "DELETE TbUsuariosAplicacionesPermisos.CorreoUsuario " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.CorreoUsuario)='" & Me.CorreoUsuario & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    m_SQL = "DELETE TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE (((TbUsuariosAplicaciones.CorreoUsuario)='" & Me.CorreoUsuario & "'));"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    Eliminar = "Usuario eliminado"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    
End Function


Private Function CorreoCambioPass( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_HTMLCabecera As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLFin As String
    Dim m_HTMLTexto As String
    Dim m_mensaje As String
    Dim m_Destinatarios As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
        
    On Error GoTo errores
    m_Asunto = "Envio de Pass (LANZADERA)"
    m_Destinatarios = Me.CorreoUsuario
    m_HTMLCabecera = DameHTMLCabecera("Cambio Pass " & Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLUsuariosConectados = DameHTMLUsuarioConectado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLFin = DameFinHTML(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTexto = "Si usted no ha solicitado una Pass hable con su administrador y borre este correo" & vbNewLine & _
                "Generada la Pass de un sólo uso para el usuario: " & Me.CorreoUsuario & vbNewLine & Me.PassIncialPlana
    
    m_mensaje = m_HTMLCabecera & vbNewLine
    m_mensaje = m_mensaje & m_HTMLUsuariosConectados & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTexto & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLFin & vbNewLine
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = Me.CorreoUsuario
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoCambioPass ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Public Function CorreoAlta( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_HTMLCabecera As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLListaPermisos As String
    Dim m_HTMLTexto As String
    Dim m_HTMLFin As String
    Dim m_mensaje As String
    Dim m_Destinatarios As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
    On Error GoTo errores
    
    m_Asunto = "Alta Usuario (LANZADERA)"
    m_Destinatarios = Me.CorreoUsuario
    m_HTMLCabecera = DameHTMLCabecera("Alta de " & Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLUsuariosConectados = DameHTMLUsuarioConectado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTexto = "Generada la Pass de un sólo uso para el usuario: " & Me.CorreoUsuario & vbNewLine & Me.PassIncialPlana
    m_HTMLListaPermisos = DameHTMLTablaPermisos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLFin = DameFinHTML(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = m_HTMLCabecera & vbNewLine
    m_mensaje = m_mensaje & m_HTMLUsuariosConectados & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & "<STRONG> ALTA DE USUARIO</STRONG>" & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTexto & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLListaPermisos & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLFin & vbNewLine
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = Me.CorreoUsuario
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlta ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function CorreoDatosUsuario( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_HTMLCabecera As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLDatosGenerales As String
    Dim m_HTMLListaPermisos As String
    Dim m_HTMLTexto As String
    Dim m_HTMLFin As String
    Dim m_mensaje As String
    Dim m_Destinatarios As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
    On Error GoTo errores
    
    m_Asunto = "Datos del Usuario (LANZADERA)"
    m_Destinatarios = Me.CorreoUsuario
    m_HTMLCabecera = DameHTMLCabecera("Datos del Usuario " & Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLUsuariosConectados = DameHTMLUsuarioConectado(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTexto = "Datos registrados del usuario: " & Me.CorreoUsuario
    m_HTMLDatosGenerales = DameHTMLTablaDatosGenerales(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLListaPermisos = DameHTMLTablaPermisos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLFin = DameFinHTML(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = m_HTMLCabecera & vbNewLine
    m_mensaje = m_mensaje & m_HTMLUsuariosConectados & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    
    m_mensaje = m_mensaje & m_HTMLTexto & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    
    m_mensaje = m_mensaje & m_HTMLDatosGenerales & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    
    m_mensaje = m_mensaje & m_HTMLListaPermisos & vbNewLine
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLFin & vbNewLine
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = Me.CorreoUsuario
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoDatosUsuario ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function DameHTMLTablaPermisos( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_mensaje As String
    Dim m_ObjCol As Scripting.Dictionary
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    Dim m_ID As Variant
    Dim m_objAplicacion As Aplicacion
    Dim m_Perfil As String
    
    On Error GoTo errores
    
    Set m_ObjCol = Me.ColAplicacionesPermisos
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
        m_mensaje = m_mensaje & "<td colspan='7' class=""ColespanArriba""> PERFILES EN APLICATIVOS</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Aplicación</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Administrador</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Calidad</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Economía</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Secretaría</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Técnico</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>CalidadAvisos</strong></td>" & vbNewLine
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        If Not m_ObjCol Is Nothing Then
            For Each m_ID In m_ObjCol
                Set m_ObjUsuarioAplicacionPermisos = m_ObjCol(m_ID)
                m_mensaje = m_mensaje & "<tr>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & m_ObjUsuarioAplicacionPermisos.Aplicacion.NombreAplicacion & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioAdministradorCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioCalidadCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioEconomiaCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioSecretariaCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioTecnicoCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                    If m_ObjUsuarioAplicacionPermisos.EsUsuarioCalidadAvisosCalculado = EnumSino.Sí Then
                        m_Perfil = "Sí"
                    Else
                        m_Perfil = "No"
                    End If
                    m_mensaje = m_mensaje & "<td> " & m_Perfil & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "</tr>" & vbNewLine
                Set m_ObjUsuarioAplicacionPermisos = Nothing
            Next
        End If
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    DameHTMLTablaPermisos = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameHTMLTablaPermisos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function DameHTMLTablaDatosGenerales( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_mensaje As String
    
    
    On Error GoTo errores
    
    
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
        m_mensaje = m_mensaje & "<td colspan='2' class=""ColespanArriba""> DATOS GENERALES DEL USUARIO</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">NOMBRE</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Nombre & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">CORREO</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.CorreoUsuario & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">USUARIO DE RED</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.UsuarioRed & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">IDUSUARIO</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.ID & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">MATRÍCULA</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Matricula & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">TELÉFONO</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Telefono & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">MÓVIL</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.Movil & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">FECHA ALTA</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.FechaAlta & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">FECHA BAJA</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & Me.FechaBaja & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">USUARIO ACTIVADO</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(Me.Activado, "Sí", "No") & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">USUARIO ADMINISTRADOR</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(Me.EsAdministradorCalculado = EnumSino.Sí, "Sí", "No") & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">USUARIO IMBORRABLE</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(Me.UsuarioImborrable = True, "Sí", "No") & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">USUARIO DE APLICACIONES DE PRUEBA</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(Me.PermisoPruebas = "Sí", "Sí", "No") & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">MANTENER LANZADERA ABIERTA</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td> " & IIf(Me.MantenerLanzaderaAbierta = True, "Sí", "No") & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">PASSWORD CADUCA</td>" & vbNewLine
            If Me.PasswordNuncaCaduca = True Then
                m_mensaje = m_mensaje & "<td>&nbsp;</td>" & vbNewLine
            Else
                m_mensaje = m_mensaje & "<td> " & Me.TextoTiempoParaExpirarPass & "</td>" & vbNewLine
            End If
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    DameHTMLTablaDatosGenerales = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameHTMLTablaDatosGenerales ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function BorrarPermisos( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    On Error GoTo errores
    If Me.CorreoUsuario = "" Then
        p_Error = "No se ha indicado la CorreoUsuario"
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & Me.CorreoUsuario & "';"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    BorrarPermisos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarPermisos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function BorrarPassHistoricas( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    On Error GoTo errores
    If Me.UsuarioRed = "" Then
        p_Error = "No se ha indicado el usuario"
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbUsuariosHistoricoContrasenias.* " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE Usuario='" & Me.UsuarioRed & "';"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    BorrarPassHistoricas = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarPassHistoricas ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Function BorrarConexiones( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    On Error GoTo errores
    If Me.CorreoUsuario = "" Then
        p_Error = "No se ha indicado la CorreoUsuario"
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbConexionesRegistro.* " & _
            "FROM TbConexionesRegistro " & _
            "WHERE Usuario='" & Me.UsuarioRed & "';"
    DoCmd.SetWarnings False
    getdb().Execute m_SQL
    DoCmd.SetWarnings True
    
    BorrarConexiones = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarConexiones ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function CampoRepetido( _
                                p_NombreCampo As String, _
                                p_Valor As String, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Where As String
    On Error GoTo errores
    If p_NombreCampo = "" Or p_Valor = "" Then
        p_Error = "Faltan datos"
        Err.Raise 1000
    End If
    
    If Me.ID <> "" Then
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "' AND ID<>" & Me.ID & ";"
    Else
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "';"
    End If
    m_SQL = "SELECT " & p_NombreCampo & " " & _
            "FROM TbUsuariosAplicaciones " & _
            m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            CampoRepetido = EnumSino.No
        Else
            CampoRepetido = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioRedRepetido ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Function Registrar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_FechaActual As String
    Dim m_FormatoCorreoValido As EnumSino
    Dim m_FormatoPassValido As EnumSino
    Dim m_ObjHistorico As HistoricoPass
    Dim m_ID As String
    Dim m_PassPlana As String
    Dim m_CampoRepetido As EnumSino
    Dim EsEdicion As EnumSino
    On Error GoTo errores
    m_FechaActual = Now()
    
    If Me.Nombre = "" Then
        p_Error = "Ha de indicar el nombre"
        Err.Raise 1000
    End If
    
    If Me.CorreoUsuario = "" Then
        p_Error = "Ha de indicar el correo"
        Err.Raise 1000
    End If
    If Me.UsuarioRed = "" Then
        p_Error = "Se ha de indicar el usuario de red"
        Err.Raise 1000
    End If
    If Me.Matricula = "" Then
        p_Error = "Se ha de indicar la Matricula"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaAlta) Then
        p_Error = "Se ha de indicar FechaAlta"
        Err.Raise 1000
    End If
    m_FormatoCorreoValido = FormatoCorreoCorrecto(Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_FormatoCorreoValido = EnumSino.No Then
        p_Error = "Ha de indicar el correo con formato válido"
        Err.Raise 1000
    End If
    If Me.ID = "" Then
        If Me.PassIncialPlana = "" Then
            Me.PassIncialPlana = Me.PassIncialPlanaCalculada
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.PassIncialPlana = "" Then
                p_Error = "Ha de indicar el password"
                Err.Raise 1000
            End If
            Me.Password = SHA256(Me.PassIncialPlana)
        End If
    End If
    m_CampoRepetido = CampoRepetido("Nombre", Me.Nombre, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos un nombre idéntico"
        Err.Raise 1000
        
    End If
    m_CampoRepetido = 0
    m_CampoRepetido = CampoRepetido("CorreoUsuario", Me.CorreoUsuario, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos un Correo idéntico"
        Err.Raise 1000
        
    End If
    m_CampoRepetido = 0
    m_CampoRepetido = CampoRepetido("UsuarioRed", Me.UsuarioRed, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos un Usuario de Red idéntico"
        Err.Raise 1000
    End If
    m_CampoRepetido = 0
    m_CampoRepetido = CampoRepetido("Matricula", Me.Matricula, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos una Matricula idéntica"
        Err.Raise 1000
    End If
    
        
    
    m_SQL = "SELECT  TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE (((TbUsuariosAplicaciones.CorreoUsuario)='" & Me.CorreoUsuario & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                m_ID = Me.IDCalculado
                p_Error = Me.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.ID = m_ID
                
                .Fields("ID") = Me.ID
                .Fields("CorreoUsuario") = LCase(Me.CorreoUsuario)
                .Fields("Password") = Me.Password
                .Fields("Nombre") = Me.Nombre
                If Me.Matricula <> "" Then
                    .Fields("Matricula") = Me.Matricula
                End If
                .Fields("UsuarioRed") = Me.UsuarioRed
                .Fields("FechaAlta") = m_FechaActual
                
                .Fields("Activado") = Me.Activado
                If IsDate(Me.FechaBloqueo) Then
                    .Fields("FechaBloqueo") = Me.FechaBloqueo
                End If
                .Fields("FechaProximoCambioContrasenia") = DateAdd("m", 3, CDate(m_FechaActual))
                
                .Fields("TieneQueCambiarLaContrasenia") = True
                .Fields("PasswordNuncaCaduca") = Me.PasswordNuncaCaduca
                If Me.Telefono <> "" Then
                    .Fields("Telefono") = Me.Telefono
                End If
                If Me.Movil <> "" Then
                    .Fields("Movil") = Me.Movil
                End If
                If m_ObjUsuarioConectadoLogin.CorreoUsuario <> "" Then
                    .Fields("Observaciones") = "Usuario dado de Registrar por " & m_ObjUsuarioConectadoLogin.CorreoUsuario
                Else
                    .Fields("Observaciones") = "Usuario dado de Registrar por " & m_UsuarioRedConectado
                End If
                If Me.PassIncialPlana <> "" Then
                    .Fields("PassIncialPlana") = Me.PassIncialPlana
                End If
                .Fields("UsuarioImborrable") = Me.UsuarioImborrable
                .Fields("EsAdministrador") = Me.EsAdministrador
                .Fields("MantenerLanzaderaAbierta") = Me.MantenerLanzaderaAbierta
            .Update
            EsEdicion = EnumSino.No
        Else
            .Edit
                
                .Fields("Nombre") = Me.Nombre
                If Me.Matricula <> "" Then
                    .Fields("Matricula") = Me.Matricula
                Else
                    .Fields("Matricula") = Null
                End If
                If Me.Telefono <> "" Then
                    .Fields("Telefono") = Me.Telefono
                End If
                If Me.Movil <> "" Then
                    .Fields("Movil") = Me.Movil
                End If
                .Fields("UsuarioRed") = Me.UsuarioRed
                .Fields("FechaAlta") = m_FechaActual
                .Fields("Activado") = Me.Activado
                If IsDate(Me.FechaBloqueo) Then
                    .Fields("FechaBloqueo") = Me.FechaBloqueo
                Else
                    .Fields("FechaBloqueo") = Null
                    .Fields("Activado") = True
                End If
                .Fields("UsuarioImborrable") = Me.UsuarioImborrable
                .Fields("EsAdministrador") = Me.EsAdministrador
                .Fields("PasswordNuncaCaduca") = Me.PasswordNuncaCaduca
                .Fields("MantenerLanzaderaAbierta") = Me.MantenerLanzaderaAbierta
                If m_ObjUsuarioConectadoLogin.CorreoUsuario <> "" Then
                    .Fields("Observaciones") = "Usuario actualizado por " & m_ObjUsuarioConectadoLogin.CorreoUsuario
                Else
                    .Fields("Observaciones") = "Usuario actualizado por " & m_UsuarioRedConectado
                End If
                
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If EsEdicion = EnumSino.Sí Then
        Set m_ObjHistorico = New HistoricoPass
        With m_ObjHistorico
            .FechaPass = Now()
            .PassAntigua = Me.Password
            .usuario = Me.CorreoUsuario
            .Save p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
        
        CorreoAlta p_Error
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Function Borrar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    
    
    On Error GoTo errores
    If Me.ID = "" Or Me.CorreoUsuario = "" Or Me.UsuarioRed = "" Then
        p_Error = "Ha de indicar el usuario"
        Err.Raise 1000
    End If
    BorrarPermisos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    BorrarPassHistoricas p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    BorrarConexiones p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT  TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE CorreoUsuario='" & Me.CorreoUsuario & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .Delete
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function



Public Function getJSON(Optional ByRef p_Error As String) As String
    Dim m_ObjPermisos As UsuarioAplicacionPermisos
    Dim m_ID As Variant
    Dim m_Campo As Variant
    Dim ValorCampo As String
    Dim m_objAplicacion As Aplicacion
    
    Dim ColDatosGeneralesUsuario As Scripting.Dictionary
    Dim ColDatosGeneralesApp As Scripting.Dictionary
    Dim colDatosPermisosApp As Scripting.Dictionary
    Dim ColApp As Scripting.Dictionary
    Dim ColApps As Scripting.Dictionary
    Dim ColUsuarioYAplicaciones As Scripting.Dictionary
    Dim Col As Scripting.Dictionary
    'ColDatosGeneralesApp("IDAplicacion")="2"
    'colDatosPermisosApp("EsUsuarioAdministrador")="Sí"
    'ColApp("DatosGenerales")=ColDatosGeneralesApp
    'ColApp("Permisos")=colDatosPermisosApp
    'ColApps("2")=ColApp
    
    'ColUsuarioYAplicaciones("DatosUsuario")= ColDatosGeneralesUsuario
    'ColUsuarioYAplicaciones("DatosAplicaciones")= ColApps
    'Col("mtr")=ColUsuarioYAplicaciones
    
    
    On Error GoTo errores
    For Each m_Campo In ColCampos
        ValorCampo = getPropiedad(CStr(m_Campo), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If ColDatosGeneralesUsuario Is Nothing Then
            Set ColDatosGeneralesUsuario = New Scripting.Dictionary
            ColDatosGeneralesUsuario.CompareMode = TextCompare
        End If
        If Not ColDatosGeneralesUsuario.Exists(m_Campo) Then
            ColDatosGeneralesUsuario.Add m_Campo, ValorCampo
        End If
    Next
    
    
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        If Nz(m_ID, "") = "" Then
            GoTo siguiente
        End If
        Debug.Print m_ID
        'If CStr(m_ID) = "51" Then Stop
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        For Each m_Campo In m_objAplicacion.ColCampos
            ValorCampo = m_objAplicacion.getPropiedad(CStr(m_Campo), p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If ColDatosGeneralesApp Is Nothing Then
                Set ColDatosGeneralesApp = New Scripting.Dictionary
                ColDatosGeneralesApp.CompareMode = TextCompare
            End If
            If Not ColDatosGeneralesApp.Exists(m_Campo) Then
                ColDatosGeneralesApp.Add m_Campo, ValorCampo
            End If
        Next
        If Not Me.ColAplicacionesPermisos Is Nothing Then
            If Me.ColAplicacionesPermisos.Exists(CStr(m_ID)) Then
                Set m_ObjPermisos = Me.ColAplicacionesPermisos(CStr(m_ID))
            End If
        End If
        
        
        
        If Not m_ObjPermisos Is Nothing Then
            
            If colDatosPermisosApp Is Nothing Then
                Set colDatosPermisosApp = New Scripting.Dictionary
                colDatosPermisosApp.CompareMode = TextCompare
            End If
            colDatosPermisosApp.Add "EsUsuarioAdministrador", m_ObjPermisos.EsUsuarioAdministrador
            colDatosPermisosApp.Add "EsUsuarioCalidad", m_ObjPermisos.EsUsuarioCalidad
            colDatosPermisosApp.Add "EsUsuarioCalidadAvisos", m_ObjPermisos.EsUsuarioCalidadAvisos
            colDatosPermisosApp.Add "EsUsuarioEconomia", m_ObjPermisos.EsUsuarioEconomia
            colDatosPermisosApp.Add "EsUsuarioSecretaria", m_ObjPermisos.EsUsuarioSecretaria
            colDatosPermisosApp.Add "EsUsuarioTecnico", m_ObjPermisos.EsUsuarioTecnico
            colDatosPermisosApp.Add "EsUsuarioSinAcceso", m_ObjPermisos.EsUsuarioSinAcceso
        End If
        If ColApp Is Nothing Then
            Set ColApp = New Scripting.Dictionary
            ColApp.CompareMode = TextCompare
        End If
        ColApp.Add "DatosGenerales", ColDatosGeneralesApp
        If colDatosPermisosApp Is Nothing Then
            ColApp.Add "Permisos", "NOTHING"
        Else
            ColApp.Add "Permisos", colDatosPermisosApp
        End If
        If ColApps Is Nothing Then
            Set ColApps = New Scripting.Dictionary
            ColApps.CompareMode = TextCompare
        End If
        ColApps.Add m_objAplicacion.IDAplicacion, ColApp
        Set m_objAplicacion = Nothing
        Set m_ObjPermisos = Nothing
        Set ColDatosGeneralesApp = Nothing
        Set colDatosPermisosApp = Nothing
        Set ColApp = Nothing
        
siguiente:
    Next
    If ColUsuarioYAplicaciones Is Nothing Then
        Set ColUsuarioYAplicaciones = New Scripting.Dictionary
        ColUsuarioYAplicaciones.CompareMode = TextCompare
    End If
    ColUsuarioYAplicaciones.Add "DatosUsuario", ColDatosGeneralesUsuario
    ColUsuarioYAplicaciones.Add "DatosAplicaciones", ColApps
    If Col Is Nothing Then
        Set Col = New Scripting.Dictionary
        Col.CompareMode = TextCompare
    End If
    Col.Add Me.UsuarioRed, ColUsuarioYAplicaciones
    
    getJSON = JsonConverter.ConvertToJson(Col)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getJSON ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbUsuariosAplicaciones")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            Set m_objColCampos = New Collection
        
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo Usuario.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "CorreoUsuario" Then
        getPropiedad = Me.CorreoUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "Password" Then
        getPropiedad = Me.Password
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioRed" Then
        getPropiedad = Me.UsuarioRed
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "Matricula" Then
        getPropiedad = Me.Matricula
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "Activado" Then
        getPropiedad = Me.Activado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaProximoCambioContrasenia" Then
        getPropiedad = Me.FechaProximoCambioContrasenia
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        getPropiedad = Me.Telefono
        Exit Function
    End If
    If m_NombrePropiedad = "Movil" Then
        getPropiedad = Me.Movil
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioImborrable" Then
        getPropiedad = Me.UsuarioImborrable
        Exit Function
    End If
    If m_NombrePropiedad = "EsAdministrador" Then
        getPropiedad = Me.EsAdministrador
        Exit Function
    End If
    If m_NombrePropiedad = "PermisosAsignados" Then
        getPropiedad = Me.PermisosAsignados
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBaja" Then
        getPropiedad = Me.FechaBaja
        Exit Function
    End If
    If m_NombrePropiedad = "PasswordNuncaCaduca" Then
        getPropiedad = Me.PasswordNuncaCaduca
        Exit Function
    End If
    If m_NombrePropiedad = "MantenerLanzaderaAbierta" Then
        getPropiedad = Me.MantenerLanzaderaAbierta
        Exit Function
    End If
    If m_NombrePropiedad = "PassIncialPlana" Then
        getPropiedad = Me.PassIncialPlana
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSSID" Then
        getPropiedad = Me.UsuarioSSID
        Exit Function
    End If
    If m_NombrePropiedad = "Id" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "JefeDelUsuario" Then
        getPropiedad = Me.JefeDelUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "PermisoPruebas" Then
        getPropiedad = Me.PermisoPruebas
        Exit Function
    End If
    If m_NombrePropiedad = "ParaTareasProgramadas" Then
        getPropiedad = Me.ParaTareasProgramadas
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBloqueo" Then
        getPropiedad = Me.FechaBloqueo
        Exit Function
    End If
    If m_NombrePropiedad = "TieneQueCambiarLaContrasenia" Then
        getPropiedad = Me.TieneQueCambiarLaPass
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimaConexion" Then
        getPropiedad = Me.FechaUltimaConexion
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Usuario.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "CorreoUsuario" Then
        Me.CorreoUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Password" Then
        Me.Password = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioRed" Then
        Me.UsuarioRed = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Matricula" Then
        Me.Matricula = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Activado" Then
        Me.Activado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaProximoCambioContrasenia" Then
        Me.FechaProximoCambioContrasenia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Telefono" Then
        Me.Telefono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Movil" Then
        Me.Movil = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioImborrable" Then
        Me.UsuarioImborrable = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsAdministrador" Then
        Me.EsAdministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PermisosAsignados" Then
        Me.PermisosAsignados = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBaja" Then
        Me.FechaBaja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PasswordNuncaCaduca" Then
        Me.PasswordNuncaCaduca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "MantenerLanzaderaAbierta" Then
        Me.MantenerLanzaderaAbierta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PassIncialPlana" Then
        Me.PassIncialPlana = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSSID" Then
        Me.UsuarioSSID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Id" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "JefeDelUsuario" Then
        Me.JefeDelUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PermisoPruebas" Then
        Me.PermisoPruebas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaTareasProgramadas" Then
        Me.ParaTareasProgramadas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBloqueo" Then
        Me.FechaBloqueo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TieneQueCambiarLaContrasenia" Then
        Me.TieneQueCambiarLaPass = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimaConexion" Then
        Me.FechaUltimaConexion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Usuario.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

