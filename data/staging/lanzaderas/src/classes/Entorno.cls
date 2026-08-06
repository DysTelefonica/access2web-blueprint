VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_ColItems As Scripting.Dictionary

Private m_sURLDirAplicaciones As String
Private m_sURLDirAplicacionLanzadera As String
Private m_sURLDirAyuda As String
Private m_sURLArchivoIniLanzadera As String
Private m_sURLDirLocal As String
Private m_sURLDirLocalGeneral As String
Private m_ObjColAplicaciones As Scripting.Dictionary
Private m_ObjColAplicacionesActivas As Scripting.Dictionary

Private m_sUsuarioSSID As String
Private m_sURLEjecutableAccess As String
Private m_eUsuarioEnLaOficina As EnumSino

Private m_sVersionAplicacion As String

Private m_sURLDirVideos As String

Private m_sURLBBDDDatos As String
Private m_sNombreBaseApp As String

Private m_objAplicacionLanzadera As Aplicacion

Private m_sURLDirRecursosLanzadera As String

Private m_sURLAyudaInstalarNetFrameWork As String
Private m_sURLDirAplicacionesLocal As String
Private m_sURLDirIconosArbol As String
Private m_sSSIDAdmin As String
Public NumeroLogins As Integer

Public URLUltimoArchivo As String
Private m_sCSS As String
Private m_sURLArchivoCSS As String
Private m_objColPermisosAplicativoPorDefecto As Scripting.Dictionary
Private m_objColPermisosTipo As Scripting.Dictionary

Private m_sTextoParaPantallaPrincipal As String
Private m_lColorMenuDesactivado As Long
Private m_lColorMenuActivado As Long

Private m_sURLArchivoInstalacionWMP As String

Private m_objColVideoCategorias As Scripting.Dictionary
Private m_objColCategoriasVideo As Scripting.Dictionary

Private m_objColEstadosWMP As Scripting.Dictionary
Private m_sURLArchivoSSID As String
Private m_ObjUsuarioConectadoMaquina As usuario
Private m_sCabeceraHTML As String
Private m_sFinHTML As String
Private m_sURLConfiguracion  As String
Private m_bPermitirDiaSinContrasenia  As Boolean
Private m_bReiniciarVersiones  As Boolean
Private m_bEstablecidoJSON  As Boolean
Public Error As String
Private m_Error As String
Private m_objColItems As Scripting.Dictionary

Private Const LANZADERA_APP_FOLDER As String = "0Lanzadera"

Private Function NormalizeEntornoDirectoryPath(ByVal p_Path As String) As String
    Dim normalizedPath As String
    normalizedPath = Trim$(Replace(p_Path, "/", "\"))
    If normalizedPath <> "" And Right$(normalizedPath, 1) <> "\" Then normalizedPath = normalizedPath & "\"
    NormalizeEntornoDirectoryPath = normalizedPath
End Function

Private Function LastDirectoryName(ByVal p_Path As String) As String
    Dim normalizedPath As String
    Dim pathWithoutTrailingSlash As String
    Dim pathParts As Variant

    normalizedPath = NormalizeEntornoDirectoryPath(p_Path)
    If normalizedPath = "" Then Exit Function

    pathWithoutTrailingSlash = Left$(normalizedPath, Len(normalizedPath) - 1)
    pathParts = Split(pathWithoutTrailingSlash, "\")
    LastDirectoryName = CStr(pathParts(UBound(pathParts)))
End Function

Private Function ParentDirectoryPath(ByVal p_Path As String) As String
    Dim normalizedPath As String
    Dim pathWithoutTrailingSlash As String
    Dim lastSeparatorPos As Long

    normalizedPath = NormalizeEntornoDirectoryPath(p_Path)
    If normalizedPath = "" Then Exit Function

    pathWithoutTrailingSlash = Left$(normalizedPath, Len(normalizedPath) - 1)
    lastSeparatorPos = InStrRev(pathWithoutTrailingSlash, "\")
    If lastSeparatorPos = 0 Then Exit Function

    ParentDirectoryPath = Left$(pathWithoutTrailingSlash, lastSeparatorPos)
End Function

Private Function ConfiguredPathIsLanzaderaFolder(ByVal p_Path As String) As Boolean
    ConfiguredPathIsLanzaderaFolder = (StrComp(LastDirectoryName(p_Path), LANZADERA_APP_FOLDER, vbTextCompare) = 0)
End Function

Private Function ResolveApplicationsRootFromConfiguredPath(ByVal p_Path As String) As String
    Dim normalizedPath As String
    Dim parentPath As String

    normalizedPath = NormalizeEntornoDirectoryPath(p_Path)
    If normalizedPath = "" Then Exit Function

    If ConfiguredPathIsLanzaderaFolder(normalizedPath) Then
        parentPath = ParentDirectoryPath(normalizedPath)
        If parentPath <> "" Then
            ResolveApplicationsRootFromConfiguredPath = parentPath
        Else
            ResolveApplicationsRootFromConfiguredPath = normalizedPath
        End If
    Else
        ResolveApplicationsRootFromConfiguredPath = normalizedPath
    End If
End Function

Public Property Get EstablecidoJSON() As Boolean

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not fso.FileExists(m_ObjEntorno.URLDirAplicaciones & "configuracion.json") Then
        m_bEstablecidoJSON = False
    Else
        m_bEstablecidoJSON = True
    End If
    EstablecidoJSON = m_bEstablecidoJSON
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EstablecidoJSON ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PermitirDiaSinContrasenia() As Boolean

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.EstablecidoJSON = False Then
        m_bPermitirDiaSinContrasenia = False
        PermitirDiaSinContrasenia = m_bPermitirDiaSinContrasenia
        Exit Property
    End If
    
    m_bPermitirDiaSinContrasenia = getParametroConf("PermitirDiaSinContrasenia", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    PermitirDiaSinContrasenia = m_bPermitirDiaSinContrasenia
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.PermitirDiaSinContrasenia ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ReiniciarVersiones() As Boolean

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.EstablecidoJSON = False Then
        m_bReiniciarVersiones = False
        ReiniciarVersiones = m_bReiniciarVersiones
        Exit Property
    End If
    m_bReiniciarVersiones = getParametroConf("ReiniciarVersiones", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    ReiniciarVersiones = m_bReiniciarVersiones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ReiniciarVersiones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get URLConfiguracion() As String

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.EstablecidoJSON = False Then
        Exit Property
    End If
    If m_sURLConfiguracion <> "" Then
        URLConfiguracion = m_sURLConfiguracion
        Exit Property
    End If
    
    m_sURLConfiguracion = Me.URLDirAplicaciones & "configuracion.json"
    
    If Not fso.FileExists(m_sURLConfiguracion) Then
        m_Error = "No se alcanza m_sURLConfiguracion" & vbNewLine & m_sURLConfiguracion
        m_sURLConfiguracion = ""
        Err.Raise 1000
    End If
    URLConfiguracion = m_sURLConfiguracion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLConfiguracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLDirAplicaciones() As String

    Dim configuredApplicationsPath As String

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    configuredApplicationsPath = ResolverRutaSegunConfiguracion(m_URLRutaAplicacionesLocal, m_URLRutaAplicacionesRemotas, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirAplicaciones = ResolveApplicationsRootFromConfiguredPath(configuredApplicationsPath)
    If Not fso.FolderExists(m_sURLDirAplicaciones) Then
        m_Error = "No se alcanza m_sURLDirAplicaciones" & vbNewLine & m_sURLDirAplicaciones
        m_sURLDirAplicaciones = ""
        Err.Raise 1000
    End If
    URLDirAplicaciones = m_sURLDirAplicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirAplicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLDirAplicacionLanzadera() As String
    
    Dim m_URLDirAplicaciones As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirAplicacionLanzadera <> "" Then
        URLDirAplicacionLanzadera = m_sURLDirAplicacionLanzadera
        Exit Property
    End If
    m_URLDirAplicaciones = Me.URLDirAplicaciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirAplicacionLanzadera = m_URLDirAplicaciones & LANZADERA_APP_FOLDER & "\"
    
    If Not fso.FolderExists(m_sURLDirAplicacionLanzadera) Then
        m_Error = "No se alcanza m_sURLDirAplicacionLanzadera" & vbNewLine & m_sURLDirAplicacionLanzadera
        m_sURLDirAplicacionLanzadera = ""
        Err.Raise 1000
    End If
    URLDirAplicacionLanzadera = m_sURLDirAplicacionLanzadera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirAplicacionLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get URLDirRecursosLanzadera() As String
    
    Dim m_URLDirAplicacionLanzadera As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirRecursosLanzadera <> "" Then
        URLDirRecursosLanzadera = m_sURLDirRecursosLanzadera
        Exit Property
    End If
    m_URLDirAplicacionLanzadera = Me.URLDirAplicacionLanzadera
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirRecursosLanzadera = m_URLDirAplicacionLanzadera & "recursos\"
    If Not fso.FolderExists(m_sURLDirRecursosLanzadera) Then
        m_Error = "No se alcanza m_sURLDirRecursosLanzadera" & vbNewLine & m_sURLDirRecursosLanzadera
        m_sURLDirRecursosLanzadera = ""
        Err.Raise 1000
    End If
    
    URLDirRecursosLanzadera = m_sURLDirRecursosLanzadera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirRecursosLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLArchivoIniLanzadera() As String
    
     Dim m_URLDirRecursosLanzadera As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLArchivoIniLanzadera <> "" Then
        URLArchivoIniLanzadera = m_sURLArchivoIniLanzadera
        Exit Property
    End If
    m_URLDirRecursosLanzadera = Me.URLDirRecursosLanzadera
    
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLArchivoIniLanzadera = m_URLDirRecursosLanzadera & fso.GetBaseName(CurrentProject.Name) & ".ini"
    
    If Not fso.FileExists(m_sURLArchivoIniLanzadera) Then
        m_Error = "No se alcanza m_sURLArchivoIniLanzadera" & vbNewLine & m_sURLArchivoIniLanzadera
        m_sURLArchivoIniLanzadera = ""
        Err.Raise 1000
    End If
    URLArchivoIniLanzadera = m_sURLArchivoIniLanzadera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoIniLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLDirLocalGeneral() As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirLocalGeneral <> "" Then
        URLDirLocalGeneral = m_sURLDirLocalGeneral
        Exit Property
    End If
    
    m_sURLDirLocalGeneral = Environ("APPDATA") & "\" & "Aplicaciones DYSN\"
    If Not fso.FolderExists(m_sURLDirLocalGeneral) Then
        fso.CreateFolder m_sURLDirLocalGeneral
    End If
    
    
    URLDirLocalGeneral = m_sURLDirLocalGeneral
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirLocalGeneral ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLDirLocal() As String
    
    Dim m_URLDirLocalGeneral As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLDirLocal <> "" Then
        URLDirLocal = m_sURLDirLocal
        Exit Property
    End If
    
    m_URLDirLocalGeneral = Me.URLDirLocalGeneral
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_sURLDirLocal = m_URLDirLocalGeneral & "Lanzadera\"
    If Not fso.FolderExists(m_sURLDirLocal) Then
        fso.CreateFolder m_sURLDirLocal
    End If
    
    URLDirLocal = m_sURLDirLocal
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColAplicaciones() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAplicaciones Is Nothing Then
        Set ColAplicaciones = m_ObjColAplicaciones
        Exit Property
    End If
    Set m_ObjColAplicaciones = Constructor.getAplicaciones(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAplicaciones = m_ObjColAplicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAplicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColAplicacionesActivas() As Scripting.Dictionary

     On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAplicacionesActivas Is Nothing Then
        Set ColAplicacionesActivas = m_ObjColAplicacionesActivas
        Exit Property
    End If
    Set m_ObjColAplicacionesActivas = Constructor.getAplicacionesActivas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAplicacionesActivas = m_ObjColAplicacionesActivas
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColAplicacionesActivas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get URLEjecutableAccess() As String

    On Error GoTo errores
    Me.Error = ""
     m_Error = ""
    If m_sURLEjecutableAccess <> "" Then
        URLEjecutableAccess = m_sURLEjecutableAccess
        Exit Property
    End If
    m_sURLEjecutableAccess = DameURLEjecutableAccess(m_Error)
    If m_Error <> "" Then
        Exit Property
    End If
    URLEjecutableAccess = m_sURLEjecutableAccess
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLEjecutableAccess ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get UsuarioEnLaOficina() As EnumSino
    
    Dim m_EnOFi As Boolean
    On Error GoTo errores
    Me.Error = ""
     m_Error = ""
    If m_eUsuarioEnLaOficina <> 0 Then
        UsuarioEnLaOficina = m_eUsuarioEnLaOficina
        Exit Property
    End If
    m_eUsuarioEnLaOficina = m_EnOficina
    
    UsuarioEnLaOficina = m_eUsuarioEnLaOficina
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuarioEnLaOficina ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get VersionAplicacion() As String

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sVersionAplicacion <> "" Then
        VersionAplicacion = m_sVersionAplicacion
        Exit Property
    End If
    m_sVersionAplicacion = LeerIni("VersionAplicacion", "", Me.Error)
    If Me.Error <> "" Then
        Exit Property
    End If
    VersionAplicacion = m_sVersionAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAyudaInstalarNetFrameWork() As String

    Dim m_URLDirRecursosLanzadera As String
    Dim m_URLAplicaciones As String
    Dim m_URLLegacyInstalador As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLAyudaInstalarNetFrameWork <> "" Then
        URLAyudaInstalarNetFrameWork = m_sURLAyudaInstalarNetFrameWork
        Exit Property
    End If
    m_URLDirRecursosLanzadera = Me.URLDirRecursosLanzadera
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    m_sURLAyudaInstalarNetFrameWork = m_URLDirRecursosLanzadera & "InstalarFrx.pdf"
    If Not fso.FileExists(m_sURLAyudaInstalarNetFrameWork) Then
        m_URLAplicaciones = Me.URLDirAplicaciones
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_URLLegacyInstalador = m_URLAplicaciones & "InstalarFrx.pdf"
        If fso.FileExists(m_URLLegacyInstalador) Then
            m_sURLAyudaInstalarNetFrameWork = m_URLLegacyInstalador
        Else
            m_Error = "No se alcanza m_sURLAyudaInstalarNetFrameWork" & vbNewLine & m_sURLAyudaInstalarNetFrameWork
            m_sURLAyudaInstalarNetFrameWork = ""
            Err.Raise 1000
        End If
    End If
    URLAyudaInstalarNetFrameWork = m_sURLAyudaInstalarNetFrameWork
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLAyudaInstalarNetFrameWork ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColPermisosTipo() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColPermisosTipo Is Nothing Then
        Set ColPermisosTipo = m_objColPermisosTipo
        Exit Property
    End If
    Set m_objColPermisosTipo = New Scripting.Dictionary
    m_objColPermisosTipo.CompareMode = TextCompare
    With m_objColPermisosTipo
        .Add "1", "EsUsuarioAdministrador"
        .Add "2", "EsUsuarioCalidad"
        .Add "3", "EsUsuarioEconomia"
        .Add "4", "EsUsuarioSecretaria"
        .Add "5", "EsUsuarioTecnico"
        .Add "6", "EsUsuarioSinAcceso"
        .Add "7", "EsUsuarioCalidadAvisos"
    End With
   
    
    Set ColPermisosTipo = m_objColPermisosTipo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPermisosTipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColVideoCategorias() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColVideoCategorias Is Nothing Then
        Set ColVideoCategorias = m_objColVideoCategorias
        Exit Property
    End If
    Set m_objColVideoCategorias = New Scripting.Dictionary
    m_objColVideoCategorias.CompareMode = TextCompare
    With m_objColVideoCategorias
        .Add "1", "Ayuda"
        .Add "2", "Encuesta"
    End With
   Set ColVideoCategorias = m_objColVideoCategorias
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColVideoCategorias ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get ColEstadosWMP() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColEstadosWMP Is Nothing Then
        Set ColEstadosWMP = m_objColEstadosWMP
        Exit Property
    End If
    Set m_objColEstadosWMP = New Scripting.Dictionary
    m_objColEstadosWMP.CompareMode = TextCompare
    With m_objColEstadosWMP
        .Add "wmppsStopped", 1
        .Add "wmppsPaused", 2
        .Add "wmppsPlaying", 3
        .Add "wmppsScanForward", 4
        .Add "wmppsScanReverse", 5
        .Add "wmppsBuffering", 6
        .Add "wmppsWaiting", 7
        .Add "wmppsMediaEnded", 8
        .Add "wmppsTransitioning", 9
        .Add "wmppsReady", 10
        .Add "wmppsReconnecting", 1
        .Add "wmppsLast2", 12
    End With
   
    
    Set ColEstadosWMP = m_objColEstadosWMP
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosWMP ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property


Public Property Get ColCategoriasVideo() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColCategoriasVideo Is Nothing Then
        Set ColCategoriasVideo = m_objColCategoriasVideo
        Exit Property
    End If
    Set m_objColCategoriasVideo = New Scripting.Dictionary
    m_objColCategoriasVideo.CompareMode = TextCompare
    With m_objColCategoriasVideo
        .Add "Ayuda", 1
        .Add "Encuesta", 2
    End With
   
    
    Set ColCategoriasVideo = m_objColCategoriasVideo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCategoriasVideo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property

Public Property Get ColorMenuDesactivado() As Long
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    m_lColorMenuDesactivado = RGB(31, 73, 125)
    
    ColorMenuDesactivado = m_lColorMenuDesactivado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColorMenuDesactivado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColorMenuActivado() As Long
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_lColorMenuActivado = RGB(76, 109, 151)
    ColorMenuActivado = m_lColorMenuActivado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColorMenuActivado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get AplicacionLanzadera() As Aplicacion
    
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_objAplicacion As Aplicacion
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objAplicacionLanzadera Is Nothing Then
        Set AplicacionLanzadera = m_objAplicacionLanzadera
        Exit Property
    End If
    Set m_Col = Me.ColAplicaciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_objAplicacion = m_Col(m_ID)
        If m_objAplicacion.NombreAplicacion = "Lanzadera" Then
            Set m_objAplicacionLanzadera = m_objAplicacion
            Set AplicacionLanzadera = m_objAplicacionLanzadera
            Exit Property
        End If
        Set m_objAplicacion = Nothing
    Next
    Set m_objAplicacion = Nothing
    
    AplicacionLanzadera = m_objAplicacionLanzadera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.AplicacionLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property
Public Property Get UsuarioConectadoMaquina() As usuario
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioConectadoMaquina Is Nothing Then
        Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
        Exit Property
    End If
    Set m_ObjUsuarioConectadoMaquina = Constructor.getUsuarioConectadoPorMaquina(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuarioConectadoMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property
Public Property Get URLDirIconosArbol() As String

    Dim m_URLDirRecursosLanzadera As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirIconosArbol <> "" Then
        URLDirIconosArbol = m_sURLDirIconosArbol
        Exit Property
    End If
    m_URLDirRecursosLanzadera = Me.URLDirRecursosLanzadera
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirIconosArbol = m_URLDirRecursosLanzadera & "IconosArbol\"
    If Not fso.FolderExists(m_sURLDirIconosArbol) Then
        m_Error = "No se alcanza m_sURLDirIconosArbol" & vbNewLine & m_sURLDirIconosArbol
        m_sURLDirIconosArbol = ""
        Err.Raise 1000
    End If
    
    URLDirIconosArbol = m_sURLDirIconosArbol
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirIconosArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CSS() As String
    
    Dim m_URLCSS As String
    Dim m_objStream As ADODB.stream
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If m_sCSS <> "" Then
        CSS = m_sCSS
        Exit Property
    End If
    m_URLCSS = Me.URLArchivoCSS
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_objStream = New ADODB.stream
    With m_objStream
        .Open
        .Type = adTypeText
        .Charset = "utf-8"
        .LoadFromFile m_URLCSS
        m_sCSS = .ReadText
        .Close
    End With
    
    Set m_objStream = Nothing
    CSS = m_sCSS
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CSS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Get URLArchivoCSS() As String
    
    Dim m_URLDirAplicaciones As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLArchivoCSS <> "" Then
        URLArchivoCSS = m_sURLArchivoCSS
        Exit Property
    End If
    m_URLDirAplicaciones = Me.URLDirAplicaciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_sURLArchivoCSS = m_URLDirAplicaciones & "css.txt"
    If Not fso.FileExists(m_sURLArchivoCSS) Then
        m_Error = "No se alcanza el archivo m_sURLArchivoCSS" & vbNewLine & m_sURLArchivoCSS
        m_sURLArchivoCSS = ""
        Err.Raise 1000
    End If
    URLArchivoCSS = m_sURLArchivoCSS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoCSS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Get URLDirAyuda() As String
    
    Dim m_URLDirAplicacion As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If m_sURLDirAyuda <> "" Then
        URLDirAyuda = m_sURLDirAyuda
        Exit Property
    End If
    m_URLDirAplicacion = Me.URLDirAplicacionLanzadera
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirAyuda = m_URLDirAplicacion & "ANEXOS\AYUDA\"
    
    If Not fso.FolderExists(m_sURLDirAyuda) Then
        m_Error = "No se alcanza m_sURLDirAyuda " & vbNewLine & m_sURLDirAyuda
        m_sURLDirAyuda = ""
        Err.Raise 1000
    End If
    
    URLDirAyuda = m_sURLDirAyuda
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirAyuda ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get URLDirVideos() As String
    
    Dim m_URLDirAplicacionAplicaciones As String
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLDirVideos <> "" Then
        URLDirVideos = m_sURLDirVideos
        Exit Property
    End If
    m_URLDirAplicacionAplicaciones = Me.URLDirAplicaciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLDirVideos = m_URLDirAplicacionAplicaciones & "00videos\"
    
    If Not fso.FolderExists(m_sURLDirVideos) Then
        m_sURLDirVideos = ""
    End If
    
    URLDirVideos = m_sURLDirVideos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLDirVideos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get URLArchivoSSID() As String
    
    Dim m_URLDirLocalGeneral As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLArchivoSSID <> "" Then
        URLArchivoSSID = m_sURLArchivoSSID
        Exit Property
    End If
    m_URLDirLocalGeneral = Me.URLDirLocalGeneral
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLArchivoSSID = m_URLDirLocalGeneral & "ssid.txt"
    
    URLArchivoSSID = m_sURLArchivoSSID
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoSSID ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CabeceraHTML() As String
    
    'variable #titulo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCabeceraHTML <> "" Then
        CabeceraHTML = m_sCabeceraHTML
        Exit Property
    End If
    m_sCabeceraHTML = "<!DOCTYPE html>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<html lang=""es"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<head>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta charset=""ISO-8859-1"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<title>#titulo</title>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<style type=""text/css"">" & vbNewLine
        m_sCabeceraHTML = m_sCabeceraHTML & m_ObjEntorno.CSS & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "</style>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "</head>" & vbNewLine
    m_sCabeceraHTML = m_sCabeceraHTML & "<body>" & vbNewLine

    CabeceraHTML = m_sCabeceraHTML
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CabeceraHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FinHTML() As String
    
    'variable #titulo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sFinHTML <> "" Then
        FinHTML = m_sFinHTML
        Exit Property
    End If
    m_sFinHTML = m_sFinHTML & "</body>" & vbNewLine
    m_sFinHTML = m_sFinHTML & "</html>" & vbNewLine
    FinHTML = m_sFinHTML
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.FinHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ColItems() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColItems Is Nothing Then
        Set ColItems = m_objColItems
        Exit Property
    End If
    Set m_objColItems = New Scripting.Dictionary
    m_objColItems.CompareMode = TextCompare
    With m_objColItems
        .Add "URLDirAplicaciones", ""
        .Add "URLDirAplicacionLanzadera", ""
        .Add "URLDirAyuda", ""
        .Add "URLDirRecursosLanzadera", ""
        .Add "ColAplicaciones", "o"
        .Add "ColAplicacionesActivas", "o"
        .Add "CSS", ""
        .Add "URLArchivoCSS", ""
        .Add "URLArchivoIniLanzadera", ""
        .Add "URLAyudaInstalarNetFrameWork", ""
        .Add "URLDirLocal", ""
        .Add "URLEjecutableAccess", ""
        .Add "VersionAplicacion", ""
        .Add "ColPermisosTipo", "o"
        .Add "AplicacionLanzadera", "o"
        .Add "ColorMenuDesactivado", ""
        .Add "ColorMenuActivado", ""
        .Add "URLDirIconosArbol", ""
        .Add "URLDirVideos", ""
        .Add "ColVideoCategorias", "o"
        .Add "ColCategoriasVideo", "o"
        .Add "ColEstadosWMP", "o"
        .Add "URLArchivoSSID", ""
        .Add "URLDirLocalGeneral", ""
        .Add "CabeceraHTML", ""
        .Add "FinHTML", ""
        .Add "URLConfiguracion", ""
        .Add "PermitirDiaSinContrasenia", ""
        .Add "ReiniciarVersiones", ""
    End With
    Set ColItems = m_objColItems
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColItems ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Function getPropiedad(p_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If p_NombrePropiedad = "URLDirAplicaciones" Then
        getPropiedad = Me.URLDirAplicaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirAplicacionLanzadera" Then
        getPropiedad = Me.URLDirAplicacionLanzadera
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirAyuda" Then
        getPropiedad = Me.URLDirAyuda
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirRecursosLanzadera" Then
        getPropiedad = Me.URLDirRecursosLanzadera
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAplicaciones" Then
        Set getPropiedad = Me.ColAplicaciones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColAplicacionesActivas" Then
        Set getPropiedad = Me.ColAplicacionesActivas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CSS" Then
        getPropiedad = Me.CSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombrePropiedad = "URLArchivoCSS" Then
        getPropiedad = Me.URLArchivoCSS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLArchivoIniLanzadera" Then
        getPropiedad = Me.URLArchivoIniLanzadera
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLAyudaInstalarNetFrameWork" Then
        getPropiedad = Me.URLAyudaInstalarNetFrameWork
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirLocal" Then
        getPropiedad = Me.URLDirLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLEjecutableAccess" Then
        getPropiedad = Me.URLEjecutableAccess
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "VersionAplicacion" Then
        getPropiedad = Me.VersionAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColPermisosTipo" Then
        Set getPropiedad = Me.ColPermisosTipo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "AplicacionLanzadera" Then
        Set getPropiedad = Me.AplicacionLanzadera
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColorMenuDesactivado" Then
        getPropiedad = Me.ColorMenuDesactivado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColorMenuActivado" Then
        getPropiedad = Me.ColorMenuActivado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirIconosArbol" Then
        getPropiedad = Me.URLDirIconosArbol
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombrePropiedad = "URLDirVideos" Then
        getPropiedad = Me.URLDirVideos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColVideoCategorias" Then
        Set getPropiedad = Me.ColVideoCategorias
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColCategoriasVideo" Then
        Set getPropiedad = Me.ColCategoriasVideo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ColEstadosWMP" Then
        Set getPropiedad = Me.ColEstadosWMP
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLArchivoSSID" Then
        getPropiedad = Me.URLArchivoSSID
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLDirLocalGeneral" Then
        getPropiedad = Me.URLDirLocalGeneral
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "CabeceraHTML" Then
        getPropiedad = Me.CabeceraHTML
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "FinHTML" Then
        getPropiedad = Me.FinHTML
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "URLConfiguracion" Then
        getPropiedad = Me.URLConfiguracion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "PermitirDiaSinContrasenia" Then
        getPropiedad = Me.PermitirDiaSinContrasenia
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombrePropiedad = "ReiniciarVersiones" Then
        getPropiedad = Me.ReiniciarVersiones
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    p_Error = "La propiedad " & p_NombrePropiedad & " no está codificada en Entorno.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Entorno.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function







