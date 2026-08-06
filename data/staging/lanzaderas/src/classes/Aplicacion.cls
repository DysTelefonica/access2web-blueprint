VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Aplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public IDAplicacion As String
Public NombreAplicacion As String
Public NombreCorto As String
Public NombreEjecutable As String
Public NombreArchivoDatos As String
Public Pass As String
Public NombreCarpeta As String
Public NombreCarpetaTemporal As String
Public NombreIcono As String
Public NombreIconoLanzadera As String
Public NombreIconoParaArbol As String
Public EjecucionEnOficina As String
Public NombreCarpetaDocumentacion As String
Public NombreDirectorioIconos As String
Public NombreDirectorioAyuda As String
Public NombreDirectorioRecursos As String
Public EnPruebas As String
Public ConIconoEnLanzadera As String
Public Comando As String
Public NombreFuncionPublicacion As String
Public TituloAplicacion As String
Public URLDIrectorioIconoAplicacion As String

Private m_sURLDIrectorioIconosUsuario As String
Private m_sURLDIrectorioIconosServidor As String
Private m_sURLDIrectorioIconos As String

Private m_sURLBBDDDatosEnTdE As String

Private m_eEjecucionEnOficinaCalculado As EnumSino
Private m_eEnPruebasCalculado As EnumSino
Private m_sURLIconoAcceso As String
Private m_sURLIconoNOAcceso As String

Private m_sURLDirectorioAplicacionServidor As String
Private m_sURLDirectorioAplicacionUsuario As String
Private m_sURLDirectorioAplicacion As String

Private m_sURLDirectorioRecursosServidor As String
Private m_sURLDirectorioRecursosUsuario As String
Private m_sURLDirectorioRecursos As String

Private m_sURLDirectorioDocumentacionServidor As String
Private m_sURLDirectorioDocumentacionUsuario As String
Private m_sURLDirectorioDocumentacion As String

Private m_sURLDirectorioAyudaServidor As String
Private m_sURLDirectorioAyudaUsuario As String
Private m_sURLDirectorioAyuda As String

Private m_sURLBmpUsuario As String
Private m_sURLBmpServidor As String
Private m_sURLBmp As String

Private m_sNombreBase As String
Private m_sURLIconoEnServidor As String
Private m_sURLIconoEnUsuario As String
Private m_sURLIcono As String

Private m_sURLIConoParaArbolEnUsuario As String
Private m_sURLIConoParaArbolEnServidor As String
Private m_sURLIconoParaArbol As String

Private m_sURLCarpetaAppData As String

Private m_sURLIniServidor As String
Private m_sURLIniUsuario As String
Private m_sURLIni As String

Private m_sNombreArchivoIni As String
Private m_sVersionServidor As String
Private m_sVersionUsuario As String
Private m_eHayCambioDeVersion As EnumSino

Private m_sURLEjecutableEnServidor As String
Private m_sURLEjecutableUsuario As String
Private m_sURLEjecutable As String

Private m_sIDAplicacionCalculada As String
Private m_eConIconoEnLanzaderaCalculado As EnumSino
Private m_sNombreIconoArbol As String
Private m_objColVideos As Scripting.Dictionary
Private m_objColVideosDeAyuda As Scripting.Dictionary
Private m_objColVideosDeEncuesta As Scripting.Dictionary
Private m_sURLDirectorioVideos As String
Private m_objUsuariosAplicacionAbierta As Scripting.Dictionary
Private m_objColCampos As Collection
Public Error As String

Private Function ResolverRutaActiva( _
                                ByVal p_RutaUsuario As String, _
                                ByVal p_RutaServidor As String _
                                ) As String
    On Error GoTo errores
    Me.Error = ""

    ResolverRutaActiva = ResolverRutaSegunConfiguracion(p_RutaUsuario, p_RutaServidor, Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.ResolverRutaActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Property Get UsuariosAplicacionAbierta() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_objUsuariosAplicacionAbierta Is Nothing Then
        Set UsuariosAplicacionAbierta = m_objUsuariosAplicacionAbierta
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objUsuariosAplicacionAbierta = Constructor.getUsuariosAplicacionAbierta(Me.IDAplicacion, Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    
    Set UsuariosAplicacionAbierta = m_objUsuariosAplicacionAbierta
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.UsuariosAplicacionAbierta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Set UsuariosAplicacionAbierta(ByVal objNewValue As Variant)

    Set m_objUsuariosAplicacionAbierta = objNewValue

End Property

Public Property Get EjecucionEnOficinaCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.EjecucionEnOficina = "Sí" Then
        m_eEjecucionEnOficinaCalculado = EnumSino.Sí
    ElseIf Me.EjecucionEnOficina = "No" Then
        m_eEjecucionEnOficinaCalculado = EnumSino.No
    End If
    EjecucionEnOficinaCalculado = m_eEjecucionEnOficinaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.EjecucionEnOficinaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property


Public Property Get EnPruebasCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    If Me.EnPruebas = "Sí" Then
        m_eEnPruebasCalculado = EnumSino.Sí
    Else
         m_eEnPruebasCalculado = EnumSino.No
        
    End If
    EnPruebasCalculado = m_eEnPruebasCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.EnPruebasCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLIconoAcceso() As String
    
    Dim m_URLDirIconos As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLIconoAcceso <> "" Then
        URLIconoAcceso = m_sURLIconoAcceso
        Exit Property
    End If
    If m_ObjEntorno.AplicacionLanzadera Is Nothing Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Me.Error = "No se ha podido obtener la aplicación de la lanzadera"
        Err.Raise 1000
    End If
    If m_ObjEntorno.AplicacionLanzadera.URLDirectorioRecursosServidor = "" Then
        Me.Error = m_ObjEntorno.AplicacionLanzadera.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Me.Error = "No se ha podido obtener AplicacionLanzadera.URLDirectorioRecursosServidor"
        Err.Raise 1000
    End If
    m_URLDirIconos = m_ObjEntorno.AplicacionLanzadera.URLDirectorioRecursosServidor & "iconos\"
    If Me.NombreIconoLanzadera = "" Then
        Exit Property
    End If
    m_sURLIconoAcceso = m_URLDirIconos & Me.NombreIconoLanzadera & "_C.png"
    URLIconoAcceso = m_sURLIconoAcceso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIconoAcceso ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLIconoNOAcceso() As String
    Dim m_URLDirIconos As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLIconoNOAcceso <> "" Then
        URLIconoNOAcceso = m_sURLIconoNOAcceso
        Exit Property
    End If
    If m_ObjEntorno.AplicacionLanzadera Is Nothing Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Me.Error = "No se ha podido obtener la aplicación de la lanzadera"
        Err.Raise 1000
    End If
    If m_ObjEntorno.AplicacionLanzadera.URLDirectorioRecursosServidor = "" Then
        Me.Error = m_ObjEntorno.AplicacionLanzadera.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Me.Error = "No se ha podido obtener AplicacionLanzadera.URLDirectorioRecursosServidor"
        Err.Raise 1000
    End If
    m_URLDirIconos = m_ObjEntorno.AplicacionLanzadera.URLDirectorioRecursosServidor & "iconos\"
    If Me.NombreIconoLanzadera = "" Then
        Exit Property
    End If
    m_sURLIconoNOAcceso = m_URLDirIconos & Me.NombreIconoLanzadera & "_BN.png"
    URLIconoNOAcceso = m_sURLIconoNOAcceso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIconoAcceso ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLDirectorioAplicacionServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAplicacionServidor <> "" Then
        URLDirectorioAplicacionServidor = m_sURLDirectorioAplicacionServidor
        Exit Property
    End If
    If m_ObjEntorno.URLDirAplicaciones = "" Then
        Me.Error = m_ObjEntorno.Error
        Exit Property
    End If
    If Me.NombreCarpeta = "" Then
        Exit Property
    End If
    m_sURLDirectorioAplicacionServidor = m_ObjEntorno.URLDirAplicaciones & Me.NombreCarpeta & "\"
    URLDirectorioAplicacionServidor = m_sURLDirectorioAplicacionServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAplicacionServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLDirectorioAplicacionUsuario() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAplicacionUsuario <> "" Then
        URLDirectorioAplicacionUsuario = m_sURLDirectorioAplicacionUsuario
        Exit Property
    End If
    If Me.URLCarpetaAppData = "" Then
        'C:\Users\adm\AppData\Roaming\Agedys\Aplicaciones DYSN\
        Exit Property
    End If
    If Me.NombreCarpetaTemporal = "" Then
        Exit Property
    End If
    
    m_sURLDirectorioAplicacionUsuario = Me.URLCarpetaAppData & Me.NombreCarpetaTemporal & "\"
    If Not fso.FolderExists(m_sURLDirectorioAplicacionUsuario) Then
        fso.CreateFolder m_sURLDirectorioAplicacionUsuario
    End If
    URLDirectorioAplicacionUsuario = m_sURLDirectorioAplicacionUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAplicacionUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLDirectorioRecursosServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioRecursosServidor <> "" Then
        URLDirectorioRecursosServidor = m_sURLDirectorioRecursosServidor
        Exit Property
    End If
    If Me.URLDirectorioAplicacionServidor = "" Then
        Exit Property
    End If
    m_sURLDirectorioRecursosServidor = Me.URLDirectorioAplicacionServidor & "recursos\"
    If Not fso.FolderExists(m_sURLDirectorioRecursosServidor) Then
        Me.Error = "No es alcanzable la carpeta: " & m_sURLDirectorioRecursosServidor
        m_sURLIconoEnServidor = ""
        Err.Raise 1000
    End If
    URLDirectorioRecursosServidor = m_sURLDirectorioRecursosServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioRecursosServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDirectorioRecursosUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioRecursosUsuario <> "" Then
        URLDirectorioRecursosUsuario = m_sURLDirectorioRecursosUsuario
        Exit Property
    End If
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    m_sURLDirectorioRecursosUsuario = Me.URLDirectorioAplicacionUsuario & "recursos\"
    If Not fso.FolderExists(m_sURLDirectorioRecursosUsuario) Then
        Me.Error = "No es alcanzable el directorio de recursos: " & m_sURLDirectorioRecursosUsuario
        m_sURLDirectorioRecursosUsuario = ""
        Err.Raise 1000
    End If
    URLDirectorioRecursosUsuario = m_sURLDirectorioRecursosUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioRecursosUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDirectorioRecursos() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioRecursos <> "" Then
        URLDirectorioRecursos = m_sURLDirectorioRecursos
        Exit Property
    End If
    m_sURLDirectorioRecursos = ResolverRutaActiva(Me.URLDirectorioRecursosUsuario, Me.URLDirectorioRecursosServidor)
    If Not fso.FolderExists(m_sURLDirectorioRecursos) Then
        Me.Error = "No es alcanzable el directorio: " & m_sURLDirectorioRecursos
        m_sURLDirectorioRecursos = ""
        Err.Raise 1000
    End If
    URLDirectorioRecursos = m_sURLDirectorioRecursos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioRecursos ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioDocumentacionServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioDocumentacionServidor <> "" Then
        URLDirectorioDocumentacionServidor = m_sURLDirectorioDocumentacionServidor
        Exit Property
    End If
    If Me.URLDirectorioAplicacionServidor = "" Then
        Exit Property
    End If
    If Me.NombreCarpetaDocumentacion = "" Then
        Exit Property
    End If
    m_sURLDirectorioDocumentacionServidor = Me.URLDirectorioAplicacionServidor & Me.NombreCarpetaDocumentacion & "\"
    
    URLDirectorioDocumentacionServidor = m_sURLDirectorioDocumentacionServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioDocumentacionServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioDocumentacionUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioDocumentacionUsuario <> "" Then
        URLDirectorioDocumentacionUsuario = m_sURLDirectorioDocumentacionUsuario
        Exit Property
    End If
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreCarpetaDocumentacion = "" Then
        Exit Property
    End If
    m_sURLDirectorioDocumentacionUsuario = Me.URLDirectorioAplicacionUsuario & Me.NombreCarpetaDocumentacion & "\"
    
    URLDirectorioDocumentacionUsuario = m_sURLDirectorioDocumentacionUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioDocumentacionUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDirectorioDocumentacion() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioDocumentacion <> "" Then
        URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
        Exit Property
    End If
    
    m_sURLDirectorioDocumentacion = ResolverRutaActiva(Me.URLDirectorioDocumentacionUsuario, Me.URLDirectorioDocumentacionServidor)
    If Not fso.FolderExists(m_sURLDirectorioDocumentacion) Then
        Me.Error = "No es alcanzable el directorio de recursos: " & m_sURLDirectorioDocumentacion
        m_sURLDirectorioDocumentacion = ""
        Err.Raise 1000
    End If
    URLDirectorioDocumentacion = m_sURLDirectorioDocumentacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioDocumentacion ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioAyudaServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAyudaServidor <> "" Then
        URLDirectorioAyudaServidor = m_sURLDirectorioAyudaServidor
        Exit Property
    End If
    If Me.URLDirectorioDocumentacionServidor = "" Then
        Exit Property
    End If
    If Me.NombreDirectorioAyuda = "" Then
        Exit Property
    End If
    m_sURLDirectorioAyudaServidor = Me.URLDirectorioDocumentacionServidor & Me.NombreDirectorioAyuda & "\"
    
    URLDirectorioAyudaServidor = m_sURLDirectorioAyudaServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAyudaServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioAyudaUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAyudaUsuario <> "" Then
        URLDirectorioAyudaUsuario = m_sURLDirectorioAyudaUsuario
        Exit Property
    End If
    If Me.URLDirectorioDocumentacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreDirectorioAyuda = "" Then
        Exit Property
    End If
    m_sURLDirectorioAyudaUsuario = Me.URLDirectorioDocumentacionUsuario & Me.NombreDirectorioAyuda & "\"
    If Not fso.FolderExists(m_sURLDirectorioAyudaUsuario) Then
        Me.Error = "No es alcanzable el directorio de recursos: " & m_sURLDirectorioAyudaUsuario
        m_sURLDirectorioAyudaUsuario = ""
        Err.Raise 1000
    End If
    URLDirectorioAyudaUsuario = m_sURLDirectorioAyudaUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAyudaUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioAyuda() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAyuda <> "" Then
        URLDirectorioAyuda = m_sURLDirectorioAyuda
        Exit Property
    End If
    
    m_sURLDirectorioAyuda = ResolverRutaActiva(Me.URLDirectorioAplicacionUsuario, Me.URLDirectorioAplicacionServidor)
    If Not fso.FolderExists(m_sURLDirectorioAyuda) Then
        Me.Error = "No es alcanzable el directorio: " & m_sURLDirectorioAyuda
        m_sURLDirectorioAyuda = ""
        Err.Raise 1000
    End If
    URLDirectorioAyuda = m_sURLDirectorioAyuda
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAyuda ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLDirectorioAplicacion() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioAplicacion <> "" Then
        URLDirectorioAplicacion = m_sURLDirectorioAplicacion
        Exit Property
    End If
    m_sURLDirectorioAplicacion = ResolverRutaActiva(Me.URLDirectorioAplicacionUsuario, Me.URLDirectorioAplicacionServidor)
    
    URLDirectorioAplicacion = m_sURLDirectorioAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioAplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Property Get URLBmpServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLBmpServidor <> "" Then
        URLBmpServidor = m_sURLBmpServidor
        Exit Property
    End If
    
    If Me.URLDirectorioRecursosServidor = "" Then
        Exit Property
    End If
    If Me.NombreBase = "" Then
        Exit Property
    End If
    
    m_sURLBmpServidor = Me.URLDirectorioRecursosServidor & Me.NombreBase & ".bmp"
    
    URLBmpServidor = m_sURLBmpServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLBmpServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLBmpUsuario() As String

     On Error GoTo errores
    Me.Error = ""
    If m_sURLBmpUsuario <> "" Then
        URLBmpUsuario = m_sURLBmpUsuario
        Exit Property
    End If
    
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreBase = "" Then
        Exit Property
    End If
    
    m_sURLBmpUsuario = Me.URLDirectorioAplicacionUsuario & Me.NombreBase & ".bmp"
    
    URLBmpUsuario = m_sURLBmpUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLBmpUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLBmp() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLBmp <> "" Then
        URLBmp = m_sURLBmp
        Exit Property
    End If
    m_sURLBmp = ResolverRutaActiva(Me.URLBmpUsuario, Me.URLBmpServidor)
    If Not fso.FileExists(m_sURLBmp) Then
        Me.Error = "No es alcanzable el archivo: " & m_sURLBmp
        m_sURLBmp = ""
        Err.Raise 1000
    End If
    URLBmp = m_sURLBmp
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLBmp ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get NombreBase() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sNombreBase <> "" Then
        NombreBase = m_sNombreBase
        Exit Property
    End If
    If Me.NombreEjecutable = "" Then
        Exit Property
    End If
    m_sNombreBase = fso.GetBaseName(Me.NombreEjecutable)
    
    NombreBase = m_sNombreBase
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.NombreBase ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLCarpetaAppData() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLCarpetaAppData <> "" Then
        URLCarpetaAppData = m_sURLCarpetaAppData
        Exit Property
    End If
    If Me.NombreCarpetaTemporal = "" Then
        Exit Property
    End If
    
    m_sURLCarpetaAppData = Environ("APPDATA") & "\" & "Aplicaciones DYSN\"
    If Not fso.FolderExists(m_sURLCarpetaAppData) Then
        fso.CreateFolder m_sURLCarpetaAppData
        m_sURLCarpetaAppData = m_sURLCarpetaAppData & "\"
    End If
    
    
    URLCarpetaAppData = m_sURLCarpetaAppData
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLCarpetaAppData ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIniServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIniServidor <> "" Then
        URLIniServidor = m_sURLIniServidor
        Exit Property
    End If
    If Me.URLDirectorioRecursosServidor = "" Then
        Exit Property
    End If
    If Me.NombreArchivoIni = "" Then
        Exit Property
    End If
    m_sURLIniServidor = Me.URLDirectorioRecursosServidor & Me.NombreArchivoIni
    
    URLIniServidor = m_sURLIniServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIniServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIniUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIniUsuario <> "" Then
        URLIniUsuario = m_sURLIniUsuario
        Exit Property
    End If
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreArchivoIni = "" Then
        Exit Property
    End If
    m_sURLIniUsuario = Me.URLDirectorioAplicacionUsuario & Me.NombreArchivoIni
    
    URLIniUsuario = m_sURLIniUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIniUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLIni() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLIni <> "" Then
        URLIni = m_sURLIni
        Exit Property
    End If
    m_sURLIni = ResolverRutaActiva(Me.URLIniUsuario, Me.URLIniServidor)
    If Not fso.FileExists(m_sURLIni) Then
        Me.Error = "No es alcanzable el archivo: " & m_sURLIni
        m_sURLIni = ""
        Err.Raise 1000
    End If
    URLIni = m_sURLIni
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIni ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get NombreArchivoIni() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sNombreArchivoIni <> "" Then
        NombreArchivoIni = m_sNombreArchivoIni
        Exit Property
    End If
    If Me.NombreBase = "" Then
        Exit Property
    End If
    
    m_sNombreArchivoIni = Me.NombreBase & ".ini"
    NombreArchivoIni = m_sNombreArchivoIni
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.NombreArchivoIni ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Property Get URLIconoEnUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIconoEnUsuario <> "" Then
        URLIconoEnUsuario = m_sURLIconoEnUsuario
        Exit Property
    End If
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreIcono = "" Then
        Exit Property
    End If
    m_sURLIconoEnUsuario = Me.URLDirectorioAplicacionUsuario & Me.NombreIcono
    
    URLIconoEnUsuario = m_sURLIconoEnUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIconoEnUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIconoEnServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIconoEnServidor <> "" Then
        URLIconoEnServidor = m_sURLIconoEnServidor
        Exit Property
    End If
    If Me.URLDirectorioRecursosServidor = "" Then
        Exit Property
    End If
    If Me.NombreIcono = "" Then
        Exit Property
    End If
    m_sURLIconoEnServidor = Me.URLDirectorioRecursosServidor & Me.NombreIcono
    
    URLIconoEnServidor = m_sURLIconoEnServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIconoEnServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIcono() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIcono <> "" Then
        URLIcono = m_sURLIcono
        Exit Property
    End If
    m_sURLIcono = ResolverRutaActiva(Me.URLIconoEnUsuario, Me.URLIconoEnServidor)
    If Not fso.FileExists(m_sURLIcono) Then
        Me.Error = "No es alcanzable el archivo: " & m_sURLIcono
        m_sURLIcono = ""
        Err.Raise 1000
    End If
    URLIcono = m_sURLIcono
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIcono ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDIrectorioIconosUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDIrectorioIconosUsuario <> "" Then
        URLDIrectorioIconosUsuario = m_sURLDIrectorioIconosUsuario
        Exit Property
    End If
    If Me.URLDirectorioRecursosUsuario = "" Then
        Exit Property
    End If
    If Me.NombreIcono = "" Then
        Exit Property
    End If
    m_sURLDIrectorioIconosUsuario = Me.URLDirectorioRecursosUsuario & "iconos\"
    
    URLDIrectorioIconosUsuario = m_sURLDIrectorioIconosUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDIrectorioIconosUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDIrectorioIconosServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDIrectorioIconosServidor <> "" Then
        URLDIrectorioIconosServidor = m_sURLDIrectorioIconosServidor
        Exit Property
    End If
    If Me.URLDirectorioRecursosServidor = "" Then
        Exit Property
    End If
    If Me.NombreIcono = "" Then
        Exit Property
    End If
    m_sURLDIrectorioIconosServidor = Me.URLDirectorioRecursosServidor & "iconos\"
    
    URLDIrectorioIconosServidor = m_sURLDIrectorioIconosServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDIrectorioIconosServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLDIrectorioIconos() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDIrectorioIconos <> "" Then
        URLDIrectorioIconos = m_sURLDIrectorioIconos
        Exit Property
    End If
    m_sURLDIrectorioIconos = ResolverRutaActiva(Me.URLDIrectorioIconosUsuario, Me.URLDIrectorioIconosServidor)
    If Not fso.FolderExists(m_sURLDIrectorioIconos) Then
        Me.Error = "No es alcanzable el directorio: " & m_sURLDIrectorioIconos
        m_sURLDIrectorioIconos = ""
        Err.Raise 1000
    End If
    URLDIrectorioIconos = m_sURLDIrectorioIconos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDIrectorioIconos ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get URLIConoParaArbolEnUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIConoParaArbolEnUsuario <> "" Then
        URLIConoParaArbolEnUsuario = m_sURLIConoParaArbolEnUsuario
        Exit Property
    End If
    If Me.URLDIrectorioIconosUsuario = "" Then
        Exit Property
    End If
    If Me.NombreIconoParaArbol = "" Then
        Exit Property
    End If
    m_sURLIConoParaArbolEnUsuario = Me.URLDIrectorioIconosUsuario & Me.NombreIconoParaArbol
    
    URLIConoParaArbolEnUsuario = m_sURLIConoParaArbolEnUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIConoParaArbolEnUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIConoParaArbolEnServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLIConoParaArbolEnServidor <> "" Then
        URLIConoParaArbolEnServidor = m_sURLIConoParaArbolEnServidor
        Exit Property
    End If
    If Me.URLDIrectorioIconosServidor = "" Then
        Exit Property
    End If
    If Me.NombreIconoParaArbol = "" Then
        Exit Property
    End If
    m_sURLIConoParaArbolEnServidor = Me.URLDIrectorioIconosServidor & Me.NombreIconoParaArbol
    
    URLIConoParaArbolEnServidor = m_sURLIConoParaArbolEnServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIConoParaArbolEnServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLIconoParaArbol() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLIconoParaArbol <> "" Then
        URLIconoParaArbol = m_sURLIconoParaArbol
        Exit Property
    End If
    If m_ObjEntorno.URLDirIconosArbol = "" Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.NombreEjecutable = "" Then
        Exit Property
    End If
    m_sURLIconoParaArbol = m_ObjEntorno.URLDirIconosArbol & fso.GetBaseName(Me.NombreEjecutable) & ".ico"
    
    
    If Not fso.FileExists(m_sURLIconoParaArbol) Then
        Me.Error = "No es alcanzable el archivo: " & m_sURLIconoParaArbol
        m_sURLIconoParaArbol = ""
        Err.Raise 1000
    End If
    URLIconoParaArbol = m_sURLIconoParaArbol
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLIconoParaArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLEjecutableEnServidor() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLEjecutableEnServidor <> "" Then
        URLEjecutableEnServidor = m_sURLEjecutableEnServidor
        Exit Property
    End If
    If Me.URLDirectorioRecursosServidor = "" Then
        Exit Property
    End If
    If Me.NombreBase = "" Then
        Exit Property
    End If
    m_sURLEjecutableEnServidor = Me.URLDirectorioRecursosServidor & Me.NombreBase & ".accde"
    
    
    URLEjecutableEnServidor = m_sURLEjecutableEnServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLEjecutableEnServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get URLEjecutableUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLEjecutableUsuario <> "" Then
        URLEjecutableUsuario = m_sURLEjecutableUsuario
        Exit Property
    End If
    If Me.URLDirectorioAplicacionUsuario = "" Then
        Exit Property
    End If
    If Me.NombreBase = "" Then
        Exit Property
    End If
    m_sURLEjecutableUsuario = Me.URLDirectorioAplicacionUsuario & Me.NombreBase & ".accde"
    URLEjecutableUsuario = m_sURLEjecutableUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLEjecutableUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLEjecutable() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLEjecutable <> "" Then
        URLEjecutable = m_sURLEjecutable
        Exit Property
    End If
    m_sURLEjecutable = ResolverRutaActiva(Me.URLEjecutableUsuario, Me.URLEjecutableEnServidor)
    If Not fso.FileExists(m_sURLEjecutable) Then
        Me.Error = "No es alcanzable el archivo: " & m_sURLEjecutable
        m_sURLEjecutable = ""
        Err.Raise 1000
    End If
    URLEjecutable = m_sURLEjecutable
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLEjecutable ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Property Get URLDirectorioVideos() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioVideos <> "" Then
        URLDirectorioVideos = m_sURLDirectorioVideos
        Exit Property
    End If
    If Me.NombreAplicacion = "" Then
        Exit Property
    End If
    If m_ObjEntorno.URLDirVideos = "" Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sURLDirectorioVideos = m_ObjEntorno.URLDirVideos & Me.NombreAplicacion & "\"
    URLDirectorioVideos = m_sURLDirectorioVideos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.URLDirectorioVideos ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get VersionServidor() As String

    On Error GoTo errores
    Me.Error = ""
    'If Me.IDAplicacion = "51" Then Stop
    If Me.URLIniServidor = "" Then
        Exit Property
    End If
    m_sVersionServidor = LeerIni(Me.URLIniServidor, "VersionAplicacion", "", Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    VersionServidor = m_sVersionServidor
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.VersionServidor ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get VersionUsuario() As String

    On Error GoTo errores
    Me.Error = ""
    'If Me.IDAplicacion = "51" Then Stop
    If Me.URLIniUsuario = "" Then
        Exit Property
    End If
    m_sVersionUsuario = LeerIni(Me.URLIniUsuario, "VersionAplicacion", "", Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    VersionUsuario = m_sVersionUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.VersionUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get IDAplicacionCalculada() As String

    On Error GoTo errores
    Me.Error = ""
    m_sIDAplicacionCalculada = DameID("TbAplicaciones", "IDAplicacion", getdb(), Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    IDAplicacionCalculada = m_sIDAplicacionCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.IDAplicacionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ConIconoEnLanzaderaCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.ConIconoEnLanzadera = "Sí" Then
        m_eConIconoEnLanzaderaCalculado = EnumSino.Sí
    ElseIf Me.ConIconoEnLanzadera = "No" Then
        m_eConIconoEnLanzaderaCalculado = EnumSino.No
    End If
    ConIconoEnLanzaderaCalculado = m_eConIconoEnLanzaderaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.ConIconoEnLanzaderaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Property Get NombreIconoArbol() As String

    On Error GoTo errores
    Me.Error = ""
    
    If m_sNombreIconoArbol <> "" Then
         NombreIconoArbol = m_sNombreIconoArbol
        Exit Property
    End If
    If Me.URLIconoParaArbol = "" Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    
    m_sNombreIconoArbol = fso.GetBaseName(Me.URLIconoParaArbol)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    NombreIconoArbol = m_sNombreIconoArbol
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.NombreIconoArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get ColVideos() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColVideos Is Nothing Then
        Set ColVideos = m_objColVideos
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objColVideos = Constructor.getVideosDeAplicacion(Me.IDAplicacion, , Me.Error)
    Set ColVideos = m_objColVideos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.ColVideos ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Public Property Get ColVideosDeAyuda() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColVideosDeAyuda Is Nothing Then
        Set ColVideosDeAyuda = m_objColVideosDeAyuda
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objColVideosDeAyuda = Constructor.getVideosDeAplicacion(Me.IDAplicacion, EnumSino.Sí, Me.Error)
    Set ColVideosDeAyuda = m_objColVideosDeAyuda
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.ColVideosDeAyuda ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ColVideosDeEncuesta() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColVideosDeEncuesta Is Nothing Then
        Set ColVideosDeEncuesta = m_objColVideosDeEncuesta
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objColVideosDeEncuesta = Constructor.getVideosDeAplicacion(Me.IDAplicacion, EnumSino.No, Me.Error)
    Set ColVideosDeEncuesta = m_objColVideosDeEncuesta
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.ColVideosDeEncuesta ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get HayCambioDeVersion() As EnumSino
    
    Dim m_VersionUsuario As String
    Dim m_VersionServidor As String
    On Error GoTo errores
    Me.Error = ""
    
    m_VersionUsuario = Me.VersionUsuario
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    m_VersionServidor = Me.VersionServidor
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionUsuario <> m_VersionServidor Then
        m_eHayCambioDeVersion = EnumSino.Sí
    Else
        m_eHayCambioDeVersion = EnumSino.No
    End If
    
    HayCambioDeVersion = m_eHayCambioDeVersion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.HayCambioDeVersion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Property

Public Function Lanzar( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_URLIniLocal As String
    Dim m_URLIniRemoto As String
    
    Dim m_URLIconoLocal As String
    Dim m_URLIconoRemoto As String
    
    Dim m_URLbmpLocal As String
    Dim m_URLbmpRemoto As String
    
    Dim m_URLEjecutableRemoto As String
    Dim m_URLEjecutableLocal As String
    
    Dim m_HayCambioVersion As EnumSino
    Dim m_db As DAO.Database
    Dim wks As DAO.Workspace
    Dim m_ComandoEjecutable As String
   
    On Error GoTo errores
    If m_CorreoUsuarioEnLogin = "" Then
        If m_ObjUsuarioConectadoLogin Is Nothing Then
            p_Error = "No se sabe el usuario que ha hecho login"
            Err.Raise 1000
        End If
        m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
    End If
    If InStr(1, m_CorreoUsuarioEnLogin, "@") = 0 Then
        p_Error = "No se sabe el usuario que ha hecho login"
        Err.Raise 1000
    End If
    
    
    
    
    
    If Not fso.FolderExists(Me.URLDirectorioAplicacionUsuario) Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se ha podido obtener la carpeta del aplicativo en local del usuario"
        Err.Raise 1000
    End If
    If Not fso.FolderExists(Me.URLDirectorioAplicacionServidor) Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se ha podido obtener la carpeta del aplicativo en el servidor"
        Err.Raise 1000
    End If
    m_URLIniRemoto = Me.URLIniServidor
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se ha podido obtener el archivo de configuración del servidor"
        Err.Raise 1000
    End If
    m_URLIniLocal = Me.URLIniUsuario
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLEjecutableRemoto = Me.URLEjecutableEnServidor
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLEjecutableRemoto) Then
        p_Error = "No se ha podido obtener el ejecutable del servidor" & vbNewLine & m_URLEjecutableRemoto
        Err.Raise 1000
    End If
    m_URLEjecutableLocal = Me.URLEjecutableUsuario
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_URLIconoRemoto = Me.URLIconoEnServidor
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLIconoRemoto) Then
        p_Error = "No se ha podido obtener m_URLIconoRemoto" & vbNewLine & m_URLIconoRemoto
        Err.Raise 1000
    End If
    m_URLIconoLocal = Me.URLIconoEnUsuario
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLbmpRemoto = Me.URLBmpServidor
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLIconoRemoto) Then
        p_Error = "No se ha podido obtener m_URLbmpRemoto" & vbNewLine & m_URLbmpRemoto
        Err.Raise 1000
    End If
    m_URLbmpLocal = Me.URLBmpUsuario
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HayCambioVersion = Me.HayCambioDeVersion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HayCambioVersion = 0 Then
        p_Error = "No se ha podido saber si hay o no cambio de versión"
        Err.Raise 1000
    End If
    
    If m_HayCambioVersion = EnumSino.Sí Then
        
        fso.CopyFile m_URLEjecutableRemoto, m_URLEjecutableLocal, True
        fso.CopyFile m_URLIconoRemoto, m_URLIconoLocal, True
        fso.CopyFile m_URLbmpRemoto, m_URLbmpLocal, True
        fso.CopyFile m_URLIniRemoto, m_URLIniLocal, True
        Set wks = DBEngine.Workspaces(0)
        If Me.Pass <> "" Then
            Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False, "MS Access;PWD=" & Me.Pass)
        Else
            Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False)
        End If
        
        AlterarPropriedad Me.URLEjecutableUsuario, "AppIcon", dbText, m_URLIconoRemoto, m_db, Me.Pass, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Me.Pass <> "" Then
            m_db.NewPassword Me.Pass, ""
        End If
        m_db.Close
        Set m_db = Nothing
        wks.Close
        Set wks = Nothing
    Else
        If Not fso.FileExists(m_URLEjecutableLocal) Then
            fso.CopyFile m_URLEjecutableRemoto, m_URLEjecutableLocal, True
            Set wks = DBEngine.Workspaces(0)
            If Me.Pass <> "" Then
                Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False, "MS Access;PWD=" & Me.Pass)
            Else
                Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False)
            End If
            AlterarPropriedad Me.URLEjecutableUsuario, "AppIcon", dbText, m_URLIconoRemoto, m_db, Me.Pass, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.Pass <> "" Then
                m_db.NewPassword Me.Pass, ""
            End If
            m_db.Close
            Set m_db = Nothing
            wks.Close
            Set wks = Nothing
        End If
        If Not fso.FileExists(m_URLIconoLocal) Then
            fso.CopyFile m_URLIconoRemoto, m_URLIconoLocal, True
        End If
        If Not fso.FileExists(m_URLbmpLocal) Then
            fso.CopyFile m_URLbmpRemoto, m_URLbmpLocal, True
        End If
        If Not fso.FileExists(m_URLIniLocal) Then
            fso.CopyFile m_URLIniRemoto, m_URLIniLocal, True
        End If
    End If
    If Me.NombreAplicacion = "CONDOR" Or Me.NombreAplicacion = "No Conformidades Prueba" Then
        If GetConfig("Runtime") = True Then
            m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " /runtime"
        Else
            m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " "
        End If
    Else
        m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " "
    End If
    
    m_ComandoEjecutable = m_ComandoEjecutable & Chr(34) & m_URLEjecutableLocal & Chr(34) & " " & " /cmd " & m_CorreoUsuarioEnLogin
    EjecutarShelllanzar m_ComandoEjecutable, EnumSino.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoLogin.MantenerLanzaderaAbierta = False Then
        
        DoCmd.Quit acQuitSaveAll
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Lanzar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
    On Error Resume Next
    If Not m_db Is Nothing Then
        m_db.Close
        Set m_db = Nothing
    End If
    If Not wks Is Nothing Then
        wks.Close
        Set wks = Nothing
    End If
    
End Function

Function ColUsuariosUsandola(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim rutaACCDB As String
    Dim rutaLACCDB As String
    Dim bytes() As Byte
    Dim i As Long
    Dim fNum As Integer
    Dim tempStr As String
    Dim posiblesUsuarios As Collection
    Dim usuario As Variant
    Dim nombresValidos As Collection
    Dim result As String
    
    Dim limpio As String
    Dim nombreCompleto As Variant
    Dim m_Usuario As usuario
    
    On Error GoTo errores
    rutaACCDB = Me.URLDirectorioAplicacion & Me.NombreArchivoDatos
    If Dir(rutaACCDB) = "" Then
        p_Error = "Archivo de la tabla de la aplicación no encontrado."
        Err.Raise 1000
    End If
    ' Verificar extensión .accdb
    If LCase(Right(rutaACCDB, 6)) <> ".accdb" Then
        p_Error = "Archivo .accdb no válido."
        Err.Raise 1000
    End If

    ' Obtener ruta del archivo .laccdb
    rutaLACCDB = Left(rutaACCDB, Len(rutaACCDB) - 6) & ".laccdb"
    If Dir(rutaLACCDB) = "" Then
        Exit Function
    End If

    ' Leer archivo binario
    fNum = FreeFile
    Open rutaLACCDB For Binary Access Read As #fNum
        ReDim bytes(LOF(fNum) - 1)
        Get #fNum, , bytes
    Close #fNum

    Set posiblesUsuarios = New Collection
    tempStr = ""

    ' Extraer cadenas ASCII legibles
    For i = 0 To UBound(bytes)
        If bytes(i) >= 32 And bytes(i) <= 126 Then
            tempStr = tempStr & Chr$(bytes(i))
        Else
            If Len(tempStr) >= 3 Then
                On Error Resume Next
                posiblesUsuarios.Add tempStr, tempStr
                On Error GoTo 0
            End If
            tempStr = ""
        End If
    Next i
    
    ' Incluir última cadena posible
    If Len(tempStr) >= 3 Then
        On Error Resume Next
        posiblesUsuarios.Add tempStr, tempStr
        On Error GoTo 0
    End If
    If posiblesUsuarios.Count = 0 Then
        Exit Function
    End If
   
   
    

    ' Buscar nombre completo correspondiente a cada usuario red
    For Each usuario In posiblesUsuarios
        limpio = Trim(CStr(usuario))
        If limpio <> "" Then
            Set m_Usuario = Constructor.getUsuario(p_UsuarioRed:=Replace(limpio, "'", "''"), p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_Usuario Is Nothing Then
                If ColUsuariosUsandola Is Nothing Then
                    Set ColUsuariosUsandola = New Scripting.Dictionary
                    ColUsuariosUsandola.CompareMode = TextCompare
                End If
                If Not ColUsuariosUsandola.Exists(m_Usuario.UsuarioRed) Then
                    ColUsuariosUsandola.Add m_Usuario.UsuarioRed, m_Usuario
                End If
            End If
            
            Set m_Usuario = Nothing
        End If
    Next usuario

   
  

    
Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Lanzar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function CadenaNombreUsuariosConectados(Optional ByRef p_Error As String) As String
    
    
    
    Dim m_Usuario As usuario
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_Cadena As String
    
    On Error GoTo errores
    
    Set m_Col = Me.ColUsuariosUsandola(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        CadenaNombreUsuariosConectados = "Nadie está usando la aplicación"
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        Set m_Usuario = m_Col(m_ID)
        
        If m_Cadena = "" Then
            m_Cadena = m_Usuario.Nombre
        Else
            m_Cadena = m_Cadena & vbNewLine & m_Usuario.Nombre
        End If
        Set m_Usuario = Nothing
    Next
    CadenaNombreUsuariosConectados = m_Cadena
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CadenaNombreUsuariosConectados ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
   
End Function
Public Function CadenaCorreoUsuariosConectados(Optional ByRef p_Error As String) As String
    
    
    
    Dim m_Usuario As usuario
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_Cadena As String
    
    On Error GoTo errores
    
    Set m_Col = Me.ColUsuariosUsandola(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        CadenaCorreoUsuariosConectados = "Nadie está usando la aplicación"
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        Set m_Usuario = m_Col(m_ID)
        
        If m_Cadena = "" Then
            m_Cadena = m_Usuario.CorreoUsuario
        Else
            m_Cadena = m_Cadena & vbNewLine & m_Usuario.CorreoUsuario
        End If
        Set m_Usuario = Nothing
    Next
    CadenaCorreoUsuariosConectados = m_Cadena
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CadenaCorreoUsuariosConectados ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
   
End Function
Private Function LeerIni( _
                        p_URL As String, _
                        key As String, _
                        Optional Default As Variant, _
                        Optional ByRef p_Error As String _
                        ) As String
   
    
    Dim Len_Value As Long
    Dim bufer As String * 256
    Dim m_nombreBase As String
    
    On Error GoTo errores
    If Not fso.FileExists(p_URL) Then
        p_Error = "No se alcanza el archivo ini: " & vbNewLine & p_URL
        Err.Raise 1000
    End If
    If Me.NombreBase = "" Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se ha podido obtener el nombre base del aplicativo"
        Err.Raise 1000
    End If
    
    Len_Value = GetPrivateProfileString(Me.NombreBase, _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         p_URL)
          
    LeerIni = Left$(bufer, CLng(Len_Value))
    
    Exit Function
errores:
    If p_Error <> "" Then
        p_Error = "El método LeerIni ha devuelto el error: " & Err.Description
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
    
    If Me.IDAplicacion <> "" Then
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "' AND IDAplicacion<>" & Me.IDAplicacion & ";"
    Else
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "';"
    End If
    m_SQL = "SELECT " & p_NombreCampo & " " & _
            "FROM TbAplicaciones " & _
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
        p_Error = "El método Aplicacion.CampoRepetido ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Function Registrar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As String
    Dim m_CampoRepetido As EnumSino
    
    On Error GoTo errores
   
    
    If Me.NombreAplicacion = "" Then
        p_Error = "Ha de indicar el NombreAplicacion"
        Err.Raise 1000
    End If
    m_CampoRepetido = CampoRepetido("NombreAplicacion", Me.NombreAplicacion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos un nombre idéntico"
        Err.Raise 1000
    End If
    If Me.NombreCorto = "" Then
        p_Error = "Ha de indicar el NombreCorto"
        Err.Raise 1000
    End If
    m_CampoRepetido = CampoRepetido("NombreCorto", Me.NombreCorto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe en la base de datos un nombre corto idéntico"
        Err.Raise 1000
    End If
    If Me.NombreEjecutable = "" Then
        p_Error = "Se ha de indicar NombreEjecutable"
        Err.Raise 1000
    End If
    If Me.NombreArchivoDatos = "" Then
        p_Error = "Se ha de indicar la NombreArchivoDatos"
        Err.Raise 1000
    End If
    If Me.NombreCarpetaTemporal = "" Then
        p_Error = "Se ha de indicar la NombreCarpetaTemporal"
        Err.Raise 1000
    End If
    If Me.NombreCarpeta = "" Then
        p_Error = "Se ha de indicar la NombreCarpeta"
        Err.Raise 1000
    End If
    If Me.NombreIcono = "" Then
        p_Error = "Se ha de indicar la NombreIcono"
        Err.Raise 1000
    End If
    If Me.NombreIconoLanzadera = "" Then
        p_Error = "Se ha de indicar la NombreIconoLanzadera"
        Err.Raise 1000
    End If
    If Me.NombreIconoParaArbol = "" Then
        p_Error = "Se ha de indicar la NombreIconoParaArbol"
        Err.Raise 1000
    End If
    If Me.EjecucionEnOficina <> "Sí" And Me.EjecucionEnOficina <> "No" Then
        p_Error = "Se ha de indicar la EjecucionEnOficina"
        Err.Raise 1000
    End If
    If Me.NombreDirectorioIconos = "" Then
        p_Error = "Se ha de indicar la NombreDirectorioIconos"
        Err.Raise 1000
    End If
    If Me.NombreDirectorioRecursos = "" Then
        p_Error = "Se ha de indicar la NombreDirectorioRecursos"
        Err.Raise 1000
    End If
    If Me.EnPruebas <> "Sí" And Me.EnPruebas <> "No" Then
        p_Error = "Se ha de indicar la EnPruebas"
        Err.Raise 1000
    End If
    If Me.ConIconoEnLanzadera <> "Sí" And Me.ConIconoEnLanzadera <> "No" Then
        p_Error = "Se ha de indicar la ConIconoEnLanzadera"
        Err.Raise 1000
    End If
    
    If Me.IDAplicacion = "" Then
        m_SQL = "TbAplicaciones"
    Else
        m_SQL = "SELECT TbAplicaciones.* " & _
            "FROM TbAplicaciones " & _
            "WHERE IDAplicacion=" & Me.IDAplicacion & ";"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                Me.IDAplicacion = Me.IDAplicacionCalculada
                p_Error = Me.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Nz(Me.NombreAplicacion, "") <> "" Then
                    .Fields("NombreAplicacion") = Me.NombreAplicacion
                End If
                If Nz(Me.NombreCorto, "") <> "" Then
                    .Fields("NombreCorto") = Me.NombreCorto
                End If
                If Nz(Me.NombreEjecutable, "") <> "" Then
                    .Fields("NombreEjecutable") = Me.NombreEjecutable
                End If
                If Nz(Me.NombreArchivoDatos, "") <> "" Then
                    .Fields("NombreArchivoDatos") = Me.NombreArchivoDatos
                End If
                If Nz(Me.Pass, "") <> "" Then
                    .Fields("Pass") = Me.Pass
                End If
                If Nz(Me.NombreCarpetaTemporal, "") <> "" Then
                    .Fields("NombreCarpetaTemporal") = Me.NombreCarpetaTemporal
                End If
                If Nz(Me.NombreCarpeta, "") <> "" Then
                    .Fields("NombreCarpeta") = Me.NombreCarpeta
                End If
                If Nz(Me.NombreIcono, "") <> "" Then
                    .Fields("NombreIcono") = Me.NombreIcono
                End If
                If Nz(Me.NombreIconoLanzadera, "") <> "" Then
                    .Fields("NombreIconoLanzadera") = Me.NombreIconoLanzadera
                End If
                If Nz(Me.NombreIconoParaArbol, "") <> "" Then
                    .Fields("NombreIconoParaArbol") = Me.NombreIconoParaArbol
                End If
                If Nz(Me.EjecucionEnOficina, "") <> "" Then
                    .Fields("EjecucionEnOficina") = Me.EjecucionEnOficina
                End If
                If Nz(Me.NombreCarpetaDocumentacion, "") <> "" Then
                    .Fields("NombreCarpetaDocumentacion") = Me.NombreCarpetaDocumentacion
                End If
                If Nz(Me.NombreDirectorioIconos, "") <> "" Then
                    .Fields("NombreDirectorioIconos") = Me.NombreDirectorioIconos
                End If
                If Nz(Me.NombreDirectorioAyuda, "") <> "" Then
                    .Fields("NombreDirectorioAyuda") = Me.NombreDirectorioAyuda
                End If
                If Nz(Me.NombreDirectorioRecursos, "") <> "" Then
                    .Fields("NombreDirectorioRecursos") = Me.NombreDirectorioRecursos
                End If
                If Nz(Me.EnPruebas, "") <> "" Then
                    .Fields("EnPruebas") = Me.EnPruebas
                End If
                If Nz(Me.ConIconoEnLanzadera, "") <> "" Then
                    .Fields("ConIconoEnLanzadera") = Me.ConIconoEnLanzadera
                End If
            .Update
            
        Else
            .Edit
                If Nz(Me.NombreAplicacion, "") <> "" Then
                    .Fields("NombreAplicacion") = Me.NombreAplicacion
                Else
                    .Fields("NombreAplicacion") = Null
                End If
                If Nz(Me.NombreCorto, "") <> "" Then
                    .Fields("NombreCorto") = Me.NombreCorto
                Else
                    .Fields("NombreCorto") = Null
                End If
                If Nz(Me.NombreEjecutable, "") <> "" Then
                    .Fields("NombreEjecutable") = Me.NombreEjecutable
                Else
                    .Fields("NombreEjecutable") = Null
                End If
                If Nz(Me.NombreArchivoDatos, "") <> "" Then
                    .Fields("NombreArchivoDatos") = Me.NombreArchivoDatos
                Else
                    .Fields("NombreArchivoDatos") = Null
                End If
                If Nz(Me.Pass, "") <> "" Then
                    .Fields("Pass") = Me.Pass
                Else
                    .Fields("Pass") = Null
                End If
                If Nz(Me.NombreCarpetaTemporal, "") <> "" Then
                    .Fields("NombreCarpetaTemporal") = Me.NombreCarpetaTemporal
                Else
                    .Fields("NombreCarpetaTemporal") = Null
                End If
                If Nz(Me.NombreCarpeta, "") <> "" Then
                    .Fields("NombreCarpeta") = Me.NombreCarpeta
                Else
                    .Fields("NombreCarpeta") = Null
                End If
                If Nz(Me.NombreIcono, "") <> "" Then
                    .Fields("NombreIcono") = Me.NombreIcono
                Else
                    .Fields("NombreIcono") = Null
                End If
                If Nz(Me.NombreIconoLanzadera, "") <> "" Then
                    .Fields("NombreIconoLanzadera") = Me.NombreIconoLanzadera
                Else
                    .Fields("NombreIconoLanzadera") = Null
                End If
                If Nz(Me.NombreIconoParaArbol, "") <> "" Then
                    .Fields("NombreIconoParaArbol") = Me.NombreIconoParaArbol
                Else
                    .Fields("NombreIconoParaArbol") = Null
                End If
                If Nz(Me.EjecucionEnOficina, "") <> "" Then
                    .Fields("EjecucionEnOficina") = Me.EjecucionEnOficina
                Else
                    .Fields("EjecucionEnOficina") = Null
                End If
                If Nz(Me.NombreCarpetaDocumentacion, "") <> "" Then
                    .Fields("NombreCarpetaDocumentacion") = Me.NombreCarpetaDocumentacion
                Else
                    .Fields("NombreCarpetaDocumentacion") = Null
                End If
                If Nz(Me.NombreDirectorioIconos, "") <> "" Then
                    .Fields("NombreDirectorioIconos") = Me.NombreDirectorioIconos
                Else
                    .Fields("NombreDirectorioIconos") = Null
                End If
                If Nz(Me.NombreDirectorioAyuda, "") <> "" Then
                    .Fields("NombreDirectorioAyuda") = Me.NombreDirectorioAyuda
                Else
                    .Fields("NombreDirectorioAyuda") = Null
                End If
                If Nz(Me.NombreDirectorioRecursos, "") <> "" Then
                    .Fields("NombreDirectorioRecursos") = Me.NombreDirectorioRecursos
                Else
                    .Fields("NombreDirectorioRecursos") = Null
                End If
                If Nz(Me.EnPruebas, "") <> "" Then
                    .Fields("EnPruebas") = Me.EnPruebas
                Else
                    .Fields("EnPruebas") = Null
                End If
                If Nz(Me.ConIconoEnLanzadera, "") <> "" Then
                    .Fields("ConIconoEnLanzadera") = Me.ConIconoEnLanzadera
                Else
                    .Fields("ConIconoEnLanzadera") = Null
                End If
                
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Function Eliminar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As String
    Dim m_CampoRepetido As EnumSino
    
    On Error GoTo errores
   
    
    If Me.IDAplicacion = "" Then
        p_Error = "Ha de indicar el IDAplicacion"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbAplicaciones.* " & _
            "FROM TbAplicaciones " & _
            "WHERE IDAplicacion=" & Me.IDAplicacion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No hay registrado ninguna aplicación con ese ID"
            Err.Raise 1000
            .Delete
                
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
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
    Set tdf = getdb().TableDefs("TbAplicaciones")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            If fld.Name = "NombreBoton" Or fld.Name = "NombreFuncionPublicacion" Or fld.Name = "Titulo" Then
                GoTo siguiente
            End If
            Set m_objColCampos = New Collection
siguiente:
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo Aplicacion.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreAplicacion" Then
        getPropiedad = Me.NombreAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCorto" Then
        getPropiedad = Me.NombreCorto
        Exit Function
    End If
    If m_NombrePropiedad = "NombreEjecutable" Then
        getPropiedad = Me.NombreEjecutable
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoDatos" Then
        getPropiedad = Me.NombreArchivoDatos
        Exit Function
    End If
    If m_NombrePropiedad = "Pass" Then
        getPropiedad = Me.Pass
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCarpeta" Then
        getPropiedad = Me.NombreCarpeta
        Exit Function
    End If
   
    If m_NombrePropiedad = "NombreCarpetaTemporal" Then
        getPropiedad = Me.NombreCarpetaTemporal
        Exit Function
    End If
    
    If m_NombrePropiedad = "NombreIcono" Then
        getPropiedad = Me.NombreIcono
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIconoLanzadera" Then
        getPropiedad = Me.NombreIconoLanzadera
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIconoParaArbol" Then
        getPropiedad = Me.NombreIconoParaArbol
        Exit Function
    End If
    If m_NombrePropiedad = "EjecucionEnOficina" Then
        getPropiedad = Me.EjecucionEnOficina
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCarpetaDocumentacion" Then
        getPropiedad = Me.NombreCarpetaDocumentacion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioIconos" Then
        getPropiedad = Me.NombreDirectorioIconos
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioAyuda" Then
        getPropiedad = Me.NombreDirectorioAyuda
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioRecursos" Then
        getPropiedad = Me.NombreDirectorioRecursos
        Exit Function
    End If
    
    If m_NombrePropiedad = "EnPruebas" Then
        getPropiedad = Me.EnPruebas
        Exit Function
    End If
    If m_NombrePropiedad = "ConIconoEnLanzadera" Then
        getPropiedad = Me.ConIconoEnLanzadera
        Exit Function
    End If
    If m_NombrePropiedad = "Comando" Then
        getPropiedad = Me.Comando
        Exit Function
    End If
    If m_NombrePropiedad = "NombreFuncionPublicacion" Then
        getPropiedad = Me.NombreFuncionPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "TituloAplicacion" Then
        getPropiedad = Me.TituloAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "URLDIrectorioIconoAplicacion" Then
        getPropiedad = Me.URLDIrectorioIconoAplicacion
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Aplicacion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Aplicacion.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreAplicacion" Then
        Me.NombreAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCorto" Then
        Me.NombreCorto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreEjecutable" Then
        Me.NombreEjecutable = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoDatos" Then
        Me.NombreArchivoDatos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Pass" Then
        Me.Pass = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCarpeta" Then
        Me.NombreCarpeta = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "NombreCarpetaTemporal" Then
        Me.NombreCarpetaTemporal = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "NombreIcono" Then
        Me.NombreIcono = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIconoLanzadera" Then
        Me.NombreIconoLanzadera = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIconoParaArbol" Then
        Me.NombreIconoParaArbol = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EjecucionEnOficina" Then
        Me.EjecucionEnOficina = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCarpetaDocumentacion" Then
        Me.NombreCarpetaDocumentacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioIconos" Then
        Me.NombreDirectorioIconos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioAyuda" Then
        Me.NombreDirectorioAyuda = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreDirectorioRecursos" Then
        Me.NombreDirectorioRecursos = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "EnPruebas" Then
        Me.EnPruebas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ConIconoEnLanzadera" Then
        Me.ConIconoEnLanzadera = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Comando" Then
        Me.Comando = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreFuncionPublicacion" Then
        Me.NombreFuncionPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TituloAplicacion" Then
        Me.TituloAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLDIrectorioIconoAplicacion" Then
        Me.URLDIrectorioIconoAplicacion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Aplicacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Aplicacion.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function









