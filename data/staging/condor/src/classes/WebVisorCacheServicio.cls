VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebVisorCacheServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' SERVICIO: WebVisorCacheServicio.cls
' DESCRIPCIÓN: Gestiona el caché de HTML estático para el WebVisor.
'              Carga, guarda y verifica archivos HTML cacheados en servidor.
' ARQUITECTURA: Servicio (lógica de negocio)
' ==========================================================================

Private fso As Object

Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
End Sub

' --------------------------------------------------------------------------
' MÉTODO: ObtenerHTMLCacheable
' DESCRIPCIÓN: Obtiene HTML del caché o genera dinámicamente.
'              Si no existe en caché, lo guarda para futuras cargas.
' --------------------------------------------------------------------------
Public Function ObtenerHTMLCacheable(ByVal tipoHTML As String, ByVal htmlGenerado As String) As String
    On Error GoTo Errores
    
    Debug.Print "[Cache] ===== ObtenerHTMLCacheable START ====="
    Debug.Print "[Cache] tipoHTML: " & tipoHTML
    Debug.Print "[Cache] htmlGenerado length: " & Len(htmlGenerado)
    
    ObtenerHTMLCacheable = htmlGenerado
    
    Dim usarCache As Boolean
    usarCache = ObtenerModoCarga()
    Debug.Print "[Cache] ObtenerModoCarga(): " & usarCache
    
    If Not usarCache Then
        Debug.Print "[Cache] No usar cache, retornando htmlGenerado"
        Exit Function
    End If
    
    Dim htmlCache As String
    Debug.Print "[Cache] Leyendo HTML cacheado..."
    htmlCache = LeerHTMLCacheado(tipoHTML)
    Debug.Print "[Cache] htmlCache length: " & Len(htmlCache)
    
    If htmlCache <> "" Then
        Debug.Print "[Cache] Usando HTML cacheado"
        ObtenerHTMLCacheable = htmlCache
    Else
        Debug.Print "[Cache] Guardando HTML en cache"
        GuardarHTMLEnCache tipoHTML, htmlGenerado
    End If
    
    Debug.Print "[Cache] ===== ObtenerHTMLCacheable END ====="
    
    Exit Function
Errores:
    ObtenerHTMLCacheable = htmlGenerado
End Function

' --------------------------------------------------------------------------
' MÉTODO: ObtenerModoCarga
' DESCRIPCIÓN: Verifica si está habilitado el modo de carga por archivo.
' --------------------------------------------------------------------------
Private Function ObtenerModoCarga() As Boolean
    On Error Resume Next
    ObtenerModoCarga = False
    If Application.TempVars("UsarRenderizadoFichero") = "Sí" Then
        ObtenerModoCarga = True
    End If
End Function

' --------------------------------------------------------------------------
' MÉTODO: LeerHTMLCacheado
' DESCRIPCIÓN: Lee un archivo HTML desde el caché del servidor.
' --------------------------------------------------------------------------
Private Function LeerHTMLCacheado(ByVal tipoHTML As String) As String
    On Error Resume Next
    LeerHTMLCacheado = ""
    
    Debug.Print "[Cache-Leer] tipoHTML: " & tipoHTML
    
    Dim ruta As String
    ruta = ObtenerRutaHTML(tipoHTML)
    Debug.Print "[Cache-Leer] ruta: " & ruta
    
    If ruta = "" Then
        Debug.Print "[Cache-Leer] ruta vacia"
        Exit Function
    End If
    
    If Not fso.FileExists(ruta) Then
        Debug.Print "[Cache-Leer] archivo NO existe: " & ruta
        Exit Function
    End If
    
    Debug.Print "[Cache-Leer] Leyendo archivo..."
    Dim ts As Object
    Set ts = fso.OpenTextFile(ruta, 1)
    LeerHTMLCacheado = ts.ReadAll
    ts.Close
    Debug.Print "[Cache-Leer] Leido, length: " & Len(LeerHTMLCacheado)
End Function

' --------------------------------------------------------------------------
' MÉTODO: GuardarHTMLEnCache
' DESCRIPCIÓN: Guarda un HTML en la carpeta de caché del servidor.
' --------------------------------------------------------------------------
Private Sub GuardarHTMLEnCache(ByVal tipoHTML As String, ByVal htmlContent As String)
    On Error Resume Next
    
    Dim ruta As String
    Dim carpeta As String
    Dim ts As Object
    
    ruta = ObtenerRutaHTML(tipoHTML)
    If ruta = "" Then Exit Sub
    
    carpeta = ObtenerDirectorioHTML
    If carpeta = "" Then Exit Sub
    
    If Not fso.FolderExists(carpeta) Then
        fso.CreateFolder carpeta
    End If
    
    Set ts = fso.OpenTextFile(ruta, 2, True, 0)
    ts.Write htmlContent
    ts.Close
    Set ts = Nothing
End Sub

' --------------------------------------------------------------------------
' MÉTODO: ExisteCache
' DESCRIPCIÓN: Verifica si existe el archivo cacheado para un tipo de HTML.
' --------------------------------------------------------------------------
Public Function ExisteCache(ByVal tipoHTML As String) As Boolean
    On Error Resume Next
    ExisteCache = False
    
    Dim ruta As String
    ruta = ObtenerRutaHTML(tipoHTML)
    
    If ruta <> "" Then
        ExisteCache = fso.FileExists(ruta)
    End If
End Function

' --------------------------------------------------------------------------
' MÉTODO: ObtenerRutaHTML
' DESCRIPCIÓN: Retorna la ruta completa del archivo HTML cacheado.
' --------------------------------------------------------------------------
Private Function ObtenerRutaHTML(ByVal tipoHTML As String) As String
    On Error Resume Next
    ObtenerRutaHTML = ""
    
    Debug.Print "[ObtenerRutaHTML] tipoHTML: " & tipoHTML
    
    Dim entorno As New entorno
    Dim directorio As String
    
    directorio = ObtenerDirectorioHTML
    Debug.Print "[ObtenerRutaHTML] directorio: " & directorio
    If directorio = "" Then
        Debug.Print "[ObtenerRutaHTML] directorio vacio, sale"
        Exit Function
    End If
    
    Select Case tipoHTML
        Case "CAPTURA_FECHA_RAC"
            ObtenerRutaHTML = directorio & "captura_fecha_rac.html"
        Case "CAPTURA_COMENTARIOS_RAC"
            ObtenerRutaHTML = directorio & "captura_comentarios_rac.html"
        Case "CAPTURA_RECHAZO_CALIDAD"
            ObtenerRutaHTML = directorio & "captura_rechazo_calidad.html"
    End Select
    
    Debug.Print "[ObtenerRutaHTML] ruta final: " & ObtenerRutaHTML
End Function

' --------------------------------------------------------------------------
' MÉTODO: ObtenerDirectorioHTML
' DESCRIPCIÓN: Retorna la ruta del directorio de caché HTML.
' --------------------------------------------------------------------------
Private Function ObtenerDirectorioHTML() As String
    On Error Resume Next
    ObtenerDirectorioHTML = ""
    
    Debug.Print "[ObtenerDirectorioHTML] START"
    
    Dim entorno As New entorno
    ObtenerDirectorioHTML = entorno.URLDirectorioHTML
    Debug.Print "[ObtenerDirectorioHTML] entorno.URLDirectorioHTML: " & ObtenerDirectorioHTML
End Function

