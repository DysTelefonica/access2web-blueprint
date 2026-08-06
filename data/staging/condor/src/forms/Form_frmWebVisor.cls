

Option Compare Database
Option Explicit

' ==========================================================================
' FORMULARIO: frmWebVisor
' DESCRIPCIÓN: Contenedor genérico para visualización de contenido HTML mediante Edge.
' ARQUITECTURA: MVVM (Vista)
'
' NOTAS TÉCNICAS (EdgeBrowserControl / WebView2):
'   - El control REQUIERE ser inicializado con un archivo local (https://msaccess/)
'     antes de poder cargar cualquier otro contenido local o dinámico.
'     Sin esta inicialización, cualquier Navigate a https://msaccess/ resulta
'     en pantalla en blanco con doble DocumentComplete (bug "No Mixed Content").
'   - El archivo _blank.html actúa como "cebador" del contexto de seguridad local.
'   - document.write() con ExecuteJavascript SÍ funciona una vez inicializado.
'   - Ref: https://www.devhut.net/access-modern-web-browser-control-dynamic-content/
' ==========================================================================

' --- Estado interno ---
Private m_htmlPendiente     As String   ' HTML a renderizar tras la inicialización
Private m_tituloPendiente   As String   ' Título pendiente de aplicar
Private m_controlInicializado As Boolean ' True tras cargar _blank.html
Private m_ultimoComando     As String
Private m_rutaArchivoTemporal As String

' --- Ruta del HTML de inicialización ---
Private Const BLANK_FILENAME As String = "_blank.html"

' ==========================================================================
' EVENTOS DE FORMULARIO
' ==========================================================================

Private Sub Form_Load()
    On Error Resume Next
    Ajustar Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.Caption = "CONDOR - Visor Web"
    m_controlInicializado = False
    m_htmlPendiente = ""
    m_tituloPendiente = ""
    
    ' CRÍTICO: inicializar el control con un archivo local antes de cualquier
    ' carga dinámica. Sin esto el control queda en contexto "web" y rechaza
    ' los archivos locales con doble DocumentComplete + pantalla en blanco.
    InicializarControl
End Sub

Private Sub Form_Close()
    On Error Resume Next
    If m_rutaArchivoTemporal <> "" Then
        Kill m_rutaArchivoTemporal
        m_rutaArchivoTemporal = ""
    End If
End Sub

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub InicializarControl()
    m_controlInicializado = False
    Dim rutaBlank As String
    rutaBlank = ObtenerOCrearBlank()
    If rutaBlank <> "" Then
        Me.webVisor.Navigate "https://msaccess/" & rutaBlank
    Else
        Me.webVisor.Navigate "about:blank"
    End If
End Sub
' --------------------------------------------------------------------------
' FUNCIÓN: ObtenerOCrearBlank
' DESCRIPCIÓN: Devuelve la ruta local del _blank.html, creándolo si no existe.
'              Se guarda junto al .accdb para que la ruta sea estable.
' --------------------------------------------------------------------------
Private Function ObtenerOCrearBlank() As String
    On Error GoTo Errores
    
    ' CRÍTICO: CurrentProject.Path puede tener espacios (bug documentado del control).
    ' Usar %TEMP%\CONDOR\ que es una ruta corta y sin espacios.
    Dim carpeta As String
    carpeta = Environ("TEMP") & "\CONDOR"
    
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta
    
    Dim rutaBlank As String
    rutaBlank = carpeta & "\" & BLANK_FILENAME
    
    If Dir(rutaBlank) = "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim ts As Object
        Set ts = fso.OpenTextFile(rutaBlank, 2, True, 0)
        ts.Write "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body></body></html>"
        ts.Close
        Set ts = Nothing
        Set fso = Nothing
    End If
    
    ObtenerOCrearBlank = rutaBlank
    Exit Function
Errores:
    ObtenerOCrearBlank = ""
End Function


' ==========================================================================
' API PÚBLICA
' ==========================================================================

' --------------------------------------------------------------------------
' MÉTODO: CargarContenido
' DESCRIPCIÓN: Carga una cadena HTML en el control Edge.
'              Si el control aún no está inicializado (Form_Open aún pendiente),
'              guarda el HTML en m_htmlPendiente y se aplica en DocumentComplete.
'              Estrategia de renderizado:
'                1) Modo archivo (TempVars "UsarRenderizadoFichero" = "Sí"):
'                   Escribe el HTML a un .html temporal y navega con https://msaccess/
'                2) Modo JS (por defecto):
'                   Inyecta via document.write() en DocumentComplete
' --------------------------------------------------------------------------
Public Sub CargarContenido(ByVal html As String, _
                            Optional ByVal titulo As String = "Visor de Contenido", _
                            Optional ByVal tipoCacheable As String = "")
    On Error GoTo Errores
    
    ' --- Resolver caché de servidor ---
    If tipoCacheable <> "" Then
        Dim cache As New WebVisorCacheServicio
        html = cache.ObtenerHTMLCacheable(tipoCacheable, html)
    End If
    
    Me.lblTitulo.Caption = titulo
    Me.TimerInterval = 0
    
    ' --- Limpiar archivo temporal anterior ---
    On Error Resume Next
    If m_rutaArchivoTemporal <> "" Then
        Kill m_rutaArchivoTemporal
        m_rutaArchivoTemporal = ""
    End If
    On Error GoTo Errores
    
    ' --- Si el control aún no está listo, encolar y esperar ---
    If Not m_controlInicializado Then
        m_htmlPendiente = html
        m_tituloPendiente = titulo
        Exit Sub
    End If
    
    ' --- Renderizar ---
    RenderizarHTML html, titulo
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "Form_frmWebVisor.CargarContenido"
    errObj.Raise
End Sub

' ==========================================================================
' RENDERIZADO INTERNO
' ==========================================================================

' --------------------------------------------------------------------------
' MÉTODO: RenderizarHTML
' DESCRIPCIÓN: Decide entre modo archivo y modo JS, y ejecuta la carga.
'              SÓLO llamar cuando m_controlInicializado = True.
' --------------------------------------------------------------------------
Private Sub RenderizarHTML(ByVal html As String, ByVal titulo As String)
    On Error GoTo Errores
    
    If ObtenerModoCarga() Then
        RenderizarModoArchivo html
    Else
        RenderizarModoJS html
    End If
    
    Exit Sub
Errores:
End Sub

' --------------------------------------------------------------------------
' MÉTODO: RenderizarModoArchivo
' DESCRIPCIÓN: Escribe el HTML a un .html temporal y navega con https://msaccess/
'              El control ya está inicializado, por lo que la carga es directa.
' --------------------------------------------------------------------------
Private Sub RenderizarModoArchivo(ByVal html As String)
    On Error GoTo Errores
    
    Dim rutaTemp As String
    RenderizarAArchivo html, rutaTemp
    
    If rutaTemp = "" Then
        RenderizarModoJS html
        Exit Sub
    End If
    
    m_rutaArchivoTemporal = rutaTemp
    
    Dim urlNavigate As String
    urlNavigate = "https://msaccess/" & rutaTemp
    Me.webVisor.Navigate urlNavigate
    Exit Sub
Errores:
End Sub

' --------------------------------------------------------------------------
' MÉTODO: RenderizarModoJS
' DESCRIPCIÓN: Inyecta el HTML directamente via document.write() con ExecuteJavascript.
'              Funciona porque el control ya está en contexto local (_blank.html).
'              IMPORTANTE: eliminar todos los vbCrLf — requisito del control Edge.
' --------------------------------------------------------------------------
Private Sub RenderizarModoJS(ByVal html As String)
    On Error GoTo Errores
    
    ' Preparar el HTML para document.write:
    '   1. Eliminar saltos de línea (CRÍTICO para el control Edge)
    '   2. Escapar comillas simples
    '   3. NO escapar backslashes globalmente (rompe rutas)
    Dim jsContent As String
    jsContent = html
    jsContent = Replace(jsContent, vbCrLf, "")
    jsContent = Replace(jsContent, vbCr, "")
    jsContent = Replace(jsContent, vbLf, "")
    jsContent = Replace(jsContent, "'", "\'")
    
    Dim jsCmd As String
    jsCmd = "document.open(); document.write('" & jsContent & "'); document.close();"
    
    Me.webVisor.ExecuteJavascript jsCmd
    Exit Sub
Errores:
    Dim logServ As New LogServicio
    logServ.RegistrarError Err.Number, Err.description, "Form_frmWebVisor", "RenderizarModoJS", ""
End Sub

' --------------------------------------------------------------------------
' MÉTODO: RenderizarAArchivo
' DESCRIPCIÓN: Escribe el HTML a un archivo .html temporal en UTF-8.
' --------------------------------------------------------------------------
Public Sub RenderizarAArchivo(ByVal html As String, ByRef ruta As String)
    On Error GoTo Errores
    
    ruta = ""
    
    Dim carpetaTemp As String
    carpetaTemp = Environ("TEMP") & "\CONDOR\WebVisor"
    
    If Dir(carpetaTemp, vbDirectory) = "" Then
        MkDir carpetaTemp
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim nombreArchivo As String
    nombreArchivo = "vis_" & Format(Now, "yyyymmddhhnnss") & "_" & _
                    Mid(fso.GetBaseName(fso.GetTempName), 1, 8) & ".html"
    Set fso = Nothing
    
    ruta = carpetaTemp & "\" & nombreArchivo
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2           ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText html
    stream.SaveToFile ruta, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
    
    Exit Sub
Errores:
    ruta = ""
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    On Error GoTo 0
    Dim logServ As New LogServicio
    logServ.RegistrarError errNum, errDesc, "Form_frmWebVisor", "RenderizarAArchivo", ""
    Err.Raise errNum
End Sub

Private Sub webVisor_DocumentComplete(URL As Variant)
    On Error GoTo Errores
    
    Dim urlStr As String
    urlStr = LCase(Trim(CStr(URL)))
    
    If Not m_controlInicializado Then
        If urlStr = "about:blank" Or urlStr = "" Then
            m_controlInicializado = True
            
            If m_htmlPendiente <> "" Then
                Dim h As String: h = m_htmlPendiente
                Dim t As String: t = m_tituloPendiente
                m_htmlPendiente = ""
                m_tituloPendiente = ""
                If t <> "" Then Me.lblTitulo.Caption = t
                RenderizarModoJS h
            End If
        End If
        Exit Sub
    End If
    
    Exit Sub
Errores:
    Dim logServ As New LogServicio
    logServ.RegistrarError Err.Number, Err.description, "Form_frmWebVisor", "webVisor_DocumentComplete", ""
End Sub

' --------------------------------------------------------------------------
' EVENTO: webVisor_BeforeNavigate
' DESCRIPCIÓN: Intercepta los comandos VBA enviados desde el HTML via vba://
' --------------------------------------------------------------------------
Private Sub webVisor_BeforeNavigate(Cancel As Integer, URL As String)
    On Error GoTo Errores
    
    If Left(URL, 4) <> "vba:" Then GoTo LimpiarYSalir
    
    Cancel = True
    
    Dim mensaje As String
    Dim partes() As String
    mensaje = Mid(URL, 5)
    partes = Split(mensaje, ":")
    
    If UBound(partes) < 1 Then GoTo LimpiarYSalir
    
    Select Case partes(0)
        Case "DATOS_RAC"
            m_ultimoComando = partes(1)
            If CurrentProject.AllForms("frmGestionSolicitud").IsLoaded Then
                Forms("frmGestionSolicitud").RecibirDatosRAC m_ultimoComando
            End If
            
        Case "DATOS_RECHAZO"
            m_ultimoComando = partes(1)
            If CurrentProject.AllForms("frmGestionSolicitud").IsLoaded Then
                Forms("frmGestionSolicitud").RecibirDatosRechazoCalidad m_ultimoComando
            End If
            
        Case "ABRIR_ADJUNTO"
            Dim idAdjunto As Long
            idAdjunto = CLng(val(partes(1)))
            If idAdjunto > 0 Then
                Dim adjServ As AdjuntosServicio
                Dim adj As Adjunto
                Set adjServ = New AdjuntosServicio
                Set adj = adjServ.getAdjuntoPorID(idAdjunto)
                If Not adj Is Nothing Then
                    Dim entorno As New entorno
                    Dim rutaBase As String
                    rutaBase = entorno.URLDirectorioDocumentacion
                    If Right$(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"
                    Dim rutaAdj As String
                    rutaAdj = rutaBase & adj.nombreArchivo
                    If rutaAdj <> "" Then Call AbrirEnLocal(rutaAdj)
                End If
            End If
            
        Case "IMPORTAR_WORD_RAC"
            If CurrentProject.AllForms("frmGestionSolicitud").IsLoaded Then
                Forms("frmGestionSolicitud").RecibirImportacionWordRAC
            End If
    End Select
    
    If partes(0) <> "IMPORTAR_WORD_RAC" Then
        DoCmd.Close acForm, Me.name
    End If

LimpiarYSalir:
    Exit Sub
Errores:
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' HELPERS PRIVADOS
' ==========================================================================

' --------------------------------------------------------------------------
' FUNCIÓN: ObtenerModoCarga
' DESCRIPCIÓN: Lee el flag global UsarRenderizadoFichero.
'              "Sí" = archivo temporal, otro valor = inyección JS.
' --------------------------------------------------------------------------
Private Function ObtenerModoCarga() As Boolean
    On Error Resume Next
    ObtenerModoCarga = False
    Dim flagValue As String
    flagValue = Application.TempVars("UsarRenderizadoFichero")
    ObtenerModoCarga = (flagValue = "Sí")
End Function

' --------------------------------------------------------------------------
' FUNCIÓN: EncodeBase64
' DESCRIPCIÓN: Codifica texto a Base64 (reservado para uso futuro con data URIs)
' --------------------------------------------------------------------------
Private Function EncodeBase64(ByVal text As String) As String
    On Error GoTo Errores
    Dim bytes() As Byte
    bytes = StrConv(text, vbFromUnicode)
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    EncodeBase64 = Replace(node.text, vbLf, "")
    Set node = Nothing
    Set xml = Nothing
    Exit Function
Errores:
    EncodeBase64 = ""
    Dim logSvcB64 As LogServicio
    Set logSvcB64 = New LogServicio
    logSvcB64.RegistrarError Err.Number, Err.description, "Form_frmWebVisor", "EncodeBase64"
End Function



