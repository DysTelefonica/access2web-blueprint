VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InformeRiesgoPDFServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================================
' SERVICIO DE GENERACION DE INFORME PDF (Refactorizado)
' =============================================================================================
' Objetivo: Generacion robusta de PDF A4 con encabezados/pies fijos y soporte Portrait/Landscape.
' Estrategia:
' 1. Generacion de HTML con layout Flexbox para garantizar posicionamiento de Header/Footer.
' 2. Separacion de orientaciones: Se generan dos archivos PDF si es necesario (Principal + Anexo Inventario).
' 3. Uso de Microsoft Edge (Chromium) en modo headless para la conversion.
' =============================================================================================

Private m_HTML As String
Private m_PageCount As Long
Private m_Proyecto As Proyecto
Private m_ObjEntorno As Entorno

' Constantes de configuracion
Private Const PAGE_WIDTH_MM_PORTRAIT As Long = 210
Private Const PAGE_HEIGHT_MM_PORTRAIT As Long = 297
Private Const MARGIN_MM As Long = 0 ' Controlado por CSS interno
Private Const MAX_ROWS_PER_PAGE_PORTRAIT As Long = 14
Private Const MAX_ROWS_PER_PAGE_LANDSCAPE As Long = 16

Private Sub Class_Initialize()
    Set m_ObjEntorno = New Entorno
End Sub

' =============================================================================================
' API PUBLICA
' =============================================================================================

Public Sub GuardarComoPDF(ByVal p_Edicion As Edicion, ByVal p_RutaPDF As String, Optional ByRef p_Error As String)
    ' Metodo principal. Orquesta la generacion de 1 o 2 PDFs segun el contenido.
    ' Si hay inventario (Landscape), se genera como Anexo.
    
    Dim sRutaMain As String
    Dim sRutaAnexo As String
    Dim sHTML_Main As String
    Dim sHTML_Anexo As String
    Dim bGenerarAnexo As Boolean
    
    On Error GoTo errores
    p_Error = ""
    
    ' 1. Preparar Entorno y Datos
    Set m_Proyecto = p_Edicion.Proyecto
    If m_Proyecto Is Nothing Then
        Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    End If
    
    ' 2. Determinar Rutas
    sRutaMain = p_RutaPDF
    sRutaAnexo = Replace(p_RutaPDF, ".pdf", "_Anexo_Inventario.pdf")
    
    ' 3. Generar HTML Principal (Portrait)
    ' Incluye: Portada, Datos, Cuadro Control, Control Cambios, Fichas de Riesgo
    sHTML_Main = GenerarHTML_Principal(p_Edicion, Format(Date, "dd/mm/yyyy"), Format(Date, "dd/mm/yyyy"), p_Error)
    If p_Error <> "" Then Exit Sub
    
    ' 4. Generar HTML Anexo (Landscape) - Inventario
    ' Verificamos si hay riesgos para decidir si generar el anexo
    bGenerarAnexo = TieneRiesgos(p_Edicion)
    If bGenerarAnexo Then
        sHTML_Anexo = GenerarHTML_Inventario(p_Edicion, Format(Date, "dd/mm/yyyy"), Format(Date, "dd/mm/yyyy"), p_Error)
        If p_Error <> "" Then Exit Sub
    End If
    
    ' 5. Convertir a PDF
    ' Principal
    ConvertirHTMLaPDF sHTML_Main, sRutaMain, False, p_Error
    If p_Error <> "" Then Exit Sub
    
    ' Anexo
    If bGenerarAnexo Then
        ConvertirHTMLaPDF sHTML_Anexo, sRutaAnexo, True, p_Error
        If p_Error <> "" Then Exit Sub
    End If
    
    Exit Sub
    
errores:
    p_Error = "Error GuardarComoPDF: " & Err.Description
End Sub

' =============================================================================================
' GENERACION HTML (MOTOR)
' =============================================================================================

Private Function GenerarHTML_Principal( _
                                        ByVal p_Edicion As Edicion, ByVal p_FechaCierre As String, _
                                        ByVal p_FechaPublicacion As String, _
                                        ByRef p_Error As String) As String
    
    On Error GoTo errores
    p_Error = ""
    m_HTML = ""
    m_PageCount = 0
    
    Dim sTituloExp As String, sExpediente As String, sNombre As String, sJefe As String, sCliente As String, sCodDoc As String, sEdicion As String
    Dim m_EsPublicacion As Boolean
    Dim m_Contenido As String
    
    sTituloExp = Nz(m_Proyecto.TituloExpediente, "")
    sExpediente = Nz(m_Proyecto.Proyecto, "")
    sNombre = Nz(m_Proyecto.NombreProyecto, "")
    sJefe = Nz(p_Edicion.Elaborado, "")
    sCliente = Nz(m_Proyecto.Cliente, "")
    sCodDoc = Nz(m_Proyecto.CodigoDocumento, "")
    sEdicion = Nz(p_Edicion.Edicion, "")
    m_EsPublicacion = IsDate(p_FechaPublicacion)
    
    IniciarHTML False ' Portrait
    
    ' 1. Portada
    m_Contenido = ConstruirSeccionPortadaHTML(sTituloExp, sExpediente)
    
    AgregarPagina m_Contenido, False, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 2. Datos Generales
    m_Contenido = ConstruirSeccionDatosGeneralesHTML(sNombre, sExpediente, sJefe, sCliente, sCodDoc, sEdicion, p_FechaPublicacion, m_EsPublicacion)
    
    AgregarPagina m_Contenido, False, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 3. Cuadro de Control
    m_Contenido = ConstruirSeccionCuadroControlHTML_Local(p_Edicion, m_Proyecto, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    AgregarPagina m_Contenido, False, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    
    ' 4. Control de Cambios (Paginado)
    
    AgregarSeccionControlCambiosPaginada p_Edicion, m_Proyecto, p_Error
    If p_Error <> "" Then Err.Raise 1000
    ' 5. Fichas de Riesgo (Paginado - 1 por pagina)
    AgregarSeccionFichasRiesgoPaginada p_Edicion, m_Proyecto, p_FechaCierre, p_FechaPublicacion, p_Error
    If p_Error <> "" Then Err.Raise 1000
    FinalizarHTML
    GenerarHTML_Principal = m_HTML
    Exit Function
    
errores:
    p_Error = "Error GenerarHTML_Principal: " & Err.Description
End Function

Private Function GenerarHTML_Inventario(ByVal p_Edicion As Edicion, ByVal p_FechaCierre As String, ByVal p_FechaPublicacion As String, ByRef p_Error As String) As String
    m_HTML = ""
    m_PageCount = 0 ' Reiniciar contador para el anexo (o podriamos continuarlo si pasaramos el parametro)
    
    IniciarHTML True ' Landscape
    
    ' Inventario (Paginado Horizontal)
    AgregarSeccionInventarioPaginada p_Edicion, p_FechaCierre, p_FechaPublicacion, p_Error
    
    FinalizarHTML
    GenerarHTML_Inventario = m_HTML
End Function

Private Sub IniciarHTML(ByVal bLandscape As Boolean)
    m_HTML = "<!DOCTYPE html>" & vbCrLf
    m_HTML = m_HTML & "<html lang='es'>" & vbCrLf
    m_HTML = m_HTML & "<head>" & vbCrLf
    m_HTML = m_HTML & "    <meta charset='UTF-8'>" & vbCrLf
    m_HTML = m_HTML & "    <title>Informe PDF</title>" & vbCrLf
    m_HTML = m_HTML & GetEstilosCSS_PDF(bLandscape)
    m_HTML = m_HTML & "</head>" & vbCrLf
    m_HTML = m_HTML & "<body>" & vbCrLf
End Sub

Private Sub FinalizarHTML()
    m_HTML = m_HTML & "</body>" & vbCrLf
    m_HTML = m_HTML & "</html>"
End Sub

' =============================================================================================
' GESTION DE PAGINAS Y LAYOUT (FLEXBOX)
' =============================================================================================

Private Sub AgregarPagina(ByVal p_Contenido As String, ByVal p_EsLandscape As Boolean, Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""
    m_PageCount = m_PageCount + 1
    
    Dim sClass As String
    ' Nota: La orientacion base la define @page en el CSS, pero la clase ayuda si mezclamos (aunque aqui separamos archivos)
    If p_EsLandscape Then
        sClass = "pdf-page landscape-page"
    Else
        sClass = "pdf-page portrait-page"
    End If
    
    m_HTML = m_HTML & "<div class='" & sClass & "'>" & vbCrLf
    
    ' Header Fijo (Flex Item)
    m_HTML = m_HTML & "  <div class='header-wrapper'>" & GetHeaderHTML() & "</div>" & vbCrLf
    
    ' Contenido (Flex Item Grow)
    m_HTML = m_HTML & "  <div class='content-wrapper'>" & vbCrLf
    m_HTML = m_HTML & p_Contenido & vbCrLf
    m_HTML = m_HTML & "  </div>" & vbCrLf
    
    ' Footer Fijo (Flex Item)
    m_HTML = m_HTML & "  <div class='footer-wrapper'>" & GetFooterHTML(m_PageCount) & "</div>" & vbCrLf
    
    m_HTML = m_HTML & "</div>" & vbCrLf
    Exit Sub
    
errores:
    p_Error = "Error AgregarPagina: " & Err.Description
End Sub

Private Function GetEstilosCSS_PDF(ByVal bMainLandscape As Boolean) As String
    Dim css As String
    
    ' Definimos el tamano de pagina base para @page
    Dim sPageSize As String
    If bMainLandscape Then sPageSize = "landscape" Else sPageSize = "portrait"
    
    css = "<style>" & vbCrLf
    css = css & "* { box-sizing: border-box; }" & vbCrLf
    css = css & "body { margin: 0; padding: 0; background: #525659; font-family: 'Segoe UI', Arial, sans-serif; -webkit-print-color-adjust: exact; }" & vbCrLf
    
    ' Layout Flexbox por Pagina
    css = css & ".pdf-page { display: flex; flex-direction: column; page-break-after: always; background: white; margin: 20px auto; box-shadow: 0 0 10px rgba(0,0,0,0.5); overflow: hidden; }" & vbCrLf
    
    ' Dimensiones explicitas
    css = css & ".portrait-page { width: 210mm; height: 297mm; }" & vbCrLf
    css = css & ".landscape-page { width: 297mm; height: 210mm; }" & vbCrLf
    
    ' Elementos Flex
    css = css & ".header-wrapper { height: 25mm; flex: 0 0 auto; border-bottom: 2px solid #0066FF; padding: 0 15mm; display: flex; align-items: center; }" & vbCrLf
    css = css & ".footer-wrapper { height: 15mm; flex: 0 0 auto; border-top: 1px solid #ccc; display: flex; align-items: center; justify-content: center; font-size: 10px; color: #666; }" & vbCrLf
    css = css & ".content-wrapper { flex: 1 1 auto; padding: 10mm 15mm; overflow: hidden; position: relative; }" & vbCrLf
    
    ' Estilos de Tablas y Textos
    css = css & "h1, h2, h3 { color: #003366; margin-top: 0; }" & vbCrLf
    css = css & "h2 { font-size: 16pt; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-bottom: 15px; }" & vbCrLf
    css = css & "h3 { font-size: 12pt; margin-bottom: 8px; margin-top: 15px; }" & vbCrLf
    css = css & ".card { background: #fff; border: 1px solid #ddd; border-radius: 4px; padding: 10px; margin-bottom: 10px; }" & vbCrLf
    
    ' Tablas Robustas
    css = css & ".report-table { width: 100%; border-collapse: collapse; font-size: 11px; table-layout: fixed; }" & vbCrLf
    css = css & ".report-table th, .report-table td { border: 1px solid #ddd; padding: 6px; text-align: left; vertical-align: top; word-wrap: break-word; overflow-wrap: break-word; }" & vbCrLf
    css = css & ".report-table th { background-color: #f5f5f5; font-weight: bold; color: #333; }" & vbCrLf
    
    ' Estilos Especificos
    css = css & ".kv { display: flex; margin-bottom: 5px; border-bottom: 1px solid #f0f0f0; padding-bottom: 2px; }" & vbCrLf
    css = css & ".kv-k { width: 150px; font-weight: bold; color: #555; }" & vbCrLf
    css = css & ".kv-v { flex: 1; color: #000; }" & vbCrLf
    
    ' Media Print (Critico para Edge Headless)
    css = css & "@media print {" & vbCrLf
    css = css & "  body { background: white; }" & vbCrLf
    css = css & "  .pdf-page { margin: 0; box-shadow: none; page-break-after: always; }" & vbCrLf
    ' Forzar tamano de papel
    css = css & "  @page { size: " & sPageSize & " A4; margin: 0; }" & vbCrLf
    css = css & "}" & vbCrLf
    
    css = css & "</style>"
    GetEstilosCSS_PDF = css
End Function

Private Function GetHeaderHTML() As String
    Dim s As String
    s = "<div style='display:flex; width:100%; align-items:center; justify-content:space-between;'>"
    s = s & "  <div style='width:180px;'>" & GetLogoSVG("100%", "auto") & "</div>"
    s = s & "  <div style='text-align:right;'>"
    s = s & "     <div style='font-size: 14px; font-weight: bold; color:#003366;'>Defensa y Seguridad</div>"
    s = s & "     <div style='font-size: 12px; color:#666;'>Informe de Gestion de Riesgos</div>"
    s = s & "  </div>"
    s = s & "</div>"
    GetHeaderHTML = s
End Function

Private Function GetFooterHTML(ByVal p_NumPagina As Long) As String
    GetFooterHTML = "Pagina " & p_NumPagina
End Function

' =============================================================================================
' CONVERSION PDF (EDGE HEADLESS)
' =============================================================================================

Private Sub ConvertirHTMLaPDF(ByVal pHtmlContent As String, ByVal pRutaPDF As String, ByVal bEsLandscape As Boolean, ByRef pError As String)
    Dim sRutaHTML As String
    Dim sEdgePath As String
    Dim sCmd As String
    Dim sArgs As String
    Dim fso As Object
    
    On Error GoTo errores
    
    ' 1. Guardar HTML temporal
    sRutaHTML = Replace(pRutaPDF, ".pdf", ".html")
    GuardarArchivoTexto sRutaHTML, pHtmlContent
    
    ' 2. Localizar Edge
    sEdgePath = EncontrarEdge()
    If sEdgePath = "" Then
        pError = "No se encontro Microsoft Edge (msedge.exe) en las rutas estandar."
        Exit Sub
    End If
    
    ' 3. Construir Comando
    ' PowerShell Start-Process es mas robusto para argumentos con comillas y espacios
    ' Parametros clave: --headless, --print-to-pdf, --disable-gpu, --no-pdf-header-footer (ya tenemos los nuestros)
    ' Nota: --run-all-compositor-stages-before-draw asegura que se renderice todo
    
    sArgs = "-ArgumentList '--headless', '--disable-gpu', '--run-all-compositor-stages-before-draw', '--print-to-pdf=""" & pRutaPDF & """', '""" & sRutaHTML & """'"
    
    sCmd = "powershell -Command ""Start-Process -FilePath '" & sEdgePath & "' " & sArgs & " -Wait"""
    
    ' 4. Ejecutar
    Shell sCmd, vbHide
    
    ' 5. Verificar
    ' Esperar un momento y comprobar existencia (Start-Process -Wait deberia bastar, pero a veces el sistema de archivos tarda)
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To 20 ' Esperar max 2 segundos extra
        If fso.FileExists(pRutaPDF) Then Exit For
        DoEvents
        ' Sleep simulado
        Dim t As Single: t = Timer: While Timer < t + 0.1: DoEvents: Wend
    Next i
    
    If Not fso.FileExists(pRutaPDF) Then
        pError = "El archivo PDF no se genero correctamente. Verifique permisos o instalacion de Edge."
    Else
        ' Limpiar HTML temporal
        On Error Resume Next
        fso.DeleteFile sRutaHTML
        On Error GoTo 0
    End If
    
    Exit Sub
    
errores:
    pError = "Error ConvertirHTMLaPDF: " & Err.Description
End Sub

Private Function EncontrarEdge() As String
    Dim rutas(1) As String
    Dim fso As Object
    Dim i As Integer
    
    rutas(0) = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    rutas(1) = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For i = 0 To 1
        If fso.FileExists(rutas(i)) Then
            EncontrarEdge = rutas(i)
            Exit Function
        End If
    Next i
    
    EncontrarEdge = ""
End Function

Private Sub GuardarArchivoTexto(ByVal p_Ruta As String, ByVal p_Contenido As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText p_Contenido
    stream.SaveToFile p_Ruta, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

' =============================================================================================
' GENERADORES DE SECCIONES (Refactorizados para Paginacion)
' =============================================================================================

Private Sub AgregarSeccionControlCambiosPaginada(ByVal p_Edicion As Edicion, ByVal p_Proyecto As Proyecto, ByRef p_Error As String)
    Dim filas As Collection
    Dim fila As Object
    Dim sRows As String
    Dim rowsInPage As Long
    On Error GoTo errores
    p_Error = ""
    Set filas = ControlCambios.ControlCambios_GetFilasHastaEdicion(p_Edicion, p_Proyecto, ControlCambios.ControlCambios_UsarCacheProyecto(CLng(p_Proyecto.IDProyecto), p_Error), p_Error)
    If filas Is Nothing Then
        AgregarPagina "<h2>Control de Cambios</h2><div class='card'><p>No hay cambios registrados.</p></div>", False
        Exit Sub
    End If
    If filas.Count = 0 Then
        AgregarPagina "<h2>Control de Cambios</h2><div class='card'><p>No hay cambios registrados.</p></div>", False
        Exit Sub
    End If
    
    
    Dim sHeader As String, sFooter As String
    sHeader = "<h2>Control de Cambios</h2><div class='card'><table class='report-table'><thead><tr><th style='width:15%'>Codigo</th><th style='width:10%'>Ed.</th><th style='width:15%'>Estado</th><th style='width:30%'>Mitigacion</th><th style='width:30%'>Contingencia</th></tr></thead><tbody>"
    sFooter = "</tbody></table></div>"
    
    sRows = ""
    rowsInPage = 0
    
    For Each fila In filas
        sRows = sRows & "<tr>" & _
                "<td>" & HTMLSafe(Nz(fila("CodigoRiesgo"), "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(fila("Edicion"), "")) & "</td>" & _
                "<td>" & Nz(fila("EstadoHtml"), "") & "</td>" & _
                "<td>" & Nz(fila("MitigacionHtml"), "") & "</td>" & _
                "<td>" & Nz(fila("ContingenciaHtml"), "") & "</td>" & _
                "</tr>"
        
        rowsInPage = rowsInPage + 1
        If rowsInPage >= MAX_ROWS_PER_PAGE_PORTRAIT Then
            AgregarPagina sHeader & sRows & sFooter, False
            sRows = ""
            rowsInPage = 0
        End If
    Next fila
    
    If rowsInPage > 0 Then AgregarPagina sHeader & sRows & sFooter, False
    Exit Sub
errores:
    p_Error = "Error en AgregarSeccionControlCambiosPaginada " & vbNewLine & Err.Description
End Sub

Private Sub AgregarSeccionInventarioPaginada(ByVal p_Edicion As Edicion, ByVal p_FechaCierre As String, ByVal p_FechaPub As String, ByRef p_Error As String)
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As riesgo
    Dim sRows As String
    Dim rowsInPage As Long
    Dim incluirCausaRaiz As Boolean
    Dim p As Proyecto
    
    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos
    If col Is Nothing Or col.Count = 0 Then Exit Sub
    
    On Error Resume Next
    Set p = p_Edicion.Proyecto
    incluirCausaRaiz = (p.RequiereRiesgoDeBibliotecaCalculado = 1)
    On Error GoTo 0
    
    Dim sHeader As String, sFooter As String
    sHeader = "<h2>Inventario de Riesgos</h2><div class='card'><table class='report-table' style='font-size:10px;'><thead><tr style='background-color:#ECF0F1;'>" & _
              "<th style='width:8%'>Codigo</th><th style='width:25%'>Descripcion</th>" & IIf(incluirCausaRaiz, "<th style='width:15%'>Causa Raiz</th>", "") & _
              "<th style='width:8%'>Detect.</th><th style='width:6%'>Plazo</th><th style='width:6%'>Coste</th><th style='width:6%'>Calidad</th>" & _
              "<th style='width:6%'>Global</th><th style='width:6%'>Vuln.</th><th style='width:6%'>Val.</th><th style='width:6%'>Prio.</th><th style='width:8%'>Estado</th></tr></thead><tbody>"
    sFooter = "</tbody></table></div>"
    
    sRows = ""
    rowsInPage = 0
    Dim i As Long: i = 0
    
    For Each k In col
        Set r = col(k)
        i = i + 1
        Dim sBg As String: sBg = IIf(i Mod 2 = 0, "#F4F6F7", "#FFFFFF")
        
        sRows = sRows & "<tr style='background-color:" & sBg & ";'>" & _
                "<td>" & HTMLSafe(Nz(r.CodigoRiesgo, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Descripcion, "")) & "</td>"
        If incluirCausaRaiz Then sRows = sRows & "<td>" & HTMLSafe(Nz(r.CausaRaiz, "")) & "</td>"
        sRows = sRows & "<td>" & HTMLSafe(Nz(r.DetectadoPor, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Plazo, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Coste, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Calidad, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.ImpactoGlobal, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Vulnerabilidad, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Valoracion, "")) & "</td>" & _
                "<td>" & HTMLSafe(Nz(r.Priorizacion, "")) & "</td>" & _
                "<td>" & HTMLSafe(r.ESTADOCalculadoTexto) & "</td>" & _
                "</tr>"
        
        rowsInPage = rowsInPage + 1
        If rowsInPage >= MAX_ROWS_PER_PAGE_LANDSCAPE Then
            AgregarPagina sHeader & sRows & sFooter, True
            sRows = ""
            rowsInPage = 0
        End If
    Next k
    
    If rowsInPage > 0 Then AgregarPagina sHeader & sRows & sFooter, True
End Sub

Private Sub AgregarSeccionFichasRiesgoPaginada(ByVal p_Edicion As Edicion, ByVal p_Proyecto As Proyecto, ByVal p_FechaCierre As String, ByVal p_FechaPub As String, ByRef p_Error As String)
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As riesgo
    Dim sFicha As String
    On Error GoTo errores
    p_Error = ""
    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos
    If col Is Nothing Then Exit Sub
    
    For Each k In col
        Set r = col(k)
        sFicha = ConstruirFichaRiesgoHTML_Local(r, p_Proyecto, p_FechaCierre, p_FechaPub, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        AgregarPagina sFicha, False
    Next k
    Exit Sub
    
errores:
    p_Error = "Error AgregarSeccionFichasRiesgoPaginada: " & Err.Description
End Sub

Private Function TieneRiesgos(ByVal p_Edicion As Edicion) As Boolean
    Dim col As Scripting.Dictionary
    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos
    If col Is Nothing Then TieneRiesgos = False Else TieneRiesgos = (col.Count > 0)
End Function

' =============================================================================================
' HELPERS HTML Y DATOS (Incluye Logica de Negocio Local)
' =============================================================================================

Private Function ConstruirSeccionPortadaHTML(ByVal p_Titulo As String, ByVal p_Expediente As String) As String
    Dim s As String
    s = "<div style='display:flex; flex-direction:column; justify-content:center; height:100%; text-align:center;'>" & vbCrLf
    s = s & "  <h1 style='font-size: 36pt; margin-bottom: 40px; color: #003366;'>INFORME DE GESTION DE RIESGOS</h1>" & vbCrLf
    s = s & "  <div style='font-size: 18pt; color: #666;'>" & HTMLSafe(p_Expediente) & "</div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    ConstruirSeccionPortadaHTML = s
End Function

Private Function ConstruirSeccionDatosGeneralesHTML( _
                                                    p_Nombre As String, p_Expediente As String, p_Jefe As String, _
                                                    p_Cliente As String, p_Codigo As String, p_Edicion As String, _
                                                    p_Fecha As String, p_EsPublicacion As Boolean) As String
    Dim s As String
    s = "<h2>Datos Generales</h2>" & vbCrLf
    s = s & "<div class='card'>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Proyecto</div><div class='kv-v'>" & HTMLSafe(UCase$(p_Expediente)) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Nombre</div><div class='kv-v'>" & HTMLSafe(UCase$(p_Nombre)) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Jefe Proyecto</div><div class='kv-v'>" & HTMLSafe(p_Jefe) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Cliente</div><div class='kv-v'>" & HTMLSafe(p_Cliente) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Codigo Doc.</div><div class='kv-v'>" & HTMLSafe(p_Codigo) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Edicion</div><div class='kv-v'>" & HTMLSafe(p_Edicion) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Fecha</div><div class='kv-v'>" & HTMLSafe(p_Fecha) & "</div></div>" & vbCrLf
    s = s & "  <div class='kv'><div class='kv-k'>Estado</div><div class='kv-v'>" & IIf(p_EsPublicacion, "PUBLICADO", "BORRADOR") & "</div></div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    ConstruirSeccionDatosGeneralesHTML = s
End Function

Private Function ConstruirSeccionCuadroControlHTML_Local( _
                                                        ByVal p_Edicion As Edicion, _
                                                        ByVal m_Proyecto As Proyecto, _
                                                        Optional ByRef p_Error As String) As String
    Dim s As String
    Dim ediciones As Scripting.Dictionary
    Dim ed As Edicion
    Dim e As Variant
    On Error GoTo errores
    p_Error = ""
    ' Obtener historial de ediciones
    Set ediciones = Constructor.getEdiciones(m_Proyecto.IDProyecto, , p_Error)
    If p_Error <> "" Then Err.Raise 1000
    s = "<h2>Cuadro de Control</h2>" & vbCrLf
    s = s & "<div class='card'><table class='report-table'><thead><tr><th style='width:10%'>Edicion</th><th style='width:15%'>Fecha</th><th style='width:25%'>Elaborado</th><th style='width:25%'>Revisado</th><th style='width:25%'>Aprobado</th></tr></thead><tbody>"
    
    If Not ediciones Is Nothing Then
        For Each e In ediciones
            Set ed = ediciones(e)
            If IsDate(ed.FechaPublicacion) Then
                s = s & "<tr><td>" & HTMLSafe(Nz(ed.Edicion, "")) & "</td><td>" & HTMLSafe(Format(ed.FechaPublicacion, "dd/mm/yyyy")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(ed.Elaborado, "")) & "</td><td>" & HTMLSafe(Nz(ed.Revisado, "")) & "</td><td>" & HTMLSafe(Nz(ed.Aprobado, "")) & "</td></tr>"
            End If
        Next
    End If
    s = s & "</tbody></table></div>"
    ConstruirSeccionCuadroControlHTML_Local = s
    Exit Function
    
errores:
    p_Error = "Error ConstruirSeccionCuadroControlHTML_Local: " & Err.Description
End Function

Private Function ConstruirFichaRiesgoHTML_Local( _
                                                p_Riesgo As riesgo, p_Proyecto As Proyecto, _
                                                p_FechaCierre As String, _
                                                p_FechaPublicacion As String, _
                                                Optional ByRef p_Error As String) As String
    Dim s As String, sError As String
    On Error GoTo errores
    p_Error = ""
    s = "<h2>Ficha de Riesgo: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</h2>"
    
    ' Seccion General (Grid)
    s = s & "<div class='card' style='display:grid; grid-template-columns: repeat(3, 1fr); gap:10px;'>"
    s = s & "  <div><strong>Estado:</strong> " & HTMLSafe(p_Riesgo.ESTADOCalculadoTexto) & "</div>"
    s = s & "  <div><strong>Prioridad:</strong> " & HTMLSafe(p_Riesgo.Priorizacion) & "</div>"
    s = s & "  <div><strong>Impacto Global:</strong> " & HTMLSafe(p_Riesgo.ImpactoGlobalCalculado) & "</div>"
    s = s & "</div>"
    
    ' Descripcion
    s = s & "<div class='card'><h3>Descripcion</h3><div>" & HTMLSafe(Nz(p_Riesgo.Descripcion, "")) & "</div>"
    If Nz(p_Riesgo.CausaRaiz, "") <> "" Then
        s = s & "<h3 style='margin-top:10px;'>Causa Raiz</h3><div>" & HTMLSafe(Nz(p_Riesgo.CausaRaiz, "")) & "</div>"
    End If
    s = s & "</div>"
    
    ' Historico Estados (Funcion Externa simulada o llamada)
    ' s = s & ConstruirSeccionHistorico(p_Riesgo) ' Si existiera
    
    ' Planes
    s = s & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Mitigacion, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    s = s & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Contingencia, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    ' Materializaciones
    s = s & ConstruirSeccionMaterializaciones(p_Riesgo, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ConstruirFichaRiesgoHTML_Local = s
    Exit Function
    
errores:
    p_Error = "Error ConstruirFichaRiesgoHTML_Local: " & Err.Description
End Function

' --- COPIA DE FUNCIONES DE APOYO ---

Public Function GetLogoSVG(Optional ByVal pWidth As String = "100%", Optional ByVal pHeight As String = "auto") As String
    ' SVG Telefnica
    GetLogoSVG = "<svg viewBox='0 0 1920 802' style='width:" & pWidth & "; height:" & pHeight & ";'>" & _
                 "<path fill='#0066FF' d='m685.79,360.16h-48.84v-29.07h127.9v29.07h-48.84v133.72h-30.23v-133.72Z'/>" & _
                 "<path fill='#0066FF' d='m866.94,458.99c-2.56,9.3-17.21,37.21-54.65,37.21-34.88,0-60.47-25.58-60.47-61.63s25.58-61.63,60.47-61.63c32.56,0,58.14,25.58,58.14,59.3,0,3.49-.47,6.28-.7,8.37l-.47,3.26h-88.37c2.56,16.51,14.88,27.91,31.4,27.91,13.72,0,22.09-7.9,24.42-12.79h30.23Zm-25.58-34.88c-2.56-15.12-12.56-26.75-29.07-26.75-17.67,0-27.91,11.63-31.4,26.75h60.47Z'/>" & _
                 "<circle fill='#0066FF' cx='275.37' cy='277.85' r='52.41'/><circle fill='#0066FF' cx='398.52' cy='277.85' r='52.41'/><circle fill='#0066FF' cx='521.67' cy='277.85' r='52.41'/>" & _
                 "</svg>"
End Function

Private Function ConstruirSeccionPlanes(p_Riesgo As riesgo, ByVal Tipo As EnumTipoPlan, Optional ByRef p_Error As String) As String
    Dim col As Scripting.Dictionary
    Dim p As Object
    Dim s As String, sTitulo As String
    Dim m_Codigo As String
    On Error GoTo errores
    p_Error = ""

    If Tipo = Mitigacion Then
        Set col = p_Riesgo.ColPMs
        sTitulo = "Planes de Mitigacion"
    Else
        Set col = p_Riesgo.ColPCs
        sTitulo = "Planes de Contingencia"
    End If
    If col Is Nothing Then Exit Function
    If col.Count = 0 Then Exit Function
    
    s = "<h3>" & sTitulo & "</h3>"
    Dim v As Variant
    For Each v In col
        Set p = col(v)
        If Tipo = Mitigacion Then
            m_Codigo = p.CodMitigacion
        Else
            m_Codigo = p.CodContingencia
        End If
        s = s & "<div style='margin-bottom:8px; padding-left:10px; border-left:3px solid #eee;'>"
        s = s & "<div><strong>" & HTMLSafe(m_Codigo) & ":</strong> " & HTMLSafe(Nz(p.DisparadorDelPlan, "")) & "</div>"
        's = s & "<div style='font-size:10px; color:#666;'>Responsable: " & HTMLSafe(Nz(p.Responsable, "")) & " | Fecha: " & HTMLSafe(Nz(p.FechaImplantacion, "")) & "</div>"
        s = s & "</div>"
    Next v
    ConstruirSeccionPlanes = s
    Exit Function
    
errores:
    p_Error = "Error ConstruirSeccionPlanes: " & Err.Description
End Function

Private Function ConstruirSeccionMaterializaciones(p_Riesgo As riesgo, Optional ByRef p_Error As String) As String
    Dim col As Scripting.Dictionary
    Dim mat As RiesgoMaterializacion
    Dim s As String
    Dim k As Variant
    On Error GoTo errores
    p_Error = ""
    Set col = p_Riesgo.ColMaterializaciones
    If col Is Nothing Then Exit Function
    If col.Count = 0 Then Exit Function

    
    s = "<h3>Materializaciones</h3>"
    s = s & "<div class='card'><table class='report-table'><thead><tr><th>Fecha</th></tr></thead><tbody>"
    For Each k In col
        Set mat = col(k)
        s = s & "<tr><td>" & HTMLSafe(Nz(mat.Fecha, "")) & "</td></tr>"
    Next k
    s = s & "</tbody></table></div>"
    ConstruirSeccionMaterializaciones = s
    Exit Function
    
errores:
    p_Error = "Error ConstruirSeccionMaterializaciones: " & Err.Description
End Function


Private Function HTMLSafe(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HTMLSafe = s
End Function



