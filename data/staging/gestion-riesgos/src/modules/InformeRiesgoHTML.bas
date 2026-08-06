Attribute VB_Name = "InformeRiesgoHTML"
Option Compare Database
Option Explicit

' =========================================================================
' Módulo: InformeRiesgoHTML
' Proyecto: GESTIÓN DE RIESGOS
' Descripción: Generador de informes 100% autosuficiente con LOGO SVG PROPIO.
'              Diseño corporativo premium basado en Indicadores.html.
' =========================================================================

Public Function GenerarInformeRiesgoHTML( _
                                        Optional ByVal p_IDRiesgo As String, _
                                        Optional ByVal p_hWnd As Long, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_Riesgo As Riesgo
    Dim m_HTML As String
    Dim m_URL As String

    On Error GoTo errores
    p_Error = ""

    ' 1. Obtención del objeto Riesgo
    If p_IDRiesgo <> "" Then
        Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo, p_Error)
        If p_Error <> "" Then err.Raise 1000
    ElseIf Not m_ObjRiesgoActivo Is Nothing Then
        Set m_Riesgo = m_ObjRiesgoActivo
    End If

    If m_Riesgo Is Nothing Then
        p_Error = "No se ha indicado un riesgo para generar el informe"
        err.Raise 1000
    End If

    ' 2. Construcción del HTML
    m_HTML = ConstruirInformeRiesgoHTML(m_Riesgo, p_Error)
    If p_Error <> "" Then err.Raise 1000

    ' 3. Guardar en formato UTF-8 Real (Garantiza tildes y eñes)
    m_URL = GuardarInformeHTML_UTF8(m_Riesgo, m_HTML, p_Error)
    If p_Error <> "" Then err.Raise 1000

    ' 4. Abrir en el navegador predeterminado
    If p_hWnd = 0 Then
        On Error Resume Next
        p_hWnd = Application.hWndAccessApp
        On Error GoTo errores
    End If
    
    Ejecutar p_hWnd, "open", m_URL, "", "", 1

    GenerarInformeRiesgoHTML = m_URL
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "Error en GenerarInformeRiesgoHTML: " & err.description
    End If
End Function

Public Function GenerarInformeEdicionHTML( _
                                        ByVal p_Edicion As Edicion, _
                                        Optional ByVal p_hWnd As Long, _
                                        Optional ByVal p_FechaCierre As String, _
                                        Optional ByVal p_FechaPublicacion As String, _
                                        Optional ByRef p_Error As String, _
                                        Optional ByVal p_GenerarPDF As Boolean = False, _
                                        Optional ByVal p_ControlCambiosAlcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceResumen3 _
                                        ) As String

    Dim m_HTML As String
    Dim m_URL As String
    Dim m_URL_Apertura As String

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        err.Raise 1000
    End If

    m_HTML = ConstruirInformeEdicionHTML(p_Edicion, p_FechaCierre, p_FechaPublicacion, p_Error, p_GenerarPDF, p_ControlCambiosAlcance)
    If p_Error <> "" Then err.Raise 1000

    m_URL = GuardarInformeEdicionHTML_UTF8(p_Edicion, m_HTML, p_Error)
    If p_Error <> "" Then err.Raise 1000
    
    m_URL_Apertura = m_URL ' Por defecto, abriremos el HTML

    ' --- NUEVA INTEGRACIÓN PDF OPCIONAL ---
    If p_GenerarPDF Then
        Dim m_URL_PDF As String
        m_URL_PDF = Replace(m_URL, ".html", ".pdf", 1, -1, vbTextCompare)
        If ConvertirHTMLaPDFConEdge(m_URL, m_URL_PDF, p_Error) Then
            m_URL_Apertura = m_URL_PDF ' Si todo va bien, cambiamos a PDF
        Else
            err.Raise 1000
        End If
    End If
    ' --------------------------------------

    If p_hWnd = 0 Then
        On Error Resume Next
        p_hWnd = Application.hWndAccessApp
        On Error GoTo errores
    End If
    
    Ejecutar p_hWnd, "open", m_URL_Apertura, "", "", 1

    GenerarInformeEdicionHTML = m_URL_Apertura
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "Error en GenerarInformeEdicionHTML: " & err.description
    End If
End Function

Public Function GenerarInformeControlCambiosHistoricoHTML( _
                                        ByVal p_Edicion As Edicion, _
                                        Optional ByVal p_hWnd As Long, _
                                        Optional ByVal p_FechaCierre As String, _
                                        Optional ByVal p_FechaPublicacion As String, _
                                        Optional ByRef p_Error As String, _
                                        Optional ByVal p_GenerarPDF As Boolean = False _
                                        ) As String
    GenerarInformeControlCambiosHistoricoHTML = GenerarInformeEdicionHTML( _
                                                p_Edicion:=p_Edicion, _
                                                p_hWnd:=p_hWnd, _
                                                p_FechaCierre:=p_FechaCierre, _
                                                p_FechaPublicacion:=p_FechaPublicacion, _
                                                p_Error:=p_Error, _
                                                p_GenerarPDF:=p_GenerarPDF, _
                                                p_ControlCambiosAlcance:=EnumControlCambiosAlcanceCompleto)
End Function

Private Function ConstruirInformeRiesgoHTML(ByVal p_Riesgo As Riesgo, ByRef p_Error As String) As String
    Dim html As String
    Dim sNemotecnico As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim badgeText As String
    Dim badgeClass As String
    On Error GoTo errores

    On Error Resume Next
    sNemotecnico = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Nemotecnico, "")
    On Error GoTo errores
    
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html lang='es'>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "    <meta charset='UTF-8'>" & vbCrLf
    html = html & "    <title>Gestión de Riesgos - Ficha " & p_Riesgo.CodigoRiesgo & "</title>" & vbCrLf
    html = html & GetEstilosCSS_Corporativos()
    html = html & "</head>" & vbCrLf
    
    
    html = html & "<body>" & vbCrLf
    
    ' HEADER CON LOGO SVG
    html = html & "    <header>" & vbCrLf
    html = html & "        <div class='header-container container'>" & vbCrLf
    html = html & "            <div class='logo-container'>" & vbCrLf
    html = html & "                <div class='logo-wrapper'>" & GetLogoSVG() & "</div>" & vbCrLf
    html = html & "                <div class='header-text'>" & vbCrLf
'    If Trim$(sNemotecnico) <> "" Then
'        html = html & "                    <h1>GESTIÓN DE RIESGOS - " & HTMLSafe(sNemotecnico) & "</h1>" & vbCrLf
'    Else
'        html = html & "                    <h1>GESTIÓN DE RIESGOS</h1>" & vbCrLf
'    End If
    html = html & "                    <h1>GESTIÓN DE RIESGOS-Publicabilidad Riesgo</h1>" & vbCrLf
    html = html & "                    <p>Detalle Técnico del Riesgo</p>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "            <div class='header-info'>" & vbCrLf
    html = html & "                <p><strong>CÓDIGO: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</strong></p>" & vbCrLf
    html = html & "                <p>Fecha: " & Format(Date, "dd/mm/yyyy") & "</p>" & vbCrLf
    On Error Resume Next
    If ConstruirDatosPublicabilidadRiesgo(p_Riesgo, m_Datos, , p_Error) = EnumSiNo.No Then
        badgeText = "RIESGO: N/A"
        badgeClass = "hdr-badge"
    Else
        Call EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, p_Error)
        Select Case m_Veredicto
            Case EnumPublicabilidadVeredicto.Publicable
                badgeText = "RIESGO: PUBLICABLE"
                badgeClass = "hdr-badge publicable"
            Case EnumPublicabilidadVeredicto.NoPublicable
                badgeText = "RIESGO: NO PUBLICABLE"
                badgeClass = "hdr-badge no-publicable"
            Case Else
                badgeText = "RIESGO: NO APLICA"
                badgeClass = "hdr-badge no-aplica"
        End Select
    End If
    On Error GoTo errores
    html = html & "                <div class='" & badgeClass & "'>" & badgeText & "</div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </header>" & vbCrLf

    ' Contenedor Principal
    html = html & "    <div class='container main-container'>" & vbCrLf
    
    ' Buscador de secciones
    html = html & "        <div class='section-search-wrapper'>" & vbCrLf
    html = html & "            <div class='section-search-box'>" & vbCrLf
    html = html & "                <span class='search-icon'>&#128269;</span>" & vbCrLf
    html = html & "                <input id='sectionSearchInput' type='text' placeholder='Buscar apartado...' autocomplete='off'>" & vbCrLf
    html = html & "                <div id='sectionSearchResults' class='section-search-results'></div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    
    ' Sistema de Pestañas
    html = html & "        <div class='tab-container'>" & vbCrLf
    html = html & "            <button class='tab-btn active' onclick=""openTab(event, 'general')"">Datos Generales</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'datosProyecto')"">Datos de Proyecto</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'mitigacion')"">Mitigación</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'contingencia')"">Contingencia</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'materializaciones')"">Materializaciones</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'publicabilidad')"">Publicabilidad</button>" & vbCrLf
    html = html & "        </div>" & vbCrLf

    ' Sección General
    html = html & "        <div id='general' class='tab-content active'>" & vbCrLf & ConstruirSeccionGeneral(p_Riesgo) & "</div>" & vbCrLf

    html = html & "        <div id='datosProyecto' class='tab-content'>" & vbCrLf & ConstruirSeccionDatosProyecto(p_Riesgo) & "</div>" & vbCrLf
    
    ' Sección Mitigación
    html = html & "        <div id='mitigacion' class='tab-content'>" & vbCrLf & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Mitigacion) & "</div>" & vbCrLf
    
    ' Sección Contingencia
    html = html & "        <div id='contingencia' class='tab-content'>" & vbCrLf & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Contingencia) & "</div>" & vbCrLf

    html = html & "        <div id='materializaciones' class='tab-content'>" & vbCrLf & ConstruirSeccionMaterializaciones(p_Riesgo) & "</div>" & vbCrLf
    html = html & "        <div id='publicabilidad' class='tab-content'>" & vbCrLf & ConstruirSeccionPublicabilidad(p_Riesgo) & "</div>" & vbCrLf

    html = html & "    </div>" & vbCrLf
    
    html = html & "    <footer><p>© " & Year(Date) & " Telefónica - Aplicación GESTIÓN DE RIESGOS</p></footer>" & vbCrLf
    html = html & GetScriptsJS()
    html = html & "</body>" & vbCrLf
    html = html & "</html>"
    
    ConstruirInformeRiesgoHTML = html
    Exit Function
errores:
    p_Error = "Error en ConstruirInformeRiesgoHTML: " & err.description
End Function

Public Function ConstruirInformeEdicionHTML( _
                                            ByVal p_Edicion As Edicion, _
                                            ByVal p_FechaCierre As String, _
                                            ByVal p_FechaPublicacion As String, _
                                            ByRef p_Error As String, _
                                            Optional ByVal p_EsPDF As Boolean = False, _
                                            Optional ByVal p_ControlCambiosAlcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceResumen3 _
                                            ) As String

    Dim html As String
    Dim m_Proyecto As Proyecto
    Dim sNombreProyecto As String, sExpediente As String, sCliente As String, sCodigoDocumento As String
    Dim sJefeProyecto As String, sEdicion As String, sFechaPub As String, sFechaCierre As String
    Dim sNemotecnico As String
    Dim sTituloExpediente As String
    Dim m_EsPublicacion As Boolean

    On Error GoTo errores
    p_Error = ""

    Set m_Proyecto = p_Edicion.Proyecto
    If m_Proyecto Is Nothing Then
        Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
        If p_Error <> "" Then err.Raise 1000
    End If

    sNombreProyecto = Nz(m_Proyecto.NombreProyecto, Nz(m_Proyecto.Proyecto, ""))
    sExpediente = Nz(m_Proyecto.Proyecto, "")
    sCliente = Nz(m_Proyecto.Cliente, "")
    sCodigoDocumento = Nz(m_Proyecto.CodigoDocumento, "")
    sJefeProyecto = Nz(p_Edicion.Elaborado, "")
    sEdicion = Nz(p_Edicion.Edicion, "")
    sNemotecnico = ""
    sTituloExpediente = m_Proyecto.TituloExpediente
    On Error Resume Next
    sNemotecnico = Nz(m_Proyecto.Expediente.Nemotecnico, "")
    On Error GoTo errores

    m_EsPublicacion = IsDate(p_FechaPublicacion)
    If IsDate(p_FechaPublicacion) Then
        sFechaPub = Format$(CDate(p_FechaPublicacion), "dd/mm/yyyy")
    Else
        sFechaPub = "-"
    End If
    
    If IsDate(sFechaCierre) Then
        sFechaCierre = Format$(CDate(p_FechaCierre), "dd/mm/yyyy")
    Else
        sFechaCierre = "-"
    End If

    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'>" & vbCrLf
    html = html & "<title>Gestión de Riesgos - Edición " & sEdicion & "</title>" & vbCrLf
    html = html & GetEstilosCSS_Corporativos()
    If Not p_EsPDF Then html = html & GetEstilosCSS_WebOnly()
    html = html & GetEstilosCSS_InformeEdicion_Print(p_EsPDF)
    html = html & "</head><body>" & vbCrLf

    html = html & ConstruirPiePaginaImpresionHTML(sCodigoDocumento, sFechaPub, sEdicion)
    
    ' Portada: se abre la sección, luego el header integrado, luego los títulos
    ' ConstruirSeccionPortadaHTML genera la apertura <section class='print-only-cover'>
    ' pero necesitamos insertar el header ANTES de los títulos
    Dim sPortada As String
    sPortada = "<section class='print-only-cover'>" & vbCrLf
    
    ' Header integrado en la portada (visible solo en print)
    sPortada = sPortada & "    <header>" & vbCrLf
    sPortada = sPortada & "        <div class='header-container container'>" & vbCrLf
    sPortada = sPortada & "            <div class='logo-container'>" & vbCrLf
    sPortada = sPortada & "                <div class='logo-wrapper'>" & GetLogoSVG("200px") & "</div>" & vbCrLf
    sPortada = sPortada & "            </div>" & vbCrLf
    sPortada = sPortada & "            <div class='header-info'>" & vbCrLf
    sPortada = sPortada & "                <p><strong>Defensa y Seguridad</strong></p>" & vbCrLf
    sPortada = sPortada & "                <p><strong>Informe:</strong> " & IIf(m_EsPublicacion, "Publicación", "Borrador") & "</p>" & vbCrLf
    sPortada = sPortada & "            </div>" & vbCrLf
    sPortada = sPortada & "        </div>" & vbCrLf
    sPortada = sPortada & "    </header>" & vbCrLf
    
    ' Títulos centrados
    sPortada = sPortada & "  <div class='cover-titles'>" & vbCrLf
    sPortada = sPortada & "    <h1 style='font-size: 22pt; font-weight: 700; margin-bottom: 30px;'>INFORME DE GESTIÓN DE RIESGOS</h1>" & vbCrLf
    sPortada = sPortada & "    <div style='font-size: 16pt; font-weight: 600; color: #333;'>" & HTMLSafe(UCase$(sTituloExpediente)) & "</div>" & vbCrLf
    sPortada = sPortada & "    <div style='font-size: 14pt; opacity: 0.6; margin-top: 15px;'>" & HTMLSafe(sExpediente) & "</div>" & vbCrLf
    sPortada = sPortada & "  </div>" & vbCrLf
    sPortada = sPortada & "</section>" & vbCrLf
    
    html = html & sPortada

    ' Header para pantalla (se oculta en print via CSS .screen-only-header)
    html = html & "    <header class='screen-only-header'>" & vbCrLf
    html = html & "        <div class='header-container container'>" & vbCrLf
    html = html & "            <div class='logo-container'>" & vbCrLf
    html = html & "                <div class='logo-wrapper'>" & GetLogoSVG("200px") & "</div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "            <div class='header-info'>" & vbCrLf
    html = html & "                <p><strong>Defensa y Seguridad</strong></p>" & vbCrLf
    html = html & "                <p><strong>Informe:</strong> " & IIf(m_EsPublicacion, "Publicación", "Borrador") & "</p>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </header>" & vbCrLf

    html = html & "<div class='container main-container'>" & vbCrLf
    If Not p_EsPDF Then
        html = html & "  <div class='section-search-wrapper'>" & vbCrLf
        html = html & "    <div class='section-search-box'>" & vbCrLf
        html = html & "      <span class='search-icon'>&#128269;</span>" & vbCrLf
        html = html & "      <input id='sectionSearchInput' type='text' placeholder='Buscar secciones' />" & vbCrLf
        html = html & "      <div id='sectionSearchResults' class='section-search-results'></div>" & vbCrLf
        html = html & "    </div>" & vbCrLf
        html = html & "  </div>" & vbCrLf
    End If
    html = html & ConstruirSeccionDatosGeneralesHTML(sNombreProyecto, sExpediente, sJefeProyecto, sCliente, sCodigoDocumento, sEdicion, sFechaPub, m_EsPublicacion, p_EsPDF)
    html = html & ConstruirSeccionCuadroControlHTML(p_Edicion, m_Proyecto, p_EsPDF)
    html = html & ConstruirSeccionControlCambiosHTML(p_Edicion, m_Proyecto, p_Error, p_EsPDF, p_ControlCambiosAlcance)
    html = html & ConstruirSeccionInventarioRiesgosHTML(p_Edicion, sFechaCierre, sFechaPub, p_Error, p_EsPDF)
    html = html & ConstruirSeccionFichasRiesgoHTML(p_Edicion, m_Proyecto, sFechaCierre, sFechaPub, p_Error, p_EsPDF)
    html = html & "</div>" & vbCrLf

    If Not p_EsPDF Then html = html & GetScriptsJS()
    html = html & "</body></html>"

    ConstruirInformeEdicionHTML = html
    Exit Function
errores:
    p_Error = "Error: " & err.description
End Function

Private Function ConstruirSeccionCuadroControlHTML( _
                                                    ByVal p_EdicionActual As Edicion, _
                                                    ByVal p_Proyecto As Proyecto, _
                                                    Optional ByVal p_EsPDF As Boolean = False _
                                                    ) As String
    Dim s As String
    Dim edByNum As Scripting.Dictionary
    Dim minEd As Long, maxEd As Long, i As Long
    Dim kEd As Variant
    Dim ed As Edicion
    Dim edNum As Long
    Dim sFecha As String
    Dim sElab As String, sRev As String, sApr As String
    
    edNum = 0
    If Not p_EdicionActual Is Nothing Then
        If IsNumeric(p_EdicionActual.Edicion) Then edNum = CLng(p_EdicionActual.Edicion)
    End If
    If edNum <= 0 Then edNum = 0
    
    Set edByNum = New Scripting.Dictionary
    edByNum.CompareMode = TextCompare
    minEd = 2147483647
    maxEd = -2147483647
    
    If Not p_Proyecto Is Nothing Then
        For Each kEd In p_Proyecto.colEdiciones
            Set ed = p_Proyecto.colEdiciones(kEd)
            If Not ed Is Nothing Then
                If IsNumeric(ed.Edicion) Then
                    If CLng(ed.Edicion) <= edNum Or edNum = 0 Then
                        If CLng(ed.Edicion) < minEd Then minEd = CLng(ed.Edicion)
                        If CLng(ed.Edicion) > maxEd Then maxEd = CLng(ed.Edicion)
                        edByNum.Add CStr(ed.Edicion), ed
                    End If
                End If
            End If
        Next kEd
    End If
    
    If p_EsPDF Then
        s = "<section class='report-section'>" & vbCrLf
        s = s & "  <h2>Cuadro de control</h2>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    Else
        s = "<details open class='report-section collapsible-section'>" & vbCrLf
        s = s & "  <summary><h2>Cuadro de control</h2></summary>" & vbCrLf
        s = s & "  <div class='collapsible-body'>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    End If

    s = s & "    <table class='report-table report-table-small'>" & vbCrLf
    s = s & "      <thead><tr><th>Edición</th><th>Fecha</th><th>Elaborado</th><th>Revisado</th><th>Aprobado</th></tr></thead>" & vbCrLf
    s = s & "      <tbody>" & vbCrLf
    
    If minEd = 2147483647 Or maxEd = -2147483647 Then
        s = s & "        <tr><td colspan='5'>No hay ediciones.</td></tr>" & vbCrLf
    Else
        For i = minEd To maxEd
            If edByNum.Exists(CStr(i)) Then
                Set ed = edByNum(CStr(i))
                
                sFecha = ""
                If IsDate(ed.FechaPublicacion) Then
                    sFecha = Format$(CDate(ed.FechaPublicacion), "dd/mm/yyyy")
                ElseIf IsDate(ed.FechaEdicion) Then
                    sFecha = Format$(CDate(ed.FechaEdicion), "dd/mm/yyyy")
                End If
                
                sElab = Nz(ed.Elaborado, "")
                sRev = Nz(ed.Revisado, "")
                sApr = Nz(ed.Aprobado, "")
                
                s = s & "        <tr>" & _
                        "<td>" & HTMLSafe(Nz(ed.Edicion, "")) & "</td>" & _
                        "<td>" & HTMLSafe(sFecha) & "</td>" & _
                        "<td>" & HTMLSafe(sElab) & "</td>" & _
                        "<td>" & HTMLSafe(sRev) & "</td>" & _
                        "<td>" & HTMLSafe(sApr) & "</td>" & _
                        "</tr>" & vbCrLf
            End If
        Next i
    End If
    
    s = s & "      </tbody>" & vbCrLf
    s = s & "      <tfoot>" & vbCrLf
    s = s & "        <tr><td colspan='5' style='height: 18mm; border: none !important; background: transparent !important;'></td></tr>" & vbCrLf
    s = s & "      </tfoot>" & vbCrLf
    s = s & "    </table>" & vbCrLf

    If p_EsPDF Then
        s = s & "  </div>" & vbCrLf
        s = s & "</section>" & vbCrLf
    Else
        s = s & "  </div>" & vbCrLf
        s = s & "  </div>" & vbCrLf
        s = s & "</details>" & vbCrLf
    End If
    
    ConstruirSeccionCuadroControlHTML = s
End Function

Private Function ConstruirPiePaginaImpresionHTML(ByVal p_CodDoc As String, ByVal p_Fecha As String, ByVal p_Edi As String) As String
    Dim s As String
    s = "<div class='print-footer'>" & vbCrLf
    s = s & "  <span>" & HTMLSafe(p_CodDoc) & " &nbsp;|&nbsp; Ed. " & HTMLSafe(p_Edi) & "</span>" & vbCrLf
    s = s & "</div>" & vbCrLf
    ConstruirPiePaginaImpresionHTML = s
End Function

Public Function InformeHTML_CompactarFechasAccionParaTest( _
                                                    ByVal p_FechaInicio As Variant, _
                                                    ByVal p_FechaFinPrevista As Variant, _
                                                    ByVal p_FechaFinReal As Variant _
                                                    ) As String
    InformeHTML_CompactarFechasAccionParaTest = InformeHTML_CompactarFechasAccion(p_FechaInicio, p_FechaFinPrevista, p_FechaFinReal)
End Function

Public Function InformeHTML_EstadoCanonicoParaTest( _
                                                ByVal p_EstadoEnum As Long, _
                                                ByVal p_EstadoCalculado As String, _
                                                ByVal p_FechaCierre As String, _
                                                ByVal p_FechaEstado As Variant _
                                                ) As String
    Dim sEstado As String
    Dim sFecha As String
    InformeHTML_ResolverEstadoCanonico p_EstadoEnum, p_EstadoCalculado, p_FechaCierre, p_FechaEstado, sEstado, sFecha
    InformeHTML_EstadoCanonicoParaTest = sEstado & "|" & sFecha
End Function

Public Function InformeHTML_ComponerSeccionControlCambiosParaTest( _
                                                            ByVal p_Inner As String, _
                                                            Optional ByVal p_EsPDF As Boolean = False _
                                                            ) As String
    InformeHTML_ComponerSeccionControlCambiosParaTest = InformeHTML_ComponerSeccionControlCambios(p_Inner, p_EsPDF)
End Function

Public Function InformeHTML_ComponerShellEdicionParaTest( _
                                                        Optional ByVal p_EsPDF As Boolean = False _
                                                        ) As String
    Dim html As String

    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'>" & vbCrLf
    html = html & GetEstilosCSS_Corporativos()
    If Not p_EsPDF Then html = html & GetEstilosCSS_WebOnly()
    html = html & GetEstilosCSS_InformeEdicion_Print(p_EsPDF)
    html = html & "</head><body>" & vbCrLf
    html = html & ConstruirPiePaginaImpresionHTML("DOC-GR", "01/02/2026", "1")
    html = html & "<section class='print-only-cover'>" & vbCrLf
    html = html & "  <div class='cover-titles'><h1>INFORME DE GESTIÓN DE RIESGOS</h1></div>" & vbCrLf
    html = html & "</section>" & vbCrLf
    html = html & "<div class='container main-container'>" & vbCrLf
    If Not p_EsPDF Then
        html = html & "  <div class='section-search-wrapper'>" & vbCrLf
        html = html & "    <div class='section-search-box'>" & vbCrLf
        html = html & "      <input id='sectionSearchInput' type='text' placeholder='Buscar secciones' />" & vbCrLf
        html = html & "      <div id='sectionSearchResults' class='section-search-results'></div>" & vbCrLf
        html = html & "    </div>" & vbCrLf
        html = html & "  </div>" & vbCrLf
    End If
    html = html & ConstruirSeccionDatosGeneralesHTML("Proyecto", "EXP-001", "JP", "Cliente", "DOC-GR", "1", "01/02/2026", True, p_EsPDF)
    html = html & InformeHTML_ComponerSeccionControlCambios("<table class='cc-table'></table>", p_EsPDF)
    html = html & "</div>" & vbCrLf
    If Not p_EsPDF Then html = html & GetScriptsJS()
    html = html & "</body></html>"

    InformeHTML_ComponerShellEdicionParaTest = html
End Function

' ----------------------------------------------------------------------------
' Punto 09 (issue #88) - Render de un check individual del informe de
' publicabilidad. Wrapper publico ParaTest que expone el helper privado
' `InformeHTML_RenderCheckPublicabilidad` para que los atomos TDD
' (Test_InformeRiesgoHTML_Publicabilidad_*) puedan ejercitar el render del
' motivo en dos partes (que se evalua + por que no se cumple) sin
' instanciar un objeto Riesgo ni un DAO.Database.
' ----------------------------------------------------------------------------
Public Function InformeHTML_RenderCheckPublicabilidadParaTest( _
                                                        ByVal p_Check As Object, _
                                                        ByVal p_Estado As EnumPublicabilidadCheckEstado _
                                                        ) As String
    InformeHTML_RenderCheckPublicabilidadParaTest = InformeHTML_RenderCheckPublicabilidad(p_Check, p_Estado)
End Function

Private Function ConstruirSeccionControlCambiosHTML( _
                                                    ByVal p_EdicionActual As Edicion, _
                                                    ByVal p_Proyecto As Proyecto, _
                                                    ByRef p_Error As String, _
                                                    Optional ByVal p_EsPDF As Boolean = False, _
                                                    Optional ByVal p_ControlCambiosAlcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceResumen3 _
                                                    ) As String
    Dim s As String
    Dim inner As String
    On Error GoTo errores
    p_Error = ""
    inner = ControlCambios_ConstruirSeccionControlCambiosHTMLConAlcance(p_EdicionActual, p_Proyecto, p_ControlCambiosAlcance, p_Error)
    inner = Replace(inner, " eliminado</span>", "</span>")
    inner = Replace(inner, "<s>", "")
    inner = Replace(inner, "</s>", "")
    inner = Replace(inner, "<section class='report-section print-page seccion-vertical'>", "")
    inner = Replace(inner, "<section class='report-section seccion-vertical'>", "")
    inner = Replace(inner, "<section class='report-section'>", "")
    inner = Replace(inner, "</section>", "")
    inner = Replace(inner, "<h2>Control de cambios</h2>", "")
    
    s = InformeHTML_ComponerSeccionControlCambios(inner, p_EsPDF)
    
    ConstruirSeccionControlCambiosHTML = s
    Exit Function
errores:
    If err.Number <> 1000 Then p_Error = "Error en ConstruirSeccionControlCambiosHTML: " & err.description
End Function

Private Function InformeHTML_ComponerSeccionControlCambios( _
                                                    ByVal p_Inner As String, _
                                                    ByVal p_EsPDF As Boolean _
                                                    ) As String
    Dim s As String
    Dim sLegend As String

    sLegend = "<div class='legend-container'>" & _
              "<span class='legend-item'><span class='dot dot-new'></span> Nuevo</span>" & _
              "<span class='legend-item'><span class='dot dot-unchanged'></span> Sin cambios</span>" & _
              "<span class='legend-item'><span class='dot dot-deleted'></span> Eliminado</span>" & _
              "</div>" & vbCrLf

    If p_EsPDF Then
        s = "<section class='report-section seccion-vertical'>" & vbCrLf
        s = s & "  <h2>Control de cambios</h2>" & vbCrLf
        s = s & p_Inner
        s = s & sLegend
        s = s & "</section>" & vbCrLf
    Else
        s = "<details open class='report-section seccion-vertical collapsible-section'>" & vbCrLf
        s = s & "  <summary><h2>Control de cambios</h2></summary>" & vbCrLf
        s = s & "  <div class='collapsible-body'>" & vbCrLf
        s = s & p_Inner
        s = s & sLegend
        s = s & "  </div>" & vbCrLf
        s = s & "</details>" & vbCrLf
    End If

    InformeHTML_ComponerSeccionControlCambios = s
End Function

Private Function SonRiesgosDiferentes_HTML(ByVal r1 As Riesgo, ByVal r2 As Riesgo) As Boolean
    On Error Resume Next
    If r2 Is Nothing Then
        SonRiesgosDiferentes_HTML = True
        Exit Function
    End If
    If r1.Estado <> r2.Estado Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.ImpactoGlobal <> r2.ImpactoGlobal Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Vulnerabilidad <> r2.Vulnerabilidad Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Valoracion <> r2.Valoracion Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Priorizacion <> r2.Priorizacion Then SonRiesgosDiferentes_HTML = True: Exit Function
    If (r1.colPMs Is Nothing) Xor (r2.colPMs Is Nothing) Then SonRiesgosDiferentes_HTML = True: Exit Function
    If Not r1.colPMs Is Nothing And Not r2.colPMs Is Nothing Then
        If r1.colPMs.Count <> r2.colPMs.Count Then SonRiesgosDiferentes_HTML = True: Exit Function
    End If
    If (r1.colPCs Is Nothing) Xor (r2.colPCs Is Nothing) Then SonRiesgosDiferentes_HTML = True: Exit Function
    If Not r1.colPCs Is Nothing And Not r2.colPCs Is Nothing Then
        If r1.colPCs.Count <> r2.colPCs.Count Then SonRiesgosDiferentes_HTML = True: Exit Function
    End If
    SonRiesgosDiferentes_HTML = False
End Function

Private Function ConstruirTextoEstadoCambiosHTML(ByVal r As Riesgo, ByVal rPrev As Riesgo, ByVal esPrimera As Boolean) As String
    Dim s As String
    If esPrimera Or rPrev Is Nothing Then
        s = ""
        s = s & "Detectado por: " & Nz(r.DetectadoPor, "") & vbCrLf
        s = s & "Origen: " & Nz(r.CausaRaiz, "") & vbCrLf
        s = s & "Impacto global: " & Nz(r.ImpactoGlobal, "") & vbCrLf
        s = s & "Vulnerabilidad: " & Nz(r.Vulnerabilidad, "") & vbCrLf
        s = s & "Valoración: " & Nz(r.Valoracion, "") & vbCrLf
        s = s & "Mitigación: " & Nz(r.Mitigacion, "") & vbCrLf
        s = s & "Contingencia: " & r.RequierePlanContingencia & vbCrLf
        s = s & "Materializado: " & IIf(IsDate(r.FechaMaterializado), "Sí", "No") & vbCrLf
        s = s & "Estado: " & Nz(r.Estado, "") & vbCrLf
        s = s & "Fecha estado: " & FormatoFecha(r.FechaEstado) & vbCrLf
        s = s & "Priorización: " & Nz(r.Priorizacion, "")
        ConstruirTextoEstadoCambiosHTML = HTMLSafeLargo(s)
        Exit Function
    End If

    s = ""
    If Nz(r.DetectadoPor, "") <> Nz(rPrev.DetectadoPor, "") Then s = s & "Detectado por: " & Nz(r.DetectadoPor, "") & vbCrLf
    If Nz(r.CausaRaiz, "") <> Nz(rPrev.CausaRaiz, "") Then s = s & "Origen: " & Nz(r.CausaRaiz, "") & vbCrLf
    If Nz(r.ImpactoGlobal, "") <> Nz(rPrev.ImpactoGlobal, "") Then s = s & "Impacto global: " & Nz(r.ImpactoGlobal, "") & vbCrLf
    If Nz(r.Vulnerabilidad, "") <> Nz(rPrev.Vulnerabilidad, "") Then s = s & "Vulnerabilidad: " & Nz(r.Vulnerabilidad, "") & vbCrLf
    If Nz(r.Valoracion, "") <> Nz(rPrev.Valoracion, "") Then s = s & "Valoración: " & Nz(r.Valoracion, "") & vbCrLf
    If Nz(r.Mitigacion, "") <> Nz(rPrev.Mitigacion, "") Then s = s & "Mitigación: " & Nz(r.Mitigacion, "") & vbCrLf
    If r.RequierePlanContingencia <> rPrev.RequierePlanContingencia Then s = s & "Contingencia: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
    If Nz(r.Estado, "") <> Nz(rPrev.Estado, "") Then s = s & "Estado: " & Nz(r.Estado, "") & vbCrLf
    If Nz(r.FechaEstado, "") <> Nz(rPrev.FechaEstado, "") Then s = s & "Fecha estado: " & FormatoFecha(r.FechaEstado) & vbCrLf
    If Nz(r.Priorizacion, "") <> Nz(rPrev.Priorizacion, "") Then s = s & "Priorización: " & Nz(r.Priorizacion, "")
    ConstruirTextoEstadoCambiosHTML = HTMLSafeLargo(s)
End Function

Private Function ConstruirTextoPlanesCambiosHTML(ByVal colActual As Scripting.Dictionary, ByVal colPrev As Scripting.Dictionary) As String
    Dim sAct As String, sPrev As String
    sAct = SerializarPlanesHTML(colActual)
    sPrev = SerializarPlanesHTML(colPrev)
    If sAct = sPrev Then
        ConstruirTextoPlanesCambiosHTML = ""
    Else
        ConstruirTextoPlanesCambiosHTML = sAct
    End If
End Function

Private Function SerializarPlanesHTML(ByVal colPlanes As Scripting.Dictionary) As String
    Dim s As String
    Dim k As Variant, kAcc As Variant
    Dim plan As Object, Accion As Object

    On Error Resume Next
    If colPlanes Is Nothing Then
        SerializarPlanesHTML = ""
        Exit Function
    End If
    If colPlanes.Count = 0 Then
        SerializarPlanesHTML = ""
        Exit Function
    End If

    For Each k In colPlanes
        Set plan = colPlanes(k)
        If plan Is Nothing Then GoTo siguientePlan

        s = s & "<div>" & HTMLSafe(Nz(plan.ESTADOCalculadoTexto, "")) & ": " & HTMLSafe(Nz(plan.DisparadorDelPlan, "")) & "</div>"

        If Not plan.colAcciones Is Nothing Then
            For Each kAcc In plan.colAcciones
                Set Accion = plan.colAcciones(kAcc)
                If Not Accion Is Nothing Then
                    s = s & "<div class='cc-plan-action'>- " & HTMLSafe(Nz(Accion.Accion, "")) & " (" & HTMLSafe(Nz(Accion.ResponsableAccion, "")) & ") " & _
                        HTMLSafe(InformeHTML_CompactarFechasAccion(Accion.FechaInicio, Accion.FechaFinPrevista, Accion.FechaFinReal)) & "</div>"
                End If
            Next kAcc
        End If

siguientePlan:
        Set plan = Nothing
    Next k

    SerializarPlanesHTML = Trim$(s)
End Function

Private Function InformeHTML_CompactarFechasAccion( _
                                            ByVal p_FechaInicio As Variant, _
                                            ByVal p_FechaFinPrevista As Variant, _
                                            ByVal p_FechaFinReal As Variant _
                                            ) As String
    Dim partes As String
    Dim sInicio As String
    Dim sFinPrev As String
    Dim sFinReal As String

    sInicio = InformeHTML_FormatoFechaOpcional(p_FechaInicio)
    sFinPrev = InformeHTML_FormatoFechaOpcional(p_FechaFinPrevista)
    sFinReal = InformeHTML_FormatoFechaOpcional(p_FechaFinReal)

    If sInicio <> "" Then partes = sInicio
    If sFinPrev <> "" Then
        If partes <> "" Then partes = partes & " / "
        partes = partes & "Fin prevista: " & sFinPrev
    End If
    If sFinReal <> "" Then
        If partes <> "" Then partes = partes & " / "
        partes = partes & "Fin real: " & sFinReal
    End If

    InformeHTML_CompactarFechasAccion = partes
End Function

Private Function InformeHTML_FormatoFechaOpcional(ByVal p_Valor As Variant) As String
    If IsDate(p_Valor) Then
        InformeHTML_FormatoFechaOpcional = Format$(CDate(p_Valor), "dd/mm/yyyy")
    Else
        InformeHTML_FormatoFechaOpcional = ""
    End If
End Function

Private Function ConstruirSeccionInventarioRiesgosHTML( _
                                                        ByVal p_Edicion As Edicion, _
                                                        ByVal p_FechaCierre As String, _
                                                        ByVal p_FechaPublicacion As String, _
                                                        ByRef p_Error As String, _
                                                        Optional ByVal p_EsPDF As Boolean = False _
                                                        ) As String

    Dim s As String
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As Riesgo
    Dim p As Proyecto
    Dim incluirCausaRaiz As Boolean
    Dim colSpanTotal As Long
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EstadoTexto As String
    Dim m_FechaEstado As Variant
    Dim iRow As Long

    On Error GoTo errores
    p_Error = ""

    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos
    
    incluirCausaRaiz = False
    On Error Resume Next
    Set p = p_Edicion.Proyecto
    If Not p Is Nothing Then
        incluirCausaRaiz = (p.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí)
    End If
    On Error GoTo errores
    
    ' Definimos el colSpan total para el tfoot
    If incluirCausaRaiz Then
        colSpanTotal = 13
    Else
        colSpanTotal = 12
    End If

    If p_EsPDF Then
        s = "<section class='report-section'>" & vbCrLf
        s = s & "  <h2>Inventario de Riesgos Detectados</h2>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    Else
        s = "<details open class='report-section collapsible-section'>" & vbCrLf
        s = s & "  <summary><h2>Inventario de Riesgos Detectados</h2></summary>" & vbCrLf
        s = s & "  <div class='collapsible-body'>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    End If
    
    ' Punto 08 / Fase A — spec Calidad 2026-07-08: mostrar UNA sola
    ' etiqueta + valor, no dos fechas simultáneas. Con publicación
    ' efectiva -> "Fecha Edición" + FechaPublicacion. Sin publicación
    ' -> "Fecha publicación anterior" + p_FechaCierre (fallback a la
    ' fecha de cierre/creación de la edición).
    If IsDate(p_FechaPublicacion) Then
        s = s & "    <div class='meta-row'><div><strong>Fecha Edición:</strong> " & HTMLSafe(Format$(CDate(p_FechaPublicacion), "dd/mm/yyyy")) & "</div></div>" & vbCrLf
    Else
        s = s & "    <div class='meta-row'><div><strong>Fecha publicación anterior:</strong> " & HTMLSafe(p_FechaCierre) & "</div></div>" & vbCrLf
    End If
    s = s & "    <div class='table-scroll'>" & vbCrLf
    s = s & "      <table class='report-table report-table-small'>" & vbCrLf
    
    ' --- CABECERA (thead): Se repetirá en cada página física ---
    s = s & "        <thead>" & vbCrLf
    s = s & "          <tr style='background-color: #D6EAF8; color: #333;'>" & vbCrLf
    If incluirCausaRaiz Then
        s = s & "            <th colspan='4' style='text-align:center; border-right: 1px solid #ccc;'>Identificación</th>" & _
                "<th colspan='7' style='text-align:center; border-right: 1px solid #ccc;'>Análisis</th>" & _
                "<th colspan='2' style='text-align:center;'>Estado</th>" & vbCrLf
    Else
        s = s & "            <th colspan='3' style='text-align:center; border-right: 1px solid #ccc;'>Identificación</th>" & _
                "<th colspan='7' style='text-align:center; border-right: 1px solid #ccc;'>Análisis</th>" & _
                "<th colspan='2' style='text-align:center;'>Estado</th>" & vbCrLf
    End If
    s = s & "          </tr>" & vbCrLf
    
    s = s & "          <tr style='background-color: #ECF0F1; color: #333;'>" & vbCrLf
    s = s & "            <th>Código Riesgo</th><th>Descripción</th>"
    If incluirCausaRaiz Then s = s & "<th>Causa Raíz</th>"
    s = s & "<th>Detectado Por</th><th>Plazo</th><th>Coste</th><th>Calidad</th><th>Global</th>" & _
            "<th>Vulnerabilidad</th><th>Valoración</th><th>Priorización</th><th>Estado</th><th>Fecha</th>" & vbCrLf
    s = s & "          </tr>" & vbCrLf
    s = s & "        </thead>" & vbCrLf

    ' --- CUERPO (tbody) ---
    s = s & "        <tbody>" & vbCrLf
    iRow = 0
    If col Is Nothing Or col.Count = 0 Then
        s = s & "          <tr><td colspan='" & CStr(colSpanTotal) & "'>La edición no tiene riesgos.</td></tr>" & vbCrLf
    Else
        For Each k In col
            Set r = col(k)
            If r Is Nothing Then GoTo siguienteInv
            
            iRow = iRow + 1
            Dim sRowStyle As String
            If iRow Mod 2 = 0 Then
                sRowStyle = " style='background-color: #F4F6F7;'"
            Else
                sRowStyle = " style='background-color: #FFFFFF;'"
            End If
            
            m_Estado = r.EstadoEnum
            m_EstadoTexto = ""
            m_FechaEstado = ""
            InformeHTML_ResolverEstadoCanonico m_Estado, Nz(r.ESTADOCalculadoTexto, ""), p_FechaCierre, r.FechaEstado, m_EstadoTexto, m_FechaEstado

            s = s & "          <tr" & sRowStyle & ">" & _
                    "<td>" & HTMLSafe(Nz(r.CodigoRiesgo, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Descripcion, "")) & "</td>"
            If incluirCausaRaiz Then s = s & "<td>" & HTMLSafe(Nz(r.CausaRaiz, "")) & "</td>"
            s = s & "<td>" & HTMLSafe(Nz(r.DetectadoPor, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Plazo, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Coste, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Calidad, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.ImpactoGlobal, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Vulnerabilidad, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Valoracion, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Priorizacion, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(m_EstadoTexto, "")) & "</td>" & _
                    "<td>" & HTMLSafe(FormatoFecha(m_FechaEstado)) & "</td>" & _
                    "</tr>" & vbCrLf
siguienteInv:
            Set r = Nothing
        Next k
    End If
    s = s & "        </tbody>" & vbCrLf

    ' --- PIE DE TABLA (tfoot): Reserva el espacio para el pie de página fijo ---
    ' Este tfoot evita que el contenido de la tabla pise el footer de 18mm de alto.
    s = s & "        <tfoot>" & vbCrLf
    s = s & "          <tr><td colspan='" & CStr(colSpanTotal) & "' style='height: 18mm; border: none !important; background: transparent !important;'></td></tr>" & vbCrLf
    s = s & "        </tfoot>" & vbCrLf

    s = s & "      </table>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "  </div>" & vbCrLf

    If p_EsPDF Then
        s = s & "  </div>" & vbCrLf
        s = s & "</section>" & vbCrLf
    Else
        s = s & "  </div>" & vbCrLf
        s = s & "</details>" & vbCrLf
    End If

    ConstruirSeccionInventarioRiesgosHTML = s
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "Error en ConstruirSeccionInventarioRiesgosHTML: " & err.description
    End If
End Function

Private Function ConstruirSeccionFichasRiesgoHTML( _
                                                ByVal p_Edicion As Edicion, _
                                                ByVal p_Proyecto As Proyecto, _
                                                ByVal p_FechaCierre As String, _
                                                ByVal p_FechaPublicacion As String, _
                                                ByRef p_Error As String, _
                                                Optional ByVal p_EsPDF As Boolean = False _
                                                ) As String

    Dim s As String
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As Riesgo

    On Error GoTo errores
    p_Error = ""

    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos

    If col Is Nothing Or col.Count = 0 Then
        ConstruirSeccionFichasRiesgoHTML = ""
        Exit Function
    End If

    For Each k In col
        Set r = col(k)
        If r Is Nothing Then GoTo siguienteFicha
        Avance "Riesgos " & r.CodigoRiesgo & " ..."
        s = s & ConstruirFichaRiesgoHTML(r, p_Proyecto, p_FechaCierre, p_FechaPublicacion, p_EsPDF)
siguienteFicha:
        Set r = Nothing
    Next k

    ConstruirSeccionFichasRiesgoHTML = s
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "Error en ConstruirSeccionFichasRiesgoHTML: " & err.description
    End If
End Function

Private Function ConstruirFichaRiesgoHTML( _
                                        ByVal p_Riesgo As Riesgo, _
                                        ByVal p_Proyecto As Proyecto, _
                                        ByVal p_FechaCierre As String, _
                                        ByVal p_FechaPublicacion As String, _
                                        Optional ByVal p_EsPDF As Boolean = False _
                                        ) As String
    Dim s As String
    Dim sError As String

    If p_EsPDF Then
        s = "<section class='report-section print-page'>" & vbCrLf
        s = s & "  <h2>Ficha de riesgo: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</h2>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    Else
        s = "<details open class='report-section collapsible-section'>" & vbCrLf
        s = s & "  <summary><h2>Ficha de riesgo: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</h2></summary>" & vbCrLf
        s = s & "  <div class='collapsible-body'>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    End If
    
    s = s & "    <div class='meta-row'><div><strong>Proyecto:</strong> " & HTMLSafe(Nz(p_Proyecto.NombreProyecto, Nz(p_Proyecto.Proyecto, ""))) & "</div><div><strong>Fecha:</strong> " & HTMLSafe(Format$(Date, "dd/mm/yyyy")) & "</div></div>" & vbCrLf
    
            ' 1. Histórico de Estados (Moved up)
    s = s & "    <div class='card-mini' style='margin-bottom:25px;'>" & vbCrLf
    s = s & ConstruirTablaEstadosHistoricosHTML(p_Riesgo, p_FechaCierre, p_FechaPublicacion, sError)
    s = s & "    </div>" & vbCrLf

    ' 2. Datos del Riesgo (Tabla Horizontal Centrada)
    s = s & "    <div style='margin-bottom:25px; text-align: center;'>" & vbCrLf
    s = s & "      <h3 class='ficha-subtitle' style='margin-top:0;'>Datos del Riesgo</h3>" & vbCrLf
    s = s & "      <table class='report-table' style='width:100%; text-align:center; margin: 0 auto;'>" & vbCrLf
    s = s & "        <thead>" & vbCrLf
    s = s & "          <tr style='background-color: #ECF0F1;'>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Código</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Detectado por</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Origen</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Impacto</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Vuln.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Valor.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Mitig.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Conting.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Mat.</th>" & vbCrLf
    s = s & "          </tr>" & vbCrLf
    s = s & "        </thead>" & vbCrLf
    s = s & "        <tbody>" & vbCrLf
    s = s & "          <tr>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.CodigoRiesgo, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.DetectadoPor, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Origen, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.ImpactoGlobal, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Vulnerabilidad, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Valoracion, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Mitigacion, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.ContingenciaCalculada, "")) & "</td>" & vbCrLf
    
    Dim sFechaMat As String, sEstiloMat As String
    sFechaMat = ""
    sEstiloMat = ""
    If IsDate(p_Riesgo.FechaMaterializado) Then
        sFechaMat = Format(p_Riesgo.FechaMaterializado, "dd/mm/yyyy")
        sEstiloMat = " style='color:red;'"
    End If
    s = s & "            <td" & sEstiloMat & ">" & HTMLSafe(sFechaMat) & "</td>" & vbCrLf
    s = s & "          </tr>" & vbCrLf
    s = s & "        </tbody>" & vbCrLf
    s = s & "      </table>" & vbCrLf
    
    s = s & "    </div>" & vbCrLf

    s = s & "    <div style='margin-bottom:25px;'>" & vbCrLf
    s = s & "      <h3 class='ficha-subtitle' style='margin-top:0;'>Descripción</h3>" & vbCrLf
    s = s & "      <div class='description-text'>" & HTMLSafeLargo(Nz(p_Riesgo.Descripcion, "")) & "</div>" & vbCrLf
    
    If Nz(p_Riesgo.CausaRaiz, "") <> "" Then
        s = s & "      <h3 class='ficha-subtitle' style='margin-top:16px;'>Causa raíz</h3>" & vbCrLf
        s = s & "      <div class='description-text'>" & HTMLSafeLargo(Nz(p_Riesgo.CausaRaiz, "")) & "</div>" & vbCrLf
    End If
    
    s = s & "    </div>" & vbCrLf

    Dim planesMit As String, planesCont As String
    planesMit = ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Mitigacion)
    planesMit = Replace(planesMit, "<div class='card'", "<div")
    planesMit = Replace(planesMit, "border-left: 6px solid var(--tele-blue);", "")
    planesMit = Replace(planesMit, "border: 1px solid var(--grey-2);", "")
    planesMit = Replace(planesMit, "margin-bottom: 25px;", "")
    s = s & "    <div style='margin-bottom:25px;'>" & vbCrLf
    s = s & "      " & planesMit & vbCrLf
    s = s & "    </div>" & vbCrLf
    planesCont = ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Contingencia)
    planesCont = Replace(planesCont, "<div class='card'", "<div")
    planesCont = Replace(planesCont, "border-left: 6px solid var(--tele-blue);", "")
    planesCont = Replace(planesCont, "border: 1px solid var(--grey-2);", "")
    planesCont = Replace(planesCont, "margin-bottom: 25px;", "")
    s = s & "    <div style='margin-bottom:25px;'>" & vbCrLf
    s = s & "      " & planesCont & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "    </div>" & vbCrLf

    If p_EsPDF Then
        s = s & "  </div>" & vbCrLf
        s = s & "</section>" & vbCrLf
    Else
        s = s & "  </div>" & vbCrLf
        s = s & "  </div>" & vbCrLf
        s = s & "</details>" & vbCrLf
    End If

    ConstruirFichaRiesgoHTML = s
End Function

Public Function GetEstilosCSS_InformeEdicion_Print(Optional ByVal p_EsPDF As Boolean = False) As String
    Dim css As String
    css = "<style>" & vbCrLf
    
    ' --- 1. ESTILOS DE TÍTULOS Y TARJETAS (PANTALLA Y PDF) ---
    css = css & "h2 { color: var(--grey-9); border-bottom: 2px solid var(--tele-blue); padding-bottom: 8px; margin-bottom: 20px; font-size: 16px; text-transform: none; }" & vbCrLf
    css = css & ".card { background: white; padding: 18px; border-radius: 18px; border: 1px solid var(--grey-2); margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }" & vbCrLf
    css = css & ".report-cover-kv { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; }" & vbCrLf
    css = css & ".report-section { margin-bottom: 60px; }" & vbCrLf
    css = css & ".meta-row { margin-bottom: 25px; }" & vbCrLf
    css = css & ".kv { border: 1px solid var(--grey-2); border-radius: 12px; padding: 12px 14px; background: white; }" & vbCrLf
    css = css & ".kv-k { font-size: 10px; color: var(--grey-6); font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }" & vbCrLf
    css = css & ".kv-v { margin-top: 4px; font-size: 12px; color: var(--grey-9); font-weight: 400; }" & vbCrLf
    If Not p_EsPDF Then
        css = css & "details.collapsible-section { margin-bottom: 0; }" & vbCrLf
        css = css & "details.collapsible-section > summary { list-style: none; cursor: pointer; display: flex; align-items: center; gap: 10px; user-select: none; }" & vbCrLf
        css = css & "details.collapsible-section > summary::-webkit-details-marker { display: none; }" & vbCrLf
        css = css & "details.collapsible-section > summary::before { display: none; }" & vbCrLf
        css = css & "details.collapsible-section > summary h2 { margin-bottom: 0; flex: 1; }" & vbCrLf
        css = css & "details[open].collapsible-section > summary h2 { margin-bottom: 0; padding-bottom: 8px; }" & vbCrLf
        css = css & "details.collapsible-section > .collapsible-body { padding-top: 20px; animation: fadeIn 0.3s ease; }" & vbCrLf
    End If
    
    ' --- 2. TABLAS ---
    css = css & ".report-table { width: 100%; border-collapse: collapse; background: white; margin-bottom: 10px; }" & vbCrLf
    css = css & ".report-table th { text-align: left; padding: 12px; border-bottom: 2px solid var(--grey-2); background: #fcfcfc; color: var(--grey-9); }" & vbCrLf
    css = css & ".report-table td { padding: 12px; border-bottom: 1px solid #f6f6f6; vertical-align: top; color: var(--grey-6); font-weight: 400; }" & vbCrLf
    css = css & ".report-table-small th, .report-table-small td{ font-size: 12px; }" & vbCrLf
    
    css = css & ".logo-wrapper { width: 350px; height: auto; display: flex; align-items: center; }" & vbCrLf
    css = css & ".cc-added { color: #0b7a28; padding: 2px; }" & vbCrLf
    css = css & ".cc-deleted { color: #b42318; padding: 2px; }" & vbCrLf
    css = css & ".ficha-action { margin-left: 0; display: block; }" & vbCrLf
    css = css & ".cc-plan-action { margin-left: 1cm; display: block; }" & vbCrLf
    css = css & ".table-scroll { overflow-x: visible; }" & vbCrLf
    css = css & ".legend-container { display: flex; align-items: center; gap: 10px; margin: 10px 0 20px; }" & vbCrLf
    css = css & ".legend-item { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--grey-6); }" & vbCrLf
    css = css & ".dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }" & vbCrLf
    css = css & ".dot.dot-new { background: #0b7a28; }" & vbCrLf
    css = css & ".dot.dot-unchanged { background: #58617A; }" & vbCrLf
    css = css & ".dot.dot-deleted { background: #b42318; }" & vbCrLf
    css = css & ".plans-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }" & vbCrLf
    css = css & ".cc-table { table-layout: fixed; }" & vbCrLf
    css = css & ".cc-table col.cc-col-codigo { width: 90px; }" & vbCrLf
    css = css & ".cc-table col.cc-col-edicion { width: 60px; }" & vbCrLf
    css = css & ".cc-table col.cc-col-estado { width: 160px; }" & vbCrLf
    css = css & ".cc-table col.cc-col-pm { width: 50%; }" & vbCrLf
    css = css & ".cc-table col.cc-col-pc { width: 50%; }" & vbCrLf
    css = css & ".action-table tr.ficha-action-row td { padding-left: 0; }" & vbCrLf
    css = css & ".cc-table td:nth-child(4), .cc-table td:nth-child(5) { padding-left: 1cm; }" & vbCrLf
    
    ' --- 3. CONTROL VISUAL (Web vs Print) ---
    css = css & ".print-only-cover { display: none; }" & vbCrLf
    css = css & ".print-footer { display: none; }" & vbCrLf
    css = css & ".logo-wrapper svg .logo-path { fill: white !important; }" & vbCrLf

    ' --- 4. CONFIGURACIÓN EXCLUSIVA IMPRESIÓN (A4 Landscape optimizado) ---
    css = css & "@media print {" & vbCrLf
    
    ' 1. ELIMINAR ANIMACIONES (¡Culpables de que el contenido sea invisible!)
    ' Forzamos opacidad a 1 y quitamos transiciones para que Edge no imprima un frame vacío
    css = css & "  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; box-sizing: border-box !important; " & _
                "font-family: Arial, Helvetica, sans-serif !important; -webkit-text-stroke: 0.2px transparent !important; " & _
                "animation: none !important; transition: none !important; opacity: 1 !important; visibility: visible !important; }" & vbCrLf
                
    ' 2. MÁRGENES NATIVOS DE A4: Devolvemos el control a Edge para que pagine bien (Sin hojas en blanco)
    css = css & "  @page { size: A4 landscape; margin: 15mm !important; }" & vbCrLf
    
    ' 3. BODY Y HTML NATURALES
    css = css & "  html, body { background: white !important; color: #1a1a1a !important; font-size: 11px !important; line-height: 1.35 !important; " & _
                "margin: 0 !important; padding: 0 !important; width: 100% !important; height: auto !important; " & _
                "overflow: visible !important; display: block !important; }" & vbCrLf
    
    ' 4. OCULTAR UI
    If p_EsPDF Then
        css = css & "  .tab-container, .subtab-container, .tab-btn, .subtab-btn, footer, .print-footer-right, .screen-only-header { display: none !important; }" & vbCrLf
    Else
        css = css & "  .section-search-wrapper, .tab-container, .subtab-container, .tab-btn, .subtab-btn, footer, .print-footer-right, .screen-only-header { display: none !important; }" & vbCrLf
    End If
    
    ' 5. CONTENIDO DE PESTAÑAS: VISIBLE SIEMPRE
    css = css & "  .tab-content, .subtab-content { display: block !important; border: none !important; box-shadow: none !important; padding: 0 !important; margin: 0 0 20px 0 !important; }" & vbCrLf
    If Not p_EsPDF Then
        css = css & "  details, details.collapsible-section { display: block !important; overflow: visible !important; page-break-inside: auto !important; break-inside: auto !important; }" & vbCrLf
        css = css & "  details[open] > summary ~ *, details > summary ~ * { display: block !important; }" & vbCrLf
        css = css & "  .collapsible-body { display: block !important; padding-top: 10px !important; overflow: visible !important; }" & vbCrLf
    End If
    
    ' 6. PORTADA
    css = css & "  .print-only-cover { display: flex !important; flex-direction: column !important; justify-content: flex-start !important; align-items: stretch !important; min-height: 90vh !important; text-align: center !important; background: white !important; page-break-after: always !important; break-after: page !important; overflow: visible !important; }" & vbCrLf
    css = css & "  .print-only-cover header { display: flex !important; justify-content: space-between !important; background: white !important; color: #1a1a1a !important; border-bottom: 2px solid #0066FF !important; padding: 8px 0 !important; margin-bottom: 30px !important; }" & vbCrLf
    css = css & "  .print-only-cover header .header-info { text-align: right !important; color: #1a1a1a !important; }" & vbCrLf
    css = css & "  .print-only-cover header .logo-wrapper svg .logo-path { fill: #0066FF !important; }" & vbCrLf
    css = css & "  .print-only-cover .cover-titles { margin: auto 0 !important; width: 100% !important; }" & vbCrLf
    
    ' 7. PIE DE PÁGINA ANCLADO ABAJO
    css = css & "  .print-footer { display: block !important; position: fixed !important; bottom: 0 !important; right: 0 !important; font-size: 9px !important; color: #666 !important; text-align: right !important; z-index: 9999 !important; }" & vbCrLf
    
    ' 8. CONTENEDORES Y TABLAS
    css = css & "  .container, .main-container { width: 100% !important; padding: 0 !important; margin: 0 !important; overflow: visible !important; }" & vbCrLf
    css = css & "  .card { box-shadow: none !important; border: 1px solid #ddd !important; border-radius: 6px !important; padding: 12px !important; margin-bottom: 12px !important; break-inside: auto !important; page-break-inside: auto !important; }" & vbCrLf
    css = css & "  .card-mini { break-inside: avoid !important; page-break-inside: avoid !important; }" & vbCrLf
    css = css & "  table, .action-table, .report-table, .cc-table { width: 100% !important; border-collapse: collapse !important; font-size: 10px !important; }" & vbCrLf
    css = css & "  tr { break-inside: avoid !important; page-break-inside: avoid !important; }" & vbCrLf
    css = css & "  th, td { padding: 5px 6px !important; font-size: 10px !important; border-bottom: 1px solid #e0e0e0 !important; word-wrap: break-word !important; }" & vbCrLf
    css = css & "  th { background: #f5f5f5 !important; font-weight: bold !important; border-bottom: 2px solid #ccc !important; }" & vbCrLf
    
    ' 9. TÍTULOS Y SALTOS
    css = css & "  .report-section { margin-bottom: 20px !important; break-before: page !important; page-break-before: always !important; overflow: visible !important; }" & vbCrLf
    css = css & "  .report-section:first-of-type { break-before: auto !important; page-break-before: auto !important; }" & vbCrLf
    css = css & "  h2 { font-size: 14px !important; margin-bottom: 10px !important; padding-bottom: 4px !important; break-after: avoid !important; color: #1a1a1a !important; overflow: visible !important; }" & vbCrLf
    css = css & "  h3 { font-size: 11px !important; break-after: avoid !important; }" & vbCrLf
    css = css & "}" & vbCrLf
    
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_InformeEdicion_Print = css
End Function
Public Function GetEstilosCSS_Correos() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & "html, body { margin:0; padding:0; height:100%; }" & vbCrLf
    css = css & "body, table, td { font-family: 'Segoe UI', Arial, sans-serif; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%; }" & vbCrLf
    css = css & "table { border-collapse: collapse !important; mso-table-lspace:0pt; mso-table-rspace:0pt; width:100%; background:white; }" & vbCrLf
    css = css & "img { -ms-interpolation-mode: bicubic; border:0; outline:none; text-decoration:none; }" & vbCrLf
    css = css & "p { margin:0 0 12px 0; }" & vbCrLf
    css = css & "a { color: #0066FF; text-decoration: none; }" & vbCrLf
    css = css & "a:hover { text-decoration: underline; }" & vbCrLf
    css = css & ".report-table { width: 100%; border-collapse: collapse !important; background: white; }" & vbCrLf
    css = css & ".report-table th { text-align: left; padding: 12px; border-bottom: 2px solid #D1D5E4; background: #fcfcfc; color: #031A34; }" & vbCrLf
    css = css & ".report-table td { padding: 12px; border-bottom: 1px solid #f6f6f6; vertical-align: top; color: #58617A; }" & vbCrLf
    css = css & "td.Cabecera, td.ColespanArriba { background: #fcfcfc; color: #031A34; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; font-size: 12px; }" & vbCrLf
    css = css & "td.ColespanArriba { text-align: left; }" & vbCrLf
    css = css & "td.centrado { text-align: center; }" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_Correos = css
End Function

Private Function GuardarInformeEdicionHTML_UTF8(ByVal p_Edicion As Edicion, ByVal p_HTML As String, ByRef p_Error As String) As String
    Dim m_Ruta As String
    Dim m_Stream As Object
    Dim m_FSO As Object

    On Error GoTo errores
    p_Error = ""

    m_Ruta = GetURLInformeEdicionHTML(p_Edicion, p_Error)
    If p_Error <> "" Then err.Raise 1000

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    If m_FSO.FileExists(m_Ruta) Then
        If FicheroAbierto(m_Ruta) Then
            p_Error = "Tiene abierto un informe anterior"
            err.Raise 1000
        End If
    End If

    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2
    m_Stream.Charset = "utf-8"
    m_Stream.Open
    m_Stream.WriteText p_HTML
    m_Stream.SaveToFile m_Ruta, 2
    m_Stream.Close

    GuardarInformeEdicionHTML_UTF8 = m_Ruta
    Exit Function
errores:
    If err.Number <> 1000 Then
        p_Error = "Error UTF-8: " & err.description
    End If
End Function

Private Function GetURLInformeEdicionHTML( _
                                        ByVal p_Edicion As Edicion, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_Cod As String
    Dim m_Proyecto As Proyecto

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "La edición hay que introducirla"
        err.Raise 1000
    End If

    Set m_Proyecto = p_Edicion.Proyecto
    If m_Proyecto Is Nothing Then
        Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
        If p_Error <> "" Then err.Raise 1000
    End If

    m_Cod = Nz(m_Proyecto.CodigoDocumento, "")
    If m_Cod = "" Then
        p_Error = "No se sabe el código del documento del proyecto al que pertenece el informe"
        err.Raise 1000
    End If

    m_Cod = Replace(m_Cod, "/", "_")
    m_Cod = Replace(m_Cod, "\", "_")
    m_Cod = Replace(m_Cod, vbNewLine, "")

    If m_Proyecto.Juridica = "TdE" Then
        GetURLInformeEdicionHTML = m_ObjEntorno.URLDirectorioLocal & m_Cod & "V" & Format(p_Edicion.Edicion, "00") & ".html"
    Else
        GetURLInformeEdicionHTML = m_ObjEntorno.URLDirectorioLocal & m_Cod & "-" & p_Edicion.Edicion & ".html"
    End If

    Exit Function
errores:
    If err.Number <> 1000 Then
        p_Error = "El método GetURLInformeEdicionHTML ha devuelto el error: " & vbNewLine & err.description
    End If
End Function

' -------------------------------------------------------------------------
' COPIA TU CÓDIGO SVG AQUÍ
' -------------------------------------------------------------------------
Public Function GetLogoSVG(Optional ByVal pWidth As String = "100%", Optional ByVal pHeight As String = "auto") As String
    ' SVG Logo Telefónica - Solo trazado
    ' Parámetros admiten valores CSS (px, %, rem, auto)
    Dim svg As String
    
    svg = "<svg id='Layer_1' data-name='Layer 1' xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1920 802' " & _
          "style='width:" & pWidth & "; height:" & pHeight & ";'>" & _
          "<defs><style>.logo-path{fill:#0066FF; transition: fill 0.3s;}</style></defs>" & _
          "<g><g>" & _
          "<circle class='logo-path' cx='275.37' cy='277.85' r='52.41'/>" & _
          "<circle class='logo-path' cx='398.52' cy='277.85' r='52.41'/>" & _
          "<circle class='logo-path' cx='521.67' cy='277.85' r='52.41'/>" & _
          "<circle class='logo-path' cx='398.52' cy='401' r='52.41'/>" & _
          "<circle class='logo-path' cx='398.52' cy='524.15' r='52.41'/>" & _
          "</g><g>" & _
          "<path class='logo-path' d='m685.79,360.16h-48.84v-29.07h127.9v29.07h-48.84v133.72h-30.23v-133.72Z'/>" & _
          "<path class='logo-path' d='m866.94,458.99c-2.56,9.3-17.21,37.21-54.65,37.21-34.88,0-60.47-25.58-60.47-61.63s25.58-61.63,60.47-61.63c32.56,0,58.14,25.58,58.14,59.3,0,3.49-.47,6.28-.7,8.37l-.47,3.26h-88.37c2.56,16.51,14.88,27.91,31.4,27.91,13.72,0,22.09-7.9,24.42-12.79h30.23Zm-25.58-34.88c-2.56-15.12-12.56-26.75-29.07-26.75-17.67,0-27.91,11.63-31.4,26.75h60.47Z'/>" & _
          "<path class='logo-path' d='m887.85,331.09h29.07v162.79h-29.07v-162.79Z'/>" & _
          "<path class='logo-path' d='m1049.46,458.99c-2.56,9.3-17.21,37.21-54.65,37.21-34.88,0-60.46-25.58-60.46-61.63s25.58-61.63,60.46-61.63c32.56,0,58.14,25.58,58.14,59.3,0,3.49-.47,6.28-.7,8.37l-.47,3.26h-88.37c2.56,16.51,14.88,27.91,31.4,27.91,13.72,0,22.09-7.9,24.42-12.79h30.23Zm-25.58-34.88c-2.56-15.12-12.56-26.75-29.07-26.75-17.67,0-27.91,11.63-31.4,26.75h60.47Z'/>" & _
          "<path class='logo-path' d='m1079.68,403.18h-19.77v-27.91h19.77v-17.44c0-17.67,11.4-29.07,29.07-29.07h25.58v25.58h-17.44c-4.65,0-8.14,3.49-8.14,8.14v12.79h25.58v27.91h-25.58v90.7h-29.07v-90.7Z'/>" & _
          "<path class='logo-path' d='m1259.61,434.58c0,36.05-25.58,61.63-60.47,61.63s-60.47-25.58-60.47-61.63,25.58-61.63,60.47-61.63,60.47,25.58,60.47,61.63Zm-29.07,0c0-20.93-13.95-34.88-31.4-34.88s-31.4,13.95-31.4,34.88,13.95,34.88,31.4,34.88,31.4-13.95,31.4-34.88Z'/>" & _
          "<path class='logo-path' d='m1277.03,375.28h26.75l2.32,11.63h1.16c2.09-2.56,4.89-4.88,7.91-6.98,5.35-3.49,13.49-6.98,24.65-6.98,26.74,0,46.51,19.77,46.51,50v70.93h-29.07v-68.6c0-15.12-10.46-25.58-25.58-25.58s-25.58,10.46-25.58,25.58v68.6h-29.07v-118.6Z'/>" & _
          "<path class='logo-path' d='m1571.36,449.69c-3.72,18.6-18.37,46.51-55.81,46.51-34.88,0-60.46-25.58-60.46-61.63s25.58-61.63,60.46-61.63c37.44,0,52.09,27.91,55.81,45.35h-29.07c-2.56-6.75-9.3-18.6-26.75-18.6s-31.4,13.95-31.4,34.88,13.95,34.88,31.4,34.88,24.19-11.63,26.75-19.77h29.07Z'/>" & _
          "<path class='logo-path' d='m1659.26,482.25h-1.16c-2.09,2.56-4.88,4.88-8.14,6.98-5.58,3.49-13.72,6.98-25.58,6.98-26.97,0-43.02-16.51-43.02-36.04,0-23.26,16.28-39.54,48.83-39.54h26.75v-2.32c0-13.02-7.91-22.09-20.93-22.09s-19.77,8.37-20.93,15.12h-29.07c2.56-19.54,18.37-38.37,50-38.37s50,20,50,45.35v75.58h-24.42l-2.32-11.63Zm-2.32-39.53h-24.42c-15.12,0-22.09,5.81-22.09,15.11s6.75,15.12,18.6,15.12c17.67,0,27.91-10.23,27.91-26.75v-3.49Z'/>" & _
          "<path class='logo-path' d='m1408.59,375.28h29.07v118.6h-29.07v-118.6Z'/>" & _
          "<circle class='logo-path' cx='1423.13' cy='341.02' r='17.2'/>" & _
          "<polygon class='logo-path' points='1209.61 325.28 1239.85 325.28 1211.93 359 1187.52 359 1209.61 325.28'/>" & _
          "</g></g></svg>"
          
    GetLogoSVG = svg
End Function
Private Function ConstruirTablaEstadosHistoricosHTML( _
                                                    ByVal p_Riesgo As Riesgo, _
                                                    ByVal p_FechaCierre As String, _
                                                    ByVal p_FechaPublicacion As String, _
                                                    ByRef p_Error As String _
                                                    ) As String
    Dim ColEstados As Scripting.Dictionary
    Dim k As Variant
    Dim s As String
    Dim sEstado As String, sFecha As String
    Dim partes As Variant
    Dim m_FechaPub As String

    On Error GoTo errores
    
    If Not IsDate(p_FechaPublicacion) Then
        m_FechaPub = Format(Date, "dd/mm/yyyy")
    Else
        m_FechaPub = p_FechaPublicacion
    End If

    Set ColEstados = getEstadosDiferentesHastaEdicion(p_Riesgo.Edicion, p_Riesgo.CodigoRiesgo, m_FechaPub, p_FechaCierre, p_Error)
    If p_Error <> "" Then Exit Function
    If ColEstados Is Nothing Then Exit Function
    If ColEstados.Count = 0 Then Exit Function

    s = ""
    s = s & "  <h3 style='margin-top:0;'>Histórico de Estados</h3>" & vbCrLf
    s = s & "  <table class='report-table report-table-small' style='width:100%;'>" & vbCrLf
    s = s & "    <thead><tr><th>Estado</th><th>Fecha</th></tr></thead>" & vbCrLf
    s = s & "    <tbody>" & vbCrLf

    For Each k In ColEstados
        partes = Split(ColEstados(k), "|")
        sEstado = partes(0)
        sFecha = ""
        If UBound(partes) >= 1 Then sFecha = partes(1)
        
        If IsDate(sFecha) Then sFecha = Format(sFecha, "dd/mm/yyyy")
        
        s = s & "      <tr>" & vbCrLf
        s = s & "        <td>" & HTMLSafe(sEstado) & "</td>" & vbCrLf
        s = s & "        <td>" & HTMLSafe(sFecha) & "</td>" & vbCrLf
        s = s & "      </tr>" & vbCrLf
    Next k

    s = s & "    </tbody>" & vbCrLf
    s = s & "  </table>" & vbCrLf

    ConstruirTablaEstadosHistoricosHTML = s
    Exit Function

errores:
    p_Error = "Error en ConstruirTablaEstadosHistoricosHTML: " & err.description
End Function

Private Function ConstruirSeccionGeneral(p_Riesgo As Riesgo) As String
    Dim s As String
    Dim sTecnico As String
    Dim sEdicion As String, sEsActivo As String
    Dim sTituloCausaRaiz As String
    Dim sCausaRaiz As String
    Dim sEstadoCanonico As String
    Dim sFechaEstadoCanonica As String
    
    ' Obtención segura del técnico de calidad
    On Error Resume Next
    sTecnico = p_Riesgo.Edicion.Proyecto.UsuarioCalidad.Nombre
    If err.Number <> 0 Then sTecnico = "-"
    On Error GoTo 0
    
    ' Obtención segura de Edición
    On Error Resume Next
    sEdicion = p_Riesgo.Edicion.Edicion
    If p_Riesgo.Edicion.EsActivo = 1 Then sEsActivo = "Sí" Else sEsActivo = "No"
    On Error GoTo 0
    
    InformeHTML_ResolverEstadoCanonico p_Riesgo.EstadoEnum, Nz(p_Riesgo.ESTADOCalculadoTexto, ""), Nz(p_Riesgo.Edicion.Proyecto.FechaCierre, ""), p_Riesgo.FechaEstado, sEstadoCanonico, sFechaEstadoCanonica

    ' Fila 1: KPIs Principales (Compacto - Una sola fila)
    s = "<div class='grid-compact' style='margin-bottom:20px;'>" & vbCrLf
    
    ' 1. Estado
    s = s & "    <div class='card-mini'><h3>Estado</h3><div class='kpi-value-mini' style='color:var(--tele-blue);'>" & HTMLSafe(sEstadoCanonico) & "</div></div>" & vbCrLf
    ' 2. Priorización
    s = s & "    <div class='card-mini'><h3>Priorización</h3><div class='kpi-value-mini'>" & p_Riesgo.Priorizacion & "</div></div>" & vbCrLf
    ' 3. Impacto Global
    s = s & "    <div class='card-mini'><h3>Impacto Global</h3><div class='kpi-value-mini'>" & HTMLSafe(p_Riesgo.ImpactoGlobalCalculado) & "</div></div>" & vbCrLf
    ' 4. Origen
    s = s & "    <div class='card-mini'><h3>Origen</h3><div class='kpi-value-mini' style='font-weight:normal;'>" & HTMLSafe(p_Riesgo.Origen) & "</div></div>" & vbCrLf
    ' 5. Edición
    s = s & "    <div class='card-mini'><h3>Edición</h3><div class='kpi-value-mini'>" & HTMLSafe(sEdicion) & "</div></div>" & vbCrLf
    ' 6. ¿Activa?
    s = s & "    <div class='card-mini'><h3>¿Activa?</h3><div class='kpi-value-mini' style='color:" & IIf(sEsActivo = "Sí", "#00C853", "#D50000") & ";'>" & sEsActivo & "</div></div>" & vbCrLf
    
    s = s & "</div>" & vbCrLf
    
    ' Fila 2: Fechas y Personas
    s = s & "<div class='grid' style='margin-top:20px;'>" & vbCrLf
    s = s & "    <div class='card'><h4>Detalles de Detección</h4>" & _
            "<p><strong>Fecha Detectado:</strong> " & FormatoFecha(p_Riesgo.FechaDetectado) & "</p>" & _
            "<p><strong>Detectado Por:</strong> " & HTMLSafe(p_Riesgo.DetectadoPor) & "</p></div>" & vbCrLf
            
    s = s & "    <div class='card'><h4>Gestión y Calidad</h4>" & _
            "<p><strong>Técnico Calidad:</strong> " & HTMLSafe(sTecnico) & "</p>" & _
            "<p><strong>Fecha Aceptado:</strong> " & FormatoFecha(p_Riesgo.FechaMitigacionAceptar) & "</p>" & _
            "<p><strong>Fecha Retirado:</strong> " & FormatoFecha(p_Riesgo.FechaRetirado) & "</p></div>" & vbCrLf
    s = s & "</div>" & vbCrLf

    ' Fila 3: Textos Largos
    s = s & "<div class='card' style='margin-top:20px;'>" & vbCrLf
    s = s & "    <h2>Descripción de Riesgo</h2><p>" & HTMLSafeLargo(p_Riesgo.Descripcion) & "</p>" & vbCrLf
    sCausaRaiz = Nz(p_Riesgo.CausaRaiz, "")
    sTituloCausaRaiz = "Análisis Causa Raíz"
    If Trim$(sCausaRaiz) = "" Then
        sTituloCausaRaiz = sTituloCausaRaiz & " (No aplica)"
        sCausaRaiz = "-"
    End If
    s = s & "    <h2 style='margin-top:20px;'>" & sTituloCausaRaiz & "</h2><p>" & HTMLSafeLargo(sCausaRaiz) & "</p>" & vbCrLf
    s = s & "</div>" & vbCrLf
    
    ConstruirSeccionGeneral = s
End Function

Private Sub InformeHTML_ResolverEstadoCanonico( _
                                        ByVal p_EstadoEnum As Long, _
                                        ByVal p_EstadoCalculado As String, _
                                        ByVal p_FechaCierre As String, _
                                        ByVal p_FechaEstadoOriginal As Variant, _
                                        ByRef p_EstadoTexto As String, _
                                        ByRef p_FechaEstado As Variant _
                                        )
    p_EstadoTexto = getEstadoRiesgoTexto(p_EstadoEnum)
    If IsNull(p_FechaEstadoOriginal) Then
        p_FechaEstado = ""
    Else
        p_FechaEstado = p_FechaEstadoOriginal
    End If

    If IsDate(p_FechaCierre) Then
        If p_EstadoEnum = EnumRiesgoEstado.Retirado Or p_EstadoEnum = EnumRiesgoEstado.Aceptado Or p_EstadoEnum = EnumRiesgoEstado.Materializado Or p_EstadoEnum = EnumRiesgoEstado.Cerrado Then
            p_EstadoTexto = "Cerrado"
            If Not IsDate(p_FechaEstado) Then p_FechaEstado = p_FechaCierre
        End If
    ElseIf Not IsDate(p_FechaEstado) And Trim$(p_EstadoCalculado) <> "" Then
        p_EstadoTexto = p_EstadoCalculado
    End If

    If Trim$(p_EstadoTexto) = "" Then p_EstadoTexto = Nz(p_EstadoCalculado, "-")
    If Not IsDate(p_FechaEstado) Then p_FechaEstado = ""
End Sub

Private Function ConstruirSeccionPlanes(p_Riesgo As Riesgo, Tipo As EnumTipoPlan, Optional ByRef p_Error As String) As String
    Dim s As String, colPlanes As Scripting.Dictionary, m_IDP As Variant, m_IDPA As Variant
    Dim m_Plan As Object, m_Accion As Object, m_NombrePropiedad As String
    Dim prefixAccion As String, idAccion As String, labelAccion As String
    On Error GoTo errores
    p_Error = ""
    ' Título de la sección y Selección de Colección
    If Tipo = Mitigacion Then
        Set colPlanes = p_Riesgo.colPMs
        s = "<h2>Planes de Mitigación</h2>"
    Else
        Set colPlanes = p_Riesgo.colPCs
        s = "<h2>Planes de Contingencia</h2>"
    End If
    
    If colPlanes Is Nothing Then
        s = s & "<div class='card'><p>No hay planes registrados para este tipo.</p></div>"
    ElseIf colPlanes.Count = 0 Then
        s = s & "<div class='card'><p>No hay planes registrados para este tipo.</p></div>"
    Else
        For Each m_IDP In colPlanes
            Set m_Plan = colPlanes(m_IDP)
            m_NombrePropiedad = "Cod" & IIf(Tipo = Mitigacion, "Mitigacion", "Contingencia")
            
            s = s & "<div class='card' style='border-left: 6px solid var(--tele-blue); margin-bottom: 25px;'>" & vbCrLf
            
            ' Cabecera del Plan con Estado
            s = s & "    <h3 style='color: var(--tele-blue);'>" & _
                    "<span>" & HTMLSafe(m_Plan.getPropiedad(m_NombrePropiedad)) & ": " & HTMLSafe(m_Plan.DisparadorDelPlan) & "</span>" & _
                    "<span class='plan-status-badge'>" & HTMLSafe(m_Plan.ESTADOCalculadoTexto) & "</span>" & _
                    "</h3>" & vbCrLf
            
            If Not m_Plan.colAcciones Is Nothing Then
                If m_Plan.colAcciones.Count > 0 Then
                    s = s & "    <div class='ficha-action'>" & vbCrLf
                    s = s & "    <table class='action-table'>" & vbCrLf
                    s = s & "        <thead><tr>" & _
                            "<th>Acción</th>" & _
                            "<th>Responsable</th>" & _
                            "<th>Fecha Inicio</th>" & _
                            "<th>F. Fin Prevista</th>" & _
                            "<th>F. Fin Real</th>" & _
                            "</tr></thead><tbody>" & vbCrLf
                            
                    For Each m_IDPA In m_Plan.colAcciones
                        Set m_Accion = m_Plan.colAcciones(m_IDPA)
                        If Tipo = Mitigacion Then
                            idAccion = Trim$(Nz(m_Accion.CodAccion, ""))
                            If idAccion = "" Then idAccion = Trim$(Nz(m_Accion.IDAccionMitigacion, ""))
                        Else
                            idAccion = Trim$(Nz(m_Accion.CodAccion, ""))
                            If idAccion = "" Then idAccion = Trim$(Nz(m_Accion.IDAccionContingencia, ""))
                        End If
                        If idAccion <> "" Then
                            labelAccion = idAccion & " - " & HTMLSafe(m_Accion.Accion)
                        Else
                            labelAccion = HTMLSafe(m_Accion.Accion)
                        End If
                        s = s & "        <tr class='ficha-action-row'>" & _
                                "<td>" & labelAccion & "</td>" & _
                                "<td>" & HTMLSafe(m_Accion.ResponsableAccion) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaInicio) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaFinPrevista) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaFinReal) & "</td>" & _
                                "</tr>" & vbCrLf
                    Next
                    s = s & "    </tbody></table>" & vbCrLf
                    s = s & "    </div>" & vbCrLf
                End If
            End If
            s = s & "</div>" & vbCrLf
        Next
    End If
    ConstruirSeccionPlanes = s
    Exit Function
    
errores:
    p_Error = "Error ConstruirSeccionPlanes: " & err.description
End Function

Private Function ConstruirSeccionMaterializaciones(p_Riesgo As Riesgo) As String
    Dim s As String
    Dim col As Scripting.Dictionary
    Dim clavesOrdenadas As Variant
    Dim i As Long
    Dim m_Mat As RiesgoMaterializacion
    Dim FechaInicio As String
    Dim fechaFin As String
    
    On Error Resume Next
    Set col = p_Riesgo.ColMaterializaciones
    On Error GoTo 0
    
    s = "<h2>Materializaciones</h2>"
    
    If col Is Nothing Then
        s = s & "<div class='card'><p>No hay materializaciones registradas para este riesgo.</p></div>"
        ConstruirSeccionMaterializaciones = s
        Exit Function
    End If
    
    If col.Count = 0 Then
        s = s & "<div class='card'><p>No hay materializaciones registradas para este riesgo.</p></div>"
        ConstruirSeccionMaterializaciones = s
        Exit Function
    End If
    
    clavesOrdenadas = OrdenarClavesMaterializacionesPorFecha(col)
    
    s = s & "<div class='card'>"
    s = s & "<table class='action-table'>"
    s = s & "<thead><tr><th>Fecha materialización</th><th>Fecha desmaterialización</th><th>Es materialización</th></tr></thead>"
    s = s & "<tbody>"
    
    FechaInicio = ""
    fechaFin = ""
    
    For i = LBound(clavesOrdenadas) To UBound(clavesOrdenadas)
        Set m_Mat = col(clavesOrdenadas(i))
        If EsTextoSi(m_Mat.EsMaterializacion) Then
            If FechaInicio <> "" Then
                s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>-</td><td>Sí</td></tr>"
            End If
            FechaInicio = m_Mat.Fecha
            fechaFin = ""
        Else
            If FechaInicio <> "" Then
                fechaFin = m_Mat.Fecha
                s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>" & FormatoFecha(fechaFin) & "</td><td>Sí</td></tr>"
                FechaInicio = ""
                fechaFin = ""
            End If
        End If
        Set m_Mat = Nothing
    Next
    
    If FechaInicio <> "" Then
        s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>-</td><td>Sí</td></tr>"
    End If
    
    s = s & "</tbody></table>"
    s = s & "</div>"
    
    ConstruirSeccionMaterializaciones = s
End Function

Private Function OrdenarClavesMaterializacionesPorFecha(ByVal p_Col As Scripting.Dictionary) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant
    
    keys = p_Col.keys
    
    If UBound(keys) <= LBound(keys) Then
        OrdenarClavesMaterializacionesPorFecha = keys
        Exit Function
    End If
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If FechaSerial(p_Col(keys(i)).Fecha) > FechaSerial(p_Col(keys(j)).Fecha) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    
    OrdenarClavesMaterializacionesPorFecha = keys
End Function

Private Function FechaSerial(ByVal p_Valor As Variant) As Double
    If IsDate(p_Valor) Then
        FechaSerial = CDbl(CDate(p_Valor))
    Else
        FechaSerial = 0
    End If
End Function

Private Function EsTextoSi(ByVal p_Valor As Variant) As Boolean
    Dim t As String
    t = Nz(p_Valor, "")
    If Len(t) = 0 Then
        EsTextoSi = False
        Exit Function
    End If
    EsTextoSi = (UCase$(Left$(t, 1)) = "S")
End Function

Private Function ConstruirSeccionPublicabilidad(p_Riesgo As Riesgo) As String
    Dim s As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Publicable As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_Key As Variant
    Dim m_Check As Scripting.Dictionary
    Dim m_Estado As EnumPublicabilidadCheckEstado
    Dim m_Label As String
    Dim m_Clase As String
    Dim m_Detalle As String
    Dim m_VeredictoTexto As String
    Dim m_VeredictoClase As String

    On Error GoTo errores

    If ConstruirDatosPublicabilidadRiesgo(p_Riesgo, m_Datos, , m_Error) = EnumSiNo.No Then
        s = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(m_Error) & "</p></div>"
        ConstruirSeccionPublicabilidad = s
        Exit Function
    End If

    m_Publicable = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        s = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(m_Error) & "</p></div>"
        ConstruirSeccionPublicabilidad = s
        Exit Function
    End If

    Select Case m_Veredicto
        Case EnumPublicabilidadVeredicto.Publicable
            m_VeredictoTexto = "Publicable"
            m_VeredictoClase = "verdict-publicable"
        Case EnumPublicabilidadVeredicto.NoPublicable
            m_VeredictoTexto = "No publicable"
            m_VeredictoClase = "verdict-no-publicable"
        Case Else
            m_VeredictoTexto = "No aplica"
            m_VeredictoClase = "verdict-no-aplica"
    End Select

    s = "<div class='card verdict-card'>"
    s = s & "<div><h2 style='margin:0 0 6px 0;'>Publicabilidad</h2>" & _
            "<div style='color:var(--grey-6); font-size:13px;'>Resumen de comprobaciones</div></div>"
    s = s & "<div class='verdict-badge " & m_VeredictoClase & "'>" & m_VeredictoTexto & "</div>"
    s = s & "</div>"

    s = s & "<div class='checklist'>"
    If Not m_Checks Is Nothing Then
        For Each m_Key In m_Checks
            Set m_Check = m_Checks(m_Key)
            m_Estado = m_Check("estado")
            s = s & InformeHTML_RenderCheckPublicabilidad(m_Check, m_Estado)
        Next
    End If
    s = s & "</div>"

    ConstruirSeccionPublicabilidad = s
    Exit Function

errores:
    ConstruirSeccionPublicabilidad = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(err.description) & "</p></div>"
End Function

Private Function EstadoPublicabilidadLabel(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String
    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            EstadoPublicabilidadLabel = "Cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            EstadoPublicabilidadLabel = "No cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            EstadoPublicabilidadLabel = "No aplica"
        Case Else
            EstadoPublicabilidadLabel = "Desconocido"
    End Select
End Function

Private Function EstadoPublicabilidadClass(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String
    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            EstadoPublicabilidadClass = "state-cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            EstadoPublicabilidadClass = "state-no-cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            EstadoPublicabilidadClass = "state-no-aplica"
        Case Else
            EstadoPublicabilidadClass = "state-no-aplica"
    End Select
End Function

' ----------------------------------------------------------------------------
' Render HTML de un check individual del informe de publicabilidad.
' Punto 09 (issue #88): el motivo (por que no se cumple un check) debe
' renderizarse SIEMPRE cuando el check es "No cumple", con dos partes
' separadas visualmente:
'   (a) que se evalua       -> <div class='check-text'>   (siempre visible)
'   (b) por que no se cumple -> <div class='check-detail'>
' Comportamiento:
'   - Si hay detalle:    check-detail = detalle (justificacion literal)
'   - Si no hay detalle: check-detail = "(sin motivo registrado)"
'   - HTML safe via HTMLSafe (<, >, &, ")
'   - Saltos de linea preservados en el detalle
'   - Para checks "Cumple" / "NoAplica" se conserva el comportamiento legacy
'     (check-detail solo si existe y no esta vacio).
' ----------------------------------------------------------------------------
Private Function InformeHTML_RenderCheckPublicabilidad( _
                                                    ByVal p_Check As Object, _
                                                    ByVal p_Estado As EnumPublicabilidadCheckEstado _
                                                    ) As String
    Dim m_Label As String
    Dim m_Clase As String
    Dim m_Detalle As String
    Dim s As String

    m_Label = EstadoPublicabilidadLabel(p_Estado)
    m_Clase = EstadoPublicabilidadClass(p_Estado)

    s = "<div class='check-item'>"
    s = s & "<div class='check-left'><div class='check-text'>" & HTMLSafe(CStr(p_Check("texto"))) & "</div>"

    ' Punto 09: motivo en dos partes - que se evalua (check-text arriba)
    ' + por que no se cumple (check-detail aqui). Para checks "No cumple"
    ' SIEMPRE se muestra el check-detail: con el detalle si existe, o
    ' "(sin motivo registrado)" si no.
    If p_Estado = EnumPublicabilidadCheckEstado.NoCumple Then
        m_Detalle = ""
        If p_Check.Exists("detalle") Then
            m_Detalle = Trim$(CStr(p_Check("detalle")))
        End If
        If Len(m_Detalle) > 0 Then
            s = s & "<div class='check-detail'>" & HTMLSafe(m_Detalle) & "</div>"
        Else
            s = s & "<div class='check-detail'>(sin motivo registrado)</div>"
        End If
    Else
        ' Comportamiento legacy para checks "Cumple" / "NoAplica": check-detail
        ' solo si existe la clave y no esta vacia.
        If p_Check.Exists("detalle") Then
            m_Detalle = CStr(p_Check("detalle"))
            If m_Detalle <> "" Then
                s = s & "<div class='check-detail'>" & HTMLSafe(m_Detalle) & "</div>"
            End If
        End If
    End If

    s = s & "</div>"
    s = s & "<div class='check-state " & m_Clase & "'>" & m_Label & "</div>"
    s = s & "</div>"

    InformeHTML_RenderCheckPublicabilidad = s
End Function
Public Function GetEstilosCSS_Corporativos() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & ":root { --tele-blue: #0066FF; --tele-blue-70: #0356C9; --tele-white: #F7F7FF; --pure-white: #FFFFFF; --grey-9: #031A34; --grey-6: #58617A; --grey-2: #D1D5E4; --space-2: 8px; --space-3: 12px; --space-4: 16px; --space-5: 24px; --space-6: 32px; --radius-card: 8px; --radius-input: 8px; --radius-button: 32px; }" & vbCrLf
    css = css & "body { font-family: 'Telefonica Sans', system-ui, -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; background-color: var(--tele-white); color: var(--grey-6); margin: 0; padding: 40px 4%; overflow-wrap: break-word; word-wrap: break-word; }" & vbCrLf
    css = css & ".report-section { margin-bottom: 56px; }" & vbCrLf
    css = css & ".meta-row { margin-bottom: var(--space-5); }" & vbCrLf
    css = css & "header { background-color: var(--tele-blue); color: white; padding: 25px 4% 25px 0; border-bottom: 4px solid var(--tele-blue-70); }" & vbCrLf
    css = css & ".container { width: 100%; margin: 0 auto; padding: 0; }" & vbCrLf
    css = css & ".header-container { display: flex; justify-content: space-between; align-items: center; }" & vbCrLf
    css = css & ".hdr-badge { display:inline-block; margin-top:8px; padding:7px 12px; border-radius:999px; font-weight:700; font-size:12px; letter-spacing:0.4px; text-transform:uppercase; border:1px solid rgba(255,255,255,0.45); }" & vbCrLf
    css = css & ".hdr-badge.publicable { background: rgba(232,245,233,0.18); }" & vbCrLf
    css = css & ".hdr-badge.no-publicable { background: rgba(255,235,238,0.18); }" & vbCrLf
    css = css & ".hdr-badge.no-aplica { background: rgba(236,239,241,0.18); }" & vbCrLf
    css = css & ".logo-container { display: flex; align-items: center; gap: 20px; }" & vbCrLf
    css = css & ".logo-wrapper { width: 140px; height: 60px; display: flex; align-items: center; justify-content: flex-start; }" & vbCrLf
    css = css & ".logo-wrapper svg { width: 100%; height: 100%; object-fit: contain; }" & vbCrLf
    css = css & ".header-text h1 { font-size: 18px; margin: 0; font-weight: 700; }" & vbCrLf
    css = css & ".header-text p { margin: 0; opacity: 0.85; font-size: 14px; }" & vbCrLf
    css = css & ".main-container { margin-top: 0; padding: 0; box-sizing: border-box; }" & vbCrLf
    css = css & ".tab-container { display: flex; gap: 10px; }" & vbCrLf
    css = css & ".tab-btn { padding: 14px 28px; border: 1px solid var(--grey-2); background: #E5F0FF; cursor: pointer; border-radius: var(--radius-button) var(--radius-button) 0 0; font-weight: 600; color: var(--grey-6); transition: 0.3s; }" & vbCrLf
    css = css & ".tab-btn.active { background: var(--pure-white); border-bottom: 2px solid var(--pure-white); color: var(--tele-blue); }" & vbCrLf
    css = css & ".tab-content { display: none; background: var(--pure-white); padding: 32px; border: 1px solid var(--grey-2); border-radius: 0 var(--radius-card) var(--radius-card) var(--radius-card); box-shadow: 0 4px 15px rgba(3,26,52,0.06); }" & vbCrLf
    css = css & ".tab-content.active { display: block; animation: fadeIn 0.4s; }" & vbCrLf
    css = css & ".subtab-container { display:flex; gap:10px; margin: 0 0 18px 0; }" & vbCrLf
    css = css & ".subtab-btn { padding: 10px 18px; border: 1px solid var(--grey-2); background: #E5F0FF; cursor: pointer; border-radius: var(--radius-button); font-weight: 600; color: var(--grey-6); transition: 0.3s; }" & vbCrLf
    css = css & ".subtab-btn.active { background: var(--pure-white); color: var(--tele-blue); border-color: var(--tele-blue); }" & vbCrLf
    css = css & ".subtab-content { display:none; }" & vbCrLf
    css = css & ".subtab-content.active { display:block; }" & vbCrLf
    css = css & ".card { background: white; padding: 24px; border-radius: var(--radius-card); border: 1px solid var(--grey-2); width: 100%; box-sizing: border-box; }" & vbCrLf
    css = css & ".card-mini { background: white; padding: 16px; border-radius: var(--radius-card); border: 1px solid var(--grey-2); display: flex; flex-direction: column; justify-content: center; min-width: 0; max-width: 50%; }" & vbCrLf
    css = css & ".verdict-card { display: flex; align-items: center; justify-content: space-between; gap: 15px; margin-bottom: 20px; }" & vbCrLf
    css = css & ".verdict-badge { padding: 8px 14px; border-radius: 999px; font-weight: 700; font-size: 12px; letter-spacing: 0.5px; text-transform: uppercase; }" & vbCrLf
    css = css & ".verdict-publicable { background: #E8F5E9; color: #1B5E20; }" & vbCrLf
    css = css & ".verdict-no-publicable { background: #FFEBEE; color: #B71C1C; }" & vbCrLf
    css = css & ".verdict-no-aplica { background: #ECEFF1; color: #455A64; }" & vbCrLf
    css = css & ".checklist { display: grid; gap: 12px; }" & vbCrLf
    css = css & ".check-item { border: 1px solid var(--grey-2); border-radius: var(--radius-card); padding: 12px 14px; display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; }" & vbCrLf
    css = css & ".check-left { flex: 1; }" & vbCrLf
    css = css & ".check-text { font-weight: 400; color: var(--grey-9); }" & vbCrLf
    css = css & ".check-detail { margin-top: 4px; font-size: 12px; color: var(--grey-6); }" & vbCrLf
    css = css & ".check-state { font-size: 11px; font-weight: bold; padding: 4px 8px; border-radius: 999px; text-transform: uppercase; }" & vbCrLf
    css = css & ".state-cumple { background: #E8F5E9; color: #1B5E20; }" & vbCrLf
    css = css & ".state-no-cumple { background: #FFEBEE; color: #B71C1C; }" & vbCrLf
    css = css & ".state-no-aplica { background: #ECEFF1; color: #455A64; }" & vbCrLf
    css = css & ".grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 25px; }" & vbCrLf
    css = css & ".grid-compact { display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; }" & vbCrLf
    css = css & "@media (max-width: 1100px) { .grid-compact { grid-template-columns: repeat(3, 1fr); } }" & vbCrLf
    css = css & "@media (max-width: 600px) { .grid-compact { grid-template-columns: repeat(2, 1fr); } }" & vbCrLf
    css = css & "h2 { color: var(--grey-9); border-bottom: 2px solid var(--tele-blue); padding-bottom: 8px; margin-bottom: 20px; }" & vbCrLf
    css = css & "h3 { font-size: 12px; margin: 0 0 5px 0; color: var(--grey-6); text-transform: uppercase; letter-spacing: 0.4px; }" & vbCrLf
    
    css = css & ".ficha-valor, .report-table td, .description-text, .action-table td, .report-table-small td { font-size: 14px !important; font-weight: 400 !important; color: var(--grey-6); }" & vbCrLf
    css = css & ".report-table th, .action-table th, .report-table-small th { font-size: 14px !important; font-weight: 700 !important; color: var(--grey-9); }" & vbCrLf
    css = css & ".ficha-subtitle, .collapsible-body h3 { font-size: 14px !important; font-weight: 700; text-transform: uppercase; margin-bottom: 8px; }" & vbCrLf
    css = css & ".plan-status-badge {" & vbCrLf
    css = css & "    font-size: 0.75em; font-weight: 700; text-transform: uppercase; padding: 2px 6px; background: #E5F0FF; border-radius: 24px; color: var(--grey-9); margin-left: 8px; display: inline-block; vertical-align: middle; }" & vbCrLf
    ' ----------------------------------------------------------------

    css = css & ".kpi-value { font-size: 16px; font-weight: 600; color: var(--grey-9); }" & vbCrLf
    css = css & ".kpi-value-mini { font-size: 14px; font-weight: 600; color: var(--grey-9); line-height: 1.2; }" & vbCrLf
    css = css & ".action-table { width: 100%; border-collapse: collapse; margin-top: 20px; }" & vbCrLf
    css = css & ".action-table th { text-align: left; padding: 12px; border-bottom: 2px solid var(--grey-2); background: #fcfcfc; }" & vbCrLf
    css = css & ".action-table td { padding: 12px; border-bottom: 1px solid #f6f6f6; font-size: 14px; font-weight: 400; }" & vbCrLf
    css = css & ".ficha-action { margin-left: 0; display: block; }" & vbCrLf
    css = css & ".cc-plan-action { margin-left: 16px; display: block; padding-top: 4px; }" & vbCrLf
    css = css & ".plans-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; }" & vbCrLf
    css = css & ".legend-container { display: flex; align-items: center; gap: 10px; margin: 10px 0 20px; }" & vbCrLf
    css = css & ".legend-item { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--grey-6); }" & vbCrLf
    css = css & ".dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }" & vbCrLf
    css = css & ".dot.dot-new { background: #0b7a28; }" & vbCrLf
    css = css & ".dot.dot-unchanged { background: #58617A; }" & vbCrLf
    css = css & ".dot.dot-deleted { background: #b42318; }" & vbCrLf
    css = css & "footer { text-align: center; padding: 40px; color: var(--grey-6); font-size: 12px; }" & vbCrLf
    css = css & "@keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_Corporativos = css
End Function

Private Function GetScriptsJS() As String
    Dim js As String
    js = "<script>" & vbCrLf
    js = js & "function openTab(evt, tabName) {" & vbCrLf
    js = js & "  var i, tabcontent, tablinks;" & vbCrLf
    js = js & "  tabcontent = document.getElementsByClassName('tab-content');" & vbCrLf
    js = js & "  for (i = 0; i < tabcontent.length; i++) { tabcontent[i].classList.remove('active'); }" & vbCrLf
    js = js & "  tablinks = document.getElementsByClassName('tab-btn');" & vbCrLf
    js = js & "  for (i = 0; i < tablinks.length; i++) { tablinks[i].classList.remove('active'); }" & vbCrLf
    js = js & "  document.getElementById(tabName).classList.add('active');" & vbCrLf
    js = js & "  evt.currentTarget.classList.add('active');" & vbCrLf
    js = js & "}" & vbCrLf
    js = js & "function openSubTab(evt, tabName, parentId) {" & vbCrLf
    js = js & "  var i, parent, tabcontent, tablinks;" & vbCrLf
    js = js & "  parent = document.getElementById(parentId);" & vbCrLf
    js = js & "  if (!parent) { return; }" & vbCrLf
    js = js & "  tabcontent = parent.getElementsByClassName('subtab-content');" & vbCrLf
    js = js & "  for (i = 0; i < tabcontent.length; i++) { tabcontent[i].classList.remove('active'); }" & vbCrLf
    js = js & "  tablinks = parent.getElementsByClassName('subtab-btn');" & vbCrLf
    js = js & "  for (i = 0; i < tablinks.length; i++) { tablinks[i].classList.remove('active'); }" & vbCrLf
    js = js & "  document.getElementById(tabName).classList.add('active');" & vbCrLf
    js = js & "  evt.currentTarget.classList.add('active');" & vbCrLf
    js = js & "}" & vbCrLf
    js = js & "(function(){" & vbCrLf
    js = js & "  var input = document.getElementById('sectionSearchInput');" & vbCrLf
    js = js & "  var results = document.getElementById('sectionSearchResults');" & vbCrLf
    js = js & "  if (!input || !results) return;" & vbCrLf
    js = js & "  var activeIdx = -1; var items = [];" & vbCrLf
    js = js & "  function buildIndex() {" & vbCrLf
    js = js & "    var entries = []; var headings = document.querySelectorAll('h2, h3');" & vbCrLf
    js = js & "    for (var i = 0; i < headings.length; i++) { var h = headings[i]; var txt = h.textContent.trim(); if (!txt || txt.length < 2) continue; entries.push({ el: h, text: txt, lc: txt.toLowerCase() }); }" & vbCrLf
    js = js & "    return entries;" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & "  var allSections = buildIndex();" & vbCrLf
    js = js & "  function highlight(text, query) {" & vbCrLf
    js = js & "    var idx = text.toLowerCase().indexOf(query.toLowerCase()); if (idx < 0) return text;" & vbCrLf
    js = js & "    return text.substring(0, idx) + ""<span class='section-search-highlight'>"" + text.substring(idx, idx + query.length) + ""</span>"" + text.substring(idx + query.length);" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & "  function showResults(query) {" & vbCrLf
    js = js & "    activeIdx = -1; results.innerHTML = '';" & vbCrLf
    js = js & "    if (!query || query.length < 1) { results.classList.remove('visible'); return; }" & vbCrLf
    js = js & "    var q = query.toLowerCase(); items = [];" & vbCrLf
    js = js & "    for (var i = 0; i < allSections.length; i++) { if (allSections[i].lc.indexOf(q) >= 0) items.push(allSections[i]); }" & vbCrLf
    js = js & "    if (items.length === 0) { results.innerHTML = ""<div class='sr-empty'>Sin resultados</div>""; results.classList.add('visible'); return; }" & vbCrLf
    js = js & "    for (var j = 0; j < items.length; j++) { var div = document.createElement('div'); div.className = 'sr-item'; div.innerHTML = highlight(items[j].text, query); div.setAttribute('data-idx', j); div.addEventListener('click', (function(idx){ return function(){ goTo(idx); }; })(j)); results.appendChild(div); }" & vbCrLf
    js = js & "    results.classList.add('visible');" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & "  function goTo(idx) {" & vbCrLf
    js = js & "    var entry = items[idx]; if (!entry) return; var el = entry.el;" & vbCrLf
    js = js & "    var tabContent = el.closest('.tab-content'); if (tabContent) {" & vbCrLf
    js = js & "      var allTabs = document.getElementsByClassName('tab-content'); for (var t = 0; t < allTabs.length; t++) allTabs[t].classList.remove('active'); tabContent.classList.add('active');" & vbCrLf
    js = js & "      var tabBtns = document.getElementsByClassName('tab-btn'); for (var b = 0; b < tabBtns.length; b++) { tabBtns[b].classList.remove('active'); if (tabBtns[b].getAttribute('onclick') && tabBtns[b].getAttribute('onclick').indexOf(tabContent.id) >= 0) { tabBtns[b].classList.add('active'); } }" & vbCrLf
    js = js & "    }" & vbCrLf
    js = js & "    var subTabContent = el.closest('.subtab-content'); if (subTabContent) {" & vbCrLf
    js = js & "      var parentDiv = subTabContent.parentElement; if (parentDiv) {" & vbCrLf
    js = js & "        var allSubs = parentDiv.getElementsByClassName('subtab-content'); for (var s = 0; s < allSubs.length; s++) allSubs[s].classList.remove('active'); subTabContent.classList.add('active');" & vbCrLf
    js = js & "        var subBtns = parentDiv.getElementsByClassName('subtab-btn'); for (var sb = 0; sb < subBtns.length; sb++) { subBtns[sb].classList.remove('active'); if (subBtns[sb].getAttribute('onclick') && subBtns[sb].getAttribute('onclick').indexOf(subTabContent.id) >= 0) { subBtns[sb].classList.add('active'); } }" & vbCrLf
    js = js & "      }" & vbCrLf
    js = js & "    }" & vbCrLf
    js = js & "    var node = el.parentElement; while (node) { if (node.tagName === 'DETAILS' && !node.open) node.open = true; node = node.parentElement; }" & vbCrLf
    js = js & "    setTimeout(function(){ el.scrollIntoView({ behavior: 'smooth', block: 'start' }); el.style.transition = 'background 0.3s'; el.style.background = '#FFE082'; setTimeout(function(){ el.style.background = ''; }, 1500); }, 100);" & vbCrLf
    js = js & "    results.classList.remove('visible'); input.value = '';" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & "  function updateActive() { var divs = results.getElementsByClassName('sr-item'); for (var i = 0; i < divs.length; i++) divs[i].classList.remove('sr-active'); if (activeIdx >= 0 && activeIdx < divs.length) { divs[activeIdx].classList.add('sr-active'); divs[activeIdx].scrollIntoView({ block: 'nearest' }); } }" & vbCrLf
    js = js & "  input.addEventListener('input', function(){ showResults(this.value); });" & vbCrLf
    js = js & "  input.addEventListener('keydown', function(e){ if (e.key === 'ArrowDown') { e.preventDefault(); if (activeIdx < items.length - 1) activeIdx++; updateActive(); } else if (e.key === 'ArrowUp') { e.preventDefault(); if (activeIdx > 0) activeIdx--; updateActive(); } else if (e.key === 'Enter' && activeIdx >= 0) { e.preventDefault(); goTo(activeIdx); } else if (e.key === 'Escape') { results.classList.remove('visible'); } });" & vbCrLf
    js = js & "  document.addEventListener('click', function(e){ if (!e.target.closest('.section-search-box')) results.classList.remove('visible'); });" & vbCrLf
    js = js & "})();" & vbCrLf
    js = js & "</script>"
    GetScriptsJS = js
End Function

Private Function ConstruirSeccionDatosProyecto(p_Riesgo As Riesgo) As String
    Dim s As String

    s = "<h2>Datos de Proyecto</h2>"
    s = s & "<div class='subtab-container'>" & vbCrLf
    s = s & "  <button class='subtab-btn active' onclick=""openSubTab(event, 'dp_generales', 'datosProyecto')"">Datos generales</button>" & vbCrLf
    s = s & "  <button class='subtab-btn' onclick=""openSubTab(event, 'dp_responsables', 'datosProyecto')"">Responsables</button>" & vbCrLf
    s = s & "</div>" & vbCrLf

    s = s & "<div id='dp_generales' class='subtab-content active'>" & vbCrLf
    s = s & ConstruirSeccionDatosProyectoGenerales(p_Riesgo)
    s = s & "</div>" & vbCrLf

    s = s & "<div id='dp_responsables' class='subtab-content'>" & vbCrLf
    s = s & ConstruirSeccionDatosProyectoResponsables(p_Riesgo)
    s = s & "</div>" & vbCrLf

    ConstruirSeccionDatosProyecto = s
End Function

Private Function ConstruirSeccionDatosProyectoGenerales(p_Riesgo As Riesgo) As String
    Dim s As String
    Dim sNemotecnico As String
    Dim sTitulo As String
    Dim sCliente As String
    Dim sNombreProyecto As String
    Dim sJuridica As String
    Dim sEnUTE As String
    Dim sFechaFirmaContrato As String
    Dim sFechaPrevistaCierre As String
    Dim sFechaCierre As String
    Dim sFechaProximaPublicacion As String
    Dim sTextoFechaProximaPublicacion As String
    Dim sCodigoDocumento As String

    sNemotecnico = "-": sTitulo = "-": sCliente = "-": sNombreProyecto = "-": sJuridica = "-"
    sEnUTE = "-": sFechaFirmaContrato = "": sFechaPrevistaCierre = "": sFechaCierre = "": sFechaProximaPublicacion = "": sTextoFechaProximaPublicacion = "-": sCodigoDocumento = "-"

    On Error Resume Next
    sNemotecnico = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Nemotecnico, "-")
    sTitulo = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Titulo, "-")
    sCliente = Nz(p_Riesgo.Edicion.Proyecto.Cliente, "-")
    sNombreProyecto = Nz(p_Riesgo.Edicion.Proyecto.NombreProyecto, "-")
    sJuridica = Nz(p_Riesgo.Edicion.Proyecto.Juridica, "-")
    sEnUTE = Nz(p_Riesgo.Edicion.Proyecto.EnUTECalculado, "-")
    sFechaFirmaContrato = Nz(p_Riesgo.Edicion.Proyecto.Expediente.FechaFirmaContrato, "")
    If sFechaFirmaContrato = "" Then sFechaFirmaContrato = Nz(p_Riesgo.Edicion.Proyecto.FechaFirmaContrato, "")
    sFechaPrevistaCierre = Nz(p_Riesgo.Edicion.Proyecto.FechaPrevistaCierre, "")
    sFechaCierre = Nz(p_Riesgo.Edicion.Proyecto.FechaCierre, "")
    sFechaProximaPublicacion = Nz(p_Riesgo.Edicion.Proyecto.FechaMaxProximaPublicacion, "")
    sCodigoDocumento = Nz(p_Riesgo.Edicion.Proyecto.CodigoDocumento, "-")
    On Error GoTo 0

    If IsDate(sFechaCierre) Then
        sTextoFechaProximaPublicacion = "No aplica"
    Else
        sTextoFechaProximaPublicacion = FormatoFecha(sFechaProximaPublicacion)
    End If

    s = "<div class='grid'>" & vbCrLf
    s = s & "  <div class='card'><h4>Identificación</h4>" & _
            "<p><strong>Nemotécnico:</strong> " & HTMLSafe(sNemotecnico) & "</p>" & _
            "<p><strong>Título:</strong> " & HTMLSafe(sTitulo) & "</p>" & _
            "<p><strong>Código Documento:</strong> " & HTMLSafe(sCodigoDocumento) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Datos generales</h4>" & _
            "<p><strong>Cliente:</strong> " & HTMLSafe(sCliente) & "</p>" & _
            "<p><strong>Nombre Proyecto:</strong> " & HTMLSafe(sNombreProyecto) & "</p>" & _
            "<p><strong>Jurídica:</strong> " & HTMLSafe(sJuridica) & "</p>" & _
            "<p><strong>En UTE:</strong> " & HTMLSafe(sEnUTE) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Fechas</h4>" & _
            "<p><strong>Firma contrato:</strong> " & FormatoFecha(sFechaFirmaContrato) & "</p>" & _
            "<p><strong>Prevista cierre:</strong> " & FormatoFecha(sFechaPrevistaCierre) & "</p>" & _
            "<p><strong>Cierre:</strong> " & FormatoFecha(sFechaCierre) & "</p>" & _
            "<p><strong>Próxima publicación:</strong> " & HTMLSafe(sTextoFechaProximaPublicacion) & "</p></div>" & vbCrLf

    s = s & "</div>" & vbCrLf
    ConstruirSeccionDatosProyectoGenerales = s
End Function

Private Function ConstruirSeccionDatosProyectoResponsables(p_Riesgo As Riesgo) As String
    Dim s As String
    Dim sCorreoResponsableCalidad As String
    Dim sTecnicoCalidad As String
    Dim sRACs As String
    Dim sCorreosRACs As String
    Dim sAutorizados As String

    sCorreoResponsableCalidad = "-": sTecnicoCalidad = "-"
    sRACs = "-": sCorreosRACs = "-": sAutorizados = "-"

    On Error Resume Next
    sTecnicoCalidad = Nz(p_Riesgo.Edicion.Proyecto.UsuarioCalidad.Nombre, "-")
    sCorreoResponsableCalidad = Nz(p_Riesgo.Edicion.Proyecto.CorreoResponsableCalidad, "-")
    sRACs = Nz(p_Riesgo.Edicion.Proyecto.Expediente.CadenaRACs, "-")
    sCorreosRACs = Nz(p_Riesgo.Edicion.Proyecto.Expediente.CadenaCorreoRACs, "-")
    sAutorizados = Nz(p_Riesgo.Edicion.Proyecto.CadenaNombreAutorizados, "-")
    On Error GoTo 0

    If sAutorizados <> "-" Then
        sAutorizados = Replace(sAutorizados, "|", vbCrLf)
    End If

    s = "<div class='grid'>" & vbCrLf
    s = s & "  <div class='card'><h4>Calidad</h4>" & _
            "<p><strong>Técnico Calidad:</strong> " & HTMLSafe(sTecnicoCalidad) & "</p>" & _
            "<p><strong>Correo Responsable:</strong> " & HTMLSafe(sCorreoResponsableCalidad) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>RAC</h4>" & _
            "<p><strong>RACs:</strong><br>" & HTMLSafeLargo(Replace(sRACs, "|", vbCrLf)) & "</p>" & _
            "<p><strong>Correos:</strong><br>" & HTMLSafeLargo(Replace(sCorreosRACs, ";", vbCrLf)) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Autorizados</h4><p>" & HTMLSafeLargo(sAutorizados) & "</p></div>" & vbCrLf
    s = s & "</div>" & vbCrLf

    ConstruirSeccionDatosProyectoResponsables = s
End Function

Private Function GuardarInformeHTML_UTF8(ByVal p_Riesgo As Riesgo, ByVal p_HTML As String, ByRef p_Error As String) As String
    Dim m_Ruta As String, m_Stream As Object
    On Error GoTo errores
    m_Ruta = Environ("TEMP") & "\HPS_Report_" & SanitizarNombreArchivo(p_Riesgo.CodigoRiesgo) & ".html"
    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2: m_Stream.Charset = "utf-8": m_Stream.Open
    m_Stream.WriteText p_HTML: m_Stream.SaveToFile m_Ruta, 2: m_Stream.Close
    GuardarInformeHTML_UTF8 = m_Ruta
    Exit Function
errores:
    p_Error = "Error UTF-8: " & err.description
End Function

Private Function HTMLSafeLargo(ByVal p_Texto As String) As String
    Dim m_T As String: m_T = HTMLSafe(p_Texto)
    m_T = Replace(m_T, vbCrLf, "<br>"): m_T = Replace(m_T, vbLf, "<br>")
    HTMLSafeLargo = m_T
End Function

Private Function FormatoFecha(ByVal p_Valor As Variant) As String
    If IsDate(p_Valor) Then FormatoFecha = Format$(CDate(p_Valor), "dd/mm/yyyy") Else FormatoFecha = "-"
End Function

Private Function SanitizarNombreArchivo(ByVal p_Texto As String) As String
    Dim m_T As String: m_T = p_Texto
    m_T = Replace(m_T, "/", "_"): m_T = Replace(m_T, "\", "_"): m_T = Replace(m_T, ":", "_")
    SanitizarNombreArchivo = m_T
End Function

' Función exclusiva para la portada centrada (con header integrado para impresión)
Private Function ConstruirSeccionPortadaHTML(ByVal p_Titulo As String, ByVal p_Expediente As String) As String
    Dim s As String
    s = "<section class='print-only-cover'>" & vbCrLf
    s = s & "  <div class='cover-titles'>" & vbCrLf
    s = s & "    <h1 style='font-size: 22pt; font-weight: 700; margin-bottom: 30px;'>INFORME DE GESTIÓN DE RIESGOS</h1>" & vbCrLf
    s = s & "    <div style='font-size: 16pt; font-weight: 600; color: #333;'>" & HTMLSafe(UCase$(p_Titulo)) & "</div>" & vbCrLf
    s = s & "    <div style='font-size: 14pt; opacity: 0.6; margin-top: 15px;'>" & HTMLSafe(p_Expediente) & "</div>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf
    ConstruirSeccionPortadaHTML = s
End Function

' Función de Datos Generales (Ya no incluye la portada)
Private Function ConstruirSeccionDatosGeneralesHTML( _
                                                    ByVal p_Nombre As String, ByVal p_Expediente As String, _
                                                    ByVal p_Jefe As String, _
                                                    ByVal p_Cliente As String, ByVal p_Codigo As String, _
                                                    ByVal p_Edicion As String, ByVal p_Fecha As String, _
                                                    ByVal p_EsPublicacion As Boolean, _
                                                    Optional ByVal p_EsPDF As Boolean = False) As String
    Dim s As String
    
    If p_EsPDF Then
        s = "<section class='report-section'>" & vbCrLf
        s = s & "  <h2>Datos generales</h2>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    Else
        s = "<details open class='report-section seccion-vertical collapsible-section'>" & vbCrLf
        s = s & "  <summary><h2>Datos generales</h2></summary>" & vbCrLf
        s = s & "  <div class='collapsible-body'>" & vbCrLf
        s = s & "  <div class='card'>" & vbCrLf
    End If

    s = s & "    <div class='report-cover-kv' style='margin-top:0;'>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Proyecto</div><div class='kv-v'>" & HTMLSafe(UCase$(p_Expediente)) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Nombre proyecto</div><div class='kv-v'>" & HTMLSafe(UCase$(p_Nombre)) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Jefe del proyecto</div><div class='kv-v'>" & HTMLSafe(p_Jefe) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Cliente</div><div class='kv-v'>" & HTMLSafe(p_Cliente) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Código documento</div><div class='kv-v'>" & HTMLSafe(p_Codigo) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Edición</div><div class='kv-v'>" & HTMLSafe(p_Edicion) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>" & IIf(p_EsPublicacion, "Fecha publicación", "Estado informe") & "</div><div class='kv-v'>" & IIf(p_EsPublicacion, HTMLSafe(p_Fecha), "Borrador") & "</div></div>" & vbCrLf
    s = s & "    </div>" & vbCrLf

    If p_EsPDF Then
        s = s & "  </div>" & vbCrLf
        s = s & "</section>" & vbCrLf
    Else
        s = s & "  </div>" & vbCrLf
        s = s & "  </div>" & vbCrLf
        s = s & "</details>" & vbCrLf
    End If

    ConstruirSeccionDatosGeneralesHTML = s
End Function



Public Function ConvertirHTMLaPDFConEdge(ByVal p_RutaHTML As String, ByVal p_RutaPDF As String, Optional ByRef p_Error As String) As Boolean
    Dim objShell As Object
    Dim comando As String
    Dim resultado As Long
    Dim tempDir As String

    On Error GoTo errores
    p_Error = ""

    ' 1. Creamos una carpeta temporal para forzar un perfil limpio de Edge
    ' Esto evita que herede configuraciones de tu navegador (como mostrar la ruta/fecha)
    tempDir = Environ("TEMP") & "\EdgePerfilLimpio"

    Set objShell = CreateObject("WScript.Shell")

    ' 2. Comando estricto. El orden es importante.
    comando = "msedge --headless=new --disable-gpu --user-data-dir=""" & tempDir & """ --no-pdf-header-footer --print-to-pdf=""" & p_RutaPDF & """ """ & p_RutaHTML & """"

    resultado = objShell.Run(comando, 0, True)

    If resultado = 0 Then
        ConvertirHTMLaPDFConEdge = True
    Else
        p_Error = "El proceso falló con código: " & CStr(resultado)
        ConvertirHTMLaPDFConEdge = False
    End If

    Set objShell = Nothing
    Exit Function

errores:
    p_Error = "Error en conversión: " & err.description
    ConvertirHTMLaPDFConEdge = False
    If Not objShell Is Nothing Then Set objShell = Nothing
End Function

Private Function GetEstilosCSS_WebOnly() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & "/* Restauración de triángulos interactivos (Solo Web) */" & vbCrLf
    css = css & "details.collapsible-section > summary {" & vbCrLf
    css = css & "    list-style: none;" & vbCrLf
    css = css & "    cursor: pointer;" & vbCrLf
    css = css & "    display: flex;" & vbCrLf
    css = css & "    align-items: center;" & vbCrLf
    css = css & "    gap: 10px;" & vbCrLf
    css = css & "    user-select: none;" & vbCrLf
    css = css & "}" & vbCrLf
    css = css & "details.collapsible-section > summary::-webkit-details-marker {" & vbCrLf
    css = css & "    display: none;" & vbCrLf
    css = css & "}" & vbCrLf
    css = css & "details.collapsible-section > summary::before {" & vbCrLf
    css = css & "    content: '\25B6'; /* Triángulo ? */" & vbCrLf
    css = css & "    font-size: 12px;" & vbCrLf
    css = css & "    color: var(--tele-blue);" & vbCrLf
    css = css & "    transition: transform 0.2s ease;" & vbCrLf
    css = css & "    flex-shrink: 0;" & vbCrLf
    css = css & "}" & vbCrLf
    css = css & "details[open].collapsible-section > summary::before {" & vbCrLf
    css = css & "    transform: rotate(90deg); /* Se convierte en ? */" & vbCrLf
    css = css & "}" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_WebOnly = css
End Function



