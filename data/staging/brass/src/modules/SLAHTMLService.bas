Attribute VB_Name = "SLAHTMLService"
Option Compare Database
Option Explicit

'===========================================================
' SLAHTMLService
' Servicio para generación de informes SLA en formato HTML
' Diseñado para reutilizar el dataset filtrado de SLAReportService
'===========================================================

'===========================================================

'-----------------------------------------------------------
' GetCSS_MISTICA — Tokens de marca Mística
'-----------------------------------------------------------
Private Function GetCSS_MISTICA() As String
    Dim s As String
    s = ":root {"
    s = s & " --color-highlight: #0066FF;"
    s = s & " --color-text-primary: #031A34;"
    s = s & " --color-bg-primary: #F2F4FF;"
    s = s & " --color-surface: #FFFFFF;"
    s = s & " --color-text-secondary: #58617A;"
    s = s & " --color-border: #D1D5E4;"
    s = s & " --color-success: #5CB85C;"
    s = s & " --color-error: #E66C64;"
    s = s & " --color-warning: #F4C495;"
    s = s & " --radius-button: 32px;"
    s = s & " --radius-container: 4px;"
    s = s & " --radius-chip: 24px;"
    s = s & " --shadow: 0 4px 12px rgba(0,102,255,0.1);"
    s = s & "}"
    s = s & "@media (prefers-color-scheme: dark) { :root { --color-highlight: #227AFF; --color-bg-primary: #0A1628; --color-surface: #162033; --color-text-primary: #F2F4FF; --color-text-secondary: #9BA5B7; } }"
    s = s & "body { font-family: 'Telefonica Sans','Helvetica Neue',Helvetica,Arial,sans-serif; background-color: var(--color-bg-primary); color: var(--color-text-secondary); margin: 0; padding: 0; }"
    s = s & "header { background-color: var(--color-highlight); color: var(--color-surface); padding: 40px 20px; border-bottom-left-radius: 50% 20px; border-bottom-right-radius: 50% 20px; }"
    s = s & "header h1 { color: var(--color-surface); font-weight: 100; font-size: 2.2rem; margin-bottom: 8px; }"
    s = s & "header p { opacity: 0.9; font-size: 1rem; }"
    s = s & "header .subtitle { opacity: 0.85; font-size: 1.15rem; font-weight: 300; margin-top: 0; }"
    s = s & ".container { max-width: 1100px; margin: 30px auto; padding: 0 20px; }"
    s = s & ".card { background-color: var(--color-surface); border-radius: var(--radius-container); padding: 25px; margin-bottom: 25px; box-shadow: var(--shadow); }"
    s = s & ".card.info { border-left: 6px solid var(--color-highlight); }"
    s = s & ".summary-card h2 { color: var(--color-text-primary); font-weight: 600; margin-top: 0; font-size: 1.1rem; }"
    s = s & ".donut-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 20px; margin-top: 20px; }"
    s = s & ".donut-item { display: flex; flex-direction: column; align-items: center; }"
    s = s & ".donut-legend { margin-top: 10px; text-align: center; }"
    s = s & ".donut-legend-item { font-size: .85rem; color: var(--color-text-secondary); margin: 4px 0; }"
    s = s & ".dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 6px; }"
    s = s & ".tabs { display: flex; gap: 10px; margin-bottom: 20px; border-bottom: 2px solid var(--color-border); padding-bottom: 10px; }"
    s = s & ".tab-btn { background: none; border: none; padding: 10px 20px; font-size: 1rem; color: var(--color-text-secondary); cursor: pointer; border-radius: var(--radius-button); transition: all .3s ease; font-family: inherit; }"
    s = s & ".tab-btn:hover { color: var(--color-highlight); background-color: rgba(0,102,255,.05); }"
    s = s & ".tab-btn.active { background-color: var(--color-highlight); color: var(--color-surface); font-weight: 500; }"
    s = s & ".tab-content { display: none; animation: fadeIn .5s; }"
    s = s & ".tab-content.active { display: block; }"
    s = s & "@keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }"
    s = s & ".expand-col { width: 28px; text-align: center; }"
    s = s & ".expand-icon { display:inline-block; width:16px; text-align:center; color: var(--color-highlight); font-weight:700; }"
    s = s & ".event-detail-row { display: none; }"
    s = s & "table { width: 100%; border-collapse: collapse; background: var(--color-surface); border-radius: var(--radius-container); overflow: hidden; }"
    s = s & "th { background-color: var(--color-highlight); color: var(--color-surface); padding: 10px 12px; text-align: left; font-size: 12px; font-weight: 700; }"
    s = s & "td { padding: 9px 12px; border-bottom: 1px solid var(--color-border); font-size: 12px; }"
    s = s & "tr:nth-child(even) td { background-color: var(--color-bg-primary); }"
    s = s & "tr:hover td { background-color: rgba(0,102,255,.05); cursor: pointer; }"
    s = s & "tr.selected td { background-color: rgba(0,102,255,.1); }"
    s = s & ".cumple-cell { color: var(--color-success); font-weight: 600; }"
    s = s & ".no-cumple-cell { color: var(--color-error); font-weight: 600; }"
    s = s & ".badge { display: inline-block; padding: 3px 10px; border-radius: var(--radius-chip); font-size: 11px; font-weight: 600; }"
    s = s & ".badge-cumple { background-color: #E8F5E9; color: var(--color-success); }"
    s = s & ".badge-no-cumple { background-color: #FFEBEE; color: var(--color-error); }"
    s = s & ".badge-na { background-color: var(--color-bg-primary); color: var(--color-text-secondary); }"
    s = s & ".event-detail { background-color: #FFF9F0; border-left: 4px solid var(--color-warning); border-radius: var(--radius-container); padding: 20px; margin-top: 15px; }"
    s = s & ".detail-section { margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px solid var(--color-border); }"
    s = s & ".detail-section:last-child { border-bottom: none; margin-bottom: 0; }"
    s = s & ".detail-section h4 { margin: 0 0 8px 0; color: var(--color-text-primary); font-size: 13px; }"
    s = s & ".detail-row { display: flex; gap: 20px; margin: 4px 0; font-size: 12px; }"
    s = s & ".detail-label { color: var(--color-text-secondary); min-width: 160px; }"
    s = s & ".detail-value { color: var(--color-text-primary); font-weight: 500; }"
    s = s & "footer { text-align: center; padding: 20px; color: var(--color-text-secondary); font-size: .85rem; margin-top: 30px; }"
    s = s & "input[type='radio'] { display: none; }"
    s = s & "#tab-events:checked ~ .tabs label[for='tab-events'], #tab-rules:checked ~ .tabs label[for='tab-rules'] { background-color: var(--color-highlight); color: var(--color-surface); }"
    s = s & "#tab-events:checked ~ .content-events { display: block; }"
    s = s & "#tab-rules:checked ~ .content-rules { display: block; }"
    GetCSS_MISTICA = s
End Function

'-----------------------------------------------------------
' f_DonutSVG
' Genera SVG inline de donut chart con stroke-dasharray
' Input: percent (0-100), colorCumple (verde), colorNoCumple (rojo), cumple (bool)
' Output: String SVG completo con <svg>, circle fondo, circle arco, texto centro
' Fórmula: circumference = 2 * PI * r, dasharray = (percent/100) * circ & " " & circ
' Rotación -90deg para empezar desde arriba
'-----------------------------------------------------------
Private Function f_DonutSVG(ByVal p_percent As Double, _
                            ByVal p_colorCumple As String, _
                            ByVal p_colorNoCumple As String, _
                            ByVal p_cumple As Boolean) As String
    Dim dblRadius As Double
    Dim dblStrokeWidth As Double
    Dim dblCirc As Double
    Dim dblDashCumple As Double
    Dim dblDashNoCumple As Double
    Dim strCirc As String
    Dim strDashCumple As String
    Dim strDashNoCumple As String

    dblRadius = 40
    dblStrokeWidth = 8
    dblCirc = 2 * 3.14159265358979 * dblRadius
    dblDashCumple = (p_percent / 100) * dblCirc
    dblDashNoCumple = dblCirc - dblDashCumple

    ' SVG requiere punto decimal, no coma regional
    strCirc = Replace(CStr(dblCirc), ",", ".")
    strDashCumple = Replace(CStr(dblDashCumple), ",", ".")
    strDashNoCumple = Replace(CStr(dblDashNoCumple), ",", ".")

    ' Donut robusto: fondo rojo completo (no cumple) + arco verde (cumple).
    ' Así SIEMPRE se ve rojo cuando p_percent < 100.
    f_DonutSVG = "<svg width=""100"" height=""100"" viewBox=""0 0 100 100"" xmlns=""http://www.w3.org/2000/svg"">" & vbCrLf & _
        "<circle cx=""50"" cy=""50"" r=""" & dblRadius & """ fill=""none"" stroke=""" & p_colorNoCumple & """ stroke-width=""" & dblStrokeWidth & """/>" & vbCrLf & _
        "<circle cx=""50"" cy=""50"" r=""" & dblRadius & """ fill=""none"" stroke=""" & p_colorCumple & """ stroke-width=""" & dblStrokeWidth & """ " & _
        "stroke-dasharray=""" & strDashCumple & " " & strCirc & """ " & _
        "stroke-linecap=""round"" transform=""rotate(-90 50 50)""/>" & vbCrLf & _
        "<text x=""50"" y=""55"" text-anchor=""middle"" font-family=""Helvetica,Arial,sans-serif"" " & _
        "font-size=""14"" font-weight=""600"" fill=""" & p_colorCumple & """>" & Round(p_percent, 0) & "%</text>" & vbCrLf & _
        "</svg>"
End Function

'-----------------------------------------------------------
' f_EscapeHTML
' Escapa caracteres especiales HTML para evitar inyección
' y mala renderización
' Input: text As String
' Output: String con & ? &amp;, < ? &lt;, > ? &gt;, " ? &quot;, ' ? &#39;
' Preserva tildes, ñ y caracteres españoles SIN escapado adicional
'-----------------------------------------------------------
Private Function f_EscapeHTML(ByVal p_Texto As String) As String
    Dim strTexto As String

    strTexto = Nz(p_Texto, "")
    strTexto = Replace(strTexto, "&", "&amp;")
    strTexto = Replace(strTexto, "<", "&lt;")
    strTexto = Replace(strTexto, ">", "&gt;")
    strTexto = Replace(strTexto, """", "&quot;")
    strTexto = Replace(strTexto, "'", "&#39;")

    f_EscapeHTML = strTexto
End Function

'-----------------------------------------------------------
' f_DonutLegendItem
' Genera HTML del item de leyenda para cada donut
' Input: label, percent, isCumple
' Output: String HTML como <span class="legend-item">...
' Color: success-green (#5CB85C) si cumple, alert-red (#E66C64) si no
'-----------------------------------------------------------
Private Function f_DonutLegendItem(ByVal p_label As String, _
                                    ByVal p_percent As Double, _
                                    ByVal p_cumple As Boolean) As String
    Dim strColor As String

    If p_cumple Then
        strColor = "#5CB85C"
    Else
        strColor = "#E66C64"
    End If

    f_DonutLegendItem = "<div class=""donut-legend-item"">" & _
        "<span class=""dot"" style=""background:" & strColor & """></span>" & _
        f_EscapeHTML(p_label) & ": " & Round(p_percent, 0) & "%" & _
        "</div>"
End Function

'-----------------------------------------------------------
' GenerarHeaderHTML
' Genera el header del HTML con gradiente tele-blue y título
' Input: fechaInicio, fechaFin
' Output: String HTML del <header>
'-----------------------------------------------------------
Private Function GenerarHeaderHTML(ByVal p_FechaInicio As String, _
                                   ByVal p_FechaFin As String) As String
    Dim strHTML As String

    strHTML = "<header>" & vbCrLf
    strHTML = strHTML & "  <div class=""container"">" & vbCrLf
    strHTML = strHTML & "    <h1>BRASS</h1>" & vbCrLf
    strHTML = strHTML & "    <p class=""subtitle"">Informe SLA — Eventos Franqueados</p>" & vbCrLf
    strHTML = strHTML & "    <p>Rango: " & f_EscapeHTML(p_FechaInicio) & " — " & f_EscapeHTML(p_FechaFin) & "</p>" & vbCrLf
    strHTML = strHTML & "  </div>" & vbCrLf
    strHTML = strHTML & "</header>" & vbCrLf

    GenerarHeaderHTML = strHTML
End Function

'===========================================================
' F2 — Tab Reglas SLA
'-----------------------------------------------------------
' GenerarTabReglasSLA
' Genera contenido del tab "Reglas SLA" con 4 cards.info
' Output: String HTML con 4 cards: TRES, TRCM, TRSS, SLA-4
'-----------------------------------------------------------
Private Function GenerarTabReglasSLA() As String
    Dim strHTML As String

    strHTML = vbCrLf & "<!-- F2: Reglas SLA -->" & vbCrLf

    '---- Card TRES ----
    strHTML = strHTML & "<div class=""card info"">" & vbCrLf
    strHTML = strHTML & "  <h3>TRES — Tiempo de Respuesta</h3>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Regla:</strong> FechaInicioContactoCliente &gt;= FechaRecepcionNotificacion</p>" & vbCrLf
    strHTML = strHTML & "  <p>El tiempo de respuesta mide el intervalo entre la recepci&oacute;n de la notificaci&oacute;n " & vbCrLf
    strHTML = strHTML & "y el primer contacto con el cliente.</p>" & vbCrLf
    strHTML = strHTML & "  <table>" & vbCrLf
    strHTML = strHTML & "    <tr><th>Criticidad</th><th>Objetivo</th></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>1 (Cr&iacute;tica)</td><td>&lt;= 1 hora laborable</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>3 (Alta)</td><td>&lt;= 3 horas laborables</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>5 (Media)</td><td>&lt;= 8 horas laborables</td></tr>" & vbCrLf
    strHTML = strHTML & "  </table>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Card TRCM ----
    strHTML = strHTML & "<div class=""card info"">" & vbCrLf
    strHTML = strHTML & "  <h3>TRCM — Tiempo de Reparaci&oacute;n / Adquisici&oacute;n de Componentes</h3>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Regla:</strong> Solo aplica si IncidenciaAveriaOReparacion = True</p>" & vbCrLf
    strHTML = strHTML & "  <p>El tiempo de reparaci&oacute;n/adquisici&oacute;n mide el intervalo entre el inicio " & vbCrLf
    strHTML = strHTML & "y el fin de la adquisici&oacute;n de componentes.</p>" & vbCrLf
    strHTML = strHTML & "  <table>" & vbCrLf
    strHTML = strHTML & "    <tr><th>Tipo</th><th>Urgente</th><th>No urgente</th></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>Contratista</td><td>7 d&iacute;as</td><td>30 d&iacute;as</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>Fabricante</td><td>60 d&iacute;as</td><td>150 d&iacute;as</td></tr>" & vbCrLf
    strHTML = strHTML & "  </table>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Card TRSS ----
    strHTML = strHTML & "<div class=""card info"">" & vbCrLf
    strHTML = strHTML & "  <h3>TRSS — Tiempo de Restablecimiento del Servicio</h3>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Regla:</strong> Solo aplica si EventoConServicioAfectado = True</p>" & vbCrLf
    strHTML = strHTML & "  <p>El tiempo de restablecimiento mide el intervalo entre la recepci&oacute;n " & vbCrLf
    strHTML = strHTML & "de la notificaci&oacute;n y el restablecimiento del servicio.</p>" & vbCrLf
    strHTML = strHTML & "  <table>" & vbCrLf
    strHTML = strHTML & "    <tr><th>Criticidad</th><th>Plazo m&aacute;ximo</th></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>1 (Cr&iacute;tica)</td><td>1 d&iacute;a laborable desde la solicitud</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>3 (Alta)</td><td>5 d&iacute;as laborables desde la solicitud</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>5 (Media)</td><td>15 d&iacute;as laborables desde la solicitud</td></tr>" & vbCrLf
    strHTML = strHTML & "  </table>" & vbCrLf
    strHTML = strHTML & "  <p><em>Solo aplica si EventoConServicioAfectado = True. Si se requieren repuestos no disponibles, el plazo se incrementa seg&uacute;n TRCM.</em></p>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Card SLA-4 ----
    strHTML = strHTML & "<div class=""card info"">" & vbCrLf
    strHTML = strHTML & "  <h3>SLA-4 — Reparaci&oacute;n In-situ</h3>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Regla:</strong> Mide el porcentaje de reparaciones in-situ sobre el total de equipos reparables</p>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Numerador:</strong> TipoRepInsitu = True AND (TipoRepNoSMT = True OR TipoRepValvulas = True)</p>" & vbCrLf
    strHTML = strHTML & "  <p><strong>Denominador:</strong> TipoRepNoSMT = True OR TipoRepValvulas = True</p>" & vbCrLf
    strHTML = strHTML & "  <p><strong>F&oacute;rmula:</strong> Porcentaje = (Numerador / Denominador) x 100</p>" & vbCrLf
    strHTML = strHTML & "  <table>" & vbCrLf
    strHTML = strHTML & "    <tr><th>Indicador</th><th>Descripci&oacute;n</th></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>TipoRepInsitu</td><td>Reparaci&oacute;n realizada in-situ</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>TipoRepNoSMT</td><td>Reparaci&oacute;n SMT no realizada</td></tr>" & vbCrLf
    strHTML = strHTML & "    <tr><td>TipoRepValvulas</td><td>Reparaci&oacute;n de v&aacute;lvulas</td></tr>" & vbCrLf
    strHTML = strHTML & "  </table>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    GenerarTabReglasSLA = strHTML
End Function

'===========================================================
' F1 — Donut Charts Card
'-----------------------------------------------------------
' GenerarDonutChartsCard
' Genera card de resumen con 4 donut charts SVG
' Input: dTRES, nTRES, dTRCM, nTRCM, dTRSS, nTRSS, dGlobal, nGlobal
'   (cumple y total de cada SLA)
' Output: String HTML de la card con 4 donuts en grid
'-----------------------------------------------------------
Private Function GenerarDonutChartsCard( _
    ByVal p_dTRES As Long, ByVal p_nTRES As Long, _
    ByVal p_dTRCM As Long, ByVal p_nTRCM As Long, _
    ByVal p_dTRSS As Long, ByVal p_nTRSS As Long, _
    ByVal p_dGlobal As Long, ByVal p_nGlobal As Long) As String

    Dim strHTML As String
    Dim dblPctTRES As Double, dblPctTRCM As Double
    Dim dblPctTRSS As Double, dblPctGlobal As Double

    If p_nTRES > 0 Then dblPctTRES = Round((p_dTRES / p_nTRES) * 100, 1) Else dblPctTRES = 0
    If p_nTRCM > 0 Then dblPctTRCM = Round((p_dTRCM / p_nTRCM) * 100, 1) Else dblPctTRCM = 0
    If p_nTRSS > 0 Then dblPctTRSS = Round((p_dTRSS / p_nTRSS) * 100, 1) Else dblPctTRSS = 0
    If p_nGlobal > 0 Then dblPctGlobal = Round((p_dGlobal / p_nGlobal) * 100, 1) Else dblPctGlobal = 0

    strHTML = vbCrLf & "<div class=""card summary-card"">" & vbCrLf
    strHTML = strHTML & "  <h2>Resumen de Cumplimiento</h2>" & vbCrLf
    strHTML = strHTML & "  <div class=""donut-grid"">" & vbCrLf

    ' Donut 1: Global
    strHTML = strHTML & "    <div class=""donut-item"">" & vbCrLf
    strHTML = strHTML & "      " & f_DonutSVG(dblPctGlobal, "#5CB85C", "#E66C64", True) & vbCrLf
    strHTML = strHTML & "      <div class=""donut-legend"">" & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("Global", dblPctGlobal, True) & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("No cumple", 100 - dblPctGlobal, False) & vbCrLf
    strHTML = strHTML & "        <div class=""donut-legend-item"" style=""font-size:0.75rem;color:var(--grey-6)"">" & p_dGlobal & " / " & p_nGlobal & "</div>" & vbCrLf
    strHTML = strHTML & "      </div>" & vbCrLf
    strHTML = strHTML & "    </div>" & vbCrLf

    ' Donut 2: TRES
    strHTML = strHTML & "    <div class=""donut-item"">" & vbCrLf
    strHTML = strHTML & "      " & f_DonutSVG(dblPctTRES, "#5CB85C", "#E66C64", True) & vbCrLf
    strHTML = strHTML & "      <div class=""donut-legend"">" & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("TRES", dblPctTRES, True) & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("No cumple", 100 - dblPctTRES, False) & vbCrLf
    strHTML = strHTML & "        <div class=""donut-legend-item"" style=""font-size:0.75rem;color:var(--grey-6)"">" & p_dTRES & " / " & p_nTRES & "</div>" & vbCrLf
    strHTML = strHTML & "      </div>" & vbCrLf
    strHTML = strHTML & "    </div>" & vbCrLf

    ' Donut 3: TRCM
    strHTML = strHTML & "    <div class=""donut-item"">" & vbCrLf
    strHTML = strHTML & "      " & f_DonutSVG(dblPctTRCM, "#5CB85C", "#E66C64", True) & vbCrLf
    strHTML = strHTML & "      <div class=""donut-legend"">" & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("TRCM", dblPctTRCM, True) & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("No cumple", 100 - dblPctTRCM, False) & vbCrLf
    strHTML = strHTML & "        <div class=""donut-legend-item"" style=""font-size:0.75rem;color:var(--grey-6)"">" & p_dTRCM & " / " & p_nTRCM & "</div>" & vbCrLf
    strHTML = strHTML & "      </div>" & vbCrLf
    strHTML = strHTML & "    </div>" & vbCrLf

    ' Donut 4: TRSS
    strHTML = strHTML & "    <div class=""donut-item"">" & vbCrLf
    strHTML = strHTML & "      " & f_DonutSVG(dblPctTRSS, "#5CB85C", "#E66C64", True) & vbCrLf
    strHTML = strHTML & "      <div class=""donut-legend"">" & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("TRSS", dblPctTRSS, True) & vbCrLf
    strHTML = strHTML & "        " & f_DonutLegendItem("No cumple", 100 - dblPctTRSS, False) & vbCrLf
    strHTML = strHTML & "        <div class=""donut-legend-item"" style=""font-size:0.75rem;color:var(--grey-6)"">" & p_dTRSS & " / " & p_nTRSS & "</div>" & vbCrLf
    strHTML = strHTML & "      </div>" & vbCrLf
    strHTML = strHTML & "    </div>" & vbCrLf

    strHTML = strHTML & "  </div>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    GenerarDonutChartsCard = strHTML
End Function

'===========================================================
' F3 — Tabla de Eventos con Detalle Expandible
'-----------------------------------------------------------
' GenerarTablaEventosHTML
' Genera tabla HTML con eventos y KPIs, con onclick para expandir detalle
' Input: p_dicEventos (Dictionary IDEVENTO -> dictCon detalle + campos base)
' Output: String HTML <table>
'-----------------------------------------------------------
Private Function GenerarTablaEventosHTML(ByVal p_dicEventos As Object, _
                                        Optional ByRef p_Error As String) As String

    Dim strHTML As String
    Dim strFila As String
    Dim strFechaAlta As String
    Dim blnCumpleTRES As Boolean, blnCumpleTRCM As Boolean
    Dim blnCumpleTRSS As Boolean, blnCumpleGlobal As Boolean
    Dim strIDEvento As String
    Dim dictEvento As Object
    Dim varKey As Variant
    Dim nKeys As Long
    Dim varItemKey As Variant

    p_Error = ""
    On Error GoTo errores

    Debug.Print "=== GenerarTablaEventosHTML ==="
    Debug.Print "p_dicEventos Is Nothing: " & (p_dicEventos Is Nothing)
    If Not p_dicEventos Is Nothing Then
        nKeys = p_dicEventos.count
        Debug.Print "dicEventos keys (count): " & nKeys
        For Each varKey In p_dicEventos.Keys
            Debug.Print "  key: " & varKey
        Next varKey
    End If

    strHTML = "<table class=""event-table"">" & vbCrLf
    strHTML = strHTML & "<thead>" & vbCrLf
    strHTML = strHTML & "<tr>" & vbCrLf
    strHTML = strHTML & "<th class=""expand-col""></th>" & vbCrLf
    strHTML = strHTML & "<th>IDEVENTO</th>" & vbCrLf
    strHTML = strHTML & "<th>Fecha Alta</th>" & vbCrLf
    strHTML = strHTML & "<th>BUI</th>" & vbCrLf
    strHTML = strHTML & "<th>Subsistema</th>" & vbCrLf
    strHTML = strHTML & "<th>Criticidad</th>" & vbCrLf
    strHTML = strHTML & "<th>Cumple TRES</th>" & vbCrLf
    strHTML = strHTML & "<th>Cumple TRCM</th>" & vbCrLf
    strHTML = strHTML & "<th>Cumple TRSS</th>" & vbCrLf
    strHTML = strHTML & "<th>Cumple Global</th>" & vbCrLf
    strHTML = strHTML & "</tr>" & vbCrLf
    strHTML = strHTML & "</thead>" & vbCrLf
    strHTML = strHTML & "<tbody>" & vbCrLf

    If Not p_dicEventos Is Nothing Then
        For Each varKey In p_dicEventos.Keys
            Set dictEvento = p_dicEventos(varKey)
            Debug.Print "Procesando key: " & varKey & " | dictEvento Is Nothing: " & (dictEvento Is Nothing)
            If Not dictEvento Is Nothing Then
                Debug.Print "  Count dictEvento: " & dictEvento.count
                For Each varItemKey In dictEvento.Keys
                    Debug.Print "    " & CStr(varItemKey) & " = " & CStr(dictEvento(varItemKey))
                Next varItemKey

                strIDEvento = dictEvento("IDEVENTO")
                strFechaAlta = dictEvento("FECHAALTA")
                blnCumpleTRES = dictEvento("TRES_Cumple")
                blnCumpleTRCM = dictEvento("TRCM_Cumple")
                blnCumpleTRSS = dictEvento("TRSS_Cumple")
                blnCumpleGlobal = dictEvento("Global_CumpleGlobal")

                strFila = "  <tr class=""event-row"" data-id=""" & f_EscapeHTML(strIDEvento) & """ " & _
                    "onclick=""toggleDetail(this)"">" & vbCrLf
                strFila = strFila & "<td class=""expand-col""><span class=""expand-icon"">&#9654;</span></td>" & vbCrLf
                strFila = strFila & "<td>" & f_EscapeHTML(strIDEvento) & "</td>" & vbCrLf
                strFila = strFila & "<td>" & f_EscapeHTML(strFechaAlta) & "</td>" & vbCrLf
                strFila = strFila & "<td>" & f_EscapeHTML(Nz(dictEvento("BUI"), "")) & "</td>" & vbCrLf
                strFila = strFila & "<td>" & f_EscapeHTML(Nz(dictEvento("SUBSISTEMA"), "")) & "</td>" & vbCrLf
                strFila = strFila & "<td>" & f_EscapeHTML(Nz(dictEvento("CRITICIDAD"), "")) & "</td>" & vbCrLf
                strFila = strFila & "<td class=""" & IIf(blnCumpleTRES, "cumple-cell", "no-cumple-cell") & """>" & _
                    IIf(blnCumpleTRES, "Sí", "No") & "</td>" & vbCrLf
                strFila = strFila & "<td class=""" & IIf(blnCumpleTRCM, "cumple-cell", "no-cumple-cell") & """>" & _
                    IIf(blnCumpleTRCM, "Sí", "No") & "</td>" & vbCrLf
                strFila = strFila & "<td class=""" & IIf(blnCumpleTRSS, "cumple-cell", "no-cumple-cell") & """>" & _
                    IIf(blnCumpleTRSS, "Sí", "No") & "</td>" & vbCrLf
                strFila = strFila & "<td class=""" & IIf(blnCumpleGlobal, "cumple-cell", "no-cumple-cell") & """>" & _
                    IIf(blnCumpleGlobal, "Sí", "No") & "</td>" & vbCrLf
                strFila = strFila & "</tr>" & vbCrLf

                strHTML = strHTML & strFila
                strHTML = strHTML & GenerarEventDetailPanel(dictEvento, p_Error)
                If p_Error <> "" Then
                    Exit For
                End If
            End If
        Next varKey
    End If

    strHTML = strHTML & "</tbody>" & vbCrLf
    strHTML = strHTML & "</table>" & vbCrLf

    GenerarTablaEventosHTML = strHTML
    Exit Function

errores:
    p_Error = "Error en GenerarTablaEventosHTML: " & Err.Number & " - " & Err.Description
    Debug.Print p_Error
    GenerarTablaEventosHTML = "<p>Error al generar tabla de eventos.</p>"
End Function

'-----------------------------------------------------------
' GenerarEventDetailPanel
' Genera HTML del panel de detalle expandible para F3
' Input: detailDict (estructura de ObtenerDetalleCumplimientoEvento)
' Output: String HTML del panel con 4 bloques: TRES, TRCM, TRSS, SLA-4
'-----------------------------------------------------------
Private Function GenerarEventDetailPanel(ByVal p_dictDetalle As Object, _
                                        Optional ByRef p_Error As String) As String
    Dim strHTML As String
    Dim strBadgeTRES As String, strBadgeTRCM As String, strBadgeTRSS As String
    Dim strBorderColor As String
    Dim strHoras As String, strObjetivoHoras As String
    Dim strDias As String, strObjetivoDias As String

    p_Error = ""
    On Error GoTo errores

    Dim blnGlobal As Boolean
    blnGlobal = False
    If Not p_dictDetalle Is Nothing Then
        If p_dictDetalle.Exists("Global_CumpleGlobal") Then
            blnGlobal = p_dictDetalle("Global_CumpleGlobal")
        End If
    End If
    strBorderColor = IIf(blnGlobal, "#5CB85C", "#E66C64")

    strHTML = "<tr class=""event-detail-row"" style=""display:none;""><td colspan=""10"">" & vbCrLf
    strHTML = strHTML & "<div class=""event-detail"" style=""border-left-color: " & strBorderColor & """>" & vbCrLf

    strHTML = strHTML & "<div style=""text-align:right; margin-bottom:10px;"">" & _
        "<button onclick=""closeDetail(this)"" style=""border:none;background:none;cursor:pointer;font-size:16px;"">Cerrar</button>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Bloque TRES ----
    strHTML = strHTML & "<div class=""detail-section"">" & vbCrLf
    strHTML = strHTML & "<h4>TRES — Tiempo de Respuesta</h4>" & vbCrLf

    If p_dictDetalle.Exists("TRES_FechaRecepcion") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Fecha Recepción:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRES_FechaRecepcion"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRES_FechaContacto") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Fecha Contacto:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRES_FechaContacto"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    strHoras = "N/A"
    strObjetivoHoras = "N/A"
    If p_dictDetalle.Exists("TRES_Horas") Then
        If p_dictDetalle("TRES_Horas") <> "" Then
            strHoras = p_dictDetalle("TRES_Horas") & " horas"
        End If
    End If
    If p_dictDetalle.Exists("TRES_ObjetivoHoras") Then
        If p_dictDetalle("TRES_ObjetivoHoras") <> "" Then
            strObjetivoHoras = "<= " & p_dictDetalle("TRES_ObjetivoHoras") & " horas"
        End If
    End If

    strHTML = strHTML & "<div class=""detail-row"">" & _
        "<span class=""detail-label"">Horas transcurridas:</span>" & _
        "<span class=""detail-value"">" & f_EscapeHTML(strHoras) & "</span>" & _
        "</div>" & vbCrLf
    strHTML = strHTML & "<div class=""detail-row"">" & _
        "<span class=""detail-label"">Objetivo:</span>" & _
        "<span class=""detail-value"">" & f_EscapeHTML(strObjetivoHoras) & "</span>" & _
        "</div>" & vbCrLf

    If p_dictDetalle.Exists("TRES_Cumple") Then
        Dim blnTRES As Boolean
        blnTRES = p_dictDetalle("TRES_Cumple")
        strBadgeTRES = IIf(blnTRES, "<span class=""badge badge-cumple"">Cumple</span>", "<span class=""badge badge-no-cumple"">No cumple</span>")
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Resultado:</span>" & _
            "<span class=""detail-value"">" & strBadgeTRES & "</span>" & _
            "</div>" & vbCrLf
    End If
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Bloque TRCM ----
    strHTML = strHTML & "<div class=""detail-section"">" & vbCrLf
    strHTML = strHTML & "<h4>TRCM — Tiempo de Reparación / Adquisición</h4>" & vbCrLf

    If p_dictDetalle.Exists("TRCM_Incidencia") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Incidencia:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRCM_Incidencia"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRCM_TipoReparacion") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Tipo Reparación:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRCM_TipoReparacion"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRCM_Urgente") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Urgente:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRCM_Urgente"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRCM_FechaInicioAdq") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Fecha Inicio Adq:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRCM_FechaInicioAdq"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRCM_FechaFinAdq") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Fecha Fin Adq:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRCM_FechaFinAdq"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    strDias = "N/A"
    strObjetivoDias = "N/A"
    If p_dictDetalle.Exists("TRCM_Dias") Then
        If p_dictDetalle("TRCM_Dias") <> "" And p_dictDetalle("TRCM_Dias") <> "N/A" Then
            strDias = p_dictDetalle("TRCM_Dias") & " días"
        End If
    End If
    If p_dictDetalle.Exists("TRCM_ObjetivoDias") Then
        If p_dictDetalle("TRCM_ObjetivoDias") <> "" And p_dictDetalle("TRCM_ObjetivoDias") <> "N/A" Then
            strObjetivoDias = "<= " & p_dictDetalle("TRCM_ObjetivoDias") & " días"
        End If
    End If

    strHTML = strHTML & "<div class=""detail-row"">" & _
        "<span class=""detail-label"">Días transcurridos:</span>" & _
        "<span class=""detail-value"">" & f_EscapeHTML(strDias) & "</span>" & _
        "</div>" & vbCrLf
    strHTML = strHTML & "<div class=""detail-row"">" & _
        "<span class=""detail-label"">Objetivo:</span>" & _
        "<span class=""detail-value"">" & f_EscapeHTML(strObjetivoDias) & "</span>" & _
        "</div>" & vbCrLf

    If p_dictDetalle.Exists("TRCM_Cumple") Then
        Dim blnTRCM As Boolean
        Dim strTRCMEstado As String
        blnTRCM = p_dictDetalle("TRCM_Cumple")
        
        ' Usar estado expandido si existe; fallback a lógica Boolean legacy
        If p_dictDetalle.Exists("TRCM_Estado") Then
            strTRCMEstado = p_dictDetalle("TRCM_Estado")
            ' Renderizar Inconsistente con badge-na (warning), no como No Cumple (red)
            Select Case strTRCMEstado
                Case "Cumple"
                    strBadgeTRCM = "<span class=""badge badge-cumple"">Cumple</span>"
                Case "Inconsistente"
                    strBadgeTRCM = "<span class=""badge badge-na"" style=""background-color:#F4C495;color:#8B4513;"">Inconsistente</span>"
                Case "No cumple"
                    strBadgeTRCM = "<span class=""badge badge-no-cumple"">No cumple</span>"
                Case "N/A"
                    strBadgeTRCM = "<span class=""badge badge-na"">N/A</span>"
                Case Else
                    strBadgeTRCM = IIf(blnTRCM, "<span class=""badge badge-cumple"">Cumple</span>", "<span class=""badge badge-no-cumple"">No cumple</span>")
            End Select
        Else
            strBadgeTRCM = IIf(blnTRCM, "<span class=""badge badge-cumple"">Cumple</span>", "<span class=""badge badge-no-cumple"">No cumple</span>")
        End If
        
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Resultado:</span>" & _
            "<span class=""detail-value"">" & strBadgeTRCM & "</span>" & _
            "</div>" & vbCrLf
    End If
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Bloque TRSS ----
    strHTML = strHTML & "<div class=""detail-section"">" & vbCrLf
    strHTML = strHTML & "<h4>TRSS — Tiempo de Restablecimiento del Servicio</h4>" & vbCrLf

    If p_dictDetalle.Exists("TRSS_ServicioAfectado") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Servicio Afectado:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRSS_ServicioAfectado"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRSS_FechaRestablecimiento") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Fecha Restablecimiento:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("TRSS_FechaRestablecimiento"), "N/A")) & "</span>" & _
            "</div>" & vbCrLf
    End If
    
    ' Mostrar días transcurridos y objetivo si están disponibles
    If p_dictDetalle.Exists("TRSS_DiasTranscurridos") And p_dictDetalle.Exists("TRSS_ObjetivoDias") Then
        Dim strTRSSDias As String, strTRSSObj As String
        strTRSSDias = Nz(p_dictDetalle("TRSS_DiasTranscurridos"), "N/A")
        strTRSSObj = Nz(p_dictDetalle("TRSS_ObjetivoDias"), "N/A")
        If strTRSSDias <> "N/A" And strTRSSDias <> "" Then
            strTRSSDias = strTRSSDias & " días"
        End If
        If strTRSSObj <> "N/A" And strTRSSObj <> "" Then
            strTRSSObj = "<= " & strTRSSObj & " días"
        End If
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Días transcurridos:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(strTRSSDias) & "</span>" & _
            "</div>" & vbCrLf
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Objetivo:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(strTRSSObj) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("TRSS_Cumple") Then
        Dim blnTRSS As Boolean
        Dim strTRSSEstado As String
        blnTRSS = p_dictDetalle("TRSS_Cumple")
        
        ' Usar estado expandido si existe
        If p_dictDetalle.Exists("TRSS_Estado") Then
            strTRSSEstado = p_dictDetalle("TRSS_Estado")
            Select Case strTRSSEstado
                Case "Cumple"
                    strBadgeTRSS = "<span class=""badge badge-cumple"">Cumple</span>"
                Case "Inconsistente"
                    strBadgeTRSS = "<span class=""badge badge-na"" style=""background-color:#F4C495;color:#8B4513;"">Inconsistente</span>"
                Case "No cumple"
                    strBadgeTRSS = "<span class=""badge badge-no-cumple"">No cumple</span>"
                Case "N/A"
                    strBadgeTRSS = "<span class=""badge badge-na"">N/A</span>"
                Case Else
                    strBadgeTRSS = IIf(blnTRSS, "<span class=""badge badge-cumple"">Cumple</span>", "<span class=""badge badge-no-cumple"">No cumple</span>")
            End Select
        Else
            strBadgeTRSS = IIf(blnTRSS, "<span class=""badge badge-cumple"">Cumple</span>", "<span class=""badge badge-no-cumple"">No cumple</span>")
        End If
        
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Resultado:</span>" & _
            "<span class=""detail-value"">" & strBadgeTRSS & "</span>" & _
            "</div>" & vbCrLf
    End If
    strHTML = strHTML & "</div>" & vbCrLf

    '---- Bloque SLA-4 ----
    strHTML = strHTML & "<div class=""detail-section"">" & vbCrLf
    strHTML = strHTML & "<h4>SLA-4 — Reparación In-situ</h4>" & vbCrLf

    If p_dictDetalle.Exists("SLA4_Numerador") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Numerador:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("SLA4_Numerador"), "0")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("SLA4_Denominador") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Denominador:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("SLA4_Denominador"), "0")) & "</span>" & _
            "</div>" & vbCrLf
    End If

    If p_dictDetalle.Exists("SLA4_Porcentaje") Then
        strHTML = strHTML & "<div class=""detail-row"">" & _
            "<span class=""detail-label"">Porcentaje:</span>" & _
            "<span class=""detail-value"">" & f_EscapeHTML(Nz(p_dictDetalle("SLA4_Porcentaje"), "0")) & "%</span>" & _
            "</div>" & vbCrLf
    End If
    strHTML = strHTML & "</div>" & vbCrLf

    strHTML = strHTML & "</div>" & vbCrLf
    strHTML = strHTML & "</td></tr>" & vbCrLf

    GenerarEventDetailPanel = strHTML
    Exit Function

errores:
    p_Error = "Error en GenerarEventDetailPanel: " & Err.Number & " - " & Err.Description
    Debug.Print p_Error
    GenerarEventDetailPanel = "<tr><td colspan=""10""><p>Error al generar detalle del evento.</p></td></tr>"
End Function

'===========================================================
' CalcularCumplimiento
' Variante corporativa con CSS Mistica
'-----------------------------------------------------------
Public Function GenerarInformeHTML_Corporate( _
    ByVal p_FechaInicio As String, _
    ByVal p_FechaFin As String, _
    Optional ByRef p_Error As String) As String

    Dim strSQL As String
    Dim rcdEventos As DAO.Recordset
    Dim strHTML As String
    Dim strRutaArchivo As String
    Dim strNombreArchivo As String
    Dim objFSO As Object
    Dim objArchivo As Object

    Dim lngTotal As Long
    Dim lngCumpleTRES As Long, lngCumpleTRCM As Long
    Dim lngCumpleTRSS As Long, lngCumpleGlobal As Long

    Dim dicEventDetails As Object
    Dim strIDEvento As String

    p_Error = ""
    Set dicEventDetails = CreateObject("Scripting.Dictionary")

    On Error GoTo errores

    If Not IsDate(p_FechaInicio) Then
        p_Error = "La fecha de inicio no es válida: " & p_FechaInicio
        GenerarInformeHTML_Corporate = "#ERR|" & p_Error
        Exit Function
    End If

    If Not IsDate(p_FechaFin) Then
        p_Error = "La fecha de fin no es válida: " & p_FechaFin
        GenerarInformeHTML_Corporate = "#ERR|" & p_Error
        Exit Function
    End If

    Dim strErrorSQL As String
    strSQL = SLAReportService.ConstruirSQLEventosFranqueados(p_FechaInicio, p_FechaFin, strErrorSQL)
    If strErrorSQL <> "" Then
        p_Error = strErrorSQL
        GenerarInformeHTML_Corporate = "#ERR|" & p_Error
        Exit Function
    End If

    Debug.Print "SQL eventos: " & strSQL
    Set rcdEventos = CurrentDb().OpenRecordset(strSQL)

    On Error Resume Next
    If Not (rcdEventos.BOF And rcdEventos.EOF) Then
        rcdEventos.MoveLast
        Debug.Print "RecordCount eventos: " & rcdEventos.RecordCount
        rcdEventos.MoveFirst
    Else
        Debug.Print "RecordCount eventos: 0"
    End If
    On Error GoTo errores

    lngTotal = 0
    lngCumpleTRES = 0
    lngCumpleTRCM = 0
    lngCumpleTRSS = 0
    lngCumpleGlobal = 0

    With rcdEventos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                lngTotal = lngTotal + 1
                strIDEvento = Nz(.Fields("IDEVENTO"), "")

                Dim blnTRES As Boolean, blnTRCM As Boolean, blnTRSS As Boolean, blnGlobal As Boolean
                Call CalcularCumplimiento(.Fields, blnTRES, blnTRCM, blnTRSS, blnGlobal)

                If blnTRES Then lngCumpleTRES = lngCumpleTRES + 1
                If blnTRCM Then lngCumpleTRCM = lngCumpleTRCM + 1
                If blnTRSS Then lngCumpleTRSS = lngCumpleTRSS + 1
                If blnGlobal Then lngCumpleGlobal = lngCumpleGlobal + 1

                Dim dictDetalle As Object
                Dim lngCrit As Long
                Dim datRecep As Date, datCont As Date
                Dim datIniAdq As Date, datFinAdq As Date
                Dim strTipoRep As String
                Dim blnUrgente As Boolean
                Dim blnRepInsitu As Boolean, blnRepNoSMT As Boolean, blnRepValv As Boolean

                Set dictDetalle = CreateObject("Scripting.Dictionary")
                dictDetalle("IDEVENTO") = strIDEvento
                If IsDate(.Fields("FECHAALTAEVENTO")) Then
                    dictDetalle("FECHAALTA") = Format(.Fields("FECHAALTAEVENTO"), "dd/mm/yyyy")
                Else
                    dictDetalle("FECHAALTA") = ""
                End If
                dictDetalle("BUI") = Nz(.Fields("BUI"), "")
                dictDetalle("SUBSISTEMA") = Nz(.Fields("SUBSISTEMA"), "")
                dictDetalle("CRITICIDAD") = Nz(.Fields("CRITICIDAD"), "")
                dictDetalle("TRES_Cumple") = blnTRES
                dictDetalle("TRCM_Cumple") = blnTRCM
                dictDetalle("TRSS_Cumple") = blnTRSS
                dictDetalle("Global_CumpleGlobal") = blnGlobal

                If IsDate(.Fields("FechaRecepcionNotificacion")) Then
                    dictDetalle("TRES_FechaRecepcion") = Format(.Fields("FechaRecepcionNotificacion"), "dd/mm/yyyy hh:mm")
                Else
                    dictDetalle("TRES_FechaRecepcion") = ""
                End If
                If IsDate(.Fields("FechaInicioContactoCliente")) Then
                    dictDetalle("TRES_FechaContacto") = Format(.Fields("FechaInicioContactoCliente"), "dd/mm/yyyy hh:mm")
                Else
                    dictDetalle("TRES_FechaContacto") = ""
                End If

                If IsDate(.Fields("FechaRecepcionNotificacion")) And IsDate(.Fields("FechaInicioContactoCliente")) Then
                    datRecep = CDate(.Fields("FechaRecepcionNotificacion"))
                    datCont = CDate(.Fields("FechaInicioContactoCliente"))
                    dictDetalle("TRES_Horas") = Round((datCont - datRecep) * 24, 2)
                Else
                    dictDetalle("TRES_Horas") = ""
                End If

                lngCrit = Nz(.Fields("CRITICIDAD"), 0)
                Select Case lngCrit
                    Case 1: dictDetalle("TRES_ObjetivoHoras") = 1
                    Case 3: dictDetalle("TRES_ObjetivoHoras") = 3
                    Case 5: dictDetalle("TRES_ObjetivoHoras") = 8
                    Case Else: dictDetalle("TRES_ObjetivoHoras") = 0
                End Select

                dictDetalle("TRCM_Incidencia") = IIf(Nz(.Fields("IncidenciaAveriaOReparacion"), False), "Sí", "No")
                dictDetalle("TRCM_TipoReparacion") = Nz(.Fields("TipoReparacion"), "")
                dictDetalle("TRCM_Urgente") = IIf(Nz(.Fields("Urgente"), False), "Sí", "No")

                If IsDate(.Fields("FechaInicioTiempoAdquisicion")) Then
                    dictDetalle("TRCM_FechaInicioAdq") = Format(.Fields("FechaInicioTiempoAdquisicion"), "dd/mm/yyyy")
                Else
                    dictDetalle("TRCM_FechaInicioAdq") = ""
                End If
                If IsDate(.Fields("FechaFinTiempoAdquisicion")) Then
                    dictDetalle("TRCM_FechaFinAdq") = Format(.Fields("FechaFinTiempoAdquisicion"), "dd/mm/yyyy")
                Else
                    dictDetalle("TRCM_FechaFinAdq") = ""
                End If

                strTipoRep = Nz(.Fields("TipoReparacion"), "")
                blnUrgente = Nz(.Fields("Urgente"), False)

                Dim lngTRCMEstRaw As Long
                lngTRCMEstRaw = 0
                If Nz(.Fields("IncidenciaAveriaOReparacion"), False) = True Then
                    If IsDate(.Fields("FechaInicioTiempoAdquisicion")) And IsDate(.Fields("FechaFinTiempoAdquisicion")) Then
                        If strTipoRep <> "" Then
                            datIniAdq = CDate(.Fields("FechaInicioTiempoAdquisicion"))
                            datFinAdq = CDate(.Fields("FechaFinTiempoAdquisicion"))
                            dictDetalle("TRCM_Dias") = Round((datFinAdq - datIniAdq), 2)
                            If strTipoRep = "Contratista" Then
                                dictDetalle("TRCM_ObjetivoDias") = IIf(blnUrgente, 7, 30)
                            ElseIf strTipoRep = "Fabricante" Then
                                dictDetalle("TRCM_ObjetivoDias") = IIf(blnUrgente, 60, 150)
                            Else
                                dictDetalle("TRCM_ObjetivoDias") = "N/A"
                                lngTRCMEstRaw = 2
                            End If
                            If lngTRCMEstRaw = 0 Then
                                lngTRCMEstRaw = IIf(datFinAdq >= datIniAdq, 1, 3)
                            End If
                        Else
                            dictDetalle("TRCM_Dias") = ""
                            dictDetalle("TRCM_ObjetivoDias") = "N/A"
                            lngTRCMEstRaw = 2
                        End If
                    Else
                        dictDetalle("TRCM_Dias") = ""
                        dictDetalle("TRCM_ObjetivoDias") = "N/A"
                        lngTRCMEstRaw = 2
                    End If
                Else
                    dictDetalle("TRCM_Dias") = "N/A"
                    dictDetalle("TRCM_ObjetivoDias") = "N/A"
                End If

                Select Case lngTRCMEstRaw
                    Case 0: dictDetalle("TRCM_Estado") = "N/A"
                    Case 1: dictDetalle("TRCM_Estado") = "Cumple"
                    Case 2: dictDetalle("TRCM_Estado") = "Inconsistente"
                    Case 3: dictDetalle("TRCM_Estado") = "No cumple"
                End Select

                dictDetalle("TRSS_ServicioAfectado") = IIf(Nz(.Fields("EventoConServicioAfectado"), False), "Sí", "No")
                If IsDate(.Fields("FechaRestablecimientoServicio")) Then
                    dictDetalle("TRSS_FechaRestablecimiento") = Format(.Fields("FechaRestablecimientoServicio"), "dd/mm/yyyy")
                Else
                    dictDetalle("TRSS_FechaRestablecimiento") = ""
                End If

                Dim lngTRSSEstRaw As Long
                Dim lngTRSSCritSvc As Long
                Dim lngTRSSObjSvc As Long
                Dim lngTRSSDiasSvc As Long
                lngTRSSEstRaw = 0
                lngTRSSDiasSvc = 0
                lngTRSSObjSvc = 0

                If Nz(.Fields("EventoConServicioAfectado"), False) = True Then
                    If IsDate(.Fields("FechaRestablecimientoServicio")) And IsDate(.Fields("FechaRecepcionNotificacion")) Then
                        Dim datRecepSvc As Date, datRestabSvc As Date
                        datRecepSvc = CDate(.Fields("FechaRecepcionNotificacion"))
                        datRestabSvc = CDate(.Fields("FechaRestablecimientoServicio"))

                        If datRestabSvc >= datRecepSvc Then
                            lngTRSSCritSvc = Nz(.Fields("CRITICIDAD"), 0)
                            lngTRSSObjSvc = SLAReportService.GetTRSSObjetivo(lngTRSSCritSvc)
                            If TRSS_USA_DIAS_LABORABLES Then
                                lngTRSSDiasSvc = DateDiff("d", datRecepSvc, datRestabSvc)
                            Else
                                lngTRSSDiasSvc = DateDiff("d", datRecepSvc, datRestabSvc)
                            End If

                            If lngTRSSObjSvc > 0 Then
                                If lngTRSSDiasSvc <= lngTRSSObjSvc Then
                                    lngTRSSEstRaw = 1
                                Else
                                    lngTRSSEstRaw = 3
                                End If
                            Else
                                lngTRSSEstRaw = 2
                            End If
                        Else
                            lngTRSSEstRaw = 2
                        End If
                    Else
                        lngTRSSEstRaw = 2
                    End If
                Else
                    lngTRSSEstRaw = 0
                End If

                Select Case lngTRSSEstRaw
                    Case 0
                        dictDetalle("TRSS_Estado") = "N/A"
                        dictDetalle("TRSS_ObjetivoDias") = "N/A"
                        dictDetalle("TRSS_DiasTranscurridos") = "N/A"
                    Case 2
                        dictDetalle("TRSS_Estado") = "Inconsistente"
                        dictDetalle("TRSS_ObjetivoDias") = "N/A"
                        dictDetalle("TRSS_DiasTranscurridos") = IIf(lngTRSSDiasSvc > 0, CStr(lngTRSSDiasSvc), "N/A")
                    Case Else
                        dictDetalle("TRSS_Estado") = IIf(lngTRSSEstRaw = 1, "Cumple", "No cumple")
                        dictDetalle("TRSS_ObjetivoDias") = CStr(lngTRSSObjSvc)
                        dictDetalle("TRSS_DiasTranscurridos") = CStr(lngTRSSDiasSvc)
                End Select

                blnRepInsitu = Nz(.Fields("TipoRepInsitu"), False)
                blnRepNoSMT = Nz(.Fields("TipoRepNoSMT"), False)
                blnRepValv = Nz(.Fields("TipoRepValvulas"), False)
                dictDetalle("SLA4_Numerador") = IIf(blnRepInsitu And (blnRepNoSMT Or blnRepValv), 1, 0)
                dictDetalle("SLA4_Denominador") = IIf((blnRepNoSMT Or blnRepValv), 1, 0)
                If dictDetalle("SLA4_Denominador") > 0 Then
                    dictDetalle("SLA4_Porcentaje") = Round((dictDetalle("SLA4_Numerador") / dictDetalle("SLA4_Denominador")) * 100, 1)
                Else
                    dictDetalle("SLA4_Porcentaje") = 0
                End If

                If Len(strIDEvento & "") > 0 Then
                    If dicEventDetails.Exists(strIDEvento) Then
                        dicEventDetails.Remove strIDEvento
                    End If
                    dicEventDetails.Add strIDEvento, dictDetalle
                Else
                    Debug.Print "WARN: IDEVENTO vacío, no se agrega al diccionario"
                End If

                .MoveNext
            Loop
        End If
    End With

    Debug.Print "=== GenerarInformeHTML_Corporate ==="
    Debug.Print "dicEventDetails.Count: " & dicEventDetails.count

    Dim strDonutCharts As String
    strDonutCharts = GenerarDonutChartsCard( _
        lngCumpleTRES, lngTotal, _
        lngCumpleTRCM, lngTotal, _
        lngCumpleTRSS, lngTotal, _
        lngCumpleGlobal, lngTotal)

    Dim strTablaEventos As String
    strTablaEventos = GenerarTablaEventosHTML(dicEventDetails, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000, "GenerarInformeHTML_Corporate", p_Error
    End If

    strHTML = "<!DOCTYPE html>" & vbCrLf
    strHTML = strHTML & "<html lang=""es"">" & vbCrLf
    strHTML = strHTML & "<head>" & vbCrLf
    strHTML = strHTML & "<meta charset=""UTF-8"">" & vbCrLf
    strHTML = strHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" & vbCrLf
    strHTML = strHTML & "<title>Informe SLA — Eventos Franqueados</title>" & vbCrLf
    strHTML = strHTML & "<style type=""text/css"">" & vbCrLf
    strHTML = strHTML & GetCSS_MISTICA() & vbCrLf
    strHTML = strHTML & "</style>" & vbCrLf
    strHTML = strHTML & "</head>" & vbCrLf
    strHTML = strHTML & "<body>" & vbCrLf

    strHTML = strHTML & GenerarHeaderHTML(p_FechaInicio, p_FechaFin)

    strHTML = strHTML & "<div class=""container"">" & vbCrLf

    strHTML = strHTML & strDonutCharts

    strHTML = strHTML & "<input type=""radio"" name=""tabs"" id=""tab-events"" checked="""">" & vbCrLf
    strHTML = strHTML & "<input type=""radio"" name=""tabs"" id=""tab-rules"">" & vbCrLf
    strHTML = strHTML & "<div class=""tabs"">" & vbCrLf
    strHTML = strHTML & "  <label for=""tab-events"" class=""tab-btn active"">Eventos</label>" & vbCrLf
    strHTML = strHTML & "  <label for=""tab-rules"" class=""tab-btn"">Reglas SLA</label>" & vbCrLf
    strHTML = strHTML & "</div>" & vbCrLf

    strHTML = strHTML & "<div class=""tab-content content-events"">" & vbCrLf
    strHTML = strHTML & strTablaEventos
    strHTML = strHTML & "</div>" & vbCrLf

    strHTML = strHTML & "<div class=""tab-content content-rules"">" & vbCrLf
    strHTML = strHTML & GenerarTabReglasSLA()
    strHTML = strHTML & "</div>" & vbCrLf

    strHTML = strHTML & "<footer>Generado: " & f_EscapeHTML(Format(Now(), "dd/mm/yyyy hh:mm:ss")) & " — BRASS</footer>" & vbCrLf

    strHTML = strHTML & "<script type=""text/javascript"">" & vbCrLf
    strHTML = strHTML & "function toggleDetail(row) {" & vbCrLf
    strHTML = strHTML & "  var detailRow = row.nextElementSibling;" & vbCrLf
    strHTML = strHTML & "  var icon = row.querySelector('.expand-icon');" & vbCrLf
    strHTML = strHTML & "  if (detailRow && detailRow.classList && detailRow.classList.contains('event-detail-row')) {" & vbCrLf
    strHTML = strHTML & "    if (detailRow.style.display === 'none' || detailRow.style.display === '') {" & vbCrLf
    strHTML = strHTML & "      detailRow.style.display = 'table-row';" & vbCrLf
    strHTML = strHTML & "      row.classList.add('selected');" & vbCrLf
    strHTML = strHTML & "      if(icon) icon.innerHTML = '&#9660;';" & vbCrLf
    strHTML = strHTML & "    } else {" & vbCrLf
    strHTML = strHTML & "      detailRow.style.display = 'none';" & vbCrLf
    strHTML = strHTML & "      row.classList.remove('selected');" & vbCrLf
    strHTML = strHTML & "      if(icon) icon.innerHTML = '&#9654;';" & vbCrLf
    strHTML = strHTML & "    }" & vbCrLf
    strHTML = strHTML & "  }" & vbCrLf
    strHTML = strHTML & "}" & vbCrLf
    strHTML = strHTML & "function closeDetail(btn) {" & vbCrLf
    strHTML = strHTML & "  var detailRow = btn;" & vbCrLf
    strHTML = strHTML & "  while (detailRow && detailRow.tagName !== 'TR') detailRow = detailRow.parentNode;" & vbCrLf
    strHTML = strHTML & "  if (detailRow && detailRow.previousElementSibling) {" & vbCrLf
    strHTML = strHTML & "    detailRow.style.display = 'none';" & vbCrLf
    strHTML = strHTML & "    var prevRow = detailRow.previousElementSibling;" & vbCrLf
    strHTML = strHTML & "    if (prevRow.classList) prevRow.classList.remove('selected');" & vbCrLf
    strHTML = strHTML & "    var icon = prevRow.querySelector ? prevRow.querySelector('.expand-icon') : null;" & vbCrLf
    strHTML = strHTML & "    if (icon) icon.innerHTML = '&#9654;';" & vbCrLf
    strHTML = strHTML & "  }" & vbCrLf
    strHTML = strHTML & "}" & vbCrLf
    strHTML = strHTML & "</script>" & vbCrLf

    strHTML = strHTML & "</div>" & vbCrLf
    strHTML = strHTML & "</body>" & vbCrLf
    strHTML = strHTML & "</html>" & vbCrLf

    strNombreArchivo = "InformeSLA_Corporate_" & Format(Now(), "yyyyMMdd_HHmm") & ".html"

    Dim strDirectorio As String
    strDirectorio = ""
    On Error Resume Next
    strDirectorio = m_ObjEntorno.URLDirectorioLocal
    On Error GoTo errores
    If strDirectorio = "" Then
        strDirectorio = Environ("USERPROFILE") & "\Documents"
    End If

    If Right(strDirectorio, 1) <> "\" And Right(strDirectorio, 1) <> "/" Then
        strDirectorio = strDirectorio & "\"
    End If

    strRutaArchivo = strDirectorio & strNombreArchivo

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objArchivo = objFSO.CreateTextFile(strRutaArchivo, True, True)
    objArchivo.Write strHTML
    objArchivo.Close
    Set objArchivo = Nothing
    Set objFSO = Nothing

    Shell "cmd /c start """ & strNombreArchivo & """ """ & strRutaArchivo & """", 0

    rcdEventos.Close
    Set rcdEventos = Nothing

    GenerarInformeHTML_Corporate = strRutaArchivo
    Exit Function

errores:
    Dim lngErr As Long
    Dim strErr As String
    lngErr = Err.Number
    strErr = Err.Description

    On Error Resume Next
    If Not rcdEventos Is Nothing Then rcdEventos.Close: Set rcdEventos = Nothing
    If Not objArchivo Is Nothing Then objArchivo.Close: Set objArchivo = Nothing
    If Not objFSO Is Nothing Then Set objFSO = Nothing
    On Error GoTo 0

    p_Error = "Error en GenerarInformeHTML_Corporate: " & lngErr & " - " & strErr
    Debug.Print p_Error
    GenerarInformeHTML_Corporate = "#ERR|" & p_Error
End Function

'===========================================================
' CalcularCumplimiento
' Calcula los KPIs de cumplimiento SLA para un evento
' Delega en SLAReportService.CalcularCumplimientoSLA
'-----------------------------------------------------------
Private Sub CalcularCumplimiento( _
    ByRef p_Fields As DAO.Fields, _
    ByRef p_CumpleTRES As Boolean, _
    ByRef p_CumpleTRCM As Boolean, _
    ByRef p_CumpleTRSS As Boolean, _
    ByRef p_CumpleGlobal As Boolean)

    Call SLAReportService.CalcularCumplimientoSLA(p_Fields, p_CumpleTRES, p_CumpleTRCM, p_CumpleTRSS, p_CumpleGlobal)
End Sub


