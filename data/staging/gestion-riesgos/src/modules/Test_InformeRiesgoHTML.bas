Attribute VB_Name = "Test_InformeRiesgoHTML"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_InformeHTML_ControlCambios_LeyendaDebajoTabla() As String
    Dim logs(0 To 4) As String
    Dim html As String
    Dim idxTable As Long
    Dim idxLegend As Long

    logs(0) = "1. Arrange: bloque interno con tabla"
    html = InformeRiesgoHTML.InformeHTML_ComponerSeccionControlCambiosParaTest("<table class='cc-table'></table>", False)

    logs(1) = "2. Act: buscar posiciones de tabla y leyenda"
    idxTable = InStr(1, html, "<table class='cc-table'>", vbTextCompare)
    idxLegend = InStr(1, html, "<div class='legend-container'>", vbTextCompare)

    logs(2) = "3. Assert: tabla debe aparecer antes que leyenda"
    logs(3) = "idxTable=" & CStr(idxTable) & ", idxLegend=" & CStr(idxLegend)
    If idxTable = 0 Or idxLegend = 0 Or idxLegend <= idxTable Then
        Test_InformeHTML_ControlCambios_LeyendaDebajoTabla = BuildFail("La leyenda no quedó debajo de la tabla", logs)
        Exit Function
    End If

    logs(4) = "OK"
    Test_InformeHTML_ControlCambios_LeyendaDebajoTabla = BuildOk("legend_below_table_ok", logs)
End Function

Public Function Test_InformeHTML_FechasVacias_NoMuestraEtiquetasFin() As String
    Dim logs(0 To 3) As String
    Dim s As String

    logs(0) = "1. Arrange: fecha inicio válida y fines vacíos"
    logs(1) = "2. Act: compactar fechas"
    s = InformeRiesgoHTML.InformeHTML_CompactarFechasAccionParaTest(#1/15/2026#, Null, "")

    logs(2) = "3. Assert: no debe incluir etiquetas Fin prevista/Fin real"
    logs(3) = "Salida=" & s
    If InStr(1, s, "Fin prevista:", vbTextCompare) > 0 Or InStr(1, s, "Fin real:", vbTextCompare) > 0 Then
        Test_InformeHTML_FechasVacias_NoMuestraEtiquetasFin = BuildFail("Se renderizaron etiquetas de fin vacías", logs)
        Exit Function
    End If

    Test_InformeHTML_FechasVacias_NoMuestraEtiquetasFin = BuildOk("empty_end_dates_hidden_ok", logs)
End Function

Public Function Test_InformeHTML_EstadoCanonico_CierreForzado() As String
    Dim logs(0 To 3) As String
    Dim salida As String

    logs(0) = "1. Arrange: riesgo retirado con fecha cierre"
    logs(1) = "2. Act: resolver estado canónico"
    salida = InformeRiesgoHTML.InformeHTML_EstadoCanonicoParaTest(EnumRiesgoEstado.Retirado, "Retirado", "31/01/2026", Null)

    logs(2) = "3. Assert: estado Cerrado y fecha de cierre"
    logs(3) = "Salida=" & salida
    If InStr(1, salida, "Cerrado|31/01/2026", vbTextCompare) = 0 Then
        Test_InformeHTML_EstadoCanonico_CierreForzado = BuildFail("No aplicó estado/fecha canónica de cierre", logs)
        Exit Function
    End If

    Test_InformeHTML_EstadoCanonico_CierreForzado = BuildOk("canonical_closed_status_ok", logs)
End Function

Public Function Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb() As String
    Dim logs(0 To 5) As String
    Dim html As String

    logs(0) = "1. Arrange: shell representativo del informe de edicion en modo PDF"
    html = InformeRiesgoHTML.InformeHTML_ComponerShellEdicionParaTest(True)

    logs(1) = "2. Assert: no debe inyectar JavaScript web"
    If InStr(1, html, "<script", vbTextCompare) > 0 Or InStr(1, html, "function openTab", vbTextCompare) > 0 Then
        Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb = BuildFail("El HTML PDF no debe incluir scripts interactivos", logs)
        Exit Function
    End If

    logs(2) = "3. Assert: no debe depender de onclick ni buscador interactivo"
    If InStr(1, html, "onclick=", vbTextCompare) > 0 Or InStr(1, html, "sectionSearchInput", vbTextCompare) > 0 Then
        Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb = BuildFail("El HTML PDF no debe incluir controles web interactivos", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: no debe usar details/collapsible-section para revelar contenido"
    If InStr(1, html, "<details", vbTextCompare) > 0 Or InStr(1, html, "collapsible-section", vbTextCompare) > 0 Then
        Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb = BuildFail("El HTML PDF debe renderizar contenido estatico visible", logs)
        Exit Function
    End If

    logs(4) = "5. Assert: debe incluir salida especifica de impresion/PDF"
    If InStr(1, html, "print-only-cover", vbTextCompare) = 0 Or InStr(1, html, "<section class='report-section'", vbTextCompare) = 0 Then
        Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb = BuildFail("El HTML PDF debe incluir portada/secciones estaticas de informe", logs)
        Exit Function
    End If

    logs(5) = "OK"
    Test_InformeHTML_EdicionPDF_ExcluyeInteractividadWeb = BuildOk("pdf_static_html_ok", logs)
End Function

Public Function Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb() As String
    Dim logs(0 To 5) As String
    Dim html As String

    logs(0) = "1. Arrange: shell representativo del informe de edicion en modo HTML"
    html = InformeRiesgoHTML.InformeHTML_ComponerShellEdicionParaTest(False)

    logs(1) = "2. Assert: HTML normal conserva JavaScript web"
    If InStr(1, html, "<script", vbTextCompare) = 0 Or InStr(1, html, "function openTab", vbTextCompare) = 0 Then
        Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb = BuildFail("El HTML normal debe conservar scripts interactivos", logs)
        Exit Function
    End If

    logs(2) = "3. Assert: HTML normal conserva buscador interactivo"
    If InStr(1, html, "sectionSearchInput", vbTextCompare) = 0 Then
        Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb = BuildFail("El HTML normal debe conservar buscador de secciones", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: HTML normal no debe ocultar globalmente el buscador"
    If InStr(1, html, ".section-search-wrapper { display: none !important; }", vbTextCompare) > 0 Then
        Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb = BuildFail("El HTML normal no debe ocultar el buscador con CSS global", logs)
        Exit Function
    End If

    logs(4) = "5. Assert: HTML normal conserva details colapsables"
    If InStr(1, html, "<details", vbTextCompare) = 0 Or InStr(1, html, "collapsible-section", vbTextCompare) = 0 Then
        Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb = BuildFail("El HTML normal debe conservar secciones colapsables", logs)
        Exit Function
    End If

    logs(5) = "OK"
    Test_InformeHTML_EdicionHTML_ConservaInteractividadWeb = BuildOk("web_html_interactive_ok", logs)
End Function

' =========================================================================
' Punto 09 (issue #88) - Informe de publicabilidad muestra motivo en dos partes
' =========================================================================
' Acta Calidad 2026-06-25 + contrato staging 2026-06-22 (fila req-cal-09).
' Escenarios de aceptacion: el motivo (por que no se cumple un check) debe
' renderizarse SIEMPRE cuando el check es "No cumple", con dos partes
' separadas visualmente:
'   (a) que se evalua  -> <div class='check-text'>   (siempre visible)
'   (b) por que no se cumple -> <div class='check-detail'>
' Comportamientos esperados:
'   - Si hay detalle:    check-detail = detalle (justificacion literal)
'   - Si no hay detalle: check-detail = "(sin motivo registrado)"
'   - HTML safe para <, >, &, ", saltos de linea preservados, >255 chars OK
'   - Para checks "Cumple"/"NoAplica" se conserva el comportamiento legacy
'     (check-detail solo si existe y no esta vacio)
' =========================================================================

Public Function Test_InformeHTML_Publicabilidad_NoCumpleConDetalle_MuestraJustificacionLiteral() As String
    Dim logs(0 To 4) As String
    Dim m_Check As Object
    Dim s As String

    logs(0) = "1. Arrange: check NoCumple con texto + detalle (justificacion concreta)"
    Set m_Check = CreateObject("Scripting.Dictionary")
    m_Check("texto") = "Plan de mitigacion con FechaFinPrevista"
    m_Check("detalle") = "PM01 con fecha de fin prevista 03/04/2026, ya cerrado en edicion anterior"

    logs(1) = "2. Act: renderizar el check individual"
    s = InformeRiesgoHTML.InformeHTML_RenderCheckPublicabilidadParaTest(m_Check, 2)
    '                                  ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ EnumPublicabilidadCheckEstado.NoCumple = 2

    logs(2) = "3. Assert: check-detail contiene la justificacion literal"
    If InStr(1, s, "<div class='check-detail'>PM01 con fecha de fin prevista 03/04/2026, ya cerrado en edicion anterior</div>", vbTextCompare) = 0 Then
        logs(3) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_NoCumpleConDetalle_MuestraJustificacionLiteral = BuildFail("check-detail no contiene la justificacion literal cuando hay detalle", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: estado 'No cumple' renderizado como label"
    If InStr(1, s, "No cumple", vbTextCompare) = 0 Then
        logs(4) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_NoCumpleConDetalle_MuestraJustificacionLiteral = BuildFail("label 'No cumple' no aparece en el HTML", logs)
        Exit Function
    End If

    logs(4) = "OK"
    Test_InformeHTML_Publicabilidad_NoCumpleConDetalle_MuestraJustificacionLiteral = BuildOk("detalle_literal_ok", logs)
End Function

Public Function Test_InformeHTML_Publicabilidad_NoCumpleSinDetalle_MuestraSinMotivoRegistrado() As String
    Dim logs(0 To 3) As String
    Dim m_Check As Object
    Dim s As String

    logs(0) = "1. Arrange: check NoCumple solo con texto, sin clave 'detalle'"
    Set m_Check = CreateObject("Scripting.Dictionary")
    m_Check("texto") = "Falta evidencia del suministrador X"

    logs(1) = "2. Act: renderizar el check individual"
    s = InformeRiesgoHTML.InformeHTML_RenderCheckPublicabilidadParaTest(m_Check, 2)

    logs(2) = "3. Assert: check-detail contiene '(sin motivo registrado)' explicito"
    If InStr(1, s, "<div class='check-detail'>(sin motivo registrado)</div>", vbTextCompare) = 0 Then
        logs(3) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_NoCumpleSinDetalle_MuestraSinMotivoRegistrado = BuildFail("check-detail no muestra '(sin motivo registrado)' cuando falta detalle", logs)
        Exit Function
    End If

    logs(3) = "OK"
    Test_InformeHTML_Publicabilidad_NoCumpleSinDetalle_MuestraSinMotivoRegistrado = BuildOk("sin_motivo_ok", logs)
End Function

Public Function Test_InformeHTML_Publicabilidad_DetalleConCaracteresEspeciales_EscapaHTML() As String
    Dim logs(0 To 4) As String
    Dim m_Check As Object
    Dim s As String
    Dim detalleIn As String

    logs(0) = "1. Arrange: check NoCumple con detalle que contiene <, >, &, "" y saltos de linea"
    detalleIn = "Cota < 100 & > 50; comentario ""critico"" con 'apostrofes'" & vbCrLf & "linea 2 del detalle"
    Set m_Check = CreateObject("Scripting.Dictionary")
    m_Check("texto") = "Validacion de cota"
    m_Check("detalle") = detalleIn

    logs(1) = "2. Act: renderizar el check individual"
    s = InformeRiesgoHTML.InformeHTML_RenderCheckPublicabilidadParaTest(m_Check, 2)

    logs(2) = "3. Assert: < > & "" estan escapados (no rompen HTML)"
    If InStr(1, s, "&lt;", vbTextCompare) = 0 Or InStr(1, s, "&gt;", vbTextCompare) = 0 Or InStr(1, s, "&amp;", vbTextCompare) = 0 Then
        logs(3) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_DetalleConCaracteresEspeciales_EscapaHTML = BuildFail("caracteres <, >, & no se escaparon", logs)
        Exit Function
    End If
    If InStr(1, s, "< 100 & > 50", vbTextCompare) > 0 Then
        logs(3) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_DetalleConCaracteresEspeciales_EscapaHTML = BuildFail("< y > aparecen sin escapar dentro del check-detail", logs)
        Exit Function
    End If

    logs(3) = "4. Assert: saltos de linea preservados (no se truncan)"
    If InStr(1, s, vbCrLf, vbTextCompare) = 0 And InStr(1, s, "linea 2 del detalle", vbTextCompare) = 0 Then
        logs(4) = "Salida=" & s
        Test_InformeHTML_Publicabilidad_DetalleConCaracteresEspeciales_EscapaHTML = BuildFail("saltos de linea / contenido multilinea no preservado", logs)
        Exit Function
    End If

    logs(4) = "OK"
    Test_InformeHTML_Publicabilidad_DetalleConCaracteresEspeciales_EscapaHTML = BuildOk("html_safe_ok", logs)
End Function

Public Function Test_InformeHTML_Publicabilidad_DetalleLargo_NoTrunca() As String
    Dim logs(0 To 3) As String
    Dim m_Check As Object
    Dim s As String
    Dim detalleLargo As String
    Dim i As Long

    logs(0) = "1. Arrange: check NoCumple con detalle > 255 caracteres"
    detalleLargo = ""
    For i = 1 To 300
        detalleLargo = detalleLargo & "x"
    Next i
    ' Sanity: que el detalle sea de 300 chars
    If Len(detalleLargo) <> 300 Then
        Test_InformeHTML_Publicabilidad_DetalleLargo_NoTrunca = BuildFail("setup invalido: detalle no llega a 300 chars (Len=" & Len(detalleLargo) & ")", logs)
        Exit Function
    End If
    Set m_Check = CreateObject("Scripting.Dictionary")
    m_Check("texto") = "Detalle muy largo"
    m_Check("detalle") = detalleLargo

    logs(1) = "2. Act: renderizar el check individual"
    s = InformeRiesgoHTML.InformeHTML_RenderCheckPublicabilidadParaTest(m_Check, 2)

    logs(2) = "3. Assert: el detalle se renderiza completo (los 300 'x' aparecen en el HTML)"
    If InStr(1, s, detalleLargo, vbTextCompare) = 0 Then
        logs(3) = "Salida length=" & Len(s) & ", detalle length=" & Len(detalleLargo)
        Test_InformeHTML_Publicabilidad_DetalleLargo_NoTrunca = BuildFail("el detalle largo (>255 chars) se trunco en el HTML", logs)
        Exit Function
    End If

    logs(3) = "OK"
    Test_InformeHTML_Publicabilidad_DetalleLargo_NoTrunca = BuildOk("no_trunca_ok", logs)
End Function
