Attribute VB_Name = "Test_InformePublicacionSalida"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_InformePublicacionSalida_BloqueaExcel() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: tipo solicitado Excel"
    logs(1) = "2. Act: ValidarTipoSalidaInformePublicacion"
    logs(2) = "3. Assert: Excel no esta permitido"

    Dim m_Error As String

    If ValidarTipoSalidaInformePublicacion(EnumTipoInformePublicacion.Excel, m_Error) Then
        Test_InformePublicacionSalida_BloqueaExcel = BuildFail("Excel no debe permitirse para informe de publicacion", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "Excel", vbTextCompare) = 0 Then
        Test_InformePublicacionSalida_BloqueaExcel = BuildFail("El error debe mencionar que Excel no esta permitido", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_BloqueaExcel = BuildOk("excel_bloqueado", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_BloqueaExcel = BuildFail("Test_InformePublicacionSalida_BloqueaExcel: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_PermiteHTML() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: tipo solicitado HTML"
    logs(1) = "2. Act: ValidarTipoSalidaInformePublicacion"
    logs(2) = "3. Assert: HTML esta permitido"

    Dim m_Error As String

    If Not ValidarTipoSalidaInformePublicacion(EnumTipoInformePublicacion.html, m_Error) Then
        Test_InformePublicacionSalida_PermiteHTML = BuildFail("HTML debe permitirse: " & m_Error, logs)
        Exit Function
    End If

    If m_Error <> "" Then
        Test_InformePublicacionSalida_PermiteHTML = BuildFail("HTML no debe informar error: " & m_Error, logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_PermiteHTML = BuildOk("html_permitido", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_PermiteHTML = BuildFail("Test_InformePublicacionSalida_PermiteHTML: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_GenerarInformeBloqueaExcel() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: GenerarInformePublicacion con tipo Excel"
    logs(1) = "2. Act: llamada al dispatcher sin edicion"
    logs(2) = "3. Assert: se bloquea antes de generar informe"
    logs(3) = "4. Assert: informa error de Excel no permitido"

    Dim m_Error As String
    Dim m_URL As String
    Dim m_Edicion As Edicion

    On Error Resume Next
    m_URL = GenerarInformePublicacion(m_Edicion, EnumTipoInformePublicacion.Excel, , , , m_Error)
    Err.Clear
    On Error GoTo EH

    If m_Error = "" Then
        Test_InformePublicacionSalida_GenerarInformeBloqueaExcel = BuildFail("GenerarInformePublicacion debe bloquear Excel", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "Excel", vbTextCompare) = 0 Then
        Test_InformePublicacionSalida_GenerarInformeBloqueaExcel = BuildFail("El error debe mencionar Excel no permitido", logs)
        Exit Function
    End If

    If m_URL <> "" Then
        Test_InformePublicacionSalida_GenerarInformeBloqueaExcel = BuildFail("No debe devolver URL al bloquear Excel", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_GenerarInformeBloqueaExcel = BuildOk("generar_informe_excel_bloqueado", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_GenerarInformeBloqueaExcel = BuildFail("Test_InformePublicacionSalida_GenerarInformeBloqueaExcel: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_PublicarGeneraPDF() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: contexto publicacion"
    logs(1) = "2. Act: InformePublicacionDebeGenerarPDF"
    logs(2) = "3. Assert: publicacion genera PDF"

    If Not InformePublicacionDebeGenerarPDF(EnumSiNo.Sí) Then
        Test_InformePublicacionSalida_PublicarGeneraPDF = BuildFail("La publicacion debe generar PDF", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_PublicarGeneraPDF = BuildOk("publicacion_pdf", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_PublicarGeneraPDF = BuildFail("Test_InformePublicacionSalida_PublicarGeneraPDF: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_ManualNoGeneraPDF() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: contexto manual"
    logs(1) = "2. Act: InformePublicacionDebeGenerarPDF"
    logs(2) = "3. Assert: manual queda HTML"

    If InformePublicacionDebeGenerarPDF(EnumSiNo.No) Then
        Test_InformePublicacionSalida_ManualNoGeneraPDF = BuildFail("El informe manual debe quedar en HTML, no PDF", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_ManualNoGeneraPDF = BuildOk("manual_html", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_ManualNoGeneraPDF = BuildFail("Test_InformePublicacionSalida_ManualNoGeneraPDF: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_TipoDocumentoPDFGeneraPDF() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: selector de tipo documento en PDF"
    logs(1) = "2. Act: InformePublicacionDebeGenerarPDFDesdeTipoDocumento"
    logs(2) = "3. Assert: PDF activa generacion PDF"

    If Not InformePublicacionDebeGenerarPDFDesdeTipoDocumento("PDF") Then
        Test_InformePublicacionSalida_TipoDocumentoPDFGeneraPDF = BuildFail("El selector PDF debe activar p_GenerarPDF", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_TipoDocumentoPDFGeneraPDF = BuildOk("tipo_documento_pdf", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_TipoDocumentoPDFGeneraPDF = BuildFail("Test_InformePublicacionSalida_TipoDocumentoPDFGeneraPDF: " & Err.Description, logs)
End Function

Public Function Test_InformePublicacionSalida_TipoDocumentoHTMLNoGeneraPDF() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: selector de tipo documento en HTML"
    logs(1) = "2. Act: InformePublicacionDebeGenerarPDFDesdeTipoDocumento"
    logs(2) = "3. Assert: HTML conserva salida HTML"

    If InformePublicacionDebeGenerarPDFDesdeTipoDocumento("HTML") Then
        Test_InformePublicacionSalida_TipoDocumentoHTMLNoGeneraPDF = BuildFail("El selector HTML no debe activar p_GenerarPDF", logs)
        Exit Function
    End If

    Test_InformePublicacionSalida_TipoDocumentoHTMLNoGeneraPDF = BuildOk("tipo_documento_html", logs)
    Exit Function
EH:
    Test_InformePublicacionSalida_TipoDocumentoHTMLNoGeneraPDF = BuildFail("Test_InformePublicacionSalida_TipoDocumentoHTMLNoGeneraPDF: " & Err.Description, logs)
End Function
