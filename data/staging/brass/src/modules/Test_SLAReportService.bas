Attribute VB_Name = "Test_SLAReportService"
Option Compare Database
Option Explicit

'===========================================================
' Test_SLAReportService - html-informe-sla
' Tests unitarios para SLAReportService.ConstruirSQLEventosFranqueados
' Ejecutar desde Immediate Window: Test_SLAReportService.RunAllTests
'===========================================================

Private Type typResult
    testName As String
    passed As Boolean
    message As String
End Type

Private m_results() As typResult
Private m_resultCount As Long
Private m_message As String

'===========================================================
' Entry Point - Ejecutar todos los tests
'===========================================================
Public Function RunAllTests() As Boolean
    Dim allPassed As Boolean
    
    On Error GoTo errores
    
    Erase m_results
    m_resultCount = 0
    allPassed = True
    
    Debug.Print ""
    Debug.Print "============================================================"
    Debug.Print " BATERÍA DE TESTS - Test_SLAReportService (html-informe-sla)"
    Debug.Print "============================================================"
    Debug.Print ""
    
    '--- SRV-001: Fecha inicio inválida ---
    Debug.Print "--- SRV-001: Fecha inicio inválida ---"
    AddResult "Test_SRV001_FechaInicioInvalida", Test_SRV001_FechaInicioInvalida
    AddResult "Test_SRV001_FechaInicioVacia", Test_SRV001_FechaInicioVacia
    
    '--- SRV-002: Fecha fin inválida ---
    Debug.Print "--- SRV-002: Fecha fin inválida ---"
    AddResult "Test_SRV002_FechaFinInvalida", Test_SRV002_FechaFinInvalida
    AddResult "Test_SRV002_FechaFinVacia", Test_SRV002_FechaFinVacia
    
    '--- SRV-003: Fecha inicio > fecha fin ---
    Debug.Print "--- SRV-003: Fecha inicio > fecha fin ---"
    AddResult "Test_SRV003_FechaInicioMayorQueFin", Test_SRV003_FechaInicioMayorQueFin
    
    '--- SRV-004: SQL contiene Franqueado = True ---
    Debug.Print "--- SRV-004: SQL contiene Franqueado = True ---"
    AddResult "Test_SRV004_ContieneFranqueadoTrue", Test_SRV004_ContieneFranqueadoTrue
    
    '--- SRV-005: SQL contiene condición de datos SLA ---
    Debug.Print "--- SRV-005: SQL contiene condición de datos SLA ---"
    AddResult "Test_SRV005_ContieneCondicionSLA", Test_SRV005_ContieneCondicionSLA
    
    '--- SRV-006: SQL contiene campos necesarios para Excel/HTML ---
    Debug.Print "--- SRV-006: SQL contiene campos necesarios ---"
    AddResult "Test_SRV006_ContieneCamposNecesarios", Test_SRV006_ContieneCamposNecesarios
    
    '--- SRV-007: SQL con fechas válidas retorna SQL no vacía ---
    Debug.Print "--- SRV-007: SQL retorna string no vacío ---"
    AddResult "Test_SRV007_SQLNoVacia", Test_SRV007_SQLNoVacia
    
    '--- SRV-008: SQL incluye ORDER BY ---
    Debug.Print "--- SRV-008: SQL incluye ORDER BY ---"
    AddResult "Test_SRV008_ContieneOrderBy", Test_SRV008_ContieneOrderBy
    
    '--- Summary ---
    PrintSummary allPassed
    
    RunAllTests = allPassed
    Exit Function
    
errores:
    Debug.Print "[ERROR] Error en RunAllTests: " & Err.Description
    RunAllTests = False
End Function

'===========================================================
' AddResult - Wrapper para ejecutar test y registrar resultado
'===========================================================
Private Sub AddResult(testName As String, result As Boolean)
    Dim r As typResult
    r.testName = testName
    r.passed = result
    r.message = m_message
    ReDim Preserve m_results(m_resultCount)
    m_results(m_resultCount) = r
    m_resultCount = m_resultCount + 1
    
    If result Then
        Debug.Print "[OK] " & testName & ": " & m_message
    Else
        Debug.Print "[FAIL] " & testName & ": " & m_message
    End If
End Sub

'===========================================================
' PrintSummary - Imprime resumen de resultados
'===========================================================
Private Sub PrintSummary(ByRef allPassed As Boolean)
    Dim i As Long
    Dim passedCount As Long
    
    passedCount = 0
    For i = 0 To m_resultCount - 1
        If m_results(i).passed Then passedCount = passedCount + 1
    Next i
    
    Debug.Print ""
    Debug.Print "============================================================"
    Debug.Print " RESUMEN: " & passedCount & "/" & m_resultCount & " tests aprobados"
    Debug.Print "============================================================"
    
    If passedCount < m_resultCount Then
        Debug.Print ""
        Debug.Print "Tests FALLIDOS:"
        For i = 0 To m_resultCount - 1
            If Not m_results(i).passed Then
                Debug.Print "  - " & m_results(i).testName & ": " & m_results(i).message
            End If
        Next i
        allPassed = False
    End If
End Sub

'===========================================================
' SRV-001: Fecha inicio inválida
'===========================================================
Private Function Test_SRV001_FechaInicioInvalida() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("no-es-fecha", "01/01/2024", strError)
    
    If strResult = "" And strError <> "" Then
        m_message = "Fecha inicio inválida detectada correctamente. Error: " & strError
        Test_SRV001_FechaInicioInvalida = True
    Else
        m_message = "NO detectada fecha inicio inválida. Resultado: '" & strResult & "', Error: '" & strError & "'"
        Test_SRV001_FechaInicioInvalida = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV001_FechaInicioInvalida = False
End Function

Private Function Test_SRV001_FechaInicioVacia() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("", "01/01/2024", strError)
    
    If strResult = "" And strError <> "" Then
        m_message = "Fecha inicio vacía detectada correctamente. Error: " & strError
        Test_SRV001_FechaInicioVacia = True
    Else
        m_message = "NO detectada fecha inicio vacía. Resultado: '" & strResult & "', Error: '" & strError & "'"
        Test_SRV001_FechaInicioVacia = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV001_FechaInicioVacia = False
End Function

'===========================================================
' SRV-002: Fecha fin inválida
'===========================================================
Private Function Test_SRV002_FechaFinInvalida() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "no-es-fecha", strError)
    
    If strResult = "" And strError <> "" Then
        m_message = "Fecha fin inválida detectada correctamente. Error: " & strError
        Test_SRV002_FechaFinInvalida = True
    Else
        m_message = "NO detectada fecha fin inválida. Resultado: '" & strResult & "', Error: '" & strError & "'"
        Test_SRV002_FechaFinInvalida = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV002_FechaFinInvalida = False
End Function

Private Function Test_SRV002_FechaFinVacia() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "", strError)
    
    If strResult = "" And strError <> "" Then
        m_message = "Fecha fin vacía detectada correctamente. Error: " & strError
        Test_SRV002_FechaFinVacia = True
    Else
        m_message = "NO detectada fecha fin vacía. Resultado: '" & strResult & "', Error: '" & strError & "'"
        Test_SRV002_FechaFinVacia = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV002_FechaFinVacia = False
End Function

'===========================================================
' SRV-003: Fecha inicio > fecha fin
'===========================================================
Private Function Test_SRV003_FechaInicioMayorQueFin() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    ' Una fecha inicio mayor que la fecha fin debe ser detectada
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "01/01/2023", strError)
    
    If strResult = "" And InStr(1, strError, "fecha de inicio no puede ser mayor", vbTextCompare) > 0 Then
        m_message = "Rango invertido detectado correctamente. Error: " & strError
        Test_SRV003_FechaInicioMayorQueFin = True
    Else
        m_message = "NO detectado rango invertido correctamente. Resultado: '" & strResult & "', Error: '" & strError & "'"
        Test_SRV003_FechaInicioMayorQueFin = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV003_FechaInicioMayorQueFin = False
End Function

'===========================================================
' SRV-004: SQL contiene Franqueado = True
'===========================================================
Private Function Test_SRV004_ContieneFranqueadoTrue() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "31/12/2024", strError)
    
    If strResult <> "" And InStr(1, strResult, "Franqueado = True", vbTextCompare) > 0 Then
        m_message = "SQL contiene 'Franqueado = True' correctamente"
        Test_SRV004_ContieneFranqueadoTrue = True
    Else
        m_message = "SQL NO contiene 'Franqueado = True'. Resultado: " & Left(strResult, 200)
        Test_SRV004_ContieneFranqueadoTrue = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV004_ContieneFranqueadoTrue = False
End Function

'===========================================================
' SRV-005: SQL contiene condición de datos SLA
'===========================================================
Private Function Test_SRV005_ContieneCondicionSLA() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "31/12/2024", strError)
    
    If strResult <> "" Then
        ' Verificar que incluye al menos una condición SLA (IS NOT NULL)
        Dim bTieneCondicionSLA As Boolean
        bTieneCondicionSLA = InStr(1, strResult, "IS NOT NULL", vbTextCompare) > 0
        
        If bTieneCondicionSLA Then
            m_message = "SQL contiene condición 'IS NOT NULL' para datos SLA"
            Test_SRV005_ContieneCondicionSLA = True
        Else
            m_message = "SQL NO contiene condición 'IS NOT NULL' para datos SLA"
            Test_SRV005_ContieneCondicionSLA = False
        End If
    Else
        m_message = "SQL retornó vacío. Error: " & strError
        Test_SRV005_ContieneCondicionSLA = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV005_ContieneCondicionSLA = False
End Function

'===========================================================
' SRV-006: SQL contiene campos necesarios para Excel/HTML
'===========================================================
Private Function Test_SRV006_ContieneCamposNecesarios() As Boolean
    Dim strError As String
    Dim strResult As String
    Dim arrCamposRequeridos As Variant
    Dim i As Integer
    Dim bFaltaCampo As Boolean
    Dim strCamposFaltantes As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "31/12/2024", strError)
    
    If strResult = "" Then
        m_message = "SQL retornó vacío. Error: " & strError
        Test_SRV006_ContieneCamposNecesarios = False
        Exit Function
    End If
    
    ' Campos requeridos para el informe (misma lista que usa cmdExportarExcel)
    arrCamposRequeridos = Array("IDEVENTO", "FECHAALTAEVENTO", "BUI", "SUBSISTEMA", "CRITICIDAD", _
        "FechaRecepcionNotificacion", "FechaInicioContactoCliente", "IncidenciaAveriaOReparacion", _
        "TipoReparacion", "Urgente", "EventoConServicioAfectado", "FechaRestablecimientoServicio", _
        "FechaInicioTiempoAdquisicion", "FechaFinTiempoAdquisicion")
    
    bFaltaCampo = False
    strCamposFaltantes = ""
    
    For i = LBound(arrCamposRequeridos) To UBound(arrCamposRequeridos)
        If InStr(1, strResult, CStr(arrCamposRequeridos(i)), vbTextCompare) = 0 Then
            bFaltaCampo = True
            strCamposFaltantes = strCamposFaltantes & arrCamposRequeridos(i) & ", "
        End If
    Next i
    
    If Not bFaltaCampo Then
        m_message = "SQL contiene todos los campos requeridos para Excel/HTML"
        Test_SRV006_ContieneCamposNecesarios = True
    Else
        m_message = "SQL faltan campos: " & strCamposFaltantes
        Test_SRV006_ContieneCamposNecesarios = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV006_ContieneCamposNecesarios = False
End Function

'===========================================================
' SRV-007: SQL con fechas válidas retorna SQL no vacía
'===========================================================
Private Function Test_SRV007_SQLNoVacia() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "31/12/2024", strError)
    
    If strResult <> "" And strError = "" Then
        m_message = "SQL retorna string no vacío con fechas válidas. Longitud: " & Len(strResult)
        Test_SRV007_SQLNoVacia = True
    Else
        m_message = "SQL retornó vacío o con error. Resultado: '" & Left(strResult, 100) & "', Error: '" & strError & "'"
        Test_SRV007_SQLNoVacia = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV007_SQLNoVacia = False
End Function

'===========================================================
' SRV-008: SQL incluye ORDER BY
'===========================================================
Private Function Test_SRV008_ContieneOrderBy() As Boolean
    Dim strError As String
    Dim strResult As String
    
    On Error GoTo errores
    
    strResult = SLAReportService.ConstruirSQLEventosFranqueados("01/01/2024", "31/12/2024", strError)
    
    If strResult <> "" And InStr(1, strResult, "ORDER BY", vbTextCompare) > 0 Then
        m_message = "SQL contiene 'ORDER BY' correctamente"
        Test_SRV008_ContieneOrderBy = True
    Else
        m_message = "SQL NO contiene 'ORDER BY'. Resultado: " & Left(strResult, 200)
        Test_SRV008_ContieneOrderBy = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SRV008_ContieneOrderBy = False
End Function

