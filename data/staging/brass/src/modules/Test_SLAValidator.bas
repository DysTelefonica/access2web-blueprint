Attribute VB_Name = "Test_SLAValidator"
Option Compare Database
Option Explicit

'===========================================================
' Test_SLAValidator - Spec-009
' Tests unitarios para SLAValidator.ValidarCamposSLA
' Ejecutar desde Immediate Window: Test_SLAValidator.RunAllTests
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
    Debug.Print " BATERÍA DE TESTS - Test_SLAValidator (Spec-009)"
    Debug.Print "============================================================"
    Debug.Print ""
    
    '--- SLA-001: Regla general - fecha fin sin fecha inicio ---
    Debug.Print "--- SLA-001: Fecha fin sin fecha inicio ---"
    AddResult "Test_SLA001_TRES_FaltaFechaRecepcion", Test_SLA001_TRES_FaltaFechaRecepcion
    AddResult "Test_SLA001_TRCM_FaltaFechaInicioAdquisicion", Test_SLA001_TRCM_FaltaFechaInicioAdquisicion
    AddResult "Test_SLA001_TRSS_FaltaServicioAfectado", Test_SLA001_TRSS_FaltaServicioAfectado
    
    '--- SLA-007: Coherencia cronológica ---
    Debug.Print ""
    Debug.Print "--- SLA-007: Coherencia cronológica ---"
    AddResult "Test_SLA007_TRES_FechaAnterior", Test_SLA007_TRES_FechaAnterior
    AddResult "Test_SLA007_TRCM_FechaAnterior", Test_SLA007_TRCM_FechaAnterior
    AddResult "Test_SLA007_TRSS_FechaAnterior", Test_SLA007_TRSS_FechaAnterior
    
    '--- SLA-002/003/004: Franqueo TRCM incompleto ---
    Debug.Print ""
    Debug.Print "--- SLA-002/003/004: Franqueo TRCM incompleto ---"
    AddResult "Test_SLA002_003_004_FranqueoIncompleto", Test_SLA002_003_004_FranqueoIncompleto
    
    '--- SLA-005: Franqueo TRSS incompleto ---
    Debug.Print ""
    Debug.Print "--- SLA-005: Franqueo TRSS incompleto ---"
    AddResult "Test_SLA005_FranqueoSinRestablecimiento", Test_SLA005_FranqueoSinRestablecimiento
    
    '--- Edición permite incompleto ---
    Debug.Print ""
    Debug.Print "--- Edición permite incompleto ---"
    AddResult "Test_EdicionPermiteIncompleto", Test_EdicionPermiteIncompleto
    
    '--- Caso válido ---
    Debug.Print ""
    Debug.Print "--- Caso válido ---"
    AddResult "Test_Valido_TodosLosCampos", Test_Valido_TodosLosCampos
    
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
' Test_SLA001_TRES_FaltaFechaRecepcion
' TRES: Existe FechaInicioContacto sin FechaRecepcion
'===========================================================
Private Function Test_SLA001_TRES_FaltaFechaRecepcion() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="", _
        p_FechaInicioContacto:="01/01/2024", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-001" And strMensaje <> "" Then
        m_message = "SLA-001 detectado correctamente. Mensaje: " & strMensaje
        Test_SLA001_TRES_FaltaFechaRecepcion = True
    Else
        m_message = "SLA-001 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA001_TRES_FaltaFechaRecepcion = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA001_TRES_FaltaFechaRecepcion = False
End Function

'===========================================================
' Test_SLA001_TRCM_FaltaFechaInicioAdquisicion
' TRCM: Existe FechaFinAdquisicion sin FechaInicioAdquisicion
'===========================================================
Private Function Test_SLA001_TRCM_FaltaFechaInicioAdquisicion() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="", _
        p_FechaInicioContacto:="", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="01/01/2024", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-001" And strMensaje <> "" Then
        m_message = "SLA-001 detectado correctamente (TRCM). Mensaje: " & strMensaje
        Test_SLA001_TRCM_FaltaFechaInicioAdquisicion = True
    Else
        m_message = "SLA-001 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA001_TRCM_FaltaFechaInicioAdquisicion = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA001_TRCM_FaltaFechaInicioAdquisicion = False
End Function

'===========================================================
' Test_SLA001_TRSS_FaltaServicioAfectado
' TRSS: Existe FechaRestablecimiento sin servicio afectado
'===========================================================
Private Function Test_SLA001_TRSS_FaltaServicioAfectado() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="01/01/2024", _
        p_FechaInicioContacto:="", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="02/01/2024", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-001" And strMensaje <> "" Then
        m_message = "SLA-001 detectado correctamente (TRSS). Mensaje: " & strMensaje
        Test_SLA001_TRSS_FaltaServicioAfectado = True
    Else
        m_message = "SLA-001 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA001_TRSS_FaltaServicioAfectado = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA001_TRSS_FaltaServicioAfectado = False
End Function

'===========================================================
' Test_SLA007_TRES_FechaAnterior
' TRES: FechaInicioContacto < FechaRecepcion
'===========================================================
Private Function Test_SLA007_TRES_FechaAnterior() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="02/01/2024", _
        p_FechaInicioContacto:="01/01/2024", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-007" And strMensaje <> "" Then
        m_message = "SLA-007 detectado correctamente (TRES). Mensaje: " & strMensaje
        Test_SLA007_TRES_FechaAnterior = True
    Else
        m_message = "SLA-007 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA007_TRES_FechaAnterior = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA007_TRES_FechaAnterior = False
End Function

'===========================================================
' Test_SLA007_TRCM_FechaAnterior
' TRCM: FechaFinAdquisicion < FechaInicioAdquisicion
'===========================================================
Private Function Test_SLA007_TRCM_FechaAnterior() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="", _
        p_FechaInicioContacto:="", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="02/01/2024", _
        p_FechaFinAdquisicion:="01/01/2024", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-007" And strMensaje <> "" Then
        m_message = "SLA-007 detectado correctamente (TRCM). Mensaje: " & strMensaje
        Test_SLA007_TRCM_FechaAnterior = True
    Else
        m_message = "SLA-007 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA007_TRCM_FechaAnterior = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA007_TRCM_FechaAnterior = False
End Function

'===========================================================
' Test_SLA007_TRSS_FechaAnterior
' TRSS: FechaRestablecimiento < FechaRecepcion
'===========================================================
Private Function Test_SLA007_TRSS_FechaAnterior() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="02/01/2024", _
        p_FechaInicioContacto:="", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="SI", _
        p_FechaRestablecimiento:="01/01/2024", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-007" And strMensaje <> "" Then
        m_message = "SLA-007 detectado correctamente (TRSS). Mensaje: " & strMensaje
        Test_SLA007_TRSS_FechaAnterior = True
    Else
        m_message = "SLA-007 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA007_TRSS_FechaAnterior = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA007_TRSS_FechaAnterior = False
End Function

'===========================================================
' Test_SLA002_003_004_FranqueoIncompleto
' Franqueo con Incidencia="Sí" sin campos de adquisición
' Verifica que los mensajes de SLA-002, SLA-003 y SLA-004
' se acumulen correctamente (el primero en setearse determina
' el código de retorno, pero todos los mensajes deben aparecer)
'===========================================================
Private Function Test_SLA002_003_004_FranqueoIncompleto() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=True, _
        p_FechaRecepcion:="01/01/2024", _
        p_FechaInicioContacto:="01/01/2024", _
        p_IncidenciaAveria:="SI", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=True, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    ' El código retornado debe ser SLA-002 (primera validación que falla)
    ' Los mensajes de las tres validaciones deben estar acumulados
    If strResultado = "SLA-002" Then
        If InStr(strMensaje, "fecha de inicio de adquisicion") > 0 And _
           InStr(strMensaje, "fecha de fin de adquisicion") > 0 And _
           InStr(strMensaje, "tipo de reparacion") > 0 Then
            m_message = "SLA-002/003/004 acumulados correctamente. Mensaje: " & strMensaje
            Test_SLA002_003_004_FranqueoIncompleto = True
        Else
            m_message = "Mensajes no acumulados completamente. Mensaje: " & strMensaje
            Test_SLA002_003_004_FranqueoIncompleto = False
        End If
    Else
        m_message = "Código esperado SLA-002, obtenido: " & strResultado & ". Mensaje: " & strMensaje
        Test_SLA002_003_004_FranqueoIncompleto = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA002_003_004_FranqueoIncompleto = False
End Function

'===========================================================
' Test_SLA005_FranqueoSinRestablecimiento
' Franqueo con ServicioAfectado="Sí" sin FechaRestablecimiento
'===========================================================
Private Function Test_SLA005_FranqueoSinRestablecimiento() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=True, _
        p_FechaRecepcion:="01/01/2024", _
        p_FechaInicioContacto:="01/01/2024", _
        p_IncidenciaAveria:="NO", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="SI", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "SLA-005" And strMensaje <> "" Then
        m_message = "SLA-005 detectado correctamente. Mensaje: " & strMensaje
        Test_SLA005_FranqueoSinRestablecimiento = True
    Else
        m_message = "SLA-005 NO detectado correctamente. Resultado: " & strResultado & ", Mensaje: " & strMensaje
        Test_SLA005_FranqueoSinRestablecimiento = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_SLA005_FranqueoSinRestablecimiento = False
End Function

'===========================================================
' Test_EdicionPermiteIncompleto
' Edición permite Incidencia="Sí" sin campos de adquisición
' (Las validaciones de franqueo no aplican en edición)
'===========================================================
Private Function Test_EdicionPermiteIncompleto() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    ' En modo edición (p_EnFranqueo=False), no debe fallar por
    ' campos TRCM faltantes cuando Incidencia="Sí"
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=False, _
        p_FechaRecepcion:="01/01/2024", _
        p_FechaInicioContacto:="01/01/2024", _
        p_IncidenciaAveria:="SI", _
        p_FechaInicioAdquisicion:="", _
        p_FechaFinAdquisicion:="", _
        p_TipoReparacion:="", _
        p_Urgente:=Null, _
        p_EventoConServicioAfectado:="NO", _
        p_FechaRestablecimiento:="", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "" Then
        m_message = "Edición permite campos TRCM incompletos (comportamiento correcto)"
        Test_EdicionPermiteIncompleto = True
    Else
        m_message = "Edición NO permite campos TRCM incompletos. Código: " & strResultado & ", Mensaje: " & strMensaje
        Test_EdicionPermiteIncompleto = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_EdicionPermiteIncompleto = False
End Function

'===========================================================
' Test_Valido_TodosLosCampos
' Caso válido: todos los campos correctamente preenchados
'===========================================================
Private Function Test_Valido_TodosLosCampos() As Boolean
    Dim strResultado As String
    Dim strMensaje As String
    
    On Error GoTo errores
    
    strResultado = SLAValidator.ValidarCamposSLA( _
        p_EnFranqueo:=True, _
        p_FechaRecepcion:="01/01/2024", _
        p_FechaInicioContacto:="02/01/2024", _
        p_IncidenciaAveria:="SI", _
        p_FechaInicioAdquisicion:="03/01/2024", _
        p_FechaFinAdquisicion:="04/01/2024", _
        p_TipoReparacion:="Contratista", _
        p_Urgente:=True, _
        p_EventoConServicioAfectado:="SI", _
        p_FechaRestablecimiento:="05/01/2024", _
        p_Mensaje:=strMensaje)
    
    If strResultado = "" And strMensaje = "" Then
        m_message = "Caso válido aceptado correctamente"
        Test_Valido_TodosLosCampos = True
    Else
        m_message = "Caso válido rechazado. Código: " & strResultado & ", Mensaje: " & strMensaje
        Test_Valido_TodosLosCampos = False
    End If
    Exit Function
    
errores:
    m_message = "Error: " & Err.Description
    Test_Valido_TodosLosCampos = False
End Function

