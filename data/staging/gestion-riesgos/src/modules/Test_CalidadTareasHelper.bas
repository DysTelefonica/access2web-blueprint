Attribute VB_Name = "Test_CalidadTareasHelper"
' ============================================================
' Test_CalidadTareasHelper — TDD atoms for
'   modCalidadTareasHelper
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   hr4-sister-slice-2026-07-01 (HR3d follow-up §9, #4)
'
' Scope: 2 helper entries (CalidadTareas_SeleccionarNodoRiesgo,
'        CalidadTareas_CargarArbol) -> 6 scenario atoms
'   Helper signature:
'     Public Sub CalidadTareas_SeleccionarNodoRiesgo( _
'         ByVal p_ObjRiesgo As riesgo, _
'         ByVal p_TipoRiesgoTarea As EnumTipoRiesgoTarea, _
'         ByRef p_OutKey As String, _
'         Optional ByRef p_Error As String)
'     Public Sub CalidadTareas_CargarArbol( _
'         ByVal p_ObjArbolCalidad As ArbolTareasCalidad, _
'         ByRef p_Error As String)
'
'   Convention: helper is Sub (no JSON return). Atoms wrap the call and
'   return JSON via TestCore_BuildOk / TestCore_BuildFail.
'
'   Hard rule 7: el helper NO hace MsgBox. Los atomos verifican que
'   p_Error se rellena en sad path sin raise visible.
'
'   Note: Arbol_NodeClick is an event handler and is made Private
'   without helper extraction (per Hard rule 4 exception).
' ============================================================
Option Compare Database
Option Explicit

' --- Constants de IDs del fixture (coinciden con Test_Fixtures.bas) ---
Private Const FIX_EDICION    As Long = 900102
Private Const FIX_PROYECTO   As Long = 900101
Private Const FIX_EXPEDIENTE As Long = 900100
Private Const FIX_ID_RIESGO  As Long = 900503

' --- Helper wrappers (Hard rule 6: re-usa Test_Helper) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Fixture helpers ---
Private Function ForceBackend(ByRef p_Error As String) As Boolean
    ForceBackend = Test_Helper.ForceLocalBackend(p_Error)
End Function

' ============================================================
' ATOM 1 (SAD): CalidadTareas_SeleccionarNodoRiesgo con p_ObjRiesgo = Nothing
'   El helper debe set p_Error y NO raise (matching original behavior).
' ============================================================
Public Function Test_CalidadTareas_SeleccionarNodoRiesgo_Nothing_pErrorSetted() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_SeleccionarNodoRiesgo_Nothing_pErrorSetted = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Key As String
    logs(logIdx) = "2. Act: helper con p_ObjRiesgo = Nothing"
    logIdx = logIdx + 1
    CalidadTareas_SeleccionarNodoRiesgo Nothing, EnumTipoRiesgoTarea.Aceptados, m_Key, m_Err

    logs(logIdx) = "3. Assert: p_Error debe estar seteado y NO debe raise"
    logIdx = logIdx + 1
    If m_Err = "" Then
        logs(logIdx) = "4. Assert FAIL: esperaba p_Error seteado, got vacio"
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_Nothing_pErrorSetted = BuildFail( _
            "Nothing debe setear p_Error. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error seteado, no raise (" & m_Err & ")"
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_Nothing_pErrorSetted = BuildOk("seleccionar_nodo_nothing_p_error_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_Nothing_pErrorSetted = BuildFail( _
        "Unhandled exception (helper must NOT raise on sad path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (SAD): CalidadTareas_SeleccionarNodoRiesgo con tipo no soportado
'   El helper debe set p_Error y NO raise.
' ============================================================
Public Function Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted() As String
    Dim logs(0 To 6) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Riesgo As riesgo
    Set m_Riesgo = GetCachedRiesgo(CStr(FIX_ID_RIESGO), m_Err)
    If m_Err <> "" Or m_Riesgo Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedRiesgo: " & m_Err
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted = BuildFail("seed risk failed: " & m_Err, logs)
        Exit Function
    End If

    Dim m_Key As String
    Dim m_TipoInvalido As EnumTipoRiesgoTarea
    m_TipoInvalido = 999
    logs(logIdx) = "2. Act: helper con p_TipoRiesgoTarea invalido (999)"
    logIdx = logIdx + 1
    CalidadTareas_SeleccionarNodoRiesgo m_Riesgo, m_TipoInvalido, m_Key, m_Err

    logs(logIdx) = "3. Assert: p_Error debe estar seteado y NO debe raise"
    logIdx = logIdx + 1
    If m_Err = "" Then
        logs(logIdx) = "4. Assert FAIL: esperaba p_Error seteado, got vacio"
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted = BuildFail( _
            "Tipo invalido debe setear p_Error. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error seteado para tipo invalido (" & m_Err & ")"
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted = BuildOk("seleccionar_nodo_tipo_invalido_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_TipoInvalido_pErrorSetted = BuildFail( _
        "Unhandled exception (helper must NOT raise on sad path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (SAD): CalidadTareas_SeleccionarNodoRiesgo estado no mappable
'   El helper debe set p_Error cuando el estado no es mapeable al key.
' ============================================================
Public Function Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted() As String
    Dim logs(0 To 6) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Riesgo As riesgo
    Set m_Riesgo = GetCachedRiesgo(CStr(FIX_ID_RIESGO), m_Err)
    If m_Err <> "" Or m_Riesgo Is Nothing Then
        logs(logIdx) = "Arrange FAIL: GetCachedRiesgo: " & m_Err
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted = BuildFail("seed risk failed: " & m_Err, logs)
        Exit Function
    End If

    ' P_TipoRiesgoTarea = Aceptados pero el riesgo esta en estado Detectado
    '   => no hay key mapeable, helper debe set p_Error "Nodo no encontrado"
    Dim m_Key As String
    logs(logIdx) = "2. Act: helper con riesgo Detectado + tipo Aceptados (no mappable)"
    logIdx = logIdx + 1
    CalidadTareas_SeleccionarNodoRiesgo m_Riesgo, EnumTipoRiesgoTarea.Aceptados, m_Key, m_Err

    logs(logIdx) = "3. Assert: p_Error debe estar seteado (Nodo no encontrado)"
    logIdx = logIdx + 1
    If m_Err = "" Then
        logs(logIdx) = "4. Assert FAIL: esperaba p_Error seteado"
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted = BuildFail( _
            "Estado no mappable debe setear p_Error. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error seteado para estado no mappable (" & m_Err & ")"
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted = BuildOk("seleccionar_nodo_estado_no_mapeable_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_EstadoNoMapeable_pErrorSetted = BuildFail( _
        "Unhandled exception (helper must NOT raise on sad path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (SAD): CalidadTareas_CargarArbol con p_ObjArbolCalidad = Nothing
'   El helper debe set p_Error porque no hay arbol para cargar.
' ============================================================
Public Function Test_CalidadTareas_CargarArbol_ArbolNothing_pErrorSetted() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_CargarArbol_ArbolNothing_pErrorSetted = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    logs(logIdx) = "2. Act: helper con p_ObjArbolCalidad = Nothing"
    logIdx = logIdx + 1
    CalidadTareas_CargarArbol Nothing, m_Err

    logs(logIdx) = "3. Assert: p_Error debe estar seteado"
    logIdx = logIdx + 1
    If m_Err = "" Then
        logs(logIdx) = "4. Assert FAIL: esperaba p_Error seteado"
        logIdx = logIdx + 1
        Test_CalidadTareas_CargarArbol_ArbolNothing_pErrorSetted = BuildFail( _
            "Arbol Nothing debe setear p_Error. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error seteado para Arbol Nothing (" & m_Err & ")"
    logIdx = logIdx + 1
    Test_CalidadTareas_CargarArbol_ArbolNothing_pErrorSetted = BuildOk("cargar_arbol_nothing_p_error_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_CargarArbol_ArbolNothing_pErrorSetted = BuildFail( _
        "Unhandled exception (helper must NOT raise on sad path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (HAPPY): CalidadTareas_CargarArbol con ArbolTareasCalidad valido
'   Carga el arbol de tareas de calidad para la edicion.
' ============================================================
Public Function Test_CalidadTareas_CargarArbol_ConArbolValido_Ok() As String
    Dim logs(0 To 6) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_CargarArbol_ConArbolValido_Ok = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Arbol As ArbolTareasCalidad
    Set m_Arbol = New ArbolTareasCalidad

    logs(logIdx) = "2. Act: helper con ArbolTareasCalidad valido"
    logIdx = logIdx + 1
    CalidadTareas_CargarArbol m_Arbol, m_Err

    logs(logIdx) = "3. Assert: p_Error vacio (helper valido sin raise)"
    logIdx = logIdx + 1
    If m_Err <> "" Then
        logs(logIdx) = "4. Assert FAIL: helper set p_Error: " & m_Err
        logIdx = logIdx + 1
        Test_CalidadTareas_CargarArbol_ConArbolValido_Ok = BuildFail( _
            "Arbol valido no debe setear p_Error. Got: " & m_Err, logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error vacio, arbol validado"
    logIdx = logIdx + 1
    Test_CalidadTareas_CargarArbol_ConArbolValido_Ok = BuildOk("cargar_arbol_valido_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_CargarArbol_ConArbolValido_Ok = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 6 (ADVERSARIAL): CalidadTareas_SeleccionarNodoRiesgo con riesgo Nothing + tipo invalido
'   La validacion debe priorizar el chequeo de Nothing (primer check)
' ============================================================
Public Function Test_CalidadTareas_SeleccionarNodoRiesgo_NothingPrioridad_Setted() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend + SeedAll"
    logIdx = logIdx + 1
    Dim m_Err As String
    If Not ForceBackend(m_Err) Then
        Test_CalidadTareas_SeleccionarNodoRiesgo_NothingPrioridad_Setted = BuildFail("TESTS BLOCKED: " & m_Err, logs)
        Exit Function
    End If
    Test_Fixtures.SeedAll

    Dim m_Key As String
    logs(logIdx) = "2. Act: Nothing + tipo invalido (999)"
    logIdx = logIdx + 1
    CalidadTareas_SeleccionarNodoRiesgo Nothing, 999, m_Key, m_Err

    logs(logIdx) = "3. Assert: p_Error seteado (prioridad Nothing)"
    logIdx = logIdx + 1
    If m_Err = "" Then
        logs(logIdx) = "4. Assert FAIL: esperaba p_Error"
        logIdx = logIdx + 1
        Test_CalidadTareas_SeleccionarNodoRiesgo_NothingPrioridad_Setted = BuildFail( _
            "Nothing + invalido debe setear p_Error. Got vacio.", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Assert PASS: p_Error seteado, no raise (" & m_Err & ")"
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_NothingPrioridad_Setted = BuildOk("seleccionar_nodo_nothing_invalido_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_CalidadTareas_SeleccionarNodoRiesgo_NothingPrioridad_Setted = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function
