Attribute VB_Name = "Test_ArbolRefrescoHelper"
' ============================================================
' Test_ArbolRefrescoHelper — TDD atoms for
'   modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   forms-thin-phase0-testeable-2026-06-25 / Tier-1 slice
'        desduplicacion gemelos bridge (Mitigacion, Materializado,
'        Retirado)
'
' Scope: 1 helper (Sub with ByRef p_Error) -> 7 scenario atoms
'   Helper signature (public, modArbolRefrescoHelper):
'     Public Sub Refresco_ArbolRiesgosCalidadBridge( _
'         ByVal p_Mutacion As String, _
'         Optional ByRef p_Error As String)
'
'   Convention: helper is Sub (no JSON return). Atoms wrap the call and
'   return JSON via TestCore_BuildOk / TestCore_BuildFail.
'
'   Hard rule 7: el helper NO hace MsgBox. Los atomos verifican que
'   p_Error se rellena en sad path sin raise visible.
'
'   Pre-condicion probada: FormRiesgosGestion NO esta abierto en tests
'   headless. El bridge hace no-op silencioso (p_Error = "") en ese caso.
'   Esto preserva la semantica original del bridge privado.
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOM 1 (HAPPY): scope "materializacion" — usado por Form_FormRiesgoMaterializado
'
' Setup: FormRiesgosGestion NO esta abierto (tests headless).
' Act:   modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "materializacion", m_Error
' Expect: m_Error = "" (no-op silencioso, no raise).
' ============================================================
Public Function Test_Refresco_ArbolRiesgosCalidad_Materializacion_Happy_NoFormOpen() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge ""materializacion"""
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "materializacion", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_ArbolRiesgosCalidad_Materializacion_Happy_NoFormOpen = TestCore_BuildFail( _
            "Materializacion happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Materializacion_Happy_NoFormOpen = TestCore_BuildOk( _
        "materializacion_happy_no_form_open", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Materializacion_Happy_NoFormOpen = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (HAPPY): scope "aceptacion" — usado por Form_FormRiesgoMitigacion
'                  (ComandoAceptacionRegistrar_Click, ListaMitigacion_Click)
' ============================================================
Public Function Test_Refresco_ArbolRiesgosCalidad_Aceptacion_Happy_NoFormOpen() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge ""aceptacion"""
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "aceptacion", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_ArbolRiesgosCalidad_Aceptacion_Happy_NoFormOpen = TestCore_BuildFail( _
            "Aceptacion happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Aceptacion_Happy_NoFormOpen = TestCore_BuildOk( _
        "aceptacion_happy_no_form_open", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Aceptacion_Happy_NoFormOpen = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (HAPPY): scope "calidad" — usado por Form_FormRiesgoMitigacion
'                  (ComandoAceptar_Click, ComandoRechazar_Click) y
'                  Form_FormRiesgoRetirado (ComandoAceptar_Click,
'                  ComandoRechazar_Click)
' ============================================================
Public Function Test_Refresco_ArbolRiesgosCalidad_Calidad_Happy_NoFormOpen() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge ""calidad"""
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "calidad", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_ArbolRiesgosCalidad_Calidad_Happy_NoFormOpen = TestCore_BuildFail( _
            "Calidad happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Calidad_Happy_NoFormOpen = TestCore_BuildOk( _
        "calidad_happy_no_form_open", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Calidad_Happy_NoFormOpen = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (HAPPY): scope "retiro" — usado por Form_FormRiesgoRetirado
'                  (ComandoRetiroRegistrar_Click, ComandoEliminar_Click)
' ============================================================
Public Function Test_Refresco_ArbolRiesgosCalidad_Retiro_Happy_NoFormOpen() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge ""retiro"""
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "retiro", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_ArbolRiesgosCalidad_Retiro_Happy_NoFormOpen = TestCore_BuildFail( _
            "Retiro happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Retiro_Happy_NoFormOpen = TestCore_BuildOk( _
        "retiro_happy_no_form_open", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_ArbolRiesgosCalidad_Retiro_Happy_NoFormOpen = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (EDGE): mutacion vacia — p_Error vacio, no raise
'
' Edge: el caller podria pasar una mutacion vacia por error.
' El helper debe tratarlo como no-op silencioso (la rama
' FormularioAbierto=False se ejecuta primero y sale).
' ============================================================
Public Function Test_Refresco_Helper_EmptyMutation_NoFormOpen_StillSucceeds() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge """" (empty mutation)"
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: helper set p_Error on empty mutation: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_Helper_EmptyMutation_NoFormOpen_StillSucceeds = TestCore_BuildFail( _
            "Empty mutation with no form open must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio, no raise on empty mutation"
    logIdx = logIdx + 1
    Test_Refresco_Helper_EmptyMutation_NoFormOpen_StillSucceeds = TestCore_BuildOk( _
        "empty_mutation_no_form_open_silent_no_op", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_Helper_EmptyMutation_NoFormOpen_StillSucceeds = TestCore_BuildFail( _
        "Helper must NOT raise on edge cases: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 6 (EDGE/GUARD): mutacion invalida con FormRiesgosGestion NO abierto
'
' Contrato del bridge: `If Not FormularioAbierto("FormRiesgosGestion")
' Then Exit Sub`. Esta guarda significa que cuando FormRiesgosGestion
' no esta abierto (modo headless de tests), TODAS las mutaciones
' (validas o invalidas) son silent no-op. La validacion de scope
' (ResolverScopeRefrescoCalidadRiesgo) solo se ejecuta cuando el form
' esta abierto.
'
' Hard rule 7: el helper NO hace MsgBox. Esta atomo verifica:
'   - No raise (sad path NO propaga Err.Raise visible).
'   - p_Error vacio (early-exit guard funciona correctamente).
'   - No hay MsgBox (helper 100% test-friendly).
' ============================================================
Public Function Test_Refresco_Helper_InvalidMutation_NoFormOpen_SilentNoOp() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "1. Act: Refresco_ArbolRiesgosCalidadBridge ""NoSoyUnaMutacionValida"" (form NOT open)"
    logIdx = logIdx + 1

    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "NoSoyUnaMutacionValida", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL: bridge populated p_Error despite early-exit guard: " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_Helper_InvalidMutation_NoFormOpen_SilentNoOp = TestCore_BuildFail( _
            "Early-exit guard should yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error vacio (early-exit guard funciona), no raise, no MsgBox"
    logIdx = logIdx + 1
    Test_Refresco_Helper_InvalidMutation_NoFormOpen_SilentNoOp = TestCore_BuildOk( _
        "invalid_mutation_early_exit_no_msgbox", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_Helper_InvalidMutation_NoFormOpen_SilentNoOp = TestCore_BuildFail( _
        "Bridge must NOT raise on guard early-exit. Unexpected raise: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 7 (ADVERSARIAL): doble llamada con scopes distintos — sin state bleed
'
' Llama al helper dos veces seguidas con scopes diferentes. Cada
' llamada debe producir p_Error="" sin contaminar el estado entre
' invocaciones (el bridge no mantiene estado entre llamadas — es Sub
' puro, solo modifica p_Error localmente).
' ============================================================
Public Function Test_Refresco_Helper_AdversarialDoble_OK() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim m_Error As String

    logs(logIdx) = "1. Act (call 1): Refresco_ArbolRiesgosCalidadBridge ""materializacion"""
    logIdx = logIdx + 1
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "materializacion", m_Error
    If m_Error <> "" Then
        logs(logIdx) = "2. Assert FAIL (call 1): " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_Helper_AdversarialDoble_OK = TestCore_BuildFail( _
            "Call 1 failed: " & m_Error, logs)
        Exit Function
    End If
    logs(logIdx) = "2. Assert PASS (call 1): p_Error vacio"
    logIdx = logIdx + 1

    logs(logIdx) = "3. Act (call 2): Refresco_ArbolRiesgosCalidadBridge ""retiro"" con m_Error limpio"
    logIdx = logIdx + 1
    m_Error = ""
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "retiro", m_Error
    If m_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL (call 2): " & m_Error
        logIdx = logIdx + 1
        Test_Refresco_Helper_AdversarialDoble_OK = TestCore_BuildFail( _
            "Call 2 failed (state bleed?): " & m_Error, logs)
        Exit Function
    End If
    logs(logIdx) = "4. Assert PASS (call 2): p_Error vacio, scopes distintos, idempotente"
    logIdx = logIdx + 1

    Test_Refresco_Helper_AdversarialDoble_OK = TestCore_BuildOk("adversarial_doble_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_Helper_AdversarialDoble_OK = TestCore_BuildFail( _
        "Unhandled exception in adversarial double-call: " & Err.Description, logs)
End Function