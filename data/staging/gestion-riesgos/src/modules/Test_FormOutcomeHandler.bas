Attribute VB_Name = "Test_FormOutcomeHandler"
' ============================================================
' Test_FormOutcomeHandler — TDD atoms for
'   modFormOutcomeHandler.FormOutcomeHandler_Run
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   forms-thin-phase0-testeable-2026-06-25 / Tier-1 slice
'        + generalize-visado-outcome-handler-2026-07-01
'
' Historia:
'   - Renombrado desde Test_VisadoOutcomeHandler al generalizarse el helper.
'     El nombre del modulo refleja ahora que cubre el contrato general del
'     handler, no solo los outcomes "visado" del gemelo Calidad.
'
' Scope: 1 helper (Sub with ByRef p_Error) -> 5 scenario atoms
'   Helper signature:
'     Public Sub FormOutcomeHandler_Run( _
'         ByVal p_Outcome As String, _
'         Optional ByRef p_Error As String)
'
'   Convention: helper is Sub (no JSON return). Atoms wrap the call and
'   return JSON via TestCore_BuildOk / TestCore_BuildFail.
'
'   Hard rule 7: el helper NO hace MsgBox. Los atomos verifican que
'   p_Error se rellena en sad path sin raise visible.
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOM 1 (HAPPY): Aprobada
'
' Setup: ForceLocalBackend (la cadena llama New TareasCalidad → requiere backend
'        alcanzable, aunque los formularios no esten abiertos).
' Act:   FormOutcomeHandler_Run("AceptacionAprobada", m_Error)
' Expect: m_Error = "" (helper no raise; bridge es no-op silencioso).
' ============================================================
Public Function Test_FormOutcomeHandler_Run_Aprobada_AcceptacionYes() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_FormOutcomeHandler_Run_Aprobada_AcceptacionYes = TestCore_BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "2. Act: FormOutcomeHandler_Run(""AceptacionAprobada"")"
    logIdx = logIdx + 1

    FormOutcomeHandler_Run "AceptacionAprobada", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_Aprobada_AcceptacionYes = TestCore_BuildFail( _
            "Happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_Aprobada_AcceptacionYes = TestCore_BuildOk("aprobada_acceptacion_yes", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_Aprobada_AcceptacionYes = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 2 (HAPPY): RechazadaQuitada
'
' Mismo flujo que Aprobada — los 4 outcomes llaman internamente
' EstablecerTareasCalidad(EnumSiNo.Sí, EnumSiNo.No) y el mismo bridge.
' Verificamos que el outcome alternativo funciona identicamente.
' ============================================================
Public Function Test_FormOutcomeHandler_Run_RechazadaQuitada_AcceptacionYes() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_FormOutcomeHandler_Run_RechazadaQuitada_AcceptacionYes = TestCore_BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim m_Error As String
    m_Error = "PRESET-WAS-NOT-CLEARED"

    logs(logIdx) = "2. Act: FormOutcomeHandler_Run(""AceptacionRechazadaQuitada"")"
    logIdx = logIdx + 1

    FormOutcomeHandler_Run "AceptacionRechazadaQuitada", m_Error

    If m_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: helper set p_Error: " & m_Error
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_RechazadaQuitada_AcceptacionYes = TestCore_BuildFail( _
            "Happy path must yield p_Error=''. Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error vacio, helper did not raise"
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_RechazadaQuitada_AcceptacionYes = TestCore_BuildOk("rechazada_quitada_acceptacion_yes", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_RechazadaQuitada_AcceptacionYes = TestCore_BuildFail( _
        "Unhandled exception (helper must NOT raise on happy path): " & Err.Description, logs)
End Function

' ============================================================
' ATOM 3 (SAD): Outcome vacio → p_Error populated, NO raise
'
' Hard rule 7 (no MsgBox) + sad path: si el outcome esta vacio o no es
' uno de los 4 soportados, el helper rellena p_Error SIN hacer Err.Raise
' visible (asi el caller puede inspeccionar y decidir).
' ============================================================
Public Function Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise = TestCore_BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim m_Error As String

    logs(logIdx) = "2. Act: FormOutcomeHandler_Run("""") — outcome vacio"
    logIdx = logIdx + 1

    FormOutcomeHandler_Run "", m_Error

    If m_Error = "" Then
        logs(logIdx) = "3. Assert FAIL: helper did not populate p_Error on empty outcome"
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise = TestCore_BuildFail( _
            "Sad path must yield p_Error <> ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "outcome", vbTextCompare) = 0 Then
        logs(logIdx) = "3. Assert FAIL: p_Error does not mention 'outcome': " & m_Error
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise = TestCore_BuildFail( _
            "Sad path p_Error should mention 'outcome' (validation rule). Got: " & m_Error, logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: p_Error populated, helper did not raise"
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise = TestCore_BuildOk("empty_outcome_sad", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_EmptyOutcome_PopulatesErrorNoRaise = TestCore_BuildFail( _
        "Sad path must NOT raise — caller inspects p_Error. Unexpected raise: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 4 (EDGE): outcome invalido (no en la lista de 4)
'
'   Verifica que cualquier outcome que NO este en la lista produce
'   p_Error sin raise. Cubre typos / valores futuros / entradas
'   maliciosas.
' ============================================================
Public Function Test_FormOutcomeHandler_Run_InvalidOutcome_PopulatesErrorNoRaise() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_FormOutcomeHandler_Run_InvalidOutcome_PopulatesErrorNoRaise = TestCore_BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim m_Error As String

    logs(logIdx) = "1. Act: FormOutcomeHandler_Run(""NoSoyUnOutcome"")"
    logIdx = logIdx + 1

    FormOutcomeHandler_Run "NoSoyUnOutcome", m_Error

    If m_Error = "" Then
        logs(logIdx) = "2. Assert FAIL: helper did not populate p_Error on invalid outcome"
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_InvalidOutcome_PopulatesErrorNoRaise = TestCore_BuildFail( _
            "Invalid outcome must yield p_Error <> ''", logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert PASS: p_Error populated, no raise"
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_InvalidOutcome_PopulatesErrorNoRaise = TestCore_BuildOk("invalid_outcome_sad", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_InvalidOutcome_PopulatesErrorNoRaise = TestCore_BuildFail( _
        "Sad path must NOT raise. Unexpected raise: " & Err.Description, logs)
End Function

' ============================================================
' ATOM 5 (ADVERSARIAL): doble llamada — idempotente y sin state bleed
'
'   Llama al helper dos veces seguidas con el mismo outcome. Cada
'   llamada debe producir p_Error="" sin contaminar el estado del
'   backend (m_ObjTareasCalidad) ni del caller (m_Error limpio al
'   inicio de cada llamada).
' ============================================================
Public Function Test_FormOutcomeHandler_Run_AdversarialDoble_OK() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_FormOutcomeHandler_Run_AdversarialDoble_OK = TestCore_BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim m_Error As String

    logs(logIdx) = "2. Act (call 1): FormOutcomeHandler_Run(""AceptacionAprobada"")"
    logIdx = logIdx + 1
    FormOutcomeHandler_Run "AceptacionAprobada", m_Error
    If m_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL (call 1): " & m_Error
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_AdversarialDoble_OK = TestCore_BuildFail( _
            "Call 1 failed: " & m_Error, logs)
        Exit Function
    End If
    logs(logIdx) = "3. Assert PASS (call 1): p_Error vacio"
    logIdx = logIdx + 1

    logs(logIdx) = "4. Act (call 2): FormOutcomeHandler_Run(""AceptacionAprobada"") con m_Error="""
    logIdx = logIdx + 1
    m_Error = ""
    FormOutcomeHandler_Run "AceptacionAprobada", m_Error
    If m_Error <> "" Then
        logs(logIdx) = "5. Assert FAIL (call 2): " & m_Error
        logIdx = logIdx + 1
        Test_FormOutcomeHandler_Run_AdversarialDoble_OK = TestCore_BuildFail( _
            "Call 2 failed (state bleed?): " & m_Error, logs)
        Exit Function
    End If
    logs(logIdx) = "5. Assert PASS (call 2): p_Error vacio, idempotente"
    logIdx = logIdx + 1

    Test_FormOutcomeHandler_Run_AdversarialDoble_OK = TestCore_BuildOk("adversarial_doble_ok", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_FormOutcomeHandler_Run_AdversarialDoble_OK = TestCore_BuildFail( _
        "Unhandled exception in adversarial double-call: " & Err.Description, logs)
End Function