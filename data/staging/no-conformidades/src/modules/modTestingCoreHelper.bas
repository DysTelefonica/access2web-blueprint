Attribute VB_Name = "modTestingCoreHelper"
Option Compare Database
Option Explicit

' ============================================
' modTestingCoreHelper
' ============================================
' Test infrastructure helpers. Pure logic - no production UI calls.
' Used by Test_*.bas modules to assert canonical JSON contract and
' p_PromptResult injection semantics.
'
' Parses JSON via JSONHelper.JSONAObjeto (JsonConverter-backed on
' Access 365). All assertions log to p_Logs (Collection) and return
' Boolean for caller convenience.
'
' SDD: openspec/changes/form-thin-helper-refactor/ (slice 0)
' Skill: access-vba-e2e-methodology rule #10
'        vba-access §1.4.1 (p_Error ByRef where applicable)
' ============================================

' === Module-level declarations (ALL AT THE TOP per vba-access §10.1) ===
' (no Private Const / Type / Enum needed for this module)

' === Private helpers (BEFORE public per vba-access §10.1) ===
' (logic inlined per public function; no shared helpers needed)

' ============================================
' PUBLIC FUNCTIONS (per-module prefix per access-vba-e2e-methodology rule #9)
' ============================================

Public Function modTestingCoreHelper_AssertJsonOk( _
    ByVal p_JsonString As String, _
    ByRef p_Logs As Collection _
) As Boolean
    On Error GoTo EH

    modTestingCoreHelper_AssertJsonOk = False

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(p_JsonString, jsonErr)

    If jsonErr <> "" Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON parse error: " & jsonErr
        Exit Function
    End If
    If dict Is Nothing Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON parser returned Nothing"
        Exit Function
    End If
    If Not dict.Exists("ok") Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON missing 'ok' field"
        Exit Function
    End If
    If dict("ok") <> True Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: expected ok=true, got ok=" & dict("ok")
        Exit Function
    End If

    TestHelper.AddLog p_Logs, "ASSERT OK: JSON ok=true"
    modTestingCoreHelper_AssertJsonOk = True
    Exit Function

EH:
    TestHelper.AddLog p_Logs, "ASSERT FAIL: AssertJsonOk error: " & Err.Description
End Function

Public Function modTestingCoreHelper_AssertJsonFail( _
    ByVal p_JsonString As String, _
    ByVal p_ExpectedErrorSubstring As String, _
    ByRef p_Logs As Collection _
) As Boolean
    On Error GoTo EH

    modTestingCoreHelper_AssertJsonFail = False

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(p_JsonString, jsonErr)

    If jsonErr <> "" Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON parse error: " & jsonErr
        Exit Function
    End If
    If dict Is Nothing Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON parser returned Nothing"
        Exit Function
    End If
    If Not dict.Exists("ok") Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON missing 'ok' field"
        Exit Function
    End If
    If dict("ok") <> False Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: expected ok=false, got ok=" & dict("ok")
        Exit Function
    End If
    If Not dict.Exists("error") Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: JSON missing 'error' field"
        Exit Function
    End If

    Dim actualError As String
    actualError = CStr(dict("error"))
    ' Case-sensitive match (per slice 0 spec). Use vbBinaryCompare so the
    ' assertion is robust against locale-dependent case folding.
    If InStr(1, actualError, p_ExpectedErrorSubstring, vbBinaryCompare) = 0 Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: expected error contains '" & p_ExpectedErrorSubstring & "', got '" & actualError & "'"
        Exit Function
    End If

    TestHelper.AddLog p_Logs, "ASSERT OK: JSON ok=false and error contains substring"
    modTestingCoreHelper_AssertJsonFail = True
    Exit Function

EH:
    TestHelper.AddLog p_Logs, "ASSERT FAIL: AssertJsonFail error: " & Err.Description
End Function

Public Function modTestingCoreHelper_AssertPromptResultInjected( _
    ByVal p_PromptResult As Long, _
    ByVal p_Expected As Long, _
    ByRef p_Logs As Collection _
) As Boolean
    On Error GoTo EH

    modTestingCoreHelper_AssertPromptResultInjected = False

    If p_PromptResult <> p_Expected Then
        TestHelper.AddLog p_Logs, "ASSERT FAIL: expected p_PromptResult=" & p_Expected & ", got " & p_PromptResult
        Exit Function
    End If

    TestHelper.AddLog p_Logs, "ASSERT OK: p_PromptResult=" & p_Expected
    modTestingCoreHelper_AssertPromptResultInjected = True
    Exit Function

EH:
    TestHelper.AddLog p_Logs, "ASSERT FAIL: AssertPromptResultInjected error: " & Err.Description
End Function
