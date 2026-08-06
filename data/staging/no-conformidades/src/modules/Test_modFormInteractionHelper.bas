Attribute VB_Name = "Test_modFormInteractionHelper"
Option Compare Database
Option Explicit

' ============================================
' TEST_MODFORMINTERACTIONHELPER
' ============================================
' Slice 0 atoms for the form-thin-helper-refactor epic.
' Covers 7 public functions across 2 helpers:
'   - modFormInteractionHelper: AskUserYesNo, AskUserInput, ShowInfo, ShowError
'   - modTestingCoreHelper: AssertJsonOk, AssertJsonFail, AssertPromptResultInjected
' 28 atoms (4 per public function: happy + sad + edge + adversarial).
' SDD: openspec/changes/form-thin-helper-refactor (slice 0)
' Skill: access-vba-tdd §1.1, §1.4, §1.8, §1.10
' ============================================

' === Module-level declarations (ALL AT THE TOP per vba-access §10.1) ===
Private Const FIX_PROMPT As String = "¿Confirma la operación? (test prompt)"
Private Const FIX_TITLE As String = "Título de prueba"
Private Const FIX_INPUT_DEFAULT As String = "valor por defecto"
Private Const FIX_INPUT_DEFAULT_UNICODE As String = "Año 2026: niño, ¿quién? ¡España!"
Private Const FIX_INFO As String = "Mensaje informativo de prueba"
Private Const FIX_ERROR As String = "MSG-01: error de prueba con acentos á é"
Private Const FIX_ERROR_LONG As String = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
Private Const FIX_TITLE_LONG As String = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
Private Const FIX_JSON_OK As String = "{""ok"":true,""error"":"""",""value"":""ok"",""logs"":[]}"
Private Const FIX_JSON_FAIL As String = "{""ok"":false,""error"":""fallo de prueba"",""value"":null,""logs"":[]}"
Private Const FIX_JSON_NO_OK As String = "{""value"":""missing ok field"",""error"":""""}"
Private Const FIX_JSON_DEEPLY_NESTED As String = "{""ok"":true,""error"":"""",""value"":{""a"":{""b"":{""c"":{""d"":{""e"":""deep""}}}}},""logs"":[]}"
Private Const FIX_JSON_MALFORMED As String = "{not valid json"
Private Const FIX_JSON_UNICODE_FAIL As String = "{""ok"":false,""error"":""MSG-Ñ: ¿fallo con acentos? ¡sí!"",""value"":null,""logs"":[]}"
Private Const FIX_SUBSTRING_PRESENT As String = "prueba"  ' REMOVED 2026-06-25: "MSG-01:" (invented; not a substring of FIX_JSON_FAIL's "error":"fallo de prueba"). Real substring: "prueba".
Private Const FIX_SUBSTRING_ABSENT As String = "esto no aparece en ningún lado"
Private Const FIX_VBYES As Long = 6
Private Const FIX_VBNO As Long = 7
Private Const FIX_VBOK As Long = 1
Private Const FIX_VBCANCEL As Long = 2

' === Private helpers (BEFORE public atoms per vba-access §10.1) ===
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs As Collection) As String
    BuildOk = TestHelper.BuildJsonOk(p_Logs, p_Value)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs As Collection) As String
    BuildFail = TestHelper.BuildJsonFail(p_Error, p_Logs)
End Function

' ============================================
' PART 1/2 — modFormInteractionHelper atoms (16)
' ============================================

' Happy: p_PromptResult=vbYes injected, helper returns vbYes without opening MsgBox
Public Function Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbYes_ReturnsVbYes_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = FIX_VBYES
    Dim result As Long
    result = modFormInteractionHelper_AskUserYesNo(FIX_PROMPT, FIX_TITLE, vbNo, promptResult)
    If result <> FIX_VBYES Then Err.Raise 1001, , "expected vbYes=" & FIX_VBYES & " got " & result
    Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbYes_ReturnsVbYes_Atomic = BuildOk("vbYes returned", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbYes_ReturnsVbYes_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: p_PromptResult=vbNo injected, helper returns vbNo (default-Yes path not taken)
Public Function Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbNo_ReturnsVbNo_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = FIX_VBNO
    Dim result As Long
    result = modFormInteractionHelper_AskUserYesNo(FIX_PROMPT, FIX_TITLE, vbYes, promptResult)
    If result <> FIX_VBNO Then Err.Raise 1001, , "expected vbNo=" & FIX_VBNO & " got " & result
    Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbNo_ReturnsVbNo_Atomic = BuildOk("vbNo returned", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserYesNo_WithPromptInjectionVbNo_ReturnsVbNo_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: no prompt injection (=-1), helper must NOT call MsgBox (would block COM) and returns -1
Public Function Test_modFormInteractionHelper_AskUserYesNo_WithoutPromptInjection_ReturnsMinusOneNoMsgBox_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = -1
    Dim result As Long
    result = modFormInteractionHelper_AskUserYesNo(FIX_PROMPT, FIX_TITLE, vbNo, promptResult)
    If result <> -1 Then Err.Raise 1001, , "expected -1 (skipped, no MsgBox), got " & result
    Test_modFormInteractionHelper_AskUserYesNo_WithoutPromptInjection_ReturnsMinusOneNoMsgBox_Atomic = BuildOk("no MsgBox", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserYesNo_WithoutPromptInjection_ReturnsMinusOneNoMsgBox_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: rapid double-click simulated by two sequential calls — the LAST call's injection wins
Public Function Test_modFormInteractionHelper_AskUserYesNo_Adversarial_DoubleClickRace_LastWins_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim pr1 As Long: pr1 = FIX_VBYES
    Dim pr2 As Long: pr2 = FIX_VBNO
    Dim r1 As Long: r1 = modFormInteractionHelper_AskUserYesNo(FIX_PROMPT, FIX_TITLE, vbNo, pr1)
    Dim r2 As Long: r2 = modFormInteractionHelper_AskUserYesNo(FIX_PROMPT, FIX_TITLE, vbYes, pr2)
    If r1 <> FIX_VBYES Then Err.Raise 1001, , "first call expected vbYes, got " & r1
    If r2 <> FIX_VBNO Then Err.Raise 1001, , "second (last) call expected vbNo, got " & r2
    Test_modFormInteractionHelper_AskUserYesNo_Adversarial_DoubleClickRace_LastWins_Atomic = BuildOk("last wins", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserYesNo_Adversarial_DoubleClickRace_LastWins_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Happy: p_PromptResult=vbOK injected, p_InputText populated with p_Default, returns vbOK
Public Function Test_modFormInteractionHelper_AskUserInput_WithPromptInjection_ReturnsDefaultText_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = FIX_VBOK
    Dim inputText As String
    Dim result As Long
    result = modFormInteractionHelper_AskUserInput(FIX_PROMPT, FIX_TITLE, FIX_INPUT_DEFAULT, promptResult, inputText)
    If result <> FIX_VBOK Then Err.Raise 1001, , "expected vbOK, got " & result
    If inputText <> FIX_INPUT_DEFAULT Then Err.Raise 1001, , "expected p_InputText='" & FIX_INPUT_DEFAULT & "', got '" & inputText & "'"
    Test_modFormInteractionHelper_AskUserInput_WithPromptInjection_ReturnsDefaultText_Atomic = BuildOk("default returned", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserInput_WithPromptInjection_ReturnsDefaultText_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: empty p_Default, p_PromptResult=vbOK, p_InputText populated with empty string
Public Function Test_modFormInteractionHelper_AskUserInput_WithEmptyDefault_PopulatesPInputText_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = FIX_VBOK
    Dim inputText As String
    Dim result As Long
    result = modFormInteractionHelper_AskUserInput(FIX_PROMPT, FIX_TITLE, "", promptResult, inputText)
    If result <> FIX_VBOK Then Err.Raise 1001, , "expected vbOK, got " & result
    If inputText <> "" Then Err.Raise 1001, , "expected p_InputText='', got '" & inputText & "'"
    Test_modFormInteractionHelper_AskUserInput_WithEmptyDefault_PopulatesPInputText_Atomic = BuildOk("empty default", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserInput_WithEmptyDefault_PopulatesPInputText_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: no prompt injection (=-1), helper returns -1 and populates p_InputText with p_Default
Public Function Test_modFormInteractionHelper_AskUserInput_WithoutPromptInjection_ReturnsMinusOne_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = -1
    Dim inputText As String
    Dim result As Long
    result = modFormInteractionHelper_AskUserInput(FIX_PROMPT, FIX_TITLE, FIX_INPUT_DEFAULT, promptResult, inputText)
    If result <> -1 Then Err.Raise 1001, , "expected -1 (skipped), got " & result
    If inputText <> FIX_INPUT_DEFAULT Then Err.Raise 1001, , "expected p_InputText='" & FIX_INPUT_DEFAULT & "', got '" & inputText & "'"
    Test_modFormInteractionHelper_AskUserInput_WithoutPromptInjection_ReturnsMinusOne_Atomic = BuildOk("skipped", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserInput_WithoutPromptInjection_ReturnsMinusOne_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: special chars (Ñ, ¿, ¡, á, é, í, ó, ú) preserved through p_Default/p_InputText
Public Function Test_modFormInteractionHelper_AskUserInput_Adversarial_SpecialChars_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = FIX_VBOK
    Dim inputText As String
    Dim result As Long
    result = modFormInteractionHelper_AskUserInput(FIX_PROMPT, FIX_TITLE, FIX_INPUT_DEFAULT_UNICODE, promptResult, inputText)
    If result <> FIX_VBOK Then Err.Raise 1001, , "expected vbOK, got " & result
    If inputText <> FIX_INPUT_DEFAULT_UNICODE Then Err.Raise 1001, , "unicode lost: expected '" & FIX_INPUT_DEFAULT_UNICODE & "', got '" & inputText & "'"
    Test_modFormInteractionHelper_AskUserInput_Adversarial_SpecialChars_Atomic = BuildOk("unicode preserved", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_AskUserInput_Adversarial_SpecialChars_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Happy: p_PromptResult=0 injected, no MsgBox opened, no error raised
Public Function Test_modFormInteractionHelper_ShowInfo_WithPromptInjection_NoMsgBoxOpened_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = 0
    modFormInteractionHelper_ShowInfo FIX_INFO, FIX_TITLE, promptResult
    Test_modFormInteractionHelper_ShowInfo_WithPromptInjection_NoMsgBoxOpened_Atomic = BuildOk("no MsgBox", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowInfo_WithPromptInjection_NoMsgBoxOpened_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: no prompt injection (=-1), helper must NOT call MsgBox (would block COM)
Public Function Test_modFormInteractionHelper_ShowInfo_WithoutPromptInjection_LogsButDoesNotBlock_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = -1
    modFormInteractionHelper_ShowInfo FIX_INFO, FIX_TITLE, promptResult
    Test_modFormInteractionHelper_ShowInfo_WithoutPromptInjection_LogsButDoesNotBlock_Atomic = BuildOk("no block", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowInfo_WithoutPromptInjection_LogsButDoesNotBlock_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: empty message, helper handles gracefully (no error)
Public Function Test_modFormInteractionHelper_ShowInfo_Edge_EmptyMessage_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = 0
    modFormInteractionHelper_ShowInfo "", FIX_TITLE, promptResult
    Test_modFormInteractionHelper_ShowInfo_Edge_EmptyMessage_Atomic = BuildOk("empty ok", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowInfo_Edge_EmptyMessage_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: very long title (200+ chars) handled without error
Public Function Test_modFormInteractionHelper_ShowInfo_Adversarial_LongTitle_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = 0
    modFormInteractionHelper_ShowInfo FIX_INFO, FIX_TITLE_LONG, promptResult
    Test_modFormInteractionHelper_ShowInfo_Adversarial_LongTitle_Atomic = BuildOk("long title ok", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowInfo_Adversarial_LongTitle_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Happy: error message propagated to p_Error (Telefónica D&S convention p_Error ByRef)
Public Function Test_modFormInteractionHelper_ShowError_WithErrorMessage_PopulatesPError_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long
    Dim pError As String
    modFormInteractionHelper_ShowError FIX_ERROR, FIX_TITLE, promptResult, pError
    If pError = "" Then Err.Raise 1001, , "expected p_Error populated, got empty"
    If InStr(pError, "MSG-01") = 0 Then Err.Raise 1001, , "expected p_Error to contain 'MSG-01', got '" & pError & "'"
    Test_modFormInteractionHelper_ShowError_WithErrorMessage_PopulatesPError_Atomic = BuildOk("pError populated", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowError_WithErrorMessage_PopulatesPError_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: empty message, p_Error populated with empty string (no error raised)
Public Function Test_modFormInteractionHelper_ShowError_EmptyMessage_DoesNothing_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long
    Dim pError As String
    modFormInteractionHelper_ShowError "", FIX_TITLE, promptResult, pError
    If pError <> "" Then Err.Raise 1001, , "expected p_Error='' for empty input, got '" & pError & "'"
    Test_modFormInteractionHelper_ShowError_EmptyMessage_DoesNothing_Atomic = BuildOk("empty handled", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowError_EmptyMessage_DoesNothing_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: no prompt injection (=-1), helper must NOT call MsgBox (would block COM)
Public Function Test_modFormInteractionHelper_ShowError_WithoutPromptInjection_DoesNotBlock_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long: promptResult = -1
    Dim pError As String
    modFormInteractionHelper_ShowError FIX_ERROR, FIX_TITLE, promptResult, pError
    If pError = "" Then Err.Raise 1001, , "expected p_Error populated even without prompt injection"
    Test_modFormInteractionHelper_ShowError_WithoutPromptInjection_DoesNotBlock_Atomic = BuildOk("no block", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowError_WithoutPromptInjection_DoesNotBlock_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: very long error message (200+ chars) handled without truncation
Public Function Test_modFormInteractionHelper_ShowError_Adversarial_VeryLongMessage_TruncatesSafely_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim promptResult As Long
    Dim pError As String
    modFormInteractionHelper_ShowError FIX_ERROR_LONG, FIX_TITLE, promptResult, pError
    If Len(pError) <> Len(FIX_ERROR_LONG) Then Err.Raise 1001, , "expected p_Error length=" & Len(FIX_ERROR_LONG) & ", got " & Len(pError)
    Test_modFormInteractionHelper_ShowError_Adversarial_VeryLongMessage_TruncatesSafely_Atomic = BuildOk("long ok", logs)
    Exit Function
EH:
    Test_modFormInteractionHelper_ShowError_Adversarial_VeryLongMessage_TruncatesSafely_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' PART 2/2 — modTestingCoreHelper atoms (12)

' Happy: valid JSON with ok=true passes assertion
Public Function Test_modTestingCoreHelper_AssertJsonOk_HappyPath_Passes_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonOk(FIX_JSON_OK, logs)
    If Not ok Then Err.Raise 1001, , "expected assertion to pass for valid ok=true JSON"
    Test_modTestingCoreHelper_AssertJsonOk_HappyPath_Passes_Atomic = BuildOk("assert passed", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonOk_HappyPath_Passes_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: malformed JSON fails assertion
Public Function Test_modTestingCoreHelper_AssertJsonOk_MalformedJson_Fails_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonOk(FIX_JSON_MALFORMED, logs)
    If ok Then Err.Raise 1001, , "expected assertion to fail for malformed JSON"
    Test_modTestingCoreHelper_AssertJsonOk_MalformedJson_Fails_Atomic = BuildOk("assert rejected", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonOk_MalformedJson_Fails_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: JSON missing 'ok' field fails assertion
Public Function Test_modTestingCoreHelper_AssertJsonOk_MissingOkField_Fails_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonOk(FIX_JSON_NO_OK, logs)
    If ok Then Err.Raise 1001, , "expected assertion to fail for JSON missing 'ok' field"
    Test_modTestingCoreHelper_AssertJsonOk_MissingOkField_Fails_Atomic = BuildOk("missing field rejected", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonOk_MissingOkField_Fails_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: deeply nested JSON with ok=true passes (parser handles nesting)
Public Function Test_modTestingCoreHelper_AssertJsonOk_Adversarial_DeeplyNested_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonOk(FIX_JSON_DEEPLY_NESTED, logs)
    If Not ok Then Err.Raise 1001, , "expected assertion to pass for deeply nested valid JSON"
    Test_modTestingCoreHelper_AssertJsonOk_Adversarial_DeeplyNested_Atomic = BuildOk("deep ok", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonOk_Adversarial_DeeplyNested_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Happy: ok=false + error substring matches passes
Public Function Test_modTestingCoreHelper_AssertJsonFail_HappyPath_ErrorSubstringMatches_Passes_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonFail(FIX_JSON_FAIL, FIX_SUBSTRING_PRESENT, logs)
    If Not ok Then Err.Raise 1001, , "expected assertion to pass when error substring matches"
    Test_modTestingCoreHelper_AssertJsonFail_HappyPath_ErrorSubstringMatches_Passes_Atomic = BuildOk("substring matched", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonFail_HappyPath_ErrorSubstringMatches_Passes_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: ok=true (not a fail) fails assertion
Public Function Test_modTestingCoreHelper_AssertJsonFail_SubstringNotPresent_Fails_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonFail(FIX_JSON_OK, FIX_SUBSTRING_ABSENT, logs)
    If ok Then Err.Raise 1001, , "expected assertion to fail when JSON has ok=true"
    Test_modTestingCoreHelper_AssertJsonFail_SubstringNotPresent_Fails_Atomic = BuildOk("not-fail rejected", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonFail_SubstringNotPresent_Fails_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: case-sensitive match — uppercase substring does NOT match lowercase error
Public Function Test_modTestingCoreHelper_AssertJsonFail_CaseSensitive_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonFail(FIX_JSON_FAIL, "FALLO DE PRUEBA", logs)
    If ok Then Err.Raise 1001, , "expected assertion to fail when case does not match (case-sensitive)"
    Test_modTestingCoreHelper_AssertJsonFail_CaseSensitive_Atomic = BuildOk("case-sensitive enforced", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonFail_CaseSensitive_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: Spanish chars (Ñ, ¿, ¡, á, é) in error message preserved through assertion
Public Function Test_modTestingCoreHelper_AssertJsonFail_Adversarial_UnicodeInErrorMessage_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertJsonFail(FIX_JSON_UNICODE_FAIL, "MSG-Ñ", logs)
    If Not ok Then Err.Raise 1001, , "expected assertion to pass for unicode-bearing error"
    Test_modTestingCoreHelper_AssertJsonFail_Adversarial_UnicodeInErrorMessage_Atomic = BuildOk("unicode preserved", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertJsonFail_Adversarial_UnicodeInErrorMessage_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Happy: matching values pass
Public Function Test_modTestingCoreHelper_AssertPromptResultInjected_HappyPath_Matches_Passes_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertPromptResultInjected(FIX_VBYES, FIX_VBYES, logs)
    If Not ok Then Err.Raise 1001, , "expected assertion to pass when values match"
    Test_modTestingCoreHelper_AssertPromptResultInjected_HappyPath_Matches_Passes_Atomic = BuildOk("match ok", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertPromptResultInjected_HappyPath_Matches_Passes_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Sad: mismatched values fail
Public Function Test_modTestingCoreHelper_AssertPromptResultInjected_Mismatch_Fails_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertPromptResultInjected(FIX_VBYES, FIX_VBNO, logs)
    If ok Then Err.Raise 1001, , "expected assertion to fail when values differ"
    Test_modTestingCoreHelper_AssertPromptResultInjected_Mismatch_Fails_Atomic = BuildOk("mismatch rejected", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertPromptResultInjected_Mismatch_Fails_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Edge: -1 sentinel (test mode without injection) matches -1 expected
Public Function Test_modTestingCoreHelper_AssertPromptResultInjected_Edge_NegativeOne_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertPromptResultInjected(-1, -1, logs)
    If Not ok Then Err.Raise 1001, , "expected -1 to match -1"
    Test_modTestingCoreHelper_AssertPromptResultInjected_Edge_NegativeOne_Atomic = BuildOk("sentinel ok", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertPromptResultInjected_Edge_NegativeOne_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function

' Adversarial: vbYes vs vbNo in MsgBox button codes must not be conflated (6 vs 7)
Public Function Test_modTestingCoreHelper_AssertPromptResultInjected_Adversarial_VbYesVbNo_Atomic() As String
    On Error GoTo EH
    Dim logs As Collection: Set logs = New Collection
    Dim ok As Boolean
    ok = modTestingCoreHelper_AssertPromptResultInjected(FIX_VBYES, FIX_VBNO, logs)
    If ok Then Err.Raise 1001, , "vbYes must not equal vbNo"
    ok = modTestingCoreHelper_AssertPromptResultInjected(FIX_VBNO, FIX_VBYES, logs)
    If ok Then Err.Raise 1001, , "vbNo must not equal vbYes"
    Test_modTestingCoreHelper_AssertPromptResultInjected_Adversarial_VbYesVbNo_Atomic = BuildOk("buttons distinct", logs)
    Exit Function
EH:
    Test_modTestingCoreHelper_AssertPromptResultInjected_Adversarial_VbYesVbNo_Atomic = BuildFail("Test: " & Err.Description, logs)
End Function
