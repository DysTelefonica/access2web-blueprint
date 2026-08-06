Attribute VB_Name = "Test_modFrmNCProyectoControlEficaciaAltaHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmNCProyectoControlEficaciaAltaHelper (slice 10 of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; constructor is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper owns the 1 message-class decision that previously lived
'           inline in Form_FormNCProyectoControlEficaciaAlta.cls:
'             1. _RenderError  -- boilerplate error handler (2 errores: blocks, 4 MsgBox: 2 in
'                                ComandoGrabar_Click + 2 in Form_Load)
' Slice:    10 of 25 (ControlEficaciaAlta pair -- NCProyecto side).
'           4 atoms total (1 public function x 4 scenario classes: happy/sad/edge/adversarial).
'
' Source-of-truth provenance:
'   - Form_FormNCProyectoControlEficaciaAlta.cls -- 4 MsgBox-having errores: blocks confirmed
'       via grep; each block follows the same pattern:
'           If Err.Number <> 1000 Then ... CorreoAlAdministrador m_Error
'           pregunta = MsgBox(m_Error, vbCritical / vbExclamation, "Error" / "Advertencia")
'   - The helper's _RenderError contract:
'       - empty text -> PROMPT_ATOM (-1), no modal
'       - p_PromptResult=PROMPT_ATOM (-1) -> PROMPT_ATOM (-1), no modal
'       - p_PromptResult=PROMPT_PRODUCTION (0) -> real MsgBox (not exercised by atom)
'       - other Long -> return that value (test simulating user response)
' =============================================================================

' ----- Message contract (byte-for-byte match with production literals) -----
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"

' ----- Sentinel constants (mirror production module) -----
Private Const PROMPT_ATOM As Long = -1
Private Const PROMPT_PRODUCTION As Long = 0
Private Const VB_OK As Long = 1
Private Const VB_CANCEL As Long = 2

' =============================================================================
' Local helpers (Private, declared at top per access-vba-tdd §1.8)
' =============================================================================

Private Function BuildOk(ByRef p_Logs As Collection, ByVal p_Value As Variant) As String
    BuildOk = TestHelper.BuildJsonOk(p_Logs, p_Value)
End Function

Private Function BuildFail(ByRef p_Logs As Collection, ByVal p_Msg As String) As String
    BuildFail = TestHelper.BuildJsonFail(p_Msg, p_Logs)
End Function

Private Function LogsToJSON(ByRef p_Logs As Collection) As String
    Dim s As String
    Dim i As Long
    s = "["
    For i = 1 To p_Logs.Count
        If i > 1 Then s = s & ","
        s = s & """" & Replace(p_Logs(i), """", "\""") & """"
    Next i
    LogsToJSON = s & "]"
End Function

' =============================================================================
' ATOMS (Public Functions returning JSON contracts)
' =============================================================================

' -----------------------------------------------------------------------------
' Happy path: p_PromptResult=PROMPT_ATOM (-1) returns -1, no real MsgBox.
'   The helper must NOT show a modal in atom mode (it would block COM).
' -----------------------------------------------------------------------------
Public Function Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Happy_NoModal_Atomic() As String
    Dim logs As Collection
    Dim helperErr As String
    Dim result As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErr = ""

    TestHelper.AddLog logs, "Arrange: prepare error text + ErrNumber for atom-mode call"

    ' Act: call helper in atom mode (PROMPT_ATOM = -1)
    result = modFrmNCProyectoControlEficaciaAltaHelper_RenderError( _
                    p_ErrorText:="Test error text for atom mode", _
                    p_ErrNumber:=9999, _
                    p_PromptResult:=PROMPT_ATOM, _
                    p_Error:=helperErr)

    TestHelper.AddLog logs, "Act: helper returned result=" & result & ", helperErr=""" & helperErr & """"

    ' Assert: PROMPT_ATOM (-1) sentinel -- no real modal was shown
    If result <> PROMPT_ATOM Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Happy_NoModal_Atomic = _
            BuildFail(logs, "expected result=PROMPT_ATOM (-1), got " & result)
        Exit Function
    End If
    If Len(helperErr) <> 0 Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Happy_NoModal_Atomic = _
            BuildFail(logs, "expected empty helperErr, got """ & helperErr & """")
        Exit Function
    End If

    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Happy_NoModal_Atomic = _
        BuildOk(logs, "atom_mode_no_modal")
    Exit Function

EH:
    TestHelper.AddLog logs, "EH: " & Err.Description
    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Happy_NoModal_Atomic = _
        BuildFail(logs, "EH: " & Err.Description)
End Function

' -----------------------------------------------------------------------------
' Sad path: ErrNumber=1000 (warning path) -- atom still no-modal.
'   The vbCritical/vbExclamation decision is gated on p_PromptResult=PROMPT_PRODUCTION.
'   In atom mode, the helper must short-circuit regardless of ErrNumber.
' -----------------------------------------------------------------------------
Public Function Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Sad_WarningPath_Atomic() As String
    Dim logs As Collection
    Dim helperErr As String
    Dim result As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErr = ""

    TestHelper.AddLog logs, "Arrange: prepare ErrNumber=1000 (controlled warning) + atom-mode call"

    ' Act: call helper with ErrNumber=1000 in atom mode
    result = modFrmNCProyectoControlEficaciaAltaHelper_RenderError( _
                    p_ErrorText:="Controlled warning message", _
                    p_ErrNumber:=1000, _
                    p_PromptResult:=PROMPT_ATOM, _
                    p_Error:=helperErr)

    TestHelper.AddLog logs, "Act: helper returned result=" & result & ", helperErr=""" & helperErr & """"

    ' Assert: PROMPT_ATOM (-1) sentinel still -- atom mode bypasses style decision
    If result <> PROMPT_ATOM Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Sad_WarningPath_Atomic = _
            BuildFail(logs, "expected result=PROMPT_ATOM (-1) even with ErrNumber=1000, got " & result)
        Exit Function
    End If
    If Len(helperErr) <> 0 Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Sad_WarningPath_Atomic = _
            BuildFail(logs, "expected empty helperErr, got """ & helperErr & """")
        Exit Function
    End If

    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Sad_WarningPath_Atomic = _
        BuildOk(logs, "atom_mode_warning_path")
    Exit Function

EH:
    TestHelper.AddLog logs, "EH: " & Err.Description
    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Sad_WarningPath_Atomic = _
        BuildFail(logs, "EH: " & Err.Description)
End Function

' -----------------------------------------------------------------------------
' Edge case: empty error text. Helper returns PROMPT_ATOM without showing any
'   modal. This guards the form from accidentally calling MsgBox("") if m_Error
'   is unexpectedly empty (e.g. Err.Raise 1000 without an error string set).
' -----------------------------------------------------------------------------
Public Function Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Edge_EmptyText_Atomic() As String
    Dim logs As Collection
    Dim helperErr As String
    Dim result As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErr = ""

    TestHelper.AddLog logs, "Arrange: prepare empty error text + ErrNumber=9999"

    ' Act: call helper with empty text in atom mode
    result = modFrmNCProyectoControlEficaciaAltaHelper_RenderError( _
                    p_ErrorText:="", _
                    p_ErrNumber:=9999, _
                    p_PromptResult:=PROMPT_ATOM, _
                    p_Error:=helperErr)

    TestHelper.AddLog logs, "Act: helper returned result=" & result & " for empty text"

    ' Assert: empty text returns PROMPT_ATOM (-1) immediately
    If result <> PROMPT_ATOM Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Edge_EmptyText_Atomic = _
            BuildFail(logs, "expected result=PROMPT_ATOM (-1) for empty text, got " & result)
        Exit Function
    End If

    ' Also test with whitespace-only text
    result = modFrmNCProyectoControlEficaciaAltaHelper_RenderError( _
                    p_ErrorText:="   ", _
                    p_ErrNumber:=9999, _
                    p_PromptResult:=PROMPT_ATOM, _
                    p_Error:=helperErr)
    If result <> PROMPT_ATOM Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Edge_EmptyText_Atomic = _
            BuildFail(logs, "expected result=PROMPT_ATOM (-1) for whitespace text, got " & result)
        Exit Function
    End If

    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Edge_EmptyText_Atomic = _
        BuildOk(logs, "atom_mode_empty_text")
    Exit Function

EH:
    TestHelper.AddLog logs, "EH: " & Err.Description
    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Edge_EmptyText_Atomic = _
        BuildFail(logs, "EH: " & Err.Description)
End Function

' -----------------------------------------------------------------------------
' Adversarial: 2000-char error text. Helper must handle arbitrarily long text
'   without buffer overflow, runtime error, or partial display. Atom mode
'   still returns PROMPT_ATOM.
' -----------------------------------------------------------------------------
Public Function Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Adversarial_VeryLongText_Atomic() As String
    Dim logs As Collection
    Dim helperErr As String
    Dim result As Long
    Dim longText As String
    Dim i As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErr = ""

    ' Build a 2000-char error text (2000 'X' characters)
    longText = ""
    For i = 1 To 2000
        longText = longText & "X"
    Next i
    TestHelper.AddLog logs, "Arrange: prepared longText with " & Len(longText) & " chars"

    ' Act: call helper with 2000-char text in atom mode
    result = modFrmNCProyectoControlEficaciaAltaHelper_RenderError( _
                    p_ErrorText:=longText, _
                    p_ErrNumber:=9999, _
                    p_PromptResult:=PROMPT_ATOM, _
                    p_Error:=helperErr)

    TestHelper.AddLog logs, "Act: helper returned result=" & result & " for " & Len(longText) & "-char text"

    ' Assert: PROMPT_ATOM (-1) sentinel -- long text doesn't break the atom-mode path
    If result <> PROMPT_ATOM Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            BuildFail(logs, "expected result=PROMPT_ATOM (-1) for 2000-char text, got " & result)
        Exit Function
    End If
    If Len(helperErr) <> 0 Then
        Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            BuildFail(logs, "expected empty helperErr, got """ & helperErr & """")
        Exit Function
    End If

    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        BuildOk(logs, "atom_mode_2000_chars")
    Exit Function

EH:
    TestHelper.AddLog logs, "EH: " & Err.Description
    Test_modFrmNCProyectoControlEficaciaAltaHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        BuildFail(logs, "EH: " & Err.Description)
End Function
