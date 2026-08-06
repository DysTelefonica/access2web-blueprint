Attribute VB_Name = "modFrmNCAuditoriaControlEficaciaHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormNCAuditoriaControlEficacia (slice 9 of form-thin-helper-refactor).
'
' This helper owns the 1 message-class decision that previously lived inline
' in Form_FormNCAuditoriaControlEficacia.cls event handlers:
'   1. _RenderError  -- boilerplate error handler (3 errores: blocks = 4 MsgBox: 2 in
'                       ComandoGrabar_Click + 2 in Form_Load)
'
' The form is THIN: builds m_Error, calls helper, helper renders MsgBox (or
' returns the contract in atom/test mode).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; constructor is pure DAO)
'           access-vba-tdd §1.6 (canonical signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology Hard Rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Slice:    9 of 25 (ControlEficacia pair -- NCAuditoria side; companion is
'           modFrmNCProyectoControlEficaciaHelper.bas on NCProyecto side).
' =============================================================================

' ----- Message contracts (byte-for-byte match with Test_modFrmNCAuditoriaControlEficaciaHelper.bas) -----
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"

' ----- Sentinel constants -----
Private Const PROMPT_ATOM As Long = -1       ' atom/test mode (no real modal)
Private Const PROMPT_PRODUCTION As Long = 0   ' production mode (real modal)

' =============================================================================
' PRIVATE HELPERS (declared at top per vba-access §10.1 + access-vba-tdd §1.8)
' =============================================================================

Private Function BuildOk(ByRef p_Logs As Collection, ByVal p_Value As Variant) As String
    BuildOk = TestHelper.BuildJsonOk(p_Logs, p_Value)
End Function

Private Function BuildFail(ByRef p_Logs As Collection, ByVal p_Msg As String) As String
    BuildFail = TestHelper.BuildJsonFail(p_Msg, p_Logs)
End Function

' =============================================================================
' PUBLIC API -- Per-form thin pattern: form calls these; helper orchestrates
' =============================================================================

' -----------------------------------------------------------------------------
' RenderError: extract the boilerplate error handler from the form's `errores:`
'   blocks (mirrors modFrmNCProyectoControlEficaciaHelper_RenderError shape).
'
'   Production (p_PromptResult = PROMPT_PRODUCTION = 0):
'     - Shows MsgBox(m_ErrorText, vbCritical/vbExclamation, "Error"/"Advertencia")
'     - vbCritical if Err.Number <> 1000 (unexpected error)
'     - vbExclamation if Err.Number = 1000 (controlled warning)
'     - Returns the user's MsgBox result (typically vbOK).
'
'   Atom/test mode (p_PromptResult = PROMPT_ATOM = -1):
'     - Returns PROMPT_ATOM sentinel without showing a real modal.
'     - Atom asserts `ok = true` on the helper return.
'
'   Test simulating user (p_PromptResult = any other Long):
'     - Returns that exact value (so the atom can inject user response).
'
'   The form pattern that uses this helper:
'     errores:
'         DoCmd.Hourglass False
'         If Err.Number <> 1000 Then
'             m_Error = "Al <Sub> se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
'             CorreoAlAdministrador m_Error
'         End If
'         pregunta = modFrmNCAuditoriaControlEficaciaHelper_RenderError(m_Error, Err.Number, , m_HelperErr)
'
'   Note: The form still calls `CorreoAlAdministrador m_Error` for non-1000 errors
'   BEFORE delegating to the helper. The helper only owns the MsgBox decision.
' -----------------------------------------------------------------------------
Public Function modFrmNCAuditoriaControlEficaciaHelper_RenderError( _
                                                ByVal p_ErrorText As String, _
                                                ByVal p_ErrNumber As Long, _
                                                Optional ByRef p_PromptResult As Long = 0, _
                                                Optional ByRef p_Error As String = "" _
                                                ) As Long
    On Error GoTo EH
    p_Error = ""

    ' Edge case: empty error text. No modal needed.
    If Len(Trim$(p_ErrorText)) = 0 Then
        modFrmNCAuditoriaControlEficaciaHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Atom/test mode: no real MsgBox.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCAuditoriaControlEficaciaHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Production: show real MsgBox.
    If p_PromptResult = PROMPT_PRODUCTION Then
        Dim style As Long
        If p_ErrNumber = 1000 Then
            style = vbExclamation
        Else
            style = vbCritical
        End If
        modFrmNCAuditoriaControlEficaciaHelper_RenderError = MsgBox(p_ErrorText, style, MSG_ERROR_TITLE)
        Exit Function
    End If

    ' Test simulating user response.
    modFrmNCAuditoriaControlEficaciaHelper_RenderError = p_PromptResult
    Exit Function

EH:
    p_Error = "modFrmNCAuditoriaControlEficaciaHelper_RenderError: " & Err.Description
    modFrmNCAuditoriaControlEficaciaHelper_RenderError = vbOK
End Function