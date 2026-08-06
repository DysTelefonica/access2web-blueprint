Attribute VB_Name = "modFormInteractionHelper"
Option Compare Database
Option Explicit

' ============================================
' modFormInteractionHelper
' ============================================
' Canonical MsgBox/InputBox wrapper utilities.
' All gated by p_PromptResult (ByRef Long, default -1).
'
' Contract (per access-vba-e2e-methodology rule #10):
'   - p_PromptResult = -1   : atom/test mode. NO real MsgBox/InputBox call.
'   - p_PromptResult >= 0   : use value as user response. NO real MsgBox.
' Production UI MsgBox rendering is the form's responsibility
' (it inspects the helper return and calls MsgBox itself when needed).
'
' SDD: openspec/changes/form-thin-helper-refactor/ (slice 0)
' Skill: access-vba-e2e-methodology rule #9, #10
'        vba-access §1.4.1 (p_Error ByRef)
' ============================================

' === Module-level declarations (ALL AT THE TOP per vba-access §10.1) ===
' (no Private Const / Type / Enum needed for this module)

' === Private helpers (BEFORE public per vba-access §10.1) ===
' (logic inlined per public function for clarity; no shared helpers needed)

' ============================================
' PUBLIC FUNCTIONS (per-module prefix per access-vba-e2e-methodology rule #9)
' ============================================

Public Function modFormInteractionHelper_AskUserYesNo( _
    ByVal p_Prompt As String, _
    ByVal p_Title As String, _
    Optional ByVal p_Default As Long = vbNo, _
    Optional ByRef p_PromptResult As Long = -1 _
) As Long
    On Error GoTo EH

    ' Atom/test mode (no injection): signal "would have prompted" without opening MsgBox.
    ' Returning -1 (not vbYes/vbNo) makes the sentinel explicit so callers can distinguish
    ' "no prompt" from "user said yes/no".
    If p_PromptResult = -1 Then
        modFormInteractionHelper_AskUserYesNo = -1
        Exit Function
    End If

    ' Injected (>= 0): use the value as the user response. Never call MsgBox.
    modFormInteractionHelper_AskUserYesNo = p_PromptResult
    Exit Function

EH:
    modFormInteractionHelper_AskUserYesNo = -1
End Function

Public Function modFormInteractionHelper_AskUserInput( _
    ByVal p_Prompt As String, _
    ByVal p_Title As String, _
    Optional ByVal p_Default As String = "", _
    Optional ByRef p_PromptResult As Long = -1, _
    Optional ByRef p_InputText As String = "" _
) As Long
    On Error GoTo EH

    ' Atom/test mode (no injection): populate p_InputText with p_Default so the
    ' caller can assert what the default value would have been. Return -1 as
    ' the "no real prompt" sentinel.
    If p_PromptResult = -1 Then
        p_InputText = p_Default
        modFormInteractionHelper_AskUserInput = -1
        Exit Function
    End If

    ' Injected (>= 0): OK vs Cancel.
    ' - vbOK: populate p_InputText with p_Default (the user accepted the default).
    ' - vbCancel: p_InputText = "" (the user dismissed without entering text).
    If p_PromptResult = vbOK Then
        p_InputText = p_Default
    Else
        p_InputText = ""
    End If
    modFormInteractionHelper_AskUserInput = p_PromptResult
    Exit Function

EH:
    p_InputText = ""
    modFormInteractionHelper_AskUserInput = -1
End Function

Public Sub modFormInteractionHelper_ShowInfo( _
    ByVal p_Message As String, _
    ByVal p_Title As String, _
    Optional ByRef p_PromptResult As Long = -1 _
)
    On Error GoTo EH

    ' Slice 0 ships the testable wrapper only. No real MsgBox is opened
    ' in atom mode (-1) or injection mode (>= 0); production forms render
    ' their own MsgBox using the helper's contract (or the helper's return).
    Exit Sub

EH:
    ' Swallow to keep the wrapper non-blocking. Production forms detect
    ' the empty return path and render their own error message.
End Sub

Public Sub modFormInteractionHelper_ShowError( _
    ByVal p_ErrorMessage As String, _
    ByVal p_Title As String, _
    Optional ByRef p_PromptResult As Long = -1, _
    Optional ByRef p_Error As String = "" _
)
    On Error GoTo EH

    ' Populate p_Error per Telefónica D&S convention (vba-access §1.4.1):
    ' every public Sub/Function ends with p_Error ByRef and surfaces the
    ' failure text there, not via MsgBox.
    p_Error = p_ErrorMessage
    Exit Sub

EH:
    p_Error = "modFormInteractionHelper_ShowError: " & Err.Description
End Sub
