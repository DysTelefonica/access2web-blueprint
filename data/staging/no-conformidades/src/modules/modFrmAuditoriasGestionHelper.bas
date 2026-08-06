Attribute VB_Name = "modFrmAuditoriasGestionHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormAuditoriasGestion (slice 4 of form-thin-helper-refactor).
'
' This helper owns the 2 message-class decisions that previously lived inline
' in Form_FormAuditoriasGestion.cls event handlers:
'   1. _RenderError                  -- boilerplate error handler (32 of 33 MsgBox,
'                                          all in the 16 errores: blocks)
'   2. _ConfirmarEliminacionAuditoria -- MsgBox("¿Desea eliminar...?") user confirmation
'                                          (1 MsgBox in ComandoEliminar_Click)
'
' The form is THIN: builds m_Error, calls helper, renders result.
' Operations class (AuditoriaOperaciones.cls) is pure DAO: no UI awareness.
'
' Skill:    access-vba-tdd §1.1 (operations class is pure DAO; helper owns UI)
'           access-vba-tdd §1.6 (canonical signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology Hard Rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Slice:    4 of 27. SINGLE form (FormAuditoriasGestion, the first single-form slice
'           of the epic; paired slices were 2a/2b/3).
' =============================================================================

' ----- Message contracts (byte-for-byte match with Test_modFrmAuditoriasGestionHelper.bas) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-AUDITORIAS-GESTION-ELIMINACION: ¿Desea eliminar definitivamente del sistema la auditoría seleccionada?"
Private Const MSG_ELIMINACION_TITLE As String = "Eliminación/Deseliminacion"
Private Const MSG_ERROR_TITLE As String = "Error"

' ----- Sentinel constants -----
Private Const PROMPT_ATOM As Long = -1       ' atom/test mode (no real modal)
Private Const PROMPT_PRODUCTION As Long = 0   ' production mode (real modal)
Private Const VB_YES As Long = 6
Private Const VB_NO As Long = 7

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
'   blocks. In production: shows MsgBox(vbCritical/vbExclamation, "Error"/"Advertencia")
'   based on Err.Number (1000 -> Exclamation, else Critical). In atom/test:
'   returns -1 sentinel, no modal.
'
'   p_PromptResult:
'     -1  = atom/test mode (no real MsgBox; return -1)
'      0  = production mode (show real MsgBox; return user's choice)
'     >=1 = test simulating user response (return that value)
'
'   Returns: Long (vbOK in production, -1 in atom, caller-injected value in test)
' -----------------------------------------------------------------------------
Public Function modFrmAuditoriasGestionHelper_RenderError( _
                                            ByVal p_ErrorText As String, _
                                            ByVal p_ErrNumber As Long, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As Long
    On Error GoTo EH
    p_Error = ""

    ' Edge case: empty error text. No modal needed.
    If Len(Trim$(p_ErrorText)) = 0 Then
        modFrmAuditoriasGestionHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Atom/test mode: no real MsgBox.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmAuditoriasGestionHelper_RenderError = PROMPT_ATOM
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
        modFrmAuditoriasGestionHelper_RenderError = MsgBox(p_ErrorText, style, MSG_ERROR_TITLE)
        Exit Function
    End If

    ' Test simulating user response.
    modFrmAuditoriasGestionHelper_RenderError = p_PromptResult
    Exit Function

EH:
    p_Error = "modFrmAuditoriasGestionHelper_RenderError: " & Err.Description
    modFrmAuditoriasGestionHelper_RenderError = vbOK
End Function

' -----------------------------------------------------------------------------
' ConfirmarEliminacionAuditoria: handle the delete-Auditoria confirmation flow.
'   This is the single user-confirmation MsgBox in ComandoEliminar_Click.
'   Production (p_PromptResult=0): shows MsgBox "¿Desea eliminar definitivamente
'     del sistema la auditoría seleccionada?". If user clicks Yes, delegates
'     to AuditoriaOperaciones.Eliminar. If user clicks No, returns "user_rejected".
'   Atom mode (p_PromptResult=-1): returns JSON with the eliminacion prompt text
'     contract, no real modal. Atom asserts the prompt text via JSON value.
'   Test simulating (p_PromptResult=vbYes/vbNo): uses that choice directly.
'
'   Notes:
'   - Unlike slice 2b's _ConfirmarBorradoNC, this helper has NO motivo prompt
'     (the original form's prompt is a simple yes/no for definitive deletion).
'   - AuditoriaOperaciones.Eliminar checks (1) NCs are Nothing, (2) no
'     documentos are open on disk. The helper propagates that error verbatim
'     via JSON ok=false / error=<opErr>.
'
'   Returns: JSON string with {ok, value, error, logs}.
' -----------------------------------------------------------------------------
Public Function modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria( _
                                            ByRef p_Auditoria As Auditoria, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim respuesta As Long
    Dim m_AudOp As AuditoriaOperaciones
    Dim opErr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_Auditoria is Nothing.
    If p_Auditoria Is Nothing Then
        modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildFail(logs, "ConfirmarEliminacionAuditoria: p_Auditoria is Nothing")
        Exit Function
    End If

    ' Atom mode: return prompt text contract. No real modal.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildOk(logs, MSG_ELIMINACION_PROMPT)
        Exit Function
    End If

    ' Production or test simulating user choice.
    If p_PromptResult = PROMPT_PRODUCTION Then
        ' Production: show real MsgBox.
        respuesta = MsgBox( _
            "¿Desea eliminar definitivamente del sistema la auditoría seleccionada?", _
            vbExclamation + vbYesNo + vbDefaultButton2, MSG_ELIMINACION_TITLE)
    Else
        respuesta = p_PromptResult
    End If

    If respuesta <> VB_YES Then
        modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildOk(logs, "user_rejected")
        Exit Function
    End If

    ' Delegate to operations class (pure DAO). Note: Eliminar operates on Me.Auditoria.
    Set m_AudOp = New AuditoriaOperaciones
    Set m_AudOp.Auditoria = p_Auditoria

    opErr = ""
    m_AudOp.Eliminar opErr
    If opErr <> "" Then
        p_Error = opErr
        modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildFail(logs, p_Error)
        Exit Function
    End If

    modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildOk(logs, "deleted")
    Exit Function

EH:
    p_Error = "modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria: " & Err.Description
    modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria = BuildFail(logs, p_Error)
End Function