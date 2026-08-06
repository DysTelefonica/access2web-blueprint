Attribute VB_Name = "modFrmNCProyectoARHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormNCProyectoAR (slice 5 of form-thin-helper-refactor).
'
' This helper owns the 3 message-class decisions that previously lived inline
' in Form_FormNCProyectoAR.cls event handlers:
'   1. _RenderError             -- boilerplate error handler (~14 of 16 MsgBox)
'   2. _ConfirmarBorradoNC      -- MsgBox("¿Desea borrar la NC?") + parent-NC Borrar flow
'   3. _ConfirmarHabilitacionNC -- MsgBox("¿Desea habilitar la NC?")
'
' The form is THIN: builds m_Error, calls helper, renders result.
' Operations class (NCProyectoOperaciones.cls) is pure DAO: no UI awareness.
'
' Skill:    access-vba-tdd §1.1 (operations class is pure DAO; helper owns UI)
'           access-vba-tdd §1.6 (canonical signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology Hard Rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Slice:    5 of 27 (AR pair: NCProyectoAR + NCAuditoriaAR;
'           mirror of slice 3 Acciones pair with name changes only).
' =============================================================================

' ----- Message contracts (byte-for-byte match with Test_modFrmNCProyectoARHelper.bas) -----
Private Const MSG_BORRADO_PROMPT As String = "MSG-AR-BORRADO: ¿Desea borrar la Acción Realizada seleccionada?"
Private Const MSG_HABILITACION_PROMPT As String = "MSG-AR-HABILITACION: ¿Desea habilitar la no conformidad de proyecto seleccionada?"
Private Const MSG_BORRADO_TITLE As String = "Borrado de NC"
Private Const MSG_HABILITACION_TITLE As String = "Habilitación de NC"
Private Const MSG_MOTIVO_TITLE As String = "Motivación"
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
Public Function modFrmNCProyectoARHelper_RenderError( _
                                            ByVal p_ErrorText As String, _
                                            ByVal p_ErrNumber As Long, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As Long
    On Error GoTo EH
    p_Error = ""

    ' Edge case: empty error text. No modal needed.
    If Len(Trim$(p_ErrorText)) = 0 Then
        modFrmNCProyectoARHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Atom/test mode: no real MsgBox.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCProyectoARHelper_RenderError = PROMPT_ATOM
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
        modFrmNCProyectoARHelper_RenderError = MsgBox(p_ErrorText, style, MSG_ERROR_TITLE)
        Exit Function
    End If

    ' Test simulating user response.
    modFrmNCProyectoARHelper_RenderError = p_PromptResult
    Exit Function

EH:
    p_Error = "modFrmNCProyectoARHelper_RenderError: " & Err.Description
    modFrmNCProyectoARHelper_RenderError = vbOK
End Function

' -----------------------------------------------------------------------------
' ConfirmarBorradoNC: handle the delete-NC confirmation flow with motivo prompt.
'   Production (p_PromptResult=0): shows MsgBox "¿Desea borrar la NC de proyecto?".
'     If user clicks Yes, shows InputBox for motivo. If motivo is empty, rejects.
'     If user clicks No or motivo empty, returns "user_rejected" or "motivo_required".
'     If user accepts, delegates to NCProyectoOperaciones.Eliminar.
'   Atom mode (p_PromptResult=-1): returns JSON with the borrado prompt text
'     contract, no real modal. Atom asserts the prompt text via JSON value.
'   Test simulating (p_PromptResult=vbYes/vbNo): uses that choice directly,
'     caller-provided motivo via p_Motivo.
'
'   Returns: JSON string with {ok, value, error, logs}.
'
'   NOTE: AR forms do NOT currently invoke _ConfirmarBorradoNC (the AR list
'   delete lives in FormNCProyectoAcciones, not in the single-AR edit form).
'   This helper is shipped for coverage parity with the slice 3 Acciones
'   pair and to provide the canonical 3-function shape per the SDD.
' -----------------------------------------------------------------------------
Public Function modFrmNCProyectoARHelper_ConfirmarBorradoNC( _
                                            ByRef p_NC As NCProyecto, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Motivo As String = "", _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim respuesta As Long
    Dim motivoFinal As String
    Dim m_NCOp As NCProyectoOperaciones
    Dim opErr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_NC is Nothing.
    If p_NC Is Nothing Then
        modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildFail(logs, "ConfirmarBorradoNC: p_NC is Nothing")
        Exit Function
    End If

    ' Atom mode: return prompt text contract. No real modal.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildOk(logs, MSG_BORRADO_PROMPT)
        Exit Function
    End If

    ' Production or test simulating user choice.
    If p_PromptResult = PROMPT_PRODUCTION Then
        ' Production: show real MsgBox.
        respuesta = MsgBox( _
            "¿Desea borrar la Acción Realizada seleccionada?", _
            vbExclamation + vbYesNo + vbDefaultButton2, MSG_BORRADO_TITLE)
        If respuesta <> VB_YES Then
            modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildOk(logs, "user_rejected")
            Exit Function
        End If
        ' Production: show real InputBox for motivo.
        motivoFinal = Nz(InputBox( _
            "Introduzca el motivo para marcar como borrada la no conformidad", _
            MSG_MOTIVO_TITLE), "")
    Else
        ' Test simulating user choice.
        respuesta = p_PromptResult
        If respuesta <> VB_YES Then
            modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildOk(logs, "user_rejected")
            Exit Function
        End If
        motivoFinal = p_Motivo
    End If

    ' Edge: motivo is empty. Reject before calling operations.
    If Len(Trim$(motivoFinal)) = 0 Then
        p_Error = "El motivo es obligatorio para marcar como borrada la no conformidad"
        modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' Update the caller's motivo buffer (in case production caller wants it back).
    p_Motivo = motivoFinal

    ' Delegate to operations class (pure DAO). Note: Eliminar operates on the NC.
    Set m_NCOp = New NCProyectoOperaciones
    Set m_NCOp.nc = p_NC
    p_NC.MotivoBorrado = motivoFinal

    opErr = ""
    m_NCOp.Eliminar EnumSino.No, opErr
    If opErr <> "" Then
        p_Error = opErr
        modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildFail(logs, p_Error)
        Exit Function
    End If

    modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildOk(logs, "deleted")
    Exit Function

EH:
    p_Error = "modFrmNCProyectoARHelper_ConfirmarBorradoNC: " & Err.Description
    modFrmNCProyectoARHelper_ConfirmarBorradoNC = BuildFail(logs, p_Error)
End Function

' -----------------------------------------------------------------------------
' ConfirmarHabilitacionNC: handle the enable-NC confirmation flow (no motivo).
'   Production (p_PromptResult=0): shows MsgBox "¿Desea habilitar la NC?".
'     If user clicks Yes and NC is currently borrada, delegates to
'     NCProyectoOperaciones.Habilitar. Otherwise rejects.
'   Atom mode (p_PromptResult=-1): returns JSON with the habilitacion prompt text.
'   Test simulating: uses caller-injected choice.
'
'   Returns: JSON string with {ok, value, error, logs}.
'
'   NOTE: AR forms do NOT currently invoke _ConfirmarHabilitacionNC (the
'   habilitacion flow lives in FormNCProyectoAcciones, not in the single-AR
'   edit form). This helper is shipped for coverage parity with the slice 3
'   Acciones pair and to provide the canonical 3-function shape per the SDD.
' -----------------------------------------------------------------------------
Public Function modFrmNCProyectoARHelper_ConfirmarHabilitacionNC( _
                                            ByRef p_NC As NCProyecto, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim respuesta As Long
    Dim m_NCOp As NCProyectoOperaciones
    Dim opErr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_NC is Nothing.
    If p_NC Is Nothing Then
        modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildFail(logs, "ConfirmarHabilitacionNC: p_NC is Nothing")
        Exit Function
    End If

    ' Atom mode: return prompt text contract. No real modal.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildOk(logs, MSG_HABILITACION_PROMPT)
        Exit Function
    End If

    ' Production or test simulating user choice.
    If p_PromptResult = PROMPT_PRODUCTION Then
        ' Production: show real MsgBox.
        respuesta = MsgBox( _
            "¿Desea habilitar la no conformidad de proyecto seleccionada?", _
            vbExclamation + vbYesNo + vbDefaultButton2, MSG_HABILITACION_TITLE)
    Else
        respuesta = p_PromptResult
    End If

    If respuesta <> VB_YES Then
        modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildOk(logs, "user_rejected")
        Exit Function
    End If

    ' Edge: NC must be borrada to be habilitable. Race-safe guard.
    If p_NC.Borrado = False Then
        p_Error = "La NC no está borrada; no se puede habilitar"
        modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' Delegate to operations class (pure DAO).
    Set m_NCOp = New NCProyectoOperaciones
    Set m_NCOp.nc = p_NC

    opErr = ""
    m_NCOp.Habilitar opErr
    If opErr <> "" Then
        p_Error = opErr
        modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildFail(logs, p_Error)
        Exit Function
    End If

    modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildOk(logs, "habilitada")
    Exit Function

EH:
    p_Error = "modFrmNCProyectoARHelper_ConfirmarHabilitacionNC: " & Err.Description
    modFrmNCProyectoARHelper_ConfirmarHabilitacionNC = BuildFail(logs, p_Error)
End Function