Attribute VB_Name = "modFrmNCProyectoDocumentosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormNCProyectoDocumentos (slice 6a of form-thin-helper-refactor).
'
' This helper owns the 3 message-class decisions that previously lived inline
' in Form_FormNCProyectoDocumentos.cls event handlers:
'   1. _RenderError               -- boilerplate error handler (7 of 7 MsgBox in errores: blocks)
'   2. _ConfirmarEliminacionDocumento -- MsgBox("¿Desea borrar el documento seleccionado?") + Eliminar
'   3. _ValidarNombreDocumento     -- the 4 validation branches in ComandoCambiarNombre_Click
'                                       (empty / same-as-current / duplicate-in-list)
'
' The form is THIN: builds m_Error, calls helper, renders result.
' Operations class (DocumentoProyectoOperaciones.cls) is pure DAO: no UI awareness.
'
' Skill:    access-vba-tdd §1.1 (operations class is pure DAO; helper owns UI)
'           access-vba-tdd §1.6 (canonical signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology Hard Rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Slice:    6a of 25 (Documentos group -- first batch: NCProyecto side; companion is
'           modFrmNCAuditoriaDocumentosHelper.bas on NCAuditoria side; remaining 3
'           forms in slice 6b).
' =============================================================================

' ----- Message contracts (byte-for-byte match with Test_modFrmNCProyectoDocumentosHelper.bas) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-NCPROYECTO-DOCUMENTOS-ELIMINACION: ¿Desea borrar el documento seleccionado?"
Private Const MSG_ELIMINACION_TITLE As String = "Eliminar documento"
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_NOMBRE_VACIO As String = "El nombre del documento no puede estar vacío."
Private Const MSG_NOMBRE_IGUAL As String = "El nombre indicado es el mismo que ya tiene el documento."
Private Const MSG_NOMBRE_DUPLICADO As String = "Ya existe otro documento con el nombre"

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
Public Function modFrmNCProyectoDocumentosHelper_RenderError( _
                                            ByVal p_ErrorText As String, _
                                            ByVal p_ErrNumber As Long, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As Long
    On Error GoTo EH
    p_Error = ""

    ' Edge case: empty error text. No modal needed.
    If Len(Trim$(p_ErrorText)) = 0 Then
        modFrmNCProyectoDocumentosHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Atom/test mode: no real MsgBox.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCProyectoDocumentosHelper_RenderError = PROMPT_ATOM
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
        modFrmNCProyectoDocumentosHelper_RenderError = MsgBox(p_ErrorText, style, MSG_ERROR_TITLE)
        Exit Function
    End If

    ' Test simulating user response.
    modFrmNCProyectoDocumentosHelper_RenderError = p_PromptResult
    Exit Function

EH:
    p_Error = "modFrmNCProyectoDocumentosHelper_RenderError: " & Err.Description
    modFrmNCProyectoDocumentosHelper_RenderError = vbOK
End Function

' -----------------------------------------------------------------------------
' ConfirmarEliminacionDocumento: handle the delete-Documento confirmation flow.
'   This is the single user-confirmation MsgBox in ComandoEliminar_Click.
'   Production (p_PromptResult=0): shows MsgBox "¿Desea borrar el documento seleccionado?".
'     If user clicks Yes, delegates to DocumentoProyectoOperaciones.Eliminar.
'     If user clicks No, returns "user_rejected".
'   Atom mode (p_PromptResult=-1): returns JSON with the eliminacion prompt text
'     contract, no real modal. Atom asserts the prompt text via JSON value.
'   Test simulating (p_PromptResult=vbYes/vbNo): uses that choice directly.
'
'   Notes:
'   - Unlike slice 3/5's _ConfirmarBorradoNC, this helper has NO motivo prompt
'     (the Documentos flow is a simple yes/no for definitive deletion; the Eliminar
'     operation does not request user-provided motivo in this codebase).
'   - DocumentoProyectoOperaciones.Eliminar checks (1) the file is not open,
'     (2) AR consistency (only one doc + cerrada), (3) cache invalidation,
'     (4) log write. The helper propagates those errors verbatim via JSON
'     ok=false / error=<opErr>.
'
'   Returns: JSON string with {ok, value, error, logs}.
' -----------------------------------------------------------------------------
Public Function modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento( _
                                            ByRef p_Doc As DocumentoProyecto, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim respuesta As Long
    Dim m_DocOp As DocumentoProyectoOperaciones
    Dim opErr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_Doc is Nothing.
    If p_Doc Is Nothing Then
        modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, "ConfirmarEliminacionDocumento: p_Doc is Nothing")
        Exit Function
    End If

    ' Atom mode: return prompt text contract. No real modal.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, MSG_ELIMINACION_PROMPT)
        Exit Function
    End If

    ' Production or test simulating user choice.
    If p_PromptResult = PROMPT_PRODUCTION Then
        ' Production: show real MsgBox.
        respuesta = MsgBox( _
            "¿Desea borrar el documento seleccionado?", _
            vbExclamation + vbYesNo + vbDefaultButton2, MSG_ELIMINACION_TITLE)
    Else
        respuesta = p_PromptResult
    End If

    If respuesta <> VB_YES Then
        modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, "user_rejected")
        Exit Function
    End If

    ' Delegate to operations class (pure DAO).
    Set m_DocOp = New DocumentoProyectoOperaciones
    Set m_DocOp.Documento = p_Doc

    opErr = ""
    m_DocOp.Eliminar opErr
    If opErr <> "" Then
        p_Error = opErr
        modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, "deleted")
    Exit Function

EH:
    p_Error = "modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento: " & Err.Description
    modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, p_Error)
End Function

' -----------------------------------------------------------------------------
' ValidarNombreDocumento: encapsulate the 4 validation branches in
'   Form_FormNCProyectoDocumentos.ComandoCambiarNombre_Click (lines 367-397):
'     (1) p_NombreNuevo is empty        -> MSG_NOMBRE_VACIO
'     (2) p_NombreNuevo equals current   -> MSG_NOMBRE_IGUAL
'     (3) p_NombreNuevo is in p_ColDocs with different ID -> MSG_NOMBRE_DUPLICADO
'     (4) OK                             -> "ok" / empty value
'
'   Atom mode (p_PromptResult=-1): returns JSON with the validation contract
'     decision (one of the MSG_* constants or "ok"); no modal.
'   Production/test simulating (>=0): currently the validation does NOT show
'     a modal; it sets p_Error. Modal rendering of the error text is the
'     form's errores: block via _RenderError. p_PromptResult is unused for
'     this helper (kept for symmetry with the slice 3/5 signature).
'
'   The collection p_ColDocs is the NC's DocumentosCompletos (a
'   Scripting.Dictionary of DocumentoProyecto keyed by IDDocumento). The form
'   passes it in directly so the helper has no implicit state lookup.
'
'   Returns: JSON string with {ok, value, error, logs}. value carries the
'   specific error message contract (one of MSG_NOMBRE_*) so atoms can assert
'   the right branch fired.
' -----------------------------------------------------------------------------
Public Function modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento( _
                                            ByVal p_NombreNuevo As String, _
                                            ByRef p_DocActual As DocumentoProyecto, _
                                            ByRef p_ColDocs As Scripting.Dictionary, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim m_vID As Variant
    Dim m_DocIterado As DocumentoProyecto

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_DocActual is Nothing.
    If p_DocActual Is Nothing Then
        modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, "ValidarNombreDocumento: p_DocActual is Nothing")
        Exit Function
    End If

    ' (1) Empty name.
    If Len(Trim$(p_NombreNuevo)) = 0 Then
        p_Error = MSG_NOMBRE_VACIO
        modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' (2) Same as current.
    If StrComp(p_NombreNuevo, p_DocActual.Documento, vbTextCompare) = 0 Then
        p_Error = MSG_NOMBRE_IGUAL
        modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' (3) Duplicate in the collection (skip the current doc by ID).
    If Not p_ColDocs Is Nothing Then
        For Each m_vID In p_ColDocs
            Set m_DocIterado = p_ColDocs(m_vID)
            If StrComp(m_DocIterado.Documento, p_NombreNuevo, vbTextCompare) = 0 Then
                If m_DocIterado.IDDocumento <> p_DocActual.IDDocumento Then
                    p_Error = MSG_NOMBRE_DUPLICADO & " '" & p_NombreNuevo & "' en esta lista."
                    modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
                    Exit Function
                End If
            End If
        Next
    End If

    ' (4) OK.
    modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildOk(logs, "ok")
    Exit Function

EH:
    p_Error = "modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento: " & Err.Description
    modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
End Function