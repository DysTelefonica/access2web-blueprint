Attribute VB_Name = "modFrmAuditoriaDocumentosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Form helper for Form_FormAuditoriaDocumentos (slice 6b of form-thin-helper-refactor).
'
' This helper owns the 3 message-class decisions that previously lived inline
' in Form_FormAuditoriaDocumentos.cls event handlers:
'   1. _RenderError                  -- boilerplate error handler (8 MsgBox in errores: blocks)
'   2. _ConfirmarEliminacionDocumento -- MsgBox("¿Desea borrar el documento seleccionado?") + Eliminar
'   3. _ValidarNombreDocumento        -- the 4 validation branches in ComandoCambiarNombre_Click
'                                          (empty / same-as-current / duplicate-in-list)
'
' The form is THIN: builds m_Error, calls helper, renders result.
' Operations class (DocumentoAuditoriaOperaciones.cls) is pure DAO: no UI awareness.
'
' NOTE on the form's "alta" mode (m_ObjColDocumentosAuditoriaParaAlta): that flow
' uses a Scripting.Dictionary keyed by Name (not by IDDocumento). The validation
' in alta mode is simpler (Existence check only) and stays inline in the form
' for now. Only the "existing" mode validation is extracted into the helper.
' Wiring up the alta-mode validation to a helper variant is a follow-up commit.
'
' Skill:    access-vba-tdd §1.1 (operations class is pure DAO; helper owns UI)
'           access-vba-tdd §1.6 (canonical signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology Hard Rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Slice:    6b of 25 (Documentos group -- second batch: standalone Auditoria form;
'           companion is modFrmARProyectoDocumentosHelper.bas + modFrmARAuditoriaDocumentosHelper.bas).
' =============================================================================

' ----- Message contracts (byte-for-byte match with Test_modFrmAuditoriaDocumentosHelper.bas) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-AUDITORIA-DOCUMENTOS-ELIMINACION: ¿Desea borrar el documento seleccionado?"
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
'   blocks. Mirror of modFrmARAuditoriaDocumentosHelper_RenderError (sister module).
' -----------------------------------------------------------------------------
Public Function modFrmAuditoriaDocumentosHelper_RenderError( _
                                            ByVal p_ErrorText As String, _
                                            ByVal p_ErrNumber As Long, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As Long
    On Error GoTo EH
    p_Error = ""

    ' Edge case: empty error text. No modal needed.
    If Len(Trim$(p_ErrorText)) = 0 Then
        modFrmAuditoriaDocumentosHelper_RenderError = PROMPT_ATOM
        Exit Function
    End If

    ' Atom/test mode: no real MsgBox.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmAuditoriaDocumentosHelper_RenderError = PROMPT_ATOM
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
        modFrmAuditoriaDocumentosHelper_RenderError = MsgBox(p_ErrorText, style, MSG_ERROR_TITLE)
        Exit Function
    End If

    ' Test simulating user response.
    modFrmAuditoriaDocumentosHelper_RenderError = p_PromptResult
    Exit Function

EH:
    p_Error = "modFrmAuditoriaDocumentosHelper_RenderError: " & Err.Description
    modFrmAuditoriaDocumentosHelper_RenderError = vbOK
End Function

' -----------------------------------------------------------------------------
' ConfirmarEliminacionDocumento: handle the delete-DocumentoAuditoria confirmation.
'   Production: shows MsgBox "¿Desea borrar el documento seleccionado?".
'     If user clicks Yes, delegates to DocumentoAuditoriaOperaciones.Eliminar.
'   Atom mode (p_PromptResult=-1): returns JSON with the eliminacion prompt text
'     contract, no real modal.
'   Test simulating (p_PromptResult=vbYes/vbNo): uses that choice directly.
' -----------------------------------------------------------------------------
Public Function modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento( _
                                            ByRef p_Doc As DocumentoAuditoria, _
                                            Optional ByRef db As DAO.Database = Nothing, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim respuesta As Long
    Dim m_DocOp As DocumentoAuditoriaOperaciones
    Dim opErr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_Doc is Nothing.
    If p_Doc Is Nothing Then
        modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, "ConfirmarEliminacionDocumento: p_Doc is Nothing")
        Exit Function
    End If

    ' Atom mode: return prompt text contract. No real modal.
    If p_PromptResult = PROMPT_ATOM Then
        modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, MSG_ELIMINACION_PROMPT)
        Exit Function
    End If

    ' Production or test simulating user choice.
    If p_PromptResult = PROMPT_PRODUCTION Then
        respuesta = MsgBox( _
            "¿Desea borrar el documento seleccionado?", _
            vbExclamation + vbYesNo + vbDefaultButton2, MSG_ELIMINACION_TITLE)
    Else
        respuesta = p_PromptResult
    End If

    If respuesta <> VB_YES Then
        modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, "user_rejected")
        Exit Function
    End If

    ' Delegate to operations class (pure DAO).
    Set m_DocOp = New DocumentoAuditoriaOperaciones
    Set m_DocOp.Documento = p_Doc

    opErr = ""
    m_DocOp.Eliminar opErr
    If opErr <> "" Then
        p_Error = opErr
        modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildOk(logs, "deleted")
    Exit Function

EH:
    p_Error = "modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento: " & Err.Description
    modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento = BuildFail(logs, p_Error)
End Function

' -----------------------------------------------------------------------------
' ValidarNombreDocumento: encapsulate the 4 validation branches in
'   Form_FormAuditoriaDocumentos.ComandoCambiarNombre_Click (existing-mode path,
'   lines 367-387). The alta-mode path (lines 345-365) uses a Dictionary keyed by
'   Name (not by IDDocumento) and stays inline in the form for now -- it follows
'   a different validation pattern (Existence check only, no ID comparison).
' -----------------------------------------------------------------------------
Public Function modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento( _
                                            ByVal p_NombreNuevo As String, _
                                            ByRef p_DocActual As DocumentoAuditoria, _
                                            ByRef p_ColDocs As Scripting.Dictionary, _
                                            Optional ByRef p_PromptResult As Long = 0, _
                                            Optional ByRef p_Error As String = "" _
                                            ) As String
    Dim logs As Collection
    Dim m_vID As Variant
    Dim m_DocIterado As DocumentoAuditoria

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    p_Error = ""

    ' Adversarial: p_DocActual is Nothing.
    If p_DocActual Is Nothing Then
        modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, "ValidarNombreDocumento: p_DocActual is Nothing")
        Exit Function
    End If

    ' (1) Empty name.
    If Len(Trim$(p_NombreNuevo)) = 0 Then
        p_Error = MSG_NOMBRE_VACIO
        modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' (2) Same as current.
    If StrComp(p_NombreNuevo, p_DocActual.Documento, vbTextCompare) = 0 Then
        p_Error = MSG_NOMBRE_IGUAL
        modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
        Exit Function
    End If

    ' (3) Duplicate in the collection (skip the current doc by ID).
    If Not p_ColDocs Is Nothing Then
        For Each m_vID In p_ColDocs
            Set m_DocIterado = p_ColDocs(m_vID)
            If StrComp(m_DocIterado.Documento, p_NombreNuevo, vbTextCompare) = 0 Then
                If m_DocIterado.IDDocumento <> p_DocActual.IDDocumento Then
                    p_Error = MSG_NOMBRE_DUPLICADO & " '" & p_NombreNuevo & "'."
                    modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
                    Exit Function
                End If
            End If
        Next
    End If

    ' (4) OK.
    modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildOk(logs, "ok")
    Exit Function

EH:
    p_Error = "modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento: " & Err.Description
    modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento = BuildFail(logs, p_Error)
End Function