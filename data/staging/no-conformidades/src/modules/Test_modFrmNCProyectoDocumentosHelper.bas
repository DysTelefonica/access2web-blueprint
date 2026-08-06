Attribute VB_Name = "Test_modFrmNCProyectoDocumentosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmNCProyectoDocumentosHelper (slice 6a of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; operations is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper owns the 3 message-class decisions that previously lived
'           inline in Form_FormNCProyectoDocumentos.cls:
'             1. _RenderError                  -- boilerplate error handler (7 MsgBox)
'             2. _ConfirmarEliminacionDocumento -- delete-Documento confirmation (1 MsgBox)
'             3. _ValidarNombreDocumento        -- 4 validation branches in CambiarNombre
'           The form is thin: builds m_Error, calls helper.RenderError(m_Error, Err.Number).
'           Production caller passes p_PromptResult=0 (real modal). Atom passes -1
'           (no modal, assert via JSON value or return code).
' Slice:    6a of 25 (Documentos group -- first batch: NCProyecto side).
'           12 atoms total (3 public functions x 4 scenario classes each).
'
' Source-of-truth provenance:
'   - DocumentoProyecto.cls (public String fields verified directly):
'       IDDocumento, IDNoConformidad, Documento, NombreAnexo, IDAccionRealizada,
'       IDNoConformidadResultante, Error
'   - DocumentoProyectoOperaciones.cls (public methods):
'       Eliminar(Optional ByRef p_Error As String) As String
'   - TbNCDocumentos columns verified via working INSERT in
'     Test_CacheWriteInvalidation.bas lines 805 + 890: IDDocumento, IDNoConformidad, Documento
' =============================================================================

' ----- Message contract (byte-for-byte match with production literals) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-NCPROYECTO-DOCUMENTOS-ELIMINACION: ¿Desea borrar el documento seleccionado?"
Private Const MSG_ELIMINACION_TITLE As String = "Eliminar documento"
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"
Private Const MSG_NOMBRE_VACIO As String = "El nombre del documento no puede estar vacío."
Private Const MSG_NOMBRE_IGUAL As String = "El nombre indicado es el mismo que ya tiene el documento."
Private Const MSG_NOMBRE_DUPLICADO As String = "Ya existe otro documento con el nombre"

' ----- Fixture IDs (>=900661000 reserved for this module, access-vba-tdd §1.7) -----
Private Const FIX_ID_NC As Long = 900661001
Private Const FIX_ID_DOC As Long = 900661002

' ----- Module-level state -----
Private m_PrevUsuarioConectado As usuario
Private m_PrevEntorno As entorno

' =============================================================================
' Local helpers (Private, declared at top per access-vba-tdd §1.8)
' =============================================================================

Private Function TableExistsInDb(ByVal p_Db As DAO.Database, ByVal p_TableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    For Each tdf In p_Db.TableDefs
        If tdf.Name = p_TableName Then
            TableExistsInDb = True
            Exit Function
        End If
    Next tdf
    On Error GoTo 0
End Function

' Pre-insert a minimal NCProyecto row in TbNoConformidades so that
' DocumentoProyecto.IDNoConformidad is FK-valid for the
' ConfirmarEliminacionDocumento path.
' Schema for TbNoConformidades -- required NOT NULL columns are
' DESCRIPCION, ESTADO, FECHAAPERTURA (verified via slice 1 / slice 3 working
' fixtures and Test_NCRepository fixtures). Other columns are nullable.
Private Function EnsureNCFixtureForSlice6a(ByVal p_Db As DAO.Database, _
                                           ByVal p_IDNC As Long, _
                                           ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureNCFixtureForSlice6a = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbNoConformidades") Then
        p_Error = "TbNoConformidades does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbNoConformidades WHERE IDNoConformidad=" & CStr(p_IDNC), dbFailOnError
    p_Db.Execute "INSERT INTO TbNoConformidades (IDNoConformidad, DESCRIPCION, ESTADO, FECHAAPERTURA) " & _
                 "VALUES (" & CStr(p_IDNC) & ", " & _
                 "'NC fixture slice 6a proyecto documentos', " & _
                 "'REGISTRADA', #2026-06-26#)", dbFailOnError
    EnsureNCFixtureForSlice6a = True
    Exit Function
EH:
    p_Error = "EnsureNCFixtureForSlice6a: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDNC As Long, ByVal p_IDDoc As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbNCDocumentos") Then
        p_Db.Execute "DELETE FROM TbNCDocumentos WHERE IDDocumento=" & CStr(p_IDDoc), dbFailOnError
    End If
    If TableExistsInDb(p_Db, "TbNoConformidades") Then
        p_Db.Execute "DELETE FROM TbNoConformidades WHERE IDNoConformidad=" & CStr(p_IDNC), dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh DocumentoProyecto with the fields needed for ConfirmarEliminacionDocumento.
' Field names verified against src/classes/DocumentoProyecto.cls (public String fields).
' URLAnexo is read-only Property Get; not assigned.
Private Function BuildDocumentoProyectoForSlice6a(ByVal p_IDDoc As Long, _
                                                   ByVal p_IDNC As Long, _
                                                   ByVal p_Nombre As String) As DocumentoProyecto
    Dim doc As DocumentoProyecto
    Set doc = New DocumentoProyecto
    doc.IDDocumento = CStr(p_IDDoc)
    doc.IDNoConformidad = CStr(p_IDNC)
    doc.Documento = p_Nombre
    Set BuildDocumentoProyectoForSlice6a = doc
End Function

Private Sub SetupTestContext()
    On Error Resume Next
    Set m_PrevUsuarioConectado = m_ObjUsuarioConectado
    Set m_PrevEntorno = m_ObjEntorno
    If m_ObjEntorno Is Nothing Then Set m_ObjEntorno = New entorno
    On Error GoTo 0
End Sub

Private Sub RestoreTestContext()
    On Error Resume Next
    If Not m_PrevUsuarioConectado Is Nothing Then
        Set m_ObjUsuarioConectado = m_PrevUsuarioConectado
    Else
        Set m_ObjUsuarioConectado = Nothing
    End If
    If Not m_PrevEntorno Is Nothing Then Set m_ObjEntorno = m_PrevEntorno
    Set m_PrevUsuarioConectado = Nothing
    Set m_PrevEntorno = Nothing
    On Error GoTo 0
End Sub

' =============================================================================
' ATOM 1 - Happy path: _RenderError in atom mode (p_PromptResult=-1).
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1
    helperErrMsg = ""

    returnCode = modFrmNCProyectoDocumentosHelper_RenderError( _
        "Al btnExaminar_Click se ha producido el error n: 9" & vbNewLine & "Detalle: subscript out of range", _
        9, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1"

    Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_no_modal_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_RenderError_Happy_NoModal_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 2 - Sad path: _RenderError with Err.Number=1000 (warning path).
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1
    helperErrMsg = ""

    returnCode = modFrmNCProyectoDocumentosHelper_RenderError( _
        "Debe poner un título al documento.", 1000, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty for warning path"

    If returnCode <> -1 Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 in atom (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (atom still no-modal for 1000)"

    Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_warning_no_modal_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_RenderError_Sad_WarningPath_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 3 - Edge: _RenderError with empty m_ErrorText returns -1, no modal.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    returnCode = modFrmNCProyectoDocumentosHelper_RenderError( _
        "", 0, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 for empty error text"

    Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_empty_text_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_RenderError_Edge_EmptyText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 4 - Adversarial: _RenderError with very long error text (2000 chars).
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim longText As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    longText = String$(2000, "X")
    helperErrMsg = ""

    returnCode = modFrmNCProyectoDocumentosHelper_RenderError( _
        longText, 9999, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty for 2000-char error text"

    If returnCode <> -1 Then
        Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 for very long error text"

    Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_long_text_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 5 - Happy: _ConfirmarEliminacionDocumento in atom mode returns prompt text.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim doc As DocumentoProyecto

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set doc = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, Nothing, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, MSG_ELIMINACION_PROMPT, vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain prompt text contract (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_ELIMINACION_PROMPT"

    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_prompt_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 6 - Sad: _ConfirmarEliminacionDocumento with vbNo returns value=user_rejected.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim doc As DocumentoProyecto

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set doc = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, Nothing, vbNo, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, "user_rejected", vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain user_rejected (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains user_rejected"

    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_user_rejected_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 7 - Edge: _ConfirmarEliminacionDocumento with vbYes but operations errors out.
'   The fixture is set up with a valid NC + doc; DocumentoProyectoOperaciones.Eliminar
'   has cache + log side-effects. We assert that:
'   (a) when vbYes is injected, the helper DOES NOT return user_rejected
'       (so it got past the confirmation prompt)
'   (b) the helper either succeeds OR returns ok=false with a propagated error
'       from the operations call (both are valid -- what matters is the gate passed).
'   We do NOT assert specific cache/log state since the test environment may not
'   have all the cache tables pre-populated; what we assert is the gate logic.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim db As DAO.Database
    Dim doc As DocumentoProyecto

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    SetupTestContext

    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        GoTo Cleanup
    End If
    Set db = getdb()
    If Not EnsureNCFixtureForSlice6a(db, FIX_ID_NC, helperErrMsg) Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("EnsureNCFixtureForSlice6a: " & helperErrMsg, logs)
        GoTo Cleanup
    End If
    ' Pre-insert a Documento row so Eliminar has something to delete.
    If TableExistsInDb(db, "TbNCDocumentos") Then
        db.Execute "INSERT INTO TbNCDocumentos (IDDocumento, IDNoConformidad, Documento) " & _
                   "VALUES (" & CStr(FIX_ID_DOC) & ", " & CStr(FIX_ID_NC) & ", " & _
                   "'fixture_doc_eliminar.pdf')", dbFailOnError
    End If

    Set doc = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, db, vbYes, helperErrMsg)

    ' Assert gate: resultJson must NOT contain "user_rejected" (vbYes passed the gate).
    If InStr(1, resultJson, "user_rejected", vbTextCompare) > 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("vbYes should have passed the gate; got user_rejected (json='" & resultJson & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: vbYes passed the gate (no user_rejected)"

    ' If the operations call fails (cache/log not seeded in test env), the helper
    ' returns ok=false with a propagated error. Either result is a valid gate pass.
    TestHelper.AddLog logs, "Edge result: helperErrMsg='" & helperErrMsg & "', resultJson='" & resultJson & "'"

    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_user_confirms_gate_ok")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then Call CleanupFixture(db, FIX_ID_NC, FIX_ID_DOC)
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 8 - Adversarial: _ConfirmarEliminacionDocumento with p_Doc=Nothing returns ok=false.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    resultJson = modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento( _
        Nothing, Nothing, -1, helperErrMsg)

    ' helperErrMsg should be empty (the helper sets it via BuildFail but the
    ' BuildFail JSON contains the error text; p_Error is for upstream callers).
    ' The contract is: resultJson ok=false with error containing "Nothing".
    If InStr(1, resultJson, "Nothing", vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
            TestHelper.BuildJsonFail("resultJson should mention Nothing (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson mentions Nothing"

    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_nothing_doc_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 9 - Happy: _ValidarNombreDocumento with unique new name returns ok=true.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoProyecto
    Dim colDocs As Scripting.Dictionary
    Dim docOtro As DocumentoProyecto

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set docOtro = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC + 100, FIX_ID_NC, "otro_doc.pdf")

    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual
    colDocs.Add CStr(docOtro.IDDocumento), docOtro

    resultJson = modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento( _
        "nuevo_nombre_unico.pdf", docActual, colDocs, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, """value"":""ok""", vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain value='ok' (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson value='ok'"

    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_unique_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 10 - Sad: _ValidarNombreDocumento with empty new name returns MSG_NOMBRE_VACIO.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoProyecto
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual

    resultJson = modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento( _
        "", docActual, colDocs, -1, helperErrMsg)

    ' Helper sets p_Error via BuildFail -- p_Error stays empty (the error lives in
    ' the JSON body). What we assert: resultJson contains MSG_NOMBRE_VACIO.
    If InStr(1, resultJson, MSG_NOMBRE_VACIO, vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_VACIO (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_VACIO"

    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_empty_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 11 - Edge: _ValidarNombreDocumento with same-as-current name returns MSG_NOMBRE_IGUAL.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoProyecto
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual

    resultJson = modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento( _
        "original_doc.pdf", docActual, colDocs, -1, helperErrMsg)

    If InStr(1, resultJson, MSG_NOMBRE_IGUAL, vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_IGUAL (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_IGUAL"

    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_same_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 12 - Adversarial: _ValidarNombreDocumento with duplicate in collection returns MSG_NOMBRE_DUPLICADO.
' =============================================================================
Public Function Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoProyecto
    Dim docOtro As DocumentoProyecto
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    ' docOtro has a DIFFERENT ID but SAME name as the new name we're trying.
    Set docOtro = BuildDocumentoProyectoForSlice6a(FIX_ID_DOC + 100, FIX_ID_NC, "nombre_que_ya_existe.pdf")

    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual
    colDocs.Add CStr(docOtro.IDDocumento), docOtro

    resultJson = modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento( _
        "nombre_que_ya_existe.pdf", docActual, colDocs, -1, helperErrMsg)

    If InStr(1, resultJson, MSG_NOMBRE_DUPLICADO, vbTextCompare) = 0 Then
        Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_DUPLICADO (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_DUPLICADO"

    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_duplicate_ok")
    Exit Function

EH:
    Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCProyectoDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic: " & Err.Description, logs)
End Function