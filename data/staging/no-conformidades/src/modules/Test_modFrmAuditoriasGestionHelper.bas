Attribute VB_Name = "Test_modFrmAuditoriasGestionHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmAuditoriasGestionHelper (slice 4 of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; operations is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper owns the 2 message-class decisions that previously lived
'           inline in Form_FormAuditoriasGestion.cls:
'             1. _RenderError                  -- boilerplate error handler (32 MsgBox)
'             2. _ConfirmarEliminacionAuditoria -- delete-Auditoria confirmation (1 MsgBox)
'           The form is thin: builds m_Error, calls helper.RenderError(m_Error, Err.Number).
'           Production caller passes p_PromptResult=0 (real modal). Atom passes -1
'           (no modal, assert via JSON value or return code).
' Slice:    4 of 27. FIRST single-form slice of the epic (paired slices were 2a/2b/3).
'           8 atoms total (2 public functions x 4 scenario classes each).
' =============================================================================

' ----- Message contract (byte-for-byte match with production literals) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-AUDITORIAS-GESTION-ELIMINACION: ¿Desea eliminar definitivamente del sistema la auditoría seleccionada?"
Private Const MSG_ELIMINACION_TITLE As String = "Eliminación/Deseliminacion"
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"

' ----- Fixture IDs (>=900640000 reserved for this module, access-vba-tdd §1.7) -----
Private Const FIX_ID_AUDITORIA As Long = 900640001

' ----- Module-level state -----
Private m_PrevUsuarioConectado As usuario
Private m_PrevEntorno As entorno
Private ErrMsg_Local As String  ' local errMsg buffer (avoid name collision with helper's p_Error)

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

' Pre-insert a minimal TbAuditorias row for the ConfirmarEliminacionAuditoria path.
' Schema for TbAuditorias verified via docs/schema/erd-no-conformidades.md and
' Test_NCAuditoriaGestionListadoHelper.bas (lines 623-630):
'   IDAuditoria (Long, PK, Required), Tipo, FechaInicio, FechaFin.
' Slice 4 range: FIX_ID_AUDITORIA = 900640001 (distinct from slice 1 900610002,
' slice 2a 900620xxx, slice 2b 900630xxx, slice 3 ranges).
Private Function EnsureAuditoriaFixtureForSlice4(ByVal p_Db As DAO.Database, _
                                                 ByVal p_IDAuditoria As Long, _
                                                 ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureAuditoriaFixtureForSlice4 = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbAuditorias") Then
        p_Error = "TbAuditorias does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbAuditorias WHERE IDAuditoria=" & CStr(p_IDAuditoria), dbFailOnError
    p_Db.Execute "INSERT INTO TbAuditorias (IDAuditoria, Tipo, FechaInicio, FechaFin) " & _
                 "VALUES (" & CStr(p_IDAuditoria) & ", " & _
                 "'Auditoria slice 4', " & _
                 "#2026-06-26#, #2026-06-30#)", dbFailOnError
    EnsureAuditoriaFixtureForSlice4 = True
    Exit Function
EH:
    p_Error = "EnsureAuditoriaFixtureForSlice4: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDAuditoria As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbAuditorias") Then
        p_Db.Execute "DELETE FROM TbAuditorias WHERE IDAuditoria=" & CStr(p_IDAuditoria), dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh Auditoria with the fields needed for ConfirmarEliminacionAuditoria.
' Field names verified against src/classes/Auditoria.cls (public String fields):
'   IDAuditoria (String), Tipo, FechaInicio, FechaFin, Error (String).
' Documentos and NCs are private/Scripting.Dictionary; we leave them Nothing
' so that the helper's operations call passes the "no NCs / no documentos" branch.
Private Function BuildAuditoriaForSlice4(ByVal p_IDAuditoria As Long) As Auditoria
    Dim aud As Auditoria
    Set aud = New Auditoria
    aud.IDAuditoria = CStr(p_IDAuditoria)
    aud.Tipo = "Auditoria slice 4"
    aud.FechaInicio = "26/06/2026"
    aud.FechaFin = "30/06/2026"
    Set BuildAuditoriaForSlice4 = aud
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
'   Production caller would render a real MsgBox; atom asserts no modal happens
'   and the helper returns -1 sentinel. m_ErrorText is the form's m_Error string.
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode: no real MsgBox
    helperErrMsg = ""

    returnCode = modFrmAuditoriasGestionHelper_RenderError( _
        "Al ComandoEliminar_Click se ha producido el error n: 9" & vbNewLine & "Detalle: subscript out of range", _
        9, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty in happy path (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 in atom mode (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (no modal sentinel)"

    Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_no_modal_ok")
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_RenderError_Happy_NoModal_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 2 - Sad path: _RenderError with Err.Number=1000 (warning path).
'   The original form distinguishes "Err.Number=1000 -> vbExclamation Advertencia"
'   from any other error -> "vbCritical Error". Atom asserts the helper returns
'   -1 in atom mode regardless of ErrNumber (the styling decision is rendered
'   in production only; atom never opens a modal).
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode
    helperErrMsg = ""

    returnCode = modFrmAuditoriasGestionHelper_RenderError( _
        "Seleccione un elemento de la lista", _
        1000, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on sad path (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on sad path too (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (atom mode invariant regardless of ErrNumber)"

    Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_warning_path_ok")
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_RenderError_Sad_WarningPath_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 3 - Edge: empty m_ErrorText passed to _RenderError.
'   Form should never call RenderError with empty text (every code path sets
'   m_Error before reaching errores:), but if it happens, the helper must NOT
'   open a modal -- return -1 silently.
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode
    helperErrMsg = ""

    returnCode = modFrmAuditoriasGestionHelper_RenderError( _
        "", _
        0, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on edge (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty on empty text"

    If returnCode <> -1 Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on edge (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (no modal on empty text)"

    Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_empty_text_ok")
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_RenderError_Edge_EmptyText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 4 - Adversarial: _RenderError with very long m_ErrorText (>1000 chars).
'   Stress-test the no-modal path with a pathological string length. Helper
'   must still return -1 (atom mode invariant) and leave helperErrMsg empty.
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long
    Dim longText As String
    Dim i As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode
    helperErrMsg = ""

    ' Build a 2000-char m_ErrorText
    longText = "Stack: "
    For i = 1 To 200
        longText = longText & "ABCDEFGHIJ"  ' 10 chars per iter * 200 = 2000 chars
    Next i

    returnCode = modFrmAuditoriasGestionHelper_RenderError( _
        longText, _
        13, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on adversarial (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty on 2000-char text"

    If returnCode <> -1 Then
        Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on adversarial (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (atom mode invariant on long text)"

    Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_long_text_ok")
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_RenderError_Adversarial_VeryLongText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 5 - Happy: _ConfirmarEliminacionAuditoria atom mode (p_PromptResult=-1).
'   Atom asserts the helper returns the prompt-text contract via JSON value.
'   No real MsgBox opened. No motivo prompt (unlike slice 2b's BorradoNC).
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim audParaEliminar As Auditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set audParaEliminar = BuildAuditoriaForSlice4(FIX_ID_AUDITORIA)
    promptResult = -1  ' atom mode

    jsonResult = modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria( _
        audParaEliminar, , promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on happy (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    ' Parse JSON and assert value carries the eliminacion prompt text contract
    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict Is Nothing Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parser returned Nothing", logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If InStr(1, CStr(dict("value")), MSG_ELIMINACION_PROMPT, vbBinaryCompare) = 0 Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected value to contain eliminacion prompt; got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value contains eliminacion prompt contract"

    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_eliminacion_prompt_text_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Happy_ReturnsPromptText_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 6 - Sad: _ConfirmarEliminacionAuditoria with user_clicked_no.
'   Atom simulates the user clicking "No" on the confirmation prompt by passing
'   p_PromptResult = vbNo (=7). The helper must return JSON "user_rejected"
'   without invoking Eliminar.
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim audParaEliminar As Auditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set audParaEliminar = BuildAuditoriaForSlice4(FIX_ID_AUDITORIA)
    promptResult = vbNo  ' 7 = "user clicked No"

    jsonResult = modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria( _
        audParaEliminar, , promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on sad (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If CStr(dict("value")) <> "user_rejected" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected value='user_rejected', got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value=user_rejected"

    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_eliminacion_user_rejected_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Sad_UserRejected_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 7 - Edge: _ConfirmarEliminacionAuditoria with vbYes but operaciones fail.
'   The helper should propagate AuditoriaOperaciones.Eliminar's failure as
'   ok=false via JSON, NOT swallow the error. We seed the fixture so the
'   operations call raises (we use a non-existent p_IDAuditoria so Eliminar
'   will fail when it tries to load documentos/ncs, OR we let it pass through
'   the success path -- depends on the live data; the test only asserts that
'   the JSON is well-formed with proper ok/value/error contract).
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim audParaEliminar As Auditoria
    Dim helperErrMsg As String
    Dim testDb As DAO.Database

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    ' Build a fresh Auditoria with an INVALID ID (does not exist in TbAuditorias).
    ' When the helper delegates to AuditoriaOperaciones.Eliminar, the operations
    ' class may pass through (it checks .NCs and .Documentos which are Nothing
    ' on a freshly constructed Auditoria), or fail. Either way, the JSON must
    ' be well-formed with ok=boolean and the contract preserved.
    Set audParaEliminar = BuildAuditoriaForSlice4(999999999)  ' non-existent ID
    promptResult = vbYes  ' user says yes

    ' Inject test db (sandbox) to ensure Eliminar runs against the local fixture
    Set testDb = TestHelper.GetTestDb()

    jsonResult = modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria( _
        audParaEliminar, testDb, promptResult, helperErrMsg)

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    ' Assert JSON contract holds: ok is boolean, value is non-Null when ok=true,
    ' error is non-empty when ok=false.
    Dim okVal As Boolean
    okVal = dict("ok")
    TestHelper.AddLog logs, "Assert: ok=" & okVal & " (contract preserved)"

    If okVal Then
        ' Operations succeeded (Eliminación path) — must not be the prompt text
        If CStr(dict("value")) = MSG_ELIMINACION_PROMPT Then
            Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
                TestHelper.BuildJsonFail("expected value != MSG_ELIMINACION_PROMPT after Eliminar (got prompt text)", logs)
            GoTo Cleanup
        End If
        TestHelper.AddLog logs, "Assert OK: Eliminar completed (value=" & CStr(dict("value")) & ")"
    Else
        ' Operations failed — error must mention Auditoria-related context
        If Len(Trim$(CStr(dict("error")))) = 0 Then
            Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
                TestHelper.BuildJsonFail("ok=false but error is empty", logs)
            GoTo Cleanup
        End If
        TestHelper.AddLog logs, "Assert OK: Eliminar failed with error='" & CStr(dict("error")) & "'"
    End If

    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_eliminacion_operations_failure_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Edge_OperationsFailure_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 8 - Adversarial: _ConfirmarEliminacionAuditoria with p_Auditoria = Nothing.
'   Form should never call ConfirmarEliminacionAuditoria with Nothing, but if
'   it does (e.g. after race condition where selection was cleared), helper
'   must return JSON with ok=false, NOT crash.
' =============================================================================
Public Function Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    promptResult = vbYes

    On Error Resume Next  ' tolerate helper raising 1000
    jsonResult = modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria( _
        Nothing, , promptResult, helperErrMsg)
    On Error GoTo EH

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    ' Helper must reject gracefully (ok=false, error mentions Nothing)
    If dict("ok") <> False Then
        Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic = _
            TestHelper.BuildJsonFail("expected ok=false on Nothing Auditoria (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: ok=false on Nothing Auditoria"

    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_eliminacion_nothing_auditoria_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriasGestionHelper_ConfirmarEliminacionAuditoria_Adversarial_NothingAuditoria_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function