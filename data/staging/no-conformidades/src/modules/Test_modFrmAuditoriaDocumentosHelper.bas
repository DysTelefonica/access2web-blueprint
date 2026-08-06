Attribute VB_Name = "Test_modFrmAuditoriaDocumentosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmAuditoriaDocumentosHelper (slice 6b of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; operations is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper owns the 3 message-class decisions that previously lived
'           inline in Form_FormAuditoriaDocumentos.cls:
'             1. _RenderError                  -- boilerplate error handler (8 MsgBox)
'             2. _ConfirmarEliminacionDocumento -- delete-DocumentoAuditoria confirmation (1 MsgBox)
'             3. _ValidarNombreDocumento        -- 4 validation branches in CambiarNombre
' Slice:    6b of 25 (Documentos group -- second batch: standalone Auditoria form).
'           12 atoms total (3 public functions x 4 scenario classes each).
'
' Source-of-truth provenance:
'   - DocumentoAuditoria.cls (public String fields verified directly):
'       IDDocumento, IDNoConformidad, Documento, NombreAnexo, IDAccionRealizada,
'       IDAuditoria, IDAuditoriaResultante, Error
'   - DocumentoAuditoriaOperaciones.cls (public methods):
'       Eliminar(Optional ByRef p_Error As String) As String
'       CambiarNombre(p_NombreDocumento As String, Optional ByRef p_Error As String) As String
'   - TbDocumentosAuditorias columns: IDDocumento, IDNoConformidad, Documento
'     (table name verified via DocumentoAuditoria.cls line 158:
'      constructor.getID("TbDocumentosAuditorias", "IDDocumento", getdb(), m_Error))
'   - TbNoConformidadesAuditoria PK is lowercase `id` (NOT IDNoConformidad)
'     -- verified via working INSERT in Test_modFrmNCAuditoriaDocumentosHelper.bas
'     slice 6a lines 84-89.
' =============================================================================

' ----- Message contract (byte-for-byte match with production literals) -----
Private Const MSG_ELIMINACION_PROMPT As String = "MSG-AUDITORIA-DOCUMENTOS-ELIMINACION: ¿Desea borrar el documento seleccionado?"
Private Const MSG_ELIMINACION_TITLE As String = "Eliminar documento"
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"
Private Const MSG_NOMBRE_VACIO As String = "El nombre del documento no puede estar vacío."
Private Const MSG_NOMBRE_IGUAL As String = "El nombre indicado es el mismo que ya tiene el documento."
Private Const MSG_NOMBRE_DUPLICADO As String = "Ya existe otro documento con el nombre"

' ----- Fixture IDs (>=900662300 reserved for this module, access-vba-tdd §1.7) -----
Private Const FIX_ID_NC As Long = 900662301
Private Const FIX_ID_DOC As Long = 900662302

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

' Pre-insert a minimal NCAuditoria row in TbNoConformidadesAuditoria.
' PK column is lowercase `id` (NOT IDNoConformidad) -- verified via working
' INSERT in Test_modFrmNCAuditoriaDocumentosHelper.bas slice 6a lines 84-89.
Private Function EnsureNCAuditoriaFixtureForSlice6b(ByVal p_Db As DAO.Database, _
                                                     ByVal p_IDNC As Long, _
                                                     ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureNCAuditoriaFixtureForSlice6b = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Error = "TbNoConformidadesAuditoria does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & CStr(p_IDNC), dbFailOnError
    p_Db.Execute "INSERT INTO TbNoConformidadesAuditoria (id, DESCRIPCION, ESTADO, FECHAAPERTURA, Tipo) " & _
                 "VALUES (" & CStr(p_IDNC) & ", " & _
                 "'NC fixture slice 6b auditoria documentos', " & _
                 "'REGISTRADA', #2026-06-26#, " & _
                 "'Auditoria')", dbFailOnError
    EnsureNCAuditoriaFixtureForSlice6b = True
    Exit Function
EH:
    p_Error = "EnsureNCAuditoriaFixtureForSlice6b: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDNC As Long, ByVal p_IDDoc As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbDocumentosAuditorias") Then
        p_Db.Execute "DELETE FROM TbDocumentosAuditorias WHERE IDDocumento=" & CStr(p_IDDoc), dbFailOnError
    End If
    If TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & CStr(p_IDNC), dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh DocumentoAuditoria with the fields needed for ConfirmarEliminacionDocumento.
' Field names verified against src/classes/DocumentoAuditoria.cls (public String fields).
Private Function BuildDocumentoAuditoriaForSlice6b(ByVal p_IDDoc As Long, _
                                                    ByVal p_IDNC As Long, _
                                                    ByVal p_Nombre As String) As DocumentoAuditoria
    Dim doc As DocumentoAuditoria
    Set doc = New DocumentoAuditoria
    doc.IDDocumento = CStr(p_IDDoc)
    doc.IDNoConformidad = CStr(p_IDNC)
    doc.Documento = p_Nombre
    Set BuildDocumentoAuditoriaForSlice6b = doc
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
' ATOM 1 - Happy: _RenderError in atom mode (p_PromptResult=-1).
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    returnCode = modFrmAuditoriaDocumentosHelper_RenderError( _
        "Al btnExaminar_Click se ha producido el error n: 9" & vbNewLine & "Detalle: subscript out of range", _
        9, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1"

    Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_no_modal_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_RenderError_Happy_NoModal_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 2 - Sad: _RenderError with Err.Number=1000 (warning path).
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    returnCode = modFrmAuditoriaDocumentosHelper_RenderError( _
        "Debe asignar un título al documento antes de aceptar.", 1000, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1"

    Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_warning_no_modal_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_RenderError_Sad_WarningPath_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 3 - Edge: _RenderError with empty m_ErrorText returns -1, no modal.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    returnCode = modFrmAuditoriaDocumentosHelper_RenderError( _
        "", 0, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1"

    Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_empty_text_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_RenderError_Edge_EmptyText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 4 - Adversarial: _RenderError with very long error text.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim longText As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    longText = String$(2000, "X")
    helperErrMsg = ""

    returnCode = modFrmAuditoriaDocumentosHelper_RenderError( _
        longText, 9999, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1"

    Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_long_text_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_RenderError_Adversarial_VeryLongText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 5 - Happy: _ConfirmarEliminacionDocumento atom mode returns prompt text.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim doc As DocumentoAuditoria

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set doc = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, Nothing, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, MSG_ELIMINACION_PROMPT, vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain prompt text (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_ELIMINACION_PROMPT"

    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_prompt_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Happy_ReturnsPromptText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 6 - Sad: _ConfirmarEliminacionDocumento with vbNo returns user_rejected.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim doc As DocumentoAuditoria

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set doc = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, Nothing, vbNo, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, "user_rejected", vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain user_rejected (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains user_rejected"

    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_user_rejected_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Sad_UserRejected_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 7 - Edge: _ConfirmarEliminacionDocumento with vbYes passes the gate.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim db As DAO.Database
    Dim doc As DocumentoAuditoria

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    SetupTestContext

    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        GoTo Cleanup
    End If
    Set db = getdb()
    If Not EnsureNCAuditoriaFixtureForSlice6b(db, FIX_ID_NC, helperErrMsg) Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("EnsureNCAuditoriaFixtureForSlice6b: " & helperErrMsg, logs)
        GoTo Cleanup
    End If
    ' Pre-insert a Documento row so Eliminar has something to delete.
    If TableExistsInDb(db, "TbDocumentosAuditorias") Then
        db.Execute "INSERT INTO TbDocumentosAuditorias (IDDocumento, IDNoConformidad, Documento) " & _
                   "VALUES (" & CStr(FIX_ID_DOC) & ", " & CStr(FIX_ID_NC) & ", " & _
                   "'fixture_doc_eliminar.pdf')", dbFailOnError
    End If

    Set doc = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "fixture_doc_eliminar.pdf")

    resultJson = modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento( _
        doc, db, vbYes, helperErrMsg)

    If InStr(1, resultJson, "user_rejected", vbTextCompare) > 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
            TestHelper.BuildJsonFail("vbYes should have passed the gate; got user_rejected (json='" & resultJson & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: vbYes passed the gate"

    TestHelper.AddLog logs, "Edge result: helperErrMsg='" & helperErrMsg & "', resultJson='" & resultJson & "'"

    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_user_confirms_gate_ok")
    GoTo Cleanup

EH:
    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Edge_UserConfirms_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then Call CleanupFixture(db, FIX_ID_NC, FIX_ID_DOC)
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 8 - Adversarial: _ConfirmarEliminacionDocumento with p_Doc=Nothing.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    resultJson = modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento( _
        Nothing, Nothing, -1, helperErrMsg)

    If InStr(1, resultJson, "Nothing", vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
            TestHelper.BuildJsonFail("resultJson should mention Nothing (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson mentions Nothing"

    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_elim_doc_nothing_doc_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ConfirmarEliminacionDocumento_Adversarial_NothingDoc_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 9 - Happy: _ValidarNombreDocumento with unique new name returns ok=true.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoAuditoria
    Dim colDocs As Scripting.Dictionary
    Dim docOtro As DocumentoAuditoria

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set docOtro = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC + 100, FIX_ID_NC, "otro_doc.pdf")

    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual
    colDocs.Add CStr(docOtro.IDDocumento), docOtro

    resultJson = modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento( _
        "nuevo_nombre_unico.pdf", docActual, colDocs, -1, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If InStr(1, resultJson, """value"":""ok""", vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain value='ok' (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson value='ok'"

    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_unique_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Happy_UniqueName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 10 - Sad: _ValidarNombreDocumento with empty new name returns MSG_NOMBRE_VACIO.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoAuditoria
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual

    resultJson = modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento( _
        "", docActual, colDocs, -1, helperErrMsg)

    If InStr(1, resultJson, MSG_NOMBRE_VACIO, vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_VACIO (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_VACIO"

    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_empty_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Sad_EmptyName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 11 - Edge: _ValidarNombreDocumento with same-as-current returns MSG_NOMBRE_IGUAL.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoAuditoria
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual

    resultJson = modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento( _
        "original_doc.pdf", docActual, colDocs, -1, helperErrMsg)

    If InStr(1, resultJson, MSG_NOMBRE_IGUAL, vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_IGUAL (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_IGUAL"

    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_same_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Edge_SameName_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 12 - Adversarial: _ValidarNombreDocumento with duplicate in collection.
' =============================================================================
Public Function Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic() As String
    Dim logs As Collection
    Dim resultJson As String
    Dim helperErrMsg As String
    Dim docActual As DocumentoAuditoria
    Dim docOtro As DocumentoAuditoria
    Dim colDocs As Scripting.Dictionary

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    helperErrMsg = ""

    Set docActual = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC, FIX_ID_NC, "original_doc.pdf")
    Set docOtro = BuildDocumentoAuditoriaForSlice6b(FIX_ID_DOC + 100, FIX_ID_NC, "nombre_que_ya_existe.pdf")

    Set colDocs = New Scripting.Dictionary
    colDocs.Add CStr(docActual.IDDocumento), docActual
    colDocs.Add CStr(docOtro.IDDocumento), docOtro

    resultJson = modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento( _
        "nombre_que_ya_existe.pdf", docActual, colDocs, -1, helperErrMsg)

    If InStr(1, resultJson, MSG_NOMBRE_DUPLICADO, vbTextCompare) = 0 Then
        Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
            TestHelper.BuildJsonFail("resultJson should contain MSG_NOMBRE_DUPLICADO (got='" & resultJson & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: resultJson contains MSG_NOMBRE_DUPLICADO"

    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
        TestHelper.BuildJsonOk(logs, "validar_nombre_duplicate_ok")
    Exit Function

EH:
    Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmAuditoriaDocumentosHelper_ValidarNombreDocumento_Adversarial_DuplicateName_Atomic: " & Err.Description, logs)
End Function