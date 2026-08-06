Attribute VB_Name = "Test_modFrmNCAuditoriaAccionesHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmNCAuditoriaAccionesHelper (slice 3 of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; operations is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper owns the 3 message-class decisions that previously lived
'           inline in Form_FormNCAuditoriaAcciones.cls:
'             1. _RenderError          -- boilerplate error handler (~48 of 52 MsgBox)
'             2. _ConfirmarBorradoNC   -- MsgBox("¿Desea borrar la AC?") + InputBox(motivo)
'             3. _ConfirmarHabilitacionNC -- MsgBox("¿Desea habilitar la NC?")
'           The form is thin: builds m_Error, calls helper.RenderError(m_Error, Err.Number).
'           Production caller passes p_PromptResult=0 (real modal). Atom passes -1
'           (no modal, assert via JSON value or return code).
' Slice:    3 of 27 (Acciones pair: NCAuditoriaAcciones; companion is
'           Test_modFrmNCProyectoAccionesHelper.bas on NCProyecto side).
' =============================================================================

' ----- Message contract (byte-for-byte match with production literals) -----
Private Const MSG_BORRADO_PROMPT As String = "MSG-AUDITORIA-ACCIONES-BORRADO: ¿Desea borrar la Acción Correctiva de auditoría seleccionada?"
Private Const MSG_HABILITACION_PROMPT As String = "MSG-AUDITORIA-ACCIONES-HABILITACION: ¿Desea habilitar la no conformidad de auditoría seleccionada?"
Private Const MSG_MOTIVO_PROMPT As String = "MSG-AUDITORIA-ACCIONES-MOTIVO: Introduzca el motivo para marcar como borrada la no conformidad de auditoría"
Private Const MSG_BORRADO_TITLE As String = "Borrado de NC"
Private Const MSG_HABILITACION_TITLE As String = "Habilitación de NC"
Private Const MSG_MOTIVO_TITLE As String = "Motivación"
Private Const MSG_ERROR_TITLE As String = "Error"
Private Const MSG_ADVERTENCIA_TITLE As String = "Advertencia"

' ----- Fixture IDs (>=900650000 reserved for this module, access-vba-tdd §1.7) -----
Private Const FIX_ID_NC As Long = 900650001

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

' Pre-insert a minimal NCAuditoria row in TbNoConformidadesAuditoria for the
' Eliminar/Habilitar path. Mirrors Test_modFrmNCAuditoriaGestionHelper.bas:
' EnsureNCAuditoriaFixtureForSlice2b but uses FIX_ID_NC in the 900650001 range
' (slice 3 NCAuditoria range, distinct from slice 2b's 900630001).
Private Function EnsureNCAuditoriaFixtureForSlice3(ByVal p_Db As DAO.Database, _
                                                    ByVal p_IDNC As Long, _
                                                    ByVal p_Borrado As Boolean, _
                                                    ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureNCAuditoriaFixtureForSlice3 = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Error = "TbNoConformidadesAuditoria does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & p_IDNC, dbFailOnError
    p_Db.Execute "INSERT INTO TbNoConformidadesAuditoria (id, " & _
                 "DESCRIPCION, ESTADO, FECHAAPERTURA, " & _
                 "Tipo, PuntoNorma, CAUSARAIZ, " & _
                 "RequiereControlEficacia, " & _
                 "Borrado, MotivoBorrado) " & _
                 "VALUES (" & p_IDNC & ", " & _
                 "'NC fixture slice 3 auditoria acciones', " & _
                 "'REGISTRADA', #2026-06-26#, " & _
                 "'Auditoria', 'TEST-PN-SLICE3', 'TEST-CAUSA-SLICE3', " & _
                 "'No', " & _
                 IIf(p_Borrado, "True", "False") & ", " & _
                 IIf(p_Borrado, "'motivo fixture slice 3'", "NULL") & ")", dbFailOnError
    EnsureNCAuditoriaFixtureForSlice3 = True
    Exit Function
EH:
    p_Error = "EnsureNCAuditoriaFixtureForSlice3: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDNC As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & p_IDNC, dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh NCAuditoria with the fields needed for ConfirmarBorradoNC /
' ConfirmarHabilitacionNC. Field names verified against docs/schema/class-fields.md
' (NCAuditoria section) and NCAuditoria.cls source (public String fields):
'   id (String), IDAuditoria, FechaApertura, Numero, Descripcion, CAUSARAIZ,
'   AccionCorrectiva, CORRECCION, FECHACIERRE, FPREVCIERRE, RESPONSABLEIMPLANTACION,
'   RequiereControlEficacia, ControlEficacia, FechaControlEficacia,
'   FechaPrevistaControlEficacia, ResultadoControlEficacia, ConformeControlEficacia,
'   RequiereAccionCorrectiva, MotivoNoAccionCorrectiva, MotivoNoRequiereControlEficacia,
'   Tipo, PuntoNorma, Estado, Borrado (Boolean), MotivoBorrado, Notas, Cerrada.
Private Function BuildNCAuditoriaForSlice3(ByVal p_IDNC As Long, _
                                            ByVal p_Borrado As Boolean, _
                                            ByVal p_MotivoBorrado As String) As NCAuditoria
    Dim nc As NCAuditoria
    Set nc = New NCAuditoria
    nc.id = CStr(p_IDNC)
    nc.Descripcion = "NC fixture slice 3 auditoria acciones " & p_IDNC
    nc.CAUSARAIZ = "TEST-CAUSA-SLICE3"
    nc.PuntoNorma = "TEST-PN-SLICE3"
    nc.Tipo = "Auditoria"
    nc.FechaApertura = #6/26/2026#
    nc.RequiereControlEficacia = "No"
    nc.RESPONSABLEIMPLANTACION = "TEST_RESPONSABLE_SLICE3"
    nc.Borrado = p_Borrado
    nc.MotivoBorrado = p_MotivoBorrado
    Set BuildNCAuditoriaForSlice3 = nc
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
Public Function Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode: no real MsgBox
    helperErrMsg = ""

    returnCode = modFrmNCAuditoriaAccionesHelper_RenderError( _
        "Al ComandoACBorrar_Click se ha producido el error n: 9" & vbNewLine & "Detalle: subscript out of range", _
        9, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty in happy path (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 in atom mode (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (no modal sentinel)"

    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_no_modal_ok")
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_RenderError_Happy_NoModal_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 2 - Sad path: _RenderError with Err.Number=1000 (warning path).
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode
    helperErrMsg = ""

    returnCode = modFrmNCAuditoriaAccionesHelper_RenderError( _
        "Seleccione un elemento de la lista superior", _
        1000, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on sad path (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    If returnCode <> -1 Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on sad path too (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (atom mode invariant regardless of ErrNumber)"

    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_warning_path_ok")
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_RenderError_Sad_WarningPath_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 3 - Edge: empty m_ErrorText passed to _RenderError.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic() As String
    Dim logs As Collection
    Dim returnCode As Long
    Dim helperErrMsg As String
    Dim promptResult As Long

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    promptResult = -1  ' atom mode
    helperErrMsg = ""

    returnCode = modFrmNCAuditoriaAccionesHelper_RenderError( _
        "", _
        0, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on edge (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty on empty text"

    If returnCode <> -1 Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on edge (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (no modal on empty text)"

    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_empty_text_ok")
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_RenderError_Edge_EmptyText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 4 - Adversarial: _RenderError with very long m_ErrorText (>1000 chars).
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic() As String
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

    longText = "Stack: "
    For i = 1 To 200
        longText = longText & "ABCDEFGHIJ"
    Next i

    returnCode = modFrmNCAuditoriaAccionesHelper_RenderError( _
        longText, _
        13, promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on adversarial (got='" & helperErrMsg & "')", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty on 2000-char text"

    If returnCode <> -1 Then
        Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic = _
            TestHelper.BuildJsonFail("returnCode should be -1 on adversarial (got=" & returnCode & ")", logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert OK: returnCode = -1 (atom mode invariant on long text)"

    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonOk(logs, "render_error_long_text_ok")
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_RenderError_Adversarial_VeryLongText_Atomic: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM 5 - Happy: _ConfirmarBorradoNC atom mode (p_PromptResult=-1).
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim motivo As String
    Dim ncParaBorrar As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncParaBorrar = BuildNCAuditoriaForSlice3(FIX_ID_NC, False, "")
    promptResult = -1  ' atom mode
    motivo = ""

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC( _
        ncParaBorrar, , promptResult, motivo, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on happy (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict Is Nothing Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parser returned Nothing", logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If InStr(1, CStr(dict("value")), MSG_BORRADO_PROMPT, vbBinaryCompare) = 0 Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected value to contain borrado prompt; got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value contains borrado prompt contract"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_borrado_prompt_text_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Happy_ReturnsPromptText_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 6 - Sad: _ConfirmarBorradoNC with user_clicked_no.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim motivo As String
    Dim ncParaBorrar As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncParaBorrar = BuildNCAuditoriaForSlice3(FIX_ID_NC, False, "")
    promptResult = vbNo
    motivo = ""

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC( _
        ncParaBorrar, , promptResult, motivo, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on sad (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If CStr(dict("value")) <> "user_rejected" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected value='user_rejected', got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value=user_rejected"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_borrado_user_rejected_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Sad_UserRejected_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 7 - Edge: _ConfirmarBorradoNC with user_clicked_yes but empty motivo.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim motivo As String
    Dim ncParaBorrar As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncParaBorrar = BuildNCAuditoriaForSlice3(FIX_ID_NC, False, "")
    promptResult = vbYes
    motivo = ""

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC( _
        ncParaBorrar, , promptResult, motivo, helperErrMsg)

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    If dict("ok") <> False Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
            TestHelper.BuildJsonFail("expected ok=false on empty motivo (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If InStr(1, CStr(dict("error")), "motivo", vbBinaryCompare) = 0 Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
            TestHelper.BuildJsonFail("expected error mentioning motivo, got='" & CStr(dict("error")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: ok=false and error mentions motivo"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_borrado_empty_motivo_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Edge_EmptyMotivo_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 8 - Adversarial: _ConfirmarBorradoNC with p_NC = Nothing.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim motivo As String
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    promptResult = vbYes
    motivo = "test motivo"

    On Error Resume Next
    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC( _
        Nothing, , promptResult, motivo, helperErrMsg)
    On Error GoTo EH

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    If dict("ok") <> False Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("expected ok=false on Nothing NC (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: ok=false on Nothing NC"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_borrado_nothing_nc_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarBorradoNC_Adversarial_NothingNC_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 9 - Happy: _ConfirmarHabilitacionNC atom mode (p_PromptResult=-1).
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim ncParaHabilitar As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncParaHabilitar = BuildNCAuditoriaForSlice3(FIX_ID_NC, True, "motivo fixture")
    promptResult = -1  ' atom mode

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC( _
        ncParaHabilitar, , promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on happy (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict Is Nothing Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("JSON parser returned Nothing", logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If InStr(1, CStr(dict("value")), MSG_HABILITACION_PROMPT, vbBinaryCompare) = 0 Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
            TestHelper.BuildJsonFail("expected value to contain habilitacion prompt; got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value contains habilitacion prompt contract"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_habilitacion_prompt_text_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Happy_ReturnsPromptText_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 10 - Sad: _ConfirmarHabilitacionNC with user_clicked_no.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim ncParaHabilitar As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncParaHabilitar = BuildNCAuditoriaForSlice3(FIX_ID_NC, True, "motivo fixture")
    promptResult = vbNo

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC( _
        ncParaHabilitar, , promptResult, helperErrMsg)

    If helperErrMsg <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("helperErrMsg should be empty on sad (got='" & helperErrMsg & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: helperErrMsg empty"

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If
    If dict("ok") <> True Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected ok=true (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    If CStr(dict("value")) <> "user_rejected" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
            TestHelper.BuildJsonFail("expected value='user_rejected', got='" & CStr(dict("value")) & "'", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: value=user_rejected"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_habilitacion_user_rejected_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Sad_UserRejected_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 11 - Edge: _ConfirmarHabilitacionNC with p_NC NOT borrada (already enabled).
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim ncNoBorrada As NCAuditoria
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    Set ncNoBorrada = BuildNCAuditoriaForSlice3(FIX_ID_NC, False, "")  ' NOT borrada
    promptResult = vbYes

    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC( _
        ncNoBorrada, , promptResult, helperErrMsg)

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    If dict("ok") <> False Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic = _
            TestHelper.BuildJsonFail("expected ok=false on non-borrada NC (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: ok=false on already-enabled NC"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_habilitacion_already_enabled_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Edge_AlreadyEnabled_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function

' =============================================================================
' ATOM 12 - Adversarial: _ConfirmarHabilitacionNC with p_NC = Nothing.
' =============================================================================
Public Function Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic() As String
    Dim logs As Collection
    Dim jsonResult As String
    Dim promptResult As Long
    Dim helperErrMsg As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    SetupTestContext

    promptResult = vbYes

    On Error Resume Next
    jsonResult = modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC( _
        Nothing, , promptResult, helperErrMsg)
    On Error GoTo EH

    Dim jsonErr As String
    Dim dict As Object
    Set dict = JSONHelper.JSONAObjeto(jsonResult, jsonErr)
    If jsonErr <> "" Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("JSON parse failed: " & jsonErr, logs)
        GoTo Cleanup
    End If

    If dict("ok") <> False Then
        Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic = _
            TestHelper.BuildJsonFail("expected ok=false on Nothing NC (got=" & dict("ok") & ")", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: ok=false on Nothing NC"

    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic = _
        TestHelper.BuildJsonOk(logs, "confirmar_habilitacion_nothing_nc_ok")

Cleanup:
    On Error Resume Next
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Exit Function

EH:
    Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic = _
        TestHelper.BuildJsonFail("Test_modFrmNCAuditoriaAccionesHelper_ConfirmarHabilitacionNC_Adversarial_NothingNC_Atomic: " & Err.Description, logs)
    Resume Cleanup
End Function
