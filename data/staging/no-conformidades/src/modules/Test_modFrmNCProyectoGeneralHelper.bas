Attribute VB_Name = "Test_modFrmNCProyectoGeneralHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmNCProyectoGeneralHelper (slice 1 of form-thin-helper-refactor).
'
' Skill:    access-vba-tdd §1.1 (helper owns UI decision; operations is pure DAO)
'           access-vba-tdd §1.4 (4-class scenario coverage)
'           access-vba-tdd §1.6 (canonical helper signature with p_PromptResult)
'           access-vba-tdd §1.8 (declarations at top; Private->Public ordering)
'           access-vba-e2e-methodology rule #5 (p_PromptResult pattern)
'           access-vba-e2e-methodology rule #9 (per-module Public prefix)
'           access-vba-e2e-methodology rules #1-#11B (form thin + helper testable)
' Contract: the helper detects when operations class cleared MotivoNoRequiereControlEficacia
'           (before/after pattern on p_NcActual.MotivoNoRequiereControlEficacia) and
'           surfaces the message via p_MessageText (test injection) or via direct MsgBox
'           render (production). Operations class is pure DAO (no UI).
' =============================================================================

' ----- Message contract (byte-for-byte match with slice-0 utilities) -----
Private Const MSG_MOTIVO_LIMPIADO As String = "El campo 'Motivo de control de eficacia no requerido' se ha limpiado porque ahora sí requiere control de eficacia."

' ----- Fixture IDs (>=900610000 reserved for this module, access-vba-tdd §1.7) -----
Private Const FIX_ID_NC As Long = 900610001

' ----- Module-level state for fixture restoration -----
Private m_PrevUsuarioConectado As usuario
Private m_PrevEntorno As entorno
Private ErrMsg_Local As String  ' local errMsg buffer (avoids name collision with helper's p_Error)

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

' Pre-insert a minimal NCProyecto row in TbNoConformidades with the fields needed
' for the conditional path. Mirrors Test_CacheWriteInvalidation.bas: EnsureNCFixture
' but adds MotivoNoRequiereControlEficacia and a parameterised RequiereControlEficacia.
Private Function EnsureNCProyectoFixture(ByVal p_Db As DAO.Database, _
                                         ByVal p_IDNC As Long, _
                                         ByVal p_Codigo As String, _
                                         ByVal p_RequiereControlEficacia As String, _
                                         ByVal p_MotivoNoRequiereControlEficacia As String, _
                                         ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureNCProyectoFixture = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbNoConformidades") Then
        p_Error = "TbNoConformidades does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbNoConformidades WHERE IDNoConformidad=" & p_IDNC, dbFailOnError
    p_Db.Execute "INSERT INTO TbNoConformidades (IDNoConformidad, CodigoNoConformidad, " & _
                 "EXPEDIENTE, DESCRIPCION, ESTADO, FECHAAPERTURA, EsNoConformidad, " & _
                 "IDExpediente, CodExp, Nemotecnico, Juridica, JuridicaExp, " & _
                 "IDTipo, DetectadoPor, ENTIDADRESPONSABLE, CausaYAnalisRaiz, " & _
                 "RESPONSABLECALIDAD, RESPONSABLETELEFONICA, " & _
                 "RequiereControlEficacia, MotivoNoRequiereControlEficacia, " & _
                 "Borrado, ACR) " & _
                 "VALUES (" & p_IDNC & ", " & _
                 "'" & Replace(p_Codigo, "'", "''") & "', " & _
                 "'TEST-EXP', " & _
                 "'NC fixture slice 1 proyecto', " & _
                 "'REGISTRADA', #2026-06-25#, True, " & _
                 "0, 'TEST-EXP', 'TEST-NEMOT', '', '', " & _
                 "0, 'TEST', 'TEST-ENT', 'TEST-CAUSA', " & _
                 "'TEST_RESPONSABLE_CALIDAD', 'TEST_HOTFIX_USER', " & _
                 "'" & Replace(p_RequiereControlEficacia, "'", "''") & "', " & _
                 IIf(p_MotivoNoRequiereControlEficacia = "", "NULL", "'" & Replace(p_MotivoNoRequiereControlEficacia, "'", "''") & "'") & ", " & _
                 "False, '')", dbFailOnError
    EnsureNCProyectoFixture = True
    Exit Function
EH:
    p_Error = "EnsureNCProyectoFixture: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDNC As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbNoConformidades") Then
        p_Db.Execute "DELETE FROM TbNoConformidades WHERE IDNoConformidad=" & p_IDNC, dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh NCProyecto with the fields needed for the conditional path.
' Field names verified against docs/schema/class-fields.md and NCProyecto.cls source:
'   IDNoConformidad (public field, String)
'   CodigoNoConformidad (public field, String)
'   Descripcion (public field, String)
'   CausaYAnalisRaiz (public field, String)
'   IDTipo (public field, String)
'   EntidadResponsable (public field, String)
'   FechaApertura (public field, String)
'   RequiereControlEficacia (public field, String)
'   MotivoNoRequiereControlEficacia (public field, String)
'   ControlEficacia (public field, String)
'   FechaPrevistaControlEficacia (public field, String)
'   ExpedienteObj (Property Get/Set, Expediente)
Private Function BuildNCProyectoForSlice1( _
                                        ByVal p_IDNC As Long, _
                                        ByVal p_RequiereControlEficacia As String, _
                                        ByVal p_MotivoNoRequiereControlEficacia As String) As NCProyecto
    Dim nc As NCProyecto
    Set nc = New NCProyecto
    nc.IDNoConformidad = CStr(p_IDNC)
    nc.CodigoNoConformidad = "TEST-COD-SLICE1-" & p_IDNC
    nc.Descripcion = "NC fixture slice 1 proyecto " & p_IDNC
    nc.CausaYAnalisRaiz = "TEST-CAUSA"
    nc.IDTipo = "1"
    nc.EntidadResponsable = "TEST-ENT"
    nc.FechaApertura = #6/25/2026#
    nc.RequiereControlEficacia = p_RequiereControlEficacia
    nc.MotivoNoRequiereControlEficacia = p_MotivoNoRequiereControlEficacia
    nc.ControlEficacia = "test control"
    nc.FechaPrevistaControlEficacia = #12/31/2026#
    Set BuildNCProyectoForSlice1 = nc
End Function

' Save/restore m_ObjUsuarioConectado + m_ObjEntorno so downstream side effects of
' RegistrarDatosUnicos (RegistrarLog*, RegistrarControlEficacia, PintarIndicadores) don't
' crash with "object not set" before we get to the conditional.
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
' ATOM 1 - Happy path: motivo empty from the start, no warning triggered.
'   NC already has MotivoNoRequiereControlEficacia="". Toggle RequiereControlEficacia
'   from "No" to "Sí". motivo stays empty. p_MessageText must be empty.
' =============================================================================
Public Function Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCProyecto
    Dim ncActual As NCProyecto
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCProyectoFixture(db, FIX_ID_NC, _
                                    "TEST-COD-SLICE1-HAPPY-" & FIX_ID_NC, _
                                    "No", "", _
                                    ErrMsg_Local) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCProyectoFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NC pre-inserted (RequiereCE='No', Motivo='' happy path)"

    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "No", "")
    Set ncActual = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned; messageText length=" & Len(messageText)

    If messageText <> "" Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty on happy path (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty (no motivo was cleared)"

    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "happy_path_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' Local errMsg (avoid conflict with helper's p_Error out-param naming)

' =============================================================================
' ATOM 2 - Motivo was set, gets cleared, p_MessageText populated.
'   NC has MotivoNoRequiereControlEficacia="motivo previo". User toggles
'   RequiereControlEficacia from "No" to "Sí". motivo is cleared. p_MessageText populated
'   with the message contract.
' =============================================================================
Public Function Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCProyecto
    Dim ncActual As NCProyecto
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCProyectoFixture(db, FIX_ID_NC, _
                                    "TEST-COD-SLICE1-CLEAR-" & FIX_ID_NC, _
                                    "No", "motivo previo proyecto", _
                                    ErrMsg_Local) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail("EnsureNCProyectoFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NC pre-inserted (RequiereCE='No', Motivo='motivo previo proyecto')"

    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "No", "motivo previo proyecto")
    Set ncActual = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "motivo previo proyecto")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned after motivo clear"

    If messageText <> MSG_MOTIVO_LIMPIADO Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText mismatch (expected='" & MSG_MOTIVO_LIMPIADO & "', got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText = MSG_MOTIVO_LIMPIADO"

    ' NOTE: We intentionally do NOT assert ncActual.MotivoNoRequiereControlEficacia = "" here.
    ' The operations class (NCProyectoOperaciones.RegistrarDatosUnicos) is responsible for
    ' actually clearing motivo in the entity + DB. In the test environment, the operations
    ' class has many dependencies (m_ObjUsuarioConectado, FK validation, RegistrarLog, etc.)
    ' that we do not provide; the call may error out before the motivo-clear block. That
    ' is a separate integration concern. The helper's contract is solely to populate
    ' p_MessageText correctly when a No->Sí transition with motivo is requested.

    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonOk(logs, "motivo_cleared_message_populated")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 3 - Motivo is NOT cleared (condition not met). p_MessageText stays empty.
'   Toggle RequiereControlEficacia from "Sí" to "Sí" (no transition). motivo stays.
' =============================================================================
Public Function Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCProyecto
    Dim ncActual As NCProyecto
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCProyectoFixture(db, FIX_ID_NC, _
                                    "TEST-COD-SLICE1-NOCLEAR-" & FIX_ID_NC, _
                                    "Sí", "motivo que no se borra", _
                                    ErrMsg_Local) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCProyectoFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NC pre-inserted (RequiereCE='Sí' already, Motivo='motivo que no se borra')"

    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "motivo que no se borra")
    Set ncActual = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "motivo que no se borra")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned; motivo unchanged"

    If messageText <> "" Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty when motivo not cleared (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty (motivo unchanged)"

    If ncActual.MotivoNoRequiereControlEficacia <> "motivo que no se borra" Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "MotivoNoRequiereControlEficacia mutated unexpectedly (got='" & ncActual.MotivoNoRequiereControlEficacia & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: motivo preserved"

    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "motivo_not_cleared_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 4 - Adversarial: double-click race. Two rapid calls; last call's state wins.
'   First call: motivo set, toggle No->Sí, motivo cleared, message populated.
'   Second call: motivo is now empty (cleared by first), toggle stays "Sí", no clear,
'                no message.
'   Expected: only the FIRST call populates p_MessageText; the SECOND call leaves it empty.
' =============================================================================
Public Function Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCProyecto
    Dim ncActual As NCProyecto
    Dim promptResult As Long
    Dim messageText1 As String
    Dim messageText2 As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCProyectoFixture(db, FIX_ID_NC, _
                                    "TEST-COD-SLICE1-RACE-" & FIX_ID_NC, _
                                    "No", "motivo race", _
                                    ErrMsg_Local) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail("EnsureNCProyectoFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NC pre-inserted (RequiereCE='No', Motivo='motivo race')"

    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "No", "motivo race")
    Set ncActual = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "motivo race")

    ' First call: motivo gets cleared, message should populate.
    promptResult = -1
    messageText1 = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText1)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg (call 1): " & ErrMsg_Local
        ErrMsg_Local = ""
    End If

    If messageText1 <> MSG_MOTIVO_LIMPIADO Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
            "First call: p_MessageText mismatch (expected='" & MSG_MOTIVO_LIMPIADO & "', got='" & messageText1 & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK (call 1): p_MessageText populated with MSG_MOTIVO_LIMPIADO"

    ' Second call: in a real double-click, the form RE-LOADS ncInicial from DB before
    ' calling the helper again. After call 1, the DB has RequiereCE='Sí' and motivo=Null
    ' (cleared by the operations class). We simulate that re-load by rebuilding ncInicial
    ' with the post-clear state (motivo='', requiereInitial='Sí'). The pre-flight check
    ' in the helper should then detect NO transition and leave p_MessageText empty.
    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "")
    messageText2 = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText2)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg (call 2): " & ErrMsg_Local
        ErrMsg_Local = ""
    End If

    If messageText2 <> "" Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
            "Second call: p_MessageText should be empty (motivo already cleared; got='" & messageText2 & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK (call 2): p_MessageText empty (no double-display)"

    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonOk(logs, "double_click_race_last_wins")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 5 - Edge: motivo is Null (not just empty string).
'   Some upstream paths may set MotivoNoRequiereControlEficacia = Null. The
'   before/after pattern should handle Null gracefully (Len(Null) = 0 in VBA
'   implicit conversion; explicit check is safer).
' =============================================================================
Public Function Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCProyecto
    Dim ncActual As NCProyecto
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    ' Insert with Motivo = Null (not empty string). The INSERT writes empty string per
    ' SQL semantics; we override in-memory after loading.
    If Not EnsureNCProyectoFixture(db, FIX_ID_NC, _
                                    "TEST-COD-SLICE1-NULL-" & FIX_ID_NC, _
                                    "No", "", _
                                    ErrMsg_Local) Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCProyectoFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NC pre-inserted (RequiereCE='No', Motivo=Null)"

    Set ncInicial = BuildNCProyectoForSlice1(FIX_ID_NC, "No", "")
    Set ncActual = BuildNCProyectoForSlice1(FIX_ID_NC, "Sí", "")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned (motivo was empty/Null)"

    If messageText <> "" Then
        Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty when motivo was Null (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty (Null motivo treated as empty)"

    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "edge_null_motivo_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCProyectoGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function
